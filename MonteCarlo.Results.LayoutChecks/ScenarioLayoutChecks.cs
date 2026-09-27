using MonteCarlo.Core;
using MonteCarlo.Excel;
using System.Reflection;

internal static class ScenarioLayoutChecks
{
    public static void Run()
    {
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        foreach (var mode in Enum.GetValues<ScenarioAnalysisMode>())
        {
            var input = new AssumptionDefinition { Name = "Long revenue assumption", SheetName = "Model", CellAddress = "A1", CellLink = "a", Distribution = DistributionType.Normal, Parameter1 = 100, Parameter2 = 15 };
            var forecasts = new[] { new ForecastDefinition { Name = "Profit", SheetName = "Model", CellAddress = "B1", CellLink = "f1" }, new ForecastDefinition { SheetName = "Model", CellAddress = "B2", CellLink = "f2" } };
            var scenarios = new[] { new ScenarioDefinition("1", "Growth", new[] { new ScenarioAdjustment("a", ScenarioAdjustmentType.PercentageChange, 15) }), new ScenarioDefinition("2", "Downside", Array.Empty<ScenarioAdjustment>()) };
            var analysis = new ScenarioAnalysisDefinition(scenarios, null, mode, 80,
                new[] { new ScenarioForecastQuestion("f1", TargetDirection.AtOrAbove, 120, TargetDirection.AtOrBelow, 70), new ScenarioForecastQuestion("f2", TargetDirection.AtOrBelow, 300, TargetDirection.AtOrAbove, 95) });
            var result = new ScenarioComparisonResult(10000, int.MinValue, "Downside", mode, 80, new[] {
                new ScenarioComparisonRow("2", "Downside", true, "f1", "Profit (Model!B1)", 100, null, mode == ScenarioAnalysisMode.ProbabilityToValue ? 80 : .8, null, 120, TargetDirection.AtOrBelow),
                new ScenarioComparisonRow("1", "Growth", false, "f1", "Profit (Model!B1)", 115, 15, mode == ScenarioAnalysisMode.ProbabilityToValue ? 100 : .6, mode == ScenarioAnalysisMode.ProbabilityToValue ? 20 : -20, 120, TargetDirection.AtOrBelow) });
            ScenarioAnalysisDefinition? saved = null;
            using var editor = new ScenarioAnalysisForm(new[] { input }, forecasts, analysis,
                items => saved = items, (_, _, _) => Task.FromResult(result));
            Prepare(editor, scale);
            var controls = Descendants(editor).ToArray();
            var grid = controls.OfType<DataGridView>().Single(g => g.Name == "Adjustments");
            var list = controls.OfType<CheckedListBox>().Single();
            var tabs = controls.OfType<TabControl>().Single();
            if (grid.Rows.Count != 1 || list.Items.Count != 2 || list.CheckedItems.Count != 2) throw new Exception("Scenario editor data missing");
            AssertBounds(editor, grid, 220, 80);
            var baseline = controls.OfType<Button>().Single(b => b.Text == "Set as Baseline");
            AssertBounds(editor, baseline, 20, 15);
            if (!controls.Single(c => c.Name == "BaselineSelection").Text.Contains("None")) throw new Exception("Invented a baseline");
            list.SelectedIndex = 1; Click(baseline);
            if (!controls.Single(c => c.Name == "BaselineSelection").Text.Contains("Downside")) throw new Exception("Baseline selection not visible");
            list.SetItemChecked(1, false); if (!list.GetItemChecked(1)) throw new Exception("Baseline excluded from comparison");
            grid.Rows[0].Cells[2].Value = "-20"; Invoke(editor, "CommitRow", 0); Invoke(editor, "Save");
            if (saved == null || saved.Scenarios[1].Adjustments.Single().Value != -20 || saved.BaselineScenarioId != "2") throw new Exception("Scenario/baseline save failed");
            grid.Rows[0].Cells[2].Value = ""; Invoke(editor, "CommitRow", 0); Invoke(editor, "Save");
            if (saved!.Scenarios[1].Adjustments.Count != 0) throw new Exception("No change did not remove adjustment");
            tabs.SelectedIndex = 1; Layout(editor);
            var questionGrid = controls.OfType<DataGridView>().Single(g => g.Name == "ForecastQuestions");
            AssertBounds(editor, questionGrid, 220, 80);
            if (questionGrid.Rows.Count != 2 || questionGrid.Columns[1].Visible != (mode == ScenarioAnalysisMode.ValueToProbability)) throw new Exception("Mode-specific forecast question controls incorrect");
            var probability = questionGrid.Columns[3];
            if (probability.Visible != (mode == ScenarioAnalysisMode.ProbabilityToValue)) throw new Exception("Probability activation incorrect");
            Invoke(editor, "Save");
            if (saved!.Questions[0].ValueDirection != TargetDirection.AtOrAbove || saved.Questions[0].ProbabilityDirection != TargetDirection.AtOrBelow || saved.Questions[1].Target != 300) throw new Exception("Mode settings mixed or lost");
            var modeSelector = controls.OfType<ComboBox>().Single(c => c.Name == "AnalysisMode");
            modeSelector.SelectedIndex = mode == ScenarioAnalysisMode.ProbabilityToValue ? 1 : 0;
            Invoke(editor, "Save");
            modeSelector.SelectedIndex = (int)mode; Invoke(editor, "Save");
            if (saved!.Questions[0].Probability != 70 || saved.Questions[1].Probability != 95 || saved.Mode != mode || saved.Questions[0].ValueDirection != TargetDirection.AtOrAbove || saved.Questions[0].ProbabilityDirection != TargetDirection.AtOrBelow || saved.Questions[0].Target != 120)
                throw new Exception("Switching analysis modes corrupted settings");
            foreach (var button in controls.OfType<Button>().Where(b => b.Text is "Save definitions" or "Run Scenario Analysis" or "Close")) AssertBounds(editor, button, 20, 15);
            using var comparison = new ScenarioComparisonForm(result); Prepare(comparison, scale);
            var comparisonGrid = Descendants(comparison).OfType<DataGridView>().Single();
            AssertBounds(comparison, comparisonGrid, 220, 80);
            if (comparisonGrid.Columns.Count != 6 || comparisonGrid.Columns.Cast<DataGridViewColumn>().Any(c => c.HeaderText.Contains("P50") || c.HeaderText.Contains("P80"))) throw new Exception("Old statistics columns retained");
            if (Convert.ToString(comparisonGrid.Rows[0].Cells[3].Value) != "—" || Convert.ToString(comparisonGrid.Rows[0].Cells[5].Value) != "—") throw new Exception("Baseline deltas must be neutral");
            if (!Convert.ToString(comparisonGrid.Rows[0].Cells[0].Value)!.Contains("Downside ★ Baseline")) throw new Exception("Baseline scenario name lost");
            if (mode == ScenarioAnalysisMode.ValueToProbability && !Convert.ToString(comparisonGrid.Rows[1].Cells[5].Value)!.EndsWith(" pp")) throw new Exception("Probability delta lacks pp units");
            var context = Descendants(comparison).Single(c => c.Name == "ForecastContext");
            if (!context.Text.Contains("Model!B1") || !context.Text.Contains("≤")) throw new Exception("Forecast question context missing");
            foreach (var button in Descendants(comparison).OfType<Button>()) AssertBounds(comparison, button, 20, 15);
            Console.WriteLine($"SCENARIO PASS {mode} scale {scale}: explicit baseline, both question settings, six decision columns, footer");
        }
    }
    private static void Click(Button button) => typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(button, new object[] { EventArgs.Empty });
    private static void Invoke(object instance, string method, params object[] args) => instance.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(instance, args);
    private static IEnumerable<Control> Descendants(Control root) { foreach (Control child in root.Controls) { yield return child; foreach (var nested in Descendants(child)) yield return nested; } }
    private static void Prepare(Form form, float scale) { _ = form.Handle; foreach (var child in Descendants(form)) _ = child.Handle; form.Scale(new SizeF(scale, scale)); Layout(form); }
    private static void Layout(Form form) { for (int i = 0; i < 4; i++) { form.PerformLayout(); foreach (var child in Descendants(form)) child.PerformLayout(); } }
    private static void AssertBounds(Form form, Control control, int minWidth, int minHeight)
    {
        var origin = form.PointToClient(control.PointToScreen(Point.Empty));
        if (control.Width < minWidth || control.Height < minHeight || origin.X < 0 || origin.Y < 0 || origin.X + control.Width > form.ClientSize.Width || origin.Y + control.Height > form.ClientSize.Height)
            throw new Exception($"Scenario control clipped: {control.Name}/{control.Text}, {origin}, {control.Size}, form {form.ClientSize}");
    }
}
