using MonteCarlo.Core.Tests;
using MonteCarlo.Excel;
using System.Reflection;

internal static class OptimizerLayoutChecks
{
    public static void Run()
    {
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        {
            var p = OptimizationTests.Problem();
            var result = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => new OptimizationTests.Book());
            OptimizationProblem? saved = null;
            var deterministicResult = OptimizationService.Run(p, [], [OptimizationTests.Forecast()], () => new OptimizationTests.Book());
            using var deterministic = new OptimizerForm([OptimizationTests.Forecast()], p, x => x,
                (_, _, _) => Task.FromResult(deterministicResult), _ => { }, _ => { }, deterministic: true);
            Prepare(deterministic, scale);
            if (Descendants(deterministic).Single(c => c.Name == "EvaluationMode").Text != "Deterministic") throw new Exception("Detected mode missing");
            if (Descendants(deterministic).OfType<ComboBox>().ElementAt(1).Items.Count != 4) throw new Exception("Deterministic probability options enabled");
            if (!Descendants(deterministic).Any(c => c.Text == "N/A – Deterministic model")) throw new Exception("Deterministic trial status missing");
            using var deterministicConstraint = new OptimizerConstraintForm(null, p.Variables, [OptimizationTests.Forecast()], () => null,
                _ => throw new Exception(), deterministic: true);
            Prepare(deterministicConstraint, scale);
            if (Descendants(deterministicConstraint).Single(c => c.Name == "ConstraintType") is not ComboBox kinds || kinds.Items.Count != 5 || kinds.Items.Cast<string>().Any(s => s.Contains("probability"))) throw new Exception("Deterministic probability constraint enabled");
            if (!Descendants(deterministic).Any(c => c.Text == "Maximum Candidate Evaluations") || !Descendants(deterministic).Any(c => c.Name == "ComputationalEstimate" && c.Text.Contains("0."))) throw new Exception("Candidate budget/estimate missing");
            using var deterministicResults = new OptimizerResultsForm(deterministicResult, _ => { }, _ => { }); Prepare(deterministicResults, scale);
            foreach (var button in Descendants(deterministicResults).OfType<Button>()) Bounds(deterministicResults, button, 30, 20);
            using var form = new OptimizerForm([OptimizationTests.Forecast()], p, x => saved = x, (_, _, _) => Task.FromResult(result), _ => { }, _ => { });
            Prepare(form, scale);
            var tabs = Descendants(form).OfType<TabControl>().Single();
            var selectors = Descendants(form).OfType<ComboBox>().ToArray();
            if (selectors[1].Items.Count != 6) throw new Exception("Optimizer objective options missing");
            foreach (var page in tabs.TabPages.Cast<TabPage>())
            {
                tabs.SelectedTab = page; Layout(form);
                foreach (var grid in Descendants(page).OfType<DataGridView>()) Bounds(form, grid, 400, 100);
                foreach (var button in Descendants(page).OfType<Button>()) Bounds(form, button, 30, 20);
            }
            tabs.SelectedIndex = 0;
            selectors[1].SelectedIndex = 4; Layout(form);
            if (!selectors[2].Enabled) throw new Exception("Optimizer target direction unavailable");
            foreach (var button in Descendants(form).OfType<Button>().Where(b => b.Text is "Save Configuration" or "Run Optimizer" or "Cancel Optimization" or "Close")) Bounds(form, button, 30, 20);
            var save = Descendants(form).OfType<Button>().Single(b => b.Text == "Save Configuration"); Click(save);
            if (saved?.Objective.Metric != OptimizationMetric.Probability) throw new Exception("Optimizer save objective failed");
            using var results = new OptimizerResultsForm(result, _ => { }, _ => { }); Prepare(results, scale);
            var resultTabs = Descendants(results).OfType<TabControl>().Single();
            foreach (TabPage page in resultTabs.TabPages) { resultTabs.SelectedTab = page; Layout(results); Bounds(results, Descendants(page).OfType<DataGridView>().Single(), 400, 150); }
            resultTabs.SelectedIndex = 0; Layout(results);
            foreach (var button in Descendants(results).OfType<Button>()) Bounds(results, button, 30, 20);
            int applied = 0, exported = 0;
            using var workflow = new OptimizerResultsForm(result, _ => applied++, _ => exported++); Prepare(workflow, scale);
            var buttons = Descendants(workflow).OfType<Button>().ToArray();
            Click(buttons.Single(b => b.Text == "Export to Excel")); if (exported != 1 || applied != 0) throw new Exception("Export applied values");
            // Close has no apply callback; successful apply is covered by engine tests without modal message boxes.
            Click(buttons.Single(b => b.Text == "Close Without Applying")); if (applied != 0) throw new Exception("Close applied values");
            int selections = 0;
            using var dialog = new DecisionVariableForm(p.Variables[0], () => { selections++; return null; }, v => new(v.Sheet, "$C$1", 50, "Suggested"), _ => { });
            Prepare(dialog, scale); foreach (var button in Descendants(dialog).OfType<Button>()) Bounds(dialog, button, 30, 20);
            foreach (var field in Descendants(dialog).OfType<TextBox>()) Bounds(dialog, field, 120, 18);
            var originalSize = dialog.ClientSize;
            dialog.ClientSize = new(dialog.ClientSize.Width, Math.Max(280, (int)(280 * scale)));
            using (var largerFont = new Font(dialog.Font.FontFamily, 11))
            {
                var originalFont = dialog.Font; dialog.Font = largerFont; Layout(dialog);
                foreach (var button in Descendants(dialog).OfType<Button>().Where(b => b.Text is "OK" or "Cancel")) Bounds(dialog, button, 30, 20);
                dialog.Font = originalFont;
            }
            dialog.ClientSize = originalSize; Layout(dialog);
            Directory.CreateDirectory("artifacts/optimizer-layout");
            using (var preview = new Bitmap(dialog.Width, dialog.Height)) { preview.SetResolution(96 * scale, 96 * scale); dialog.DrawToBitmap(preview, new Rectangle(Point.Empty, preview.Size)); preview.Save($"artifacts/optimizer-layout/decision-variable-{scale}.png"); }
            var reference = Descendants(dialog).Single(c => c.Name == "SelectedCell");
            if (reference.Text != "Model!$C$1" || Descendants(dialog).Single(c => c.Name == "CurrentValue").Text != "50") throw new Exception("Selected reference/current value missing");
            Click(Descendants(dialog).OfType<Button>().Single(b => b.Text == "Select"));
            if (selections != 1 || reference.Text != "Model!$C$1") throw new Exception("Picker cancellation changed selection");
            Click(Descendants(dialog).OfType<Button>().Single(b => b.Text == "OK"));
            if (dialog.Variable?.Name != "X" || dialog.Variable.Cell != "$C$1") throw new Exception("Edit/name retention failed");
            using var added = new DecisionVariableForm(null, () => new("Model", "$C$2", 103, "Selling Price"), _ => throw new Exception(), _ => { });
            Prepare(added, scale); Click(Descendants(added).OfType<Button>().Single(b => b.Text == "Select"));
            if (Descendants(added).Single(c => c.Name == "VariableName").Text != "Selling Price") throw new Exception("Name suggestion missing");
            Click(Descendants(added).OfType<Button>().Single(b => b.Text == "OK"));
            if (added.Variable?.Cell != "$C$2") throw new Exception("New selection not accepted");
            using var tracked = new DecisionVariableForm(p.Variables[0] with { CellLink = "tracked" }, () => new("Renamed", "$C$1", 50, "Other"), _ => new("Renamed", "$C$1", 50, null), _ => { });
            Prepare(tracked, scale); Click(Descendants(tracked).OfType<Button>().Single(b => b.Text == "Select")); Click(Descendants(tracked).OfType<Button>().Single(b => b.Text == "OK"));
            if (tracked.Variable?.CellLink != "tracked" || tracked.Variable.Name != "X") throw new Exception("Tracked rename or entered name lost");
            using var bitmap = new Bitmap(results.Width, results.Height); bitmap.SetResolution(96 * scale, 96 * scale); results.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
            Directory.CreateDirectory("artifacts/optimizer-layout"); bitmap.Save($"artifacts/optimizer-layout/results-{scale}.png");
            Console.WriteLine($"OPTIMIZER PASS scale {scale}: six objectives, grids, target controls, save, results/export/close and editor bounds");
        }
    }
    private static void Click(Button b) => typeof(Button).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(b, [EventArgs.Empty]);
    private static IEnumerable<Control> Descendants(Control root) { foreach (Control child in root.Controls) { yield return child; foreach (var c in Descendants(child)) yield return c; } }
    private static void Layout(Form form) { for (int i = 0; i < 4; i++) { form.PerformLayout(); foreach (var c in Descendants(form)) c.PerformLayout(); } }
    private static void Prepare(Form form, float scale) { _ = form.Handle; foreach (var c in Descendants(form)) _ = c.Handle; form.Scale(new SizeF(scale, scale)); Layout(form); }
    private static void Bounds(Form form, Control c, int width, int height)
    {
        var origin = form.PointToClient(c.PointToScreen(Point.Empty));
        if (c.Width < width || c.Height < height || origin.X < 0 || origin.Y < 0 || origin.X + c.Width > form.ClientSize.Width || origin.Y + c.Height > form.ClientSize.Height)
            throw new Exception($"Optimizer control clipped: {c.Text}/{c.Name}, {origin}, {c.Size}, client {form.ClientSize}");
    }
}
