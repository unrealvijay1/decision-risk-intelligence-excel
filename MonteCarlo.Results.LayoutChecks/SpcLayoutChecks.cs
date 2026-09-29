using MonteCarlo.Core;
using MonteCarlo.Excel;

internal static class SpcLayoutChecks
{
    public static void Run()
    {
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        {
            using var editor = new SpcConfigurationForm(null, "Data!A1:A50", "Dates!B1:B50", _ => null, () => { }, _ => { });
            Prepare(editor, scale); CheckButtons(editor);
            Preview(editor, $"spc-configuration-{scale}.png");
            if (editor.ReadConfiguration().Rules != new SpcRules()) throw new Exception("SPC rule defaults changed.");
            var checks = Descendants(editor).OfType<CheckBox>().ToArray();
            string[] ruleLabels = ["Beyond control limits (3σ)", "Eight-point run", "Two of three beyond 2σ", "Four of five beyond 1σ", "Six-point trend"];
            if (!checks.Select(c => c.Text).Order().SequenceEqual(ruleLabels.Order())) throw new Exception("SPC rule controls are missing.");
            foreach (var check in checks)
            {
                if (check.Width < check.PreferredSize.Width || check.Height < check.PreferredSize.Height || check.Right > check.Parent!.ClientSize.Width)
                    throw new Exception($"SPC rule clipped at {scale}: {check.Text}");
                foreach (var other in checks.Where(c => c != check && c.Parent == check.Parent))
                    if (check.Bounds.IntersectsWith(other.Bounds)) throw new Exception("SPC rule controls overlap.");
                var scroll = Descendants(editor).OfType<Panel>().Single(p => p.AutoScroll);
                scroll.ScrollControlIntoView(check);
                var visible = scroll.RectangleToClient(check.RectangleToScreen(check.ClientRectangle));
                if (!scroll.ClientRectangle.Contains(visible)) throw new Exception($"SPC rule cannot be scrolled into view: {check.Text}");
                check.Checked = false;
            }
            Preview(editor, $"spc-configuration-rules-{scale}.png");
            if (editor.ReadConfiguration().Rules != new SpcRules(false, false, false, false, false)) throw new Exception("SPC rule selections were not read.");
            var values = Enumerable.Range(0, 50).Select(i => 10d + (i % 2 == 0 ? -1 : 1)).ToArray();
            for (int i = 30; i < 38; i++) values[i] = 35;
            var ranges = values.Select((v, i) => i == 0 ? (double?)null : Math.Abs(v - values[i - 1])).ToArray();
            var baseline = new SpcBaseline(1, 20, 10, 2, 4.68, 15.32, 6.536);
            var result = new SpcResult(new() { ProcessName = "Long process name for layout validation", Target = new(14, TargetDirection.AtOrBelow), MeasurementUnit = "days" },
                values, Enumerable.Range(1, 50).Select(i => new DateTime(2026, 1, 1).AddDays(i).ToString("dd MMM yyyy")).ToArray(), ranges, baseline,
                SpcSignalDetector.Detect(values, ranges, baseline, new()));
            using var results = new SpcResultsForm(result, () => { }); Prepare(results, scale); CheckButtons(results);
            if (!result.Signals.Any(s => s.RuleId == 3) || !result.Signals.Any(s => s.RuleId == 4)) throw new Exception("SPC multi-rule fixture failed.");
            var grid = Descendants(results).OfType<DataGridView>().Single();
            ((TabControl)grid.Parent!.Parent!).SelectedTab = (TabPage)grid.Parent;
            results.PerformLayout(); grid.PerformLayout();
            if (!grid.Columns.Contains("QualifyingObservations") || !grid.Columns.Contains("WindowValues") || grid.Rows.Count != result.Signals.Count)
                throw new Exception("SPC signal details are incomplete.");
            Preview(results, $"spc-results-{scale}.png");
            foreach (var chart in Descendants(results).OfType<SpcChartPanel>())
            {
                using var bitmap = new Bitmap(chart.Width, chart.Height); bitmap.SetResolution(96 * scale, 96 * scale);
                chart.DrawToBitmap(bitmap, chart.ClientRectangle);
                if (chart.PlotBounds.Width < 100 || chart.PlotBounds.Height < 50) throw new Exception($"SPC chart plot is too small: {scale}, panel {chart.Size}, font {chart.Font.Size}, plot {chart.PlotBounds}.");
                int colored = 0;
                for (int y = 0; y < bitmap.Height; y += 3) for (int x = 0; x < bitmap.Width; x += 3)
                { var pixel = bitmap.GetPixel(x, y); if (Math.Abs(pixel.R - pixel.B) > 25) colored++; }
                if (colored < 30) throw new Exception("SPC chart did not render its series and reference lines.");
                if (Environment.GetEnvironmentVariable("RESULTS_LAYOUT_PREVIEW") == "1") bitmap.Save(Path.Combine(AppContext.BaseDirectory, $"spc-{chart.Chart}-{scale}.png"));
            }
            results.Size = new Size((int)(950 * scale), (int)(790 * scale));
            for (int i = 0; i < 4; i++) { results.PerformLayout(); foreach (var c in Descendants(results)) c.PerformLayout(); }
            CheckButtons(results);
            Console.WriteLine($"SPC PASS scale {scale}: configuration, rules, results, both charts, footer and resizing");
        }
    }
    private static void Prepare(Form form, float scale)
    {
        _ = form.Handle; foreach (var c in Descendants(form)) _ = c.Handle;
        form.Scale(new SizeF(scale, scale));
        for (int i = 0; i < 4; i++) { form.PerformLayout(); foreach (var c in Descendants(form)) c.PerformLayout(); }
    }
    private static void Preview(Form form, string name)
    {
        if (Environment.GetEnvironmentVariable("RESULTS_LAYOUT_PREVIEW") != "1") return;
        using var bitmap = new Bitmap(form.Width, form.Height); form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size));
        bitmap.Save(Path.Combine(AppContext.BaseDirectory, name));
    }
    private static void CheckButtons(Form form)
    {
        foreach (var c in Descendants(form).OfType<Button>())
        {
            var point = form.PointToClient(c.PointToScreen(Point.Empty));
            if (point.X < 0 || point.Y < 0 || point.X + c.Width > form.ClientSize.Width || point.Y + c.Height > form.ClientSize.Height)
                throw new Exception($"SPC button clipped: {form.Text}, {c.Text}, at {point}, size {c.Size}, client {form.ClientSize}");
        }
    }
    private static IEnumerable<Control> Descendants(Control root)
    { foreach (Control child in root.Controls) { yield return child; foreach (var c in Descendants(child)) yield return c; } }
}
