using MonteCarlo.Excel;
using MonteCarlo.Core.Tests;
using ExcelDna.Integration;
using System.Drawing;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        if (args.Contains("--correlation-performance")) { CorrelationPerformanceChecks.Run(); return; }
        if (args.Contains("--live-scenarios")) { ScenarioLiveChecks.Run(); return; }
        if (args.Contains("--live-spc")) { SpcLiveChecks.Run(); return; }
        CorrelationLayoutChecks.Run();
        SpcLayoutChecks.Run();
        ScenarioLayoutChecks.Run();
        ExcelDnaUtil.Application = new PersistenceApplication { ActiveWorkbook = new PersistenceWorkbook() };
        foreach (var mode in new[] { SimulationSeedMode.Automatic, SimulationSeedMode.Fixed })
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        {
            var forecast = new ForecastDefinition { Name = "Test", SheetName = "Model", CellAddress = "B2", TargetSettings = new(296.68, MonteCarlo.Core.TargetDirection.AtOrBelow, null) };
            var second = new ForecastDefinition { Name = "Second", SheetName = "Model", CellAddress = "B3", TargetSettings = new(500, MonteCarlo.Core.TargetDirection.AtOrAbove, null) };
            var app = new PersistenceApplication { ActiveWorkbook = new PersistenceWorkbook() };
            app.ActiveWorkbook.Worksheets.Add().Name = "Model";
            ExcelDnaUtil.Application = app;
            SimulationModel.Clear(); SimulationModel.Forecasts.AddRange(new[] { forecast, second }); WorkbookPersistence.SaveModel();
            double[] samples = Enumerable.Range(1, 5000).Select(i => 113.0 + i * 264.0 / 5000).ToArray();
            var result = new SimulationRunResult(5000, new() { [forecast] = samples, [second] = samples.Select(v => v * 2).ToArray() },
                new() { ["Long assumption name for sensitivity verification"] = samples })
                { Settings = new(5000, mode, mode == SimulationSeedMode.Fixed ? int.MinValue : null), ActualSeedUsed = int.MinValue };
            using var form = new ResultsForm(result);
            _ = form.Handle;
            foreach (Control control in Descendants(form)) _ = control.Handle;
            form.Scale(new SizeF(scale, scale));
            Layout(form);
            var controls = Descendants(form).ToArray();
            var tabs = controls.OfType<TabControl>().Single(t => t.Name == "ResultsTabs");
            if (!tabs.TabPages.Cast<TabPage>().Select(p => p.Text).SequenceEqual(new[] { "Forecast", "Sensitivity", "Statistics" }) || tabs.SelectedIndex != 0)
                throw new Exception("Primary tabs/default incorrect");
            var forecastSelector = Field<ComboBox>(form, "cmbForecast");
            if (AncestorTab(forecastSelector) != null) throw new Exception("Forecast selector must be shared");
            if (AncestorTab(Field<Label>(form, "lblP50"))?.Text != "Statistics" || AncestorTab(Field<Label>(form, "lblP80"))?.Text != "Statistics")
                throw new Exception("Percentiles remain headline KPIs");
            if (AncestorTab(Field<Label>(form, "lblSuccess"))?.Text != "Forecast" || Field<Label>(form, "lblSuccess").Font.Size <= Field<Label>(form, "lblTargetSummary").Font.Size)
                throw new Exception("Probability hierarchy lost");
            var panel = (Panel)typeof(ResultsForm).GetField("histogramPanel", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.GetValue(form)!;
            using var bitmap = new Bitmap(Math.Max(1, panel.Width), Math.Max(1, panel.Height));
            bitmap.SetResolution(96 * scale, 96 * scale);
            using var graphics = Graphics.FromImage(bitmap);
            using var markerFont = new Font("Segoe UI", 8, FontStyle.Bold);
            int top = 8 + 3 * ((int)Math.Ceiling(markerFont.GetHeight(graphics)) + 8);
            int plotHeight = panel.ClientSize.Height - top - 55;
            Console.WriteLine($"GEOMETRY {mode} {scale}: panel={panel.ClientSize}, top={top}, plotHeight={plotHeight}, samples={result.ForecastResults[0].Values.Length}");

            if (plotHeight < 30) throw new Exception("Insufficient drawable plot height");
            var data = result.ForecastResults[0];
            if (data.Values.Length != 5000 || !double.IsFinite(data.Minimum) || !double.IsFinite(data.Maximum) || data.Maximum <= data.Minimum)
                throw new Exception("Invalid chart data");
            int binCount = Math.Min(30, Math.Max(15, (int)Math.Sqrt(data.Trials)));
            var bins = new int[binCount];
            foreach (double value in data.Values)
                bins[Math.Clamp((int)((value - data.Minimum) / ((data.Maximum - data.Minimum) / binCount)), 0, binCount - 1)]++;
            if (bins.Sum() != data.Trials || bins.Max() <= 0) throw new Exception("Empty bins");
            var plot = new RectangleF(55, top, panel.ClientSize.Width - 80, plotHeight);
            float barWidth = plot.Width / binCount;
            foreach (int count in bins.Where(n => n > 0))
                if (barWidth <= 2 || (float)count / bins.Max() * plotHeight <= 0) throw new Exception("Invisible bar");
            foreach (double marker in new[] { data.P50, data.P80, 296.68 })
            {
                double x = plot.Left + (marker - data.Minimum) / (data.Maximum - data.Minimum) * plot.Width;
                if (!double.IsFinite(x) || x < plot.Left || x > plot.Right) throw new Exception("Marker outside plot");
            }
            var selector = (ComboBox)typeof(ResultsForm).GetField("cmbChartView", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.GetValue(form)!;
            foreach (int view in new[] { 0, 1 })
            {
                selector.SelectedIndex = view;
                var paint = typeof(ResultsForm).GetMethod("HistogramPanel_Paint", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
                paint.Invoke(form, new object[] { panel, new PaintEventArgs(graphics, panel.ClientRectangle) });
                int chartPixels = 0, targetPixels = 0, percentilePixels = 0;
                for (int y = top; y < panel.ClientSize.Height - 55; y++)
                for (int x = 55; x < panel.ClientSize.Width - 25; x++)
                {
                    int color = bitmap.GetPixel(x, y).ToArgb();
                    if (color == SystemColors.Highlight.ToArgb()) chartPixels++;
                    if (color == Color.DarkOrange.ToArgb()) targetPixels++;
                    if (color == SystemColors.ControlDarkDark.ToArgb()) percentilePixels++;
                }
                // Presence checks, not pixel-perfect snapshots: axes alone cannot pass.
                if (chartPixels == 0 || targetPixels == 0 || percentilePixels == 0)
                    throw new Exception($"Blank/incomplete view {view}: chart={chartPixels}, target={targetPixels}, percentiles={percentilePixels}");
                Console.WriteLine($"RENDER PASS view={view}, bins={bins.Length}, populated={bins.Count(n => n > 0)}");
            }            tabs.SelectedIndex = 1; Layout(form);
            var tornado = Field<Panel>(form, "tornadoPanel");
            if (AncestorTab(tornado)?.Text != "Sensitivity" || tornado.Width < 200 || tornado.Height < 200)
                throw new Exception("Sensitivity lacks room");
            tabs.SelectedIndex = 2; Layout(form);
            foreach (string name in new[] { "Distribution Statistics", "Simulation Run", "Trials", "Mean", "Std Dev", "Minimum", "Maximum", "P10", "P50", "P80", "P90", "Seed", "Seed Mode" })
                if (!Descendants(tabs.SelectedTab!).OfType<Label>().Any(x => x.Text == name)) throw new Exception("Missing " + name);
            foreach (Label label in Descendants(tabs.SelectedTab!).OfType<Label>())
            {
                if (!label.Parent!.ClientRectangle.Contains(label.Bounds)) throw new Exception("Outside statistics: " + label.Text);
                var preferred = label.GetPreferredSize(Size.Empty);
                if (label.Height < preferred.Height || label.Width < preferred.Width) throw new Exception($"Clipped {label.Text}: {label.Size} needs {preferred}");
            }
            foreach (string text in new[] { "Export Report", "Close" })
            {
                var button = controls.OfType<Button>().Single(x => x.Text == text);
                var rect = new Rectangle(Position(button, form), button.Size);
                var panelRect = new Rectangle(Position(tabs, form), tabs.Size);
                if (AncestorTab(button) != null || !form.ClientRectangle.Contains(rect) || panelRect.IntersectsWith(rect)) throw new Exception("Footer overlap " + text);
            }
            if (!Descendants(tabs.SelectedTab!).OfType<Label>().Any(x => x.Text == "-2147483648")) throw new Exception("Missing seed");
            // Shared selector refreshes the same forecast object used by all renderers/statistics.
            forecastSelector.SelectedIndex = 1;
            if (!ReferenceEquals(Field<ForecastRunResult>(form, "currentForecastResult"), result.ForecastResults[1])) throw new Exception("Stale forecast");
            var format = typeof(ResultsForm).GetMethod("FormatValue", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
            string Format(double v) => (string)format.Invoke(null, new object[] { v })!;
            if (Field<Label>(form, "lblP50").Text != Format(result.ForecastResults[1].P50)) throw new Exception("Statistics did not refresh");
            if (Field<TextBox>(form, "txtTarget").Text != Format(500)) throw new Exception("Decision target did not refresh");
            forecastSelector.SelectedIndex = 0;
            tabs.SelectedIndex = 0; Layout(form);
            var direction = Field<ComboBox>(form, "cmbSuccessDirection");
            direction.SelectedIndex = 1;
            if (!Field<Label>(form, "lblTargetSummary").Text.Contains("≥")) throw new Exception("Direction did not refresh");
            Field<TextBox>(form, "txtTarget").Text = "250";
            typeof(ResultsForm).GetMethod("CalculateProbability", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .Invoke(form, new object?[] { null, EventArgs.Empty });
            if (Field<double?>(form, "currentTarget") != 250) throw new Exception("Target calculator failed");
            Field<TextBox>(form, "txtConfidence").Text = "80";
            typeof(ResultsForm).GetMethod("CalculateConfidenceTarget", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.Invoke(form, null);
            if (Field<double?>(form, "acceptedConfidence") != 80 || !Field<double?>(form, "currentTarget").HasValue) throw new Exception("Confidence calculator failed");
            var markerText = typeof(ResultsForm).GetMethod("MarkerText", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
            foreach (var pair in new[] { ("P50", data.P50), ("P80", data.P80), ("Target", 296.68) })
            {
                string label = (string)markerText.Invoke(null, new object?[] { pair.Item1, pair.Item2, pair.Item1 == "Target" ? MonteCarlo.Core.TargetDirection.AtOrBelow : null })!;
                if (!label.Contains(pair.Item1) || !label.Contains(Format(pair.Item2)) || (pair.Item1 == "Target" && !label.Contains("≤"))) throw new Exception("Marker value missing");
            }
            double exactTarget = Field<double?>(form, "currentTarget")!.Value;
            if (Field<TextBox>(form, "txtTarget").Text != Format(exactTarget)) throw new Exception("Raw target precision shown");
            typeof(ResultsForm).GetMethod("CalculateProbability", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
                .Invoke(form, new object?[] { null, EventArgs.Empty });
            if (Field<double?>(form, "currentTarget") != exactTarget) throw new Exception("Formatted display rounded the calculation");
            if (!Field<Label>(form, "lblCalculatedProbability").Text.Contains("%")) throw new Exception("Calculator probability result missing");
            var left = Descendants(form).Single(c => c.Name == "ForecastControls");
            if (left.Width <= 0 || Position(left, form).X >= Position(panel, form).X) throw new Exception("Columns misplaced");
            if (panel.Width <= left.Width || panel.Height < 300) throw new Exception("Chart space not preserved");
            var calculators = Descendants(left).OfType<TabControl>().Single();
            if (!calculators.TabPages.Cast<TabPage>().Select(p => p.Text).SequenceEqual(new[] { "Target → Probability", "Probability → Target" })) throw new Exception("Old terminology");
            foreach (TabPage page in calculators.TabPages)
            {
                calculators.SelectedTab = page; Layout(form);
                foreach (Control c in Descendants(page).Where(c => c is TextBox || c is Button || c is Label))
                {
                    if (!c.Parent!.ClientRectangle.Contains(c.Bounds)) throw new Exception("Calculator control clipped: " + c.Text);
                    if (!new Rectangle(Point.Empty, tabs.SelectedTab!.ClientSize).Contains(new Rectangle(Position(c, tabs.SelectedTab), c.Size))) throw new Exception("Calculator outside Forecast tab");
                }
            }
            typeof(ResultsForm).GetMethod("CalculateConfidenceTarget", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.Invoke(form, null);
            if (!Field<Label>(form, "lblRequiredTarget").Text.Contains(Format(Field<double?>(form, "currentTarget")!.Value))) throw new Exception("Calculator target result missing");            Layout(form);
            var resultLabel = Field<Label>(form, "lblRequiredTarget");
            var resultPage = AncestorTab(resultLabel)!;
            if (!resultPage.ClientRectangle.Contains(new Rectangle(Position(resultLabel, resultPage), resultLabel.Size))) throw new Exception("Calculated result clipped after refresh");
            if (resultLabel.Height < resultLabel.GetPreferredSize(new Size(resultLabel.Width, 0)).Height) throw new Exception("Calculated result text clipped");
            if (Field<Label>(form, "lblProbabilityCaption").Text != "Probability of achieving target") throw new Exception("Repeated target in summary");
            var labelBounds = typeof(ResultsForm).GetMethod("MarkerLabelBounds", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
            int labelRowHeight = (int)Math.Ceiling(markerFont.GetHeight(graphics)) + 8;
            foreach (float anchor in new[] { 55f, panel.Width / 2f, panel.Width - 25f })
            {
                RectangleF? previous = null;
                for (int row = 0; row < 3; row++)
                {
                    var bounds = (RectangleF)labelBounds.Invoke(null, new object?[] { panel.ClientRectangle, 180f, anchor, 4 + row * labelRowHeight, labelRowHeight, 55 })!;
                    if (!((RectangleF)panel.ClientRectangle).Contains(bounds) || anchor < bounds.Left || anchor > bounds.Right || bounds.Bottom >= top)
                        throw new Exception("Marker label not associated with its line/in label band");
                    if (previous.HasValue && previous.Value.IntersectsWith(bounds)) throw new Exception("Coincident marker labels overlap");
                    previous = bounds;
                }
            }            // Statistics tables and sensitivity retain the engine's existing values and ordering.
            if (data.Sensitivities.Count == 0 || tornado.ClientSize.Height < 200) throw new Exception("Missing sensitivity data/space");            if (Environment.GetEnvironmentVariable("RESULTS_LAYOUT_PREVIEW") == "1" && mode == SimulationSeedMode.Automatic && scale == 1)
            {
                using var preview = new Bitmap(form.Width, form.Height);
                form.DrawToBitmap(preview, new Rectangle(Point.Empty, preview.Size));
                Directory.CreateDirectory("artifacts/results-tabs-preview");
                preview.Save("artifacts/results-tabs-preview/forecast.png");
            }
            Console.WriteLine($"PASS {mode} scale {scale}: three tabs, statistics, shared selector/footer");        }
    }
    static T Field<T>(ResultsForm form, string name) => (T)typeof(ResultsForm).GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.GetValue(form)!;
    static TabPage? AncestorTab(Control control) { for (Control? c = control.Parent; c != null; c = c.Parent) if (c is TabPage page) return page; return null; }
    static Point Position(Control c, Control ancestor) { var p = Point.Empty; while(c != ancestor) { p.Offset(c.Location); c = c.Parent!; } return p; }
    static IEnumerable<Control> Descendants(Control c) => c.Controls.Cast<Control>().SelectMany(x => new[] { x }.Concat(Descendants(x)));
    static void Layout(Control c) { c.PerformLayout(); foreach (Control child in c.Controls) Layout(child); c.PerformLayout(); }
}
namespace MonteCarlo.Excel { public static class ReportExporter { public static void Export(SimulationRunResult result) => throw new NotSupportedException("Export not exercised by layout harness"); } }
