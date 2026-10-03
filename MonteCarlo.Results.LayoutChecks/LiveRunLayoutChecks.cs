using MonteCarlo.Excel;
using MonteCarlo.Core.Tests;

internal static class LiveRunLayoutChecks
{
    public static void Run()
    {
        Directory.CreateDirectory("artifacts/live-progress-layout");
        foreach (float scale in new[] { 1f, 1.25f, 1.5f })
        foreach (bool optimizer in new[] { false, true })
        {
            using var cts = new CancellationTokenSource();
            using var form = new LiveRunForm("Live run", cts, optimizer);
            _ = form.Handle; form.Scale(new SizeF(scale, scale)); form.Show();
            var output = OptimizationTests.Forecast(); output.TargetSettings = scale == 1f ? new(105, MonteCarlo.Core.TargetDirection.AtOrAbove, null) : scale == 1.25f ? null : new(null, MonteCarlo.Core.TargetDirection.AtOrAbove, null);
            var values = Enumerable.Range(0, 4096).Select(i => 100 + 10 * Math.Sin(i * .17)).ToArray();
            if (optimizer) form.CandidateStarted(new(1, [21], []), OptimizationTests.Problem());
            form.Trial(4096, new Dictionary<ForecastDefinition, double[]> { [output] = values });
            if (optimizer && !Descendants(form).OfType<TextBox>().Single().Text.Contains("Current Candidate: X: 21")) throw new Exception("Current candidate disappeared during trials");
            if (optimizer)
            {
                var run = new SimulationRunResult(values.Length, new() { [output] = values }, new());
                for (int i = 1; i <= 30; i++) form.Evaluation(new(i, 100, new([i], 100 + i / 3d, [], run)) { MonteCarloEvaluations = i, FeasibleEvaluations = i, RejectedBeforeSimulation = 5 }, OptimizationTests.Problem(), [output]);
                form.CandidateStarted(new(31, [22], []), OptimizationTests.Problem());
                form.Trial(4096, new Dictionary<ForecastDefinition, double[]> { [output] = values });
                if (!Descendants(form).OfType<TextBox>().Single().Text.Contains("Rejected before simulation: 5")) throw new Exception("Counters disappeared during trials");
            }
            form.PerformLayout();
            var charts = Descendants(form).OfType<LiveChart>().ToArray();
            if (charts.Length != (optimizer ? 2 : 1) || charts.Any(c => c.Width < 160 || c.Height < 130)) throw new Exception("Live charts clipped.");
            if (optimizer)
            {
                var graph = charts.Single(c => c.Series);
                if (scale == 1f && graph.Target != 105 || scale == 1.25f && (graph.Target == null || graph.TargetLabel != "Initial P80") || scale == 1.5f && graph.Target != null)
                    throw new Exception("Optimizer target/default/No Target reference incorrect.");
            }
            using var image = new Bitmap(form.Width, form.Height); form.DrawToBitmap(image, new(Point.Empty, image.Size));
            image.Save($"artifacts/live-progress-layout/{(optimizer ? "optimizer" : "simulation")}-{scale}.png");
            foreach (var chart in charts)
            foreach (var samples in new[] { new[] { 0d }, new[] { double.MaxValue }, new[] { double.MaxValue, -double.MaxValue } })
            { chart.Values = samples; using var bmp = new Bitmap(chart.Width, chart.Height); chart.DrawToBitmap(bmp, new(Point.Empty, bmp.Size)); }
            if (optimizer)
            {
                form.Size = form.MinimumSize; form.PerformLayout();
                if (charts.Any(c => c.Height < 130 || c.Width < 160)) throw new Exception("Optimizer charts clipped at the minimum size");
            }
            form.Finish(); Console.WriteLine($"LIVE LAYOUT PASS optimizer={optimizer} scale={scale}");
        }
    }
    private static IEnumerable<Control> Descendants(Control parent)
    { foreach (Control control in parent.Controls) { yield return control; foreach (var child in Descendants(control)) yield return child; } }
}
