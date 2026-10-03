using System.Drawing.Drawing2D;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

internal sealed class LiveRunForm : Form
{
    private readonly TextBox status = new() { Dock = DockStyle.Top, Height = 54, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
    private readonly ComboBox forecast = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly LiveChart distribution = new() { Dock = DockStyle.Fill };
    private readonly LiveChart convergence = new() { Dock = DockStyle.Fill };
    private readonly CancellationTokenSource cancellation;
    private IReadOnlyDictionary<ForecastDefinition, double[]> samples = new Dictionary<ForecastDefinition, double[]>();
    private readonly List<double> best = [];
    private double? initialReference;
    private OptimizationCandidateStatus? current;
    private OptimizationProblem? currentProblem;
    private bool currentDeterministic;
    private string lastBest = "Best Feasible Candidate: none found";
    private string lastCounters = "Feasible evaluated: 0; Rejected before simulation: 0; Monte Carlo evaluations: 0";
    public void CandidateStarted(OptimizationCandidateStatus candidate, OptimizationProblem problem, bool deterministic = false)
    {
        current = candidate; currentProblem = problem; currentDeterministic = deterministic;
        status.Text = CandidateSummary("Pending") + "\r\nChecking constraints before simulation…";
        Refresh(); Application.DoEvents(); cancellation.Token.ThrowIfCancellationRequested();
    }
    private string CandidateSummary(string objective)
    {
        if (current == null || currentProblem == null) return "";
        return $"Candidate Evaluations: {current.Number:N0} / {currentProblem.Settings.MaximumEvaluations:N0}; Trials per Candidate: {(currentDeterministic ? "N/A" : currentProblem.Settings.Trials.ToString("N0"))}\r\nCurrent Candidate: " +
            string.Join("; ", currentProblem.Variables.Select((v, i) => $"{v.Name}: {current.Values[i]:G6}")) +
            $"\r\nCurrent Objective (preview): {objective}\r\n{lastBest}\r\n{lastCounters}\r\n" +
            $"Deterministic constraints: {(current.DeterministicConstraints.All(c => c.Met) ? "Passed; distribution constraints pending" : "Rejected")}";
    }
    public LiveRunForm(string title, CancellationTokenSource cancellation, bool optimizer = false)
    {
        this.cancellation = cancellation;
        if (optimizer) status.Height = 130;
        Text = title; ClientSize = new(720, optimizer ? 620 : 420); MinimumSize = new(520, optimizer ? 540 : 350);
        StartPosition = FormStartPosition.CenterScreen;
        var cancel = new Button { Text = "Cancel", Dock = DockStyle.Bottom, Height = 36 };
        cancel.Click += (_, _) => cancellation.Cancel();
        FormClosing += (_, e) => { e.Cancel = true; cancellation.Cancel(); };
        var plots = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = optimizer ? 2 : 1, ColumnCount = 1 };
        plots.RowStyles.Add(new(SizeType.Percent, optimizer ? 50 : 100));
        plots.Controls.Add(distribution, 0, 0);
        if (optimizer) { plots.RowStyles.Add(new(SizeType.Percent, 50)); plots.Controls.Add(convergence, 0, 1); }
        Controls.Add(plots); Controls.Add(forecast); Controls.Add(status); Controls.Add(cancel);
        forecast.SelectedIndexChanged += (_, _) => SelectForecast();
    }
    public void Trial(int completed, IReadOnlyDictionary<ForecastDefinition, double[]> values)
    {
        samples = values;
        if (forecast.Items.Count == 0)
        {
            forecast.Items.AddRange(values.Keys.Select(ScenarioSimulationService.Display).Cast<object>().ToArray());
            forecast.SelectedIndex = 0;
        }
        string preview = "Pending";
        if (currentProblem != null)
        {
            var entry = values.FirstOrDefault(x => ScenarioSimulationService.Identity(x.Key).Equals(currentProblem.Objective.ForecastId, StringComparison.OrdinalIgnoreCase));
            if (entry.Value is { Length: > 0 })
                preview = OptimizationService.Statistic(new ForecastRunResult(entry.Key, entry.Value, []), currentProblem.Objective.Metric,
                    currentProblem.Objective.Target, currentProblem.Objective.Direction).ToString("G8");
        }
        status.Text = (current == null ? "" : CandidateSummary(preview) + "\r\n") +
            $"Trials completed for current candidate: {completed:N0}\r\nOriginal model values will be restored when the run ends.";
        SelectForecast(); Refresh(); Application.DoEvents();
        cancellation.Token.ThrowIfCancellationRequested();
    }
    private void SelectForecast()
    {
        if (forecast.SelectedIndex < 0 || forecast.SelectedIndex >= samples.Count) return;
        var entry = samples.ElementAt(forecast.SelectedIndex);
        distribution.Values = entry.Value; distribution.Target = entry.Key.TargetSettings?.Target;
        distribution.Caption = "Live forecast distribution — sampled results"; distribution.Invalidate();
    }
    public void Evaluation(OptimizationProgress progress, OptimizationProblem problem, IReadOnlyList<ForecastDefinition> outputs)
    {
        if (progress.BestFeasible != null) best.Add(progress.BestFeasible.Objective);
        convergence.Values = best.ToArray(); convergence.Series = true;
        var output = outputs.Single(f => ScenarioSimulationService.Identity(f).Equals(problem.Objective.ForecastId, StringComparison.OrdinalIgnoreCase));
        if (best.Count == 1 && !initialReference.HasValue && progress.BestFeasible != null && output.TargetSettings == null && problem.Objective.Metric != OptimizationMetric.Probability)
            initialReference = progress.Best.Run.ForecastResults.Single(f => ScenarioSimulationService.Identity(f.Forecast).Equals(problem.Objective.ForecastId, StringComparison.OrdinalIgnoreCase)).P80;
        // A forecast value target and a probability percentage have different units.
        convergence.Target = problem.Objective.Metric == OptimizationMetric.Probability
            ? (output.TargetSettings?.Target == problem.Objective.Target && output.TargetSettings.Direction == problem.Objective.Direction ? output.TargetSettings.RequestedConfidence : null)
            : output.TargetSettings == null ? initialReference : output.TargetSettings.Target;
        convergence.TargetLabel = output.TargetSettings == null && problem.Objective.Metric != OptimizationMetric.Probability ? "Initial P80" : "Target";
        convergence.Caption = $"Best {problem.Objective.Metric} vs evaluation — {(progress.Best.Feasible ? "feasible" : "seeking feasibility")}" +
            (problem.Objective.Metric == OptimizationMetric.Probability ? $"; probability at target {problem.Objective.Target:G6} (%)" : "");
        status.Text = OptimizationProgressText.Format(progress, problem, progress.EvaluationMode == OptimizationEvaluationMode.Deterministic);
        lastBest = progress.BestFeasible is { } feasible
            ? "Best Feasible Candidate: " + string.Join("; ", problem.Variables.Select((v, i) => $"{v.Name}: {feasible.Values[i]:G6}")) + $"\r\nBest {problem.Objective.Metric}: {feasible.Objective:G8}"
            : "Best Feasible Candidate: none found";
        lastCounters = $"Feasible evaluated: {progress.FeasibleEvaluations:N0}; Rejected before simulation: {progress.RejectedBeforeSimulation:N0}; Monte Carlo evaluations: {progress.MonteCarloEvaluations:N0}";
        Refresh(); Application.DoEvents(); cancellation.Token.ThrowIfCancellationRequested();
    }
    public void Finish() { Hide(); }
}

internal sealed class LiveChart : Panel
{
    internal double[] Values = [];
    internal double? Target;
    internal bool Series;
    internal string Caption = "Waiting for trials…";
    internal string TargetLabel = "Target";
    public LiveChart() { DoubleBuffered = true; BackColor = Color.White; }
    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e); var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
        g.DrawString(Caption, Font, Brushes.Black, new RectangleF(8, 5, Width - 16, 35));
        if (Values.Length == 0 || Width < 160 || Height < 130) return;
        var r = new RectangleF(65, 45, Width - 90, Height - 85);
        double min = Values.Min(), max = Values.Max();
        if (Target.HasValue) { min = Math.Min(min, Target.Value); max = Math.Max(max, Target.Value); }
        if (min == max)
        {
            double padding = Math.Max(1, Math.Abs(min) * .01);
            if (double.IsFinite(min - padding)) min -= padding;
            if (double.IsFinite(max + padding)) max += padding;
        }
        // Normalize in scaled coordinates so extreme finite values cannot overflow the axis span.
        double scale = Math.Max(1, Math.Max(Math.Abs(min), Math.Abs(max)));
        double lo = min / scale, span = max / scale - lo;
        float Position(double v) => span > 0 ? (float)((v / scale - lo) / span) : .5f;
        using var pen = new Pen(Color.RoyalBlue, 2);
        g.DrawRectangle(Pens.Gray, r.X, r.Y, r.Width, r.Height);
        if (Series)
        {
            var points = Values.Select((v, i) => new PointF(r.Left + r.Width * i / Math.Max(1, Values.Length - 1), r.Bottom - r.Height * Position(v))).ToArray();
            for (int i = 1; i < points.Length; i++)
            { g.DrawLine(pen, points[i - 1], new PointF(points[i].X, points[i - 1].Y)); g.DrawLine(pen, new PointF(points[i].X, points[i - 1].Y), points[i]); }
            if (points.Length == 1) g.FillEllipse(Brushes.RoyalBlue, points[0].X - 3, points[0].Y - 3, 6, 6);
            g.DrawString($"{min:G5}", Font, Brushes.Black, 2, r.Bottom - 15);
            g.DrawString($"{max:G5}", Font, Brushes.Black, 2, r.Top);
            g.DrawString($"Evaluation 1 → {Values.Length:N0}", Font, Brushes.Black, r.Left, r.Bottom + 8);
        }
        else
        {
            var bins = new int[32];
            foreach (double value in Values) bins[Math.Clamp((int)(Position(value) * bins.Length), 0, bins.Length - 1)]++;
            int peak = Math.Max(1, bins.Max());
            var points = bins.Select((n, i) => new PointF(r.Left + (i + .5f) * r.Width / bins.Length, r.Bottom - r.Height * n / peak)).ToArray();
            for (int i = 0; i < bins.Length; i++) g.FillRectangle(Brushes.LightSteelBlue, r.Left + i * r.Width / bins.Length, points[i].Y, r.Width / bins.Length, r.Bottom - points[i].Y);
            g.DrawLines(pen, points);
            g.DrawString($"{min:G5}", Font, Brushes.Black, r.Left, r.Bottom + 8);
            string maximum = $"{max:G5}";
            g.DrawString(maximum, Font, Brushes.Black, r.Right - g.MeasureString(maximum, Font).Width, r.Bottom + 8);
            g.DrawString("Frequency", Font, Brushes.Black, 2, r.Top);
        }
        if (Target.HasValue)
        {
            using var targetPen = new Pen(Color.DarkOrange, 2) { DashStyle = DashStyle.Dash };
            if (Series) { float y = r.Bottom - r.Height * Position(Target.Value); g.DrawLine(targetPen, r.Left, y, r.Right, y); }
            else { float x = r.Left + r.Width * Position(Target.Value); g.DrawLine(targetPen, x, r.Top, x, r.Bottom); }
            g.DrawString($"{TargetLabel}: {Target.Value:G6}", Font, Brushes.DarkOrange, r.Right - 160, r.Top + 4);
        }
    }
}

