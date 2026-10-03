using MonteCarlo.Core;
namespace MonteCarlo.Excel;

public static class OptimizerExporter
{
    public static string Direction(TargetDirection d) => d == TargetDirection.AtOrAbove ? ">=" : "<=";
    public static string Operator(OptimizationConstraint c) => c.EffectiveOperator switch
    {
        OptimizationConstraintOperator.LessThanOrEqual => "<=", OptimizationConstraintOperator.GreaterThanOrEqual => ">=",
        OptimizationConstraintOperator.Equal => "=", _ => throw new ArgumentException("Unsupported constraint operator.")
    };
    public static string Objective(OptimizationProblem p, IReadOnlyList<ForecastDefinition> forecasts) =>
        $"{(p.Objective.Maximize ? "Maximize" : "Minimize")} {p.Objective.Metric} {Subject(p.Objective.ForecastId, p, forecasts)}" +
        (p.Objective.Metric == OptimizationMetric.Probability ? $" ({Direction(p.Objective.Direction)} {p.Objective.Target:G8})" : "");
    private static string Subject(string id, OptimizationProblem p, IReadOnlyList<ForecastDefinition> forecasts) =>
        p.Variables.FirstOrDefault(v => v.Id == id)?.Name ?? forecasts.FirstOrDefault(f => ScenarioSimulationService.Identity(f) == id)?.Name ?? id;
    public static string Constraint(OptimizationConstraint c, OptimizationProblem p, IReadOnlyList<ForecastDefinition> forecasts) =>
        (c.Kind == OptimizationConstraintKind.Linear ? OptimizationFeasibleSpace.Expression(c, p.Variables) : c.Kind == OptimizationConstraintKind.Cell ? c.CellDescription : $"{c.Kind} {Subject(c.SubjectId, p, forecasts)}") + (c.Kind == OptimizationConstraintKind.Probability ? $" ({Direction(c.TargetDirection)} {c.Target:G8})" : "") +
        $" {Operator(c)} {c.Value:G8}" + (c.Kind == OptimizationConstraintKind.Probability ? "%" : "");
    public static object?[,] Table(OptimizationResult r)
    {
        var p = r.Problem; var forecasts = r.ConfiguredForecasts.Count > 0 ? r.ConfiguredForecasts : r.Best.Run.ForecastResults.Select(f => f.Forecast).ToArray();
        bool deterministic = r.EvaluationMode == OptimizationEvaluationMode.Deterministic;
        var rows = new List<object?[]> {
            new object?[] { "MONTE CARLO OPTIMIZER", r.Best.Feasible ? "Best feasible solution found" : "No feasible solution found" },
            new object?[] { "Objective", Objective(p, forecasts) }, new object?[] { "Best feasible objective", r.BestFeasible?.Objective },
            new object?[] { "Search method", "Automatic: feasible grid / elite neighborhoods / progressive refinement / global exploration" },
            new object?[] { "Evaluations", r.Evaluations, "Maximum", p.Settings.MaximumEvaluations },
            new object?[] { "Rejected before simulation", r.RejectedBeforeSimulation },
            new object?[] { "Feasible candidates evaluated", r.FeasibleEvaluations },
            new object?[] { "Monte Carlo evaluations completed", r.MonteCarloEvaluations },
            new object?[] { "Evaluation mode", deterministic ? "Deterministic" : "Monte Carlo" },
            new object?[] { "Trials per evaluation", deterministic ? "N/A – Deterministic model" : (object)p.Settings.Trials }, new object?[] { deterministic ? "Search seed" : "Common seed", p.Settings.Seed },
            new object?[] { "Duration (seconds)", r.Duration.TotalSeconds }, new object?[] { "Interpretation", deterministic ? "Calculated forecast values; best observed candidate, no global optimality guarantee." : "Best observed candidate; stochastic estimates, no global optimality guarantee." },
            new object?[] { "DECISION VARIABLES", "Cell", "Original", r.BestFeasible != null ? "Optimized" : "Infeasible diagnostic candidate", "Change", "Minimum", "Maximum", "Step" }
        };
        for (int i = 0; i < p.Variables.Count; i++) { var v = p.Variables[i]; rows.Add([v.Name, v.Sheet + "!" + v.Cell, r.OriginalValues[i], r.Best.Values[i], r.Best.Values[i] - r.OriginalValues[i], v.Minimum, v.Maximum, v.Step]); }
        rows.Add(["CONSTRAINTS", "Result", "Status"]);
        foreach (var c in r.Best.Constraints) rows.Add([Constraint(c.Constraint, p, forecasts), c.Actual, c.Met ? "Met" : "Not met"]);
        rows.Add(["FORECAST RESULTS", "Mean", "P50", "P80", "Standard deviation", "Target probability (%)", "Target", "Direction"]);
        foreach (var f in r.Best.Run.ForecastResults)
        {
            bool selected = ScenarioSimulationService.Identity(f.Forecast) == p.Objective.ForecastId && p.Objective.Metric == OptimizationMetric.Probability;
            double? target = selected ? p.Objective.Target : f.Forecast.TargetSettings?.Target;
            var direction = selected ? p.Objective.Direction : f.Forecast.TargetSettings?.Direction ?? TargetDirection.AtOrAbove;
            rows.Add([ScenarioSimulationService.Display(f.Forecast), f.Mean, f.P50, f.P80, f.StandardDeviation,
                !deterministic && target.HasValue ? StatisticProbability(f, target.Value, direction) : null, target, target.HasValue ? Direction(direction) : null]);
        }
        var table = new object?[rows.Count, 8];
        for (int i = 0; i < rows.Count; i++) for (int j = 0; j < rows[i].Length; j++) table[i, j] = rows[i][j];
        return table;
    }
    private static double StatisticProbability(ForecastRunResult f, double target, TargetDirection d) => OptimizationService.Statistic(f, OptimizationMetric.Probability, target, d);
    public static void Export(object workbook, OptimizationResult result)
    {
        dynamic book = workbook; var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 1; i <= (int)book.Worksheets.Count; i++) names.Add((string)book.Worksheets[i].Name);
        string stem = "Optimizer_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"), name = stem;
        for (int i = 2; names.Contains(name); i++) name = stem + "_" + i;
        dynamic sheet = book.Worksheets.Add(); sheet.Name = name; var data = Table(result);
        // Per-cell text formatting prevents formula injection while preserving numeric report cells.
        for (int r = 0; r < data.GetLength(0); r++) for (int c = 0; c < data.GetLength(1); c++)
        {
            dynamic cell = sheet.Cells[r + 1, c + 1]; cell.NumberFormat = data[r, c] is string ? "@" : "General"; cell.Value2 = data[r, c];
        }
        sheet.Range["A:A"].Font.Bold = true; sheet.Range["A:H"].Columns.AutoFit(); sheet.Range["A:A"].ColumnWidth = 55;
        sheet.Range["A:A"].WrapText = true; sheet.Activate();
    }
}
