using System.Globalization;
using System.Text.RegularExpressions;
namespace MonteCarlo.Excel;

/// <summary>Bounded grid projection. Excel still verifies every generated candidate before simulation.</summary>
public sealed class OptimizationFeasibleSpace
{
    private readonly OptimizationProblem problem;
    private readonly IReadOnlyList<OptimizationConstraint> linear;
    public double[] Minimum { get; }
    public double[] Maximum { get; }
    public OptimizationFeasibleSpace(OptimizationProblem problem, IReadOnlyList<OptimizationConstraint>? guides = null)
    {
        this.problem = problem;
        linear = problem.Constraints.Where(c => c.Kind == OptimizationConstraintKind.Linear).Concat(guides ?? []).ToArray();
        Minimum = problem.Variables.Select(v => v.Minimum).ToArray(); Maximum = problem.Variables.Select(v => v.Snap(v.Maximum)).ToArray();
        for (int i = 0; i < Minimum.Length; i++)
        {
            var v = problem.Variables[i];
            foreach (var c in problem.Constraints.Where(c => c.Kind == OptimizationConstraintKind.DecisionVariable && c.SubjectId == v.Id))
            {
                if (c.EffectiveOperator != OptimizationConstraintOperator.LessThanOrEqual)
                    Minimum[i] = Math.Max(Minimum[i], v.Minimum + Math.Ceiling((c.Value - v.Minimum) / v.Step - 1e-10) * v.Step);
                if (c.EffectiveOperator != OptimizationConstraintOperator.GreaterThanOrEqual)
                    Maximum[i] = Math.Min(Maximum[i], v.Minimum + Math.Floor((c.Value - v.Minimum) / v.Step + 1e-10) * v.Step);
            }
        }
    }
    public double[] Random(Random random) => problem.Variables.Select((v, i) =>
        Minimum[i] + random.NextInt64((long)Math.Round((Maximum[i] - Minimum[i]) / v.Step) + 1) * v.Step).ToArray();
    public double[] Repair(double[] source, Random random)
    {
        var values = source.Select((x, i) => Math.Clamp(problem.Variables[i].Snap(x), Minimum[i], Maximum[i])).ToArray();
        for (int pass = 0; pass < 32; pass++)
        {
            bool changed = false;
            foreach (var c in linear)
            {
                double actual = Value(c, problem.Variables, values);
                if (OptimizationService.EvaluateConstraint(c, actual).Met) continue;
                var terms = c.Terms!.OrderBy(_ => random.Next()).ToArray();
                foreach (var term in terms)
                {
                    int i = Index(problem.Variables, term.VariableId); double old = values[i];
                    double next = Math.Clamp(problem.Variables[i].Snap(old + (c.Value - actual) / term.Coefficient), Minimum[i], Maximum[i]);
                    values[i] = next; double updated = Value(c, problem.Variables, values); values[i] = old;
                    if (c.EffectiveOperator != OptimizationConstraintOperator.Equal && !OptimizationService.EvaluateConstraint(c, updated).Met)
                    {
                        double direction = c.EffectiveOperator == OptimizationConstraintOperator.LessThanOrEqual ? -1 : 1;
                        next = Math.Clamp(problem.Variables[i].Snap(next + direction * Math.Sign(term.Coefficient) * problem.Variables[i].Step), Minimum[i], Maximum[i]);
                        values[i] = next; updated = Value(c, problem.Variables, values); values[i] = old;
                    }
                    if (OptimizationService.EvaluateConstraint(c, updated).Violation < OptimizationService.EvaluateConstraint(c, actual).Violation)
                    { values[i] = next; actual = updated; changed = true; }
                    if (OptimizationService.EvaluateConstraint(c, actual).Met) break;
                }
            }
            if (!changed || linear.All(c => OptimizationService.EvaluateConstraint(c, Value(c, problem.Variables, values)).Met)) break;
        }
        return values;
    }
    private static int Index(IReadOnlyList<DecisionVariable> variables, string id) => variables.Select((v, i) => (v, i)).Single(x => x.v.Id == id).i;
    public static double Value(OptimizationConstraint c, IReadOnlyList<DecisionVariable> variables, double[] values) =>
        c.Terms!.Sum(t => t.Coefficient * values[Index(variables, t.VariableId)]);

    // Deliberately conservative: no UDFs, indirect references, nested formulas or unknown inputs.
    // Unsupported cell formulas remain black-box pre-simulation constraints.
    public static OptimizationConstraint? Guide(OptimizationConstraint constraint, ISimulationCell cell, IReadOnlyList<DecisionVariable> variables)
    {
        var direct = variables.FirstOrDefault(v => v.Sheet.Equals(constraint.Sheet, StringComparison.OrdinalIgnoreCase) &&
            v.Cell.Replace("$", "").Equals(constraint.Cell.Replace("$", ""), StringComparison.OrdinalIgnoreCase));
        if (direct != null) return constraint with { Kind = OptimizationConstraintKind.Linear, Terms = [new(direct.Id, 1)] };
        var content = cell.Capture();
        if (!content.IsFormula || content.Value is not string formula) return null;
        try
        {
            string text = formula.Trim().TrimStart('=');
            if (text.Contains("SUM", StringComparison.OrdinalIgnoreCase) && !Regex.IsMatch(text, @"^SUM\([^()]*\)$", RegexOptions.IgnoreCase)) return null;
            text = Regex.Replace(text, @"SUM\(([^()]*)\)", m =>
            {
                var parts = new List<string>();
                foreach (string item in m.Groups[1].Value.Split(','))
                {
                    var range = item.Trim().Split(':');
                    if (range.Length == 1) { parts.Add(range[0]); continue; }
                    var a = Regex.Match(range[0], @"^\$?([A-Za-z]+)\$?(\d+)$");
                    var b = Regex.Match(range[1], @"^\$?([A-Za-z]+)\$?(\d+)$");
                    if (!a.Success || !b.Success || !a.Groups[1].Value.Equals(b.Groups[1].Value, StringComparison.OrdinalIgnoreCase)) throw new FormatException();
                    int first = int.Parse(a.Groups[2].Value), last = int.Parse(b.Groups[2].Value);
                    if (last < first || last - first > 100) throw new FormatException();
                    for (int r = first; r <= last; r++) parts.Add(a.Groups[1].Value + r);
                }
                return string.Join("+", parts);
            }, RegexOptions.IgnoreCase);
            var terms = Parse(text, variables, token => variables.SingleOrDefault(v =>
                v.Sheet.Equals(constraint.Sheet, StringComparison.OrdinalIgnoreCase) && v.Cell.Replace("$", "").Equals(token.Replace("$", ""), StringComparison.OrdinalIgnoreCase)));
            return constraint with { Kind = OptimizationConstraintKind.Linear, Terms = terms };
        }
        catch (Exception ex) when (ex is ArgumentException or FormatException or OverflowException or InvalidOperationException) { return null; }
    }
    public static IReadOnlyList<OptimizationLinearTerm> Parse(string expression, IReadOnlyList<DecisionVariable> variables,
        Func<string, DecisionVariable?>? resolve = null)
    {
        resolve ??= token => variables.SingleOrDefault(v => v.Name.Equals(token, StringComparison.OrdinalIgnoreCase) || v.Id.Equals(token, StringComparison.OrdinalIgnoreCase));
        var result = new Dictionary<string, double>();
        string text = expression.Trim(); int offset = 0;
        foreach (Match match in Regex.Matches(text, @"(?:^|[+-])\s*(?:(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\s*\*\s*)?[^+\-*]+"))
        {
            if (match.Index != offset) throw new ArgumentException("Linear expression: use variable names with +, - and numeric coefficient * variable.");
            offset += match.Length;
            string term = match.Value.Trim(); double sign = 1;
            if (term.StartsWith('-')) { sign = -1; term = term[1..].Trim(); } else if (term.StartsWith('+')) term = term[1..].Trim();
            var parts = term.Split('*'); double coefficient = sign;
            if (parts.Length == 2) coefficient *= double.Parse(parts[0], CultureInfo.InvariantCulture);
            var variable = resolve(parts[^1].Trim()) ?? throw new ArgumentException($"Linear expression: unknown or ambiguous variable '{parts[^1].Trim()}'.");
            result[variable.Id] = result.GetValueOrDefault(variable.Id) + coefficient;
        }
        if (offset != text.Length || result.Count == 0 || result.Values.Any(x => !double.IsFinite(x))) throw new ArgumentException("Linear expression: enter a valid expression using decision variables.");
        var terms = result.Where(x => x.Value != 0).Select(x => new OptimizationLinearTerm(x.Key, x.Value)).ToArray();
        if (terms.Length == 0) throw new ArgumentException("Linear expression must contain a nonzero coefficient.");
        return terms;
    }
    public static string Expression(OptimizationConstraint c, IReadOnlyList<DecisionVariable> variables) => string.Concat(c.Terms!.Select((t, i) =>
        $"{(t.Coefficient < 0 ? " - " : i > 0 ? " + " : "")}{Math.Abs(t.Coefficient).ToString("G17", CultureInfo.InvariantCulture)}*{variables.FirstOrDefault(v => v.Id == t.VariableId)?.Name ?? t.VariableId}")).Trim();
}

public static class OptimizationProgressText
{
    public static string Format(OptimizationProgress progress, OptimizationProblem problem, bool deterministic)
    {
        string Values(OptimizationCandidate c) => string.Join("; ", problem.Variables.Select((v, i) => $"{v.Name}: {c.Values[i]:G6}"));
        var current = progress.Current ?? progress.Best;
        return $"Candidate Evaluations: {progress.Evaluations:N0} / {progress.MaximumEvaluations:N0}; Trials per Candidate: {(deterministic ? "N/A" : problem.Settings.Trials.ToString("N0"))}\r\n" +
            $"Current Candidate: {Values(current)}\r\nCurrent Objective: {(current.Evaluated ? current.Objective.ToString("G8") : "Not evaluated — rejected before simulation")}\r\n" +
            (progress.BestFeasible is { } best ? $"Best Feasible Candidate: {Values(best)}\r\nBest {problem.Objective.Metric}: {best.Objective:G8}\r\n" : "Best Feasible Candidate: none found\r\n") +
            $"Constraint Status: {(current.Feasible ? "Feasible" : "Infeasible")}; Feasible evaluated: {progress.FeasibleEvaluations:N0}; Rejected before simulation: {progress.RejectedBeforeSimulation:N0}; Monte Carlo evaluations: {progress.MonteCarloEvaluations:N0}\r\n" +
            $"Search: {progress.Phase}; radius: {progress.RadiusSteps:G6} grid steps";
    }
}
