using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public enum OptimizationMetric { Mean, Median, Probability }
public enum OptimizationConstraintKind { DecisionVariable, Mean, Median, Probability, Cell, Linear }
public enum OptimizationConstraintOperator { FromDirection, LessThanOrEqual, GreaterThanOrEqual, Equal }
public sealed record DecisionVariable(string Id, string Name, string Sheet, string Cell, double Minimum, double Maximum, double Step, string CellLink = "")
{
    // The lattice is anchored at Minimum. An unaligned Maximum is not a permitted value.
    public long Levels => (long)Math.Floor((Maximum - Minimum) / Step + 1e-10);
    public double Snap(double value) => Math.Min(Maximum, Minimum + Math.Clamp(Math.Round((value - Minimum) / Step), 0, Levels) * Step);
}
public sealed record OptimizationObjective(string ForecastId, OptimizationMetric Metric = OptimizationMetric.Mean,
    bool Maximize = true, double Target = 0, TargetDirection Direction = TargetDirection.AtOrAbove);
public sealed record OptimizationConstraint(OptimizationConstraintKind Kind, string SubjectId, TargetDirection Direction,
    double Value, double Target = 0, TargetDirection TargetDirection = TargetDirection.AtOrAbove,
    string Sheet = "", string Cell = "", string Name = "", string CellLink = "",
    OptimizationConstraintOperator Operator = OptimizationConstraintOperator.FromDirection,
    IReadOnlyList<OptimizationLinearTerm>? Terms = null)
{
    // Missing Operator in V1 JSON preserves the original inclusive direction semantics.
    public OptimizationConstraintOperator EffectiveOperator => Operator == OptimizationConstraintOperator.FromDirection
        ? Direction == TargetDirection.AtOrAbove ? OptimizationConstraintOperator.GreaterThanOrEqual : OptimizationConstraintOperator.LessThanOrEqual : Operator;
    public string CellDescription => string.IsNullOrWhiteSpace(Name) ? $"{Sheet}!{Cell}" : $"{Name} ({Sheet}!{Cell})";
}
public sealed record OptimizationLinearTerm(string VariableId, double Coefficient);
public sealed record OptimizationSettings(int MaximumEvaluations = 250, int Trials = 1000, int Seed = 12345);
public sealed record OptimizationProblem(OptimizationObjective Objective, IReadOnlyList<DecisionVariable> Variables,
    IReadOnlyList<OptimizationConstraint> Constraints, OptimizationSettings Settings)
{
    public void Validate() => Validate(false);
    public void Validate(bool deterministic)
    {
        if (string.IsNullOrWhiteSpace(Objective.ForecastId)) throw new ArgumentException("Objective: select a configured forecast.");
        if (!Enum.IsDefined(Objective.Metric) || !Enum.IsDefined(Objective.Direction) || !double.IsFinite(Objective.Target))
            throw new ArgumentException("Objective: select a valid metric/direction and enter a finite target.");
        if (Variables.Count == 0) throw new ArgumentException("Decision variables: add at least one numeric input cell.");
        if (Variables.Count > 100) throw new ArgumentException("Decision variables: V1 supports at most 100 inputs.");
        if (Variables.Select(v => v.Id).Distinct(StringComparer.OrdinalIgnoreCase).Count() != Variables.Count)
            throw new ArgumentException("Decision variables: remove duplicate variable IDs.");
        foreach (var v in Variables)
            if (string.IsNullOrWhiteSpace(v.Name) || string.IsNullOrWhiteSpace(v.Sheet) || string.IsNullOrWhiteSpace(v.Cell) ||
                !double.IsFinite(v.Minimum) || !double.IsFinite(v.Maximum) || !double.IsFinite(v.Step) ||
                v.Minimum >= v.Maximum || v.Step <= 0 || v.Minimum + v.Step <= v.Minimum || !double.IsFinite((v.Maximum - v.Minimum) / v.Step) ||
                (v.Maximum - v.Minimum) / v.Step > 1e12 || v.Step > v.Maximum - v.Minimum)
                throw new ArgumentException($"{v.Name}: enter a valid cell, Minimum < Maximum and a positive Step no larger than the range (at most 10¹² steps).");
        if (Settings.MaximumEvaluations < 1 || Settings.MaximumEvaluations > 100000 || !deterministic && SimulationSettings.ValidateTrials(Settings.Trials) != null)
            throw new ArgumentException("Settings: enter 1–100,000 evaluations and 1–1,000,000 trials per evaluation.");
        foreach (var c in Constraints)
        {
            if (!Enum.IsDefined(c.Kind) || !Enum.IsDefined(c.Direction) || !Enum.IsDefined(c.TargetDirection) || !Enum.IsDefined(c.Operator) ||
                !double.IsFinite(c.Value) || !double.IsFinite(c.Target) || string.IsNullOrWhiteSpace(c.SubjectId))
                throw new ArgumentException("Constraint: select a subject/direction and enter finite values.");
            if (c.Kind == OptimizationConstraintKind.Cell && (string.IsNullOrWhiteSpace(c.Sheet) || string.IsNullOrWhiteSpace(c.Cell)))
                throw new ArgumentException("Constraint: select or enter a single numeric cell, including its worksheet.");
            if (c.Kind == OptimizationConstraintKind.Probability && (c.Value < 0 || c.Value > 100))
                throw new ArgumentException("Probability constraint: enter a requirement from 0 to 100 percent.");
            if (c.Kind == OptimizationConstraintKind.Linear && (c.Terms is not { Count: > 0 } ||
                c.Terms.Any(t => !double.IsFinite(t.Coefficient) || t.Coefficient == 0 || !Variables.Any(v => v.Id == t.VariableId)) ||
                c.Terms.Select(t => t.VariableId).Distinct().Count() != c.Terms.Count))
                throw new ArgumentException("Linear constraint: enter nonzero finite coefficients for existing decision variables, without duplicates.");
            if (c.Kind == OptimizationConstraintKind.DecisionVariable)
            {
                var v = Variables.SingleOrDefault(v => v.Id == c.SubjectId) ?? throw new ArgumentException("Constraint variable is missing. Select an existing variable.");
                if (c.EffectiveOperator != OptimizationConstraintOperator.LessThanOrEqual && c.Value > v.Snap(v.Maximum) ||
                    c.EffectiveOperator != OptimizationConstraintOperator.GreaterThanOrEqual && c.Value < v.Minimum)
                    throw new ArgumentException($"{v.Name}: the constraint is impossible within the bounds. Adjust the requirement or bounds.");
            }
        }
        foreach (var v in Variables)
        {
            double lo = v.Minimum, hi = v.Snap(v.Maximum);
            foreach (var c in Constraints.Where(c => c.Kind == OptimizationConstraintKind.DecisionVariable && c.SubjectId == v.Id))
            {
                if (c.EffectiveOperator != OptimizationConstraintOperator.LessThanOrEqual) lo = Math.Max(lo, c.Value);
                if (c.EffectiveOperator != OptimizationConstraintOperator.GreaterThanOrEqual) hi = Math.Min(hi, c.Value);
            }
            if (lo > hi || Math.Ceiling((lo - v.Minimum) / v.Step - 1e-10) > Math.Floor((hi - v.Minimum) / v.Step + 1e-10))
                throw new ArgumentException($"{v.Name}: constraints allow no stepped value. Adjust the requirements or step.");
        }
    }
}
public sealed record OptimizationConstraintResult(OptimizationConstraint Constraint, double Actual, bool Met, double Violation);
public sealed record OptimizationCandidate(double[] Values, double Objective, IReadOnlyList<OptimizationConstraintResult> Constraints, SimulationRunResult Run)
{
    public bool Evaluated { get; init; } = true;
    public bool Feasible => Constraints.All(c => c.Met);
    public double Violation => Constraints.Sum(c => c.Violation);
}
public sealed record OptimizationProgress(int Evaluations, int MaximumEvaluations, OptimizationCandidate Best)
{
    public OptimizationCandidate? Current { get; init; }
    public int RejectedBeforeSimulation { get; init; }
    public int FeasibleEvaluations { get; init; }
    public int MonteCarloEvaluations { get; init; }
    public string Phase { get; init; } = "Exploration";
    public double RadiusSteps { get; init; }
    public OptimizationEvaluationMode EvaluationMode { get; init; } = OptimizationEvaluationMode.MonteCarlo;
    public OptimizationCandidate? BestFeasible => Best.Feasible && Best.Evaluated ? Best : null;
}
public sealed record OptimizationResult(OptimizationProblem Problem, double[] OriginalValues, OptimizationCandidate Best,
    int Evaluations, TimeSpan Duration)
{
    public OptimizationEvaluationMode EvaluationMode { get; init; } = OptimizationEvaluationMode.MonteCarlo;
    public IReadOnlyList<ForecastDefinition> ConfiguredForecasts { get; init; } = [];
    public int RejectedBeforeSimulation { get; init; }
    public int MonteCarloEvaluations { get; init; }
    public int FeasibleEvaluations { get; init; }
    public OptimizationCandidate? BestFeasible => Best.Feasible && Best.Evaluated ? Best : null;
    public IReadOnlyList<OptimizationEvaluation> History { get; init; } = [];
    public IReadOnlyList<SimulationCellContent> OriginalContents { get; init; } = [];
}
public sealed record OptimizationEvaluation(int Number, double[] Values, double Objective, bool Feasible, double Violation)
{
    public bool Evaluated { get; init; } = true;
    public IReadOnlyList<OptimizationConstraintResult> Constraints { get; init; } = [];
    public double? BestFeasibleObjective { get; init; }
    public string Phase { get; init; } = "Exploration";
    public double RadiusSteps { get; init; }
}
public sealed record OptimizationCandidateStatus(int Number, double[] Values, IReadOnlyList<OptimizationConstraintResult> DeterministicConstraints);

public enum OptimizationEvaluationMode { Deterministic, MonteCarlo }

/// <summary>Evaluates forecasts after the caller has written and recalculated a candidate.</summary>
public interface IOptimizationCandidateEvaluator
{
    SimulationRunResult Evaluate(CancellationToken cancellation);
}

public sealed class DeterministicCandidateEvaluator : IOptimizationCandidateEvaluator
{
    private readonly Dictionary<ForecastDefinition, ISimulationCell> outputs;
    private readonly Action<int, IReadOnlyDictionary<ForecastDefinition, double[]>>? live;
    public DeterministicCandidateEvaluator(ISimulationWorkbook workbook, IReadOnlyList<ForecastDefinition> forecasts,
        Action<int, IReadOnlyDictionary<ForecastDefinition, double[]>>? live = null)
    {
        this.live = live;
        outputs = forecasts.ToDictionary(f => f, f =>
        {
            ConstraintCellSelection.Absolute(f.CellAddress);
            var cell = workbook.Resolve(f.SheetName, f.CellAddress);
            ConstraintCellSelection.Number(cell.ReadValue(), SimulationValidation.Location("Forecast", f.Name, f.SheetName, f.CellAddress));
            return cell;
        });
    }
    public SimulationRunResult Evaluate(CancellationToken cancellation)
    {
        cancellation.ThrowIfCancellationRequested();
        var samples = outputs.ToDictionary(x => x.Key, x => new[] { ConstraintCellSelection.Number(x.Value.ReadValue(),
            SimulationValidation.Location("Forecast", x.Key.Name, x.Key.SheetName, x.Key.CellAddress)) });
        live?.Invoke(1, samples.ToDictionary(x => x.Key, x => x.Value.ToArray()));
        cancellation.ThrowIfCancellationRequested();
        return new SimulationRunResult(1, samples, []);
    }
}

public sealed class MonteCarloCandidateEvaluator(OptimizationSettings settings, IReadOnlyList<AssumptionDefinition> inputs,
    IReadOnlyList<ForecastDefinition> outputs, ISimulationWorkbook workbook, IReadOnlyList<AssumptionCorrelation>? correlations,
    Action<int, IReadOnlyDictionary<ForecastDefinition, double[]>>? live) : IOptimizationCandidateEvaluator
{
    public SimulationRunResult Evaluate(CancellationToken cancellation)
    {
        var outcome = SimulationExecution.Run(new SimulationSettings(settings.Trials, SimulationSeedMode.Fixed, settings.Seed),
            inputs, outputs, workbook, cancellation: cancellation, correlations: correlations, liveProgress: live);
        cancellation.ThrowIfCancellationRequested();
        if (!outcome.Succeeded) throw new SimulationRunException(outcome);
        return outcome.Result!;
    }
}

public interface IOptimizationAlgorithm
{
    OptimizationCandidate Search(OptimizationProblem problem, double[] initial, Func<double[], OptimizationCandidate> evaluate,
        CancellationToken cancellation, Action<OptimizationProgress>? progress, out int evaluations);
}

/// <summary>Exhaust small lattices; otherwise seeded global exploration with shrinking coordinate polls.</summary>
public sealed class AutomaticOptimizationAlgorithm(IReadOnlyList<OptimizationConstraint>? guides = null) : IOptimizationAlgorithm
{
    public static bool Better(OptimizationCandidate a, OptimizationCandidate b, bool maximize) =>
        a.Feasible != b.Feasible ? a.Feasible : !a.Feasible && a.Violation != b.Violation ? a.Violation < b.Violation :
        maximize ? a.Objective > b.Objective : a.Objective < b.Objective;

    public OptimizationCandidate Search(OptimizationProblem p, double[] initial, Func<double[], OptimizationCandidate> evaluate,
        CancellationToken cancellation, Action<OptimizationProgress>? progress, out int evaluations)
    {
        var random = new Random(p.Settings.Seed); var seen = new HashSet<string>();
        OptimizationCandidate? best = null; int count = 0;
        var space = new OptimizationFeasibleSpace(p, guides);
        var elite = new List<OptimizationCandidate>();
        string phase = "Exploration"; double radius = 0;
        void Visit(double[] values)
        {
            cancellation.ThrowIfCancellationRequested();
            if (count >= p.Settings.MaximumEvaluations) return;
            values = space.Repair(values, random);
            string key = string.Join("|", values.Select(x => x.ToString("R", System.Globalization.CultureInfo.InvariantCulture)));
            if (!seen.Add(key)) return;
            var candidate = evaluate(values); count++;
            if (best == null || Better(candidate, best, p.Objective.Maximize)) best = candidate;
            elite.Add(candidate with { Run = new SimulationRunResult(1, [], []) });
            elite.Sort((a, b) => Better(a, b, p.Objective.Maximize) ? -1 : Better(b, a, p.Objective.Maximize) ? 1 : 0);
            if (elite.Count > 6) elite.RemoveAt(elite.Count - 1);
            progress?.Invoke(new(count, p.Settings.MaximumEvaluations, best) { Current = candidate, Phase = phase, RadiusSteps = radius });
        }
        double size = p.Variables.Aggregate(1d, (n, v) => n * (v.Levels + 1));
        if (size <= p.Settings.MaximumEvaluations)
        {
            var values = new double[p.Variables.Count];
            void Enumerate(int d)
            {
                if (d == values.Length) { Visit(values.ToArray()); return; }
                var v = p.Variables[d];
                for (long i = 0; i <= v.Levels; i++) { values[d] = v.Minimum + i * v.Step; Enumerate(d + 1); }
            }
            Enumerate(0);
        }
        else
        {
            Visit(initial); Visit(space.Minimum); Visit(space.Maximum);
            for (int i = 0; i < p.Variables.Count; i++)
            {
                var corner = space.Minimum.ToArray(); corner[i] = space.Maximum[i]; Visit(corner);
                corner = space.Maximum.ToArray(); corner[i] = space.Minimum[i]; Visit(corner);
            }
            for (int i = 0; i < Math.Min(12, p.Settings.MaximumEvaluations / 8); i++) Visit(space.Random(random));
            int attempts = 0;
            int local = 0; int dimensions = p.Variables.Count;
            double largest = p.Variables.Max(v => v.Levels);
            int scales = Math.Max(1, (int)Math.Ceiling(Math.Log2(Math.Max(1, largest))) + 1);
            while (count < p.Settings.MaximumEvaluations && attempts++ < p.Settings.MaximumEvaluations * 40)
            {
                if (attempts % 8 == 0)
                { phase = "Global exploration"; radius = 0; Visit(space.Random(random)); }
                else
                {
                    phase = "Elite neighborhood";
                    var parent = local % 8 == 7 ? elite[random.Next(elite.Count)] : best!;
                    var values = parent.Values.ToArray(); int d = local / 2 % dimensions;
                    int level = local / (2 * dimensions) % scales;
                    radius = Math.Max(1, Math.Floor(p.Variables[d].Levels / Math.Pow(2, level + 1)));
                    values[d] += (local % 2 == 0 ? 1 : -1) * radius * p.Variables[d].Step;
                    local++; Visit(values);
                }
            }
        }
        evaluations = count; return best!;
    }
}

public static class OptimizationService
{
    public static void ValidateMode(OptimizationProblem p, bool deterministic)
    {
        if (deterministic && (p.Objective.Metric == OptimizationMetric.Probability || p.Constraints.Any(c => c.Kind == OptimizationConstraintKind.Probability)))
            throw new ArgumentException("Deterministic optimization has no uncertainty. Choose Mean or Median and replace probability constraints with forecast-value or cell constraints.");
    }
    public static IReadOnlyDictionary<OptimizationConstraint, ISimulationCell> ValidateConstraintCells(
        IReadOnlyList<OptimizationConstraint> constraints, ISimulationWorkbook workbook)
    {
        var cells = new Dictionary<OptimizationConstraint, ISimulationCell>();
        foreach (var c in constraints.Where(c => c.Kind == OptimizationConstraintKind.Cell))
        {
            try
            {
                ConstraintCellSelection.Absolute(c.Cell);
                var cell = workbook.Resolve(c.Sheet, c.Cell);
                ConstraintCellSelection.Number(cell.ReadValue(), c.CellDescription);
                cells[c] = cell;
            }
            catch (SimulationCellAccessException ex) { throw new ArgumentException($"Constraint \"{c.CellDescription}\": the reference is unavailable. Select one existing cell in this workbook.", ex); }
        }
        return cells;
    }
    public static OptimizationConstraintResult EvaluateConstraint(OptimizationConstraint c, double actual)
    {
        if (!double.IsFinite(actual)) throw new ArgumentException($"Constraint \"{c.CellDescription}\": the calculated value is not finite.");
        double gap = c.EffectiveOperator switch
        {
            OptimizationConstraintOperator.GreaterThanOrEqual => c.Value - actual,
            OptimizationConstraintOperator.LessThanOrEqual => actual - c.Value,
            OptimizationConstraintOperator.Equal => Math.Abs(actual - c.Value) - 1e-9 * Math.Max(1, Math.Abs(c.Value)),
            _ => throw new ArgumentException("Constraint: select a supported operator.")
        };
        return new(c, actual, gap <= 0, Math.Max(0, gap) / Math.Max(1, Math.Abs(c.Value)));
    }
    public static double Statistic(ForecastRunResult f, OptimizationMetric metric, double target, TargetDirection direction) => metric switch
    {
        OptimizationMetric.Mean => f.Mean, OptimizationMetric.Median => f.P50,
        OptimizationMetric.Probability => 100 * (direction == TargetDirection.AtOrAbove ? f.ProbabilityGreaterThanOrEqual(target) : f.ProbabilityLessThanOrEqual(target)),
        _ => throw new ArgumentException("Unknown objective metric.")
    };
    public static ISimulationCell[] ValidateCells(OptimizationProblem p, ISimulationWorkbook workbook,
        IReadOnlyList<AssumptionDefinition> inputs, IReadOnlyList<ForecastDefinition> outputs)
    {
        var cells = p.Variables.Select(v =>
        {
            try { return workbook.Resolve(v.Sheet, v.Cell); }
            catch (SimulationCellAccessException ex) { throw new ArgumentException($"{v.Name} ({v.Sheet}!{v.Cell}): the cell is unavailable. Edit the variable and select one existing numeric cell.", ex); }
        }).ToArray();
        if (cells.Select(c => c.Identity).Distinct(StringComparer.OrdinalIgnoreCase).Count() != cells.Length)
            throw new ArgumentException("Decision variables: duplicate cells. Keep one variable per cell.");
        var reserved = inputs.Select(a => workbook.Resolve(a.SheetName, a.CellAddress).Identity)
            .Concat(outputs.Select(f => workbook.Resolve(f.SheetName, f.CellAddress).Identity)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < cells.Length; i++)
        {
            var c = cells[i];
            if (c.InputProblem != null) throw new ArgumentException($"{p.Variables[i].Name} ({c.Identity}): {c.InputProblem} Select an editable numeric input.");
            if (reserved.Contains(c.Identity)) throw new ArgumentException($"{c.Identity}: a decision variable cannot also be an assumption or forecast. Select a controllable input.");
            Numeric(c.ReadValue(), c.Identity);
        }
        return cells;
    }
    private static double Numeric(object? value, string location)
    {
        if (value is not (double or float or int or long or short or byte or decimal) || !double.IsFinite(Convert.ToDouble(value)))
            throw new ArgumentException($"{location}: current value must be numeric and finite. Enter a numeric input.");
        return Convert.ToDouble(value);
    }
    public static OptimizationResult Run(OptimizationProblem p, IReadOnlyList<AssumptionDefinition> inputs,
        IReadOnlyList<ForecastDefinition> outputs, Func<IScenarioSandbox> createSandbox, CancellationToken cancellation = default,
        Action<OptimizationProgress>? progress = null, IReadOnlyList<AssumptionCorrelation>? correlations = null, IOptimizationAlgorithm? algorithm = null,
        Action<int, IReadOnlyDictionary<ForecastDefinition, double[]>>? liveProgress = null,
        Action<OptimizationCandidateStatus>? candidateStarted = null)
    {
        bool deterministic = inputs.Count == 0;
        p.Validate(deterministic); ValidateMode(p, deterministic);
        p = p with { Variables = p.Variables.ToArray(), Constraints = p.Constraints.Select(c => c with { Terms = c.Terms?.ToArray() }).ToArray() };
        inputs = inputs.Select(ScenarioSimulationService.Clone).ToArray(); outputs = outputs.Select(ScenarioSimulationService.Clone).ToArray();
        if (!deterministic)
        {
            var validation = SimulationValidation.ValidateModel(p.Settings.Trials, inputs, outputs);
            if (!validation.IsValid) throw new ArgumentException(validation.UserMessage);
        }
        else if (outputs.Count == 0) throw new ArgumentException("Objective: define at least one configured forecast.");
        ForecastDefinition Find(string id) => outputs.SingleOrDefault(f => ScenarioSimulationService.Identity(f).Equals(id, StringComparison.OrdinalIgnoreCase))
            ?? throw new ArgumentException($"Forecast {id} is missing. Select a configured forecast.");
        Find(p.Objective.ForecastId);
        foreach (var c in p.Constraints.Where(c => c.Kind is not (OptimizationConstraintKind.DecisionVariable or OptimizationConstraintKind.Cell or OptimizationConstraintKind.Linear))) Find(c.SubjectId);
        cancellation.ThrowIfCancellationRequested(); var timer = System.Diagnostics.Stopwatch.StartNew();
        using var sandbox = createSandbox(); var workbook = sandbox.Workbook;
        var cells = ValidateCells(p, workbook, inputs, outputs);
        var constraintCells = ValidateConstraintCells(p.Constraints, workbook);
        var guides = constraintCells.Select(x => OptimizationFeasibleSpace.Guide(x.Key, x.Value, p.Variables)).Where(x => x != null).Cast<OptimizationConstraint>().ToArray();
        IOptimizationCandidateEvaluator evaluator = deterministic
            ? new DeterministicCandidateEvaluator(workbook, outputs, liveProgress)
            : new MonteCarloCandidateEvaluator(p.Settings, inputs, outputs, workbook, correlations, liveProgress);
        var contents = cells.Select(c => c.Capture()).ToArray(); var original = cells.Select(c => Numeric(c.ReadValue(), c.Identity)).ToArray();
        bool screen = workbook.ScreenUpdating;
        try
        {
            if (deterministic && liveProgress != null) workbook.ScreenUpdating = true;
            var history = new List<OptimizationEvaluation>();
            int rejected = 0, simulated = 0, feasible = 0; double? bestObjective = null;
            OptimizationCandidate Record(OptimizationCandidate candidate)
            {
                if (candidate.Feasible && candidate.Evaluated)
                {
                    feasible++;
                    if (!bestObjective.HasValue || (p.Objective.Maximize ? candidate.Objective > bestObjective.Value : candidate.Objective < bestObjective.Value)) bestObjective = candidate.Objective;
                }
                history.Add(new(history.Count + 1, candidate.Values.ToArray(), candidate.Objective, candidate.Feasible, candidate.Violation)
                { Evaluated = candidate.Evaluated, Constraints = candidate.Constraints.ToArray(), BestFeasibleObjective = bestObjective });
                return candidate;
            }
            OptimizationCandidate Evaluate(double[] values)
            {
                cancellation.ThrowIfCancellationRequested();
                for (int i = 0; i < cells.Length; i++) cells[i].WriteSample(values[i]);
                workbook.Calculate();
                // Read formulas afresh for this candidate, before Monte Carlo temporarily samples assumptions.
                var cellValues = constraintCells.ToDictionary(x => x.Key, x => ConstraintCellSelection.Number(x.Value.ReadValue(), x.Key.CellDescription));
                var deterministicConstraints = p.Constraints.Where(c => c.Kind is OptimizationConstraintKind.Cell or OptimizationConstraintKind.DecisionVariable or OptimizationConstraintKind.Linear)
                    .Select(c => EvaluateConstraint(c, c.Kind == OptimizationConstraintKind.Cell ? cellValues[c] : c.Kind == OptimizationConstraintKind.Linear
                        ? OptimizationFeasibleSpace.Value(c, p.Variables, values) : values[p.Variables.ToList().FindIndex(v => v.Id == c.SubjectId)])).ToArray();
                candidateStarted?.Invoke(new(history.Count + 1, values.ToArray(), deterministicConstraints));
                cancellation.ThrowIfCancellationRequested();
                if (deterministicConstraints.Any(c => !c.Met))
                {
                    rejected++;
                    return Record(new OptimizationCandidate(values.ToArray(), double.NaN, deterministicConstraints, new SimulationRunResult(1, [], [])) { Evaluated = false });
                }
                var run = evaluator.Evaluate(cancellation);
                if (!deterministic) simulated++;
                ForecastRunResult Result(string id) => run.ForecastResults.Single(f => ScenarioSimulationService.Identity(f.Forecast).Equals(id, StringComparison.OrdinalIgnoreCase));
                var constraints = p.Constraints.Select(c =>
                {
                    double actual = c.Kind == OptimizationConstraintKind.Linear ? OptimizationFeasibleSpace.Value(c, p.Variables, values) : c.Kind == OptimizationConstraintKind.Cell ? cellValues[c] : c.Kind == OptimizationConstraintKind.DecisionVariable ? values[p.Variables.ToList().FindIndex(v => v.Id == c.SubjectId)] :
                        Statistic(Result(c.SubjectId), c.Kind == OptimizationConstraintKind.Mean ? OptimizationMetric.Mean : c.Kind == OptimizationConstraintKind.Median ? OptimizationMetric.Median : OptimizationMetric.Probability, c.Target, c.TargetDirection);
                    return EvaluateConstraint(c, actual);
                }).ToArray();
                double objective = Statistic(Result(p.Objective.ForecastId), p.Objective.Metric, p.Objective.Target, p.Objective.Direction);
                if (!double.IsFinite(objective)) throw new InvalidOperationException("Objective is outside the supported numeric range.");
                var candidate = new OptimizationCandidate(values.ToArray(), objective, constraints, run);
                return Record(candidate);
            }
            var best = (algorithm ?? new AutomaticOptimizationAlgorithm(guides)).Search(p, original, Evaluate, cancellation, x =>
            {
                if (history.Count > 0) history[^1] = history[^1] with { Phase = x.Phase, RadiusSteps = x.RadiusSteps };
                progress?.Invoke(x with { RejectedBeforeSimulation = rejected, MonteCarloEvaluations = simulated, FeasibleEvaluations = feasible,
                    EvaluationMode = deterministic ? OptimizationEvaluationMode.Deterministic : OptimizationEvaluationMode.MonteCarlo });
            }, out int evaluations);
            cancellation.ThrowIfCancellationRequested();
            return new(p, original, best, evaluations, timer.Elapsed) { History = history, OriginalContents = contents,
                EvaluationMode = deterministic ? OptimizationEvaluationMode.Deterministic : OptimizationEvaluationMode.MonteCarlo,
                RejectedBeforeSimulation = rejected, MonteCarloEvaluations = simulated, FeasibleEvaluations = feasible, ConfiguredForecasts = outputs };
        }
        finally { try { Restore(workbook, cells, contents); } finally { workbook.ScreenUpdating = screen; } }
    }
    private static void Restore(ISimulationWorkbook workbook, ISimulationCell[] cells, SimulationCellContent[] contents)
    {
        var failures = new List<Exception>();
        for (int i = 0; i < cells.Length; i++) try { cells[i].Restore(contents[i]); } catch (Exception ex) { failures.Add(ex); }
        try { workbook.Calculate(); } catch (Exception ex) { failures.Add(ex); }
        if (failures.Count > 0) throw new AggregateException("Optimizer could not fully restore its workbook. Check the reported cells before saving.", failures);
    }
    public static void Apply(OptimizationResult result, ISimulationWorkbook workbook, IReadOnlyList<AssumptionDefinition> inputs, IReadOnlyList<ForecastDefinition> outputs)
    {
        if (!result.Best.Feasible) throw new InvalidOperationException("No feasible solution was found. Adjust constraints or increase evaluations.");
        var cells = ValidateCells(result.Problem, workbook, inputs, outputs); var contents = cells.Select(c => c.Capture()).ToArray();
        var constraintCells = ValidateConstraintCells(result.Problem.Constraints, workbook);
        for (int i = 0; i < cells.Length; i++)
            if (Numeric(cells[i].ReadValue(), cells[i].Identity) != result.OriginalValues[i] ||
                (result.OriginalContents.Count == cells.Length && cells[i].Capture() != result.OriginalContents[i]))
                throw new InvalidOperationException($"{cells[i].Identity} changed since optimization. Run Optimizer again before applying.");
        bool events = workbook.EnableEvents;
        try
        {
            workbook.EnableEvents = false;
            try
            {
                for (int i = 0; i < cells.Length; i++) cells[i].WriteSample(result.Best.Values[i]); workbook.Calculate();
                foreach (var c in result.Problem.Constraints.Where(c => c.Kind == OptimizationConstraintKind.Linear))
                    if (!EvaluateConstraint(c, OptimizationFeasibleSpace.Value(c, result.Problem.Variables, result.Best.Values)).Met)
                        throw new InvalidOperationException("Linear constraint is no longer satisfied. Run Optimizer again.");
                foreach (var pair in constraintCells)
                    if (!EvaluateConstraint(pair.Key, ConstraintCellSelection.Number(pair.Value.ReadValue(), pair.Key.CellDescription)).Met)
                        throw new InvalidOperationException($"Constraint \"{pair.Key.CellDescription}\" is no longer satisfied. Original decisions were restored; run Optimizer again before applying.");
            }
            catch { Restore(workbook, cells, contents); throw; }
        }
        finally { workbook.EnableEvents = events; }
    }
}
