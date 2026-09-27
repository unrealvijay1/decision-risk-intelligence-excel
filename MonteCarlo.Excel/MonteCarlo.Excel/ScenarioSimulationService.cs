using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public interface IScenarioSandbox : IDisposable
{
    ISimulationWorkbook Workbook { get; }
}

public sealed record ScenarioComparisonRow(string ScenarioId, string Scenario, bool IsBaseline, string ForecastId,
    string Forecast, double Mean, double? MeanDelta, double DecisionValue, double? DecisionDelta,
    double? Target, TargetDirection Direction, double? Probability = null);
public sealed record ScenarioComparisonResult(int Trials, int Seed, string BaselineName,
    ScenarioAnalysisMode Mode, double Probability, IReadOnlyList<ScenarioComparisonRow> Rows);

public static class ScenarioSimulationService
{
    public static string Identity(AssumptionDefinition a) => string.IsNullOrEmpty(a.CellLink)
        ? $"cell:{a.SheetName}!{a.CellAddress}" : a.CellLink;
    public static string Identity(ForecastDefinition f) => string.IsNullOrEmpty(f.CellLink)
        ? $"cell:{f.SheetName}!{f.CellAddress}" : f.CellLink;
    public static string Display(ForecastDefinition f) => string.IsNullOrWhiteSpace(f.Name) || f.Name == $"{f.SheetName}!{f.CellAddress}"
        ? $"{f.SheetName}!{f.CellAddress}" : $"{f.Name} ({f.SheetName}!{f.CellAddress})";

    public static void ValidateAnalysis(ScenarioAnalysisDefinition analysis, IReadOnlyList<ForecastDefinition> forecasts)
    {
        analysis.ValidateStored();
        if (string.IsNullOrEmpty(analysis.BaselineScenarioId) || !analysis.Scenarios.Any(s => s.Id == analysis.BaselineScenarioId))
            throw new ArgumentException("Select one existing scenario as Baseline before running Scenario Analysis.");
        var available = forecasts.Select(Identity).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var question in analysis.Questions)
            if (!available.Contains(question.ForecastId)) throw new ArgumentException($"A configured forecast is missing ({question.ForecastId}). Remove it from Analysis Question before running.");
        if (forecasts.Count == 0) throw new ArgumentException("Define at least one forecast before running Scenario Analysis.");
        foreach (var forecast in forecasts)
        {
            var question = analysis.Questions.SingleOrDefault(q => q.ForecastId.Equals(Identity(forecast), StringComparison.OrdinalIgnoreCase));
            var direction = analysis.Mode == ScenarioAnalysisMode.ProbabilityToValue ? question?.ValueDirection : question?.ProbabilityDirection;
            if (!direction.HasValue) throw new ArgumentException($"{Display(forecast)}: select a direction in Analysis Question.");
            if (analysis.Mode == ScenarioAnalysisMode.ValueToProbability && question?.Target == null)
                throw new ArgumentException($"{Display(forecast)}: enter a target value in Analysis Question.");
        }
    }

    public static AssumptionDefinition Clone(AssumptionDefinition a) => new()
    {
        Name = a.Name, CellLink = a.CellLink, SheetName = a.SheetName, CellAddress = a.CellAddress,
        Distribution = a.Distribution, Parameter1 = a.Parameter1, Parameter2 = a.Parameter2,
        Parameter3 = a.Parameter3, Parameter4 = a.Parameter4
    };
    public static ForecastDefinition Clone(ForecastDefinition f) => new()
    { Name = f.Name, CellLink = f.CellLink, SheetName = f.SheetName, CellAddress = f.CellAddress, TargetSettings = f.TargetSettings };

    public static void Validate(IReadOnlyList<AssumptionDefinition> assumptions, IReadOnlyList<ScenarioDefinition> scenarios)
    {
        ScenarioValidation.Validate(scenarios);
        foreach (var scenario in scenarios)
        foreach (var adjustment in scenario.Adjustments)
        {
            var assumption = assumptions.SingleOrDefault(a => Identity(a).Equals(adjustment.AssumptionId, StringComparison.OrdinalIgnoreCase));
            if (assumption == null) throw new ArgumentException($"{scenario.Name}: an adjusted assumption is missing ({adjustment.AssumptionId}). Remove that adjustment or select an existing assumption.");
            try
            {
                ScenarioTransformation.Apply(new((DistributionKind)assumption.Distribution, assumption.Parameter1,
                    assumption.Parameter2, assumption.Parameter3, assumption.Parameter4), adjustment.Value);
            }
            catch (ArgumentException ex) { throw new ArgumentException($"{scenario.Name} / {assumption.Name}: {ex.Message}", ex); }
        }
    }

    public static ScenarioComparisonResult Run(SimulationSettings settings, IReadOnlyList<AssumptionDefinition> assumptions,
        IReadOnlyList<ForecastDefinition> forecasts, ScenarioAnalysisDefinition analysis,
        Func<IScenarioSandbox> createSandbox, CancellationToken cancellation = default, Action<string>? progress = null, IReadOnlyList<AssumptionCorrelation>? correlations = null)
    {
        if (settings.Validate() is string error) throw new ArgumentException(error);
        ValidateAnalysis(analysis, forecasts);
        correlations = correlations?.ToArray();
        // Freeze caller-owned models before execution, without assigning the global SimulationModel lists.
        var inputs = assumptions.Select(Clone).ToArray();
        var outputs = forecasts.Select(Clone).ToArray();
        var plans = analysis.Scenarios.Select(s => s with { Adjustments = s.Adjustments.ToArray() }).ToArray();
        analysis = analysis with { Scenarios = plans, ForecastQuestions = analysis.Questions.ToArray() };
        Validate(inputs, plans);
        var validation = SimulationValidation.ValidateModel(settings.TrialCount, inputs, outputs);
        if (!validation.IsValid) throw new ArgumentException(validation.UserMessage);
        cancellation.ThrowIfCancellationRequested();
        int seed = settings.SeedMode == SimulationSeedMode.Fixed ? settings.FixedSeed!.Value :
            System.Security.Cryptography.RandomNumberGenerator.GetInt32(int.MaxValue);
        var common = settings with { SeedMode = SimulationSeedMode.Fixed, FixedSeed = seed };
        using var sandbox = createSandbox();
        SimulationRunResult Execute(string name, IReadOnlyDictionary<string, double> changes)
        {
            cancellation.ThrowIfCancellationRequested();
            progress?.Invoke($"Running {name}…");
            var temporary = inputs.Select(Clone).ToArray();
            var outcome = SimulationExecution.Run(common, temporary, outputs.Select(Clone).ToArray(), sandbox.Workbook,
                sampleTransform: (a, sample) => changes.TryGetValue(Identity(a), out double percent)
                    ? ScenarioTransformation.ScaleSample(sample, percent) : sample, cancellation: cancellation, correlations: correlations);
            cancellation.ThrowIfCancellationRequested();
            if (!outcome.Succeeded) throw new InvalidOperationException($"{name}: {outcome.UserMessage}", outcome.DiagnosticException);
            return outcome.Result!;
        }
        var baselinePlan = plans.Single(s => s.Id == analysis.BaselineScenarioId);
        SimulationRunResult ExecutePlan(ScenarioDefinition plan) => Execute(plan.Name,
            plan.Adjustments.ToDictionary(a => a.AssumptionId, a => a.Value, StringComparer.OrdinalIgnoreCase));
        var baseline = ExecutePlan(baselinePlan);
        var rows = new List<ScenarioComparisonRow>();
        void Add(ScenarioDefinition plan, SimulationRunResult run)
        {
            for (int i = 0; i < outputs.Length; i++)
            {
                var b = baseline.ForecastResults[i]; var r = run.ForecastResults[i];
                var question = analysis.Questions.Single(q => q.ForecastId.Equals(Identity(outputs[i]), StringComparison.OrdinalIgnoreCase));
                bool valueMode = analysis.Mode == ScenarioAnalysisMode.ProbabilityToValue;
                TargetDirection direction = (valueMode ? question.ValueDirection : question.ProbabilityDirection)!.Value;
                double Decision(ForecastRunResult result) => valueMode
                    ? result.GetTargetForConfidence(analysis.ProbabilityFor(question), direction)
                    : direction == TargetDirection.AtOrBelow ? result.ProbabilityLessThanOrEqual(question.Target!.Value)
                        : result.ProbabilityGreaterThanOrEqual(question.Target!.Value);
                double decision = Decision(r), reference = Decision(b);
                bool isBaseline = plan.Id == baselinePlan.Id;
                rows.Add(new(plan.Id, plan.Name, isBaseline, Identity(outputs[i]), Display(outputs[i]), r.Mean,
                    isBaseline ? null : r.Mean - b.Mean, decision,
                    isBaseline ? null : (decision - reference) * (valueMode ? 1 : 100),
                    valueMode ? null : question.Target, direction, valueMode ? analysis.ProbabilityFor(question) : null));
            }
        }
        Add(baselinePlan, baseline);
        foreach (var plan in plans)
        {
            if (plan.Id != baselinePlan.Id) Add(plan, ExecutePlan(plan));
        }
        return new(settings.TrialCount, seed, baselinePlan.Name, analysis.Mode, analysis.Probability, rows);
    }
}
