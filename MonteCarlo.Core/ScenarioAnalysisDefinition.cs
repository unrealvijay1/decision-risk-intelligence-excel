namespace MonteCarlo.Core;

public enum ScenarioAnalysisMode { ProbabilityToValue, ValueToProbability }
public sealed record ScenarioForecastQuestion(string ForecastId, TargetDirection? ValueDirection = null,
    double? Target = null, TargetDirection? ProbabilityDirection = null, double? Probability = null);

public sealed record ScenarioAnalysisDefinition(IReadOnlyList<ScenarioDefinition> Scenarios,
    string? BaselineScenarioId = null, ScenarioAnalysisMode Mode = ScenarioAnalysisMode.ProbabilityToValue,
    double Probability = 80, IReadOnlyList<ScenarioForecastQuestion>? ForecastQuestions = null)
{
    [System.Text.Json.Serialization.JsonIgnore]
    public IReadOnlyList<ScenarioForecastQuestion> Questions => ForecastQuestions ?? Array.Empty<ScenarioForecastQuestion>();
    public ScenarioAnalysisDefinition WithScenarios(IReadOnlyList<ScenarioDefinition> scenarios) => this with
    { Scenarios = scenarios, BaselineScenarioId = scenarios.Any(s => s.Id == BaselineScenarioId) ? BaselineScenarioId : null };

    public double ProbabilityFor(ScenarioForecastQuestion question) => question.Probability ?? Probability; // Legacy common-probability fallback.

    // Incomplete drafts (no baseline/direction/target yet) may be saved, but never executed.
    public void ValidateStored()
    {
        if (Scenarios == null) throw new ArgumentException("Scenario definitions are missing.");
        ScenarioValidation.Validate(Scenarios);
        if (!Enum.IsDefined(Mode)) throw new ArgumentException("Select a supported analysis mode.");
        if (!double.IsFinite(Probability) || Probability < 0 || Probability > 100)
            throw new ArgumentException("Enter a probability from 0 to 100 percent.");
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var question in Questions)
        {
            if (question == null || string.IsNullOrWhiteSpace(question.ForecastId) || !ids.Add(question.ForecastId))
                throw new ArgumentException("Configure each forecast only once.");
            if (question.Probability is double p && (!double.IsFinite(p) || p < 0 || p > 100))
                throw new ArgumentException("Enter a probability from 0 to 100 percent for each forecast.");
            if (question.Target is double target && !double.IsFinite(target)) throw new ArgumentException("Enter a finite target value.");
            if (question.ValueDirection.HasValue && !Enum.IsDefined(question.ValueDirection.Value) ||
                question.ProbabilityDirection.HasValue && !Enum.IsDefined(question.ProbabilityDirection.Value))
                throw new ArgumentException("Select At or Above or At or Below for the forecast direction.");
        }
    }
}
