namespace MonteCarlo.Core;

public enum SpcBaselineMode { AllObservations, FirstN, SelectedPeriod }
public enum SpcChart { Individuals, MovingRange }
public enum SpcDirection { Above, Below, Increasing, Decreasing }
public sealed record SpcRules(bool BeyondLimits = true, bool SustainedRun = true, bool SustainedTrend = true,
    bool TwoOfThree = true, bool FourOfFive = true);
public sealed record SpcTarget(double Value, TargetDirection Direction)
{
    public bool IsMet(double value) => Direction == TargetDirection.AtOrBelow ? value <= Value : value >= Value;
}
// All indices in the public SPC contract are one-based, inclusive observation positions.
public sealed record SpcBaseline(int Start, int End, double Mean, double AverageMovingRange,
    double IndividualsLower, double IndividualsUpper, double MovingRangeUpper)
{
    // Derived only from the saved baseline, never from the current observations.
    [System.Text.Json.Serialization.JsonIgnore] public double Sigma => AverageMovingRange / 1.128;
    [System.Text.Json.Serialization.JsonIgnore] public double UpperOneSigma => Mean + Sigma;
    [System.Text.Json.Serialization.JsonIgnore] public double LowerOneSigma => Mean - Sigma;
    [System.Text.Json.Serialization.JsonIgnore] public double UpperTwoSigma => Mean + 2 * Sigma;
    [System.Text.Json.Serialization.JsonIgnore] public double LowerTwoSigma => Mean - 2 * Sigma;
}
public sealed record SpcSignal(int RuleId, string Description, SpcChart Chart, int ObservationIndex,
    int StartIndex, int EndIndex, SpcDirection Direction, double ActualValue, double ReferenceValue)
{
    public IReadOnlyList<int> QualifyingIndices { get; init; } = Array.Empty<int>();
    public IReadOnlyList<double> ActualValues { get; init; } = Array.Empty<double>();
    public string RuleName => RuleId switch
    {
        1 => "Beyond control limits (3σ)", 2 => "Eight-point run",
        3 => "Two of three beyond 2σ", 4 => "Four of five beyond 1σ", 5 => "Six-point trend", _ => "Unknown rule"
    };
    public string QualifyingObservations => string.Join(", ", QualifyingIndices.Select(i => $"#{i}"));
    public string Values => string.Join(", ", ActualValues.Select(v => v.ToString("G6")));
}
public sealed record SpcConfiguration
{
    public string ProcessName { get; init; } = "Process";
    public string MeasurementUnit { get; init; } = "";
    public SpcBaselineMode BaselineMode { get; init; }
    public int BaselineStart { get; init; } = 1;
    public int BaselineEnd { get; init; } = 20;
    public SpcRules Rules { get; init; } = new();
    public SpcTarget? Target { get; init; }
}
public sealed record SpcResult(SpcConfiguration Configuration, IReadOnlyList<double> Observations,
    IReadOnlyList<string> Labels, IReadOnlyList<double?> MovingRanges, SpcBaseline Baseline,
    IReadOnlyList<SpcSignal> Signals)
{
    public bool BaselineHasSignals => Signals.Any(s => s.StartIndex >= Baseline.Start && s.EndIndex <= Baseline.End);
    public int? TargetMetCount => Configuration.Target is { } t ? Observations.Count(t.IsMet) : null;
    public double? TargetMetPercent => TargetMetCount is { } count ? 100.0 * count / Observations.Count : null;
    public string Status => Signals.Count == 0 ? "No signals detected" : "Special-cause signals detected";
    public IEnumerable<string> Insights()
    {
        yield return Signals.Count == 0
            ? "The process has no detected special-cause signals under the selected rules. This does not prove statistical stability."
            : "The process has detected special-cause signals. Investigate the process without assuming a specific operational cause.";
        if (BaselineHasSignals) yield return "The baseline requires investigation before its limits are considered reliable.";
        yield return "Twenty baseline observations are a practical minimum for initial analysis, not a guarantee of reliable limits.";
        if (Signals.Count > 1) yield return "Multiple special-cause signals were detected.";
        if (!Configuration.Rules.BeyondLimits && !Configuration.Rules.SustainedRun && !Configuration.Rules.SustainedTrend &&
            !Configuration.Rules.TwoOfThree && !Configuration.Rules.FourOfFive)
            yield return "All signal rules are disabled; no special-cause signals can be detected.";
        foreach (var group in Signals.GroupBy(s => (s.RuleId, s.Chart, s.Direction)))
            yield return $"{group.Count()} {group.Key.Chart} signal(s): {group.First().Description}.";
        if (Configuration.Target is { } target)
            yield return $"Business target: {TargetMetCount} of {Observations.Count} observations ({TargetMetPercent:0.##}%) meet the target. Baseline mean {(target.IsMet(Baseline.Mean) ? "meets" : "misses")} target; latest observation {(target.IsMet(Observations[^1]) ? "meets" : "misses")} target. Target attainment is separate from statistical predictability.";
    }
}
