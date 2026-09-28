namespace MonteCarlo.Core;

public enum SpcBaselineMode { AllObservations, FirstN, SelectedPeriod }
public enum SpcChart { Individuals, MovingRange }
public enum SpcDirection { Above, Below, Increasing, Decreasing }
public sealed record SpcRules(bool BeyondLimits = true, bool SustainedRun = true, bool SustainedTrend = true);
public sealed record SpcTarget(double Value, TargetDirection Direction)
{
    public bool IsMet(double value) => Direction == TargetDirection.AtOrBelow ? value <= Value : value >= Value;
}
// All indices in the public SPC contract are one-based, inclusive observation positions.
public sealed record SpcBaseline(int Start, int End, double Mean, double AverageMovingRange,
    double IndividualsLower, double IndividualsUpper, double MovingRangeUpper);
public sealed record SpcSignal(int RuleId, string Description, SpcChart Chart, int ObservationIndex,
    int StartIndex, int EndIndex, SpcDirection Direction, double ActualValue, double ReferenceValue);
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
        if (!Configuration.Rules.BeyondLimits && !Configuration.Rules.SustainedRun && !Configuration.Rules.SustainedTrend)
            yield return "All signal rules are disabled; no special-cause signals can be detected.";
        foreach (var group in Signals.GroupBy(s => (s.RuleId, s.Chart, s.Direction)))
            yield return $"{group.Count()} {group.Key.Chart} signal(s): {group.First().Description}.";
        if (Configuration.Target is { } target)
            yield return $"Business target: {TargetMetCount} of {Observations.Count} observations ({TargetMetPercent:0.##}%) meet the target. Baseline mean {(target.IsMet(Baseline.Mean) ? "meets" : "misses")} target; latest observation {(target.IsMet(Observations[^1]) ? "meets" : "misses")} target. Target attainment is separate from statistical predictability.";
    }
}
