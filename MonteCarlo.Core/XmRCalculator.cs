namespace MonteCarlo.Core;

public static class XmRCalculator
{
    public const int MinimumBaselineObservations = 20;
    // V1 deliberately uses the conventional factors specified for this product.
    // They are not silently replaced by distribution-derived factors (~2.659, ~3.267).
    public const double IndividualsFactor = 2.66;
    public const double MovingRangeFactor = 3.268;

    public static SpcBaseline EstablishBaseline(IReadOnlyList<double> observations, SpcConfiguration configuration)
    {
        ValidateInput(observations, configuration);
        int start = configuration.BaselineMode == SpcBaselineMode.SelectedPeriod ? configuration.BaselineStart : 1;
        int end = configuration.BaselineMode == SpcBaselineMode.AllObservations ? observations.Count : configuration.BaselineEnd;
        ValidatePeriod(start, end, observations.Count);
        double mean = Mean(observations.Skip(start - 1).Take(end - start + 1).ToArray());
        var ranges = new double[end - start];
        for (int i = start; i < end; i++) ranges[i - start] = Difference(observations[i], observations[i - 1]);
        double mr = Mean(ranges);
        if (mr == 0) throw new ArgumentException("The baseline has zero measurable variation. Control limits cannot be established. Check measurement resolution or select a representative baseline with variation.");
        var baseline = new SpcBaseline(start, end, mean, mr, mean - IndividualsFactor * mr, mean + IndividualsFactor * mr, MovingRangeFactor * mr);
        ValidateBaseline(baseline, observations.Count);
        return baseline;
    }

    public static SpcResult Analyze(IReadOnlyList<double> observations, SpcConfiguration configuration,
        SpcBaseline fixedBaseline, IReadOnlyList<string>? labels = null)
    {
        ValidateInput(observations, configuration);
        ValidateBaseline(fixedBaseline, observations.Count);
        if (labels != null && (labels.Count != observations.Count || labels.Any(x => x == null)))
            throw new ArgumentException("Provide exactly one date/sequence label for each observation.");
        double[] values = observations.ToArray();
        var ranges = new double?[values.Length]; // First observation has no moving range.
        for (int i = 1; i < values.Length; i++) ranges[i] = Difference(values[i], values[i - 1]);
        var signals = SpcSignalDetector.Detect(values, ranges, fixedBaseline, configuration.Rules);
        return new(configuration, Array.AsReadOnly(values),
            Array.AsReadOnly(labels?.ToArray() ?? Enumerable.Range(1, values.Length).Select(i => i.ToString()).ToArray()),
            Array.AsReadOnly(ranges), fixedBaseline, signals);
    }
    public static void ValidateBaseline(SpcBaseline baseline, int observationCount)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ValidatePeriod(baseline.Start, baseline.End, observationCount);
        if (!double.IsFinite(baseline.Mean) || !double.IsFinite(baseline.AverageMovingRange) || baseline.AverageMovingRange <= 0 ||
            !double.IsFinite(baseline.IndividualsLower) || !double.IsFinite(baseline.IndividualsUpper) || !double.IsFinite(baseline.MovingRangeUpper) ||
            baseline.IndividualsLower >= baseline.Mean || baseline.IndividualsUpper <= baseline.Mean || baseline.MovingRangeUpper <= 0 ||
            baseline.IndividualsLower != baseline.Mean - IndividualsFactor * baseline.AverageMovingRange ||
            baseline.IndividualsUpper != baseline.Mean + IndividualsFactor * baseline.AverageMovingRange ||
            baseline.MovingRangeUpper != MovingRangeFactor * baseline.AverageMovingRange)
            throw new ArgumentException("The baseline cannot represent valid XmR limits at this numerical scale, or its saved limits are invalid. Rescale the measurement units or explicitly establish a valid baseline.");
    }
    private static void ValidatePeriod(int start, int end, int count)
    {
        if (start < 1 || end > count || end < start || end - start + 1 < MinimumBaselineObservations)
            throw new ArgumentException("Select a contiguous baseline inside the observations containing at least 20 valid observations. Twenty is a practical minimum, not a guarantee of reliable limits.");
    }
    private static void ValidateInput(IReadOnlyList<double> observations, SpcConfiguration configuration)
    {
        ArgumentNullException.ThrowIfNull(observations); ArgumentNullException.ThrowIfNull(configuration);
        if (observations.Count < MinimumBaselineObservations) throw new ArgumentException("At least 20 valid consecutive observations are required.");
        for (int i = 0; i < observations.Count; i++)
            if (!double.IsFinite(observations[i])) throw new ArgumentException($"Observation {i + 1} is missing or non-finite. Correct it before analysis; gaps cannot be skipped.");
        if (string.IsNullOrWhiteSpace(configuration.ProcessName)) throw new ArgumentException("Enter a process name.");
        if (!Enum.IsDefined(configuration.BaselineMode) || configuration.Rules == null)
            throw new ArgumentException("Select a supported baseline mode and signal rules.");
        if (configuration.Target is { } t && (!double.IsFinite(t.Value) || !Enum.IsDefined(t.Direction)))
            throw new ArgumentException("The business target must be finite with a valid target direction.");
    }
    private static double Difference(double a, double b)
    {
        double difference = Math.Abs(a - b);
        if (!double.IsFinite(difference)) throw new ArgumentException("A moving range exceeds numerical precision. Rescale the measurement units before analysis.");
        return difference;
    }
    private static double Mean(double[] values)
    {
        double scale = values.Max(v => Math.Abs(v));
        if (scale == 0) return 0;
        // Scaled compensated sum avoids overflowing the sum of otherwise finite data.
        double sum = 0, compensation = 0;
        foreach (double value in values)
        {
            double term = value / scale - compensation;
            double next = sum + term; compensation = (next - sum) - term; sum = next;
        }
        return (sum / values.Length) * scale;
    }
}
