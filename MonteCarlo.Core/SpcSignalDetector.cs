namespace MonteCarlo.Core;

public static class SpcSignalDetector
{
    public static IReadOnlyList<SpcSignal> Detect(IReadOnlyList<double> values, IReadOnlyList<double?> ranges,
        SpcBaseline baseline, SpcRules rules)
    {
        if (values.Count != ranges.Count || values.Any(v => !double.IsFinite(v)) ||
            (ranges.Count > 0 && ranges[0].HasValue) || ranges.Skip(1).Any(v => !v.HasValue) ||
            ranges.Any(v => v.HasValue && (!double.IsFinite(v.Value) || v.Value < 0)))
            throw new ArgumentException("Signal detection requires finite consecutive observations and aligned moving ranges; remove gaps before analysis.");
        var signals = new List<SpcSignal>();
        int runStart = 0, runSide = 0, trendStart = 0, trendSide = 0;
        for (int i = 0; i < values.Count; i++)
        {
            double x = values[i];
            if (rules.BeyondLimits)
            {
                if (x > baseline.IndividualsUpper || x < baseline.IndividualsLower)
                    signals.Add(new(1, x > baseline.IndividualsUpper ? "Observation above the upper control limit" : "Observation below the lower control limit", SpcChart.Individuals, i + 1, i + 1, i + 1,
                        x > baseline.IndividualsUpper ? SpcDirection.Above : SpcDirection.Below, x,
                        x > baseline.IndividualsUpper ? baseline.IndividualsUpper : baseline.IndividualsLower));
                if (ranges[i] is { } mr && mr > baseline.MovingRangeUpper)
                    signals.Add(new(1, "Moving range beyond the upper control limit", SpcChart.MovingRange, i + 1, i, i + 1,
                        SpcDirection.Above, mr, baseline.MovingRangeUpper));
            }
            int side = x.CompareTo(baseline.Mean);
            if (side == 0 || side != runSide) runStart = i;
            runSide = side;
            // One signal at the first qualifying endpoint of each uninterrupted episode.
            if (rules.SustainedRun && side != 0 && i - runStart == 7)
                signals.Add(new(2, side > 0 ? "Eight consecutive observations strictly above the baseline mean" : "Eight consecutive observations strictly below the baseline mean", SpcChart.Individuals,
                    i + 1, runStart + 1, i + 1, side > 0 ? SpcDirection.Above : SpcDirection.Below, x, baseline.Mean));
            if (i == 0) continue;
            int slope = x.CompareTo(values[i - 1]);
            if (slope == 0) trendStart = i;
            else if (slope != trendSide) trendStart = i - 1;
            trendSide = slope;
            if (rules.SustainedTrend && slope != 0 && i - trendStart == 5)
                signals.Add(new(3, slope > 0 ? "Six consecutive observations strictly increasing" : "Six consecutive observations strictly decreasing", SpcChart.Individuals,
                    i + 1, trendStart + 1, i + 1, slope > 0 ? SpcDirection.Increasing : SpcDirection.Decreasing, x, values[trendStart]));
        }
        return signals.AsReadOnly();
    }
}
