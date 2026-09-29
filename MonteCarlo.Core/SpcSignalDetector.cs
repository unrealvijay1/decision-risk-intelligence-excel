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
        var activeZones = new bool[2, 2];
        void Zone(int end, int window, int required, int ruleId, double lower, double upper)
        {
            if (end + 1 < window) return;
            for (int side = 0; side < 2; side++)
            {
                double threshold = side == 0 ? lower : upper;
                int start = end - window + 1;
                int[] qualifying = Enumerable.Range(start, window)
                    .Where(j => side == 0 ? values[j] < threshold : values[j] > threshold).Select(j => j + 1).ToArray();
                bool matches = qualifying.Length >= required;
                // A nonqualifying window ends this rule/direction's episode.
                if (matches && !activeZones[ruleId - 3, side])
                    signals.Add(new(ruleId,
                        $"{(ruleId == 3 ? "Two of three" : "Four of five")} consecutive observations {(side == 0 ? "fall below the lower" : "exceed the upper")} {(ruleId == 3 ? "two" : "one")}-sigma threshold",
                        SpcChart.Individuals, end + 1, start + 1, end + 1,
                        side == 0 ? SpcDirection.Below : SpcDirection.Above, values[end], threshold)
                    {
                        QualifyingIndices = Array.AsReadOnly(qualifying),
                        ActualValues = Array.AsReadOnly(Enumerable.Range(start, window).Select(j => values[j]).ToArray())
                    });
                activeZones[ruleId - 3, side] = matches;
            }
        }
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
            if (rules.TwoOfThree) Zone(i, 3, 2, 3, baseline.LowerTwoSigma, baseline.UpperTwoSigma);
            if (rules.FourOfFive) Zone(i, 5, 4, 4, baseline.LowerOneSigma, baseline.UpperOneSigma);
            if (i == 0) continue;
            int slope = x.CompareTo(values[i - 1]);
            if (slope == 0) trendStart = i;
            else if (slope != trendSide) trendStart = i - 1;
            trendSide = slope;
            if (rules.SustainedTrend && slope != 0 && i - trendStart == 5)
                signals.Add(new(5, slope > 0 ? "Six consecutive observations strictly increasing" : "Six consecutive observations strictly decreasing", SpcChart.Individuals,
                    i + 1, trendStart + 1, i + 1, slope > 0 ? SpcDirection.Increasing : SpcDirection.Decreasing, x, values[trendStart]));
        }
        return signals.Select(s => s.ActualValues.Count > 0 ? s : s with
        {
            QualifyingIndices = Array.AsReadOnly(s.Chart == SpcChart.MovingRange ? new[] { s.ObservationIndex }
                : Enumerable.Range(s.StartIndex, s.EndIndex - s.StartIndex + 1).ToArray()),
            ActualValues = Array.AsReadOnly(s.Chart == SpcChart.MovingRange ? new[] { s.ActualValue }
                : Enumerable.Range(s.StartIndex - 1, s.EndIndex - s.StartIndex + 1).Select(j => values[j]).ToArray())
        }).ToList().AsReadOnly();
    }
}
