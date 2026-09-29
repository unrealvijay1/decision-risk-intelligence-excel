using MonteCarlo.Core;

namespace MonteCarlo.Core.Tests;

public class SpcZoneTests
{
    // Independent reference: MRbar=1.128 gives sigma=1, zones ±1/±2, limits ±3.00048.
    private static readonly SpcBaseline Baseline = new(1, 20, 0, 1.128, -3.00048, 3.00048, 3.686304);
    private static IReadOnlyList<SpcSignal> Detect(double[] values, int rule, bool enabled = true) =>
        SpcSignalDetector.Detect(values, values.Select((x, i) => i == 0 ? (double?)null : Math.Abs(x - values[i - 1])).ToArray(),
            Baseline, new(false, false, false, rule == 3 && enabled, rule == 4 && enabled));

    [Theory]
    [InlineData(3, 1)] [InlineData(3, -1)] [InlineData(4, 1)] [InlineData(4, -1)]
    public void RequiredSameSidePointsWithRemainingPointAnywhere(int rule, int side)
    {
        double[] values = rule == 3 ? [2.5 * side, -10 * side, 2.5 * side] : [1.5 * side, 1.5 * side, -10 * side, 1.5 * side, 1.5 * side];
        var s = Assert.Single(Detect(values, rule));
        Assert.Equal(rule, s.RuleId); Assert.Equal(SpcChart.Individuals, s.Chart);
        Assert.Equal(side > 0 ? SpcDirection.Above : SpcDirection.Below, s.Direction);
        Assert.Equal(1, s.StartIndex); Assert.Equal(values.Length, s.EndIndex); Assert.Equal(values.Length, s.ObservationIndex);
        Assert.Equal((rule == 3 ? 2 : 1) * side, s.ReferenceValue);
        Assert.Equal(rule == 3 ? new[] { 1, 3 } : new[] { 1, 2, 4, 5 }, s.QualifyingIndices);
        Assert.Equal(values, s.ActualValues);
    }
    [Theory] [InlineData(3)] [InlineData(4)]
    public void AllPointsQualify(int rule)
    {
        int window = rule == 3 ? 3 : 5;
        var s = Assert.Single(Detect(Enumerable.Repeat(2.5, window).ToArray(), rule));
        Assert.Equal(Enumerable.Range(1, window), s.QualifyingIndices);
    }
    [Theory] [InlineData(3)] [InlineData(4)]
    public void WindowEndpointNeedNotBeAQualifyingObservation(int rule)
    {
        double[] values = rule == 3 ? [2.5, 2.5, 0] : [1.5, 1.5, 1.5, 1.5, 0];
        var signal = Assert.Single(Detect(values, rule));
        Assert.Equal(values.Length, signal.EndIndex);
        Assert.DoesNotContain(signal.EndIndex, signal.QualifyingIndices);
        Assert.Equal(values, signal.ActualValues);
    }
    [Theory] [InlineData(3)] [InlineData(4)]
    public void TooFewAndOppositeSidesDoNotQualify(int rule)
    {
        Assert.Empty(Detect(rule == 3 ? [2.5, 0, 0] : [1.5, 1.5, 1.5, 0, 0], rule));
        Assert.Empty(Detect(rule == 3 ? [2.5, -2.5, 0] : [1.5, 1.5, 1.5, -1.5, 0], rule));
        Assert.Empty(Detect(rule == 3 ? [2.5, 2.5] : [1.5, 1.5, 1.5, 1.5], rule));
    }
    [Theory] [InlineData(3, 1)] [InlineData(3, -1)] [InlineData(4, 1)] [InlineData(4, -1)]
    public void EqualityNeverQualifies(int rule, int side)
    {
        double threshold = (rule == 3 ? 2 : 1) * side;
        Assert.Empty(Detect(rule == 3 ? [2.5 * side, threshold, threshold] : [1.5 * side, 1.5 * side, 1.5 * side, threshold, threshold], rule));
    }
    [Theory] [InlineData(3)] [InlineData(4)]
    public void ContinuingWindowsEmitOnceAndNonmatchingWindowResetsEpisode(int rule)
    {
        int window = rule == 3 ? 3 : 5;
        double[] values = Enumerable.Repeat(2.5, 12).Concat(Enumerable.Repeat(0d, window)).Concat(Enumerable.Repeat(2.5, window)).ToArray();
        var signals = Detect(values, rule);
        Assert.Equal(2, signals.Count); Assert.Equal(window, signals[0].EndIndex);
        Assert.Equal(values.Length - 1, signals[1].EndIndex);
    }
    [Theory] [InlineData(3)] [InlineData(4)]
    public void DirectionChangeCanStartIndependentEpisode(int rule)
    {
        var signals = Detect(Enumerable.Repeat(2.5, 5).Concat(Enumerable.Repeat(-2.5, 5)).ToArray(), rule);
        Assert.Equal(new[] { SpcDirection.Above, SpcDirection.Below }, signals.Select(s => s.Direction));
    }
    [Theory] [InlineData(3)] [InlineData(4)]
    public void DisabledRuleDoesNotEmit(int rule) => Assert.Empty(Detect(Enumerable.Repeat(5d, 10).ToArray(), rule, false));

    [Fact]
    public void FixedBaselineZonesAndControlLimitsUseIndependentReference()
    {
        double[] original = Enumerable.Range(0, 20).Select(i => i % 2 == 0 ? 9d : 11d).ToArray();
        var b = XmRCalculator.EstablishBaseline(original, new());
        Assert.Equal(1.773049645390071, b.Sigma, 12);
        Assert.Equal(8.226950354609929, b.LowerOneSigma, 12); Assert.Equal(11.77304964539007, b.UpperOneSigma, 12);
        Assert.Equal(6.453900709219858, b.LowerTwoSigma, 12); Assert.Equal(13.546099290780142, b.UpperTwoSigma, 12);
        Assert.Equal(4.68, b.IndividualsLower, 12); Assert.Equal(15.32, b.IndividualsUpper, 12); Assert.Equal(6.536, b.MovingRangeUpper, 12);
        var result = XmRCalculator.Analyze(original.Concat(Enumerable.Repeat(14d, 8)).ToArray(), new(), b);
        Assert.Same(b, result.Baseline);
        Assert.Contains(result.Signals, s => s.RuleId == 3 && s.EndIndex == 22);
        Assert.Contains(result.Signals, s => s.RuleId == 4 && s.EndIndex == 24);
        Assert.Equal(result.Signals.OrderBy(s => s.EndIndex), result.Signals);
        var replaced = XmRCalculator.EstablishBaseline(original.Select(x => x * 2).ToArray(), new());
        Assert.Equal(b.UpperTwoSigma * 2, replaced.UpperTwoSigma); Assert.Equal(b.LowerOneSigma * 2, replaced.LowerOneSigma);
        Assert.Contains(result.Insights(), x => x.Contains("Multiple special-cause"));
        Assert.Contains(result.Insights(), x => x.Contains("upper two-sigma"));
    }
    [Fact]
    public void MissingMovingRangeBoundaryIsRejected()
    {
        Assert.Throws<ArgumentException>(() => SpcSignalDetector.Detect([3, 3, 3], [null, null, 0], Baseline, new()));
    }
}
