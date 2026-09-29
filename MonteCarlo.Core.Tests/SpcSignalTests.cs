using MonteCarlo.Core;

namespace MonteCarlo.Core.Tests;

public class SpcSignalTests
{
    private static readonly SpcBaseline Baseline = new(1, 20, 0, 10, -20, 20, 30);
    private static IReadOnlyList<SpcSignal> Detect(double[] values, SpcRules? rules = null) =>
        SpcSignalDetector.Detect(values, values.Select((x, i) => i == 0 ? (double?)null : Math.Abs(x - values[i - 1])).ToArray(), Baseline, rules ?? new());

    [Fact] public void BeyondLimitsAreStrictAndBothChartsAreEvaluated()
    {
        var signals = Detect([20, -20, 21, -21], new(true, false, false, false, false));
        Assert.Equal(2, signals.Count(s => s.Chart == SpcChart.Individuals));
        Assert.Equal(3, signals.Count(s => s.Chart == SpcChart.MovingRange));
        Assert.Contains(signals, s => s.ObservationIndex == 3 && s.Direction == SpcDirection.Above && s.ReferenceValue == 20);
        Assert.Contains(signals, s => s.ObservationIndex == 4 && s.Direction == SpcDirection.Below && s.ReferenceValue == -20);
    }
    [Fact] public void MovingRangeExactlyAtLimitIsNotASignal() => Assert.Empty(Detect([-15, 15], new(true, false, false, false, false)));
    [Theory] [InlineData(1)] [InlineData(-1)]
    public void RunSignalsOncePerUninterruptedEpisode(int direction)
    {
        var signals = Detect(Enumerable.Repeat((double)direction, 24).ToArray(), new(false, true, false, false, false));
        var signal = Assert.Single(signals); Assert.Equal(8, signal.ObservationIndex); Assert.Equal(1, signal.StartIndex);
        Assert.Equal(direction > 0 ? SpcDirection.Above : SpcDirection.Below, signal.Direction);
    }
    [Fact] public void CenterAndSideChangesInterruptRuns()
    {
        Assert.Empty(Detect([1, 1, 1, 1, 1, 1, 1, 0, 1, 1, 1, 1, 1, 1, 1], new(false, true, false, false, false)));
        Assert.Empty(Detect([1, 1, 1, 1, 1, 1, 1, -1, 1, 1, 1, 1, 1, 1, 1], new(false, true, false, false, false)));
        Assert.Equal(2, Detect(Enumerable.Repeat(1d, 8).Append(0).Concat(Enumerable.Repeat(1d, 8)).ToArray(), new(false, true, false, false, false)).Count);
    }
    [Theory] [InlineData(1)] [InlineData(-1)]
    public void TrendUsesSixObservationsAndReportsOnlyFirstEndpoint(int direction)
    {
        var signal = Assert.Single(Detect(Enumerable.Range(1, 12).Select(i => (double)i * direction).ToArray(), new(false, false, true, false, false)));
        Assert.Equal(6, signal.ObservationIndex); Assert.Equal(1, signal.StartIndex); Assert.Equal(6 * direction, signal.ActualValue);
    }
    [Fact] public void EqualValuesAndReversalsInterruptTrends()
    {
        Assert.Empty(Detect([1, 2, 3, 4, 5, 5, 6, 7, 8, 9], new(false, false, true, false, false)));
        Assert.DoesNotContain(Detect([1, 2, 3, 4, 5, 4, 3, 2, 1, 0], new(false, false, true, false, false)), s => s.ObservationIndex < 10);
        Assert.Single(Detect([1, 2, 3, 4, 5, 4, 3, 2, 1, 0], new(false, false, true, false, false)));
    }
    [Fact] public void ObservationCanParticipateInMultipleRules()
    {
        var signals = Detect([1, 1, 1, 2, 3, 4, 5, 25]);
        Assert.Equal(3, signals.Count(s => s.ObservationIndex == 8));
    }
    [Fact] public void DisabledRulesProduceNoSignals() => Assert.Empty(Detect([1, 2, 3, 4, 5, 6, 7, 40], new(false, false, false, false, false)));
    [Theory] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)] [InlineData(double.NegativeInfinity)]
    public void MissingOrNonfiniteBoundariesAreRejected(double invalid) => Assert.Throws<ArgumentException>(() => Detect([1, 2, invalid, 3, 4, 5, 6, 7, 8]));
}
