using MonteCarlo.Core;

namespace MonteCarlo.Core.Tests;

public class XmRCalculatorTests
{
    private static double[] Alternating(int count = 20) => Enumerable.Range(0, count).Select(i => i % 2 == 0 ? 9d : 11d).ToArray();
    [Fact] public void IndependentAlternatingReferenceDataset()
    {
        // Ten 9s and ten 11s: sum 200; nineteen ranges all 2.
        var values = Alternating(); var baseline = XmRCalculator.EstablishBaseline(values, new());
        Assert.Equal(10, baseline.Mean, 12); Assert.Equal(2, baseline.AverageMovingRange, 12);
        Assert.Equal(4.68, baseline.IndividualsLower, 12); Assert.Equal(15.32, baseline.IndividualsUpper, 12);
        Assert.Equal(6.536, baseline.MovingRangeUpper, 12);
        var result = XmRCalculator.Analyze(values, new(), baseline);
        Assert.Null(result.MovingRanges[0]); Assert.All(result.MovingRanges.Skip(1), mr => Assert.Equal(2, mr)); Assert.Empty(result.Signals);
    }
    [Fact] public void IndependentIrregularReferenceDataset()
    {
        // Manual sums: X=209; MR=[2,3,2,1,3,5,4,1,2,1,4,7,4,1,2,3,4,5,2], sum 56.
        double[] values = [10, 12, 9, 11, 10, 13, 8, 12, 11, 9, 10, 14, 7, 11, 10, 12, 9, 13, 8, 10];
        var b = XmRCalculator.EstablishBaseline(values, new());
        Assert.Equal(209d, values.Sum()); // Detect reference transcription errors.
        Assert.Equal(10.45, b.Mean, 12);
        Assert.Equal(56d / 19, b.AverageMovingRange, 12);
        Assert.Equal(10.45 - 2.66 * 56 / 19, b.IndividualsLower, 12);
        Assert.Equal(10.45 + 2.66 * 56 / 19, b.IndividualsUpper, 12);
        Assert.Equal(3.268 * 56 / 19, b.MovingRangeUpper, 12);
        Assert.Equal(new double?[] { null, 2, 3, 2, 1, 3, 5, 4, 1, 2, 1, 4, 7, 4, 1, 2, 3, 4, 5, 2 }, XmRCalculator.Analyze(values, new(), b).MovingRanges);
    }
    [Fact] public void FirstNUsesOnlyBaselineAndKeepsTransitionMovingRange()
    {
        var config = new SpcConfiguration { BaselineMode = SpcBaselineMode.FirstN, BaselineEnd = 20 };
        double[] values = Alternating().Concat([100d, 200d]).ToArray();
        var b = XmRCalculator.EstablishBaseline(values, config);
        Assert.Equal(10, b.Mean); Assert.Equal(2, b.AverageMovingRange);
        var result = XmRCalculator.Analyze(values, config, b);
        Assert.Equal(89, result.MovingRanges[20]); Assert.Equal(100, result.MovingRanges[21]);
        Assert.Contains(result.Signals, s => s.Chart == SpcChart.Individuals && s.ObservationIndex == 21);
        Assert.False(result.BaselineHasSignals); // The transition MR uses observation 20 but is a subsequent signal.
    }
    [Fact] public void SelectedPeriodExcludesBothBoundaryMovingRanges()
    {
        var config = new SpcConfiguration { BaselineMode = SpcBaselineMode.SelectedPeriod, BaselineStart = 3, BaselineEnd = 22 };
        double[] values = new[] { 1000d, 1000d }.Concat(Alternating()).Append(2000).ToArray();
        var b = XmRCalculator.EstablishBaseline(values, config);
        Assert.Equal(10, b.Mean); Assert.Equal(2, b.AverageMovingRange); Assert.Equal(3, b.Start); Assert.Equal(22, b.End);
    }
    [Fact] public void PersistedAllObservationsBaselineDoesNotRecalculateAfterAppendOrEdit()
    {
        var original = Alternating(); var b = XmRCalculator.EstablishBaseline(original, new());
        double[] edited = original.Concat(Enumerable.Repeat(50d, 20)).ToArray(); edited[0] = 100;
        var result = XmRCalculator.Analyze(edited, new(), b);
        Assert.Same(b, result.Baseline); Assert.Equal(10, result.Baseline.Mean); Assert.Equal(20, result.Baseline.End);
        Assert.True(result.BaselineHasSignals); Assert.Contains(result.Signals, s => s.ObservationIndex > 20);
    }
    [Fact] public void SignalsAreEvaluatedInsideBaseline()
    {
        double[] values = Enumerable.Range(1, 20).Select(i => (double)i).ToArray(); var b = XmRCalculator.EstablishBaseline(values, new());
        var result = XmRCalculator.Analyze(values, new(), b); Assert.True(result.BaselineHasSignals);
        Assert.Contains(result.Signals, s => s.RuleId == 3 && s.ObservationIndex == 6);
        Assert.Contains(result.Insights(), s => s.Contains("baseline requires investigation"));
    }
    [Fact] public void ConstantBaselineIsRejected() => Assert.Contains("zero", Assert.Throws<ArgumentException>(() => XmRCalculator.EstablishBaseline(Enumerable.Repeat(42d, 20).ToArray(), new())).Message);
    [Theory] [InlineData(0)] [InlineData(19)] public void InsufficientObservationsRejected(int n) => Assert.Throws<ArgumentException>(() => XmRCalculator.EstablishBaseline(Alternating(n), new()));
    [Theory] [InlineData(1, 19)] [InlineData(0, 20)] [InlineData(2, 21)] [InlineData(10, 2)]
    public void InvalidPeriodRejected(int start, int end) => Assert.Throws<ArgumentException>(() => XmRCalculator.EstablishBaseline(Alternating(), new() { BaselineMode = SpcBaselineMode.SelectedPeriod, BaselineStart = start, BaselineEnd = end }));
    [Theory] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)] [InlineData(double.NegativeInfinity)]
    public void NonfiniteValuesRejectedEvenOutsideBaseline(double invalid)
    {
        var values = Alternating().Append(invalid).ToArray(); Assert.Throws<ArgumentException>(() => XmRCalculator.EstablishBaseline(values, new() { BaselineMode = SpcBaselineMode.FirstN, BaselineEnd = 20 }));
    }
    [Fact] public void OverflowingRangesAndLimitsAreRejected()
    {
        Assert.Throws<ArgumentException>(() => XmRCalculator.EstablishBaseline(Alternating().Select(v => v == 9 ? -double.MaxValue : double.MaxValue).ToArray(), new()));
        Assert.Throws<ArgumentException>(() => XmRCalculator.EstablishBaseline(Alternating().Select(v => v == 9 ? 0d : double.MaxValue / 2).ToArray(), new()));
    }
    [Fact] public void LargeFiniteMeanDoesNotOverflowSum()
    {
        var b = XmRCalculator.EstablishBaseline(Alternating().Select(v => v * 1e306).ToArray(), new());
        Assert.True(double.IsFinite(b.Mean)); Assert.InRange(b.Mean / 1e307, .999999999999, 1.000000000001);
    }
    [Fact] public void InvalidSavedLimitsAndLabelsAreRejected()
    {
        var values = Alternating(); var b = XmRCalculator.EstablishBaseline(values, new());
        Assert.Throws<ArgumentException>(() => XmRCalculator.Analyze(values, new(), b with { IndividualsUpper = 99 }));
        Assert.Throws<ArgumentException>(() => XmRCalculator.Analyze(values, new(), b, ["only one"]));
    }
    [Theory] [InlineData(TargetDirection.AtOrBelow, 9, 10)] [InlineData(TargetDirection.AtOrAbove, 11, 10)] [InlineData(TargetDirection.AtOrBelow, 11, 20)]
    public void TargetAttainmentIsInclusiveAndIndependent(TargetDirection direction, double target, int met)
    {
        var values = Alternating(); var config = new SpcConfiguration { Target = new(target, direction) };
        var result = XmRCalculator.Analyze(values, config, XmRCalculator.EstablishBaseline(values, config));
        Assert.Equal(met, result.TargetMetCount); Assert.Equal(met * 5d, result.TargetMetPercent); Assert.Empty(result.Signals);
    }
    [Fact] public void ResultsOwnImmutableCopies()
    {
        var values = Alternating(); var result = XmRCalculator.Analyze(values, new(), XmRCalculator.EstablishBaseline(values, new())); values[0] = 100;
        Assert.Equal(9, result.Observations[0]);
    }
}
