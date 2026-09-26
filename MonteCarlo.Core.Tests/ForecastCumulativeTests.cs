using MonteCarlo.Core;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

public class ForecastCumulativeTests
{
    [Theory]
    [InlineData(new double[] { 4, 1, 2, 2 })]
    [InlineData(new double[] { 10, 0 })]
    [InlineData(new double[] { 7 })]
    [InlineData(new double[] { 3, 3, 3 })]
    [InlineData(new double[] { 1.123456789, 5.987654321, 2.345678901 })]
    public void ConfidenceMatchesStoredPercentilesExactly(double[] samples)
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), samples, new());
        var original = forecast.Values.ToArray();
        Assert.Equal(forecast.P50, forecast.GetTargetForConfidence(50, TargetDirection.AtOrBelow));
        Assert.Equal(forecast.P80, forecast.GetTargetForConfidence(80, TargetDirection.AtOrBelow));
        Assert.Equal(forecast.P90, forecast.GetTargetForConfidence(90, TargetDirection.AtOrBelow));
        Assert.Equal(forecast.P80, forecast.GetTargetForConfidence(80));
        Assert.Equal(forecast.P50, forecast.GetTargetForConfidence(50, TargetDirection.AtOrAbove));
        Assert.Equal(original, forecast.Values);
    }

    [Theory]
    [InlineData(80, TargetDirection.AtOrAbove, 2)]
    [InlineData(90, TargetDirection.AtOrAbove, 1)]
    [InlineData(0, TargetDirection.AtOrAbove, 10)]
    [InlineData(100, TargetDirection.AtOrAbove, 0)]
    [InlineData(0, TargetDirection.AtOrBelow, 0)]
    [InlineData(100, TargetDirection.AtOrBelow, 10)]
    [InlineData(87.5, TargetDirection.AtOrBelow, 8.75)]
    [InlineData(73, TargetDirection.AtOrBelow, 7.3)]
    public void ConfidenceUsesCorrectTail(double confidence, TargetDirection direction, double expected)
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), [10, 0], new());
        Assert.Equal(expected, forecast.GetTargetForConfidence(confidence, direction), 12);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(101)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void InvalidConfidenceIsRejected(double confidence)
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), [1], new());
        Assert.Throws<ArgumentOutOfRangeException>(() => forecast.GetTargetForConfidence(confidence));
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void NonFiniteResultsAreNotFiltered(double invalid)
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), [1, invalid, 2], new());
        Assert.Throws<InvalidOperationException>(() => forecast.GetTargetForConfidence(80));
    }

    [Fact]
    public void UnrepresentableInterpolationIsRejected()
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), [-double.MaxValue, double.MaxValue], new());
        Assert.Throws<InvalidOperationException>(() => forecast.GetTargetForConfidence(50));
    }

    [Theory]
    [InlineData(1, 0, TargetDirection.AtOrBelow)]
    [InlineData(1, 100, TargetDirection.AtOrBelow)]
    [InlineData(1, 0, TargetDirection.AtOrAbove)]
    [InlineData(1, 100, TargetDirection.AtOrAbove)]
    [InlineData(4, 0, TargetDirection.AtOrBelow)]
    [InlineData(4, 100, TargetDirection.AtOrBelow)]
    [InlineData(4, 0, TargetDirection.AtOrAbove)]
    [InlineData(4, 100, TargetDirection.AtOrAbove)]
    public void ConstantAndSingleSamplesSupportBoundaryConfidence(int count, double confidence, TargetDirection direction)
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), Enumerable.Repeat(7d, count).ToArray(), new());
        Assert.Equal(7, forecast.GetTargetForConfidence(confidence, direction));
        Assert.Equal(1, forecast.ProbabilityLessThanOrEqual(7));
        Assert.Equal(1, forecast.ProbabilityGreaterThanOrEqual(7));
    }

    [Fact]
    public void FullPrecisionTargetIsUsedForActualProbability()
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), [1.001, 1.002], new());
        double target = forecast.GetTargetForConfidence(80);
        Assert.Equal(forecast.P80, target);
        Assert.NotEqual(Math.Round(target, 2), target);
        Assert.Equal(.5, forecast.ProbabilityLessThanOrEqual(target));
        Assert.Equal(0, forecast.ProbabilityLessThanOrEqual(Math.Round(target, 2)));
    }

    [Fact]
    public void DirectionChangesUseSameSamplesAndInclusiveActualProbability()
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), [0, 2, 2, 10], new());
        double upper = forecast.GetTargetForConfidence(80, TargetDirection.AtOrBelow);
        double lower = forecast.GetTargetForConfidence(80, TargetDirection.AtOrAbove);
        Assert.True(lower < upper);
        Assert.Equal(upper, forecast.GetTargetForConfidence(80, TargetDirection.AtOrBelow));
        double median = forecast.GetTargetForConfidence(50, TargetDirection.AtOrAbove);
        Assert.Equal(2, median);
        Assert.Equal(.75, forecast.ProbabilityGreaterThanOrEqual(median));
        Assert.Equal(.75, forecast.ProbabilityLessThanOrEqual(median));
    }

    [Theory]
    [InlineData(new double[] { 4, 1, 2, 2 }, 2, 2.8)]
    [InlineData(new double[] { 10, 0 }, 5, 8)]
    [InlineData(new double[] { 7 }, 7, 7)]
    [InlineData(new double[] { 3, 3, 3 }, 3, 3)]
    public void BuildingCurvePreservesActualForecastPercentiles(double[] samples, double p50, double p80)
    {
        var forecast = new ForecastRunResult(new ForecastDefinition(), samples, new());
        double[] original = forecast.Values.ToArray();
        double originalP50 = forecast.P50;
        double originalP80 = forecast.P80;
        var curve = CumulativeProbability.Build(forecast.Values);
        Assert.Equal(p50, forecast.P50, 12);
        Assert.Equal(p80, forecast.P80, 12);
        Assert.Equal(originalP50, forecast.P50);
        Assert.Equal(originalP80, forecast.P80);
        Assert.Equal(original, forecast.Values);
        foreach (var point in curve.Points)
        {
            Assert.Equal(forecast.ProbabilityLessThanOrEqual(point.X), point.Probability, 12);
            Assert.Equal(forecast.ProbabilityGreaterThanOrEqual(point.X), 1 - point.ProbabilityBefore, 12);
        }
    }

    [Fact]
    public void NewRunUsesNewOutcomes()
    {
        var first = new ForecastRunResult(new ForecastDefinition(), [1, 2], new());
        var second = new ForecastRunResult(new ForecastDefinition(), [10, 20], new());
        Assert.Equal(2, CumulativeProbability.Build(first.Values).Points[^1].X);
        Assert.Equal(20, CumulativeProbability.Build(second.Values).Points[^1].X);
    }
}
