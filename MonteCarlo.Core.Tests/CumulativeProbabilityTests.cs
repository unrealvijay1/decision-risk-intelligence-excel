using MonteCarlo.Core;

namespace MonteCarlo.Core.Tests;

public class CumulativeProbabilityTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GroupsDuplicatesAndDoesNotMutateInput(bool sorted)
    {
        double[] samples = sorted ? [1, 2, 2, 4] : [4, 2, 1, 2];
        double[] original = samples.ToArray();
        var result = CumulativeProbability.Build(samples);
        Assert.Null(result.ValidationMessage);
        Assert.Equal(original, samples);
        Assert.Equal(new[] { new CumulativePoint(1, 0, .25),
            new CumulativePoint(2, .25, .75), new CumulativePoint(4, .75, 1) }, result.Points);
        foreach (var point in result.Points)
        {
            Assert.Equal(EmpiricalProbability.LessThanOrEqual(samples, point.X), point.Probability);
            Assert.Equal(EmpiricalProbability.GreaterThanOrEqual(samples, point.X), 1 - point.ProbabilityBefore);
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(5)]
    public void ConstantSamplesMakeOneCompleteJump(int count)
    {
        Assert.Equal(new[] { new CumulativePoint(7, 0, 1) },
            CumulativeProbability.Build(Enumerable.Repeat(7d, count).ToArray()).Points);
    }

    [Theory]
    [InlineData(-10, 0, 1)]
    [InlineData(10, 1, 0)]
    [InlineData(2, .75, .75)]
    public void TargetProbabilitiesKeepInclusiveSemantics(double target, double below, double above)
    {
        double[] samples = [1, 2, 2, 4];
        Assert.Equal(below, EmpiricalProbability.LessThanOrEqual(samples, target));
        Assert.Equal(above, EmpiricalProbability.GreaterThanOrEqual(samples, target));
    }

    [Fact]
    public void CurveDoesNotRequireATarget()
    {
        var result = CumulativeProbability.Build([3, 1]);
        Assert.Equal(0, result.Points[0].ProbabilityBefore);
        Assert.Equal(1, result.Points[^1].Probability);
    }

    [Fact]
    public void EmptySamplesHaveNoCurve() => Assert.Empty(CumulativeProbability.Build([]).Points);

    [Fact]
    public void NullSamplesAreRejected() =>
        Assert.Throws<ArgumentNullException>(() => CumulativeProbability.Build(null!));

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void NonFiniteSamplesInvalidateWholeCurve(double invalid)
    {
        var result = CumulativeProbability.Build([1, invalid, 2]);
        Assert.NotNull(result.ValidationMessage);
        Assert.Empty(result.Points);
    }

    [Fact]
    public void FiniteExtremesRemainValid()
    {
        var result = CumulativeProbability.Build([-double.MaxValue, double.MaxValue]);
        Assert.Null(result.ValidationMessage);
        Assert.Equal(.5, result.Points[0].Probability);
        Assert.Equal(1, result.Points[1].Probability);
    }
}
