using MonteCarlo.Core;

namespace MonteCarlo.Core.Tests;

public class EmpiricalProbabilityTests
{
    public static TheoryData<double[], double, double, double, double> Cases => new()
    {
        { new double[] { 0, 1, 1, 2 }, 1, .75, .25, .75 },
        { new double[] { 0, 1, 1, 2 }, -1, 0, 1, 1 },
        { new double[] { 0, 1, 1, 2 }, 3, 1, 0, 0 },
        { new double[] { 0, 1, 1, 2 }, 0, .25, .75, 1 },
        { new double[] { 0, 1, 1, 2 }, 2, 1, 0, .25 },
        { new double[] { 1, 1, 1 }, 1, 1, 0, 1 },
        { new double[] { 1, 1, 1 }, 0, 0, 1, 1 },
        { new double[] { 1, 1, 1 }, 2, 1, 0, 0 },
        { new double[] { -3, -2, -1 }, -2, 2.0 / 3, 1.0 / 3, 2.0 / 3 },
        { new double[] { -2, 0, 2 }, 0, 2.0 / 3, 1.0 / 3, 2.0 / 3 },
        { new double[] { 0, 1, 2 }, double.NaN, 0, 0, 0 },
        { new double[] { 0, 1, 2 }, double.PositiveInfinity, 1, 0, 0 },
        { new double[] { 0, 1, 2 }, double.NegativeInfinity, 0, 1, 1 },
        { new double[] { double.NaN, double.NegativeInfinity, 0, double.PositiveInfinity }, 0, .5, .25, .5 },
        { new double[] { double.NaN, double.NegativeInfinity, 0, double.PositiveInfinity }, double.PositiveInfinity, .75, 0, .25 },
        { new double[] { double.NaN, double.NegativeInfinity, 0, double.PositiveInfinity }, double.NegativeInfinity, .25, .5, .75 }
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void ComparisonsCountSamplesIncludingEqualityCorrectly(
        double[] samples, double target, double below, double above, double atOrAbove)
    {
        Assert.Equal(below, EmpiricalProbability.LessThanOrEqual(samples, target), 12);
        Assert.Equal(above, EmpiricalProbability.GreaterThan(samples, target), 12);
        Assert.Equal(atOrAbove, EmpiricalProbability.GreaterThanOrEqual(samples, target), 12);

        // Regression: the existing methods used these exact comparisons and denominator.
        Assert.Equal((double)samples.Count(x => x <= target) / samples.Length,
            EmpiricalProbability.LessThanOrEqual(samples, target));
        Assert.Equal((double)samples.Count(x => x > target) / samples.Length,
            EmpiricalProbability.GreaterThan(samples, target));
    }

    [Fact]
    public void EmptySamplesReturnUndefinedProbability()
    {
        Assert.True(double.IsNaN(EmpiricalProbability.LessThanOrEqual(Array.Empty<double>(), 1)));
        Assert.True(double.IsNaN(EmpiricalProbability.GreaterThan(Array.Empty<double>(), 1)));
        Assert.True(double.IsNaN(EmpiricalProbability.GreaterThanOrEqual(Array.Empty<double>(), 1)));
    }

    [Fact]
    public void NullSamplesAreRejected()
    {
        Assert.Throws<ArgumentNullException>(() => EmpiricalProbability.LessThanOrEqual(null!, 1));
        Assert.Throws<ArgumentNullException>(() => EmpiricalProbability.GreaterThan(null!, 1));
        Assert.Throws<ArgumentNullException>(() => EmpiricalProbability.GreaterThanOrEqual(null!, 1));
    }
}
