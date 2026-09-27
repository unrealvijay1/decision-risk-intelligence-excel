namespace MonteCarlo.Core.Tests;

public class ScenarioTransformationTests
{
    public static IEnumerable<object[]> Distributions()
    {
        foreach (double percent in new[] { 15d, -20d, 0d })
        {
            yield return new object[] { new ScenarioDistribution(DistributionKind.Normal, 100, 15), percent };
            yield return new object[] { new ScenarioDistribution(DistributionKind.Lognormal, 2, .5), percent };
            yield return new object[] { new ScenarioDistribution(DistributionKind.Uniform, -10, 20), percent };
            yield return new object[] { new ScenarioDistribution(DistributionKind.Triangular, 40, 50, 70), percent };
            yield return new object[] { new ScenarioDistribution(DistributionKind.Pert, 40, 50, 70), percent };
            yield return new object[] { new ScenarioDistribution(DistributionKind.Beta, 10, 40, .5, 3), percent };
        }
    }
    [Theory, MemberData(nameof(Distributions))]
    public void EveryDistributionTransformsItsValueParametersOnly(ScenarioDistribution source, double percent)
    {
        var original = source with { };
        var changed = ScenarioTransformation.Apply(source, percent);
        double factor = 1 + percent / 100;
        Assert.Equal(original, source);
        Assert.Null(DistributionPreview.Validate(changed.Kind, changed.P1, changed.P2, changed.P3, changed.P4));
        if (source.Kind == DistributionKind.Lognormal)
        { Assert.Equal(source.P1 + Math.Log(factor), changed.P1); Assert.Equal(source.P2, changed.P2); }
        else
        { Assert.Equal(source.P1 * factor, changed.P1); Assert.Equal(source.P2 * factor, changed.P2); }
        Assert.Equal(source.Kind is DistributionKind.Pert or DistributionKind.Triangular ? source.P3 * factor : source.P3, changed.P3);
        Assert.Equal(source.P4, changed.P4);
        // The transformed distribution is equivalent to scaling the source samples for the same draws.
        var first = new Random(42); var second = new Random(42);
        for (int i = 0; i < 50; i++)
        {
            double expected = ScenarioTransformation.ScaleSample(DistributionSampler.Sample(source.Kind, source.P1, source.P2, source.P3, source.P4, first), percent);
            double actual = DistributionSampler.Sample(changed.Kind, changed.P1, changed.P2, changed.P3, changed.P4, second);
            Assert.Equal(expected, actual, 8);
        }
    }
    [Theory]
    [InlineData(DistributionKind.Normal)] [InlineData(DistributionKind.Lognormal)] [InlineData(DistributionKind.Uniform)]
    [InlineData(DistributionKind.Triangular)] [InlineData(DistributionKind.Pert)] [InlineData(DistributionKind.Beta)]
    public void MinusOneHundredIsAnExplicitZeroConstant(DistributionKind kind)
    {
        var result = ScenarioTransformation.Apply(new(kind, 1, 2, 4, 2), -100);
        Assert.Equal(0, result.Constant); Assert.Equal(0, ScenarioTransformation.ScaleSample(123, -100));
    }
    [Theory]
    [InlineData(10, 110)] [InlineData(-20, 80)] [InlineData(-100, 0)]
    public void ConstantScaling(double change, double expected) =>
        Assert.Equal(expected, ScenarioTransformation.Apply(new(DistributionKind.Normal, 0, 0, Constant: 100), change).Constant!.Value, 10);
    [Theory]
    [InlineData(-101)] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)] [InlineData(double.NegativeInfinity)]
    public void InvalidPercentageRejected(double percent) => Assert.Throws<ArgumentException>(() => ScenarioTransformation.Apply(new(DistributionKind.Normal, 100, 15), percent));
    [Fact]
    public void OverflowRejected() => Assert.Throws<ArgumentException>(() => ScenarioTransformation.Apply(new(DistributionKind.Normal, 1000, 15), double.MaxValue));
    [Fact]
    public void InvalidSourceNotHiddenByZeroScale() => Assert.Throws<ArgumentException>(() => ScenarioTransformation.Apply(new(DistributionKind.Normal, 10, -1), -100));
    [Theory]
    [InlineData("")] [InlineData(" ")] [InlineData("\t")]
    public void InvalidNameRejected(string name) => Assert.Throws<ArgumentException>(() => ScenarioValidation.Validate([new("1", name, [])]));
    [Fact]
    public void DuplicateNamesAndReferencesRejected()
    {
        Assert.Throws<ArgumentException>(() => ScenarioValidation.Validate([new("1", "Growth", []), new("2", " growth ", [])]));
        var adjustment = new ScenarioAdjustment("a", ScenarioAdjustmentType.PercentageChange, 5);
        Assert.Throws<ArgumentException>(() => ScenarioValidation.Validate([new("1", "Growth", [adjustment, adjustment])]));
        Assert.Throws<ArgumentException>(() => ScenarioValidation.Validate([new("1", "Growth", [adjustment with { Type = (ScenarioAdjustmentType)99 }])]));
    }
}
