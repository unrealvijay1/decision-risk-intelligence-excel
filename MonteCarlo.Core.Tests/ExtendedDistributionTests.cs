using System.Diagnostics;
using MonteCarlo.Core;

namespace MonteCarlo.Core.Tests;

public class ExtendedDistributionTests
{
    public static DiscreteTable Table => new([new(30, .3), new(-10, .2), new(2.5, .5)]);
    public static IEnumerable<object[]> Kinds => Enum.GetValues<DistributionKind>().Where(k => k >= DistributionKind.Exponential).Select(k => new object[] { k });
    [Theory] [MemberData(nameof(Kinds))]
    public void SamplingIsSeededSupportedAndPreviewIsAnalytical(DistributionKind kind)
    {
        double[] p = DistributionCatalog.Defaults(kind);
        var first = new Random(418); var second = new Random(418);
        var table = kind == DistributionKind.Discrete ? Table : null;
        var preview = DistributionPreview.Generate(kind, p[0], p[1], p[2], p[3], table);
        Assert.True(preview.IsValid, preview.ValidationMessage);
        Assert.Equal(DistributionCatalog.IsDiscrete(kind), preview.IsDiscrete);
        Assert.All(preview.Points, point => Assert.True(double.IsFinite(point.Density) && point.Density >= 0));
        for (int i = 0; i < 5000; i++)
        {
            double x = DistributionSampler.Sample(kind, p[0], p[1], p[2], p[3], first, table);
            Assert.Equal(x, DistributionSampler.Sample(kind, p[0], p[1], p[2], p[3], second, table));
            Assert.True(double.IsFinite(x));
            if (kind is DistributionKind.Poisson or DistributionKind.Binomial or DistributionKind.Bernoulli) Assert.Equal(Math.Truncate(x), x);
            if (kind == DistributionKind.TruncatedNormal) Assert.InRange(x, p[2], p[3]);
            else if (kind == DistributionKind.Discrete) Assert.Contains(x, table!.Outcomes.Select(r => r.Outcome));
            else Assert.True(x >= 0);
        }
    }
    [Theory]
    [InlineData(DistributionKind.Exponential, 2, 0, 2, .18393972058572117, .6321205588285577)]
    [InlineData(DistributionKind.Poisson, 3, 0, 2, .22404180765538775, .4231900811268435)]
    [InlineData(DistributionKind.Binomial, 10, .3, 2, .2334744405, .3827827864)]
    [InlineData(DistributionKind.Weibull, 2, 3, 3, .24525296078096154, .6321205588285577)]
    [InlineData(DistributionKind.Gamma, 2, 3, 3, .12262648039048078, .26424111765711533)]
    [InlineData(DistributionKind.Bernoulli, .3, 0, 0, .7, .7)]
    public void IndependentAnalyticalReferences(DistributionKind kind, double a, double b, double x, double mass, double cdf)
    {
        Assert.Equal(mass, ExtendedDistributions.Probability(kind, x, a, b), 11);
        Assert.Equal(cdf, ExtendedDistributions.Cdf(kind, x, a, b), 11);
        Assert.Equal(0, ExtendedDistributions.Cdf(kind, double.NegativeInfinity, a, b));
        Assert.Equal(1, ExtendedDistributions.Cdf(kind, double.PositiveInfinity, a, b));
    }
    [Theory] [MemberData(nameof(Kinds))]
    public void QuantilesSatisfyInverseDefinition(DistributionKind kind)
    {
        double[] p = DistributionCatalog.Defaults(kind); var table = kind == DistributionKind.Discrete ? Table : null;
        foreach (double u in new[] { .00001, .01, .2, .5, .9, .99999 })
        {
            double q = DistributionQuantile.Inverse(kind, u, p[0], p[1], p[2], p[3], table);
            double cdf = ExtendedDistributions.Cdf(kind, q, p[0], p[1], p[2], p[3], table);
            if (DistributionCatalog.IsDiscrete(kind))
            {
                Assert.True(cdf >= u - 1e-12);
                Assert.True(ExtendedDistributions.Cdf(kind, Math.BitDecrement(q), p[0], p[1], p[2], p[3], table) < u + 1e-12);
            }
            else Assert.Equal(u, cdf, 9);
        }
        Assert.Throws<ArgumentException>(() => DistributionQuantile.Inverse(kind, double.NaN, p[0], p[1], p[2], p[3], table));
        Assert.Throws<ArgumentException>(() => DistributionQuantile.Inverse(kind, 1.1, p[0], p[1], p[2], p[3], table));
    }
    [Theory]
    [InlineData(DistributionKind.Exponential, 2, 0, 2, 4)]
    [InlineData(DistributionKind.Poisson, 3, 0, 3, 3)]
    [InlineData(DistributionKind.Poisson, 1000000, 0, 1000000, 1000000)]
    [InlineData(DistributionKind.Binomial, 1000000, .4, 400000, 240000)]
    [InlineData(DistributionKind.Binomial, 1000000, .00001, 10, 9.9999)]
    [InlineData(DistributionKind.Weibull, 2, 3, 2.658680776358274, 1.931416529422966)]
    [InlineData(DistributionKind.Gamma, 2, 3, 6, 18)]
    [InlineData(DistributionKind.Gamma, .1, 3, .3, .9)]
    [InlineData(DistributionKind.Gamma, 1000000, 1, 1000000, 1000000)]
    [InlineData(DistributionKind.Bernoulli, .3, 0, .3, .21)]
    public void SamplingMomentsMatchTheory(DistributionKind kind, double a, double b, double mean, double variance)
    {
        var rng = new Random(1093); const int count = 60000;
        double sum = 0, square = 0;
        for (int i = 0; i < count; i++) { double delta = DistributionSampler.Sample(kind, a, b, random: rng) - mean; sum += delta; square += delta * delta; }
        Assert.InRange(sum / count, -6 * Math.Sqrt(variance / count), 6 * Math.Sqrt(variance / count));
        Assert.InRange(square / count, .91 * variance, 1.09 * variance);
    }
    [Theory]
    [InlineData(DistributionKind.Exponential, 0, 0, 0, 0)]
    [InlineData(DistributionKind.Poisson, -1, 0, 0, 0)]
    [InlineData(DistributionKind.Poisson, 1000001, 0, 0, 0)]
    [InlineData(DistributionKind.Binomial, 1.5, .5, 0, 0)]
    [InlineData(DistributionKind.Binomial, 10, -0.1, 0, 0)]
    [InlineData(DistributionKind.Bernoulli, 1.1, 0, 0, 0)]
    [InlineData(DistributionKind.Weibull, 0, 1, 0, 0)]
    [InlineData(DistributionKind.Gamma, 1, 0, 0, 0)]
    [InlineData(DistributionKind.Gamma, .0001, 1, 0, 0)]
    [InlineData(DistributionKind.TruncatedNormal, 0, 0, -1, 1)]
    [InlineData(DistributionKind.TruncatedNormal, 0, 1, 2, 1)]
    [InlineData(DistributionKind.TruncatedNormal, 0, 1, 10001, 10002)]
    [InlineData(DistributionKind.Exponential, double.NaN, 0, 0, 0)]
    [InlineData(DistributionKind.Discrete, 0, 0, 0, 0)]
    public void InvalidParametersAreRejectedEverywhere(DistributionKind kind, double a, double b, double c, double d)
    {
        Assert.NotNull(DistributionPreview.Validate(kind, a, b, c, d));
        Assert.False(DistributionPreview.Generate(kind, a, b, c, d).IsValid);
        Assert.Throws<ArgumentException>(() => DistributionSampler.Sample(kind, a, b, c, d));
        Assert.Throws<ArgumentException>(() => DistributionQuantile.Inverse(kind, .5, a, b, c, d));
    }
    [Theory] [InlineData(0)] [InlineData(1)]
    public void DeterministicProbabilityCases(double p)
    {
        foreach (double u in new[] { 0, .5, 1d })
        {
            Assert.Equal(p, DistributionQuantile.Inverse(DistributionKind.Bernoulli, u, p, 0));
            Assert.Equal(100 * p, DistributionQuantile.Inverse(DistributionKind.Binomial, u, 100, p));
        }
        Assert.Equal(p, DistributionSampler.Sample(DistributionKind.Bernoulli, p, 0));
        Assert.Equal(100 * p, DistributionSampler.Sample(DistributionKind.Binomial, 100, p));
    }
    [Fact] public void DiscreteTableSortingToleranceAndBoundaries()
    {
        Assert.Equal(new[] { -10d, 2.5, 30 }, Table.Outcomes.Select(r => r.Outcome));
        Assert.Equal(.2, Table.Cdf(-10)); Assert.Equal(.7, Table.Cdf(2.5)); Assert.Equal(.5, Table.Mass(2.5));
        Assert.Equal(-10, Table.Quantile(.2)); Assert.Equal(2.5, Table.Quantile(Math.BitIncrement(.2)));
        var zeros = new DiscreteTable([new(-100, 0), new(10, 1), new(100, 0)]);
        Assert.Equal(10, zeros.Quantile(0)); Assert.Equal(10, zeros.Quantile(1));
        Assert.Equal(1, new DiscreteTable([new(1, .5), new(2, .5 + 1e-11)]).Outcomes.Sum(r => r.Probability), 14);
        var preview = DistributionPreview.Generate(DistributionKind.Discrete, 0, 0, table: Table);
        Assert.Equal(Table.Outcomes.Select(r => r.Outcome), preview.Points.Select(p => p.X));
    }
    [Theory] [InlineData("count")] [InlineData("duplicate")] [InlineData("sum")] [InlineData("negative")] [InlineData("infinite")] [InlineData("too-many")]
    public void InvalidTablesAreRejected(string mode)
    {
        DiscreteOutcome[] rows = mode switch
        {
            "count" => [new(1, 1)], "duplicate" => [new(1, .5), new(1, .5)], "sum" => [new(1, .4), new(2, .5)],
            "negative" => [new(1, -.2), new(2, 1.2)], "infinite" => [new(double.PositiveInfinity, .5), new(2, .5)],
            _ => Enumerable.Range(0, 101).Select(i => new DiscreteOutcome(i, 1d / 101)).ToArray()
        };
        Assert.Throws<ArgumentException>(() => new DiscreteTable(rows));
    }
    [Theory] [InlineData(-1, 1)] [InlineData(8, 9)] [InlineData(40, 41)] [InlineData(-41, -40)] [InlineData(40, 40.000001)] [InlineData(-.000000001, .000000001)]
    public void TruncatedNormalTailsAndNarrowIntervals(double lower, double upper)
    {
        var rng = new Random(1729); double sum = 0;
        for (int i = 0; i < 10000; i++)
        {
            double x = DistributionSampler.Sample(DistributionKind.TruncatedNormal, 0, 1, lower, upper, rng);
            Assert.InRange(x, lower, upper); Assert.NotEqual(lower, x); Assert.NotEqual(upper, x); sum += x;
        }
        foreach (double p in new[] { .01, .25, .5, .9, .99 })
        {
            double x = DistributionQuantile.Inverse(DistributionKind.TruncatedNormal, p, 0, 1, lower, upper);
            Assert.InRange(ExtendedDistributions.Cdf(DistributionKind.TruncatedNormal, x, 0, 1, lower, upper), p - 1e-7, p + 1e-7);
        }
        if (lower == -1) Assert.InRange(sum / 10000, -.025, .025);
        if (lower == 40 && upper == 41) Assert.InRange(sum / 10000, 40.023, 40.027);
    }
    [Fact] public void TruncatedNormalAnalyticalReference()
    {
        Assert.Equal(.5, ExtendedDistributions.Cdf(DistributionKind.TruncatedNormal, 0, 0, 1, -1, 1), 13);
        Assert.Equal(.5843685672568167, ExtendedDistributions.Probability(DistributionKind.TruncatedNormal, 0, 0, 1, -1, 1), 12);
        Assert.Equal(-1, DistributionQuantile.Inverse(DistributionKind.TruncatedNormal, 0, 0, 1, -1, 1));
        Assert.Equal(1, DistributionQuantile.Inverse(DistributionKind.TruncatedNormal, 1, 0, 1, -1, 1));
    }
    [Theory] [InlineData(DistributionKind.Poisson)] [InlineData(DistributionKind.Binomial)] [InlineData(DistributionKind.Gamma)]
    public void MillionLargeParameterSamplesHaveBoundedPracticalCost(DistributionKind kind)
    {
        var rng = new CountingRandom(); var watch = Stopwatch.StartNew();
        double b = kind == DistributionKind.Binomial ? .4 : 1;
        for (int i = 0; i < 1000000; i++) Assert.True(double.IsFinite(DistributionSampler.Sample(kind, 1000000, b, random: rng)));
        Assert.True(rng.Calls < 150_000_000, $"Draw count {rng.Calls}");
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(45), $"Elapsed {watch.Elapsed}");
    }
    [Theory]
    [InlineData(DistributionKind.Poisson, 1000000, 0)]
    [InlineData(DistributionKind.Binomial, 1000000, .5)]
    [InlineData(DistributionKind.Gamma, 1000000, 1)]
    public void LargeParameterQuantilesRemainAccurate(DistributionKind kind, double a, double b)
    {
        foreach (double u in new[] { 1e-8, .01, .5, .99, 1 - 1e-8 })
        {
            double q = DistributionQuantile.Inverse(kind, u, a, b);
            double cdf = ExtendedDistributions.Cdf(kind, q, a, b);
            if (kind == DistributionKind.Gamma) Assert.InRange(cdf, u - 1e-8, u + 1e-8);
            else { Assert.True(cdf >= u - 1e-10); Assert.True(ExtendedDistributions.Cdf(kind, q - 1, a, b) < u + 1e-10); }
        }
    }
    [Fact] public void RareEventsAndSmallTailMassRetainEndpoints()
    {
        Assert.Equal(0, ExtendedDistributions.Probability(DistributionKind.Poisson, double.MaxValue, 1000000));
        Assert.Equal(1, ExtendedDistributions.Cdf(DistributionKind.Poisson, double.MaxValue, 1000000));
        Assert.Equal(Math.Exp(-1e-12), ExtendedDistributions.Cdf(DistributionKind.Binomial, 0, 1000000, 1e-18), 15);
        var table = new DiscreteTable([new(1, 1), new(2, 1e-18)]);
        Assert.Equal(2, table.Quantile(1)); Assert.InRange(table.Cdf(1), 0, 1);
        var random = new Random(1);
        for (int i = 0; i < 100; i++) Assert.InRange(DistributionSampler.Sample(DistributionKind.Binomial, 1000000, Math.BitDecrement(1d), random: random), 0, 1000000);
    }
    private sealed class CountingRandom : Random
    { public long Calls; public CountingRandom() : base(987) { } public override double NextDouble() { Calls++; return base.NextDouble(); } }
}
