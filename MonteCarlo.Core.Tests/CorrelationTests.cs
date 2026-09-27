using MonteCarlo.Excel;
using static MonteCarlo.Core.Tests.ScenarioSimulationTests;

namespace MonteCarlo.Core.Tests;

public class CorrelationTests
{
    internal static double Pearson(double[] a, double[] b)
    {
        double x = a.Average(), y = b.Average(), xy = 0, xx = 0, yy = 0;
        for (int i = 0; i < a.Length; i++) { xy += (a[i] - x) * (b[i] - y); xx += (a[i] - x) * (a[i] - x); yy += (b[i] - y) * (b[i] - y); }
        return xy / Math.Sqrt(xx * yy);
    }
    private static (double[], double[]) Draw(double coefficient, bool mixed = false)
    {
        var sampler = new GaussianDependence(["a", "b"], [new("a", "b", coefficient)]);
        var random = new Random(1234); var z = new double[sampler.Ids.Count];
        var a = new double[20000]; var b = new double[a.Length];
        for (int i = 0; i < a.Length; i++)
        {
            if (coefficient == 0) { a[i] = NormalDistribution.Sample(0, 1, random); b[i] = NormalDistribution.Sample(0, 1, random); }
            else { sampler.NextGaussian(random, z); a[i] = z[0]; b[i] = z[1]; }
            if (mixed) { a[i] = DistributionQuantile.FromGaussian(DistributionKind.Triangular, a[i], 50, 70, 100); b[i] = DistributionQuantile.FromGaussian(DistributionKind.Normal, b[i], 35, 8); }
        }
        return (a, b);
    }
    [Theory]
    [InlineData(-1)] [InlineData(-.8)] [InlineData(0)] [InlineData(.8)] [InlineData(1)]
    public void LatentCorrelationAndMixedDependence(double coefficient)
    {
        var (a, b) = Draw(coefficient);
        Assert.InRange(Pearson(a, b), coefficient - .02, coefficient + .02);
        var (x, y) = Draw(coefficient, true);
        if (coefficient > 0) Assert.True(Pearson(x, y) > .7);
        else if (coefficient < 0) Assert.True(Pearson(x, y) < -.7);
        else Assert.InRange(Pearson(x, y), -.03, .03);
        Assert.InRange(x.Average(), 73, 73.7); Assert.InRange(y.Average(), 34.8, 35.2);
        var again = Draw(coefficient, true); Assert.Equal(x, again.Item1); Assert.Equal(y, again.Item2);
    }
    [Fact]
    public void StrongerDependenceProducesStrongerObservedDependence()
    {
        var weak = Draw(.2, true); var strong = Draw(.8, true);
        Assert.True(Pearson(strong.Item1, strong.Item2) > Pearson(weak.Item1, weak.Item2) + .4);
    }
    [Theory]
    [InlineData(DistributionKind.Normal, 10, 2, 0, 0, 10, 4)]
    [InlineData(DistributionKind.Lognormal, 0, .5, 0, 0, 1.133148453, .364695854)]
    [InlineData(DistributionKind.Uniform, 2, 8, 0, 0, 5, 3)]
    [InlineData(DistributionKind.Triangular, 0, 2, 10, 0, 4, 4.666666667)]
    [InlineData(DistributionKind.Pert, 0, 5, 10, 0, 5, 3.571428571)]
    [InlineData(DistributionKind.Beta, 0, 1, 2, 5, .285714286, .025510204)]
    public void AllMarginalsPreserveMeanVarianceAndDistribution(DistributionKind kind, double p1, double p2, double p3, double p4, double mean, double variance)
    {
        var sampler = new GaussianDependence(["a", "b"], [new("a", "b", .8)]);
        var z = new double[2]; var random = new Random(42); var samples = new double[12000];
        for (int i = 0; i < samples.Length; i++) { sampler.NextGaussian(random, z); samples[i] = DistributionQuantile.FromGaussian(kind, z[1], p1, p2, p3, p4); }
        double actualMean = samples.Average(), actualVariance = samples.Select(x => (x - actualMean) * (x - actualMean)).Average();
        Assert.InRange(actualMean, mean - .05 * Math.Sqrt(variance), mean + .05 * Math.Sqrt(variance));
        Assert.InRange(actualVariance, .92 * variance, 1.08 * variance);
        Array.Sort(samples);
        foreach (double probability in new[] { .1, .5, .9 })
            Assert.InRange(samples[(int)(samples.Length * probability)],
                DistributionQuantile.Inverse(kind, probability - .02, p1, p2, p3, p4), DistributionQuantile.Inverse(kind, probability + .02, p1, p2, p3, p4));
    }
    [Theory]
    [InlineData(.2, .3)] [InlineData(1, 1)] [InlineData(2, 5)] [InlineData(50, 60)]
    public void BetaInverseRoundTrips(double a, double b)
    {
        foreach (double p in new[] { .0001, .01, .2, .5, .8, .99, .9999 })
        {
            double x = DistributionQuantile.Inverse(DistributionKind.Beta, p, 0, 1, a, b);
            if (x > 0 && x < 1) Assert.InRange(DistributionQuantile.BetaCdf(x, a, b), p - 2e-5, p + 2e-5);
        }
    }
    [Fact]
    public void ValidThreeWayAndSingularMatrices()
    {
        var sampler = new GaussianDependence(["a", "b", "c"], [new("a", "b", .5), new("a", "c", .3), new("b", "c", -.2)]);
        var z = new double[3]; sampler.NextGaussian(new Random(1), z); Assert.All(z, x => Assert.True(double.IsFinite(x)));
        var singular = new GaussianDependence(["a", "b", "c"], [new("a", "b", 1), new("a", "c", -1), new("b", "c", -1)]);
        singular.NextGaussian(new Random(1), z); Assert.Equal(z[0], z[1]); Assert.Equal(-z[0], z[2]);
    }
    [Fact]
    public void RejectInconsistentMatrix() => Assert.Throws<ArgumentException>(() => new GaussianDependence(["a", "b", "c"],
        [new("a", "b", .9), new("a", "c", .9), new("b", "c", -.9)]));
    [Theory]
    [InlineData("a", "a", .5)] [InlineData("a", "missing", .5)] [InlineData("a", "b", 1.01)]
    [InlineData("a", "b", double.NaN)] [InlineData("a", "b", double.PositiveInfinity)]
    public void InvalidRelationship(string a, string b, double rho) => Assert.Throws<ArgumentException>(() => new GaussianDependence(["a", "b"], [new(a, b, rho)]));
    [Fact]
    public void ReversedDuplicatesRejected() => Assert.Throws<ArgumentException>(() => new GaussianDependence(["a", "b"], [new("a", "b", .4), new("B", "A", .4)]));
    [Fact]
    public void ZeroRelationshipsPreserveExactIndependentSequence()
    {
        var settings = new SimulationSettings(300, SimulationSeedMode.Fixed, 12); var inputs = new[] { Input(), Input("b", "A2") };
        var original = SimulationExecution.Run(settings, inputs, [Output()], new TestModel());
        foreach (var relationships in new AssumptionCorrelation[][] { [], [new("a", "b", 0)] })
        {
            var run = SimulationExecution.Run(settings, inputs, [Output()], new TestModel(), correlations: relationships);
            Assert.True(run.Succeeded, run.UserMessage); Assert.Equal(original.Result!.ForecastResults[0].Values, run.Result!.ForecastResults[0].Values);
        }
    }
    [Fact]
    public void ExecutionSensitivityRestorationAndScenarioCommonDraws()
    {
        var inputs = new[] { Input(), Input("b", "A2") }; var settings = new SimulationSettings(500, SimulationSeedMode.Fixed, 23);
        AssumptionCorrelation[] correlations = [new("a", "b", 1)];
        var source = new TestModel();
        var first = SimulationExecution.Run(settings, inputs, [Output()], source, correlations: correlations);
        var second = SimulationExecution.Run(settings, inputs, [Output()], new TestModel(), correlations: correlations);
        Assert.True(first.Succeeded, first.UserMessage); Assert.Equal(first.Result!.ForecastResults[0].Values, second.Result!.ForecastResults[0].Values);
        Assert.All(first.Result.ForecastResults[0].Sensitivities, s => Assert.InRange(s.Correlation, .999, 1.001));
        Assert.Equal("=50*2", source.A.Formula); Assert.Equal(100, source.A.Value);
        var sandbox = new TestSandbox(); int writes = source.A.Writes;
        var comparison = ScenarioSimulationService.Run(settings, inputs, [Output()], Comparison(Scenario("Growth", Change("a", 50))), () => sandbox, correlations: correlations);
        Assert.Equal(comparison.Rows[0].Mean * 1.5, comparison.Rows[1].Mean, 8);
        Assert.Equal(writes, source.A.Writes); Assert.Equal(1, sandbox.Disposals);
    }
    [Fact]
    public void InvalidCorrelationNeverWritesInputs()
    {
        var model = new TestModel(); var run = SimulationExecution.Run(new SimulationSettings(10), [Input(), Input("b", "A2")], [Output()], model, correlations: [new("a", "b", 2)]);
        Assert.False(run.Succeeded); Assert.Equal(0, model.A.Writes); Assert.Equal(0, model.B.Writes);
    }
    [Theory]
    [InlineData(.001)] [InlineData(.1)] [InlineData(.5)] [InlineData(.99)]
    public void BetaQuantilesMatchClosedForms(double p)
    {
        Assert.Equal(p, DistributionQuantile.Inverse(DistributionKind.Beta, p, 0, 1, 1, 1), 10);
        Assert.Equal(Math.Sqrt(p), DistributionQuantile.Inverse(DistributionKind.Beta, p, 0, 1, 2, 1), 10);
        Assert.Equal(Math.Pow(Math.Sin(Math.PI * p / 2), 2), DistributionQuantile.Inverse(DistributionKind.Beta, p, 0, 1, .5, .5), 10);
    }
    [Fact]
    public void ThreeWayLatentCovarianceAndUnlinkedUniformLognormal()
    {
        var sampler = new GaussianDependence(["a", "b", "c", "unlinked"], [new("a", "b", .5), new("a", "c", .3), new("b", "c", -.2)]);
        Assert.Equal(3, sampler.Ids.Count);
        var arrays = Enumerable.Range(0, 3).Select(_ => new double[20000]).ToArray(); var random = new Random(42); var z = new double[3];
        for (int i = 0; i < arrays[0].Length; i++) { sampler.NextGaussian(random, z); for (int j = 0; j < 3; j++) arrays[j][i] = z[j]; }
        Assert.InRange(Pearson(arrays[0], arrays[1]), .48, .52);
        Assert.InRange(Pearson(arrays[0], arrays[2]), .28, .32);
        Assert.InRange(Pearson(arrays[1], arrays[2]), -.22, -.18);
        var uniform = arrays[0].Select(x => DistributionQuantile.FromGaussian(DistributionKind.Uniform, x, 0, 1)).ToArray();
        var lognormal = arrays[1].Select(x => DistributionQuantile.FromGaussian(DistributionKind.Lognormal, x, 0, .5)).ToArray();
        Assert.InRange(Pearson(uniform, lognormal), .35, .6);
    }
    [Theory]
    [InlineData("failure")] [InlineData("cancel")] [InlineData("restore")]
    public void CorrelatedScenarioFailureAndCancellationCleanUp(string mode)
    {
        using var token = new CancellationTokenSource(); var sandbox = new TestSandbox();
        sandbox.Model.OnCalculate = () => { if (mode == "cancel") token.Cancel(); if (mode == "failure") throw new InvalidOperationException("failed"); };
        sandbox.Model.A.FailRestore = mode == "restore";
        Assert.ThrowsAny<Exception>(() => ScenarioSimulationService.Run(new(10), [Input(), Input("b", "A2")], [Output()], Comparison(),
            () => sandbox, token.Token, correlations: [new("a", "b", .8)]));
        Assert.Equal(1, sandbox.Disposals); Assert.True(sandbox.Model.B.Restores > 0); Assert.True(sandbox.Model.EnableEvents);
    }
}
