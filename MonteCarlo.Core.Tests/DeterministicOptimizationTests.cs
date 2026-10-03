using MonteCarlo.Excel;
using MonteCarlo.Core;
namespace MonteCarlo.Core.Tests;

public sealed class DeterministicOptimizationTests
{
    private static OptimizationResult Run(OptimizationProblem p, OptimizationTests.Book b, CancellationToken token = default,
        Action<OptimizationProgress>? progress = null, Action<int, IReadOnlyDictionary<ForecastDefinition, double[]>>? live = null) =>
        OptimizationService.Run(p, [], [OptimizationTests.Forecast()], () => b, token, progress, liveProgress: live);

    [Fact] public void PriceDemandRevenueKnownOptimumUsesOneEvaluationPerCandidate()
    {
        var b = new OptimizationTests.Book();
        b.Cells["C1"].Formula = true;
        b.Function = book => { double price = Convert.ToDouble(book.Cells["C1"].Value); return price * (1000 - 5 * price); };
        var p = OptimizationTests.Problem() with { Variables = [new("x", "Price", "Model", "C1", 10, 180, 1)], Settings = new(200, 10000, 42) };
        int frames = 0;
        var r = Run(p, b, live: (n, values) => { frames++; Assert.Equal(1, n); Assert.All(values.Values, v => Assert.Single(v)); });
        Assert.Equal(OptimizationEvaluationMode.Deterministic, r.EvaluationMode);
        Assert.Equal(100, r.Best.Values[0]); Assert.Equal(500, 1000 - 5 * r.Best.Values[0]); Assert.Equal(50000, r.Best.Objective);
        Assert.Equal(171, r.Evaluations); Assert.Equal(r.Evaluations + 1, b.Calculations); Assert.Equal(r.Evaluations, frames);
        var f = Assert.Single(r.Best.Run.ForecastResults); Assert.Single(f.Values); Assert.Equal(50000, f.Mean); Assert.Equal(50000, f.P50); Assert.Equal(0, f.StandardDeviation);
        Assert.Equal(0, b.Cells["A1"].Writes); Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Cells["C1"].Formula); Assert.True(b.Disposed);
    }
    [Theory] [InlineData(true, OptimizationMetric.Mean, 200)] [InlineData(false, OptimizationMetric.Mean, 0)]
    [InlineData(true, OptimizationMetric.Median, 200)] [InlineData(false, OptimizationMetric.Median, 0)]
    public void DeterministicDirectionsAndStatistics(bool maximize, OptimizationMetric metric, double expected)
    {
        var p = OptimizationTests.Problem(maximize) with { Objective = new("forecast", metric, maximize) };
        var r = Run(p, new()); Assert.Equal(expected, r.Best.Objective); Assert.Equal(1, r.Best.Run.Trials);
    }
    [Fact] public void MultipleVariablesAndFormulaConstraints()
    {
        var (b, p, f) = CellConstraintTests.ProductMix();
        var r = OptimizationService.Run(p, [], [f], () => b);
        Assert.Equal(new double[] { 80, 60 }, r.Best.Values); Assert.Equal(88000, r.Best.Objective); Assert.True(r.Best.Feasible);
        Assert.Equal(r.Evaluations + 1, b.Calculations); Assert.Equal(10d, b.Cells["B2"].Value); Assert.Equal(10d, b.Cells["B3"].Value);
    }
    [Theory] [InlineData(OptimizationConstraintKind.Mean)] [InlineData(OptimizationConstraintKind.Median)]
    [InlineData(OptimizationConstraintKind.DecisionVariable)]
    public void LegacyValueConstraintsWork(OptimizationConstraintKind kind)
    {
        var p = OptimizationTests.Problem() with { Constraints = [new(kind, kind == OptimizationConstraintKind.DecisionVariable ? "x" : "forecast", TargetDirection.AtOrBelow, kind == OptimizationConstraintKind.DecisionVariable ? 40 : 80)] };
        var r = Run(p, new()); Assert.Equal(40, r.Best.Values[0]); Assert.True(r.Best.Feasible);
    }
    [Theory] [InlineData(true)] [InlineData(false)]
    public void ProbabilityConfigurationRejectedBeforeMutation(bool objective)
    {
        var p = OptimizationTests.Problem(); p = objective ? p with { Objective = new("forecast", OptimizationMetric.Probability) }
            : p with { Constraints = [new(OptimizationConstraintKind.Probability, "forecast", TargetDirection.AtOrAbove, 80, 20)] };
        bool opened = false;
        var ex = Assert.Throws<ArgumentException>(() => OptimizationService.Run(p, [], [OptimizationTests.Forecast()], () => { opened = true; return new OptimizationTests.Book(); }));
        Assert.Contains("no uncertainty", ex.Message); Assert.False(opened);
    }
    [Fact] public void CancellationRestoresFormulaAndValues()
    {
        using var cts = new CancellationTokenSource(); var b = new OptimizationTests.Book(); b.Cells["C1"].Formula = true;
        Assert.Throws<OperationCanceledException>(() => Run(OptimizationTests.Problem(), b, cts.Token, _ => cts.Cancel()));
        Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Cells["C1"].Formula); Assert.True(b.Disposed);
    }
    [Fact] public void LiveCallbackFailureRestores()
    {
        var b = new OptimizationTests.Book();
        Assert.Throws<InvalidOperationException>(() => Run(OptimizationTests.Problem(), b, live: (_, _) => throw new InvalidOperationException("preview")));
        Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Disposed);
    }
    [Fact] public void CloseWithoutApplyingAndExplicitApply()
    {
        var b = new OptimizationTests.Book(); var r = Run(OptimizationTests.Problem(), b);
        Assert.Equal(50d, b.Cells["C1"].Value); OptimizationService.Apply(r, b, [], [OptimizationTests.Forecast()]);
        Assert.Equal(100d, b.Cells["C1"].Value); Assert.Equal(200d, b.Cells["B1"].Value); Assert.True(b.EnableEvents);
        Assert.Contains("Deterministic", OptimizerExporter.Table(r).Cast<object?>());
    }
    [Fact] public void NormalSimulationStillRejectsNoAssumptions()
    {
        var b = new OptimizationTests.Book(); var outcome = SimulationExecution.Run(10, [], [OptimizationTests.Forecast()], b);
        Assert.False(outcome.Succeeded); Assert.Contains(outcome.Validation.Errors, e => e.Location == "Assumptions"); Assert.Equal(0, b.Calculations);
    }
    [Theory] [InlineData(1)] [InlineData(2)]
    public void AssumptionsSelectMonteCarloAndKeepCommonSeed(int count)
    {
        var b = new OptimizationTests.Book { Function = x => Convert.ToDouble(x.Cells["C1"].Value) + Convert.ToDouble(x.Cells["A1"].Value) };
        var inputs = new List<AssumptionDefinition> { OptimizationTests.Assumption() };
        if (count == 2) { var a = OptimizationTests.Assumption(); a.Name = "Other"; a.CellAddress = "C2"; inputs.Add(a); }
        var p = OptimizationTests.Problem() with { Settings = new(3, 7, 123) }; var streams = new List<double[]>();
        var r = OptimizationService.Run(p, inputs, [OptimizationTests.Forecast()], () => b,
            liveProgress: (n, values) => { if (n == 7) streams.Add(values.Values.Single().Select(v => v - Convert.ToDouble(b.Cells["C1"].Value)).ToArray()); });
        Assert.Equal(OptimizationEvaluationMode.MonteCarlo, r.EvaluationMode); Assert.Equal(7, r.Best.Run.Trials);
        Assert.Equal(3, streams.Count); Assert.True(streams[0].Distinct().Count() > 1);
        Assert.All(streams, s => { for (int i = 0; i < s.Length; i++) Assert.Equal(streams[0][i], s[i], 11); });
        Assert.Equal(21, b.Cells["A1"].Writes); if (count == 2) Assert.Equal(21, b.Cells["C2"].Writes);
    }
    [Theory] [InlineData(null)] [InlineData("text")] [InlineData(double.NaN)]
    public void InvalidForecastRejectedBeforeWrites(object? value)
    {
        var b = new OptimizationTests.Book(); b.Cells["B1"].Value = value;
        Assert.Throws<ArgumentException>(() => Run(OptimizationTests.Problem(), b)); Assert.Equal(0, b.Cells["C1"].Writes);
    }
    [Fact] public void InvalidForecastDuringCandidateRestores()
    {
        var b = new OptimizationTests.Book { Function = _ => double.NaN };
        Assert.Throws<ArgumentException>(() => Run(OptimizationTests.Problem(), b)); Assert.Equal(50d, b.Cells["C1"].Value);
    }
    [Fact] public void TrialsAreIgnoredOnlyByDeterministicEvaluator()
    {
        var p = OptimizationTests.Problem() with { Settings = new(1, 0, 1) }; Assert.Equal(1, Run(p, new()).Best.Run.Trials);
        Assert.Throws<ArgumentException>(() => OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => new OptimizationTests.Book()));
    }
}
