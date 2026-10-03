using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

public class LiveSimulationTests
{
    private static OptimizationTests.Book Book() => new() { Function = b => 2 * Convert.ToDouble(b.Cells["C1"].Value) + Convert.ToDouble(b.Cells["A1"].Value) };
    [Fact]
    public void PreviewIsBoundedDetachedAndDoesNotChangeSeededRun()
    {
        var b = Book(); var forecast = OptimizationTests.Forecast();
        var settings = new SimulationSettings(5000, SimulationSeedMode.Fixed, 42);
        double[]? preview = null; int last = 0;
        var live = SimulationExecution.Run(settings, [OptimizationTests.Assumption()], [forecast], b, liveProgress: (n, values) =>
        {
            Assert.True(b.ScreenUpdating); Assert.False(b.EnableEvents);
            if (n == 1) Assert.Equal(Convert.ToDouble(b.Cells["B1"].Value), values[forecast][0]);
            Assert.InRange(values[forecast].Length, 1, 4096); last = n;
            preview = values[forecast]; preview[0] = -999;
        });
        var ordinary = SimulationExecution.Run(settings, [OptimizationTests.Assumption()], [forecast], Book());
        Assert.True(live.Succeeded); Assert.Equal(5000, last);
        Assert.Equal(ordinary.Result!.ForecastResults[0].Values, live.Result!.ForecastResults[0].Values);
        Assert.Equal(.5, b.Cells["A1"].Value); Assert.True(b.ScreenUpdating); Assert.True(b.EnableEvents);
    }
    [Theory] [InlineData(false)] [InlineData(true)]
    public void PreviewFailureOrCancellationRestoresInputsAndSettings(bool cancel)
    {
        var b = Book(); b.ScreenUpdating = false; b.StatusBar = "original";
        using var c = new CancellationTokenSource();
        var outcome = SimulationExecution.Run(20, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], b,
            cancellation: c.Token, liveProgress: (_, _) =>
            { if (cancel) { c.Cancel(); c.Token.ThrowIfCancellationRequested(); } throw new InvalidOperationException("chart failure"); });
        Assert.False(outcome.Succeeded); Assert.NotNull(outcome.DiagnosticException);
        Assert.Empty(outcome.RestorationFailures); Assert.Equal(.5, b.Cells["A1"].Value);
        Assert.False(b.ScreenUpdating); Assert.True(b.EnableEvents); Assert.Equal("original", b.StatusBar);
    }
    [Fact]
    public void OptimizerLiveTrialsUseCandidateValuesAndRestoreSource()
    {
        var b = Book(); int callbacks = 0;
        var problem = OptimizationTests.Problem() with { Settings = new(3, 10, 42) };
        var result = OptimizationService.Run(problem, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => b,
            liveProgress: (_, values) =>
            {
                callbacks++; Assert.True(b.ScreenUpdating);
                double decision = Convert.ToDouble(b.Cells["C1"].Value);
                Assert.All(values.Values.SelectMany(x => x), x => Assert.InRange(x, decision * 2, decision * 2 + 1));
            });
        Assert.True(callbacks >= result.Evaluations);
        Assert.Equal(50d, b.Cells["C1"].Value); Assert.Equal(.5, b.Cells["A1"].Value);
        Assert.Equal(100.5, Convert.ToDouble(b.Cells["B1"].Value));
    }
}
