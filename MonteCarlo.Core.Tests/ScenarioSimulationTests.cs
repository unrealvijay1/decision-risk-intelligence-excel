using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

public class ScenarioSimulationTests
{
    internal static AssumptionDefinition Input(string id = "a", string address = "A1") => new()
    { CellLink = id, Name = id, SheetName = "Model", CellAddress = address, Distribution = DistributionType.Normal, Parameter1 = 100, Parameter2 = 15 };
    internal static ForecastDefinition Output(TargetDirection? direction = TargetDirection.AtOrBelow) => new()
    { Name = "Profit", SheetName = "Model", CellAddress = "B1", TargetSettings = new(200, direction, null) };
    internal static ScenarioDefinition Scenario(string name, params ScenarioAdjustment[] adjustments) => new(name, name, adjustments);
    internal static ScenarioAdjustment Change(string id, double value) => new(id, ScenarioAdjustmentType.PercentageChange, value);
    // Explicit user definitions in test fixtures, not a production-generated baseline.
    internal static ScenarioAnalysisDefinition Comparison(params ScenarioDefinition[] scenarios) => new(
        new[] { Scenario("Current Plan") }.Concat(scenarios).ToArray(), "Current Plan",
        ForecastQuestions: [new(ScenarioSimulationService.Identity(Output()), TargetDirection.AtOrBelow)]);

    [Theory, MemberData(nameof(ScenarioTransformationTests.Distributions), MemberType = typeof(ScenarioTransformationTests))]
    public void CommonDrawsProduceComparableResultsAndRestoreSandbox(ScenarioDistribution distribution, double percent)
    {
        var input = Input(); input.Distribution = (DistributionType)distribution.Kind;
        input.Parameter1 = distribution.P1; input.Parameter2 = distribution.P2; input.Parameter3 = distribution.P3; input.Parameter4 = distribution.P4;
        var original = ScenarioSimulationService.Clone(input);
        var sandbox = new TestSandbox(); int creates = 0;
        var result = ScenarioSimulationService.Run(new(100, SimulationSeedMode.Fixed, 42), [input], [Output()],
            Comparison(Scenario("Change", Change("a", percent))), () => { creates++; return sandbox; });
        Assert.Equal(1, creates); Assert.Equal(1, sandbox.Disposals);
        Assert.Equal(result.Rows[0].Mean * (1 + percent / 100), result.Rows[1].Mean, 8);
        Assert.Equal(result.Rows[0].DecisionValue * (1 + percent / 100), result.Rows[1].DecisionValue, 8);
        Assert.Equal(original.Parameter1, input.Parameter1); Assert.Equal(original.Parameter2, input.Parameter2);
        Assert.Equal(100d, sandbox.Model.A.Value); Assert.Equal("=50*2", sandbox.Model.A.Formula);
        Assert.True(sandbox.Model.EnableEvents); Assert.Equal("Ready", sandbox.Model.StatusBar);
        Assert.Equal(2, sandbox.Model.A.Restores);
    }
    [Fact]
    public void MultipleAdjustmentsApplyTogetherWithoutChangingSourceOrNormalRun()
    {
        var inputs = new[] { Input(), Input("b", "A2") };
        var source = new TestModel();
        var settings = new SimulationSettings(100, SimulationSeedMode.Fixed, -42);
        var before = SimulationExecution.Run(settings, inputs, [Output()], source);
        var beforeWrites = source.A.Writes;
        var sandbox = new TestSandbox();
        var comparison = ScenarioSimulationService.Run(settings, inputs, [Output()],
            Comparison(Scenario("Both", Change("a", 50), Change("b", 50)), Scenario("Empty")), () => sandbox);
        Assert.Equal(beforeWrites, source.A.Writes);
        Assert.Equal(comparison.Rows[0].Mean * 2.25, comparison.Rows[1].Mean, 8);
        Assert.Equal(comparison.Rows[0].Mean * 1.25, comparison.Rows[1].MeanDelta!.Value, 8);
        Assert.Equal(comparison.Rows[0].Mean, comparison.Rows[2].Mean);
        var after = SimulationExecution.Run(settings, inputs, [Output()], source);
        Assert.Equal(before.Result!.ForecastResults[0].Values, after.Result!.ForecastResults[0].Values);
        Assert.Equal(100, inputs[0].Parameter1); Assert.Equal(15, inputs[1].Parameter2);
    }
    [Theory]
    [InlineData(TargetDirection.AtOrBelow, 1)] [InlineData(TargetDirection.AtOrAbove, 0)]
    public void SuccessUsesExistingTargetDirectionAndPercentagePointDelta(TargetDirection direction, double success)
    {
        var result = ScenarioSimulationService.Run(new(100, SimulationSeedMode.Fixed, 42), [Input()], [Output(direction)],
            Comparison(Scenario("Zero", Change("a", -100))) with { Mode = ScenarioAnalysisMode.ValueToProbability, ForecastQuestions = new[] { new ScenarioForecastQuestion(ScenarioSimulationService.Identity(Output()), Target: 200, ProbabilityDirection: direction) } }, () => new TestSandbox());
        Assert.Equal(success, result.Rows[1].DecisionValue);
        Assert.Equal((success - result.Rows[0].DecisionValue) * 100, result.Rows[1].DecisionDelta);
        Assert.Equal(0, result.Rows[1].Mean);
    }
    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void MissingAnalysisTargetOrDirectionIsRejected(bool missingTarget)
    {
        var analysis = Comparison(Scenario("Empty")) with { Mode = ScenarioAnalysisMode.ValueToProbability,
            ForecastQuestions = [new(ScenarioSimulationService.Identity(Output()), Target: missingTarget ? null : 200,
                ProbabilityDirection: missingTarget ? TargetDirection.AtOrBelow : null)] };
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.Run(new(10), [Input()], [Output()], analysis, () => new TestSandbox()));
    }
    [Fact]
    public void AutomaticComparisonSeedCanBeReplayed()
    {
        var settings = new SimulationSettings(50);
        var scenarios = Comparison(Scenario("Growth", Change("a", 10)));
        var first = ScenarioSimulationService.Run(settings, [Input()], [Output()], scenarios, () => new TestSandbox());
        var second = ScenarioSimulationService.Run(new(50, SimulationSeedMode.Fixed, first.Seed), [Input()], [Output()], scenarios, () => new TestSandbox());
        Assert.Equal(first.Rows, second.Rows); Assert.Equal(SimulationSeedMode.Automatic, settings.SeedMode); Assert.Null(settings.FixedSeed);
    }
    [Fact]
    public void MissingAssumptionFailsBeforeCreatingCopy()
    {
        int creates = 0;
        var exception = Assert.Throws<ArgumentException>(() => ScenarioSimulationService.Run(new(10), [Input()], [Output()],
            Comparison(Scenario("Missing", Change("deleted", 5))), () => { creates++; return new TestSandbox(); }));
        Assert.Contains("missing", exception.Message); Assert.Equal(0, creates);
    }
    [Theory]
    [InlineData("failure")] [InlineData("cancel")] [InlineData("restore")]
    public void FailureOrCancellationRestoresAndDisposes(string mode)
    {
        using var cancellation = new CancellationTokenSource();
        var sandbox = new TestSandbox();
        sandbox.Model.OnCalculate = () => {
            if (mode == "cancel") cancellation.Cancel();
            if (mode == "failure") throw new InvalidOperationException("calculation failed");
        };
        sandbox.Model.A.FailRestore = mode == "restore";
        Assert.ThrowsAny<Exception>(() => ScenarioSimulationService.Run(new(10), [Input()], [Output()],
            Comparison(Scenario("Growth", Change("a", 10))), () => sandbox, cancellation.Token));
        Assert.Equal(1, sandbox.Disposals); Assert.True(sandbox.Model.A.Restores > 0);
        Assert.True(sandbox.Model.EnableEvents); Assert.Equal("Ready", sandbox.Model.StatusBar);
        if (mode != "restore") Assert.Equal("=50*2", sandbox.Model.A.Formula);
    }
    [Fact]
    public void CancelledBeforeStartDoesNotCreateSandbox()
    {
        Assert.Throws<OperationCanceledException>(() => ScenarioSimulationService.Run(new(10), [Input()], [Output()],
            Comparison(Scenario("Empty")), () => throw new Exception("Should not create"), new CancellationToken(true)));
    }
    internal sealed class TestSandbox : IScenarioSandbox
    {
        public TestModel Model = new(); public int Disposals;
        public ISimulationWorkbook Workbook => Model;
        public void Dispose() => Disposals++;
    }
    internal sealed class TestCell(string name) : ISimulationCell
    {
        public string Identity => name; public string? InputProblem => null;
        public double Value = 100; public string? Formula = "=50*2";
        public int Writes, Restores; public bool FailRestore;
        public object ReadValue() => Value;
        public SimulationCellContent Capture() => Formula == null ? new(Value) : new(Formula, true, true);
        public void WriteSample(double value) { Writes++; Value = value; Formula = null; }
        public void Restore(SimulationCellContent content)
        {
            Restores++; if (FailRestore) throw new Exception("restore failed");
            Formula = content.IsFormula ? (string?)content.Value : null;
            Value = content.IsFormula ? 100 : Convert.ToDouble(content.Value);
        }
    }
    internal sealed class TestModel : ISimulationWorkbook
    {
        public TestCell A = new("Model!A1"), B = new("Model!A2"), Output = new("Model!B1"), OutputTwo = new("Model!B2");
        public bool ScreenUpdating { get; set; } = true;
        public bool EnableEvents { get; set; } = true;
        public object? StatusBar { get; set; } = "Ready";
        public Action? OnCalculate;
        public ISimulationCell Resolve(string sheet, string address) => address switch { "A1" => A, "A2" => B, "B2" => OutputTwo, _ => Output };
        public void Calculate() { OnCalculate?.Invoke(); Output.Value = A.Value * B.Value / 100; OutputTwo.Value = 500 - 2 * Output.Value; }
    }
}

