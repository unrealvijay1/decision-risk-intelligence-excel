using MonteCarlo.Excel;
using static MonteCarlo.Core.Tests.ScenarioSimulationTests;

namespace MonteCarlo.Core.Tests;

public class ScenarioDecisionTests
{
    private static ForecastDefinition[] Forecasts() => [Output(), new() { Name = "Duration", SheetName = "Model", CellAddress = "B2" }];
    private static ScenarioAnalysisDefinition Analysis(ScenarioAnalysisMode mode, double probability = 80) => new(
        [Scenario("Growth", Change("a", 50)), Scenario("Current Plan", Change("a", 20))], "Current Plan", mode, probability,
        [new(ScenarioSimulationService.Identity(Forecasts()[0]), TargetDirection.AtOrAbove, 130, TargetDirection.AtOrBelow),
         new(ScenarioSimulationService.Identity(Forecasts()[1]), TargetDirection.AtOrBelow, 240, TargetDirection.AtOrAbove)]);

    [Theory]
    [InlineData(0)] [InlineData(50)] [InlineData(80)] [InlineData(87.5)] [InlineData(100)]
    public void ProbabilityValueUsesExistingCalculatorForEveryForecastAndUserBaseline(double probability)
    {
        CheckDecisions(Analysis(ScenarioAnalysisMode.ProbabilityToValue, probability));
    }
    [Fact]
    public void ValueProbabilityUsesCommonPerForecastTargetsAndPointDeltas() => CheckDecisions(Analysis(ScenarioAnalysisMode.ValueToProbability));

    [Fact]
    public void IndependentForecastProbabilitiesDriveValuesAndDeltas()
    {
        var analysis = Analysis(ScenarioAnalysisMode.ProbabilityToValue);
        CheckDecisions(analysis with { ForecastQuestions = [
            analysis.Questions[0] with { Probability = 65 },
            analysis.Questions[1] with { Probability = 95 }] });
    }

    [Theory]
    [InlineData(-1)] [InlineData(101)] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)]
    public void InvalidForecastProbabilityRejected(double p)
    {
        var analysis = Analysis(ScenarioAnalysisMode.ProbabilityToValue);
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(
            analysis with { ForecastQuestions = [analysis.Questions[0] with { Probability = p }, analysis.Questions[1]] }, Forecasts()));
    }
    private static void CheckDecisions(ScenarioAnalysisDefinition analysis)
    {
        var settings = new SimulationSettings(100, SimulationSeedMode.Fixed, 42);
        var inputs = new[] { Input() }; var outputs = Forecasts();
        var result = ScenarioSimulationService.Run(settings, inputs, outputs, analysis, () => new TestSandbox());
        Assert.Equal(4, result.Rows.Count); Assert.Equal("Current Plan", result.BaselineName);
        Assert.DoesNotContain(result.Rows, r => r.Scenario == "Baseline");
        Assert.Equal(2, result.Rows.Count(r => r.IsBaseline));
        var referenceRuns = new Dictionary<string, SimulationRunResult>();
        foreach (var scenario in analysis.Scenarios)
        {
            var copy = ScenarioSimulationService.Clone(inputs[0]);
            var transformed = ScenarioTransformation.Apply(new(DistributionKind.Normal, copy.Parameter1, copy.Parameter2), scenario.Adjustments[0].Value);
            copy.Parameter1 = transformed.P1; copy.Parameter2 = transformed.P2;
            var run = SimulationExecution.Run(settings, [copy], outputs, new TestModel());
            Assert.True(run.Succeeded); referenceRuns[scenario.Id] = run.Result!;
        }
        foreach (var row in result.Rows)
        {
            int i = Array.FindIndex(outputs, f => ScenarioSimulationService.Identity(f) == row.ForecastId);
            var expected = referenceRuns[row.ScenarioId].ForecastResults[i];
            var baseline = referenceRuns[analysis.BaselineScenarioId!].ForecastResults[i];
            var question = analysis.Questions[i];
            double Decision(ForecastRunResult f) => analysis.Mode == ScenarioAnalysisMode.ProbabilityToValue
                ? f.GetTargetForConfidence(analysis.ProbabilityFor(question), question.ValueDirection)
                : question.ProbabilityDirection == TargetDirection.AtOrBelow ? f.ProbabilityLessThanOrEqual(question.Target!.Value) : f.ProbabilityGreaterThanOrEqual(question.Target!.Value);
            Assert.Equal(expected.Mean, row.Mean, 9); Assert.Equal(Decision(expected), row.DecisionValue, 9);
            Assert.Contains(outputs[i].Name, row.Forecast); Assert.Contains(outputs[i].CellAddress, row.Forecast);
            if (row.IsBaseline) { Assert.Null(row.MeanDelta); Assert.Null(row.DecisionDelta); }
            else
            {
                Assert.Equal(expected.Mean - baseline.Mean, row.MeanDelta!.Value, 9);
                Assert.Equal((Decision(expected) - Decision(baseline)) * (analysis.Mode == ScenarioAnalysisMode.ValueToProbability ? 100 : 1), row.DecisionDelta!.Value, 9);
            }
        }
        Assert.Equal(100, inputs[0].Parameter1);
        Assert.Equal(200, outputs[0].TargetSettings!.Target); // Separate scenario question never overwrites normal Results target.
    }
    [Theory]
    [InlineData(null)] [InlineData("")] [InlineData("deleted")]
    public void NoExistingBaselineRejectedBeforeSandbox(string? id)
    {
        var analysis = Analysis(ScenarioAnalysisMode.ProbabilityToValue) with { BaselineScenarioId = id };
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.Run(new(2), [Input()], Forecasts(), analysis, () => throw new Exception("Must not open")));
    }
    [Theory]
    [InlineData(-1)] [InlineData(101)] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)]
    public void InvalidCommonProbabilityRejected(double p) => Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(Analysis(ScenarioAnalysisMode.ProbabilityToValue, p), Forecasts()));
    [Theory]
    [InlineData(ScenarioAnalysisMode.ProbabilityToValue)] [InlineData(ScenarioAnalysisMode.ValueToProbability)]
    public void MissingDirectionRejected(ScenarioAnalysisMode mode)
    {
        var analysis = Analysis(mode); analysis = analysis with { ForecastQuestions = [analysis.Questions[0] with { ValueDirection = null, ProbabilityDirection = null }, analysis.Questions[1]] };
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(analysis, Forecasts()));
    }
    [Theory]
    [InlineData(null)] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)]
    public void InvalidTargetRejected(double? target)
    {
        var analysis = Analysis(ScenarioAnalysisMode.ValueToProbability);
        analysis = analysis with { ForecastQuestions = [analysis.Questions[0] with { Target = target }, analysis.Questions[1]] };
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(analysis, Forecasts()));
    }
    [Fact]
    public void MissingOrDeletedForecastRequiresCorrection()
    {
        var analysis = Analysis(ScenarioAnalysisMode.ValueToProbability);
        Assert.Contains("missing", Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(analysis, [Forecasts()[0]])).Message);
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(analysis with { ForecastQuestions = [analysis.Questions[0]] }, Forecasts()));
    }
    [Fact]
    public void RenameKeepsBaselineIdentityDeletionClearsIt()
    {
        var analysis = Analysis(ScenarioAnalysisMode.ProbabilityToValue);
        var renamed = analysis.WithScenarios(analysis.Scenarios.Select(s => s.Id == analysis.BaselineScenarioId ? s with { Name = "New name" } : s).ToArray());
        Assert.Equal(analysis.BaselineScenarioId, renamed.BaselineScenarioId);
        var result = ScenarioSimulationService.Run(new(10), [Input()], Forecasts(), renamed, () => new TestSandbox());
        Assert.Equal("New name", result.BaselineName);
        var removed = renamed.WithScenarios(renamed.Scenarios.Where(s => s.Id != renamed.BaselineScenarioId).ToArray());
        Assert.Null(removed.BaselineScenarioId);
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(removed, Forecasts()));
    }
    [Fact]
    public void SingleUserScenarioIsOnlyRunAndCanBeNamedBaseline()
    {
        var sandbox = new TestSandbox();
        var analysis = Comparison() with { Scenarios = [new("chosen", "Baseline", [Change("a", 50)])], BaselineScenarioId = "chosen" };
        var result = ScenarioSimulationService.Run(new(5), [Input()], [Output()], analysis, () => sandbox);
        var row = Assert.Single(result.Rows); Assert.True(row.IsBaseline); Assert.Equal("chosen", row.ScenarioId);
        Assert.Equal(1, sandbox.Model.A.Restores); Assert.Equal(5, sandbox.Model.A.Writes);
    }
    [Fact]
    public void DuplicateScenarioIdentityCannotRepresentTwoBaselines()
    {
        var analysis = Analysis(ScenarioAnalysisMode.ProbabilityToValue);
        Assert.Throws<ArgumentException>(() => (analysis with { Scenarios = [analysis.Scenarios[1], analysis.Scenarios[1] with { Name = "Other" }] }).ValidateStored());
    }
}
