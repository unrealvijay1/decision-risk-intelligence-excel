using ExcelDna.Integration;
using MonteCarlo.Excel;
using static MonteCarlo.Core.Tests.ScenarioSimulationTests;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public class ExtendedDistributionIntegrationTests
{
    public static IEnumerable<object[]> Kinds => ExtendedDistributionTests.Kinds;
    public static AssumptionDefinition InputFor(DistributionKind kind)
    {
        var p = DistributionCatalog.Defaults(kind); var input = Input();
        input.Distribution = (DistributionType)kind;
        input.Parameter1 = p[0]; input.Parameter2 = p[1]; input.Parameter3 = p[2]; input.Parameter4 = p[3];
        input.ProbabilityTable = kind == DistributionKind.Discrete ? ExtendedDistributionTests.Table : null;
        return input;
    }
    [Theory, MemberData(nameof(Kinds))]
    public void ExecutionPreservesSeedsRestorationResultsAndZeroCorrelation(DistributionKind kind)
    {
        var input = InputFor(kind); var settings = new SimulationSettings(200, SimulationSeedMode.Fixed, 123);
        var book = new TestModel();
        var first = SimulationExecution.Run(settings, [input], [Output()], book).Result!;
        var second = SimulationExecution.Run(settings, [input], [Output()], book, correlations: []).Result!;
        Assert.Equal(first.ForecastResults[0].Values, second.ForecastResults[0].Values);
        Assert.All(first.ForecastResults[0].Values, x => Assert.True(double.IsFinite(x)));
        Assert.Equal("=50*2", book.A.Formula); Assert.Equal(100, book.A.Value);
        Assert.True(book.ScreenUpdating && book.EnableEvents); Assert.Equal("Ready", book.StatusBar);
        var random = new Random(123); var p = DistributionCatalog.Defaults(kind);
        var expected = Enumerable.Range(0, 200).Select(_ => DistributionSampler.Sample(kind, p[0], p[1], p[2], p[3], random, input.ProbabilityTable) * 100 / 100).ToArray();
        Assert.Equal(expected.Order(), first.ForecastResults[0].Values.Order());
    }
    [Theory, MemberData(nameof(Kinds))]
    public void SingularCorrelationsAndMixedMarginalsReachExecution(DistributionKind kind)
    {
        var first = InputFor(kind); var second = ScenarioSimulationService.Clone(first); second.CellLink = "b"; second.CellAddress = "A2";
        var book = new TestModel(); var observed = new List<(double, double)>();
        book.OnCalculate = () => observed.Add((book.A.Value, book.B.Value));
        var run = SimulationExecution.Run(new SimulationSettings(200, SimulationSeedMode.Fixed, 123), [first, second], [Output()], book, correlations: [new("a", "b", 1)]);
        Assert.NotNull(run.Result); Assert.All(observed, pair => Assert.Equal(pair.Item1, pair.Item2));
        var gaussian = new GaussianDependence(["a", "b"], [new("a", "b", .8)]); var z = new double[2]; var random = new Random(7);
        var x = new double[2000]; var y = new double[x.Length]; var p = DistributionCatalog.Defaults(kind);
        for (int i = 0; i < x.Length; i++) { gaussian.NextGaussian(random, z); x[i] = z[0]; y[i] = DistributionQuantile.FromGaussian(kind, z[1], p[0], p[1], p[2], p[3], first.ProbabilityTable); }
        Assert.True(CorrelationTests.Pearson(x, y) > .5);
        if (DistributionCatalog.IsDiscrete(kind)) Assert.True(y.Distinct().Count() < y.Length / 2);
        foreach (double extreme in new[] { -40d, 40d }) Assert.True(double.IsFinite(DistributionQuantile.FromGaussian(kind, extreme, p[0], p[1], p[2], p[3], first.ProbabilityTable)));
    }
    [Theory, MemberData(nameof(Kinds))]
    public void ScenariosUseCommonDrawsSelectedBaselineAndUnroundedSampleScaling(DistributionKind kind)
    {
        var input = InputFor(kind); var sourceTable = input.ProbabilityTable;
        var scenarios = Comparison(Scenario("Growth", Change("a", 50)), Scenario("Zero", Change("a", -100))) with { BaselineScenarioId = "Growth" };
        var result = ScenarioSimulationService.Run(new(400, SimulationSeedMode.Fixed, 5), [input], [Output()], scenarios, () => new TestSandbox());
        Assert.Equal(result.Rows[1].Mean * 1.5, result.Rows[0].Mean, 9);
        Assert.Equal(result.Rows[1].DecisionValue * 1.5, result.Rows[0].DecisionValue, 9);
        Assert.True(result.Rows[0].IsBaseline); Assert.Equal(0, result.Rows[2].Mean);
        Assert.Equal(-result.Rows[0].Mean, result.Rows[2].MeanDelta!.Value, 9);
        Assert.Same(sourceTable, input.ProbabilityTable);
        var p = DistributionCatalog.Defaults(kind); var distribution = new ScenarioDistribution(kind, p[0], p[1], p[2], p[3], ProbabilityTable: sourceTable);
        var transformed = ScenarioTransformation.Apply(distribution, 50);
        Assert.Equal(1.5, transformed.SampleScale); Assert.Equal(distribution.P1, transformed.P1); Assert.Equal(distribution.P2, transformed.P2);
        Assert.Equal(1.5, ScenarioTransformation.ScaleSample(1, 50));
        Assert.Throws<ArgumentException>(() => ScenarioTransformation.Apply(distribution, -101));
    }
    [Theory, MemberData(nameof(Kinds))]
    public void WorkbookRoundTripEditingReferenceTrackingAndDeletion(DistributionKind kind)
    {
        var app = new PersistenceApplication { ActiveWorkbook = new() }; ExcelDnaUtil.Application = app;
        try
        {
            SimulationModel.Clear(); var sheet = app.ActiveWorkbook.Worksheets.Add(); sheet.Name = "Model"; sheet.Range["A1"].Value2 = 100d;
            var input = InputFor(kind); input.CellLink = ""; SimulationModel.Assumptions.Add(input); SimulationModel.Forecasts.Add(Output());
            WorkbookPersistence.SaveModel(); app.ActiveWorkbook = app.ActiveWorkbook.Reopen(); SimulationModel.Clear();
            Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
            var restored = Assert.Single(SimulationModel.Assumptions); Assert.Equal(kind, (DistributionKind)restored.Distribution);
            Assert.Equal(input.ParameterDescription, restored.ParameterDescription); Assert.False(string.IsNullOrWhiteSpace(restored.CellLink));
            if (kind == DistributionKind.Discrete) Assert.Equal(input.ProbabilityTable!.Outcomes, restored.ProbabilityTable!.Outcomes);
            restored.Name = "Edited"; WorkbookPersistence.SaveModel();
            app.ActiveWorkbook.Worksheets["Model"].Range["A1"].Location = "C4";
            app.ActiveWorkbook.Worksheets["Model"].Name = "Renamed";
            Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
            restored = Assert.Single(SimulationModel.Assumptions); Assert.Equal("Edited", restored.Name); Assert.Equal("C4", restored.CellAddress); Assert.Equal("Renamed", restored.SheetName);
            ModelDefinitionDeletion.Delete(restored, (_, _) => true);
            Assert.Empty(SimulationModel.Assumptions);
        }
        finally { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    }
    [Theory]
    [InlineData(DistributionKind.Bernoulli)] [InlineData(DistributionKind.Binomial)]
    [InlineData(DistributionKind.Poisson)] [InlineData(DistributionKind.Discrete)]
    public void DiscreteResultsRetainInclusiveMassAndSensitivity(DistributionKind kind)
    {
        var run = SimulationExecution.Run(new SimulationSettings(500, SimulationSeedMode.Fixed, 19), [InputFor(kind)], [Output()], new TestModel()).Result!;
        var result = run.ForecastResults[0]; double target = result.Values.GroupBy(x => x).MaxBy(g => g.Count())!.Key;
        Assert.Equal(result.Values.Count(x => x <= target) / 500d, result.ProbabilityLessThanOrEqual(target));
        Assert.Equal(result.Values.Count(x => x >= target) / 500d, result.ProbabilityGreaterThanOrEqual(target));
        Assert.True(result.ProbabilityLessThanOrEqual(target) + result.ProbabilityGreaterThanOrEqual(target) > 1);
        Assert.Equal(1, Assert.Single(result.Sensitivities).Correlation, 12);
    }
    [Theory]
    [InlineData(null)] [InlineData("garbage")] [InlineData("{\"Version\":2,\"Outcomes\":[]}")]
    [InlineData("{\"Version\":1,\"Outcomes\":[{\"Outcome\":1,\"Probability\":0.2},{\"Outcome\":2,\"Probability\":0.2}]}")]
    public void InvalidSavedTableIsActionable(string? json)
    {
        var error = Assert.Throws<ArgumentException>(() => DistributionDataPersistence.Deserialize(DistributionType.Discrete, DistributionDataPersistence.Header, json));
        Assert.Contains("table", error.Message, StringComparison.OrdinalIgnoreCase);
    }
}

