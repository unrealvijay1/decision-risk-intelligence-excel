using ExcelDna.Integration;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class ScenarioPersistenceTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    public ScenarioPersistenceTests()
    {
        ExcelDnaUtil.Application = app; SimulationModel.Clear();
        Book.Worksheets.Add().Name = "Model";
        Book.Worksheets[1].Range["A1"].Value2 = 100d; Book.Worksheets[1].Range["B1"].Value2 = 100d;
        var input = ScenarioSimulationTests.Input(); input.CellLink = "";
        SimulationModel.Assumptions.Add(input); SimulationModel.Forecasts.Add(ScenarioSimulationTests.Output());
        WorkbookPersistence.SaveModel();
    }
    [Fact]
    public void PerForecastProbabilitiesRoundTripIndependently()
    {
        var analysis = new ScenarioAnalysisDefinition(Plans(), "1", ForecastQuestions: [
            new("f1", TargetDirection.AtOrAbove, Probability: 65),
            new("f2", TargetDirection.AtOrBelow, Probability: 95)]);
        WorkbookPersistence.SaveScenarioAnalysis(Book, analysis, [], []);
        var loaded = WorkbookPersistence.LoadScenarioAnalysis(Book.Reopen());
        Assert.Equal(65, loaded.ProbabilityFor(loaded.Questions[0]));
        Assert.Equal(95, loaded.ProbabilityFor(loaded.Questions[1]));
    }

    [Fact]
    public void OlderV2CommonProbabilityRemainsFallbackForEachForecast()
    {
        var sheet = Book.Worksheets["__MonteCarloConfig"];
        sheet.Cells[1, 22].Value2 = "ScenarioDefinitionsV2";
        sheet.Cells[1, 23].Value2 = 1;
        sheet.Cells[2, 22].Value2 = """{"Scenarios":[],"Probability":87.5,"ForecastQuestions":[{"ForecastId":"f1"},{"ForecastId":"f2"}]}""";
        var loaded = WorkbookPersistence.LoadScenarioAnalysis(Book.Reopen());
        Assert.All(loaded.Questions, q => Assert.Equal(87.5, loaded.ProbabilityFor(q)));
    }
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    private List<ScenarioDefinition> Plans() => [new("1", "Growth", [new(SimulationModel.Assumptions[0].CellLink, ScenarioAdjustmentType.PercentageChange, 15)]), new("2", "Empty", [])];
    [Theory]
    [InlineData(ScenarioAnalysisMode.ProbabilityToValue)] [InlineData(ScenarioAnalysisMode.ValueToProbability)]
    public void AnalysisV2RoundTripPreservesBothModesBaselineAndIndependentForecastSettings(ScenarioAnalysisMode mode)
    {
        var forecast = SimulationModel.Forecasts[0]; var normal = forecast.TargetSettings;
        var analysis = new ScenarioAnalysisDefinition(Plans(), "1", mode, 87.5,
            [new(ScenarioSimulationService.Identity(forecast), TargetDirection.AtOrAbove, 123.456789, TargetDirection.AtOrBelow)]);
        WorkbookPersistence.SaveScenarioAnalysis(Book, analysis, SimulationModel.Assumptions, SimulationModel.Forecasts);
        Assert.Equal("ScenarioDefinitionsV2", Book.Worksheets["__MonteCarloConfig"].Cells[1, 22].Value2);
        WorkbookPersistence.SaveModel(); app.ActiveWorkbook = Book.Reopen();
        var saved = WorkbookPersistence.LoadScenarioAnalysis(Book);
        Assert.Equal("1", saved.BaselineScenarioId); Assert.Equal(mode, saved.Mode); Assert.Equal(87.5, saved.Probability);
        Assert.Equal(analysis.Questions, saved.Questions); Assert.Equal(normal, forecast.TargetSettings);
        WorkbookPersistence.SaveScenarios(Book, saved.Scenarios.Select(s => s.Id == "1" ? s with { Name = "Renamed" } : s).ToArray(), []);
        Assert.Equal("1", WorkbookPersistence.LoadScenarioAnalysis(Book).BaselineScenarioId);
        WorkbookPersistence.SaveScenarios(Book, [saved.Scenarios[1]], []);
        Assert.Null(WorkbookPersistence.LoadScenarioAnalysis(Book).BaselineScenarioId);
    }
    [Fact]
    public void V1MigrationNeverChoosesBusinessBaselineOrForecastDirections()
    {
        var sheet = Book.Worksheets["__MonteCarloConfig"];
        sheet.Cells[1, 22].Value2 = "ScenarioDefinitionsV1"; sheet.Cells[1, 23].Value2 = 1;
        sheet.Cells[2, 22].Value2 = System.Text.Json.JsonSerializer.Serialize(Plans());
        var loaded = WorkbookPersistence.LoadScenarioAnalysis(Book);
        Assert.Equal(2, loaded.Scenarios.Count); Assert.Null(loaded.BaselineScenarioId); Assert.Empty(loaded.Questions);
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(loaded, SimulationModel.Forecasts));
        WorkbookPersistence.SaveScenarioAnalysis(Book, loaded, SimulationModel.Assumptions, SimulationModel.Forecasts);
        Assert.Null(WorkbookPersistence.LoadScenarioAnalysis(Book).BaselineScenarioId);
    }
    [Fact]
    public void MissingForecastConfigurationIsRetainedForCorrection()
    {
        var analysis = new ScenarioAnalysisDefinition(Plans(), "1", ForecastQuestions: [new("deleted", TargetDirection.AtOrAbove)]);
        WorkbookPersistence.SaveScenarioAnalysis(Book, analysis, SimulationModel.Assumptions, []);
        var loaded = WorkbookPersistence.LoadScenarioAnalysis(Book);
        Assert.Equal("deleted", Assert.Single(loaded.Questions).ForecastId);
        Assert.Contains("missing", Assert.Throws<ArgumentException>(() => ScenarioSimulationService.ValidateAnalysis(loaded, SimulationModel.Forecasts)).Message);
    }
    [Fact]
    public void ForecastReferenceTracksRenameAndLegacyUpgrade()
    {
        var forecast = SimulationModel.Forecasts[0]; forecast.CellLink = "";
        Book.Worksheets["__MonteCarloConfig"].Cells[3, 10].Value2 = null;
        var analysis = new ScenarioAnalysisDefinition(Plans(), "1", ForecastQuestions: [new(ScenarioSimulationService.Identity(forecast), TargetDirection.AtOrAbove)]);
        var saved = WorkbookPersistence.SaveScenarioAnalysis(Book, analysis, SimulationModel.Assumptions, [forecast]);
        Assert.StartsWith("_MC_Cell_", saved.Questions[0].ForecastId);
        Book.Worksheets[1].Name = "Renamed"; Book.Worksheets[1].Range["B1"].Location = "D4";
        var validation = WorkbookPersistence.LoadModelForSimulation(false); Assert.True(validation.IsValid, validation.UserMessage);
        ScenarioSimulationService.ValidateAnalysis(WorkbookPersistence.LoadScenarioAnalysis(Book), SimulationModel.Forecasts);
        Assert.Equal("D4", SimulationModel.Forecasts[0].CellAddress);
    }
    [Fact]
    public void LegacyWorkbookWithoutScenariosLoadsNormally() => Assert.Empty(WorkbookPersistence.LoadScenarios(Book));
    [Fact]
    public void RoundTripMultipleScenariosRenameDeleteAndNoResults()
    {
        var plans = Plans(); WorkbookPersistence.SaveScenarios(Book, plans, SimulationModel.Assumptions);
        app.ActiveWorkbook = Book.Reopen(); var restored = WorkbookPersistence.LoadScenarios(Book);
        Assert.Equal("Growth", restored[0].Name); Assert.Equal(plans[0].Adjustments, restored[0].Adjustments); Assert.Empty(restored[1].Adjustments);
        WorkbookPersistence.SaveScenarios(Book, [restored[0] with { Name = "Upside" }], SimulationModel.Assumptions);
        Assert.Equal("Upside", Assert.Single(WorkbookPersistence.LoadScenarios(Book)).Name);
        WorkbookPersistence.SaveScenarios(Book, [], SimulationModel.Assumptions); Assert.Empty(WorkbookPersistence.LoadScenarios(Book));
        Assert.Equal(100d, Book.Worksheets[1].Range["A1"].Value2);
    }
    [Fact]
    public void DefinitionSettingsSavesAndClearPreserveScenarios()
    {
        var plans = Plans(); WorkbookPersistence.SaveScenarios(Book, plans, SimulationModel.Assumptions);
        WorkbookPersistence.SaveSimulationSettings(new(42, SimulationSeedMode.Fixed, 7));
        SimulationModel.Assumptions[0].Name = "Renamed"; WorkbookPersistence.SaveModel();
        Assert.Equal(plans[0].Adjustments, WorkbookPersistence.LoadScenarios(Book)[0].Adjustments);
        WorkbookPersistence.ClearSavedModel();
        Assert.Equal(2, WorkbookPersistence.LoadScenarios(Book).Count);
        Assert.Throws<ArgumentException>(() => ScenarioSimulationService.Validate(SimulationModel.Assumptions, WorkbookPersistence.LoadScenarios(Book)));
        Assert.Equal(42, WorkbookPersistence.LoadSimulationSettings(new()).TrialCount);
    }
    [Fact]
    public void TrackedIdentitySurvivesRenameMoveAndDeletedReferenceIsFriendly()
    {
        WorkbookPersistence.SaveScenarios(Book, Plans(), SimulationModel.Assumptions);
        Book.Worksheets[1].Name = "Renamed"; Book.Worksheets[1].Range["A1"].Location = "C5";
        var validation = WorkbookPersistence.LoadModelForSimulation(false); Assert.True(validation.IsValid, validation.UserMessage);
        ScenarioSimulationService.Validate(SimulationModel.Assumptions, WorkbookPersistence.LoadScenarios(Book));
        Assert.Equal("C5", SimulationModel.Assumptions[0].CellAddress);
        SimulationModel.Assumptions.Clear(); WorkbookPersistence.SaveModel();
        var error = Assert.Throws<ArgumentException>(() => ScenarioSimulationService.Validate(SimulationModel.Assumptions, WorkbookPersistence.LoadScenarios(Book)));
        Assert.Contains("missing", error.Message);
    }
    [Fact]
    public void LegacyCoordinateIdentitiesUpgradeWithoutChangingParametersOrFill()
    {
        var a = SimulationModel.Assumptions[0];
        a.CellLink = ""; Book.Worksheets["__MonteCarloConfig"].Cells[2, 10].Value2 = null;
        var cell = Book.Worksheets[1].Range["A1"]; double fill = cell.Interior.Color;
        var saved = WorkbookPersistence.SaveScenarios(Book,
            [new("1", "Legacy", [new(ScenarioSimulationService.Identity(a), ScenarioAdjustmentType.PercentageChange, 10)])], [a]);
        Assert.StartsWith("_MC_Cell_", saved[0].Adjustments[0].AssumptionId);
        Assert.Equal(saved[0].Adjustments[0].AssumptionId, a.CellLink);
        Assert.Equal(100, a.Parameter1); Assert.Equal(fill, cell.Interior.Color); Assert.Equal(100d, cell.Value2);
    }
    [Theory]
    [InlineData("UnknownVersion")] [InlineData("ScenarioDefinitionsV1")]
    public void InvalidScenarioBlockDoesNotBreakNormalSimulationOrGetErased(string marker)
    {
        var sheet = Book.Worksheets["__MonteCarloConfig"];
        sheet.Cells[1, 22].Value2 = marker; sheet.Cells[1, 23].Value2 = 1; sheet.Cells[2, 22].Value2 = "corrupt";
        Assert.Throws<InvalidOperationException>(() => WorkbookPersistence.LoadScenarios(Book));
        WorkbookPersistence.SaveModel();
        Assert.Equal("corrupt", sheet.Cells[2, 22].Value2);
        Assert.True(SimulationService.TryRun(3).Succeeded);
    }
    [Fact]
    public void ChunkedPersistenceExceedsOneCellWithoutTruncation()
    {
        var plans = Enumerable.Range(0, 500).Select(i => new ScenarioDefinition(i.ToString(), "Scenario " + i,
            new[] { new ScenarioAdjustment("_MC_Cell_" + new string('x', 50) + i, ScenarioAdjustmentType.PercentageChange, i) })).ToArray();
        WorkbookPersistence.SaveScenarios(Book, plans, []);
        Assert.True(Convert.ToInt32(Book.Worksheets["__MonteCarloConfig"].Cells[1, 23].Value2) > 1);
        app.ActiveWorkbook = Book.Reopen(); Assert.Equal(500, WorkbookPersistence.LoadScenarios(Book).Count);
        WorkbookPersistence.SaveScenarios(Book, [], []); Assert.Empty(WorkbookPersistence.LoadScenarios(Book));
    }
    [Fact]
    public void SavePinsOriginalWorkbookAndRejectsInvalidNamesBeforeMutation()
    {
        var source = Book; var other = new PersistenceWorkbook(); other.Worksheets.Add(); app.ActiveWorkbook = other;
        WorkbookPersistence.SaveScenarios(source, Plans(), SimulationModel.Assumptions);
        Assert.Empty(WorkbookPersistence.LoadScenarios(other)); Assert.Equal(2, WorkbookPersistence.LoadScenarios(source).Count);
        Assert.Throws<ArgumentException>(() => WorkbookPersistence.SaveScenarios(source, [new("1", " ", [])], []));
        Assert.Equal(2, WorkbookPersistence.LoadScenarios(source).Count);
    }
}
