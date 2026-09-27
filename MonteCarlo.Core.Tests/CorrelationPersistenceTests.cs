using ExcelDna.Integration;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class CorrelationPersistenceTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    public CorrelationPersistenceTests()
    {
        ExcelDnaUtil.Application = app; SimulationModel.Clear(); Book.Worksheets.Add().Name = "Model";
        foreach (string address in new[] { "A1", "A2", "B1" }) Book.Worksheets[1].Range[address].Value2 = 100d;
        var a = ScenarioSimulationTests.Input(); var b = ScenarioSimulationTests.Input("b", "A2"); a.CellLink = b.CellLink = "";
        SimulationModel.Assumptions.AddRange([a, b]); SimulationModel.Forecasts.Add(ScenarioSimulationTests.Output()); WorkbookPersistence.SaveModel();
    }
    private AssumptionCorrelation Pair(double coefficient = .8) => new(SimulationModel.Assumptions[0].CellLink, SimulationModel.Assumptions[1].CellLink, coefficient);
    private void Save(double coefficient = .8) => WorkbookPersistence.SaveCorrelations(Book, SimulationModel.Assumptions, [Pair(coefficient)]);
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    [Fact]
    public void AbsentBlockIsIndependent() { Assert.Empty(WorkbookPersistence.LoadCorrelations(Book)); Assert.True(SimulationService.TryRun(10).Succeeded); }
    [Fact]
    public void SaveReopenRestoreAndNormalSimulationUseCorrelations()
    {
        Save(1); app.ActiveWorkbook = Book.Reopen(); Assert.Equal(1, Assert.Single(WorkbookPersistence.LoadCorrelations(Book)).Correlation);
        Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
        var result = SimulationService.TryRun(100); Assert.True(result.Succeeded, result.UserMessage);
    }
    [Fact]
    public void RenameMoveRetainsIdentity()
    {
        Save(); var pair = Pair(); Book.Worksheets[1].Name = "Renamed"; Book.Worksheets[1].Range["A1"].Location = "C3";
        Assert.True(WorkbookPersistence.LoadModelForSimulation(false).IsValid);
        Assert.Equal(pair, Assert.Single(WorkbookPersistence.LoadCorrelations(Book)));
        Assert.Equal("C3", SimulationModel.Assumptions[0].CellAddress);
    }
    [Fact]
    public void DeleteRemovesRelationshipsAndClearModelRemovesBlock()
    {
        Save(); ModelDefinitionDeletion.Delete(SimulationModel.Assumptions[0], (_, _) => true);
        Assert.Empty(WorkbookPersistence.LoadCorrelations(Book));
        WorkbookPersistence.SaveSimulationSettings(new(12)); WorkbookPersistence.ClearSavedModel();
        Assert.Empty(WorkbookPersistence.LoadCorrelations(Book)); Assert.Equal(12, WorkbookPersistence.LoadSimulationSettings(new()).TrialCount);
    }
    [Fact]
    public void ZeroRemovesRelationshipAndEditingPreservesOtherBlocks()
    {
        WorkbookPersistence.SaveScenarios(Book, [new("s", "Plan", [])], SimulationModel.Assumptions);
        WorkbookPersistence.SaveSimulationSettings(new(21)); Save();
        SimulationModel.Assumptions[0].Name = "Renamed input"; WorkbookPersistence.SaveModel();
        Assert.Equal(.8, Assert.Single(WorkbookPersistence.LoadCorrelations(Book)).Correlation);
        Save(-.4); Assert.Equal(-.4, Assert.Single(WorkbookPersistence.LoadCorrelations(Book)).Correlation);
        Save(0); Assert.Empty(WorkbookPersistence.LoadCorrelations(Book));
        Assert.Single(WorkbookPersistence.LoadScenarios(Book)); Assert.Equal(21, WorkbookPersistence.LoadSimulationSettings(new()).TrialCount);
        Assert.Equal(100d, Book.Worksheets[1].Range["A1"].Value2);
    }
    [Fact]
    public void MissingReferenceWarnsAndCanBeRemoved()
    {
        Save(); Book.Worksheets["__MonteCarloConfig"].Cells[2, 10].Value2 = "_MC_Cell_missing";
        var validation = WorkbookPersistence.LoadModelForSimulation(false); Assert.False(validation.IsValid);
        Assert.Contains("correlated assumption", validation.UserMessage, StringComparison.OrdinalIgnoreCase);
        WorkbookPersistence.SaveCorrelations(Book, SimulationModel.Assumptions, []); Assert.Empty(WorkbookPersistence.LoadCorrelations(Book));
    }
    [Fact]
    public void CorruptMetadataBlocksRunButUnrelatedSavesPreserveIt()
    {
        var sheet = Book.Worksheets["__MonteCarloConfig"]; sheet.Cells[1, 25].Value2 = "Unknown";
        Assert.False(SimulationService.TryRun(10).Succeeded); WorkbookPersistence.SaveModel(); Assert.Equal("Unknown", sheet.Cells[1, 25].Value2);
    }
    [Fact]
    public void InvalidSaveLeavesPriorConfiguration()
    {
        Save(); Assert.Throws<ArgumentException>(() => Save(2)); Assert.Equal(.8, Assert.Single(WorkbookPersistence.LoadCorrelations(Book)).Correlation);
    }
    [Fact]
    public void LegacyCoordinatesUpgradeAndSaveIsPinned()
    {
        var a = SimulationModel.Assumptions[0]; var b = SimulationModel.Assumptions[1]; a.CellLink = b.CellLink = "";
        Book.Worksheets["__MonteCarloConfig"].Cells[2, 10].Value2 = null; Book.Worksheets["__MonteCarloConfig"].Cells[3, 10].Value2 = null;
        var source = Book; app.ActiveWorkbook = new(); app.ActiveWorkbook.Worksheets.Add();
        WorkbookPersistence.SaveCorrelations(source, [a, b], [new(ScenarioSimulationService.Identity(a), ScenarioSimulationService.Identity(b), .4)]);
        Assert.StartsWith("_MC_Cell_", Assert.Single(WorkbookPersistence.LoadCorrelations(source)).AssumptionA);
        Assert.Empty(WorkbookPersistence.LoadCorrelations(Book));
    }
}
