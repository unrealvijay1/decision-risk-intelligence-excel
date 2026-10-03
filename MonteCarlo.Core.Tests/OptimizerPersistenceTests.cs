using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel;
namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class OptimizerPersistenceTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    public OptimizerPersistenceTests()
    { ExcelDnaUtil.Application = app; SimulationModel.Clear(); Book.Worksheets.Add().Name = "Model"; Book.Worksheets[1].Range["C1"].Value2 = 50d; }
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    [Fact] public void SaveReopenRoundTripAllFieldsWithoutCandidateValues()
    {
        var p = OptimizationTests.Problem() with { Objective = new("forecast", OptimizationMetric.Probability, false, 123, TargetDirection.AtOrBelow),
            Constraints = [new(OptimizationConstraintKind.Probability, "forecast", TargetDirection.AtOrAbove, 80, 1000), new(OptimizationConstraintKind.DecisionVariable, "x", TargetDirection.AtOrBelow, 90)] };
        var saved = WorkbookPersistence.SaveOptimizer(Book, p); app.ActiveWorkbook = Book.Reopen(); var loaded = WorkbookPersistence.LoadOptimizer(Book)!;
        Assert.Equal(p.Objective, loaded.Objective); Assert.Equal(p.Settings, loaded.Settings); Assert.Equal(p.Constraints, loaded.Constraints); Assert.Equal(saved.Variables, loaded.Variables); Assert.StartsWith("_MC_Cell_", loaded.Variables[0].CellLink);
        Assert.Equal(50d, Book.Worksheets[1].Range["C1"].Value2); Assert.Equal(2, Book.Worksheets["__MonteCarloConfig"].Visible);
    }
    [Fact] public void ModelSaveAndClearPreserveOptimizerBlock()
    {
        var saved = WorkbookPersistence.SaveOptimizer(Book, OptimizationTests.Problem()); WorkbookPersistence.SaveModel(); Assert.Equal(saved.Variables, WorkbookPersistence.LoadOptimizer(Book)!.Variables);
        WorkbookPersistence.ClearSavedModel(); Assert.NotNull(WorkbookPersistence.LoadOptimizer(Book));
    }
    [Fact] public void InvalidEnvelopePreservedByModelSave()
    {
        WorkbookPersistence.SaveOptimizer(Book, OptimizationTests.Problem()); var sheet = Book.Worksheets["__MonteCarloConfig"]; sheet.Cells[1, 30].Value2 = "future";
        Assert.Throws<InvalidOperationException>(() => WorkbookPersistence.LoadOptimizer(Book)); WorkbookPersistence.SaveModel(); Assert.Equal("future", sheet.Cells[1, 30].Value2);
    }
    [Fact] public void ClearPreservesMalformedBlockWithMissingMarker()
    {
        WorkbookPersistence.SaveOptimizer(Book, OptimizationTests.Problem()); var sheet = Book.Worksheets["__MonteCarloConfig"]; sheet.Cells[1, 30].Value2 = null;
        WorkbookPersistence.ClearSavedModel(); Assert.NotNull(Book.Worksheets["__MonteCarloConfig"].Cells[2, 30].Value2);
        Assert.Throws<InvalidOperationException>(() => WorkbookPersistence.LoadOptimizer(Book));
    }
    [Fact] public void TrackedRenameAndMoveResolvedWithoutCoordinateFallback()
    {
        var saved = WorkbookPersistence.SaveOptimizer(Book, OptimizationTests.Problem()); var cell = Book.Worksheets[1].Range["C1"]; cell.Location = "D4"; cell.Worksheet.Name = "Renamed";
        var loaded = WorkbookPersistence.LoadOptimizer(Book)!; Assert.Equal("Renamed", loaded.Variables[0].Sheet); Assert.Equal("D4", loaded.Variables[0].Cell);
        cell.Deleted = true; loaded = WorkbookPersistence.LoadOptimizer(Book)!; Assert.Equal("#REF!", loaded.Variables[0].Cell);
        Assert.Throws<ArgumentException>(() => WorkbookPersistence.SaveOptimizer(Book, loaded)); Assert.Equal(saved.Variables[0].CellLink, loaded.Variables[0].CellLink);
    }
    [Fact] public void JsonNullAndMalformedChunksFailFriendly()
    {
        WorkbookPersistence.SaveOptimizer(Book, OptimizationTests.Problem()); var sheet = Book.Worksheets["__MonteCarloConfig"]; sheet.Cells[2, 30].Value2 = "null";
        Assert.Throws<InvalidOperationException>(() => WorkbookPersistence.LoadOptimizer(Book)); sheet.Cells[2, 30].Value2 = "{"; Assert.Throws<InvalidOperationException>(() => WorkbookPersistence.LoadOptimizer(Book));
    }
    [Fact] public void LegacyForecastIdentityUpgradesOnModelSave()
    {
        var f = OptimizationTests.Forecast(); f.CellLink = ""; SimulationModel.Forecasts.Add(f); WorkbookPersistence.SaveModel();
        f.CellLink = ""; Book.Worksheets["__MonteCarloConfig"].Cells[2, 10].Value2 = null;
        var p = OptimizationTests.Problem() with { Objective = new(ScenarioSimulationService.Identity(f)), Constraints = [new(OptimizationConstraintKind.Mean, ScenarioSimulationService.Identity(f), TargetDirection.AtOrAbove, 0)] };
        WorkbookPersistence.SaveOptimizer(Book, p, [f]); WorkbookPersistence.SaveModel();
        Assert.StartsWith("_MC_Cell_", f.CellLink); Assert.Equal(f.CellLink, WorkbookPersistence.LoadOptimizer(Book)!.Objective.ForecastId);
        Assert.Equal(f.CellLink, WorkbookPersistence.LoadOptimizer(Book)!.Constraints[0].SubjectId);
    }
}
