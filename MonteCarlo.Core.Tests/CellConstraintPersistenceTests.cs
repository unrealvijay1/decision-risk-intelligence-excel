using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class CellConstraintPersistenceTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    public CellConstraintPersistenceTests()
    { ExcelDnaUtil.Application = app; SimulationModel.Clear(); var model = Book.Worksheets.Add(); model.Name = "Model"; model.Range["C1"].Value2 = 50d; model.Range["C2"].Value2 = 2d; }
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    private static OptimizationProblem Problem() => OptimizationTests.Problem() with { Constraints = [CellConstraintTests.Cell("C2", 2, OptimizationConstraintOperator.Equal, "Labor Used"), new(OptimizationConstraintKind.Mean, "forecast", TargetDirection.AtOrAbove, 0)] };
    [Fact] public void LinearTermsPersistByVariableIdentityAndSurviveReopen()
    {
        var p = OptimizationTests.Problem() with { Constraints = [IterativeOptimizationTests.Linear(40, OptimizationConstraintOperator.Equal, ("x", 2))] };
        var saved = WorkbookPersistence.SaveOptimizer(Book, p); app.ActiveWorkbook = Book.Reopen();
        var loaded = WorkbookPersistence.LoadOptimizer(Book)!;
        Assert.Equal(saved.Constraints[0].Terms, loaded.Constraints[0].Terms);
        Assert.Equal(OptimizationConstraintKind.Linear, loaded.Constraints[0].Kind);
        var r = OptimizationService.Run(loaded, [], [OptimizationTests.Forecast()], () => new OptimizationTests.Book());
        Assert.Equal(20, r.Best.Values[0]); Assert.NotNull(r.BestFeasible);
    }
    [Fact] public void ZeroAssumptionConfigurationReopensAndRunsDeterministically()
    {
        var saved = WorkbookPersistence.SaveOptimizer(Book, Problem()); app.ActiveWorkbook = Book.Reopen();
        var loaded = WorkbookPersistence.LoadOptimizer(Book)!;
        Assert.Empty(SimulationModel.Assumptions); Assert.Equal(saved.Settings, loaded.Settings);
        var b = new OptimizationTests.Book();
        var result = OptimizationService.Run(loaded, [], [OptimizationTests.Forecast()], () => b);
        Assert.Equal(OptimizationEvaluationMode.Deterministic, result.EvaluationMode); Assert.Equal(100, result.Best.Values[0]);
        Assert.Equal(50d, b.Cells["C1"].Value);
    }
    [Fact] public void SaveReopenPersistsAllCellFieldsAndEqualityAlongsideLegacyConstraints()
    {
        var saved = WorkbookPersistence.SaveOptimizer(Book, Problem()); app.ActiveWorkbook = Book.Reopen();
        var loaded = WorkbookPersistence.LoadOptimizer(Book)!;
        Assert.Equal(saved.Constraints, loaded.Constraints); Assert.StartsWith("_MC_Cell_", loaded.Constraints[0].CellLink);
        Assert.Equal("Labor Used", loaded.Constraints[0].Name); Assert.Equal(OptimizationConstraintOperator.Equal, loaded.Constraints[0].EffectiveOperator);
        Assert.Equal(2d, Book.Worksheets[1].Range["C2"].Value2);
    }
    [Fact] public void SavingAgainReusesTrackedCellNameAndPreservesModelSaveClear()
    {
        var saved = WorkbookPersistence.SaveOptimizer(Book, Problem()); int names = Book.Names.Items.Count;
        saved = WorkbookPersistence.SaveOptimizer(Book, saved); Assert.Equal(names, Book.Names.Items.Count);
        WorkbookPersistence.SaveModel(); Assert.Equal(saved.Constraints, WorkbookPersistence.LoadOptimizer(Book)!.Constraints);
        WorkbookPersistence.ClearSavedModel(); Assert.Equal(saved.Constraints, WorkbookPersistence.LoadOptimizer(Book)!.Constraints);
    }
    [Fact] public void CellConstraintTracksRenameMovementAndDeletionWithoutRebinding()
    {
        var saved = WorkbookPersistence.SaveOptimizer(Book, Problem()); var cell = Book.Worksheets[1].Range["C2"];
        cell.Location = "D9"; cell.Worksheet.Name = "Renamed";
        var loaded = WorkbookPersistence.LoadOptimizer(Book)!; Assert.Equal("Renamed", loaded.Constraints[0].Sheet); Assert.Equal("D9", loaded.Constraints[0].Cell);
        cell.Deleted = true; loaded = WorkbookPersistence.LoadOptimizer(Book)!; Assert.Equal("#REF!", loaded.Constraints[0].Cell);
        Assert.Equal(saved.Constraints[0].CellLink, loaded.Constraints[0].CellLink);
        Assert.Throws<ArgumentException>(() => WorkbookPersistence.SaveOptimizer(Book, loaded));
    }
    [Fact] public void ForeignTrackedCellIsReportedAsBroken()
    {
        var saved = WorkbookPersistence.SaveOptimizer(Book, Problem()); var foreign = new PersistenceWorkbook { Name = "Foreign.xlsx" };
        Book.Names.Items[saved.Constraints[0].CellLink] = new(foreign.Worksheets.Add().Range["C2"], false);
        Assert.Equal("#REF!", WorkbookPersistence.LoadOptimizer(Book)!.Constraints[0].Cell);
    }
    [Fact] public void MissingOptionalFieldsInHistoricalV1JsonKeepOriginalDirections()
    {
        var p = OptimizationTests.Problem() with { Constraints = [new(OptimizationConstraintKind.DecisionVariable, "x", TargetDirection.AtOrBelow, 40), new(OptimizationConstraintKind.Probability, "forecast", TargetDirection.AtOrAbove, 80, 100)] };
        WorkbookPersistence.SaveOptimizer(Book, p); var sheet = Book.Worksheets["__MonteCarloConfig"];
        var json = JsonNode.Parse((string)sheet.Cells[2, 30].Value2!)!;
        foreach (var c in json["Constraints"]!.AsArray()) foreach (string field in new[] { "Operator", "Sheet", "Cell", "Name", "CellLink", "EffectiveOperator", "CellDescription" }) c!.AsObject().Remove(field);
        sheet.Cells[2, 30].Value2 = json.ToJsonString();
        var loaded = WorkbookPersistence.LoadOptimizer(Book)!; Assert.Equal(p.Constraints, loaded.Constraints);
        Assert.Equal(OptimizationConstraintOperator.LessThanOrEqual, loaded.Constraints[0].EffectiveOperator);
        Assert.Equal(OptimizationConstraintOperator.GreaterThanOrEqual, loaded.Constraints[1].EffectiveOperator);
    }
    [Fact] public void UnsupportedOperatorBlocksOptimizerWithoutDestroyingSavedBlock()
    {
        WorkbookPersistence.SaveOptimizer(Book, Problem()); var sheet = Book.Worksheets["__MonteCarloConfig"];
        var json = JsonNode.Parse((string)sheet.Cells[2, 30].Value2!)!; json["Constraints"]![0]!["Operator"] = 999;
        string content = json.ToJsonString(); sheet.Cells[2, 30].Value2 = content;
        Assert.Throws<InvalidOperationException>(() => WorkbookPersistence.LoadOptimizer(Book)); WorkbookPersistence.SaveModel();
        Assert.Equal(content, Book.Worksheets["__MonteCarloConfig"].Cells[2, 30].Value2);
    }
    [Fact] public void ProductMixConstraintsPersistWithoutRegisteringLaborOrMaterialAsForecasts()
    {
        var (_, p, f) = CellConstraintTests.ProductMix(); var model = Book.Worksheets[1];
        model.Range["B2"].Value2 = 10d; model.Range["B3"].Value2 = 10d; model.Range["B6"].Value2 = 60d; model.Range["B7"].Value2 = 50d;
        var saved = WorkbookPersistence.SaveOptimizer(Book, p); app.ActiveWorkbook = Book.Reopen();
        var loaded = WorkbookPersistence.LoadOptimizer(Book)!; Assert.Equal(saved.Constraints, loaded.Constraints); Assert.Equal(saved.Variables, loaded.Variables);
        Assert.Equal(new[] { "B6", "B7", "B2" }, loaded.Constraints.Where(c => c.Kind == OptimizationConstraintKind.Cell).Select(c => c.Cell));
        Assert.Empty(SimulationModel.Forecasts);
    }
}
