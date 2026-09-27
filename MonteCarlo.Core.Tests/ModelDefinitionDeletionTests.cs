using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel;
using System.Text.Json;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class ModelDefinitionDeletionTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    public ModelDefinitionDeletionTests()
    {
        ExcelDnaUtil.Application = app; SimulationModel.Clear();
        var sheet = Book.Worksheets.Add(); sheet.Name = "Model";
        foreach (string address in new[] { "B2", "B3", "B4", "B5" })
        {
            var cell = sheet.Range[address]; cell.Value2 = 42d; cell.Formula = "=6*7";
            cell.NumberFormat = "0.00"; cell.Interior.Pattern = 1; cell.Interior.ColorIndex = 6; cell.Interior.Color = 65535;
        }
        SimulationModel.Assumptions.Add(new() { Name = "First", SheetName = "Model", CellAddress = "B2", Distribution = DistributionType.Uniform, Parameter1 = 1, Parameter2 = 2 });
        SimulationModel.Assumptions.Add(new() { Name = "Other", SheetName = "Model", CellAddress = "B4", Distribution = DistributionType.Uniform, Parameter1 = 3, Parameter2 = 4 });
        SimulationModel.Forecasts.Add(new() { Name = "Total", SheetName = "Model", CellAddress = "B3", TargetSettings = new(40, TargetDirection.AtOrAbove, 80) });
        SimulationModel.Forecasts.Add(new() { Name = "Other total", SheetName = "Model", CellAddress = "B5" });
        WorkbookPersistence.SaveModel();
    }
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    private string Snapshot() => JsonSerializer.Serialize(Book.Worksheets.Items.Select(s => new {
        s.Name, s.Visible, Cells = s.Cells.Items.OrderBy(x => x.Key.Row).ThenBy(x => x.Key.Column)
        .Where(x => x.Value.Value2 != null || s.Name != "__MonteCarloConfig")
        .Select(x => new { x.Key.Row, x.Key.Column, x.Value.Value2, x.Value.Formula, x.Value.NumberFormat, x.Value.Interior }) }));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectedDefinitionIsConfirmedDeletedAndStaysDeleted(bool forecast)
    {
        string address = forecast ? "B3" : "B2";
        string kind = forecast ? "Forecast" : "Assumption";
        var selected = Book.Worksheets[1].Range[address];
        var otherAssumption = JsonSerializer.Serialize(SimulationModel.Assumptions[1]);
        var otherForecast = JsonSerializer.Serialize(SimulationModel.Forecasts[1]);
        int calls = 0;
        Assert.Equal(DeleteDefinitionResult.Deleted, ModelDefinitionDeletion.ClearSelectedCell(selected, (title, message) => {
            calls++; Assert.Equal($"Remove {kind}?", title); Assert.Contains($"Model!{address}", message);
            Assert.Contains("value/formula will not be changed", message); return true;
        }));
        Assert.Equal(1, calls);
        Assert.Equal(forecast ? 2 : 1, SimulationModel.Assumptions.Count);
        Assert.Equal(forecast ? 1 : 2, SimulationModel.Forecasts.Count);
        Assert.Equal(otherAssumption, JsonSerializer.Serialize(SimulationModel.Assumptions.Last()));
        Assert.Equal(otherForecast, JsonSerializer.Serialize(SimulationModel.Forecasts.Last()));
        Assert.Equal(65535, selected.Interior.Color); Assert.Equal(42d, selected.Value2); Assert.Equal("=6*7", selected.Formula);
        Assert.Equal("0.00", selected.NumberFormat); Assert.Equal("Original font", selected.Font);
        Assert.Equal("Original borders", selected.Borders); Assert.Equal("Original alignment", selected.Alignment); Assert.Equal("Original protection", selected.Protection);
        Assert.Equal(ModelCellHighlight.AssumptionColor, Book.Worksheets[1].Range["B4"].Interior.Color);
        Assert.Equal(ModelCellHighlight.ForecastColor, Book.Worksheets[1].Range["B5"].Interior.Color);
        app.ActiveWorkbook = Book.Reopen(); WorkbookPersistence.LoadModel();
        Assert.Null(forecast ? (object?)SimulationModel.FindForecast("Model", address) : SimulationModel.FindAssumption("Model", address));
        if (forecast) Assert.All(SimulationModel.Forecasts, f => Assert.Null(f.TargetSettings));
        Assert.DoesNotContain(Book.Worksheets["__MonteCarloConfig"].Cells.Items.Values, c => Equals(c.Value2, address));
    }

    [Theory]
    [InlineData("B2")]
    [InlineData("B3")]
    public void CancelDoesNotMigrateLegacyFillOrChangeMemory(string address)
    {
        Book.Worksheets["__MonteCarloConfig"].Cells[1, 15].Value2 = null;
        var selected = Book.Worksheets[1].Range[address]; selected.Interior.Color = 123;
        string before = Snapshot(); var assumption = SimulationModel.Assumptions[0]; var forecast = SimulationModel.Forecasts[0];
        Assert.Equal(DeleteDefinitionResult.Cancelled, ModelDefinitionDeletion.ClearSelectedCell(selected, (_, _) => false));
        Assert.Equal(before, Snapshot()); Assert.Same(assumption, SimulationModel.Assumptions[0]); Assert.Same(forecast, SimulationModel.Forecasts[0]);
    }

    [Fact]
    public void OrdinaryCellChangesNothing()
    {
        var cell = Book.Worksheets[1].Range["C9"]; string before = Snapshot();
        Assert.Equal(DeleteDefinitionResult.NoDefinition, ModelDefinitionDeletion.ClearSelectedCell(cell, (_, _) => throw new Exception("Unexpected confirmation")));
        Assert.Equal(before, Snapshot());
    }
    [Theory]
    [InlineData(2, 1)]
    [InlineData(2, 2)]
    [InlineData(1, 2)]
    public void MultiSelectionRejectedBeforeLoading(int cells, int areas)
    {
        string before = Snapshot();
        Assert.Equal(DeleteDefinitionResult.InvalidSelection, ModelDefinitionDeletion.ClearSelectedCell(new SelectionDouble(cells, areas), (_, _) => throw new Exception()));
        Assert.Equal(before, Snapshot());
    }
    [Fact]
    public void NonCellSelectionRejected() => Assert.Equal(DeleteDefinitionResult.InvalidSelection,
        ModelDefinitionDeletion.ClearSelectedCell(new object(), (_, _) => throw new Exception()));

    [Fact]
    public void TrackedRenameAndMoveCanBeCleared()
    {
        var cell = Book.Worksheets[1].Range["B2"]; cell.Location = "D9"; cell.Worksheet.Name = "Renamed";
        Assert.Equal(DeleteDefinitionResult.Deleted, ModelDefinitionDeletion.ClearSelectedCell(cell, (_, message) => { Assert.Contains("Renamed!D9", message); return true; }));
        Assert.Single(SimulationModel.Assumptions); Assert.Equal(65535, cell.Interior.Color);
    }
    [Fact]
    public void ModelManagerUsesSameDeletePath()
    {
        var definition = SimulationModel.Forecasts[0];
        Assert.Equal(DeleteDefinitionResult.Cancelled, ModelDefinitionDeletion.Delete(definition, (_, _) => false));
        Assert.Equal(DeleteDefinitionResult.Deleted, ModelDefinitionDeletion.Delete(definition, (_, _) => true));
        Assert.Single(SimulationModel.Forecasts); Assert.Equal(65535, Book.Worksheets[1].Range["B3"].Interior.Color);
    }
    [Fact]
    public void ClearModelStillClearsAllAndRestoresAllFills()
    {
        WorkbookPersistence.ClearSavedModel();
        Assert.Empty(SimulationModel.Assumptions); Assert.Empty(SimulationModel.Forecasts);
        Assert.Single(Book.Worksheets.Items);
        foreach (var cell in Book.Worksheets[1].Cells.Items.Values) Assert.Equal(65535, cell.Interior.Color);
    }
}
public sealed class SelectionDouble(int cells, int areas)
{
    public PersistenceCount Cells => new(cells);
    public PersistenceCount Areas => new(areas);
}
