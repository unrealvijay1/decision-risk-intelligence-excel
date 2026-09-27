using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel;
using System.Reflection;
using System.Text.Json;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class ForecastRecreationTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    private PersistenceSheet Config => Book.Worksheets["__MonteCarloConfig"];
    private PersistenceCell Cell => Book.Worksheets[1].Range["B3"];
    public ForecastRecreationTests()
    {
        SimulationModel.Clear(); ExcelDnaUtil.Application = app;
        Book.Worksheets.Add().Name = "Model";
        Cell.Value2 = 42d; Cell.Formula = "=6*7"; Cell.NumberFormat = "0.00";
        Cell.Interior.Pattern = 1; Cell.Interior.ColorIndex = 6; Cell.Interior.Color = 65535;
        SimulationModel.Assumptions.Add(new() { Name = "Input", SheetName = "Model", CellAddress = "B2", Distribution = DistributionType.Uniform, Parameter1 = 1, Parameter2 = 2 });
        SimulationModel.Forecasts.Add(new() { Name = "Other", SheetName = "Model", CellAddress = "B4", TargetSettings = new(99, TargetDirection.AtOrAbove, 95) });
    }
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    private static ForecastTargetSettings? Settings(int mode) => mode switch {
        0 => null, 1 => new(50, TargetDirection.AtOrBelow, 80), _ => new(null, null, null)
    };
    [Theory]
    [InlineData(0, 0)] [InlineData(0, 1)] [InlineData(0, 2)]
    [InlineData(1, 0)] [InlineData(1, 1)] [InlineData(1, 2)]
    [InlineData(2, 0)] [InlineData(2, 1)] [InlineData(2, 2)]
    public void DefineClearRedefineRoundTripsWithoutLeakingSettings(int oldMode, int newMode)
    {
        SimulationModel.Forecasts.Add(new() { Name = "Old", SheetName = "Model", CellAddress = "B3", TargetSettings = Settings(oldMode) });
        WorkbookPersistence.SaveModel();
        var other = JsonSerializer.Serialize(SimulationModel.Forecasts[0]);
        string originalFill = (string)Config.Cells[4, 15].Value2!;
        Assert.Equal(DeleteDefinitionResult.Deleted, ModelDefinitionDeletion.ClearSelectedCell(Cell, (_, _) => true));
        Assert.Null(SimulationModel.FindForecast("Model", "B3"));
        for (int c = 1; c <= 15; c++) Assert.Null(Config.Cells[4, c].Value2);
        Assert.Equal(65535, Cell.Interior.Color);
        var recreated = new ForecastDefinition { Name = "New", SheetName = "Model", CellAddress = "B3", TargetSettings = Settings(newMode) };
        SimulationModel.Forecasts.Add(recreated);
        WorkbookPersistence.SaveModel();
        Assert.Equal(originalFill, Config.Cells[4, 15].Value2);
        Assert.Equal(ModelCellHighlight.ForecastColor, Cell.Interior.Color);
        app.ActiveWorkbook = Book.Reopen();
        Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
        Assert.Equal(Settings(newMode), SimulationModel.FindForecast("Model", "B3")!.TargetSettings);
        Assert.Equal(other, JsonSerializer.Serialize(SimulationModel.Forecasts[0]));
        Assert.Equal(42d, Cell.Value2); Assert.Equal("=6*7", Cell.Formula); Assert.Equal("0.00", Cell.NumberFormat);
        Assert.Equal(DeleteDefinitionResult.Deleted, ModelDefinitionDeletion.ClearSelectedCell(Cell, (_, _) => true));
        Assert.Equal(65535, Cell.Interior.Color);
    }
    public static IEnumerable<object?[]> ValidStates()
    {
        yield return new object?[] { null };
        foreach (double? target in new double?[] { null, 12.5 })
        foreach (TargetDirection? direction in new TargetDirection?[] { null, TargetDirection.AtOrBelow, TargetDirection.AtOrAbove })
        foreach (double? confidence in target.HasValue ? new double?[] { null, 0, 80, 100 } : new double?[] { null })
            yield return new object?[] { new ForecastTargetSettings(target, direction, confidence) };
    }
    [Theory]
    [MemberData(nameof(ValidStates))]
    public void WriteTargetSettingsRepresentsEveryNullableStateAndClearsPreviousValues(ForecastTargetSettings? settings)
    {
        WorkbookPersistence.SaveModel();
        for (int c = 11; c <= 14; c++) Config.Cells[3, c].Value2 = "stale";
        typeof(WorkbookPersistence).GetMethod("WriteTargetSettings", BindingFlags.NonPublic | BindingFlags.Static)!
            .Invoke(null, new object?[] { Config, 3, settings });
        Assert.Equal(settings == null ? null : "Yes", Config.Cells[3, 11].Value2);
        Assert.Equal((object?)settings?.Target, Config.Cells[3, 12].Value2);
        Assert.Equal(settings?.Direction?.ToString(), Config.Cells[3, 13].Value2);
        Assert.Equal((object?)settings?.RequestedConfidence, Config.Cells[3, 14].Value2);
        Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
        Assert.Equal(settings, SimulationModel.Forecasts[0].TargetSettings);
    }
    [Fact]
    public void DynamicExcelSettersNeverReceiveNullableNumericArguments()
    {
        // COM's dynamic argument marshaler may unwrap Nullable<T>. Box optional numbers
        // before that boundary, rather than relying on the managed fake's object setter.
        var nullableArguments = typeof(WorkbookPersistence).GetNestedTypes(BindingFlags.NonPublic)
            .SelectMany(t => t.GetFields(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
            .Where(f => f.FieldType.IsGenericType && f.FieldType.GetGenericTypeDefinition() == typeof(System.Runtime.CompilerServices.CallSite<>))
            .SelectMany(f => f.FieldType.GetGenericArguments()[0].GetMethod("Invoke")!.GetParameters())
            .Where(p => Nullable.GetUnderlyingType(p.ParameterType) == typeof(double));
        Assert.Empty(nullableArguments);
    }
    [Fact]
    public void AssumptionCanAlsoBeClearedAndRecreated()
    {
        WorkbookPersistence.SaveModel();
        var cell = Book.Worksheets[1].Range["B2"];
        Assert.Equal(DeleteDefinitionResult.Deleted, ModelDefinitionDeletion.ClearSelectedCell(cell, (_, _) => true));
        SimulationModel.Assumptions.Add(new() { Name = "New input", SheetName = "Model", CellAddress = "B2", Distribution = DistributionType.Uniform, Parameter1 = 10, Parameter2 = 20 });
        WorkbookPersistence.SaveModel(); app.ActiveWorkbook = Book.Reopen(); WorkbookPersistence.LoadModel();
        Assert.Equal(10, Assert.Single(SimulationModel.Assumptions).Parameter1);
    }
}
