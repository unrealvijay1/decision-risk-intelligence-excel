using ExcelDna.Integration;
using MonteCarlo.Excel;
using System.Text.Json;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class ModelCellHighlightTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    private PersistenceCell Cell => Book.Worksheets[1].Range["B2"];
    private PersistenceSheet Config => Book.Worksheets["__MonteCarloConfig"];
    public ModelCellHighlightTests()
    {
        SimulationModel.Clear();
        ExcelDnaUtil.Application = app;
        Book.Worksheets.Add().Name = "Model";
    }
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    private void Define(bool forecast)
    {
        if (forecast) SimulationModel.Forecasts.Add(new() { SheetName = "Model", CellAddress = "B2" });
        else SimulationModel.Assumptions.Add(new() { SheetName = "Model", CellAddress = "B2", Distribution = DistributionType.Uniform, Parameter1 = 1, Parameter2 = 2 });
    }
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void CreateEditReopenDeletePreservesOriginalFill(bool forecast, bool filled)
    {
        Cell.Interior.Pattern = filled ? 9 : -4142;
        Cell.Interior.ColorIndex = filled ? 6 : -4142;
        Cell.Interior.Color = 65535;
        Cell.Interior.ThemeColor = filled ? 4 : 0;
        Cell.Interior.TintAndShade = .25;
        Cell.Interior.PatternColor = 123;
        Cell.Interior.PatternColorIndex = 3;
        Cell.Interior.PatternThemeColor = 5;
        Cell.Interior.PatternTintAndShade = -.2;
        Cell.Value2 = 42d; Cell.Formula = "=6*7"; Cell.NumberFormat = "0.00";
        var other = Book.Worksheets[1].Range["C3"];
        string unrelated = JsonSerializer.Serialize(other.Interior);
        Define(forecast);
        WorkbookPersistence.SaveModel();
        int color = forecast ? ModelCellHighlight.ForecastColor : ModelCellHighlight.AssumptionColor;
        Assert.Equal(color, Cell.Interior.Color);
        string snapshot = (string)Config.Cells[2, 15].Value2!;
        for (int i = 0; i < 3; i++)
        {
            if (forecast) SimulationModel.Forecasts[0].Name = "Edited " + i;
            else SimulationModel.Assumptions[0].Parameter2 = 5 + i;
            WorkbookPersistence.SaveModel();
            Assert.Equal(snapshot, Config.Cells[2, 15].Value2);
            Assert.Equal(color, Cell.Interior.Color);
        }
        Assert.Equal(unrelated, JsonSerializer.Serialize(other.Interior));
        Assert.Equal("Original font", Cell.Font);
        Assert.Equal("Original borders", Cell.Borders);
        Assert.Equal("Original alignment", Cell.Alignment);
        Assert.Equal("Original protection", Cell.Protection);
        app.ActiveWorkbook = Book.Reopen();
        Cell.Interior.Color = 0; // Reload reapplies the configured style.
        Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
        Assert.Equal(color, Cell.Interior.Color);
        SimulationModel.Clear();
        WorkbookPersistence.SaveModel();
        Assert.Equal(filled ? 9 : -4142, Cell.Interior.Pattern);
        Assert.Equal(filled ? 4 : 0, Cell.Interior.ThemeColor);
        Assert.Equal(.25, Cell.Interior.TintAndShade);
        Assert.Equal(5, Cell.Interior.PatternThemeColor);
        Assert.Equal(-.2, Cell.Interior.PatternTintAndShade);
        Assert.Equal(42d, Cell.Value2); Assert.Equal("=6*7", Cell.Formula); Assert.Equal("0.00", Cell.NumberFormat);
        Assert.Null(Config.Cells[2, 15].Value2);
    }
    [Theory]
    [InlineData(4000)]
    [InlineData(4001)]
    public void GradientFillIsRestored(int pattern)
    {
        Cell.Interior.Pattern = pattern;
        Cell.Interior.Gradient.Degree = 45;
        Cell.Interior.Gradient.RectangleLeft = .2;
        Cell.Interior.Gradient.RectangleBottom = .7;
        Cell.Interior.Gradient.ColorStops.Add(0).Color = 65535;
        var stop = Cell.Interior.Gradient.ColorStops.Add(1);
        stop.ThemeColor = 4; stop.TintAndShade = .3;
        Define(false); WorkbookPersistence.SaveModel();
        Cell.Interior.Gradient = new();
        SimulationModel.Clear(); WorkbookPersistence.SaveModel();
        Assert.Equal(pattern, Cell.Interior.Pattern);
        Assert.Equal(pattern == 4000 ? 45 : 0, Cell.Interior.Gradient.Degree);
        Assert.Equal(pattern == 4001 ? .2 : 0, Cell.Interior.Gradient.RectangleLeft);
        Assert.Equal(65535, Cell.Interior.Gradient.ColorStops.Item(1).Color);
        Assert.Equal(4, Cell.Interior.Gradient.ColorStops.Item(2).ThemeColor);
        Assert.Equal(.3, Cell.Interior.Gradient.ColorStops.Item(2).TintAndShade);
    }
    [Fact]
    public void StylesAreDistinct() => Assert.NotEqual(ModelCellHighlight.AssumptionColor, ModelCellHighlight.ForecastColor);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyMetadataIsCapturedOnceAndClearRestores(bool forecast)
    {
        Define(forecast); WorkbookPersistence.SaveModel();
        Config.Cells[1, 15].Value2 = null; Config.Cells[2, 15].Value2 = null;
        Cell.Interior.Pattern = 1; Cell.Interior.ColorIndex = 6; Cell.Interior.Color = 65535;
        Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
        string snapshot = (string)Config.Cells[2, 15].Value2!;
        WorkbookPersistence.LoadModel();
        Assert.Equal(snapshot, Config.Cells[2, 15].Value2);
        WorkbookPersistence.ClearSavedModel();
        Assert.Equal(65535, Cell.Interior.Color);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TrackedMoveAndRenameRestoreCurrentCell(bool forecast)
    {
        Cell.Interior.Pattern = 1; Cell.Interior.ColorIndex = 6; Cell.Interior.Color = 65535;
        var moved = Cell;
        Define(forecast); WorkbookPersistence.SaveModel();
        moved.Location = "D8"; Book.Worksheets[1].Name = "Renamed";
        moved.Interior.Color = 0;
        Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
        Assert.Equal(forecast ? ModelCellHighlight.ForecastColor : ModelCellHighlight.AssumptionColor, moved.Interior.Color);
        SimulationModel.Clear(); WorkbookPersistence.SaveModel();
        Assert.Equal(65535, moved.Interior.Color);
    }
    [Fact]
    public void BrokenTrackedLinkDoesNotTouchReplacementCell()
    {
        Define(false); WorkbookPersistence.SaveModel();
        var removed = Cell; removed.Deleted = true;
        var replacement = new PersistenceCell(Book.Worksheets[1], 2, 2);
        Book.Worksheets[1].Cells.Items[(2, 2)] = replacement;
        WorkbookPersistence.LoadModel();
        SimulationModel.Clear(); WorkbookPersistence.SaveModel();
        Assert.Equal(-4142, replacement.Interior.Pattern);
    }
}
