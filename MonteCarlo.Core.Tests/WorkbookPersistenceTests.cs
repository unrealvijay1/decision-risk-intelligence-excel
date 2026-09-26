using ExcelDna.Integration;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

[CollectionDefinition("Workbook persistence", DisableParallelization = true)]
public sealed class WorkbookPersistenceCollection;

[Collection("Workbook persistence")]
public sealed class WorkbookPersistenceTests : IDisposable
{
    private readonly PersistenceApplication app = new();
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    private PersistenceSheet Config => Book.Worksheets["__MonteCarloConfig"];

    public WorkbookPersistenceTests()
    {
        SimulationModel.Clear();
        ExcelDnaUtil.Application = app;
        app.ActiveWorkbook = new();
        Book.Worksheets.Add().Name = "Inputs' Plan";
        Book.Worksheets[1].Range["B2"].Value2 = 10.0;
        Book.Worksheets[1].Range["B3"].Value2 = 20.0;
        SimulationModel.Assumptions.Add(new()
        {
            Name = "Cost", SheetName = "Inputs' Plan", CellAddress = "B2",
            Distribution = DistributionType.Uniform, Parameter1 = 1, Parameter2 = 3
        });
        SimulationModel.Forecasts.Add(new() { Name = "Total", SheetName = "Inputs' Plan", CellAddress = "B3" });
    }
    public void Dispose()
    {
        SimulationModel.Clear(); WorkbookPersistence.ReportRestoreProblem = null;
        ExcelDnaUtil.Application = null!;
    }
    private ValidationResult Reopen()
    {
        app.ActiveWorkbook = Book.Reopen();
        SimulationModel.Clear();
        return WorkbookPersistence.LoadModelForSimulation();
    }

    [Theory]
    [InlineData(DistributionType.Normal, 2, 1, 9, 7)]
    [InlineData(DistributionType.Lognormal, 2, 1, 9, 7)]
    [InlineData(DistributionType.Uniform, 1, 5, 9, 7)]
    [InlineData(DistributionType.Triangular, 1, 3, 5, 7)]
    [InlineData(DistributionType.Pert, 1, 3, 5, 7)]
    [InlineData(DistributionType.Beta, 1, 5, 2, 4)]
    public void AllDistributionsAndEveryParameterRoundTrip(DistributionType kind, double p1, double p2, double p3, double p4)
    {
        var assumption = SimulationModel.Assumptions[0];
        assumption.Distribution = kind;
        assumption.Parameter1 = p1; assumption.Parameter2 = p2; assumption.Parameter3 = p3; assumption.Parameter4 = p4;
        WorkbookPersistence.SaveModel();
        Assert.True(Reopen().IsValid);
        var restored = Assert.Single(SimulationModel.Assumptions);
        Assert.Equal(kind, restored.Distribution);
        Assert.Equal(new[] { p1, p2, p3, p4 }, new[] { restored.Parameter1, restored.Parameter2, restored.Parameter3, restored.Parameter4 });
        Assert.Equal("Cost", restored.Name);
        Assert.Equal("Inputs' Plan", restored.SheetName);
        Assert.Equal("B2", restored.CellAddress);
        Assert.StartsWith("_MC_Cell_", restored.CellLink);
        Assert.Equal("Total", Assert.Single(SimulationModel.Forecasts).Name);
        Assert.Equal(2, Config.Visible);
        Assert.All(Book.Names.Items.Values, name => Assert.False(name.Visible));
        Assert.Equal(10.0, Book.Worksheets[1].Range["B2"].Value2);
    }

    [Theory]
    [InlineData(null, null, null)]
    [InlineData(10608.923456789, TargetDirection.AtOrBelow, 70.0)]
    [InlineData(-25.123456789, TargetDirection.AtOrAbove, null)]
    [InlineData(10.0, null, null)]
    [InlineData(10.0, TargetDirection.AtOrBelow, 0.0)]
    [InlineData(10.0, TargetDirection.AtOrAbove, 100.0)]
    public void TargetSettingsRoundTrip(double? target, TargetDirection? direction, double? requested)
    {
        var settings = new ForecastTargetSettings(target, direction, requested);
        SimulationModel.Forecasts[0].TargetSettings = settings;
        WorkbookPersistence.SaveModel();
        Assert.True(Reopen().IsValid);
        Assert.Equal(settings, SimulationModel.Forecasts[0].TargetSettings);
    }

    [Fact]
    public void TargetOnlyUpdatePreservesOtherRowsAndDefaultDistinction()
    {
        SimulationModel.Forecasts.Add(new() { Name = "Second", SheetName = "Inputs' Plan", CellAddress = "B4" });
        WorkbookPersistence.SaveModel();
        Config.Cells[2, 5].Value2 = "damaged";
        var settings = new ForecastTargetSettings(null, TargetDirection.AtOrAbove, null);
        WorkbookPersistence.SaveTargetSettings(SimulationModel.Forecasts[0], settings);
        Assert.Equal("damaged", Config.Cells[2, 5].Value2);
        Assert.False(Reopen().IsValid);
        Assert.Equal(settings, SimulationModel.Forecasts[0].TargetSettings);
        Assert.Null(SimulationModel.Forecasts[1].TargetSettings);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OlderLayoutsWithoutOptionalFieldsStillLoad(bool fourParameters)
    {
        var sheet = Book.Worksheets.Add(); sheet.Name = "__MonteCarloConfig";
        sheet.Cells[1, 8].Value2 = fourParameters ? "Parameter4" : "Name";
        if (fourParameters) sheet.Cells[1, 9].Value2 = "Name";
        sheet.Cells[2, 1].Value2 = "Assumption"; sheet.Cells[2, 2].Value2 = "Inputs' Plan";
        sheet.Cells[2, 3].Value2 = "B2"; sheet.Cells[2, 4].Value2 = "Pert";
        sheet.Cells[2, 5].Value2 = 1.0; sheet.Cells[2, 6].Value2 = 2.0; sheet.Cells[2, 7].Value2 = 3.0;
        sheet.Cells[2, fourParameters ? 9 : 8].Value2 = "Legacy";
        sheet.Cells[3, 1].Value2 = "Forecast"; sheet.Cells[3, 2].Value2 = "Inputs' Plan";
        sheet.Cells[3, 3].Value2 = "B3"; sheet.Cells[3, fourParameters ? 9 : 8].Value2 = "Output";
        Assert.True(Reopen().IsValid);
        Assert.Equal("Legacy", SimulationModel.Assumptions[0].Name);
        Assert.Equal(0, SimulationModel.Assumptions[0].Parameter4);
        Assert.Null(SimulationModel.Forecasts[0].TargetSettings);
        Assert.Equal("", SimulationModel.Forecasts[0].CellLink);
        WorkbookPersistence.SaveTargetSettings(SimulationModel.Forecasts[0], new(12, TargetDirection.AtOrBelow, null));
        Assert.True(Reopen().IsValid);
        Assert.Equal(12, SimulationModel.Forecasts[0].TargetSettings!.Target);
        WorkbookPersistence.SaveModel();
        Assert.True(Reopen().IsValid);
        Assert.NotEmpty(SimulationModel.Forecasts[0].CellLink);
    }

    [Fact]
    public void TrackedLinksFollowRenamedSheetAndMovedCell()
    {
        WorkbookPersistence.SaveModel();
        Book.Worksheets[1].Name = "Renamed";
        Book.Worksheets[1].Range["B2"].Location = "C8";
        Assert.True(Reopen().IsValid);
        Assert.Equal("Renamed", SimulationModel.Assumptions[0].SheetName);
        Assert.Equal("C8", SimulationModel.Assumptions[0].CellAddress);
        Assert.Equal("Renamed", SimulationModel.Forecasts[0].SheetName);
    }

    [Theory]
    [InlineData("sheet")]
    [InlineData("cell")]
    [InlineData("name")]
    [InlineData("untrusted")]
    public void BrokenTrackedLinksNeverFallBackToOldAddress(string damage)
    {
        WorkbookPersistence.SaveModel();
        if (damage == "sheet") Book.Worksheets[1].Delete();
        else if (damage == "cell") Book.Worksheets[1].Range["B2"].Deleted = true;
        else if (damage == "name") Book.Names.Items.Clear();
        else Config.Cells[2, 10].Value2 = "OtherUserName";
        var validation = WorkbookPersistence.LoadModelForSimulation();
        Assert.False(validation.IsValid);
        Assert.Contains("Saved assumption row 2", validation.UserMessage);
        Assert.Contains("redefine", validation.UserMessage);
        Assert.Equal("#REF!", SimulationModel.Assumptions[0].CellAddress);
        var outcome = SimulationService.TryRun(10);
        Assert.False(outcome.Succeeded);
        Assert.Equal(0, app.CalculationCount);
    }

    [Theory]
    [InlineData("missingParameter")]
    [InlineData("errorParameter")]
    [InlineData("distribution")]
    [InlineData("missingType")]
    [InlineData("unknownType")]
    [InlineData("missingSheet")]
    [InlineData("badRange")]
    public void CorruptMetadataIsReportedAtRestore(string damage)
    {
        WorkbookPersistence.SaveModel();
        switch (damage)
        {
            case "missingParameter": Config.Cells[2, 5].Value2 = null; break;
            case "errorParameter": Config.Cells[2, 5].Value2 = new ExcelCellError("#N/A"); break;
            case "distribution": Config.Cells[2, 4].Value2 = "Unknown"; break;
            case "missingType": Config.Cells[2, 1].Value2 = null; break;
            case "unknownType": Config.Cells[2, 1].Value2 = "Unknown"; break;
            case "missingSheet": Config.Cells[2, 10].Value2 = null; Config.Cells[2, 2].Value2 = "Missing"; break;
            case "badRange": Config.Cells[2, 10].Value2 = null; Config.Cells[2, 3].Value2 = "#REF!"; break;
        }
        string? message = null;
        WorkbookPersistence.ReportRestoreProblem = value => message = value;
        WorkbookPersistence.LoadModel();
        Assert.False(WorkbookPersistence.LastLoadValidation.IsValid);
        Assert.NotNull(message);
        Assert.Contains("row 2", message);
        Assert.DoesNotContain("COMException", message);
        Assert.Single(SimulationModel.Forecasts); // Damage does not truncate later valid rows.
    }

    [Theory]
    [InlineData(11, "Unknown")]
    [InlineData(12, "NaN")]
    [InlineData(13, "Unknown")]
    [InlineData(14, "101")]
    public void InvalidOptionalTargetSettingsAreReported(int column, string value)
    {
        SimulationModel.Forecasts[0].TargetSettings = new(10, TargetDirection.AtOrBelow, 70);
        WorkbookPersistence.SaveModel();
        Config.Cells[3, column].Value2 = value;
        Assert.False(Reopen().IsValid);
    }

    [Fact]
    public void SaveReopenThenSimulationUsesRestoredModelAndRestoresInputs()
    {
        SimulationModel.Assumptions.Add(new() { Name = "Other", SheetName = "Inputs' Plan", CellAddress = "C2", Distribution = DistributionType.Uniform, Parameter1 = 10, Parameter2 = 20 });
        Book.Worksheets[1].Range["C2"].Value2 = 15.0;
        SimulationModel.Forecasts[0].TargetSettings = new(30, TargetDirection.AtOrBelow, null);
        WorkbookPersistence.SaveModel();
        Assert.True(Reopen().IsValid);
        // Deliberately poison memory: TryRun must reload saved settings, not use this value.
        SimulationModel.Assumptions[0].Parameter1 = 9999;
        app.OnCalculate = () => Book.Worksheets[1].Range["B3"].Value2 =
            2 * Convert.ToDouble(Book.Worksheets[1].Range["B2"].Value2) + Convert.ToDouble(Book.Worksheets[1].Range["C2"].Value2);
        var outcome = SimulationService.TryRun(30);
        Assert.True(outcome.Succeeded, outcome.UserMessage);
        Assert.Equal(2, SimulationModel.Assumptions.Count);
        Assert.All(outcome.Result!.ForecastResults[0].Values, value => Assert.InRange(value, 12, 26));
        Assert.Equal(30, outcome.Result.ForecastResults[0].Forecast.TargetSettings!.Target);
        Assert.Equal(10.0, Book.Worksheets[1].Range["B2"].Value2);
        Assert.Equal(15.0, Book.Worksheets[1].Range["C2"].Value2);
        Assert.True(app.EnableEvents);
    }

    [Fact]
    public void LoadingUnconfiguredWorkbookClearsPreviousWorkbookModel()
    {
        WorkbookPersistence.SaveModel();
        app.ActiveWorkbook = new();
        WorkbookPersistence.LoadModel();
        Assert.Empty(SimulationModel.Assumptions);
        Assert.Empty(SimulationModel.Forecasts);
        Assert.True(WorkbookPersistence.LastLoadValidation.IsValid);
    }

    [Fact]
    public void MultipleForecastTargetsStayIndependentAcrossReloadAndSave()
    {
        SimulationModel.Forecasts[0].TargetSettings = new(20, TargetDirection.AtOrBelow, 80);
        SimulationModel.Forecasts.Add(new() { Name = "Revenue", SheetName = "Inputs' Plan", CellAddress = "B4", TargetSettings = new(100, TargetDirection.AtOrAbove, 50) });
        WorkbookPersistence.SaveModel();
        Assert.True(Reopen().IsValid);
        WorkbookPersistence.SaveTargetSettings(SimulationModel.Forecasts[0], new(25, null, null));
        WorkbookPersistence.SaveModel();
        Assert.True(Reopen().IsValid);
        Assert.Equal(new ForecastTargetSettings(25, null, null), SimulationModel.Forecasts[0].TargetSettings);
        Assert.Equal(new ForecastTargetSettings(100, TargetDirection.AtOrAbove, 50), SimulationModel.Forecasts[1].TargetSettings);
    }

    [Fact]
    public void TargetUpdateCannotWriteIntoAnotherWorkbookWithSameCellAddress()
    {
        WorkbookPersistence.SaveModel();
        var firstForecast = SimulationModel.Forecasts[0];
        var other = new PersistenceWorkbook { Name = "Other.xlsx" };
        other.Worksheets.Add().Name = "Inputs' Plan";
        app.ActiveWorkbook = other;
        SimulationModel.Clear();
        SimulationModel.Forecasts.Add(new() { SheetName = "Inputs' Plan", CellAddress = "B3" });
        WorkbookPersistence.SaveModel();
        Assert.Throws<InvalidOperationException>(() => WorkbookPersistence.SaveTargetSettings(firstForecast, new(42, null, null)));
        Assert.Null(Config.Cells[2, 12].Value2);
    }

    [Fact]
    public void FormulaErrorInOptionalTargetIsRejected()
    {
        SimulationModel.Forecasts[0].TargetSettings = new(20, null, null);
        WorkbookPersistence.SaveModel();
        Config.Cells[3, 12].Value2 = new ExcelCellError("#DIV/0!");
        Assert.False(Reopen().IsValid);
    }

    [Fact]
    public void TargetUpdateUsesSourceWorkbookAfterReportChangesActiveWorkbook()
    {
        WorkbookPersistence.SaveModel();
        var source = Book;
        var forecast = SimulationModel.Forecasts[0];
        app.ActiveWorkbook = new() { Name = "Report.xlsx" };
        WorkbookPersistence.SaveTargetSettings(forecast, new(42, TargetDirection.AtOrAbove, null), source);
        Assert.Empty(Book.Worksheets.Items);
        app.ActiveWorkbook = source;
        Assert.True(Reopen().IsValid);
        Assert.Equal(42, SimulationModel.Forecasts[0].TargetSettings!.Target);
    }

    [Fact]
    public void LegacyReferenceToDeletedWorksheetProducesFriendlyWarning()
    {
        WorkbookPersistence.SaveModel();
        Config.Cells[1, 10].Value2 = null;
        Book.Worksheets[1].Delete();
        var validation = WorkbookPersistence.LoadModelForSimulation();
        Assert.False(validation.IsValid);
        Assert.Contains("Inputs' Plan!B2", validation.UserMessage);
        Assert.Contains("no longer exists", validation.UserMessage);
    }
}
