using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel;
using System.Globalization;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class ForecastEditTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    private readonly ForecastDefinition forecast = new() { Name = "Total", SheetName = "Model", CellAddress = "B3", TargetSettings = new(50, TargetDirection.AtOrBelow, 80) };
    public ForecastEditTests()
    {
        SimulationModel.Clear(); ExcelDnaUtil.Application = app;
        var sheet = Book.Worksheets.Add(); sheet.Name = "Model";
        sheet.Range["B2"].Value2 = 3d;
        sheet.Range["B3"].Value2 = 6d; sheet.Range["B3"].Formula = "=B2*2";
        sheet.Range["B3"].Interior.Pattern = 1;
        sheet.Range["B3"].Interior.ColorIndex = 6;
        sheet.Range["B3"].Interior.Color = 65535;
        SimulationModel.Assumptions.Add(new() { Name = "Input", SheetName = "Model", CellAddress = "B2", Distribution = DistributionType.Uniform, Parameter1 = 2, Parameter2 = 4 });
        SimulationModel.Forecasts.Add(forecast);
        WorkbookPersistence.SaveModel();
    }
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }

    [Fact]
    public void DraftLoadsCompleteSettingsAndCancelDoesNotMutate()
    {
        var original = forecast.TargetSettings;
        var draft = new ForecastEditDraft(forecast);
        Assert.Equal("Total", draft.Name); Assert.Equal("50", draft.TargetText);
        Assert.Equal(ForecastTargetMode.Manual, draft.Mode);
        Assert.Equal(TargetDirection.AtOrBelow, draft.Direction); Assert.Equal(80, draft.RequestedConfidence);
        draft.Name = "Cancelled"; draft.TargetText = "100"; draft.Direction = TargetDirection.AtOrAbove;
        Assert.Equal("Total", forecast.Name); Assert.Same(original, forecast.TargetSettings);
        Assert.Equal(50d, Book.Worksheets["__MonteCarloConfig"].Cells[3, 12].Value2);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void NameOnlyEditPreservesDefaultExplicitNoneAndConfidence(int mode)
    {
        forecast.TargetSettings = mode == 0 ? null : mode == 1 ? new(null, TargetDirection.AtOrAbove, null) : new(50, TargetDirection.AtOrBelow, 80);
        var original = forecast.TargetSettings;
        var draft = new ForecastEditDraft(forecast) { Name = "Renamed" };
        Assert.Equal(mode == 0 ? ForecastTargetMode.DefaultP80 : mode == 1 ? ForecastTargetMode.None : ForecastTargetMode.Manual, draft.Mode);
        Assert.True(draft.TryApply(forecast, out _));
        Assert.Equal(original, forecast.TargetSettings); Assert.Equal("Renamed", forecast.Name);
    }

    [Theory]
    [InlineData("")]
    [InlineData("bad")]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("1e999")]
    public void InvalidTargetDoesNotPartiallyApply(string text)
    {
        var original = forecast.TargetSettings;
        var draft = new ForecastEditDraft(forecast) { Name = "Invalid edit", TargetText = text };
        Assert.False(draft.TryApply(forecast, out var error)); Assert.NotNull(error);
        Assert.Equal("Total", forecast.Name); Assert.Same(original, forecast.TargetSettings);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ChangingTargetOrDirectionClearsStaleConfidence(bool changeTarget)
    {
        var draft = new ForecastEditDraft(forecast);
        if (changeTarget) draft.TargetText = "60";
        else draft.Direction = TargetDirection.AtOrAbove;
        Assert.True(draft.TryApply(forecast, out _));
        Assert.Null(forecast.TargetSettings!.RequestedConfidence);
        Assert.Equal(changeTarget ? 60 : 50, forecast.TargetSettings.Target);
        Assert.Equal(changeTarget ? TargetDirection.AtOrBelow : TargetDirection.AtOrAbove, forecast.TargetSettings.Direction);
    }

    [Fact]
    public void ExplicitClearAndReturnToDefaultAreDistinct()
    {
        var draft = new ForecastEditDraft(forecast) { Mode = ForecastTargetMode.None };
        Assert.True(draft.TryApply(forecast, out _));
        Assert.NotNull(forecast.TargetSettings); Assert.Null(forecast.TargetSettings.Target);
        draft.Mode = ForecastTargetMode.DefaultP80;
        Assert.True(draft.TryApply(forecast, out _)); Assert.Null(forecast.TargetSettings);
    }

    [Fact]
    public void EditRoundTripPreservesIdentityHighlightOtherMetadataAndSimulationThenDeleteRestoresFill()
    {
        string link = forecast.CellLink;
        var config = Book.Worksheets["__MonteCarloConfig"];
        object?[] assumptionRow = Enumerable.Range(1, 15).Select(c => config.Cells[2, c].Value2).ToArray();
        object? fill = config.Cells[3, 15].Value2;
        var draft = new ForecastEditDraft(forecast) { Name = "New total", TargetText = "7", Direction = TargetDirection.AtOrAbove };
        Assert.True(draft.TryApply(forecast, out _));
        WorkbookPersistence.SaveModel();
        Assert.Single(SimulationModel.Forecasts); Assert.Equal(link, forecast.CellLink);
        Assert.Equal(fill, config.Cells[3, 15].Value2);
        Assert.Equal(assumptionRow, Enumerable.Range(1, 15).Select(c => config.Cells[2, c].Value2).ToArray());
        app.ActiveWorkbook = Book.Reopen();
        Assert.True(WorkbookPersistence.LoadModelForSimulation().IsValid);
        var restored = Assert.Single(SimulationModel.Forecasts);
        Assert.Equal("New total", restored.Name); Assert.Equal(link, restored.CellLink);
        Assert.Equal(new ForecastTargetSettings(7, TargetDirection.AtOrAbove, null), restored.TargetSettings);
        Assert.Equal(ModelCellHighlight.ForecastColor, Book.Worksheets[1].Range["B3"].Interior.Color);
        app.OnCalculate = () => Book.Worksheets[1].Range["B3"].Value2 = 2 * Convert.ToDouble(Book.Worksheets[1].Range["B2"].Value2);
        var result = SimulationService.TryRun(10);
        Assert.True(result.Succeeded, result.UserMessage);
        Assert.Equal(restored.TargetSettings, result.Result!.ForecastResults[0].Forecast.TargetSettings);
        SimulationModel.Forecasts.Clear(); WorkbookPersistence.SaveModel();
        Assert.Equal(65535, Book.Worksheets[1].Range["B3"].Interior.Color);
        Assert.Equal("=B2*2", Book.Worksheets[1].Range["B3"].Formula);
        app.ActiveWorkbook = Book.Reopen(); WorkbookPersistence.LoadModel();
        Assert.Empty(SimulationModel.Forecasts);
        Assert.Null(Book.Worksheets["__MonteCarloConfig"].Cells[3, 12].Value2);
    }

    [Fact]
    public void TargetUsesCurrentCultureWithoutLosingPrecision()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            forecast.TargetSettings = new(12.3456789012345, TargetDirection.AtOrBelow, null);
            var draft = new ForecastEditDraft(forecast);
            Assert.Contains(",", draft.TargetText);
            Assert.True(draft.TryApply(forecast, out _));
            Assert.Equal(12.3456789012345, forecast.TargetSettings!.Target);
        }
        finally { CultureInfo.CurrentCulture = previous; }
    }
}
