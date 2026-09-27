using ExcelDna.Integration;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public sealed class SimulationSettingsTests : IDisposable
{
    private readonly PersistenceApplication app = new() { ActiveWorkbook = new() };
    private PersistenceWorkbook Book => app.ActiveWorkbook!;
    public SimulationSettingsTests()
    {
        SimulationModel.Clear(); ExcelDnaUtil.Application = app;
        Book.Worksheets.Add().Name = "Model";
        Book.Worksheets[1].Range["B2"].Value2 = 2d;
        Book.Worksheets[1].Range["B3"].Value2 = 4d;
        SimulationModel.Assumptions.Add(new() { Name = "Input", SheetName = "Model", CellAddress = "B2", Distribution = DistributionType.Uniform, Parameter1 = 1, Parameter2 = 3 });
        SimulationModel.Forecasts.Add(new() { Name = "Output", SheetName = "Model", CellAddress = "B3" });
        WorkbookPersistence.SaveModel();
        app.OnCalculate = () => Book.Worksheets[1].Range["B3"].Value2 = Convert.ToDouble(Book.Worksheets[1].Range["B2"].Value2) * 2;
    }
    public void Dispose() { SimulationModel.Clear(); ExcelDnaUtil.Application = null!; }
    [Theory]
    [InlineData("1000")] [InlineData("5000")] [InlineData("10000")] [InlineData("50000")] [InlineData("123")] [InlineData("1")] [InlineData("1000000")]
    public void PresetsAndCustomAccepted(string trials)
    {
        Assert.True(SimulationSettings.TryParse(trials, SimulationSeedMode.Automatic, "", out var settings, out _));
        Assert.Equal(int.Parse(trials), settings.TrialCount);
    }
    [Theory]
    [InlineData("")] [InlineData("0")] [InlineData("-1")] [InlineData("1.5")] [InlineData("bad")] [InlineData("1000001")]
    public void InvalidTrialsRejected(string trials) => Assert.False(SimulationSettings.TryParse(trials, SimulationSeedMode.Automatic, "", out _, out _));
    [Theory]
    [InlineData("")] [InlineData("1.5")] [InlineData("bad")] [InlineData("2147483648")]
    public void InvalidFixedSeedRejected(string seed) => Assert.False(SimulationSettings.TryParse("10", SimulationSeedMode.Fixed, seed, out _, out _));
    [Theory]
    [InlineData("-2147483648")] [InlineData("2147483647")] [InlineData("0")] [InlineData("42")]
    public void ValidFixedSeedAccepted(string seed) => Assert.True(SimulationSettings.TryParse("10", SimulationSeedMode.Fixed, seed, out _, out _));
    [Fact]
    public void LegacyWorkbookUsesDefaults()
    {
        var validation = new ValidationResult();
        Assert.Equal(new SimulationSettings(), WorkbookPersistence.LoadSimulationSettings(validation));
        Assert.True(validation.IsValid);
    }
    [Fact]
    public void WorkbookSettingsSurviveEditDeleteClearAndReopen()
    {
        var settings = new SimulationSettings(123, SimulationSeedMode.Fixed, -42);
        WorkbookPersistence.SaveSimulationSettings(settings);
        SimulationModel.Assumptions[0].Parameter2 = 4; WorkbookPersistence.SaveModel();
        ModelDefinitionDeletion.ClearSelectedCell(Book.Worksheets[1].Range["B3"], (_, _) => true);
        Assert.Equal(settings, WorkbookPersistence.LoadSimulationSettings(new()));
        WorkbookPersistence.ClearSavedModel();
        app.ActiveWorkbook = Book.Reopen();
        Assert.Equal(settings, WorkbookPersistence.LoadSimulationSettings(new()));
        WorkbookPersistence.LoadModel(); Assert.Empty(SimulationModel.Assumptions); Assert.Empty(SimulationModel.Forecasts);
        app.ActiveWorkbook = new(); Book.Worksheets.Add();
        Assert.Equal(new SimulationSettings(), WorkbookPersistence.LoadSimulationSettings(new()));
    }
    [Theory]
    [InlineData(17, "WrongVersion")] [InlineData(18, "0")] [InlineData(18, "1.5")] [InlineData(19, "Unknown")] [InlineData(20, "bad")]
    public void MalformedSettingsPreventCalculation(int column, string value)
    {
        WorkbookPersistence.SaveSimulationSettings(new(10, SimulationSeedMode.Fixed, 42));
        Book.Worksheets["__MonteCarloConfig"].Cells[1, column].Value2 = value;
        var result = SimulationService.TryRun();
        Assert.False(result.Succeeded); Assert.Contains("Simulation Settings", result.UserMessage);
        Assert.Equal(0, app.CalculationCount);
    }
    [Theory]
    [InlineData(DistributionType.Uniform, 1, 3, 0, 0)]
    [InlineData(DistributionType.Normal, 2, 1, 0, 0)]
    [InlineData(DistributionType.Lognormal, 1, .2, 0, 0)]
    [InlineData(DistributionType.Triangular, 1, 2, 4, 0)]
    [InlineData(DistributionType.Pert, 1, 2, 4, 0)]
    [InlineData(DistributionType.Beta, 1, 4, .5, 2)]
    public void FixedSeedControlsEntireSamplingPath(DistributionType distribution, double p1, double p2, double p3, double p4)
    {
        var a = SimulationModel.Assumptions[0]; a.Distribution = distribution;
        a.Parameter1 = p1; a.Parameter2 = p2; a.Parameter3 = p3; a.Parameter4 = p4;
        WorkbookPersistence.SaveModel(); WorkbookPersistence.SaveSimulationSettings(new(40, SimulationSeedMode.Fixed, 42));
        var first = SimulationService.TryRun(); var second = SimulationService.TryRun();
        Assert.True(first.Succeeded, first.UserMessage); Assert.True(second.Succeeded, second.UserMessage);
        Assert.Equal(first.Result!.ForecastResults[0].Values, second.Result!.ForecastResults[0].Values);
        Assert.Equal(40, first.Result.Trials); Assert.Equal(40, first.Result.ForecastResults[0].Values.Length);
        Assert.Equal(42, first.Result.ActualSeedUsed); Assert.Equal(SimulationSeedMode.Fixed, first.Result.Settings!.SeedMode);
        WorkbookPersistence.SaveSimulationSettings(new(40, SimulationSeedMode.Fixed, 43));
        var third = SimulationService.TryRun(); Assert.True(third.Succeeded, third.UserMessage);
        Assert.False(first.Result.ForecastResults[0].Values.SequenceEqual(third.Result!.ForecastResults[0].Values));
    }
    [Fact]
    public void AutomaticSeedCanBeReplayedWithoutChangingConfiguredMode()
    {
        WorkbookPersistence.SaveSimulationSettings(new(25));
        var run = SimulationService.TryRun(); Assert.True(run.Succeeded, run.UserMessage);
        Assert.NotNull(run.Result!.ActualSeedUsed); Assert.Equal(SimulationSeedMode.Automatic, run.Result.Settings!.SeedMode);
        Assert.Equal(new SimulationSettings(25), WorkbookPersistence.LoadSimulationSettings(new()));
        Assert.Equal(26, app.CalculationCount); // 25 trials plus restoration recalculation.
        WorkbookPersistence.SaveSimulationSettings(new(25, SimulationSeedMode.Fixed, run.Result.ActualSeedUsed));
        var replay = SimulationService.TryRun();
        Assert.Equal(run.Result.ForecastResults[0].Values, replay.Result!.ForecastResults[0].Values);
    }
    [Theory]
    [InlineData(0, SimulationSeedMode.Automatic)] [InlineData(1000001, SimulationSeedMode.Automatic)] [InlineData(10, SimulationSeedMode.Fixed)]
    public void EngineIndependentlyRejectsInvalidSettings(int count, SimulationSeedMode mode)
    {
        var result = SimulationExecution.Run(new SimulationSettings(count, mode), SimulationModel.Assumptions, SimulationModel.Forecasts, new ExcelSimulationWorkbook(app, Book));
        Assert.False(result.Succeeded); Assert.Equal(0, app.CalculationCount);
    }
}
