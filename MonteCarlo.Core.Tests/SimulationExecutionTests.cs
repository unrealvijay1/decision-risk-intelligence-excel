using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

public class SimulationExecutionTests
{
    private static AssumptionDefinition Input(string address = "A1") => new()
        { Name = address, SheetName = "Sheet1", CellAddress = address, Distribution = DistributionType.Uniform, Parameter1 = 1, Parameter2 = 2 };
    private static ForecastDefinition Output() => new() { Name = "Cost", SheetName = "Sheet1", CellAddress = "B1" };

    [Fact]
    public void InvalidConfigurationNeverTouchesWorkbook()
    {
        var workbook = new FakeWorkbook();
        var assumption = Input();
        assumption.Parameter2 = 0;
        var outcome = SimulationExecution.Run(2, [assumption], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Equal(0, workbook.Resolutions);
        Assert.Equal(0, workbook.StateWrites);
        Assert.Equal(0, workbook.Input.Writes);
    }

    [Fact]
    public void InvalidSavedParametersPreventExecution()
    {
        var workbook = new FakeWorkbook();
        var saved = new ValidationResult();
        SimulationValidation.ValidateParameters(saved, "Saved row 2", DistributionType.Normal, [null, 1]);
        var outcome = SimulationExecution.Run(2, [Input()], [Output()], workbook, saved);
        Assert.False(outcome.Succeeded);
        Assert.Equal(0, workbook.Resolutions);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("A1:B2")]
    [InlineData("#REF!")]
    public void BadReferencesPreventMutation(string address)
    {
        var workbook = new FakeWorkbook();
        var forecast = Output();
        forecast.CellAddress = address;
        var outcome = SimulationExecution.Run(2, [Input()], [forecast], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Contains(address, outcome.UserMessage);
        Assert.Equal(0, workbook.Input.Writes);
        Assert.Equal(0, workbook.StateWrites);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("bad")]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void BadForecastValuePreventsMutation(object? value)
    {
        var workbook = new FakeWorkbook();
        workbook.Output.Value = value;
        var outcome = SimulationExecution.Run(2, [Input()], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Equal(0, workbook.Input.Writes);
    }

    [Fact]
    public void DuplicateInputsAndProtectedInputsAreRejected()
    {
        var workbook = new FakeWorkbook();
        workbook.Input.InputProblem = "The cell is protected.";
        var outcome = SimulationExecution.Run(2, [Input(), Input()], [Output()], workbook);
        Assert.Contains("protected", outcome.UserMessage);
        Assert.Contains("same input cell", outcome.UserMessage);
        Assert.Equal(0, workbook.Input.Writes);
    }

    [Fact]
    public void SuccessfulRunPreservesSampleCountAndOriginalFormulaAndSettings()
    {
        var workbook = new FakeWorkbook();
        var outcome = SimulationExecution.Run(3, [Input()], [Output()], workbook);
        Assert.True(outcome.Succeeded);
        Assert.Equal(new double[] { 42, 42, 42 }, outcome.Result!.ForecastResults[0].Values);
        Assert.Equal(3, workbook.Input.Writes);
        AssertRestored(workbook);
    }

    [Theory]
    [InlineData("calculate")]
    [InlineData("write")]
    [InlineData("events")]
    [InlineData("status")]
    public void RuntimeFailuresRestoreInputsAndState(string failure)
    {
        var workbook = new FakeWorkbook { Failure = failure };
        var outcome = SimulationExecution.Run(2, [Input()], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.NotNull(outcome.DiagnosticException);
        Assert.DoesNotContain("internal failure", outcome.UserMessage);
        AssertRestored(workbook);
    }

    [Fact]
    public void FormulaErrorMidRunHasTrialAndCellContextAndRestores()
    {
        var workbook = new FakeWorkbook { Failure = "formula" };
        var outcome = SimulationExecution.Run(2, [Input()], [Output()], workbook);
        Assert.Contains("#DIV/0!", outcome.UserMessage);
        Assert.Contains("trial 1", outcome.UserMessage);
        Assert.Contains("B1", outcome.UserMessage);
        AssertRestored(workbook);
    }

    [Fact]
    public void RestorationFailureDoesNotSkipOtherCellsOrSettings()
    {
        var workbook = new FakeWorkbook();
        workbook.Input.FailRestore = true;
        var outcome = SimulationExecution.Run(2, [Input(), Input("A2")], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Contains("Sheet1!A1", outcome.RestorationFailures);
        Assert.Equal(2, workbook.Input.Restores);
        Assert.Equal("=6*7", workbook.Second.Formula);
        Assert.True(workbook.EnableEvents);
        Assert.True(workbook.ScreenUpdating);
        Assert.Equal("Ready", workbook.StatusBar);
        Assert.Contains("before saving", outcome.UserMessage);
    }

    [Fact]
    public void SnapshotFailureOccursBeforeAnyMutation()
    {
        var workbook = new FakeWorkbook();
        workbook.Second.FailCapture = true;
        var outcome = SimulationExecution.Run(2, [Input(), Input("A2")], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Equal(0, workbook.Input.Writes);
        Assert.Equal(0, workbook.StateWrites);
    }

    [Fact]
    public void InputFormulaErrorPreventsAnyMutation()
    {
        var workbook = new FakeWorkbook();
        workbook.Input.Value = new ExcelCellError("#REF!");
        var outcome = SimulationExecution.Run(2, [Input()], [Output()], workbook);
        Assert.Contains("#REF!", outcome.UserMessage);
        Assert.Contains("A1", outcome.UserMessage);
        Assert.Equal(0, workbook.Input.Writes);
        Assert.Equal(0, workbook.StateWrites);
    }

    [Fact]
    public void FailureAfterFirstInputWriteRestoresBothInputs()
    {
        var workbook = new FakeWorkbook();
        workbook.Second.FailWrite = true;
        var outcome = SimulationExecution.Run(2, [Input(), Input("A2")], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Equal(1, workbook.Input.Writes);
        Assert.Equal(1, workbook.Second.Writes);
        AssertRestored(workbook);
        Assert.Equal("=6*7", workbook.Second.Formula);
    }

    [Fact]
    public void TransientRestorationFailureIsRetried()
    {
        var workbook = new FakeWorkbook();
        workbook.Input.FailFirstRestore = true;
        var outcome = SimulationExecution.Run(2, [Input()], [Output()], workbook);
        Assert.True(outcome.Succeeded);
        Assert.Equal(2, workbook.Input.Restores);
        AssertRestored(workbook);
    }

    [Fact]
    public void FailedRecalculationDuringCleanupStillRestoresSettings()
    {
        var workbook = new FakeWorkbook { Failure = "cleanup" };
        var outcome = SimulationExecution.Run(1, [Input()], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Contains("workbook recalculation", outcome.RestorationFailures);
        AssertRestored(workbook);
    }

    [Fact]
    public void PlainValueInputsRetainValuesRatherThanBecomingFormulas()
    {
        var workbook = new FakeWorkbook();
        workbook.Input.Formula = null;
        workbook.Input.Value = 123.456;
        var outcome = SimulationExecution.Run(2, [Input()], [Output()], workbook);
        Assert.True(outcome.Succeeded);
        Assert.Null(workbook.Input.Formula);
        Assert.Equal(123.456, workbook.Input.Value);
        Assert.False(workbook.Input.LastRestore!.IsFormula);
    }

    [Fact]
    public void OriginalDisabledSettingsAndBooleanStatusArePreserved()
    {
        var workbook = new FakeWorkbook { ScreenUpdating = false, EnableEvents = false, StatusBar = false };
        var outcome = SimulationExecution.Run(1, [Input()], [Output()], workbook);
        Assert.True(outcome.Succeeded);
        Assert.False(workbook.ScreenUpdating);
        Assert.False(workbook.EnableEvents);
        Assert.Equal(false, workbook.StatusBar);
    }

    [Fact]
    public void FailedSettingRestorationDoesNotSkipRemainingSettings()
    {
        var workbook = new FakeWorkbook { Failure = "restoreScreen" };
        var outcome = SimulationExecution.Run(1, [Input()], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Contains("Screen Updating", outcome.RestorationFailures);
        Assert.True(workbook.EnableEvents);
        Assert.Equal("Ready", workbook.StatusBar);
        Assert.Equal("=6*7", workbook.Input.Formula);
    }

    [Fact]
    public void NonFiniteGeneratedSampleIsStoppedAndRestored()
    {
        var workbook = new FakeWorkbook();
        var assumption = Input();
        assumption.Parameter1 = -double.MaxValue;
        assumption.Parameter2 = double.MaxValue;
        var outcome = SimulationExecution.Run(1, [assumption], [Output()], workbook);
        Assert.False(outcome.Succeeded);
        Assert.Contains("numeric range", outcome.UserMessage);
        Assert.Equal(0, workbook.Input.Writes);
        AssertRestored(workbook);
    }

    private static void AssertRestored(FakeWorkbook workbook)
    {
        Assert.Equal("=6*7", workbook.Input.Formula);
        Assert.True(workbook.Input.LastRestore!.UsesFormula2);
        Assert.True(workbook.ScreenUpdating);
        Assert.True(workbook.EnableEvents);
        Assert.Equal("Ready", workbook.StatusBar);
    }

    private sealed class FakeCell(string identity) : ISimulationCell
    {
        public string Identity => identity;
        public string? InputProblem { get; set; }
        public object? Value = 42d;
        public string? Formula = "=6*7";
        public bool FailRestore, FailCapture, FailWrite, FailFirstRestore;
        public int Writes, Restores;
        public SimulationCellContent? LastRestore;
        public object? ReadValue() => Value;
        public SimulationCellContent Capture() => FailCapture ? throw new Exception("internal failure")
            : Formula != null ? new(Formula, true, true) : new(Value);
        public void WriteSample(double value)
        {
            Writes++;
            Formula = null;
            Value = value;
            if (FailWrite) throw new Exception("internal failure");
        }
        public void Restore(SimulationCellContent content)
        {
            Restores++;
            if (FailRestore || FailFirstRestore && Restores == 1) throw new Exception("internal failure");
            LastRestore = content;
            Formula = content.IsFormula ? (string?)content.Value : null;
            Value = content.IsFormula ? 42d : content.Value;
        }
    }

    private sealed class FakeWorkbook : ISimulationWorkbook
    {
        public FakeCell Input = new("Sheet1!A1"), Second = new("Sheet1!A2"), Output = new("Sheet1!B1");
        public string? Failure;
        public int Resolutions, StateWrites, Calculations;
        private bool screen = true, events = true;
        private object? status = "Ready";
        public bool ScreenUpdating { get => screen; set { StateWrites++; if (value && Failure == "restoreScreen") throw new Exception("internal failure"); screen = value; } }
        public bool EnableEvents { get => events; set { StateWrites++; events = value; if (!value && Failure == "events") throw new Exception("internal failure"); } }
        public object? StatusBar { get => status; set { StateWrites++; status = value; if (Failure == "status" && !Equals(value, "Ready")) throw new Exception("internal failure"); } }
        public ISimulationCell Resolve(string sheet, string address)
        {
            Resolutions++;
            Input.FailWrite = Failure == "write";
            return address switch { "A1" => Input, "A2" => Second, "B1" => Output, _ => throw new SimulationCellAccessException("internal reference details") };
        }
        public void Calculate()
        {
            Calculations++;
            if (Failure == "calculate" && Calculations == 1) throw new Exception("internal failure");
            if (Failure == "cleanup" && Calculations > 1) throw new Exception("internal failure");
            if (Failure == "formula" && Calculations == 1) Output.Value = new ExcelCellError("#DIV/0!");
        }
    }
}
