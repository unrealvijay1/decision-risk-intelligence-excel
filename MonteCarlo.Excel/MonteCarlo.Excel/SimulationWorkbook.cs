namespace MonteCarlo.Excel;

// Minimal boundary around Excel operations; the engine can be tested without COM.
public sealed record SimulationCellContent(object? Value, bool IsFormula = false, bool UsesFormula2 = false);

public interface ISimulationCell
{
    string Identity { get; }
    string? InputProblem { get; }
    object? ReadValue();
    SimulationCellContent Capture();
    void WriteSample(double value);
    void Restore(SimulationCellContent content);
}

public interface ISimulationWorkbook
{
    ISimulationCell Resolve(string sheet, string address);
    bool ScreenUpdating { get; set; }
    bool EnableEvents { get; set; }
    object? StatusBar { get; set; }
    void Calculate();
}

public sealed class SimulationCellAccessException : Exception
{
    public SimulationCellAccessException(string message, Exception? inner = null) : base(message, inner) { }
}

public sealed class SimulationRunException : Exception
{
    public SimulationRunException(SimulationExecutionResult outcome)
        : base(outcome.UserMessage, outcome.DiagnosticException) { }
}

public sealed class SimulationExecutionResult
{
    public SimulationRunResult? Result { get; internal set; }
    public ValidationResult Validation { get; } = new();
    public Exception? DiagnosticException { get; internal set; }
    public List<string> RestorationFailures { get; } = new();
    public bool Succeeded => Result != null && Validation.IsValid && DiagnosticException == null && RestorationFailures.Count == 0;
    public string UserMessage
    {
        get
        {
            string message = !Validation.IsValid ? Validation.UserMessage : DiagnosticException != null
                ? "The simulation could not finish. Check that the workbook is open and available, then try again."
                : "";
            if (RestorationFailures.Count != 0)
                message += "\n\nExcel could not restore: " + string.Join(", ", RestorationFailures) +
                    ". Check these cells/settings before saving or running again. If necessary, reopen the last saved workbook without saving this run.";
            return message.Trim();
        }
    }
}
