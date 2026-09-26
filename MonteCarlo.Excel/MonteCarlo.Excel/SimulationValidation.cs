using System.Globalization;
using System.Runtime.InteropServices;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public enum ValidationSeverity { Warning, Error }

public sealed record ValidationError(string Location, string UserMessage, string SuggestedCorrection,
    ValidationSeverity Severity = ValidationSeverity.Error);

public sealed class ValidationResult
{
    public List<ValidationError> Errors { get; } = new();
    public bool IsValid => !Errors.Any(error => error.Severity == ValidationSeverity.Error);
    public void Add(string location, string message, string correction) =>
        Errors.Add(new(location, message, correction));
    public string UserMessage => string.Join("\n\n", Errors.Select(error =>
        $"{error.Location}: {error.UserMessage}\n{error.SuggestedCorrection}"));
}

/// <summary>Explicit Excel error identity, supplied by the COM adapter before numeric conversion.</summary>
public sealed record ExcelCellError(string Name);

public static class SimulationValidation
{
    public static bool TryNumber(object? value, out double number)
    {
        number = 0;
        if (value == null || value is bool || value is ExcelCellError || value is ErrorWrapper || value is Array)
            return false;
        // COM may marshal VT_ERROR as its signed HRESULT. Ordinary numeric 2042 is not #N/A.
        if (value is int code && (code & unchecked((int)0xFFFF0000)) == unchecked((int)0x800A0000))
            return false;
        if (value is string text)
            return double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.CurrentCulture, out number) && double.IsFinite(number);
        if (value is not (double or float or decimal or byte or sbyte or short or ushort or int or uint or long or ulong))
            return false;
        number = Convert.ToDouble(value, CultureInfo.CurrentCulture);
        return double.IsFinite(number);
    }

    public static void ValidateValue(ValidationResult result, string location, object? value)
    {
        if (TryNumber(value, out _)) return;
        string message = value is ExcelCellError error ? $"The cell contains Excel error {error.Name}."
            : value == null || value is string text && string.IsNullOrWhiteSpace(text)
                ? "The cell is empty."
                : "The cell does not contain a finite numeric value (it may contain an Excel error).";
        result.Add(location, message, "Correct the cell value or formula so it returns a number, then run again.");
    }

    public static void ValidateParameters(ValidationResult result, string location,
        DistributionType distribution, IReadOnlyList<object?> parameters)
    {
        string[] names = distribution switch
        {
            DistributionType.Normal => ["Mean", "Standard deviation"],
            DistributionType.Lognormal => ["Log mean", "Log standard deviation"],
            DistributionType.Uniform => ["Minimum", "Maximum"],
            DistributionType.Triangular or DistributionType.Pert => ["Minimum", "Most likely", "Maximum"],
            DistributionType.Beta => ["Minimum", "Maximum", "Alpha", "Beta"],
            _ => []
        };
        if (names.Length == 0)
        {
            result.Add(location, "The distribution is not supported.", "Choose a supported distribution in Define Assumption.");
            return;
        }
        double[] values = new double[4];
        bool valid = true;
        for (int i = 0; i < names.Length; i++)
        {
            object? value = i < parameters.Count ? parameters[i] : null;
            if (!TryNumber(value, out values[i]))
            {
                valid = false;
                result.Add(location, $"{names[i]} is missing or is not a finite number.",
                    $"Open Define Assumption and enter a numeric {names[i]}.");
            }
        }
        if (!valid) return;
        // Same rules used by the existing assumption editor; sampling remains unchanged.
        string? error = DistributionPreview.Validate((DistributionKind)distribution,
            values[0], values[1], values[2], values[3]);
        if (error != null)
            result.Add(location, error, "Correct these distribution parameters in Define Assumption.");
    }

    public static ValidationResult ValidateModel(int trials, IReadOnlyList<AssumptionDefinition> assumptions,
        IReadOnlyList<ForecastDefinition> forecasts)
    {
        var result = new ValidationResult();
        if (trials <= 0) result.Add("Simulation settings", "The trial count must be positive.", "Choose a positive number of trials.");
        if (assumptions.Count == 0) result.Add("Assumptions", "No assumptions are configured.", "Select an input cell and choose Define Assumption.");
        if (forecasts.Count == 0) result.Add("Forecasts", "No forecast is configured.", "Select an output cell and choose Define Forecast.");
        foreach (var assumption in assumptions)
        {
            string location = Location("Assumption", assumption.Name, assumption.SheetName, assumption.CellAddress);
            ValidateReference(result, location, assumption.SheetName, assumption.CellAddress);
            ValidateParameters(result, location, assumption.Distribution,
                [assumption.Parameter1, assumption.Parameter2, assumption.Parameter3, assumption.Parameter4]);
        }
        foreach (var forecast in forecasts)
            ValidateReference(result, Location("Forecast", forecast.Name, forecast.SheetName, forecast.CellAddress),
                forecast.SheetName, forecast.CellAddress);
        return result;
    }

    public static string Location(string kind, string name, string sheet, string address) =>
        $"{kind} {name} ({sheet}!{address})";

    private static void ValidateReference(ValidationResult result, string location, string sheet, string address)
    {
        if (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(address))
            result.Add(location, "A worksheet or cell reference is missing.", "Select one cell in the correct worksheet and redefine this item.");
    }
}
