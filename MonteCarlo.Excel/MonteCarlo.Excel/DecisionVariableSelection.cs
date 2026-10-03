using System.Text.RegularExpressions;
namespace MonteCarlo.Excel;

public sealed record DecisionVariableSelection(string Sheet, string Cell, double CurrentValue, string? SuggestedName)
{
    public string Display => Sheet + "!" + Cell;
    public static DecisionVariableSelection? ReadRange(object source, object application, object? selected, ISimulationWorkbook workbook)
    {
        if (selected == null || selected is bool) return null;
        dynamic range = selected, app = application;
        if (!string.Equals((string)range.Worksheet.Parent.Name, (string)((dynamic)source).Name, StringComparison.OrdinalIgnoreCase) ||
            Convert.ToDouble(range.Cells.CountLarge) != 1 || (int)range.Areas.Count != 1)
            throw new ArgumentException("Select exactly one cell in the workbook where Optimizer was opened.");
        object? label = null;
        if ((int)range.Column > 1)
        {
            dynamic left = range.Worksheet.Cells[(int)range.Row, (int)range.Column - 1];
            if (!(bool)app.WorksheetFunction.IsError(left)) label = left.Value2;
        }
        return Read(workbook, (string)range.Worksheet.Name, (string)range.Address[true, true], label);
    }
    public static string Absolute(string address)
    {
        var match = Regex.Match(address, @"^\$?([A-Za-z]+)\$?([1-9][0-9]*)$");
        if (!match.Success) throw new ArgumentException("Select exactly one existing cell.");
        return "$" + match.Groups[1].Value.ToUpperInvariant() + "$" + match.Groups[2].Value;
    }
    public static string SuggestName(string entered, object? label) => !string.IsNullOrWhiteSpace(entered) ? entered :
        label is string text && !string.IsNullOrWhiteSpace(text) && text.Trim().Length <= 100 && !text.Any(char.IsControl) ? text.Trim() : entered;
    public static DecisionVariableSelection Read(ISimulationWorkbook workbook, string sheet, string cell, object? label = null)
    {
        var input = workbook.Resolve(sheet, cell);
        if (input.InputProblem != null) throw new ArgumentException($"{sheet}!{cell}: {input.InputProblem} Select an editable numeric input.");
        object? value = input.ReadValue();
        if (value is not (double or float or decimal or int or long or short or byte) || !double.IsFinite(Convert.ToDouble(value)))
            throw new ArgumentException($"{sheet}!{cell}: the cell is empty, text, an Excel error or nonnumeric. Select a finite numeric input.");
        return new(sheet, Absolute(cell), Convert.ToDouble(value), SuggestName("", label) is { Length: > 0 } name ? name : null);
    }
}
