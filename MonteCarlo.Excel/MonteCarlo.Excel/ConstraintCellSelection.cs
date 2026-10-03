using System.Globalization;
using System.Text.RegularExpressions;

namespace MonteCarlo.Excel;

public static class ConstraintCellSelection
{
    public static string Reference(string sheet, string cell) => "'" + sheet.Replace("'", "''") + "'!" + Absolute(cell);
    public static (string Sheet, string Cell) Parse(string text)
    {
        text = text.Trim();
        var match = Regex.Match(text, @"^(?:'((?:[^']|'')+)'|([^'!]+))!([^!]+)$");
        if (!match.Success) throw new ArgumentException("Constraint: enter a single cell with its worksheet, for example Sheet1!$B$6 or 'My Sheet'!$B$6.");
        string sheet = match.Groups[1].Success ? match.Groups[1].Value.Replace("''", "'") : match.Groups[2].Value.Trim();
        if (string.IsNullOrWhiteSpace(sheet) || sheet.IndexOfAny(['[', ']', ':', '\\', '/', '?', '*']) >= 0)
            throw new ArgumentException("Constraint: select a worksheet in this workbook; external workbook references are unsupported.");
        return (sheet, Absolute(match.Groups[3].Value.Trim()));
    }
    public static string Absolute(string address)
    {
        var match = Regex.Match(address, @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$");
        int column = 0;
        if (match.Success) foreach (char ch in match.Groups[1].Value.ToUpperInvariant()) column = column * 26 + ch - 'A' + 1;
        if (!match.Success || column > 16384 || !int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int row) || row > 1048576)
            throw new ArgumentException($"Constraint reference '{address}' is invalid. Select one existing cell, not a range or deleted reference.");
        return DecisionVariableSelection.Absolute(address);
    }
    public static double Number(object? value, string description)
    {
        if (value is string || !SimulationValidation.TryNumber(value, out double number))
            throw new ArgumentException($"Constraint \"{description}\" does not currently contain a finite numeric value (it is empty, text, or an Excel error{(value is ExcelCellError error ? ": " + error.Name : "")}). Correct the formula/value or select another cell.");
        return number;
    }
    public static DecisionVariableSelection Read(ISimulationWorkbook workbook, string sheet, string cell, object? label = null)
    {
        cell = Absolute(cell);
        try
        {
            var resolved = workbook.Resolve(sheet, cell);
            // A constraint is read-only: protected, array and spill formula outputs are valid.
            return new(sheet, cell, Number(resolved.ReadValue(), sheet + "!" + cell), DecisionVariableSelection.SuggestName("", label));
        }
        catch (SimulationCellAccessException ex) { throw new ArgumentException($"Constraint {sheet}!{cell}: the worksheet/reference is unavailable. Select one existing cell in this workbook.", ex); }
    }
    public static DecisionVariableSelection? ReadRange(object source, object application, object? selected, ISimulationWorkbook workbook)
    {
        if (selected == null || selected is bool) return null;
        dynamic range = selected, app = application;
        if (!string.Equals((string)range.Worksheet.Parent.Name, (string)((dynamic)source).Name, StringComparison.OrdinalIgnoreCase) ||
            Convert.ToDouble(range.Cells.CountLarge) != 1 || (int)range.Areas.Count != 1)
            throw new ArgumentException("Constraint: select exactly one cell in the workbook where Optimizer was opened.");
        object? label = null;
        if ((int)range.Column > 1)
        {
            dynamic left = range.Worksheet.Cells[(int)range.Row, (int)range.Column - 1];
            if (!(bool)app.WorksheetFunction.IsError(left)) label = left.Value2;
        }
        return Read(workbook, (string)range.Worksheet.Name, (string)range.Address[true, true], label);
    }
}
