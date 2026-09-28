using System.Globalization;

namespace MonteCarlo.Excel;

public static class SpcDataReader
{
    public const int MaximumObservations = 100000;
    public static double[] ParseValues(object? raw, int count)
    {
        if (count < 20 || count > MaximumObservations)
            throw new ArgumentException($"Select 20–{MaximumObservations:N0} consecutive numerical observations, excluding headers.");
        var cells = Flatten(raw).ToArray();
        if (cells.Length != count) throw new ArgumentException("The selected range could not be read. Select one contiguous row or column.");
        return cells.Select((v, i) => v is double or float or int or long or short or decimal
            && double.IsFinite(Convert.ToDouble(v, CultureInfo.InvariantCulture))
            ? Convert.ToDouble(v, CultureInfo.InvariantCulture)
            : throw new ArgumentException($"Observation {i + 1} is blank, text, an Excel error or non-finite. Correct that cell or select a complete consecutive period. Gaps are not skipped.")).ToArray();
    }
    private static IEnumerable<object?> Flatten(object? raw)
    {
        if (raw is object[,] array)
        {
            for (int r = array.GetLowerBound(0); r <= array.GetUpperBound(0); r++)
                for (int c = array.GetLowerBound(1); c <= array.GetUpperBound(1); c++) yield return array[r, c];
        }
        else yield return raw;
    }
    public static int ValidateRange(object workbook, dynamic range)
    {
        dynamic book = workbook;
        if (!string.Equals((string)range.Worksheet.Parent.Name, (string)book.Name, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Select a range in the workbook where SPC Analysis was opened.");
        if ((int)range.Areas.Count != 1 || ((int)range.Rows.Count != 1 && (int)range.Columns.Count != 1))
            throw new ArgumentException("Select one contiguous row or column in observation order.");
        int count = checked((int)(double)range.Cells.CountLarge);
        if (count > MaximumObservations) throw new ArgumentException($"Select at most {MaximumObservations:N0} observations.");
        return count;
    }
    public static double[] ReadValues(object workbook, dynamic range)
    {
        int count = ValidateRange(workbook, range);
        // Excel can marshal an error as an integer HRESULT. Count excludes errors; never
        // mistake such a value for a numerical observation (ordinary 2042 remains valid).
        dynamic app = range.Application;
        if ((double)app.WorksheetFunction.Count(range) != count)
            for (int i = 0; i < count; i++)
                if ((bool)app.WorksheetFunction.IsError(range.Cells[i + 1]))
                    throw new ArgumentException($"Observation {i + 1} contains an Excel error. Correct it before analysis; gaps cannot be skipped.");
        return ParseValues((object?)range.Value2, count);
    }
    public static string Describe(dynamic range) => $"{range.Worksheet.Name}!{range.Address[false, false]}";
    public static string Track(object workbook, object range)
    {
        dynamic book = workbook, selected = range;
        ValidateRange(workbook, selected);
        string link = "_MC_SPC_" + Guid.NewGuid().ToString("N");
        book.Names.Add(Name: link, RefersTo: "='" + ((string)selected.Worksheet.Name).Replace("'", "''") + "'!" + selected.Address[true, true], Visible: false);
        return link;
    }
    public static string TrackOrReuse(object workbook, object range, string? previousLink)
    {
        if (previousLink != null)
        {
            try
            {
                dynamic old = Resolve(workbook, previousLink), next = range;
                if ((string)old.Worksheet.Name == (string)next.Worksheet.Name &&
                    (string)old.Address[true, true] == (string)next.Address[true, true]) return previousLink;
            }
            catch (ArgumentException) { } // A newly selected range deliberately repairs a broken name.
        }
        return Track(workbook, range);
    }
    public static object Resolve(object workbook, string link)
    {
        try
        {
            if (!link.StartsWith("_MC_SPC_", StringComparison.Ordinal)) throw new ArgumentException();
            dynamic book = workbook;
            object range = book.Names[link].RefersToRange;
            ValidateRange(workbook, range);
            return range;
        }
        catch (Exception ex) { throw new ArgumentException("An SPC source or baseline reference is broken. Reselect the source and explicitly establish a new baseline; stale addresses are never reused.", ex); }
    }
    public static (int Start, int End) LocateBaseline(object workbook, dynamic data, dynamic baseline)
    {
        int n = ValidateRange(workbook, baseline);
        if ((string)data.Worksheet.Name != (string)baseline.Worksheet.Name)
            throw new ArgumentException("The baseline must be inside the observation range.");
        bool vertical = (int)data.Columns.Count == 1;
        int start = vertical ? (int)baseline.Row - (int)data.Row + 1 : (int)baseline.Column - (int)data.Column + 1;
        if ((vertical ? (int)baseline.Column != (int)data.Column || (int)baseline.Columns.Count != 1
                      : (int)baseline.Row != (int)data.Row || (int)baseline.Rows.Count != 1) ||
            start < 1 || start + n - 1 > ValidateRange(workbook, data))
            throw new ArgumentException("The baseline must be a contiguous period inside the observation range.");
        return (start, start + n - 1);
    }
    public static string[] ReadLabels(object workbook, dynamic range, int count)
    {
        if (ValidateRange(workbook, range) != count) throw new ArgumentException("Date/sequence range must contain exactly one label per observation.");
        var labels = new string[count];
        for (int i = 0; i < count; i++)
        {
            dynamic cell = range.Cells[i + 1];
            object? value = cell.Value;
            labels[i] = value is DateTime date ? date.ToString("d", CultureInfo.CurrentCulture) : Convert.ToString(cell.Text) ?? "";
            if (labels[i].Length > 0 && labels[i].All(c => c == '#')) labels[i] = Convert.ToString(cell.Value2, CultureInfo.CurrentCulture) ?? "";
        }
        return labels;
    }
}
