using System.Text.Json;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public static partial class WorkbookPersistence
{
    // Y1 version, Z1 chunk count, Y2 onward JSON. Independent of settings and scenarios.
    private const string CorrelationVersion = "AssumptionCorrelationsV1";
    private static object?[,] CaptureCorrelationBlock(dynamic sheet)
    {
        int last = (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1;
        var raw = new object?[last, 2];
        for (int r = 0; r < last; r++) for (int c = 0; c < 2; c++) raw[r, c] = sheet.Cells[r + 1, 25 + c].Value2;
        return raw;
    }
    private static void WriteCorrelationBlock(dynamic sheet, object?[,] raw)
    {
        for (int r = 0; r < raw.GetLength(0); r++) for (int c = 0; c < 2; c++)
        {
            sheet.Cells[r + 1, 25 + c].NumberFormat = "@";
            sheet.Cells[r + 1, 25 + c].Value2 = raw[r, c];
        }
    }
    private static List<AssumptionCorrelation> ParseCorrelations(object?[,] raw)
    {
        if (raw.Cast<object?>().All(x => x == null)) return [];
        try
        {
            if (!Equals(raw[0, 0], CorrelationVersion) || !int.TryParse(Convert.ToString(raw[0, 1]), out int count) ||
                count < 1 || count >= raw.GetLength(0) || count > 1000) throw new FormatException();
            var json = new System.Text.StringBuilder();
            for (int i = 1; i <= count; i++) json.Append(raw[i, 0] is string chunk ? chunk : throw new FormatException());
            var values = JsonSerializer.Deserialize<List<AssumptionCorrelation>>(json.ToString()) ?? throw new FormatException();
            if (values.Any(x => x == null)) throw new FormatException();
            return values;
        }
        catch (Exception ex) when (ex is JsonException or FormatException)
        { throw new ArgumentException("Saved Assumption Correlations are invalid or use an unsupported version. Restore a valid workbook copy before running.", ex); }
    }
    public static List<AssumptionCorrelation> LoadCorrelations(object workbook)
    {
        dynamic? sheet = FindConfigSheet(workbook);
        return sheet == null ? [] : ParseCorrelations(CaptureCorrelationBlock(sheet));
    }
    private static object?[,] CorrelationBlock(IReadOnlyList<AssumptionCorrelation> values, int minimumRows = 1)
    {
        if (values.Count == 0) return new object?[minimumRows, 2];
        string json = JsonSerializer.Serialize(values);
        int chunks = (json.Length + 29999) / 30000;
        var raw = new object?[Math.Max(minimumRows, chunks + 1), 2];
        raw[0, 0] = CorrelationVersion; raw[0, 1] = chunks;
        for (int i = 0; i < chunks; i++) raw[i + 1, 0] = json.Substring(i * 30000, Math.Min(30000, json.Length - i * 30000));
        return raw;
    }
    public static void SaveCorrelations(object workbook, IReadOnlyList<AssumptionDefinition> assumptions, IReadOnlyList<AssumptionCorrelation> values)
    {
        _ = new GaussianDependence(assumptions.Select(ScenarioSimulationService.Identity).ToArray(), values);
        dynamic sheet = GetOrCreateConfigSheet(workbook);
        object?[,] previous = CaptureCorrelationBlock(sheet);
        var remap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var a in assumptions)
        {
            string old = ScenarioSimulationService.Identity(a);
            string link = EnsureCellLink(workbook, a.CellLink, a.SheetName, a.CellAddress);
            if (string.IsNullOrEmpty(link)) throw new ArgumentException($"The reference for {a.Name} cannot be tracked. Redefine the assumption first.");
            // Persist the link in its definition row without altering values, targets or formatting.
            int last = (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1;
            for (int row = 2; row <= last; row++)
                if (Equals(sheet.Cells[row, 1].Value2, "Assumption") &&
                    (Equals(sheet.Cells[row, 10].Value2, a.CellLink) && !string.IsNullOrEmpty(a.CellLink) ||
                     Equals(sheet.Cells[row, 2].Value2, a.SheetName) && Equals(sheet.Cells[row, 3].Value2, a.CellAddress)))
                {
                    sheet.Cells[1, 10].Value2 = "CellLink"; sheet.Cells[row, 10].NumberFormat = "@"; sheet.Cells[row, 10].Value2 = link;
                    break;
                }
            remap[old] = link; a.CellLink = link;
        }
        var saved = values.Where(r => r.Correlation != 0).Select(r => r with { AssumptionA = remap[r.AssumptionA], AssumptionB = remap[r.AssumptionB] }).ToArray();
        var block = CorrelationBlock(saved, previous.GetLength(0));
        var rollback = new object?[block.GetLength(0), 2];
        for (int r = 0; r < previous.GetLength(0); r++) for (int c = 0; c < 2; c++) rollback[r, c] = previous[r, c];
        try { WriteCorrelationBlock(sheet, block); }
        catch { WriteCorrelationBlock(sheet, rollback); throw; }
    }
    private static object?[,] CorrelationsForModelSave(dynamic sheet)
    {
        object?[,] raw = CaptureCorrelationBlock(sheet);
        try
        {
            var values = ParseCorrelations(raw);
            var ids = SimulationModel.Assumptions.Select(ScenarioSimulationService.Identity).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return CorrelationBlock(values.Where(r => ids.Contains(r.AssumptionA) && ids.Contains(r.AssumptionB)).ToArray(), raw.GetLength(0));
        }
        catch (ArgumentException) { return raw; } // Unrelated definition edits must not erase an unreadable block.
    }
}
