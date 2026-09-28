using System.Text.Json;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed record SpcSavedAnalysis(SpcConfiguration Configuration, string DataLink, string? LabelLink,
    string BaselineLink, SpcBaseline Baseline);

public static partial class WorkbookPersistence
{
    // AB:AC is independent of model definitions, settings, scenarios and correlations.
    private static object?[,] CaptureSpcBlock(dynamic sheet)
    {
        int last = (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1;
        var raw = new object?[last, 2];
        for (int r = 0; r < last; r++) for (int c = 0; c < 2; c++) raw[r, c] = sheet.Cells[r + 1, 28 + c].Value2;
        return raw;
    }
    private static void WriteSpcBlock(dynamic sheet, object?[,] raw)
    {
        for (int r = 0; r < raw.GetLength(0); r++) for (int c = 0; c < 2; c++)
        {
            sheet.Cells[r + 1, 28 + c].NumberFormat = "@";
            sheet.Cells[r + 1, 28 + c].Value2 = raw[r, c];
        }
    }
    public static SpcSavedAnalysis? LoadSpcAnalysis(object workbook)
    {
        dynamic? sheet = FindConfigSheet(workbook);
        if (sheet == null) return null;
        object?[,] raw = CaptureSpcBlock(sheet);
        if (raw.Cast<object?>().All(x => x == null)) return null;
        try
        {
            if (!Equals(raw[0, 0], "SpcAnalysisV1") || !int.TryParse(Convert.ToString(raw[0, 1]), out int count) ||
                count < 1 || count >= raw.GetLength(0) || count > 1000) throw new FormatException();
            string json = string.Concat(Enumerable.Range(1, count).Select(i => raw[i, 0] as string ?? throw new FormatException()));
            var saved = JsonSerializer.Deserialize<SpcSavedAnalysis>(json) ?? throw new FormatException();
            if (saved.Configuration == null || saved.Baseline == null || string.IsNullOrEmpty(saved.DataLink) || string.IsNullOrEmpty(saved.BaselineLink))
                throw new FormatException();
            if (string.IsNullOrWhiteSpace(saved.Configuration.ProcessName) || !Enum.IsDefined(saved.Configuration.BaselineMode) || saved.Configuration.Rules == null ||
                saved.Configuration.Target is { } target && (!double.IsFinite(target.Value) || !Enum.IsDefined(target.Direction))) throw new FormatException();
            XmRCalculator.ValidateBaseline(saved.Baseline, saved.Baseline.End);
            return saved;
        }
        catch (Exception ex) when (ex is JsonException or FormatException or ArgumentException)
        { throw new ArgumentException("Saved SPC configuration is unreadable or uses an unsupported version. Restore a valid workbook copy.", ex); }
    }
    public static void SaveSpcAnalysis(object workbook, SpcSavedAnalysis saved)
    {
        string json = JsonSerializer.Serialize(saved);
        int chunks = (json.Length + 29999) / 30000;
        dynamic sheet = GetOrCreateConfigSheet(workbook);
        object?[,] old = CaptureSpcBlock(sheet);
        var next = new object?[Math.Max(old.GetLength(0), chunks + 1), 2];
        var rollback = new object?[next.GetLength(0), 2];
        for (int r = 0; r < old.GetLength(0); r++) for (int c = 0; c < 2; c++) rollback[r, c] = old[r, c];
        next[0, 0] = "SpcAnalysisV1"; next[0, 1] = chunks;
        for (int i = 0; i < chunks; i++) next[i + 1, 0] = json.Substring(i * 30000, Math.Min(30000, json.Length - i * 30000));
        try { WriteSpcBlock(sheet, next); sheet.Visible = 2; }
        catch { WriteSpcBlock(sheet, rollback); throw; }
    }
}
