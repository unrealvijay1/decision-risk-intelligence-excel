using System.Text.Json;
namespace MonteCarlo.Excel;

public static partial class WorkbookPersistence
{
    // AD:AE independent chunked envelope, preserved even when unreadable.
    private static object?[,] CaptureOptimizerBlock(dynamic sheet)
    {
        int last = (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1;
        var block = new object?[last, 2];
        for (int r = 0; r < last; r++) for (int c = 0; c < 2; c++) block[r, c] = sheet.Cells[r + 1, 30 + c].Value2;
        return block;
    }
    private static void RestoreOptimizerBlock(dynamic sheet, object?[,] block)
    {
        for (int r = 0; r < block.GetLength(0); r++) for (int c = 0; c < 2; c++)
        { if (block[r, c] == null) continue; sheet.Cells[r + 1, 30 + c].NumberFormat = "@"; sheet.Cells[r + 1, 30 + c].Value2 = block[r, c]; }
    }
    public static OptimizationProblem? LoadOptimizer(object source)
    {
        dynamic? sheet = FindConfigSheet(source); if (sheet == null) return null;
        object?[,] raw = CaptureOptimizerBlock(sheet); if (raw.Cast<object?>().All(v => v == null)) return null;
        try
        {
            if (!Equals(raw[0, 0], "OptimizerV1") || !int.TryParse(Convert.ToString(raw[0, 1]), out int count) || count < 1 || count > 1000 || count >= raw.GetLength(0)) throw new FormatException();
            string json = "";
            for (int i = 1; i <= count; i++) json += raw[i, 0] as string ?? throw new FormatException();
            var p = JsonSerializer.Deserialize<OptimizationProblem>(json) ?? throw new FormatException(); p.Validate();
            return ResolveOptimizer(source, p);
        }
        catch (Exception ex) when (ex is JsonException or FormatException or ArgumentException or NullReferenceException)
        { throw new InvalidOperationException("Saved Optimizer settings are invalid or unsupported. Restore a valid workbook copy. Other analyses remain available.", ex); }
    }
    public static OptimizationProblem ResolveOptimizer(object source, OptimizationProblem p) => p with
    {
        Variables = p.Variables.Select(v =>
        {
            string sheet = v.Sheet, cell = v.Cell;
            ResolveSavedLink(source, v.CellLink, ref sheet, ref cell, null, v.Name);
            return v with { Sheet = sheet, Cell = cell };
        }).ToArray(),
        Constraints = p.Constraints.Select(c =>
        {
            if (c.Kind != OptimizationConstraintKind.Cell) return c;
            string sheet = c.Sheet, cell = c.Cell;
            ResolveSavedLink(source, c.CellLink, ref sheet, ref cell, null, c.CellDescription);
            return c with { Sheet = sheet, Cell = cell };
        }).ToArray()
    };
    public static OptimizationProblem SaveOptimizer(object source, OptimizationProblem p, IReadOnlyList<ForecastDefinition>? forecasts = null)
    {
        p.Validate(); dynamic workbook = source; dynamic sheet = GetOrCreateConfigSheet(workbook);
        var remap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in forecasts ?? [])
        {
            if (!string.IsNullOrEmpty(f.CellLink)) continue;
            int last = (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1;
            for (int row = 2; row <= last; row++)
            {
                if (!Equals(sheet.Cells[row, 1].Value2, "Forecast") || !string.Equals(Convert.ToString(sheet.Cells[row, 2].Value2), f.SheetName, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Convert.ToString(sheet.Cells[row, 3].Value2), f.CellAddress, StringComparison.OrdinalIgnoreCase)) continue;
                string link = EnsureCellLink(workbook, Convert.ToString(sheet.Cells[row, 10].Value2) ?? "", f.SheetName, f.CellAddress);
                if (string.IsNullOrEmpty(link)) throw new ArgumentException($"{f.Name}: cannot track forecast. Redefine this forecast before saving Optimizer settings.");
                remap[ScenarioSimulationService.Identity(f)] = link; f.CellLink = link;
                sheet.Cells[1, 10].Value2 = "CellLink"; sheet.Cells[row, 10].NumberFormat = "@"; sheet.Cells[row, 10].Value2 = link;
                break;
            }
        }
        p = p with { Objective = remap.TryGetValue(p.Objective.ForecastId, out var objectiveLink) ? p.Objective with { ForecastId = objectiveLink } : p.Objective,
            Constraints = p.Constraints.Select(c => c.Kind is not (OptimizationConstraintKind.DecisionVariable or OptimizationConstraintKind.Cell or OptimizationConstraintKind.Linear) && remap.TryGetValue(c.SubjectId, out var link) ? c with { SubjectId = link } : c).ToArray() };
        p = ResolveOptimizer(source, p);
        p = p with { Variables = p.Variables.Select(v =>
        {
            if (v.Cell == "#REF!") throw new ArgumentException($"{v.Name}: tracked cell was deleted. Edit or remove the variable.");
            string link = EnsureCellLink(workbook, v.CellLink, v.Sheet, v.Cell);
            if (string.IsNullOrEmpty(link)) throw new ArgumentException($"{v.Name}: cannot track this cell. Select an existing cell.");
            return v with { CellLink = link };
        }).ToArray(), Constraints = p.Constraints.Select(c =>
        {
            if (c.Kind != OptimizationConstraintKind.Cell) return c;
            ConstraintCellSelection.Absolute(c.Cell);
            string link = EnsureCellLink(workbook, c.CellLink, c.Sheet, c.Cell);
            if (string.IsNullOrEmpty(link)) throw new ArgumentException($"Constraint \"{c.CellDescription}\": cannot track this cell. Select an existing cell in this workbook.");
            return c with { CellLink = link };
        }).ToArray() };
        var previous = CaptureOptimizerBlock(sheet); string json = JsonSerializer.Serialize(p);
        int count = (json.Length + 29999) / 30000;
        if (count > 1000) throw new ArgumentException("Optimizer configuration is too large to save.");
        void Clear() { for (int r = 1; r <= Math.Max(previous.GetLength(0), count + 1); r++) for (int c = 30; c <= 31; c++) sheet.Cells[r, c].Value2 = null; }
        try
        {
            Clear(); sheet.Cells[1, 30].Value2 = "OptimizerV1"; sheet.Cells[1, 31].Value2 = count;
            for (int i = 0; i < count; i++) { sheet.Cells[i + 2, 30].NumberFormat = "@"; sheet.Cells[i + 2, 30].Value2 = json.Substring(i * 30000, Math.Min(30000, json.Length - i * 30000)); }
            sheet.Visible = 2;
        }
        catch { Clear(); RestoreOptimizerBlock(sheet, previous); throw; }
        return p;
    }
}
