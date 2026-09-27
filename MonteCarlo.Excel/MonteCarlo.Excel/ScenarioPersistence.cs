using System.Text.Json;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public static partial class WorkbookPersistence
{
    // Independent versioned block: V1 marker, W1 chunk count, V2 onward JSON chunks.
    // Chunking respects Excel's 32,767-character cell limit, including multi-input scenarios.
    private const string ScenarioVersion = "ScenarioDefinitionsV2";
    private const int ScenarioChunkLength = 30000;
    private static object?[,] CaptureScenarioBlock(dynamic sheet)
    {
        int last = (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1;
        var block = new object?[last, 2];
        for (int r = 0; r < last; r++)
        for (int c = 0; c < 2; c++) block[r, c] = sheet.Cells[r + 1, 22 + c].Value2;
        return block;
    }
    private static void RestoreScenarioBlock(dynamic sheet, object?[,] block)
    {
        for (int r = 0; r < block.GetLength(0); r++)
        for (int c = 0; c < 2; c++)
        {
            if (block[r, c] == null) continue;
            sheet.Cells[r + 1, 22 + c].NumberFormat = "@";
            sheet.Cells[r + 1, 22 + c].Value2 = block[r, c];
        }
    }

    public static List<ScenarioDefinition> LoadScenarios(object sourceWorkbook) => LoadScenarioAnalysis(sourceWorkbook).Scenarios.ToList();

    public static ScenarioAnalysisDefinition LoadScenarioAnalysis(object sourceWorkbook)
    {
        dynamic? sheet = FindConfigSheet(sourceWorkbook);
        if (sheet == null) return new([]);
        object?[,] raw = CaptureScenarioBlock(sheet);
        if (raw.Cast<object?>().All(v => v == null)) return new([]);
        try
        {
            bool legacy = Equals(raw[0, 0], "ScenarioDefinitionsV1");
            if ((!legacy && !Equals(raw[0, 0], ScenarioVersion)) || !int.TryParse(Convert.ToString(raw[0, 1]), out int count) ||
                count < 1 || count > 1000 || count >= raw.GetLength(0)) throw new FormatException();
            var json = new System.Text.StringBuilder();
            for (int i = 1; i <= count; i++)
            {
                if (raw[i, 0] is not string chunk) throw new FormatException();
                json.Append(chunk);
            }
            var analysis = legacy
                ? new ScenarioAnalysisDefinition(JsonSerializer.Deserialize<List<ScenarioDefinition>>(json.ToString()) ?? throw new FormatException())
                : JsonSerializer.Deserialize<ScenarioAnalysisDefinition>(json.ToString()) ?? throw new FormatException();
            analysis.ValidateStored();
            return analysis;
        }
        catch (Exception ex) when (ex is JsonException or ArgumentException or FormatException)
        {
            throw new InvalidOperationException("The saved scenario definitions are invalid or use an unsupported version. Restore a valid workbook copy before editing scenarios. Normal simulations remain available.", ex);
        }
    }

    public static List<ScenarioDefinition> SaveScenarios(object sourceWorkbook, IReadOnlyList<ScenarioDefinition> scenarios,
        IReadOnlyList<AssumptionDefinition> assumptions) => SaveScenarioAnalysis(sourceWorkbook,
            LoadScenarioAnalysis(sourceWorkbook).WithScenarios(scenarios), assumptions, []).Scenarios.ToList();

    public static ScenarioAnalysisDefinition SaveScenarioAnalysis(object sourceWorkbook, ScenarioAnalysisDefinition analysis,
        IReadOnlyList<AssumptionDefinition> assumptions, IReadOnlyList<ForecastDefinition> forecasts)
    {
        analysis.ValidateStored();
        var scenarios = analysis.Scenarios;
        dynamic workbook = sourceWorkbook;
        dynamic sheet = GetOrCreateConfigSheet(workbook);
        // Upgrade coordinate-only legacy definitions by saving their tracked identity, not their parameters/fills.
        var remap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        int last = (int)sheet.UsedRange.Row + (int)sheet.UsedRange.Rows.Count - 1;
        foreach (var a in assumptions.Where(a => string.IsNullOrEmpty(a.CellLink)))
        {
            for (int row = 2; row <= last; row++)
            {
                if (!Equals(sheet.Cells[row, 1].Value2, "Assumption") ||
                    !string.Equals(Convert.ToString(sheet.Cells[row, 2].Value2), a.SheetName, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Convert.ToString(sheet.Cells[row, 3].Value2), a.CellAddress, StringComparison.OrdinalIgnoreCase)) continue;
                string old = ScenarioSimulationService.Identity(a);
                string link = Convert.ToString(sheet.Cells[row, 10].Value2) ?? "";
                link = EnsureCellLink(workbook, link, a.SheetName, a.CellAddress);
                if (string.IsNullOrEmpty(link)) throw new InvalidOperationException("The assumption reference could not be tracked. Check the worksheet and cell.");
                sheet.Cells[1, 10].Value2 = "CellLink";
                sheet.Cells[row, 10].NumberFormat = "@";
                sheet.Cells[row, 10].Value2 = link;
                remap[old] = link;
                a.CellLink = link;
                break;
            }
        }
        var saved = scenarios.Select(s => s with { Name = s.Name.Trim(), Adjustments = s.Adjustments.Select(a =>
            remap.TryGetValue(a.AssumptionId, out var link) ? a with { AssumptionId = link } : a).ToArray() }).ToList();
        var forecastRemap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var forecast in forecasts.Where(f => string.IsNullOrEmpty(f.CellLink)))
        {
            for (int row = 2; row <= last; row++)
            {
                if (!Equals(sheet.Cells[row, 1].Value2, "Forecast") ||
                    !string.Equals(Convert.ToString(sheet.Cells[row, 2].Value2), forecast.SheetName, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Convert.ToString(sheet.Cells[row, 3].Value2), forecast.CellAddress, StringComparison.OrdinalIgnoreCase)) continue;
                string old = ScenarioSimulationService.Identity(forecast);
                string link = EnsureCellLink(workbook, Convert.ToString(sheet.Cells[row, 10].Value2) ?? "", forecast.SheetName, forecast.CellAddress);
                if (string.IsNullOrEmpty(link)) throw new InvalidOperationException("The forecast reference could not be tracked. Check the cell.");
                sheet.Cells[1, 10].Value2 = "CellLink"; sheet.Cells[row, 10].NumberFormat = "@"; sheet.Cells[row, 10].Value2 = link;
                forecastRemap[old] = link; forecast.CellLink = link;
                break;
            }
        }
        var savedAnalysis = analysis.WithScenarios(saved) with { ForecastQuestions = analysis.Questions.Select(q =>
            forecastRemap.TryGetValue(q.ForecastId, out var link) ? q with { ForecastId = link } : q).ToArray() };
        string json = JsonSerializer.Serialize(savedAnalysis);
        int count = (json.Length + ScenarioChunkLength - 1) / ScenarioChunkLength;
        if (count > 1000) throw new ArgumentException("There are too many scenario adjustments to save.");
        var previous = CaptureScenarioBlock(sheet);
        try
        {
            for (int row = 1; row <= Math.Max(last, count + 1); row++)
            { sheet.Cells[row, 22].Value2 = null; sheet.Cells[row, 23].Value2 = null; }
            sheet.Cells[1, 22].Value2 = ScenarioVersion;
            sheet.Cells[1, 23].Value2 = count;
            for (int i = 0; i < count; i++)
            {
                sheet.Cells[i + 2, 22].NumberFormat = "@";
                sheet.Cells[i + 2, 22].Value2 = json.Substring(i * ScenarioChunkLength, Math.Min(ScenarioChunkLength, json.Length - i * ScenarioChunkLength));
            }
            sheet.Visible = 2;
        }
        catch
        {
            for (int row = 1; row <= Math.Max(last, count + 1); row++)
            { sheet.Cells[row, 22].Value2 = null; sheet.Cells[row, 23].Value2 = null; }
            RestoreScenarioBlock(sheet, previous);
            throw;
        }
        return savedAnalysis;
    }
}
