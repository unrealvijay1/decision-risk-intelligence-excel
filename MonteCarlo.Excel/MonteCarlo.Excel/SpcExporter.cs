using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public static class SpcExporter
{
    public const int HeaderRow = 24;
    public static object?[,] ObservationTable(SpcResult result)
    {
        string[] headers = ["Observation", "Date / sequence", "Actual value", "Moving range", "Baseline", "Individuals UCL", "Individuals center", "Individuals LCL",
            "MR UCL", "MR center", "MR LCL", "Business target", "Signal", "Signal description", "Individuals signal", "MR signal", "Baseline start", "Baseline end", "Lower 1σ", "Upper 1σ", "Lower 2σ", "Upper 2σ"];
        var data = new object?[result.Observations.Count + 1, headers.Length];
        for (int c = 0; c < headers.Length; c++) data[0, c] = headers[c];
        var byIndex = result.Signals.SelectMany(s => s.QualifyingIndices.Select(index => (index, signal: s))).GroupBy(x => x.index).ToDictionary(g => g.Key, g => g.Select(x => x.signal).ToArray());
        var b = result.Baseline;
        for (int i = 0; i < result.Observations.Count; i++)
        {
            int r = i + 1; SpcSignal[] signals = byIndex.GetValueOrDefault(r) ?? [];
            data[r, 0] = r; data[r, 1] = result.Labels[i]; data[r, 2] = result.Observations[i]; data[r, 3] = result.MovingRanges[i];
            data[r, 4] = r >= b.Start && r <= b.End ? "Yes" : "No";
            data[r, 5] = b.IndividualsUpper; data[r, 6] = b.Mean; data[r, 7] = b.IndividualsLower;
            data[r, 8] = b.MovingRangeUpper; data[r, 9] = b.AverageMovingRange; data[r, 10] = 0;
            data[r, 11] = (object?)result.Configuration.Target?.Value;
            data[r, 12] = string.Join(", ", signals.Select(s => $"Rule {s.RuleId} ({s.Chart})").Distinct());
            data[r, 13] = string.Join("; ", signals.Select(s => $"Rule {s.RuleId}, {s.Chart}, #{s.StartIndex}–#{s.EndIndex}, {s.Direction}: {s.Description}; qualifying {s.QualifyingObservations}; window values [{s.Values}], reference {s.ReferenceValue:G6}"));
            data[r, 14] = signals.Any(s => s.Chart == SpcChart.Individuals) ? result.Observations[i] : null;
            data[r, 15] = signals.Any(s => s.Chart == SpcChart.MovingRange) ? result.MovingRanges[i] : null;
            data[r, 16] = r == b.Start ? result.Observations[i] : null;
            data[r, 17] = r == b.End ? result.Observations[i] : null;
            data[r, 18] = b.LowerOneSigma; data[r, 19] = b.UpperOneSigma;
            data[r, 20] = b.LowerTwoSigma; data[r, 21] = b.UpperTwoSigma;
        }
        return data;
    }
    public static void Export(object workbook, SpcResult result, string sourceDescription)
    {
        dynamic book = workbook;
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 1; i <= (int)book.Worksheets.Count; i++) names.Add((string)book.Worksheets[i].Name);
        string stem = "SPC_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"), name = stem;
        for (int i = 2; names.Contains(name); i++) name = stem + "_" + i;
        dynamic sheet = book.Worksheets.Add(); sheet.Name = name;
        var b = result.Baseline; var c = result.Configuration;
        string target = c.Target is { } t ? $"{t.Value:G17}; {(t.Direction == TargetDirection.AtOrBelow ? "Lower" : "Higher")} is better" : "Not configured";
        object?[,] summary = {
            { "SPC PROCESS ANALYSIS", c.ProcessName }, { "Source", sourceDescription }, { "Measurement unit", c.MeasurementUnit },
            { "Observations", result.Observations.Count }, { "Baseline mode", c.BaselineMode.ToString() },
            { "Baseline period", $"#{b.Start}–#{b.End}: {result.Labels[b.Start - 1]} → {result.Labels[b.End - 1]}" },
            { "Baseline count", b.End - b.Start + 1 }, { "Baseline mean", b.Mean }, { "Average moving range", b.AverageMovingRange },
            { "Individuals UCL", b.IndividualsUpper }, { "Individuals LCL", b.IndividualsLower }, { "Moving range UCL", b.MovingRangeUpper },
            { "Moving range LCL", 0 }, { "Business target", target }, { "Latest observation", result.Observations[^1] },
            { "Status", result.Status }, { "Signals", result.Signals.Count },
            { "Baseline investigation", result.BaselineHasSignals ? "Required" : "No signals detected; reliability is not guaranteed" },
            { "Selected rules", $"Beyond limits: {c.Rules.BeyondLimits}; eight-point run: {c.Rules.SustainedRun}; two of three: {c.Rules.TwoOfThree}; four of five: {c.Rules.FourOfFive}; six-point trend: {c.Rules.SustainedTrend}" },
            { "Target attainment", c.Target == null ? "Not configured" : $"{result.TargetMetCount}/{result.Observations.Count} ({result.TargetMetPercent:0.##}%)" },
            { "Interpretation", string.Join(" ", result.Insights()) },
            { "Chart conventions", "Fixed limits. Red rings mark qualifying signal observations; green diamonds mark baseline start/end. Blank first MR is intentional." }
        };
        // Write text as text (including user labels beginning with =); numerical data stay numerical.
        sheet.Range["A1:B22"].NumberFormat = "@"; sheet.Range["A1:B22"].Value2 = summary;
        object?[,] data = ObservationTable(result); int last = HeaderRow + result.Observations.Count;
        sheet.Range[$"A{HeaderRow}:V{last}"].NumberFormat = "General";
        foreach (string column in new[] { "B", "E", "M", "N" }) sheet.Range[$"{column}{HeaderRow}:{column}{last}"].NumberFormat = "@";
        sheet.Range[$"A{HeaderRow}:V{last}"].Value2 = data;
        sheet.Range[$"A{HeaderRow}:V{HeaderRow}"].Font.Bold = true;
        sheet.Range["A1:A22"].Font.Bold = true; sheet.Range["A:V"].ColumnWidth = 18; sheet.Range["B:B"].ColumnWidth = 27;
        sheet.Range["N:N"].ColumnWidth = 65;
        AddChart(sheet, result, false, last, 20); AddChart(sheet, result, true, last, 365);
        sheet.Activate();
    }
    private static void AddChart(dynamic sheet, SpcResult result, bool mr, int last, double top)
    {
        dynamic chart = sheet.ChartObjects().Add(650, top, 900, 325).Chart;
        chart.ChartType = 65; // xlLineMarkers, category positions preserve observation alignment.
        chart.HasTitle = true; chart.ChartTitle.Text = mr ? "Moving Range — fixed baseline" : "Individuals — fixed baseline";
        chart.DisplayBlanksAs = 1; // xlNotPlotted: no invented MR at observation 1.
        while ((int)chart.SeriesCollection().Count > 0) chart.SeriesCollection(1).Delete();
        void Series(string caption, string column, int color, int dash = 1, bool markersOnly = false)
        {
            dynamic series = chart.SeriesCollection().NewSeries(); series.Name = caption;
            series.XValues = sheet.Range[$"B{HeaderRow + 1}:B{last}"];
            series.Values = sheet.Range[$"{column}{HeaderRow + 1}:{column}{last}"];
            series.Format.Line.ForeColor.RGB = color; series.Format.Line.DashStyle = dash;
            series.MarkerStyle = markersOnly ? (caption.StartsWith("Baseline") ? 2 : 8) : -4142; // green diamonds / red circles
            if (markersOnly) { series.Format.Line.Visible = 0; series.MarkerSize = 7; series.MarkerForegroundColor = color; series.MarkerBackgroundColorIndex = -4142; }
        }
        int blue = 0xA07030, red = 0x3030B0, orange = 0x008CFF, green = 0x508030;
        Series(mr ? "Moving range" : "Observation", mr ? "D" : "C", 0x705020);
        Series("Center", mr ? "J" : "G", blue); Series("UCL", mr ? "I" : "F", red, 4); Series("LCL", mr ? "K" : "H", red, 4);
        if (!mr && result.Configuration.Target != null) Series("Business target", "L", orange, 5);
        Series("Signal observation", mr ? "P" : "O", red, markersOnly: true);
        if (!mr) { Series("Baseline start", "Q", green, markersOnly: true); Series("Baseline end", "R", green, markersOnly: true); }
        chart.HasLegend = true;
    }
}
