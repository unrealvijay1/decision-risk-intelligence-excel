using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public sealed record SpcAnalysisOutcome(SpcResult Result, SpcSavedAnalysis Saved);

public static class SpcAnalysisService
{
    public static SpcAnalysisOutcome Analyze(object workbook, object dataRange, object? labelRange,
        SpcConfiguration configuration, bool replaceBaseline, SpcSavedAnalysis? saved)
    {
        dynamic data = dataRange;
        int count = SpcDataReader.ValidateRange(workbook, dataRange);
        double[] values = SpcDataReader.ReadValues(workbook, dataRange);
        string[] labels = labelRange == null ? Enumerable.Range(1, count).Select(i => i.ToString()).ToArray() : SpcDataReader.ReadLabels(workbook, labelRange, count);
        SpcBaseline baseline;
        object baselineRange;
        var config = configuration;
        if (replaceBaseline)
        {
            baseline = XmRCalculator.EstablishBaseline(values, config);
            dynamic first = data.Cells[baseline.Start], last = data.Cells[baseline.End];
            baselineRange = data.Worksheet.Range[first, last];
        }
        else
        {
            if (saved == null) throw new ArgumentException("Establish a baseline before analyzing observations.");
            baselineRange = SpcDataReader.Resolve(workbook, saved.BaselineLink);
            var period = SpcDataReader.LocateBaseline(workbook, dataRange, baselineRange);
            if (period.End - period.Start != saved.Baseline.End - saved.Baseline.Start)
                throw new ArgumentException("The tracked baseline size changed. Review the data and explicitly replace the baseline.");
            baseline = saved.Baseline with { Start = period.Start, End = period.End };
            config = config with { BaselineMode = saved.Configuration.BaselineMode };
        }
        config = config with { BaselineStart = baseline.Start, BaselineEnd = baseline.End };
        var result = XmRCalculator.Analyze(values, config, baseline, labels);
        // Names and metadata are written only by explicit actions, after complete validation.
        string dataLink = SpcDataReader.TrackOrReuse(workbook, dataRange, saved?.DataLink);
        string? labelsLink = labelRange == null ? null : SpcDataReader.TrackOrReuse(workbook, labelRange, saved?.LabelLink);
        string baselineLink = replaceBaseline ? SpcDataReader.TrackOrReuse(workbook, baselineRange, saved?.BaselineLink) : saved!.BaselineLink;
        var next = new SpcSavedAnalysis(config, dataLink, labelsLink, baselineLink, baseline);
        WorkbookPersistence.SaveSpcAnalysis(workbook, next);
        return new(result, next);
    }
}
