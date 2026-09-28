using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel.Licensing;

namespace MonteCarlo.Excel;

internal static class SpcAnalysisCommand
{
    public static void Show()
    {
        LicenseService.EnsureAccess();
        dynamic app = ExcelDnaUtil.Application;
        object workbook = app.ActiveWorkbook ?? throw new ArgumentException("Open a workbook containing your observations first.");
        var saved = WorkbookPersistence.LoadSpcAnalysis(workbook);
        object? dataRange = null, labelRange = null;
        bool labelNeedsRepair = false;
        string dataDescription = "Select numerical observations, without headers", labelDescription = "Observation numbers";
        if (saved != null)
        {
            try
            {
                dataRange = SpcDataReader.Resolve(workbook, saved.DataLink); dataDescription = SpcDataReader.Describe(dataRange);
            }
            catch (ArgumentException) { dataDescription = "Saved reference is broken. Reselect ranges and replace baseline."; }
            if (saved.LabelLink != null)
            {
                try { labelRange = SpcDataReader.Resolve(workbook, saved.LabelLink); labelDescription = SpcDataReader.Describe(labelRange); }
                catch (ArgumentException) { labelNeedsRepair = true; labelDescription = "Saved label reference is broken. Reselect or Clear it."; }
            }
        }
        string? Select(bool labels)
        {
            object selected;
            try { selected = app.InputBox(Prompt: labels ? "Select one date/sequence label per observation." : "Select one contiguous row or column of observations, excluding headers.", Title: "SPC range selection", Type: 8); }
            catch (System.Runtime.InteropServices.COMException) { return null; }
            if (selected is bool) return null;
            SpcDataReader.ValidateRange(workbook, selected);
            if (labels) { labelRange = selected; labelNeedsRepair = false; } else dataRange = selected;
            return SpcDataReader.Describe(selected);
        }
        SpcConfigurationForm? form = null;
        void Analyze(SpcDraft draft)
        {
            if (dataRange == null) throw new ArgumentException("Select the numerical observation range.");
            if (labelNeedsRepair) throw new ArgumentException("Reselect or explicitly clear the broken date/sequence reference.");
            var outcome = SpcAnalysisService.Analyze(workbook, dataRange, labelRange, draft.Configuration, draft.ReplaceBaseline, saved);
            var result = outcome.Result; saved = outcome.Saved;
            string source = SpcDataReader.Describe(dataRange);
            using var results = new SpcResultsForm(result, () => SpcExporter.Export(workbook, result, source));
            results.ShowDialog(form);
        }
        using (form = new SpcConfigurationForm(saved, dataDescription, labelDescription, Select, () => { labelRange = null; labelNeedsRepair = false; }, Analyze)) form.ShowDialog();
    }
}
