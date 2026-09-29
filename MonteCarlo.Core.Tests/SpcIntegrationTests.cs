using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

[Collection("Workbook persistence")]
public class SpcIntegrationTests
{
    private static SpcSavedAnalysis Saved => new(new() { ProcessName = "Lead time", Target = new(12, TargetDirection.AtOrBelow) },
        "_MC_SPC_data", "_MC_SPC_labels", "_MC_SPC_baseline", new(1, 20, 10, 2, 4.68, 15.32, 6.536));
    private static PersistenceWorkbook Book()
    {
        var book = new PersistenceWorkbook(); book.Worksheets.Add().Name = "Data";
        ExcelDnaUtil.Application = new PersistenceApplication { ActiveWorkbook = book }; SimulationModel.Clear(); return book;
    }
    [Fact] public void OldWorkbookHasNoSpcConfiguration() => Assert.Null(WorkbookPersistence.LoadSpcAnalysis(Book()));
    [Theory]
    [InlineData(true, false, true, false, true)]
    [InlineData(false, true, false, true, false)]
    [InlineData(false, false, false, false, false)]
    public void AllFiveSelectionsRoundTrip(bool beyond, bool run, bool trend, bool two, bool four)
    {
        var book = Book();
        var saved = Saved with { Configuration = Saved.Configuration with { Rules = new(beyond, run, trend, two, four) } };
        WorkbookPersistence.SaveSpcAnalysis(book, saved);
        Assert.Equal(saved, WorkbookPersistence.LoadSpcAnalysis(book.Reopen()));
    }
    [Fact] public void LegacyThreeRuleSettingsPreserveSelectionsAndEnableNewRules()
    {
        var book = Book();
        var saved = Saved with { Configuration = Saved.Configuration with { Rules = new(false, true, false) } };
        WorkbookPersistence.SaveSpcAnalysis(book, saved);
        var sheet = book.Worksheets["__MonteCarloConfig"];
        var json = System.Text.Json.Nodes.JsonNode.Parse((string)sheet.Cells[2, 28].Value2!)!;
        var rules = json["Configuration"]!["Rules"]!.AsObject();
        rules.Remove("TwoOfThree"); rules.Remove("FourOfFive");
        foreach (string property in new[] { "Sigma", "LowerOneSigma", "UpperOneSigma", "LowerTwoSigma", "UpperTwoSigma" })
            json["Baseline"]!.AsObject().Remove(property);
        sheet.Cells[2, 28].Value2 = json.ToJsonString();
        var restored = WorkbookPersistence.LoadSpcAnalysis(book.Reopen())!;
        Assert.Equal(new SpcRules(false, true, false, true, true), restored.Configuration.Rules);
        Assert.Equal(saved.Baseline, restored.Baseline);
    }
    [Fact] public void ExportPreservesAllRulesAtQualifyingObservationsAndZoneColumns()
    {
        double[] values = Enumerable.Range(0, 20).Select(i => i % 2 == 0 ? 9d : 11d).Concat(Enumerable.Repeat(16d, 8)).ToArray();
        var result = XmRCalculator.Analyze(values, new(), Saved.Baseline);
        var table = SpcExporter.ObservationTable(result);
        Assert.Contains("Rule 1", (string)table[21, 12]!);
        Assert.Contains("Rule 2", (string)table[21, 12]!);
        Assert.Contains("Rule 3", (string)table[21, 12]!);
        Assert.Contains("Rule 4", (string)table[21, 12]!);
        Assert.Contains("qualifying #21, #22", (string)table[21, 13]!);
        Assert.Equal(Saved.Baseline.LowerOneSigma, table[21, 18]);
        Assert.Equal(Saved.Baseline.UpperTwoSigma, table[21, 21]);
        Assert.Equal(16d, table[21, 14]);
    }
    [Fact] public void ConfigurationAndFixedLimitsSurviveSaveReopenAndNormalModelSave()
    {
        var book = Book(); WorkbookPersistence.SaveSpcAnalysis(book, Saved);
        WorkbookPersistence.SaveModel();
        Assert.Equal(Saved, WorkbookPersistence.LoadSpcAnalysis(book.Reopen()));
        Assert.Equal(2, book.Worksheets["__MonteCarloConfig"].Visible);
    }
    [Fact] public void ClearModelPreservesSpcEvenWhenNoOtherBlocksExist()
    {
        var book = Book(); WorkbookPersistence.SaveSpcAnalysis(book, Saved); WorkbookPersistence.ClearSavedModel();
        Assert.Equal(Saved, WorkbookPersistence.LoadSpcAnalysis(book));
    }
    [Fact] public void MalformedSpcSurvivesUnrelatedSaveAndClear()
    {
        var book = Book(); WorkbookPersistence.SaveSpcAnalysis(book, Saved);
        book.Worksheets["__MonteCarloConfig"].Cells[1, 28].Value2 = "FutureSpcV2";
        WorkbookPersistence.SaveModel(); WorkbookPersistence.ClearSavedModel();
        Assert.Equal("FutureSpcV2", book.Worksheets["__MonteCarloConfig"].Cells[1, 28].Value2);
        Assert.Throws<ArgumentException>(() => WorkbookPersistence.LoadSpcAnalysis(book));
    }
    [Fact] public void SpcSavePreservesModelAndOtherMetadata()
    {
        var book = Book(); WorkbookPersistence.SaveModel(); var sheet = book.Worksheets["__MonteCarloConfig"];
        sheet.Cells[1, 17].Value2 = "Settings sentinel"; sheet.Cells[1, 22].Value2 = "Scenario sentinel"; sheet.Cells[1, 25].Value2 = "Correlation sentinel";
        WorkbookPersistence.SaveSpcAnalysis(book, Saved);
        Assert.Equal("Settings sentinel", sheet.Cells[1, 17].Value2); Assert.Equal("Scenario sentinel", sheet.Cells[1, 22].Value2);
        Assert.Equal("Correlation sentinel", sheet.Cells[1, 25].Value2); Assert.Equal("Type", sheet.Cells[1, 1].Value2);
    }
    [Theory] [InlineData(null)] [InlineData("")] [InlineData("12")] [InlineData(true)] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)]
    public void StrictReaderRejectsGapsAndInvalidCells(object? invalid)
    {
        var raw = new object?[20, 1]; for (int i = 0; i < 20; i++) raw[i, 0] = i;
        raw[9, 0] = invalid;
        Assert.Contains("Observation 10", Assert.Throws<ArgumentException>(() => SpcDataReader.ParseValues(raw, 20)).Message);
    }
    [Fact] public void StrictReaderPreservesOrderAndNumericExcelErrorCodes()
    {
        var raw = Array.CreateInstance(typeof(object), [1, 20], [1, 1]);
        for (int i = 1; i <= 20; i++) raw.SetValue((double)(2043 - i), 1, i);
        var values = SpcDataReader.ParseValues(raw, 20); Assert.Equal(2042, values[0]); Assert.Equal(2023, values[^1]);
    }
    [Theory] [InlineData(19)] [InlineData(100001)]
    public void ReaderEnforcesSupportedSize(int count) => Assert.Throws<ArgumentException>(() => SpcDataReader.ParseValues(null, count));
    [Fact] public void ExportRetainsBlankFirstMovingRangeAndTargetIndependence()
    {
        var values = Enumerable.Range(1, 20).Select(i => (double)i).ToArray();
        var result = new SpcResult(new() { Target = new(0, TargetDirection.AtOrBelow) }, values,
            Enumerable.Range(1, 20).Select(i => i.ToString()).ToArray(), values.Select((v, i) => i == 0 ? (double?)null : 1).ToArray(), Saved.Baseline, []);
        var table = SpcExporter.ObservationTable(result);
        Assert.Null(table[1, 3]); Assert.Equal(1d, table[2, 3]); Assert.Equal(0d, table[1, 11]);
        Assert.Equal("No signals detected", result.Status); Assert.Equal(0, result.TargetMetCount); Assert.Equal(0d, result.TargetMetPercent);
        Assert.Equal(15.32, table[20, 5]);
    }
}
