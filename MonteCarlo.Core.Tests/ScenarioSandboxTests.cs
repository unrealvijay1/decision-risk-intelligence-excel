using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

public class ScenarioSandboxTests
{
    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void TemporaryFileDeletedOnCompletionOrSaveFailure(bool failure)
    {
        string? path = null;
        if (failure)
            Assert.Throws<IOException>(() => ScenarioTemporaryFile.Create("Model.xlsx", target => { path = target; File.WriteAllText(target, "copy"); throw new IOException("save failed"); }));
        else
        {
            var file = ScenarioTemporaryFile.Create("Model.xlsx", target => { path = target; File.WriteAllText(target, "copy"); });
            Assert.True(File.Exists(path)); file.Dispose(); file.Dispose();
        }
        Assert.NotNull(path); Assert.False(File.Exists(path)); Assert.False(Directory.Exists(Path.GetDirectoryName(path)));
    }
    [Theory]
    [InlineData("success")] [InlineData("open")] [InlineData("setup")] [InlineData("calculate")]
    public void ProductionSandboxClosesOwnedExcelAndLeavesSourceUntouched(string failure)
    {
        var source = new SandboxSource(); var original = source.Data.Worksheets[1].Range["A1"].Value2;
        string path;
        var app = new SandboxApp(source.Data.Reopen()) { Failure = failure };
        using (var copy = ScenarioExcelSandbox.Prepare(source))
        {
            path = copy.Path;
            if (failure == "success")
            {
                var result = ScenarioSimulationService.Run(new(4, SimulationSeedMode.Fixed, 42), [ScenarioSimulationTests.Input()], [ScenarioSimulationTests.Output()],
                    ScenarioSimulationTests.Comparison(ScenarioSimulationTests.Scenario("Growth", ScenarioSimulationTests.Change("a", 15))),
                    () => ScenarioExcelSandbox.Open(copy.Path, new(true, 100, .001), () => app));
                Assert.Equal(2, result.Rows.Count); Assert.Equal(1, app.Workbooks.Opens); Assert.Equal(1, source.Copies);
                Assert.Equal(3, app.AutomationSecurity); Assert.False(app.AskToUpdateLinks);
            }
            else Assert.ThrowsAny<Exception>(() => {
                using var sandbox = ScenarioExcelSandbox.Open(copy.Path, new(true, 100, .001), () => app);
                sandbox.Workbook.Calculate();
            });
            Assert.Equal(1, app.Quits);
            Assert.Equal(failure == "open" ? 0 : 1, app.Workbooks.Book.Closes);
            Assert.Equal(original, source.Data.Worksheets[1].Range["A1"].Value2);
            Assert.Equal(0, source.Closes);
        }
        Assert.False(File.Exists(path));
    }
    [Theory]
    [InlineData("vba")] [InlineData("links")] [InlineData("connections")]
    public void UnsupportedSourceFailsBeforeCopy(string kind)
    {
        var source = new SandboxSource { HasVBProject = kind == "vba", Links = kind == "links" };
        source.Connections.Count = kind == "connections" ? 1 : 0;
        Assert.Throws<InvalidOperationException>(() => ScenarioExcelSandbox.Prepare(source)); Assert.Equal(0, source.Copies);
    }
    [Fact]
    public void CloseFailureStillQuitsExcelAndReportsFailure()
    {
        var app = new SandboxApp(new SandboxSource().Data) { Failure = "close" };
        var sandbox = ScenarioExcelSandbox.Open("dummy", new(false, 100, .001), () => app);
        Assert.Throws<AggregateException>(sandbox.Dispose); Assert.Equal(1, app.Quits); Assert.Equal(2, app.Workbooks.Book.Closes);
    }
    public sealed class SandboxSource
    {
        public PersistenceWorkbook Data { get; } = new();
        public SandboxSource() { Data.Worksheets.Add().Name = "Model"; Data.Worksheets[1].Range["A1"].Value2 = 100d; Data.Worksheets[1].Range["B1"].Value2 = 100d; }
        public string Name => "Model.xlsx"; public int FileFormat => 51;
        public bool HasVBProject { get; set; } public bool Links { get; set; }
        public ConnectionCount Connections { get; } = new(); public int Copies, Closes;
        public object? LinkSources(int type) => Links ? new[] { "external.xlsx" } : null;
        public void SaveCopyAs(string path) { Copies++; File.WriteAllText(path, "test snapshot"); }
    }
    public sealed class ConnectionCount { public int Count { get; set; } }
    public sealed class SandboxApp
    {
        public SandboxApp(PersistenceWorkbook data) { Workbooks = new(this, data); }
        public string Failure = ""; public int Quits;
        public bool Visible { get; set; } public bool DisplayAlerts { get; set; } public bool EnableEvents { get; set; }
        public bool ScreenUpdating { get; set; } public object? StatusBar { get; set; }
        public int AutomationSecurity { get; set; } public bool AskToUpdateLinks { get; set; }
        public int Calculation { get { return -4135; } set { if (Failure == "setup") throw new Exception("setup failed"); } }
        public bool Iteration { get; set; } public int MaxIterations { get; set; } public double MaxChange { get; set; }
        public SandboxBooks Workbooks { get; } public PersistenceFunctions WorksheetFunction { get; } = new();
        public void Quit() => Quits++;
        public void Calculate()
        {
            if (Failure == "calculate") throw new Exception("failed calculation");
            Workbooks.Book.Data.Worksheets[1].Range["B1"].Value2 = Workbooks.Book.Data.Worksheets[1].Range["A1"].Value2;
        }
    }
    public sealed class SandboxBooks(SandboxApp app, PersistenceWorkbook data)
    {
        public SandboxBook Book { get; } = new(app, data); public int Opens;
        public SandboxBook Open(string path, int UpdateLinks, bool ReadOnly, bool IgnoreReadOnlyRecommended, bool AddToMru)
        { Opens++; Assert.Equal(0, UpdateLinks); Assert.False(AddToMru); if (app.Failure == "open") throw new Exception("open failed"); return Book; }
    }
    public sealed class SandboxBook(SandboxApp app, PersistenceWorkbook data)
    {
        public PersistenceWorkbook Data => data; public PersistenceSheets Worksheets => data.Worksheets; public int Closes;
        public object? LinkSources(int type) => null;
        public void Close(bool SaveChanges) { Assert.False(SaveChanges); Closes++; if (app.Failure == "close") throw new Exception("close failed"); }
    }
}
