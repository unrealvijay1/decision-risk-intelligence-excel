using System.Diagnostics;
using System.Runtime.InteropServices;
using MonteCarlo.Core;
using MonteCarlo.Excel;

// Opt-in integration check. Creates its own synthetic source workbook; never attaches to the user's Excel.
internal static class ScenarioLiveChecks
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    public static void Run()
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        var owned = new List<object>();
        var processes = new HashSet<int>();
        dynamic Own(object value) { owned.Add(value); return value; }
        object NewExcel()
        {
            dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out uint pid);
            if (existing.Contains((int)pid)) { Marshal.ReleaseComObject(app); throw new Exception("Refusing to use an existing Excel process."); }
            processes.Add((int)pid);
            return app;
        }
        dynamic? application = null, book = null;
        try
        {
            application = Own(NewExcel()); application.Visible = false; application.DisplayAlerts = false; application.EnableEvents = false;
            application.AutomationSecurity = 3;
            dynamic books = Own(application.Workbooks); book = Own(books.Add());
            dynamic sheets = Own(book.Worksheets); dynamic model = Own(sheets[1]); model.Name = "Model";
            dynamic input = Own(model.Range["A1"]), multiplier = Own(model.Range["A2"]), output = Own(model.Range["B1"]);
            input.Formula2 = "=50*2"; input.NumberFormat = "0.000";
            dynamic fill = Own(input.Interior); fill.Color = 65535;
            multiplier.Value2 = 100;
            dynamic names = Own(book.Names); Own(names.Add(Name: "RevenueInput", RefersTo: "='Model'!$A$1"));
            output.Formula2 = "=RevenueInput*A2/100";
            dynamic calculations = Own(sheets.Add()); calculations.Name = "Calculations";
            dynamic forecast = Own(calculations.Range["B2"]); forecast.Formula2 = "='Model'!B1+10";
            application.Calculate(); book.Saved = true;
            string originalFormula = (string)input.Formula2, originalOutput = (string)forecast.Formula2;
            var inputs = new[] { new AssumptionDefinition { Name = "Revenue", CellLink = "a", SheetName = "Model", CellAddress = "A1", Distribution = DistributionType.Normal, Parameter1 = 100, Parameter2 = 15 },
                new AssumptionDefinition { Name = "Multiplier", CellLink = "b", SheetName = "Model", CellAddress = "A2", Distribution = DistributionType.Normal, Parameter1 = 100, Parameter2 = 15 } };
            var outputs = new[] { new ForecastDefinition { Name = "Profit", SheetName = "Calculations", CellAddress = "B2", TargetSettings = new(130, TargetDirection.AtOrBelow, null) } };
            var plans = new[] { new ScenarioDefinition("current", "Current Plan", new[] { new ScenarioAdjustment("a", ScenarioAdjustmentType.PercentageChange, 10) }),
                new ScenarioDefinition("1", "Growth", new[] { new ScenarioAdjustment("a", ScenarioAdjustmentType.PercentageChange, 15) }),
                new ScenarioDefinition("2", "Zero", new[] { new ScenarioAdjustment("a", ScenarioAdjustmentType.PercentageChange, -100) }) };
            foreach (string mode in new[] { "success", "cancel", "failure" })
            {
                using var token = new CancellationTokenSource();
                int opens = 0; string path;
                var stopwatch = Stopwatch.StartNew();
                using (var copy = ScenarioExcelSandbox.Prepare((object)book))
                {
                    path = copy.Path;
                    IScenarioSandbox Create()
                    {
                        opens++;
                        return ScenarioExcelSandbox.Open(copy.Path, new(false, 100, .001), NewExcel);
                    }
                    var runOutputs = mode == "failure" ? new[] { new ForecastDefinition { SheetName = "Missing", CellAddress = "B2" } } : outputs;
                    try
                    {
                        var analysis = new ScenarioAnalysisDefinition(plans, "current", ForecastQuestions: new[] { new ScenarioForecastQuestion(ScenarioSimulationService.Identity(runOutputs[0]), TargetDirection.AtOrBelow) });
                        var result = ScenarioSimulationService.Run(new(50, SimulationSeedMode.Fixed, 42), inputs, runOutputs, analysis, Create,
                            token.Token, message => { if (mode == "cancel" && message.Contains("Growth")) token.Cancel(); }, correlations: new[] { new AssumptionCorrelation("a", "b", .8) });
                        if (mode != "success") throw new Exception("Expected run to fail/cancel.");
                        if (result.Rows.Count != 3 || result.Rows[0].Scenario != "Current Plan" || !result.Rows[0].IsBaseline ||
                            Math.Abs(result.Rows[1].Mean - ((result.Rows[0].Mean - 10) / 1.1 * 1.15 + 10)) > 1e-8 || result.Rows[2].Mean != 10)
                            throw new Exception("Named/cross-sheet formula calculation or scaling failed.");
                    }
                    catch (OperationCanceledException) when (mode == "cancel") { }
                    catch (InvalidOperationException) when (mode == "failure") { }
                    if (opens != 1) throw new Exception("More than one disposable Excel workbook opened.");
                }
                if (File.Exists(path) || Directory.Exists(Path.GetDirectoryName(path))) throw new Exception("Temporary files left behind.");
                if ((string)input.Formula2 != originalFormula || (double)input.Value2 != 100 || (string)forecast.Formula2 != originalOutput ||
                    (double)forecast.Value2 != 110 || (double)multiplier.Value2 != 100 || (string)input.NumberFormat != "0.000" || (double)fill.Color != 65535 || !(bool)book.Saved)
                    throw new Exception("Source workbook was modified.");
                Console.WriteLine($"LIVE SCENARIO PASS {mode}: one copy, source/formulas/format preserved, cleanup; {stopwatch.Elapsed.TotalSeconds:F2}s");
            }
        }
        finally
        {
            try { if (book != null) book.Close(SaveChanges: false); }
            finally
            {
                try { if (application != null) application.Quit(); }
                finally { for (int i = owned.Count - 1; i >= 0; i--) if (Marshal.IsComObject(owned[i])) Marshal.ReleaseComObject(owned[i]); }
            }
        }
        foreach (int pid in processes)
        {
            try { using var process = Process.GetProcessById(pid); if (!process.WaitForExit(5000)) throw new Exception($"Owned Excel process {pid} did not exit."); }
            catch (ArgumentException) { } // Already exited.
        }
        Console.WriteLine("LIVE SCENARIO PASS: all owned Excel processes exited; existing Excel processes were not used.");
    }
}
