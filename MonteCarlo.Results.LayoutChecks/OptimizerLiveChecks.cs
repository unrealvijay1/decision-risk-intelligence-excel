using System.Diagnostics;
using System.Runtime.InteropServices;
using MonteCarlo.Excel;
using MonteCarlo.Core.Tests;
using ExcelDna.Integration;

internal static class OptimizerLiveChecks
{
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    public static void Run()
    {
        int[] processes = []; Exception? failure = null;
        // End the fixture's COM apartment before asserting process exit. The host STA
        // survives for the layout harness and can otherwise retain dynamic COM proxies.
        var worker = new Thread(() => { try { processes = RunPrivate(); } catch (Exception ex) { failure = ex; } });
        worker.SetApartmentState(ApartmentState.STA); worker.Start(); worker.Join();
        if (failure != null) throw new InvalidOperationException("Private Excel optimizer check failed.", failure);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        foreach (int pid in processes) { try { using var process = Process.GetProcessById(pid); if (!process.WaitForExit(10000)) throw new Exception($"Owned Excel process {pid} did not exit."); } catch (ArgumentException) { } }
        Console.WriteLine("LIVE OPTIMIZER PASS: all owned Excel processes exited; user instances untouched");
    }
    private static int[] RunPrivate()
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet(); var processes = new HashSet<int>(); var owned = new List<object>();
        dynamic Own(object value) { owned.Add(value); return value; }
        object NewExcel()
        {
            dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out uint pid);
            if (existing.Contains((int)pid)) { Marshal.ReleaseComObject(app); throw new Exception("Refusing to use existing Excel."); }
            processes.Add((int)pid); Console.WriteLine($"Private Excel PID {pid}"); return app;
        }
        string directory = Path.Combine(Path.GetTempPath(), "MonteCarloOptimizerLive-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        dynamic? application = null, book = null;
        try
        {
            application = Own(NewExcel()); application.Visible = false; application.DisplayAlerts = false; application.EnableEvents = false; application.AutomationSecurity = 3;
            dynamic books = Own(application.Workbooks); book = Own(books.Add()); dynamic sheets = Own(book.Worksheets); dynamic model = Own(sheets[1]); model.Name = "Model";
            dynamic noise = Own(model.Range["A1"]), output = Own(model.Range["B1"]), decision = Own(model.Range["C1"]);
            noise.Value2 = .5; decision.Formula2 = "=25*2"; decision.NumberFormat = "0.000"; output.Formula2 = "=2*C1+A1"; application.Calculate();
            var input = OptimizationTests.Assumption(); var forecast = OptimizationTests.Forecast(); forecast.CellLink = "";
            ExcelDnaUtil.Application = application; SimulationModel.Clear(); SimulationModel.Assumptions.Add(input); SimulationModel.Forecasts.Add(forecast); WorkbookPersistence.SaveModel();
            var problem = OptimizationTests.Problem() with { Objective = new(ScenarioSimulationService.Identity(forecast)), Variables = [OptimizationTests.Problem().Variables[0] with { Step = 10 }], Settings = new(20, 30, 456) };
            problem = WorkbookPersistence.SaveOptimizer((object)book, problem, [forecast]); book.Saved = true;
            OptimizationResult? successful = null; double? replay = null;
            foreach (string mode in new[] { "success", "replay", "cancel", "failure" })
            {
                using var cts = new CancellationTokenSource(); string path; int opens = 0;
                using (var copy = ScenarioExcelSandbox.Prepare((object)book))
                {
                    path = copy.Path;
                    var outputs = mode == "failure" ? new[] { ScenarioSimulationService.Clone(forecast) } : [forecast]; if (mode == "failure") outputs[0].SheetName = "Missing";
                    try
                    {
                        var result = OptimizationService.Run(problem, [input], outputs, () => { opens++; return ScenarioExcelSandbox.Open(path, new(false, 100, .001), NewExcel); }, cts.Token,
                            p => { if (mode == "cancel") cts.Cancel(); });
                        if (mode is "cancel" or "failure") throw new Exception("Expected cancellation/failure.");
                        if (result.Best.Values[0] != 100 || !result.Best.Feasible) throw new Exception("Live optimum incorrect.");
                        var trialResults = result.Best.Run.ForecastResults[0];
                        if (trialResults.Values.Distinct().Count() <= 1 || trialResults.StandardDeviation <= 0 ||
                            trialResults.Values.Any(value => value < 200 || value > 201))
                            throw new Exception("Optimizer did not propagate Uniform(0,1) assumption samples through the Excel formula.");
                        if (mode == "success") Console.WriteLine($"LIVE SAMPLING: X=100, Y=2X+Uniform(0,1), {trialResults.Trials} trials; min={trialResults.Minimum:G10}, max={trialResults.Maximum:G10}, SD={trialResults.StandardDeviation:G10}; source assumption remains 0.5");
                        if (mode == "success") { successful = result; replay = result.Best.Objective; } else if (result.Best.Objective != replay) throw new Exception("Live seed replay differs.");
                    }
                    catch (OperationCanceledException) when (mode == "cancel") { }
                    catch (Exception ex) when (mode == "failure" && ex is SimulationCellAccessException or SimulationRunException) { }
                    if (opens != 1) throw new Exception("Expected one sandbox per optimization run.");
                }
                if (File.Exists(path) || Directory.Exists(Path.GetDirectoryName(path))) throw new Exception("Optimizer temporary file remains.");
                if ((string)decision.Formula2 != "=25*2" || (double)decision.Value2 != 50 || (double)noise.Value2 != .5 || (string)decision.NumberFormat != "0.000" || !(bool)book.Saved)
                    throw new Exception("Optimizer altered source values/formulas/format/Saved state.");
                Console.WriteLine($"LIVE OPTIMIZER PASS {mode}: one sandbox, source preserved, copy removed");
            }
            // Production COM adapter is linked into the test assembly as an internal class.
            var adapterType = typeof(OptimizationTests).Assembly.GetType("MonteCarlo.Excel.ExcelSimulationWorkbook")!;
            var adapter = (ISimulationWorkbook)Activator.CreateInstance(adapterType, new object?[] { application, book, null })!;
            foreach (string mode in new[] { "success", "cancel", "failure" })
            {
                using var cancellation = new CancellationTokenSource(); int frames = 0;
                try
                {
                    var result = OptimizationService.Run(problem, [input], [forecast], () => new LiveSource(adapter), cancellation.Token,
                        liveProgress: (completed, samples) =>
                        {
                            frames++;
                            if (!adapter.ScreenUpdating || adapter.EnableEvents) throw new Exception("Live Excel settings incorrect.");
                            double currentDecision = (double)decision.Value2, currentNoise = (double)noise.Value2, currentOutput = (double)output.Value2;
                            if (currentNoise < 0 || currentNoise > 1 || Math.Abs(currentOutput - (2 * currentDecision + currentNoise)) > 1e-10)
                                throw new Exception("Live source values did not recalculate.");
                            if (completed == 1 && samples.Values.Single()[0] != currentOutput) throw new Exception("Preview differs from worksheet.");
                            if (mode == "cancel") { cancellation.Cancel(); cancellation.Token.ThrowIfCancellationRequested(); }
                            if (mode == "failure") throw new InvalidOperationException("Synthetic chart failure");
                        });
                    if (mode != "success") throw new Exception("Expected live interruption.");
                    if (result.Best.Values[0] != 100) throw new Exception("Live source optimum incorrect.");
                }
                catch (OperationCanceledException) when (mode == "cancel") { }
                catch (SimulationRunException) when (mode == "failure") { }
                if (frames == 0 || (string)decision.Formula2 != "=25*2" || (double)noise.Value2 != .5 || (double)output.Value2 != 100.5 || (string)decision.NumberFormat != "0.000")
                    throw new Exception("Live source restoration failed.");
                Console.WriteLine($"LIVE SOURCE PASS {mode}: actual trial cells changed, forecasts recalculated, originals restored");
            }
            OptimizationService.Apply(successful!, adapter, [input], [forecast]); if ((double)decision.Value2 != 100) throw new Exception("Explicit Apply failed.");
            OptimizerExporter.Export((object)book, successful!);
            dynamic report = Own(book.ActiveSheet); if (!((string)report.Name).StartsWith("Optimizer_")) throw new Exception("Optimizer report missing.");
            string file = Path.Combine(directory, "Optimizer.xlsx"); book.SaveAs(file, 51); book.Close(false);
            book = Own(books.Open(file)); var reloaded = WorkbookPersistence.LoadOptimizer((object)book)!;
            if (reloaded.Settings != problem.Settings || reloaded.Variables[0].CellLink != problem.Variables[0].CellLink) throw new Exception("Optimizer disk persistence failed.");
            dynamic reopenedModel = Own(book.Worksheets["Model"]); dynamic reopenedDecision = Own(reopenedModel.Range["C1"]);
            if ((double)reopenedDecision.Value2 != 100) throw new Exception("Applied values did not persist.");
            reopenedModel.Name = "Renamed"; var renamed = WorkbookPersistence.LoadOptimizer((object)book)!;
            if (renamed.Variables[0].Sheet != "Renamed") throw new Exception("Tracked rename failed.");
            dynamic deleteRange = Own(reopenedModel.Range["C:C"]); deleteRange.Delete();
            if (WorkbookPersistence.LoadOptimizer((object)book)!.Variables[0].Cell != "#REF!") throw new Exception("Deleted tracked reference silently rebound.");
            Console.WriteLine("LIVE OPTIMIZER PASS: explicit apply, report export, disk save/reopen and tracked rename/delete");
        }
        finally
        {
            SimulationModel.Clear(); ExcelDnaUtil.Application = null!;
            try { if (book != null) book.Close(false); }
            finally
            {
                try { if (application != null) application.Quit(); }
                finally
                {
                    // The fixture exclusively owns these RCWs. Persistence revisits some objects,
                    // increasing their wrapper count; fully release fixture wrappers after Quit.
                    foreach (object value in owned.Distinct(ReferenceEqualityComparer.Instance).Reverse())
                        if (Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
                    owned.Clear();
                }
            }
            // Unique directory is created and owned solely by this harness.
            Directory.Delete(directory, true);
        }
        return processes.ToArray();
    }
    private sealed class LiveSource(ISimulationWorkbook workbook) : IScenarioSandbox
    {
        public ISimulationWorkbook Workbook => workbook;
        public void Dispose() { }
    }
}

