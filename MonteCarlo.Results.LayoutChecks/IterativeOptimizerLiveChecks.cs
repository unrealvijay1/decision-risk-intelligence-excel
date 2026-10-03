using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using MonteCarlo.Core.Tests;
using MonteCarlo.Excel;

internal static class IterativeOptimizerLiveChecks
{
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    public static void Run()
    {
        int pid = 0; Exception? failure = null;
        var worker = new Thread(() => { try { pid = RunPrivate(); } catch (Exception ex) { failure = ex; } });
        worker.SetApartmentState(ApartmentState.STA); worker.Start(); worker.Join();
        if (failure != null) throw new InvalidOperationException("Private iterative optimizer checks failed", failure);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        try { using var process = Process.GetProcessById(pid); if (!process.WaitForExit(10000)) throw new Exception("Owned Excel did not exit"); } catch (ArgumentException) { }
        Console.WriteLine("ITERATIVE LIVE PASS: owned Excel exited; user instances untouched");
    }
    private static int RunPrivate()
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet(); var owned = new List<object>();
        dynamic Own(object value) { owned.Add(value); return value; }
        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
        GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out uint id);
        if (existing.Contains((int)id)) { Marshal.ReleaseComObject(app); throw new Exception("Refusing an existing Excel instance"); }
        Own(app); dynamic? book = null;
        string directory = Path.Combine(Path.GetTempPath(), "MonteCarloIterative-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        try
        {
            app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false; app.AutomationSecurity = 3;
            dynamic books = Own(app.Workbooks); book = Own(books.Add()); dynamic sheets = Own(book.Worksheets), model = Own(sheets[1]); model.Name = "Model";
            app.Calculation = -4135; // Manual calculation in this owned benchmark fixture; production explicitly recalculates each trial.
            var cells = new Dictionary<string, object>();
            foreach (string address in new[] { "A1", "A2", "A3", "B1", "B3", "C1", "C2", "C3", "D1", "D2", "D3", "D4" }) cells[address] = Own(model.Range[address]);
            dynamic Cell(string address) => cells[address];
            Cell("B3").Value2 = 10000; Cell("A1").Value2 = .08; Cell("A2").Value2 = .12; Cell("A3").Value2 = .07;
            Cell("C1").Formula2 = "=0.3"; Cell("C2").Value2 = .4; Cell("C3").Value2 = .3;
            Cell("D1").Formula2 = "=SUM(C1:C3)";
            Cell("D2").Formula2 = "=$B$3*C1*A1"; Cell("D3").Formula2 = "=$B$3*C2*A2"; Cell("D4").Formula2 = "=$B$3*C3*A3";
            Cell("B1").Formula2 = "=SUM(D2:D4)"; app.Calculate();
            var fixture = IterativeOptimizationTests.Portfolio(); var inputs = fixture.Inputs;
            var forecast = fixture.Forecast; forecast.CellLink = ""; forecast.Name = "Yearly Return";
            ExcelDnaUtil.Application = app; SimulationModel.Clear(); SimulationModel.Assumptions.AddRange(inputs); SimulationModel.Forecasts.Add(forecast); WorkbookPersistence.SaveModel();
            var problem = fixture.Problem with { Objective = new(ScenarioSimulationService.Identity(forecast)),
                Constraints = [CellConstraintTests.Cell("D1", 1, OptimizationConstraintOperator.Equal, "Total Allocation")] };
            problem = WorkbookPersistence.SaveOptimizer((object)book, problem, [forecast]);
            var adapterType = typeof(OptimizationTests).Assembly.GetType("MonteCarlo.Excel.ExcelSimulationWorkbook")!;
            ISimulationWorkbook Adapter(object source) => (ISimulationWorkbook)Activator.CreateInstance(adapterType, new object?[] { app, source, null })!;
            var adapter = Adapter((object)book);
            void Restored()
            {
                if ((string)Cell("C1").Formula2 != "=0.3" || (double)Cell("C2").Value2 != .4 || (double)Cell("C3").Value2 != .3 ||
                    (double)Cell("A1").Value2 != .08 || (double)Cell("A2").Value2 != .12 || (double)Cell("A3").Value2 != .07 || (double)Cell("D1").Value2 != 1)
                    throw new Exception("Portfolio source not restored");
            }
            var rows = new List<string> { "Adapter,Candidates,TrialsPerCandidate,MCTrials,Seconds,CandidatesPerSecond,MCTrialsPerSecond,BestMean" };
            foreach (int budget in new[] { 50, 250, 500 })
            {
                var watch = Stopwatch.StartNew();
                var result = OptimizationService.Run(problem with { Settings = problem.Settings with { MaximumEvaluations = budget } }, inputs, [forecast], () => new Source(adapter),
                    progress: p => { if (p.Evaluations % 25 == 0) Console.WriteLine($"BENCHMARK PROGRESS {budget}: {p.Evaluations} candidates, {watch.Elapsed.TotalSeconds:N1}s"); });
                watch.Stop(); Restored();
                if (!result.Best.Values.SequenceEqual(new double[] { 0, 1, 0 }) || result.Evaluations != budget || result.MonteCarloEvaluations != budget || result.RejectedBeforeSimulation != 0)
                    throw new Exception("Uncapped portfolio/search budget incorrect");
                if (Math.Abs(result.Best.Objective - 1200) > 25 || result.Best.Run.ForecastResults[0].StandardDeviation < 100) throw new Exception("Portfolio uncertainty not simulated");
                rows.Add(string.Join(",", new object[] { "Excel COM (manual calculation)", budget, 1000, budget * 1000, watch.Elapsed.TotalSeconds, budget / watch.Elapsed.TotalSeconds, budget * 1000 / watch.Elapsed.TotalSeconds, result.Best.Objective }.Select(x => Convert.ToString(x, CultureInfo.InvariantCulture))));
                Console.WriteLine($"BENCHMARK PASS {budget}x1000: {watch.Elapsed.TotalSeconds:N3}s; {budget / watch.Elapsed.TotalSeconds:N3} candidates/s; {budget * 1000 / watch.Elapsed.TotalSeconds:N1} trials/s; allocations=0/1/0; mean={result.Best.Objective:N3}");
                Directory.CreateDirectory("artifacts"); File.WriteAllLines("artifacts/iterative-optimizer-benchmarks.csv", rows);
            }
            var capped = problem with { Constraints = [.. problem.Constraints, CellConstraintTests.Cell("C2", .6, name: "Shares cap")] };
            var cappedResult = OptimizationService.Run(capped, inputs, [forecast], () => new Source(adapter),
                progress: p => { if (p.Evaluations % 50 == 0) Console.WriteLine($"CAPPED PORTFOLIO PROGRESS {p.Evaluations}/250"); });
            Restored();
            if (!cappedResult.Best.Values.SequenceEqual(new[] { .4, .6, 0 }) || Math.Abs(cappedResult.Best.Objective - 1040) > 25) throw new Exception("Capped portfolio optimum incorrect");
            foreach (string mode in new[] { "cancel", "failure" })
            {
                using var cancel = new CancellationTokenSource();
                try
                {
                    OptimizationService.Run(capped, inputs, [forecast], () => new Source(adapter), cancel.Token, liveProgress: (_, _) =>
                    { if (mode == "cancel") { cancel.Cancel(); cancel.Token.ThrowIfCancellationRequested(); } throw new InvalidOperationException("Synthetic preview failure"); });
                    throw new Exception("Expected interruption");
                }
                catch (OperationCanceledException) when (mode == "cancel") { }
                catch (SimulationRunException) when (mode == "failure") { }
                Restored();
            }
            var impossible = capped with { Constraints = [CellConstraintTests.Cell("B3", 1)] };
            var rejected = OptimizationService.Run(impossible with { Settings = impossible.Settings with { MaximumEvaluations = 5 } }, inputs, [forecast], () => new Source(adapter));
            if (rejected.BestFeasible != null || rejected.MonteCarloEvaluations != 0 || rejected.RejectedBeforeSimulation != 5) throw new Exception("Deterministic rejection spent simulation trials"); Restored();
            var saved = WorkbookPersistence.SaveOptimizer((object)book, capped, [forecast]);
            string file = Path.Combine(directory, "Portfolio.xlsx"); book.SaveAs(file, 51); book.Close(false); book = Own(books.Open(file));
            var loaded = WorkbookPersistence.LoadOptimizer((object)book)!;
            if (loaded.Settings != saved.Settings || loaded.Constraints.Count != 2) throw new Exception("Saved portfolio optimizer not retained");
            var loadedAdapter = Adapter((object)book);
            OptimizationService.Apply(cappedResult with { Problem = loaded }, loadedAdapter, inputs, [forecast]);
            dynamic reopenedModel = Own(book.Worksheets["Model"]), sum = Own(reopenedModel.Range["D1"]), shares = Own(reopenedModel.Range["C2"]);
            if ((double)sum.Value2 != 1 || (double)shares.Value2 != .6) throw new Exception("Portfolio Apply failed");
            Console.WriteLine($"PORTFOLIO LIVE PASS: worksheet SUM equality and cell cap; 0/1/0 mean≈{1200}; .4/.6/0 mean={cappedResult.Best.Objective:N3}; cancellation/failure/rejection restore; disk reopen and Apply");
            return (int)id;
        }
        finally
        {
            SimulationModel.Clear(); ExcelDnaUtil.Application = null!;
            try { book?.Close(false); } catch { }
            try { app.Quit(); } catch { }
            foreach (var item in owned.Distinct(ReferenceEqualityComparer.Instance).Reverse()) try { if (Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item); } catch { }
            if (Path.GetFullPath(directory).StartsWith(Path.GetFullPath(Path.GetTempPath()), StringComparison.OrdinalIgnoreCase)) Directory.Delete(directory, true);
        }
    }
    private sealed class Source(ISimulationWorkbook workbook) : IScenarioSandbox { public ISimulationWorkbook Workbook => workbook; public void Dispose() { } }
}
