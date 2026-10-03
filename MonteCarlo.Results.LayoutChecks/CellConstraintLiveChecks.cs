using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using MonteCarlo.Core.Tests;
using MonteCarlo.Excel;

internal static class CellConstraintLiveChecks
{
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    public static void Run()
    {
        int pid = 0; Exception? failure = null;
        var worker = new Thread(() => { try { pid = RunPrivate(); } catch (Exception ex) { failure = ex; } });
        worker.SetApartmentState(ApartmentState.STA); worker.Start(); worker.Join();
        if (failure != null) throw new InvalidOperationException("Private Excel cell-constraint checks failed.", failure);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        try { using var process = Process.GetProcessById(pid); if (!process.WaitForExit(10000)) throw new Exception("Owned constraint-test Excel did not exit."); } catch (ArgumentException) { }
        Console.WriteLine("LIVE CELL CONSTRAINT PASS: owned Excel exited; user instances untouched");
    }
    private static int RunPrivate()
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet(); var owned = new List<object>();
        dynamic Own(object item) { owned.Add(item); return item; }
        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
        GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out uint id);
        if (existing.Contains((int)id)) { Marshal.ReleaseComObject(app); throw new Exception("Refusing to use an existing Excel instance."); }
        Own(app); dynamic? book = null;
        string directory = Path.Combine(Path.GetTempPath(), "MonteCarloCellConstraints-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
        string file = Path.Combine(directory, "ProductMix.xlsx");
        try
        {
            app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false; app.AutomationSecurity = 3;
            dynamic books = Own(app.Workbooks); book = Own(books.Add()); dynamic sheets = Own(book.Worksheets); dynamic model = Own(sheets[1]); model.Name = "Model";
            dynamic noise = Own(model.Range["A1"]), a = Own(model.Range["B2"]), b = Own(model.Range["B3"]), profit = Own(model.Range["B5"]), labor = Own(model.Range["B6"]), material = Own(model.Range["B7"]);
            noise.Value2 = .5; a.Formula2 = "=5*2"; b.Value2 = 10;
            profit.Formula2 = "=B2*500+B3*800"; labor.Formula2 = "=B2*2+B3*4"; material.Formula2 = "=B2*3+B3*2";
            app.Calculate();
            var input = OptimizationTests.Assumption(); var forecast = new ForecastDefinition { Name = "Total Profit", SheetName = "Model", CellAddress = "B5" };
            ExcelDnaUtil.Application = app; SimulationModel.Clear(); SimulationModel.Assumptions.Add(input); SimulationModel.Forecasts.Add(forecast); WorkbookPersistence.SaveModel();
            var problem = CellConstraintTests.ProductMix().Problem with { Objective = new(ScenarioSimulationService.Identity(forecast)) };
            problem = WorkbookPersistence.SaveOptimizer((object)book, problem, [forecast]);
            if (SimulationModel.Forecasts.Count != 1) throw new Exception("Labor/material were registered as Forecasts.");
            // Formula outputs remain locked; only the existing writable model inputs are unlocked.
            noise.Locked = false; a.Locked = false; b.Locked = false; model.Protect();
            var adapterType = typeof(OptimizationTests).Assembly.GetType("MonteCarlo.Excel.ExcelSimulationWorkbook")!;
            ISimulationWorkbook Adapter(object workbook) => (ISimulationWorkbook)Activator.CreateInstance(adapterType, new object?[] { app, workbook, null })!;
            var adapter = Adapter((object)book); OptimizationResult? optimum = null;
            foreach (string mode in new[] { "success", "cancel", "failure" })
            {
                using var cancellation = new CancellationTokenSource(); int frames = 0;
                try
                {
                    var result = OptimizationService.Run(problem, [input], [forecast], () => new Source(adapter), cancellation.Token,
                        liveProgress: (_, _) =>
                        {
                            frames++; double av = (double)a.Value2, bv = (double)b.Value2;
                            if ((double)labor.Value2 != av * 2 + bv * 4 || (double)material.Value2 != av * 3 + bv * 2) throw new Exception("Excel dependencies not recalculated.");
                            if (mode == "cancel") { cancellation.Cancel(); cancellation.Token.ThrowIfCancellationRequested(); }
                            if (mode == "failure") throw new InvalidOperationException("Synthetic chart failure");
                        });
                    if (mode != "success") throw new Exception("Expected interruption.");
                    optimum = result;
                    if (!result.Best.Feasible || result.Best.Values[0] != 80 || result.Best.Values[1] != 60 || result.Best.Objective != 88000 || result.Best.Constraints[0].Actual != 400 || result.Best.Constraints[1].Actual != 360)
                        throw new Exception("Product-mix optimum incorrect.");
                    if (!result.History.Any(h => !h.Feasible)) throw new Exception("Infeasible candidates were not evaluated correctly.");
                }
                catch (OperationCanceledException) when (mode == "cancel") { }
                catch (SimulationRunException) when (mode == "failure") { }
                if (frames == 0 || (string)a.Formula2 != "=5*2" || (double)a.Value2 != 10 || (double)b.Value2 != 10 || (double)noise.Value2 != .5 ||
                    (string)profit.Formula2 != "=B2*500+B3*800" || (string)labor.Formula2 != "=B2*2+B3*4" || (string)material.Formula2 != "=B2*3+B3*2")
                    throw new Exception("Source/formula restoration failed.");
                Console.WriteLine($"LIVE CELL CONSTRAINT PASS {mode}: protected formulas recalculated, read-only constraints, source restored");
            }
            model.Unprotect();
            dynamic errorCell = Own(model.Range["E1"]); errorCell.Formula2 = "=1/0"; app.Calculate();
            try
            {
                OptimizationService.Run(problem with { Constraints = [CellConstraintTests.Cell("E1", 0)] }, [input], [forecast], () => new Source(adapter));
                throw new Exception("Excel error constraint was accepted.");
            }
            catch (ArgumentException ex) when (ex.Message.Contains("Constraint")) { }
            labor.Formula2 = "=IF(B2=0,1/0,B2*2+B3*4)"; app.Calculate();
            try { OptimizationService.Run(problem, [input], [forecast], () => new Source(adapter)); throw new Exception("Candidate formula error was accepted."); }
            catch (ArgumentException ex) when (ex.Message.Contains("Constraint")) { }
            if ((double)a.Value2 != 10 || (string)a.Formula2 != "=5*2" || (string)labor.Formula2 != "=IF(B2=0,1/0,B2*2+B3*4)") throw new Exception("Dynamic error restoration failed.");
            labor.Formula2 = "=B2*2+B3*4"; app.Calculate();
            OptimizationService.Apply(optimum!, adapter, [input], [forecast]);
            if ((double)labor.Value2 != 400 || (double)material.Value2 != 360 || (double)profit.Value2 != 88000) throw new Exception("Applied solution violates calculated constraints.");
            book.SaveAs(file, 51); book.Close(false); book = Own(books.Open(file));
            var loaded = WorkbookPersistence.LoadOptimizer((object)book)!;
            if (!loaded.Constraints.SequenceEqual(problem.Constraints)) throw new Exception("Disk save/reopen lost constraint fields.");
            dynamic reopened = Own(book.Worksheets["Model"]); dynamic reopenedProfit = Own(reopened.Range["B5"]);
            if ((string)reopenedProfit.Formula2 != "=B2*500+B3*800" || (double)reopenedProfit.Value2 != 88000) throw new Exception("Profit formula/save failed.");
            reopened.Unprotect(); reopened.Name = "Renamed"; dynamic firstRow = Own(reopened.Range["1:1"]); firstRow.Insert();
            loaded = WorkbookPersistence.LoadOptimizer((object)book)!;
            if (loaded.Constraints[0].Sheet != "Renamed" || loaded.Constraints[0].Cell.Replace("$", "") != "B7" || loaded.Variables[0].Cell.Replace("$", "") != "B3") throw new Exception("Constraint tracking did not follow rename/row insertion.");
            dynamic deletedRow = Own(reopened.Range["7:7"]); deletedRow.Delete();
            loaded = WorkbookPersistence.LoadOptimizer((object)book)!;
            if (loaded.Constraints[0].Cell != "#REF!") throw new Exception("Deleted constraint silently rebound.");
            try { WorkbookPersistence.SaveOptimizer((object)book, loaded); throw new Exception("Deleted reference saved."); } catch (ArgumentException) { }
            Console.WriteLine("LIVE PRODUCT MIX PASS: Qty A=80, Qty B=60, profit=88,000; labor=400/material=360; formulas preserved; disk reopen and tracked rename/insert/delete passed");
            book.Close(false); book = Own(books.Add()); dynamic deterministicSheets = Own(book.Worksheets);
            dynamic priceModel = Own(deterministicSheets[1]); priceModel.Name = "Model";
            dynamic price = Own(priceModel.Range["A1"]), demand = Own(priceModel.Range["A2"]), revenue = Own(priceModel.Range["A4"]);
            price.Formula2 = "=25*2"; demand.Formula2 = "=1000-5*A1"; revenue.Formula2 = "=A1*A2"; app.Calculate();
            var revenueDefinition = new ForecastDefinition { Name = "Revenue", SheetName = "Model", CellAddress = "A4" };
            SimulationModel.Clear(); SimulationModel.Forecasts.Add(revenueDefinition); WorkbookPersistence.SaveModel();
            var deterministicProblem = new OptimizationProblem(new(ScenarioSimulationService.Identity(revenueDefinition)),
                [new("price", "Price", "Model", "A1", 10, 180, 1)], [], new(200, 10000, 42));
            deterministicProblem = WorkbookPersistence.SaveOptimizer((object)book, deterministicProblem, [revenueDefinition]);
            var priceAdapter = Adapter((object)book); OptimizationResult? priceOptimum = null;
            foreach (string mode in new[] { "success", "cancel", "failure" })
            {
                using var cancellation = new CancellationTokenSource(); int frames = 0;
                try
                {
                    var result = OptimizationService.Run(deterministicProblem, [], [revenueDefinition], () => new Source(priceAdapter), cancellation.Token,
                        liveProgress: (n, samples) =>
                        {
                            frames++;
                            if (n != 1 || samples.Values.Single().Length != 1 || (double)revenue.Value2 != (double)price.Value2 * (double)demand.Value2)
                                throw new Exception("Deterministic model was not recalculated exactly once per candidate.");
                            if (mode == "cancel") cancellation.Cancel();
                            if (mode == "failure") throw new InvalidOperationException("Synthetic deterministic preview failure");
                        });
                    if (mode != "success") throw new Exception("Expected deterministic interruption.");
                    priceOptimum = result;
                    if (result.EvaluationMode != OptimizationEvaluationMode.Deterministic || result.Best.Values[0] != 100 || result.Best.Objective != 50000 || result.Best.Run.Trials != 1 || frames != result.Evaluations)
                        throw new Exception("Price/revenue deterministic optimum incorrect.");
                }
                catch (OperationCanceledException) when (mode == "cancel") { }
                catch (InvalidOperationException ex) when (mode == "failure" && ex.Message.Contains("preview failure")) { }
                if (frames == 0 || (string)price.Formula2 != "=25*2" || (double)price.Value2 != 50 || (string)revenue.Formula2 != "=A1*A2")
                    throw new Exception("Deterministic input/formula restoration failed.");
            }
            string priceFile = Path.Combine(directory, "DeterministicPrice.xlsx");
            book.SaveAs(priceFile, 51); book.Close(false); book = Own(books.Open(priceFile));
            var priceLoaded = WorkbookPersistence.LoadOptimizer((object)book)!;
            WorkbookPersistence.LoadModelForSimulation(false);
            if (SimulationModel.Assumptions.Count != 0 || SimulationModel.Forecasts.Count != 1) throw new Exception("Deterministic disk model lost its mode.");
            var loadedAdapter = Adapter((object)book);
            priceOptimum = OptimizationService.Run(priceLoaded, [], SimulationModel.Forecasts.ToArray(), () => new Source(loadedAdapter));
            OptimizationService.Apply(priceOptimum, loadedAdapter, [], SimulationModel.Forecasts.ToArray());
            dynamic loadedModel = Own(book.Worksheets["Model"]), loadedPrice = Own(loadedModel.Range["A1"]), loadedDemand = Own(loadedModel.Range["A2"]), loadedRevenue = Own(loadedModel.Range["A4"]);
            if ((double)loadedPrice.Value2 != 100 || (double)loadedDemand.Value2 != 500 || (double)loadedRevenue.Value2 != 50000) throw new Exception("Deterministic Apply failed.");
            Console.WriteLine("LIVE DETERMINISTIC PASS: zero assumptions; Price=100/Demand=500/Revenue=50,000; live updates, success/cancel/failure restoration, disk reopen and Apply");
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
