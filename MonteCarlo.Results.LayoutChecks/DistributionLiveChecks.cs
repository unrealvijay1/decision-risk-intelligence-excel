using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel;

internal static class DistributionLiveChecks
{
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    public static void Run()
    {
        int pid = 0; Exception? failure = null;
        var worker = new Thread(() =>
        {
            try { pid = ExerciseWorkbook(); }
            catch (Exception ex) { failure = ex; }
            finally { GC.Collect(); GC.WaitForPendingFinalizers(); Marshal.CleanupUnusedObjectsInCurrentContext(); }
        });
        worker.SetApartmentState(ApartmentState.STA); worker.Start(); worker.Join();
        if (failure != null) throw new Exception("Live distribution check failed.", failure);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
        try { using var process = Process.GetProcessById(pid); if (!process.WaitForExit(30000)) throw new Exception($"Owned Excel process {pid} did not exit."); } catch (ArgumentException) { }
        Console.WriteLine("LIVE DISTRIBUTION PASS: private Excel exited; existing Excel instances untouched.");
    }
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static int ExerciseWorkbook()
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        dynamic? app = null, book = null, sheet = null; int pid = 0;
        string path = Path.GetFullPath("artifacts/distributions-live.xlsx");
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out uint processId); pid = (int)processId;
            if (existing.Contains(pid)) { Marshal.ReleaseComObject(app); app = null; throw new Exception("Refusing to use an existing Excel instance."); }
            app.Visible = false; app.DisplayAlerts = false; app.AutomationSecurity = 3;
            ExcelDnaUtil.Application = app;
            book = app.Workbooks.Add(); sheet = book.Worksheets[1]; sheet.Name = "Model";
            SimulationModel.Clear();
            foreach (var kind in Enum.GetValues<DistributionKind>().Where(k => k >= DistributionKind.Exponential))
            {
                int row = (int)kind - 5; var p = DistributionCatalog.Defaults(kind);
                sheet.Range["A" + row].Formula2 = "=50*2"; sheet.Range["A" + row].NumberFormat = "0.000";
                sheet.Range["B" + row].Formula2 = "=A" + row;
                SimulationModel.Assumptions.Add(new() { Name = kind.ToString(), SheetName = "Model", CellAddress = "A" + row,
                    Distribution = (DistributionType)kind, Parameter1 = p[0], Parameter2 = p[1], Parameter3 = p[2], Parameter4 = p[3],
                    ProbabilityTable = kind == DistributionKind.Discrete ? new([new(-10, .2), new(2.5, .5), new(30, .3)]) : null });
                SimulationModel.Forecasts.Add(new() { Name = kind.ToString(), SheetName = "Model", CellAddress = "B" + row, TargetSettings = new(1, TargetDirection.AtOrBelow, 50) });
            }
            WorkbookPersistence.SaveModel(); WorkbookPersistence.SaveSimulationSettings(new(100, SimulationSeedMode.Fixed, 912));
            Marshal.ReleaseComObject(sheet); sheet = null;
            book.SaveAs(path, 51); book.Close(false); Marshal.ReleaseComObject(book); book = app.Workbooks.Open(path);
            var validation = WorkbookPersistence.LoadModelForSimulation(); if (!validation.IsValid) throw new Exception(validation.ToString());
            if (SimulationModel.Assumptions.Count != 8 || SimulationModel.Assumptions.Single(a => a.Distribution == DistributionType.Discrete).ProbabilityTable?.Outcomes.Count != 3) throw new Exception("Disk round trip failed.");
            var first = SimulationService.TryRun(); var replay = SimulationService.TryRun();
            if (!first.Succeeded || !replay.Succeeded) throw new Exception("Live execution failed: " + first.DiagnosticException + " " + first.Validation);
            for (int i = 0; i < 8; i++) if (!first.Result!.ForecastResults[i].Values.SequenceEqual(replay.Result!.ForecastResults[i].Values)) throw new Exception("Seed replay failed.");
            // Configure a valid mixed continuous/discrete latent relationship and exercise production loading.
            var inputs = SimulationModel.Assumptions.ToArray();
            WorkbookPersistence.SaveCorrelations((object)book, inputs, [new(inputs[0].CellLink, inputs[1].CellLink, .8)]);
            var correlated = SimulationService.TryRun(); if (!correlated.Succeeded) throw new Exception("Mixed correlation failed: " + correlated.DiagnosticException);
            sheet = book.Worksheets["Model"];
            for (int row = 1; row <= 8; row++) if ((string)sheet.Range["A" + row].Formula2 != "=50*2" || (string)sheet.Range["A" + row].NumberFormat != "0.000") throw new Exception("Input restoration failed.");
            foreach (var input in inputs)
            {
                string description = (string)typeof(ReportExporter).GetMethod("GetAssumptionDescription", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!.Invoke(null, [input])!;
                if (description != input.ParameterDescription) throw new Exception("Report parameters missing.");
            }
            book.Save();
            Console.WriteLine("LIVE DISTRIBUTION PASS: all eight disk round trips; independent/replayed/mixed correlated simulations; eight forecasts and sensitivities; input formulas/format restored; report parameter descriptions.");
        }
        finally
        {
            SimulationModel.Clear(); ExcelDnaUtil.Application = null!;
            if (sheet != null) { Marshal.ReleaseComObject(sheet); sheet = null; }
            if (book != null) { book.Close(false); Marshal.ReleaseComObject(book); }
            if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); }
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
        }
        return pid;
    }
}

