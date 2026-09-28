using System.Diagnostics;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel;

internal static class SpcLiveChecks
{
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    public static void Run()
    {
        uint pid = 0; Exception? failure = null;
        var worker = new Thread(() =>
        {
            try { pid = RunWorkbook(); }
            catch (Exception ex) { failure = ex; }
            finally { GC.Collect(); GC.WaitForPendingFinalizers(); Marshal.CleanupUnusedObjectsInCurrentContext(); }
        });
        worker.SetApartmentState(ApartmentState.STA); worker.Start(); worker.Join();
        if (failure != null) throw new Exception("Live SPC check failed.", failure);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
        try { using var process = Process.GetProcessById((int)pid); if (!process.WaitForExit(30000)) throw new Exception($"Owned SPC Excel process {pid} did not exit."); }
        catch (ArgumentException) { }
        Console.WriteLine("LIVE SPC PASS: private Excel process exited; existing Excel processes were not used.");
    }
    private static uint RunWorkbook()
    {
        var existing = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        dynamic? app = null, book = null; uint pid = 0;
        string path = Path.Combine(Path.GetTempPath(), "MonteCarlo_SPC_" + Guid.NewGuid().ToString("N") + ".xlsx");
        var owned = new List<object>();
        dynamic Own(object item) { owned.Add(item); return item; }
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out pid);
            if (existing.Contains((int)pid)) throw new Exception("Refusing to use an existing Excel process.");
            Own(app); app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false; app.AutomationSecurity = 3;
            book = Own(app.Workbooks.Add()); ExcelDnaUtil.Application = app;
            dynamic dataSheet = Own(book.Worksheets[1]); dataSheet.Name = "Data";
            var raw = new object[40, 1]; for (int i = 0; i < 40; i++) raw[i, 0] = i == 30 ? 35d : 10d + (i % 2 == 0 ? -1 : 1);
            dynamic data = Own(dataSheet.Range["A1:A40"]); data.Value2 = raw;
            dynamic labels = Own(dataSheet.Range["B1:B40"]); labels.NumberFormat = "@";
            var labelValues = new object[40, 1]; for (int i = 0; i < 40; i++) labelValues[i, 0] = i == 0 ? "=2+2" : "Observation " + (i + 1);
            labels.Value2 = labelValues;
            double[] values = SpcDataReader.ReadValues((object)book, (object)data);
            string[] text = SpcDataReader.ReadLabels((object)book, (object)labels, 40);
            if (values[30] != 35 || text[0] != "=2+2") throw new Exception("Range/label reading failed.");
            dynamic dateLabels = Own(dataSheet.Range["C1:C40"]); dateLabels.NumberFormat = "dd mmm yyyy"; dateLabels.ColumnWidth = 2;
            var dateValues = new object[40, 1]; for (int i = 0; i < 40; i++) dateValues[i, 0] = new DateTime(2026, 1, 1).AddDays(i).ToOADate();
            dateLabels.Value2 = dateValues;
            string[] dates = SpcDataReader.ReadLabels((object)book, (object)dateLabels, 40);
            if (dates[0].Contains('#') || !dates[0].Contains("2026")) throw new Exception("Narrow source columns obscured date labels.");
            dataSheet.Range["A10"].Formula = "=NA()";
            try { SpcDataReader.ReadValues((object)book, (object)data); throw new Exception("Excel error was accepted."); }
            catch (ArgumentException ex) { if (!ex.Message.Contains("10")) throw; }
            data.Value2 = raw;
            var config = new SpcConfiguration { ProcessName = "=SPC text", BaselineMode = SpcBaselineMode.FirstN, Target = new(12, TargetDirection.AtOrBelow) };
            var initial = SpcAnalysisService.Analyze((object)book, (object)dataSheet.Range["A1:A20"], (object)dataSheet.Range["B1:B20"], config, true, null);
            dataSheet.Range["A2"].Formula = "=11";
            var outcome = SpcAnalysisService.Analyze((object)book, (object)data, (object)labels, config, false, initial.Saved);
            var result = outcome.Result; var saved = outcome.Saved;
            if (result.Baseline != initial.Result.Baseline || result.Observations.Count != 40 || (string)dataSheet.Range["A2"].Formula != "=11")
                throw new Exception("Appended observations changed fixed limits or source formulas.");
            int nameCount = (int)book.Names.Count;
            SpcAnalysisService.Analyze((object)book, (object)data, (object)labels, config, false, saved);
            if ((int)book.Names.Count != nameCount) throw new Exception("Unchanged references created duplicate names.");
            string dataLink = saved.DataLink, baselineLink = saved.BaselineLink;
            dataSheet.Name = "Renamed Data"; dataSheet.Rows[1].Insert();
            dynamic moved = Own(SpcDataReader.Resolve((object)book, dataLink));
            dynamic movedBaseline = Own(SpcDataReader.Resolve((object)book, baselineLink));
            if ((int)moved.Row != 2 || (string)moved.Worksheet.Name != "Renamed Data") throw new Exception("Tracked range did not move.");
            var period = SpcDataReader.LocateBaseline((object)book, (object)moved, (object)movedBaseline);
            if (period != (1, 20)) throw new Exception("Tracked baseline alignment failed.");
            SpcExporter.Export((object)book, result, "Renamed Data!A2:A41");
            dynamic report = Own(app.ActiveSheet);
            if ((int)report.ChartObjects().Count != 2 || report.Cells[SpcExporter.HeaderRow + 1, 4].Value2 != null)
                throw new Exception("Native charts or blank first MR failed.");
            if ((bool)report.Cells[SpcExporter.HeaderRow + 1, 2].HasFormula || (bool)report.Cells[1, 2].HasFormula)
                throw new Exception("User text was interpreted as a formula.");
            if ((double)dataSheet.Range["A32"].Value2 != 35) throw new Exception("Export modified source data.");
            SpcExporter.Export((object)book, result, "Same source");
            if ((string)app.ActiveSheet.Name == (string)report.Name) throw new Exception("Export overwrote a worksheet.");
            book.SaveAs(path, 51); book.Close(false); book = Own(app.Workbooks.Open(path));
            var restored = WorkbookPersistence.LoadSpcAnalysis((object)book);
            if (restored != saved) throw new Exception("SPC metadata save/reopen changed fixed limits.");
            dynamic reopened = Own(SpcDataReader.Resolve((object)book, dataLink));
            if ((int)reopened.Row != 2) throw new Exception("Source name did not survive reopen.");
            book.Worksheets["Renamed Data"].Delete();
            try { SpcDataReader.Resolve((object)book, dataLink); throw new Exception("Broken name fell back to stale coordinates."); }
            catch (ArgumentException) { }
            Console.WriteLine("LIVE SPC PASS: strict errors, labels, source isolation, native charts, unique exports, tracked rename/insert/delete and save/reopen.");
        }
        finally
        {
            if (pid != 0 && !existing.Contains((int)pid))
            {
                try { if (book != null) book.Close(false); }
                finally { if (app != null) app.Quit(); }
            }
            for (int i = owned.Count - 1; i >= 0; i--) if (Marshal.IsComObject(owned[i])) Marshal.ReleaseComObject(owned[i]);
            ExcelDnaUtil.Application = null!;
            if (File.Exists(path)) File.Delete(path);
        }
        return pid;
    }
}
