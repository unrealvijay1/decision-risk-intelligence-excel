using ExcelDna.Integration;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Excel
{
    public static class ExcelModelCommands
    {
        [ExcelCommand(
            Name = "MC_RUN_PROJECTMODEL",
            Description = "Runs the Monte Carlo project model")]
        public static void RunProjectModel()
        {
            ExcelInterop.Application? excelApp = null;
            object? originalStatus = null;
            bool statusCaptured = false;
            try
            {
                excelApp = (ExcelInterop.Application)ExcelDnaUtil.Application;
                originalStatus = excelApp.StatusBar;
                statusCaptured = true;
                ExcelInterop.Worksheet worksheet = (ExcelInterop.Worksheet)excelApp.ActiveSheet;
                int trials = 100;

                object[,] results =
                    ExcelModelSimulator
                        .RunProjectModel(trials);

                ExcelInterop.Range startCell =
                    worksheet.Range["D2"];

                int rows =
                    results.GetLength(0);

                int columns =
                    results.GetLength(1);

                ExcelInterop.Range outputRange =
                    startCell.Resize[
                        rows,
                        columns];

                outputRange.Value2 =
                    results;

                excelApp.StatusBar =
                    $"Monte Carlo simulation completed: {trials} trials";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.TraceError(ex.ToString());
                System.Windows.Forms.MessageBox.Show(
                    ex is SimulationRunException ? ex.Message :
                        "The simulation or report could not finish. Check that the workbook and output cells are available, then try again.",
                    "Monte Carlo Simulation",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                if (statusCaptured && excelApp != null)
                {
                    try { excelApp.StatusBar = originalStatus; }
                    catch (Exception ex) { System.Diagnostics.Trace.TraceError(ex.ToString()); }
                }
            }
        }
    }
}