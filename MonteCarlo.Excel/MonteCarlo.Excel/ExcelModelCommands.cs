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
            ExcelInterop.Application excelApp =
                (ExcelInterop.Application)ExcelDnaUtil.Application;

            ExcelInterop.Worksheet worksheet =
                (ExcelInterop.Worksheet)excelApp.ActiveSheet;

            try
            {
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
                System.Windows.Forms.MessageBox.Show(
                    ex.Message,
                    "Monte Carlo Simulation",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                excelApp.StatusBar = false;
            }
        }
    }
}