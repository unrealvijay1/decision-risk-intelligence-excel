using System;
using System.Diagnostics;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace MonteCarlo.Excel
{
    public class MonteCarloAddIn : IExcelAddIn
    {
        private dynamic? excelApp;

        public void AutoOpen()
        {
            try
            {
                WorkbookPersistence.ReportRestoreProblem = message => MessageBox.Show(message,
                    "Monte Carlo — model restore", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                excelApp = ExcelDnaUtil.Application;

                // When a workbook is opened
                excelApp.WorkbookOpen +=
                    new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(OnWorkbookOpen);

                // When user switches between workbooks
                excelApp.WorkbookActivate +=
                    new Microsoft.Office.Interop.Excel.AppEvents_WorkbookActivateEventHandler(OnWorkbookActivate);
            }
            catch (Exception ex)
            {
                // Do not prevent Excel from loading the add-in.
                Trace.TraceError("Workbook event registration failed: {0}", ex);
            }
            // Initial restoration must not depend on successful event registration.
            TryLoadActiveWorkbook();
        }


        public void AutoClose()
        {
            try
            {
                if (excelApp != null)
                {
                    excelApp.WorkbookOpen -=
                        new Microsoft.Office.Interop.Excel.AppEvents_WorkbookOpenEventHandler(OnWorkbookOpen);

                    excelApp.WorkbookActivate -=
                        new Microsoft.Office.Interop.Excel.AppEvents_WorkbookActivateEventHandler(OnWorkbookActivate);
                }
            }
            catch
            {
            }

            excelApp = null;
            WorkbookPersistence.ReportRestoreProblem = null;
        }


        // =========================================================
        // WORKBOOK OPENED
        // =========================================================

        private void OnWorkbookOpen(
            dynamic workbook)
        {
            TryLoadActiveWorkbook();
        }


        // =========================================================
        // WORKBOOK ACTIVATED
        // =========================================================

        private void OnWorkbookActivate(
            dynamic workbook)
        {
            TryLoadActiveWorkbook();
        }


        // =========================================================
        // LOAD MONTE CARLO MODEL
        // =========================================================

        private static void TryLoadActiveWorkbook()
        {
            try
            {
                WorkbookPersistence.LoadModel();
            }
            catch (Exception ex)
            {
                Trace.TraceError("Model restore failed: {0}", ex);
                MessageBox.Show("The saved Monte Carlo model could not be restored. Check the workbook and its model definitions in Model Manager.",
                    "Monte Carlo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
