using System;
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
                excelApp = ExcelDnaUtil.Application;

                // When a workbook is opened
                excelApp.WorkbookOpen +=
                    new Action<dynamic>(OnWorkbookOpen);

                // When user switches between workbooks
                excelApp.WorkbookActivate +=
                    new Action<dynamic>(OnWorkbookActivate);

                // Load current workbook if one is already open
                TryLoadActiveWorkbook();
            }
            catch
            {
                // Do not prevent Excel from loading the add-in.
            }
        }


        public void AutoClose()
        {
            try
            {
                if (excelApp != null)
                {
                    excelApp.WorkbookOpen -=
                        new Action<dynamic>(OnWorkbookOpen);

                    excelApp.WorkbookActivate -=
                        new Action<dynamic>(OnWorkbookActivate);
                }
            }
            catch
            {
            }

            excelApp = null;
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
            catch
            {
                // Intentionally silent.
                //
                // A workbook without a Monte Carlo model
                // is perfectly valid.
            }
        }
    }
}