using System;
using ExcelDna.Integration;

namespace MonteCarlo.Excel
{
    public static class WorkbookPersistence
    {
        private const string ConfigSheetName =
            "__MonteCarloConfig";


        // =========================================================
        // SAVE MODEL TO WORKBOOK
        // =========================================================

        public static void SaveModel()
        {
            dynamic excelApp =
                ExcelDnaUtil.Application;

            dynamic workbook =
                excelApp.ActiveWorkbook;


            if (workbook == null)
            {
                throw new Exception(
                    "No active workbook.");
            }


            dynamic configSheet =
                GetOrCreateConfigSheet(
                    workbook);


            // Clear previous configuration.
            configSheet.Cells.Clear();


            // =====================================================
            // HEADER
            // =====================================================

            configSheet.Cells[1, 1].Value2 =
                "Type";

            configSheet.Cells[1, 2].Value2 =
                "Sheet";

            configSheet.Cells[1, 3].Value2 =
                "Cell";

            configSheet.Cells[1, 4].Value2 =
                "Distribution";

            configSheet.Cells[1, 5].Value2 =
                "Parameter1";

            configSheet.Cells[1, 6].Value2 =
                "Parameter2";

            configSheet.Cells[1, 7].Value2 =
                "Parameter3";

            configSheet.Cells[1, 8].Value2 =
                "Name";


            int row =
                2;


            // =====================================================
            // SAVE ASSUMPTIONS
            // =====================================================

            foreach (
                AssumptionDefinition assumption
                in SimulationModel.Assumptions)
            {
                configSheet.Cells[
                    row,
                    1].Value2 =
                    "Assumption";


                configSheet.Cells[
                    row,
                    2].Value2 =
                    assumption.SheetName;


                configSheet.Cells[
                    row,
                    3].Value2 =
                    assumption.CellAddress;


                configSheet.Cells[
                    row,
                    4].Value2 =
                    assumption
                        .Distribution
                        .ToString();


                configSheet.Cells[
                    row,
                    5].Value2 =
                    assumption.Parameter1;


                configSheet.Cells[
                    row,
                    6].Value2 =
                    assumption.Parameter2;


                configSheet.Cells[
                    row,
                    7].Value2 =
                    assumption.Parameter3;


                configSheet.Cells[
                    row,
                    8].Value2 =
                    assumption.Name;


                row++;
            }


            // =====================================================
            // SAVE FORECASTS
            // =====================================================

            foreach (
                ForecastDefinition forecast
                in SimulationModel.Forecasts)
            {
                configSheet.Cells[
                    row,
                    1].Value2 =
                    "Forecast";


                configSheet.Cells[
                    row,
                    2].Value2 =
                    forecast.SheetName;


                configSheet.Cells[
                    row,
                    3].Value2 =
                    forecast.CellAddress;


                configSheet.Cells[
                    row,
                    4].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    5].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    6].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    7].Value2 =
                    "";


                configSheet.Cells[
                    row,
                    8].Value2 =
                    forecast.Name;


                row++;
            }


            // =====================================================
            // VERY HIDDEN CONFIG SHEET
            // =====================================================

            configSheet.Visible =
                2;
        }


        // =========================================================
        // LOAD MODEL FROM WORKBOOK
        // =========================================================

        public static void LoadModel()
        {
            dynamic excelApp =
                ExcelDnaUtil.Application;

            dynamic workbook =
                excelApp.ActiveWorkbook;


            // Important:
            // Always clear current in-memory state first.
            SimulationModel.Clear();


            if (workbook == null)
            {
                return;
            }


            dynamic configSheet =
                FindConfigSheet(
                    workbook);


            if (configSheet == null)
            {
                return;
            }


            int row =
                2;


            while (true)
            {
                object typeValue =
                    configSheet.Cells[
                        row,
                        1].Value2;


                if (typeValue == null)
                {
                    break;
                }


                string type =
                    Convert.ToString(
                        typeValue)
                    ?? "";


                if (
                    string.Equals(
                        type,
                        "Assumption",
                        StringComparison.OrdinalIgnoreCase))
                {
                    LoadAssumption(
                        configSheet,
                        row);
                }


                else if (
                    string.Equals(
                        type,
                        "Forecast",
                        StringComparison.OrdinalIgnoreCase))
                {
                    LoadForecast(
                        configSheet,
                        row);
                }


                row++;
            }
        }


        // =========================================================
        // CLEAR SAVED MODEL
        // =========================================================

        public static void ClearSavedModel()
        {
            dynamic excelApp =
                ExcelDnaUtil.Application;

            dynamic workbook =
                excelApp.ActiveWorkbook;


            SimulationModel.Clear();


            if (workbook == null)
            {
                return;
            }


            dynamic configSheet =
                FindConfigSheet(
                    workbook);


            if (configSheet == null)
            {
                return;
            }


            bool originalDisplayAlerts =
                excelApp.DisplayAlerts;


            try
            {
                excelApp.DisplayAlerts =
                    false;


                configSheet.Delete();
            }
            finally
            {
                excelApp.DisplayAlerts =
                    originalDisplayAlerts;
            }
        }


        // =========================================================
        // LOAD ASSUMPTION
        // =========================================================

        private static void LoadAssumption(
            dynamic configSheet,
            int row)
        {
            string sheetName =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        2].Value2)
                ?? "";


            string cellAddress =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        3].Value2)
                ?? "";


            string distributionText =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        4].Value2)
                ?? "";


            double parameter1 =
                GetDoubleValue(
                    configSheet.Cells[
                        row,
                        5].Value2);


            double parameter2 =
                GetDoubleValue(
                    configSheet.Cells[
                        row,
                        6].Value2);


            double parameter3 =
                GetDoubleValue(
                    configSheet.Cells[
                        row,
                        7].Value2);


            string name =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        8].Value2)
                ?? "";


            if (
                string.IsNullOrWhiteSpace(
                    sheetName)
                ||
                string.IsNullOrWhiteSpace(
                    cellAddress))
            {
                return;
            }


            if (!Enum.TryParse(
                    distributionText,
                    true,
                    out DistributionType distribution))
            {
                return;
            }


            // Backward compatibility:
            // older workbooks may have no assumption name.
            if (string.IsNullOrWhiteSpace(name))
            {
                name =
                    $"{sheetName}!{cellAddress}";
            }


            SimulationModel.Assumptions.Add(
                new AssumptionDefinition
                {
                    Name =
                        name,

                    SheetName =
                        sheetName,

                    CellAddress =
                        cellAddress,

                    Distribution =
                        distribution,

                    Parameter1 =
                        parameter1,

                    Parameter2 =
                        parameter2,

                    Parameter3 =
                        parameter3
                });
        }


        // =========================================================
        // LOAD FORECAST
        // =========================================================

        private static void LoadForecast(
            dynamic configSheet,
            int row)
        {
            string sheetName =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        2].Value2)
                ?? "";


            string cellAddress =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        3].Value2)
                ?? "";


            string name =
                Convert.ToString(
                    configSheet.Cells[
                        row,
                        8].Value2)
                ?? "";


            if (
                string.IsNullOrWhiteSpace(
                    sheetName)
                ||
                string.IsNullOrWhiteSpace(
                    cellAddress))
            {
                return;
            }


            if (string.IsNullOrWhiteSpace(name))
            {
                name =
                    $"{sheetName}!{cellAddress}";
            }


            SimulationModel.Forecasts.Add(
                new ForecastDefinition
                {
                    Name =
                        name,

                    SheetName =
                        sheetName,

                    CellAddress =
                        cellAddress
                });
        }


        // =========================================================
        // GET OR CREATE CONFIG SHEET
        // =========================================================

        private static dynamic GetOrCreateConfigSheet(
            dynamic workbook)
        {
            dynamic existingSheet =
                FindConfigSheet(
                    workbook);


            if (existingSheet != null)
            {
                return existingSheet;
            }


            dynamic newSheet =
                workbook.Worksheets.Add();


            newSheet.Name =
                ConfigSheetName;


            return newSheet;
        }


        // =========================================================
        // FIND CONFIG SHEET
        // =========================================================

        private static dynamic FindConfigSheet(
            dynamic workbook)
        {
            int sheetCount =
                workbook.Worksheets.Count;


            for (
                int i = 1;
                i <= sheetCount;
                i++)
            {
                dynamic sheet =
                    workbook.Worksheets[
                        i];


                string sheetName =
                    Convert.ToString(
                        sheet.Name)
                    ?? "";


                if (
                    string.Equals(
                        sheetName,
                        ConfigSheetName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return
                        sheet;
                }
            }


            return
                null;
        }


        // =========================================================
        // SAFE DOUBLE CONVERSION
        // =========================================================

        private static double GetDoubleValue(
            object value)
        {
            if (value == null)
            {
                return
                    0;
            }


            try
            {
                return
                    Convert.ToDouble(
                        value);
            }
            catch
            {
                return
                    0;
            }
        }
    }
}