using System;
using ExcelDna.Integration;

namespace MonteCarlo.Excel
{
    public static class ReportExporter
    {
        // =========================================================
        // EXPORT COMPLETE SIMULATION REPORT
        // =========================================================

        public static void Export(
            SimulationRunResult simulationResult)
        {
            if (simulationResult == null)
            {
                throw new ArgumentNullException(
                    nameof(simulationResult));
            }


            dynamic excelApp =
                ExcelDnaUtil.Application;


            dynamic workbook =
                excelApp.ActiveWorkbook;


            if (workbook == null)
            {
                throw new Exception(
                    "No active workbook.");
            }


            bool originalScreenUpdating =
                excelApp.ScreenUpdating;


            try
            {
                excelApp.ScreenUpdating =
                    false;


                // =================================================
                // CREATE REPORT SHEET
                // =================================================

                string sheetName =
                    CreateUniqueSheetName(
                        workbook);


                dynamic reportSheet =
                    workbook.Worksheets.Add();


                reportSheet.Name =
                    sheetName;


                // =================================================
                // TITLE
                // =================================================

                reportSheet.Cells[
                    1,
                    1].Value2 =
                    "MONTE CARLO SIMULATION REPORT";


                dynamic titleRange =
                    reportSheet.Range[
                        "A1:H1"];


                titleRange.Merge();


                titleRange.Font.Bold =
                    true;


                titleRange.Font.Size =
                    18;


                titleRange.HorizontalAlignment =
                    -4108;


                // =================================================
                // RUN INFORMATION
                // =================================================

                reportSheet.Cells[
                    3,
                    1].Value2 =
                    "RUN INFORMATION";


                reportSheet.Cells[
                    3,
                    1].Font.Bold =
                    true;


                reportSheet.Cells[
                    4,
                    1].Value2 =
                    "Generated";


                reportSheet.Cells[
                    4,
                    2].Value2 =
                    DateTime.Now.ToString(
                        "yyyy-MM-dd HH:mm:ss");


                reportSheet.Cells[
                    5,
                    1].Value2 =
                    "Trials";


                reportSheet.Cells[
                    5,
                    2].Value2 =
                    simulationResult.Trials;


                reportSheet.Cells[
                    6,
                    1].Value2 =
                    "Assumptions";


                reportSheet.Cells[
                    6,
                    2].Value2 =
                    SimulationModel
                        .Assumptions
                        .Count;


                reportSheet.Cells[
                    7,
                    1].Value2 =
                    "Forecasts";


                reportSheet.Cells[
                    7,
                    2].Value2 =
                    simulationResult
                        .ForecastResults
                        .Count;


                // =================================================
                // ASSUMPTIONS
                // =================================================

                int currentRow =
                    10;


                reportSheet.Cells[
                    currentRow,
                    1].Value2 =
                    "ASSUMPTIONS";


                reportSheet.Cells[
                    currentRow,
                    1].Font.Bold =
                    true;


                currentRow +=
                    2;


                reportSheet.Cells[
                    currentRow,
                    1].Value2 =
                    "Name";


                reportSheet.Cells[
                    currentRow,
                    2].Value2 =
                    "Sheet";


                reportSheet.Cells[
                    currentRow,
                    3].Value2 =
                    "Cell";


                reportSheet.Cells[
                    currentRow,
                    4].Value2 =
                    "Distribution";


                reportSheet.Cells[
                    currentRow,
                    5].Value2 =
                    "Parameter 1";


                reportSheet.Cells[
                    currentRow,
                    6].Value2 =
                    "Parameter 2";


                reportSheet.Cells[
                    currentRow,
                    7].Value2 =
                    "Parameter 3";


                reportSheet.Cells[
                    currentRow,
                    8].Value2 =
                    "Interpretation";


                dynamic assumptionHeader =
                    reportSheet.Range[
                        reportSheet.Cells[
                            currentRow,
                            1],
                        reportSheet.Cells[
                            currentRow,
                            8]];


                assumptionHeader.Font.Bold =
                    true;


                currentRow++;


                foreach (
                    AssumptionDefinition assumption
                    in SimulationModel.Assumptions)
                {
                    reportSheet.Cells[
                        currentRow,
                        1].Value2 =
                        GetAssumptionName(
                            assumption);


                    reportSheet.Cells[
                        currentRow,
                        2].Value2 =
                        assumption.SheetName;


                    reportSheet.Cells[
                        currentRow,
                        3].Value2 =
                        assumption.CellAddress;


                    reportSheet.Cells[
                        currentRow,
                        4].Value2 =
                        assumption
                            .Distribution
                            .ToString();


                    reportSheet.Cells[
                        currentRow,
                        5].Value2 =
                        assumption.Parameter1;


                    reportSheet.Cells[
                        currentRow,
                        6].Value2 =
                        assumption.Parameter2;


                    reportSheet.Cells[
                        currentRow,
                        7].Value2 =
                        assumption.Parameter3;


                    reportSheet.Cells[
                        currentRow,
                        8].Value2 =
                        GetAssumptionDescription(
                            assumption);


                    currentRow++;
                }


                // =================================================
                // FORECAST RESULTS
                // =================================================

                currentRow +=
                    3;


                reportSheet.Cells[
                    currentRow,
                    1].Value2 =
                    "FORECAST RESULTS";


                reportSheet.Cells[
                    currentRow,
                    1].Font.Bold =
                    true;


                currentRow +=
                    2;


                foreach (
                    ForecastRunResult forecastResult
                    in simulationResult.ForecastResults)
                {
                    string forecastName =
                        GetForecastName(
                            forecastResult);


                    // ---------------------------------------------
                    // FORECAST HEADER
                    // ---------------------------------------------

                    reportSheet.Cells[
                        currentRow,
                        1].Value2 =
                        forecastName;


                    reportSheet.Cells[
                        currentRow,
                        1].Font.Bold =
                        true;


                    reportSheet.Cells[
                        currentRow,
                        2].Value2 =
                        $"{forecastResult.Forecast.SheetName}!" +
                        $"{forecastResult.Forecast.CellAddress}";


                    currentRow +=
                        2;


                    // ---------------------------------------------
                    // SUMMARY HEADER
                    // ---------------------------------------------

                    reportSheet.Cells[
                        currentRow,
                        1].Value2 =
                        "Metric";


                    reportSheet.Cells[
                        currentRow,
                        2].Value2 =
                        "Value";


                    dynamic summaryHeader =
                        reportSheet.Range[
                            reportSheet.Cells[
                                currentRow,
                                1],
                            reportSheet.Cells[
                                currentRow,
                                2]];


                    summaryHeader.Font.Bold =
                        true;


                    currentRow++;


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "Trials",
                        forecastResult.Trials);


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "Mean",
                        forecastResult.Mean);


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "Std Dev",
                        forecastResult.StandardDeviation);


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "Minimum",
                        forecastResult.Minimum);


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "Maximum",
                        forecastResult.Maximum);


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "P10",
                        forecastResult.P10);


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "P50",
                        forecastResult.P50);


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "P80",
                        forecastResult.P80);


                    WriteMetric(
                        reportSheet,
                        ref currentRow,
                        "P90",
                        forecastResult.P90);


                    // ---------------------------------------------
                    // SENSITIVITY
                    // ---------------------------------------------

                    currentRow +=
                        2;


                    reportSheet.Cells[
                        currentRow,
                        1].Value2 =
                        "Sensitivity Analysis";


                    reportSheet.Cells[
                        currentRow,
                        1].Font.Bold =
                        true;


                    currentRow++;


                    reportSheet.Cells[
                        currentRow,
                        1].Value2 =
                        "Assumption";


                    reportSheet.Cells[
                        currentRow,
                        2].Value2 =
                        "Spearman Correlation";


                    dynamic sensitivityHeader =
                        reportSheet.Range[
                            reportSheet.Cells[
                                currentRow,
                                1],
                            reportSheet.Cells[
                                currentRow,
                                2]];


                    sensitivityHeader.Font.Bold =
                        true;


                    currentRow++;


                    foreach (
                        SensitivityResult sensitivity
                        in forecastResult.Sensitivities)
                    {
                        reportSheet.Cells[
                            currentRow,
                            1].Value2 =
                            sensitivity.Name;


                        reportSheet.Cells[
                            currentRow,
                            2].Value2 =
                            sensitivity.Correlation;


                        currentRow++;
                    }


                    currentRow +=
                        4;
                }


                // =================================================
                // FORMAT REPORT
                // =================================================

                dynamic usedRange =
                    reportSheet.UsedRange;


                usedRange.Columns.AutoFit();


                // Cap very wide columns.
                if (
                    reportSheet.Columns[
                        8].ColumnWidth >
                    55)
                {
                    reportSheet.Columns[
                        8].ColumnWidth =
                        55;
                }


                reportSheet.Columns[
                    8].WrapText =
                    true;


                reportSheet.Columns[
                    1].ColumnWidth =
                    Math.Max(
                        18,
                        reportSheet.Columns[
                            1].ColumnWidth);


                reportSheet.Columns[
                    2].ColumnWidth =
                    Math.Max(
                        14,
                        reportSheet.Columns[
                            2].ColumnWidth);


                // Generic number formatting.
                reportSheet.Columns[
                    5].NumberFormat =
                    "#,##0.00";


                reportSheet.Columns[
                    6].NumberFormat =
                    "#,##0.00";


                reportSheet.Columns[
                    7].NumberFormat =
                    "#,##0.00";


                // Activate and freeze top row.
                reportSheet.Activate();


                excelApp.ActiveWindow.SplitRow =
                    1;


                excelApp.ActiveWindow.FreezePanes =
                    true;


                reportSheet.Range[
                    "A1"].Select();


                MessageBox.Show(
                    $"Simulation report created successfully.\n\n" +
                    $"Worksheet: {sheetName}",
                    "Monte Carlo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            finally
            {
                excelApp.ScreenUpdating =
                    originalScreenUpdating;
            }
        }


        // =========================================================
        // WRITE METRIC
        // =========================================================

        private static void WriteMetric(
            dynamic sheet,
            ref int row,
            string name,
            object value)
        {
            sheet.Cells[
                row,
                1].Value2 =
                name;


            sheet.Cells[
                row,
                2].Value2 =
                value;


            row++;
        }


        // =========================================================
        // ASSUMPTION NAME
        // =========================================================

        private static string GetAssumptionName(
            AssumptionDefinition assumption)
        {
            if (
                !string.IsNullOrWhiteSpace(
                    assumption.Name))
            {
                return
                    assumption.Name;
            }


            return
                $"{assumption.SheetName}!" +
                $"{assumption.CellAddress}";
        }


        // =========================================================
        // FORECAST NAME
        // =========================================================

        private static string GetForecastName(
            ForecastRunResult result)
        {
            if (
                !string.IsNullOrWhiteSpace(
                    result.Forecast.Name))
            {
                return
                    result.Forecast.Name;
            }


            return
                $"{result.Forecast.SheetName}!" +
                $"{result.Forecast.CellAddress}";
        }


        // =========================================================
        // ASSUMPTION DESCRIPTION
        // =========================================================

        private static string GetAssumptionDescription(
            AssumptionDefinition assumption)
        {
            switch (
                assumption.Distribution)
            {
                case DistributionType.Normal:

                    return
                        $"Mean={assumption.Parameter1:N2}, " +
                        $"SD={assumption.Parameter2:N2}";


                case DistributionType.Pert:

                    return
                        $"Min={assumption.Parameter1:N2}, " +
                        $"Most Likely={assumption.Parameter2:N2}, " +
                        $"Max={assumption.Parameter3:N2}";


                case DistributionType.Triangular:

                    return
                        $"Min={assumption.Parameter1:N2}, " +
                        $"Mode={assumption.Parameter2:N2}, " +
                        $"Max={assumption.Parameter3:N2}";


                case DistributionType.Uniform:

                    return
                        $"Min={assumption.Parameter1:N2}, " +
                        $"Max={assumption.Parameter2:N2}";


                case DistributionType.Lognormal:
                    {
                        double mu =
                            assumption.Parameter1;


                        double sigma =
                            assumption.Parameter2;


                        double median =
                            Math.Exp(
                                mu);


                        double mean =
                            Math.Exp(
                                mu +
                                sigma *
                                sigma /
                                2.0);


                        double variance =
                            (
                                Math.Exp(
                                    sigma *
                                    sigma)
                                -
                                1.0
                            )
                            *
                            Math.Exp(
                                2.0 *
                                mu
                                +
                                sigma *
                                sigma);


                        double sd =
                            Math.Sqrt(
                                variance);


                        return
                            $"Mean≈{mean:N2}, " +
                            $"Median≈{median:N2}, " +
                            $"SD≈{sd:N2}, " +
                            $"Log μ={mu:N4}, " +
                            $"Log σ={sigma:N4}";
                    }


                default:

                    return
                        "";
            }
        }


        // =========================================================
        // UNIQUE REPORT SHEET NAME
        // =========================================================

        private static string CreateUniqueSheetName(
            dynamic workbook)
        {
            string baseName =
                "MonteCarlo_Report_" +
                DateTime.Now.ToString(
                    "HHmmss");


            string candidate =
                baseName;


            int counter =
                1;


            while (
                SheetExists(
                    workbook,
                    candidate))
            {
                candidate =
                    baseName +
                    "_" +
                    counter;


                counter++;
            }


            return
                candidate;
        }


        // =========================================================
        // SHEET EXISTS
        // =========================================================

        private static bool SheetExists(
            dynamic workbook,
            string name)
        {
            int count =
                workbook.Worksheets.Count;


            for (
                int i = 1;
                i <= count;
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
                        name,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return
                        true;
                }
            }


            return
                false;
        }
    }
}