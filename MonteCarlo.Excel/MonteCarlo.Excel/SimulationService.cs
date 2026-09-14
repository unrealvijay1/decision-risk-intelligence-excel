using System;
using System.Collections.Generic;
using System.Linq;

using ExcelDna.Integration;

namespace MonteCarlo.Excel
{
    public static class SimulationService
    {
        // =========================================================
        // RUN SIMULATION
        // =========================================================

        public static SimulationRunResult Run(
            int trials)
        {
            if (trials <= 0)
            {
                throw new ArgumentException(
                    "Trials must be greater than zero.",
                    nameof(trials));
            }


            WorkbookPersistence.LoadModel();


            if (SimulationModel.Assumptions.Count == 0)
            {
                throw new InvalidOperationException(
                    "No saved assumptions were found.");
            }


            if (SimulationModel.Forecasts.Count == 0)
            {
                throw new InvalidOperationException(
                    "No saved forecasts were found.");
            }


            dynamic excelApp =
                ExcelDnaUtil.Application;


            // =====================================================
            // ASSUMPTION SENSITIVITY NAMES
            // =====================================================

            var assumptionSensitivityNames =
                new Dictionary<
                    AssumptionDefinition,
                    string>();


            foreach (
                AssumptionDefinition assumption
                in SimulationModel.Assumptions)
            {
                string desiredName =
                    GetAssumptionDisplayName(
                        assumption);


                string finalName =
                    desiredName;


                int suffix =
                    2;


                while (
                    assumptionSensitivityNames
                        .Values
                        .Any(
                            x =>
                                string.Equals(
                                    x,
                                    finalName,
                                    StringComparison.OrdinalIgnoreCase)))
                {
                    finalName =
                        $"{desiredName} ({suffix})";


                    suffix++;
                }


                assumptionSensitivityNames[
                    assumption] =
                    finalName;
            }


            // =====================================================
            // ASSUMPTION SAMPLES
            // =====================================================

            var assumptionSamples =
                new Dictionary<
                    string,
                    double[]>();


            foreach (
                AssumptionDefinition assumption
                in SimulationModel.Assumptions)
            {
                assumptionSamples[
                    assumptionSensitivityNames[
                        assumption]] =
                    new double[trials];
            }


            // =====================================================
            // FORECAST SAMPLES
            // =====================================================

            var forecastSamples =
                new Dictionary<
                    ForecastDefinition,
                    double[]>();


            foreach (
                ForecastDefinition forecast
                in SimulationModel.Forecasts)
            {
                forecastSamples[
                    forecast] =
                    new double[trials];
            }


            // =====================================================
            // SAVE ORIGINAL ASSUMPTION VALUES
            // =====================================================

            var originalValues =
                new Dictionary<
                    AssumptionDefinition,
                    object?>();


            foreach (
                AssumptionDefinition assumption
                in SimulationModel.Assumptions)
            {
                dynamic sheet =
                    excelApp.Worksheets[
                        assumption.SheetName];


                dynamic cell =
                    sheet.Range[
                        assumption.CellAddress];


                originalValues[
                    assumption] =
                    cell.Value2;
            }


            // =====================================================
            // SAVE EXCEL STATE
            // =====================================================

            bool originalScreenUpdating =
                excelApp.ScreenUpdating;


            bool originalEnableEvents =
                excelApp.EnableEvents;


            object originalStatusBar =
                excelApp.StatusBar;


            excelApp.ScreenUpdating =
                false;


            excelApp.EnableEvents =
                false;


            int progressInterval =
                Math.Max(
                    1,
                    trials / 100);


            try
            {
                // =================================================
                // MAIN SIMULATION LOOP
                // =================================================

                for (
                    int i = 0;
                    i < trials;
                    i++)
                {
                    // ---------------------------------------------
                    // SAMPLE ALL ASSUMPTIONS
                    // ---------------------------------------------

                    foreach (
                        AssumptionDefinition assumption
                        in SimulationModel.Assumptions)
                    {
                        dynamic sheet =
                            excelApp.Worksheets[
                                assumption.SheetName];


                        dynamic cell =
                            sheet.Range[
                                assumption.CellAddress];


                        double sample =
                            GenerateSample(
                                assumption);


                        string sensitivityName =
                            assumptionSensitivityNames[
                                assumption];


                        assumptionSamples[
                            sensitivityName][i] =
                            sample;


                        cell.Value2 =
                            sample;
                    }


                    // ---------------------------------------------
                    // CALCULATE WORKBOOK
                    // ---------------------------------------------

                    excelApp.Calculate();


                    // ---------------------------------------------
                    // CAPTURE ALL FORECASTS
                    // ---------------------------------------------

                    foreach (
                        ForecastDefinition forecast
                        in SimulationModel.Forecasts)
                    {
                        dynamic forecastSheet =
                            excelApp.Worksheets[
                                forecast.SheetName];


                        dynamic forecastCell =
                            forecastSheet.Range[
                                forecast.CellAddress];


                        object forecastValue =
                            forecastCell.Value2;


                        if (forecastValue == null)
                        {
                            throw new Exception(
                                $"Forecast " +
                                $"{forecast.SheetName}!" +
                                $"{forecast.CellAddress} " +
                                $"returned no value.");
                        }


                        forecastSamples[
                            forecast][i] =
                            Convert.ToDouble(
                                forecastValue);
                    }


                    // ---------------------------------------------
                    // PROGRESS
                    // ---------------------------------------------

                    if (
                        i % progressInterval == 0
                        ||
                        i == trials - 1)
                    {
                        int percentage =
                            (int)(
                                (i + 1) *
                                100.0 /
                                trials);


                        excelApp.StatusBar =
                            $"Monte Carlo Simulation: " +
                            $"{percentage}% " +
                            $"({i + 1:N0}/{trials:N0})";
                    }
                }
            }
            finally
            {
                // =================================================
                // RESTORE ORIGINAL INPUT VALUES
                // =================================================

                foreach (
                    KeyValuePair<
                        AssumptionDefinition,
                        object?> item
                    in originalValues)
                {
                    dynamic sheet =
                        excelApp.Worksheets[
                            item.Key.SheetName];


                    dynamic cell =
                        sheet.Range[
                            item.Key.CellAddress];


                    cell.Value2 =
                        item.Value;
                }


                // Recalculate original workbook state.
                excelApp.Calculate();


                // Restore Excel settings.
                excelApp.ScreenUpdating =
                    originalScreenUpdating;


                excelApp.EnableEvents =
                    originalEnableEvents;


                try
                {
                    excelApp.StatusBar =
                        originalStatusBar;
                }
                catch
                {
                    excelApp.StatusBar =
                        false;
                }
            }


            // =====================================================
            // CREATE RESULT
            // =====================================================

            return
                new SimulationRunResult(
                    trials,
                    forecastSamples,
                    assumptionSamples);
        }


        // =========================================================
        // ASSUMPTION DISPLAY NAME
        // =========================================================

        private static string GetAssumptionDisplayName(
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
        // DISTRIBUTION SAMPLING
        // =========================================================

        private static double GenerateSample(
            AssumptionDefinition assumption)
        {
            MonteCarlo.Core.DistributionKind distribution =
                assumption.Distribution switch
                {
                    DistributionType.Normal =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Normal,

                    DistributionType.Triangular =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Triangular,

                    DistributionType.Pert =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Pert,

                    DistributionType.Uniform =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Uniform,

                    DistributionType.Lognormal =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Lognormal,

                    DistributionType.Beta =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Beta,

                    _ =>
                        throw new InvalidOperationException(
                            $"Unsupported distribution: " +
                            $"{assumption.Distribution}")
                };


            return
                MonteCarlo.Core
                    .DistributionSampler
                    .Sample(
                        distribution,
                        assumption.Parameter1,
                        assumption.Parameter2,
                        assumption.Parameter3,
                        assumption.Parameter4);
        }
    }
}