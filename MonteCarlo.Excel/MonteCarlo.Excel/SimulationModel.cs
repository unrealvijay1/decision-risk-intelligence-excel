using System.Collections.Generic;
using MonteCarlo.Core;

namespace MonteCarlo.Excel
{
    // =============================================================
    // DISTRIBUTION TYPES USED BY EXCEL MODEL
    // =============================================================

    public enum DistributionType
    {
        Normal,
        Triangular,
        Pert,
        Uniform,
        Lognormal,
        Beta
    }


    // =============================================================
    // ASSUMPTION DEFINITION
    // =============================================================

    public class AssumptionDefinition
    {
        // Optional hidden workbook name; Excel tracks renames and structural cell edits.
        public string CellLink { get; set; } = "";
        // Friendly name shown in Model Manager,
        // sensitivity analysis and reports.
        //
        // Examples:
        // Duration
        // Team Size
        // Daily Rate
        // Productivity
        public string Name { get; set; } =
            "";


        // Excel worksheet containing the assumption cell.
        public string SheetName { get; set; } =
            "";


        // Excel address of the assumption cell.
        //
        // Example:
        // B3
        public string CellAddress { get; set; } =
            "";


        public DistributionType Distribution { get; set; }


        // =========================================================
        // DISTRIBUTION PARAMETERS
        //
        // NORMAL
        //
        // Parameter1 = Mean
        // Parameter2 = Standard Deviation
        //
        //
        // LOGNORMAL
        //
        // Parameter1 = Log Mean
        // Parameter2 = Log Standard Deviation
        //
        //
        // UNIFORM
        //
        // Parameter1 = Minimum
        // Parameter2 = Maximum
        //
        //
        // TRIANGULAR
        //
        // Parameter1 = Minimum
        // Parameter2 = Most Likely / Mode
        // Parameter3 = Maximum
        //
        //
        // PERT
        //
        // Parameter1 = Minimum
        // Parameter2 = Most Likely
        // Parameter3 = Maximum
        //
        //
        // BETA
        //
        // Parameter1 = Minimum
        // Parameter2 = Maximum
        // Parameter3 = Alpha
        // Parameter4 = Beta
        //
        // =========================================================


        public double Parameter1 { get; set; }


        public double Parameter2 { get; set; }


        public double Parameter3 { get; set; }


        public double Parameter4 { get; set; }
    }


    // =============================================================
    // FORECAST DEFINITION
    // =============================================================

    public sealed record ForecastTargetSettings(double? Target, TargetDirection? Direction, double? RequestedConfidence);

    public class ForecastDefinition
    {
        // Null preserves the historical P80 default; a non-null record may explicitly clear the target.
        public ForecastTargetSettings? TargetSettings { get; set; }
        public string CellLink { get; set; } = "";
        // Friendly forecast name.
        //
        // Examples:
        // Total Cost
        // Completion Date
        // Revenue
        public string Name { get; set; } =
            "";


        public string SheetName { get; set; } =
            "";


        public string CellAddress { get; set; } =
            "";
    }


    // =============================================================
    // SIMULATION MODEL
    // =============================================================

    public static class SimulationModel
    {
        // =========================================================
        // ASSUMPTIONS
        // =========================================================

        public static List<AssumptionDefinition> Assumptions { get; }
            =
            new List<AssumptionDefinition>();


        // =========================================================
        // FORECASTS
        // =========================================================

        public static List<ForecastDefinition> Forecasts { get; }
            =
            new List<ForecastDefinition>();


        // =========================================================
        // CLEAR COMPLETE MODEL
        // =========================================================

        public static void Clear()
        {
            Assumptions.Clear();

            Forecasts.Clear();
        }


        // =========================================================
        // CLEAR ASSUMPTIONS ONLY
        // =========================================================

        public static void ClearAssumptions()
        {
            Assumptions.Clear();
        }


        // =========================================================
        // CLEAR FORECASTS ONLY
        // =========================================================

        public static void ClearForecasts()
        {
            Forecasts.Clear();
        }


        // =========================================================
        // FIND ASSUMPTION
        // =========================================================

        public static AssumptionDefinition? FindAssumption(
            string sheetName,
            string cellAddress)
        {
            foreach (
                AssumptionDefinition assumption
                in Assumptions)
            {
                if (
                    string.Equals(
                        assumption.SheetName,
                        sheetName,
                        System.StringComparison.OrdinalIgnoreCase)
                    &&
                    string.Equals(
                        assumption.CellAddress,
                        cellAddress,
                        System.StringComparison.OrdinalIgnoreCase))
                {
                    return
                        assumption;
                }
            }


            return
                null;
        }


        // =========================================================
        // FIND FORECAST
        // =========================================================

        public static ForecastDefinition? FindForecast(
            string sheetName,
            string cellAddress)
        {
            foreach (
                ForecastDefinition forecast
                in Forecasts)
            {
                if (
                    string.Equals(
                        forecast.SheetName,
                        sheetName,
                        System.StringComparison.OrdinalIgnoreCase)
                    &&
                    string.Equals(
                        forecast.CellAddress,
                        cellAddress,
                        System.StringComparison.OrdinalIgnoreCase))
                {
                    return
                        forecast;
                }
            }


            return
                null;
        }
    }
}
