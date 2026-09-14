using System.Collections.Generic;

namespace MonteCarlo.Excel
{
    public enum DistributionType
    {
        Normal,
        Triangular,
        Pert,
        Uniform,
        Lognormal
    }


    // =============================================================
    // ASSUMPTION
    // =============================================================

    public class AssumptionDefinition
    {
        // Friendly business name.
        //
        // Examples:
        // Daily Rate
        // Team Size
        // Duration
        // Demand
        public string Name { get; set; } = "";


        public string SheetName { get; set; } = "";


        public string CellAddress { get; set; } = "";


        public DistributionType Distribution { get; set; }


        // ---------------------------------------------------------
        // Parameter meanings
        //
        // Normal
        // Parameter1 = Mean
        // Parameter2 = Standard Deviation
        //
        // Lognormal
        // Parameter1 = Log Mean
        // Parameter2 = Log Standard Deviation
        //
        // Uniform
        // Parameter1 = Minimum
        // Parameter2 = Maximum
        //
        // Triangular / PERT
        // Parameter1 = Minimum
        // Parameter2 = Most Likely
        // Parameter3 = Maximum
        // ---------------------------------------------------------

        public double Parameter1 { get; set; }


        public double Parameter2 { get; set; }


        public double Parameter3 { get; set; }
    }


    // =============================================================
    // FORECAST
    // =============================================================

    public class ForecastDefinition
    {
        public string Name { get; set; } = "";


        public string SheetName { get; set; } = "";


        public string CellAddress { get; set; } = "";
    }


    // =============================================================
    // SIMULATION MODEL
    // =============================================================

    public static class SimulationModel
    {
        // ---------------------------------------------------------
        // ASSUMPTIONS
        // ---------------------------------------------------------

        public static List<AssumptionDefinition> Assumptions { get; }
            =
            new List<AssumptionDefinition>();


        // ---------------------------------------------------------
        // FORECASTS
        // ---------------------------------------------------------

        public static List<ForecastDefinition> Forecasts { get; }
            =
            new List<ForecastDefinition>();


        // =========================================================
        // CLEAR EVERYTHING
        // =========================================================

        public static void Clear()
        {
            Assumptions.Clear();

            Forecasts.Clear();
        }


        // =========================================================
        // CLEAR ASSUMPTIONS
        // =========================================================

        public static void ClearAssumptions()
        {
            Assumptions.Clear();
        }


        // =========================================================
        // CLEAR FORECASTS
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
                    return assumption;
                }
            }


            return null;
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
                    return forecast;
                }
            }


            return null;
        }
    }
}