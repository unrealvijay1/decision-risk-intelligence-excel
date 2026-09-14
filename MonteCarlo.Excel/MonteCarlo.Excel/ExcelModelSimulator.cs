using ExcelDna.Integration;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace MonteCarlo.Excel
{
    public static class ExcelModelSimulator
    {
        public static object[,] RunProjectModel(int trials)
        {
            if (trials <= 0)
            {
                throw new ArgumentException(
                    "Trials must be greater than zero.",
                    nameof(trials));
            }

            ExcelInterop.Application excelApp =
                (ExcelInterop.Application)ExcelDnaUtil.Application;

            ExcelInterop.Worksheet worksheet =
                (ExcelInterop.Worksheet)excelApp.ActiveSheet;

            ExcelInterop.Range durationCell =
                worksheet.Range["B2"];

            ExcelInterop.Range rateCell =
                worksheet.Range["B3"];

            ExcelInterop.Range outputCell =
                worksheet.Range["B5"];

            object originalDuration =
                durationCell.Value2;

            object originalRate =
                rateCell.Value2;

            bool originalScreenUpdating =
                excelApp.ScreenUpdating;

            bool originalEnableEvents =
                excelApp.EnableEvents;

            double[] results =
                new double[trials];

            try
            {
                excelApp.ScreenUpdating = false;
                excelApp.EnableEvents = false;

                for (int i = 0; i < trials; i++)
                {
                    double duration =
                        MonteCarlo.Core
                            .PertDistribution
                            .Sample(
                                80,
                                100,
                                160);

                    double dailyRate =
                        MonteCarlo.Core
                            .PertDistribution
                            .Sample(
                                5000,
                                6000,
                                8000);

                    durationCell.Value2 =
                        duration;

                    rateCell.Value2 =
                        dailyRate;

                    // Recalculate ONLY the output cell
                    outputCell.Calculate();

                    object outputValue =
                        outputCell.Value2;

                    if (outputValue == null)
                    {
                        throw new Exception(
                            "B5 returned no value during simulation.");
                    }

                    results[i] =
                        Convert.ToDouble(
                            outputValue);

                    // Show progress occasionally
                    if (i % 10 == 0)
                    {
                        excelApp.StatusBar =
                            $"Monte Carlo simulation: {i + 1} / {trials}";
                    }
                }
            }
            finally
            {
                // Restore inputs
                durationCell.Value2 =
                    originalDuration;

                rateCell.Value2 =
                    originalRate;

                outputCell.Calculate();

                excelApp.ScreenUpdating =
                    originalScreenUpdating;

                excelApp.EnableEvents =
                    originalEnableEvents;

                excelApp.StatusBar =
                    false;
            }

            Array.Sort(results);

            double mean =
                results.Average();

            double variance =
                results
                    .Select(x =>
                        Math.Pow(
                            x - mean,
                            2))
                    .Average();

            double stdDev =
                Math.Sqrt(variance);

            return new object[,]
            {
                { "Metric", "Value" },
                { "Trials", trials },
                { "Mean", mean },
                { "Std Dev", stdDev },
                { "Minimum", results.First() },
                { "Maximum", results.Last() },
                { "P10", Percentile(results, 0.10) },
                { "P50", Percentile(results, 0.50) },
                { "P80", Percentile(results, 0.80) },
                { "P90", Percentile(results, 0.90) }
            };
        }


        private static double Percentile(
            double[] sortedValues,
            double percentile)
        {
            double position =
                percentile *
                (sortedValues.Length - 1);

            int lower =
                (int)Math.Floor(position);

            int upper =
                (int)Math.Ceiling(position);

            if (lower == upper)
            {
                return sortedValues[lower];
            }

            double weight =
                position - lower;

            return
                sortedValues[lower] *
                (1.0 - weight)
                +
                sortedValues[upper] *
                weight;
        }
    }
}