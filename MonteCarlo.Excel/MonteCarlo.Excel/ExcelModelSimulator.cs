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

            var assumptions = new AssumptionDefinition[]
            {
                new() { Name = "Duration", SheetName = worksheet.Name, CellAddress = "B2",
                    Distribution = DistributionType.Pert, Parameter1 = 80, Parameter2 = 100, Parameter3 = 160 },
                new() { Name = "Rate", SheetName = worksheet.Name, CellAddress = "B3",
                    Distribution = DistributionType.Pert, Parameter1 = 5000, Parameter2 = 6000, Parameter3 = 8000 }
            };
            var forecast = new ForecastDefinition { SheetName = worksheet.Name, CellAddress = "B5" };
            var workbook = new ExcelSimulationWorkbook(excelApp, worksheet.Parent,
                () => worksheet.Range["B5"].Calculate());
            var outcome = SimulationExecution.Run(trials, assumptions, new[] { forecast }, workbook);
            if (!outcome.Succeeded) throw new SimulationRunException(outcome);
            double[] results = outcome.Result!.ForecastResults[0].Values;

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