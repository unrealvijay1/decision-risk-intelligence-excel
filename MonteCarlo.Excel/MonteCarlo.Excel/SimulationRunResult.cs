using System;
using System.Collections.Generic;
using System.Linq;

namespace MonteCarlo.Excel
{
    // =============================================================
    // ONE FORECAST RESULT
    // =============================================================

    public class ForecastRunResult
    {
        public ForecastDefinition Forecast { get; }

        public double[] Values { get; }

        public int Trials =>
            Values.Length;

        public double Mean { get; }

        public double StandardDeviation { get; }

        public double Minimum { get; }

        public double Maximum { get; }

        public double P10 { get; }

        public double P50 { get; }

        public double P80 { get; }

        public double P90 { get; }

        public List<SensitivityResult> Sensitivities { get; }


        public ForecastRunResult(
            ForecastDefinition forecast,
            double[] values,
            Dictionary<string, double[]> assumptionSamples)
        {
            Forecast =
                forecast;

            if (values == null)
            {
                throw new ArgumentNullException(
                    nameof(values));
            }


            if (values.Length == 0)
            {
                throw new ArgumentException(
                    "Forecast results cannot be empty.",
                    nameof(values));
            }


            Values =
                values.ToArray();


            // -----------------------------------------------------
            // BASIC STATISTICS
            // -----------------------------------------------------

            Mean =
                Values.Average();


            Minimum =
                Values.Min();


            Maximum =
                Values.Max();


            double variance =
                Values
                    .Select(
                        value =>
                            Math.Pow(
                                value - Mean,
                                2))
                    .Average();


            StandardDeviation =
                Math.Sqrt(
                    variance);


            // -----------------------------------------------------
            // PERCENTILES
            // -----------------------------------------------------

            double[] sortedValues =
                Values
                    .OrderBy(
                        value => value)
                    .ToArray();


            P10 =
                Percentile(
                    sortedValues,
                    0.10);


            P50 =
                Percentile(
                    sortedValues,
                    0.50);


            P80 =
                Percentile(
                    sortedValues,
                    0.80);


            P90 =
                Percentile(
                    sortedValues,
                    0.90);


            // -----------------------------------------------------
            // SENSITIVITY
            // -----------------------------------------------------

            Sensitivities =
                new List<SensitivityResult>();


            foreach (
                KeyValuePair<string, double[]> assumption
                in assumptionSamples)
            {
                if (assumption.Value == null)
                {
                    continue;
                }


                if (assumption.Value.Length !=
                    Values.Length)
                {
                    continue;
                }


                double correlation =
                    SpearmanCorrelation(
                        assumption.Value,
                        Values);


                Sensitivities.Add(
                    new SensitivityResult
                    {
                        Name =
                            assumption.Key,

                        Correlation =
                            correlation
                    });
            }


            Sensitivities.Sort(
                (a, b) =>
                    Math.Abs(
                        b.Correlation)
                    .CompareTo(
                        Math.Abs(
                            a.Correlation)));
        }


        // =========================================================
        // PROBABILITY <= TARGET
        // =========================================================

        public double ProbabilityLessThanOrEqual(
            double target)
        {
            int count =
                Values.Count(
                    value =>
                        value <= target);


            return
                (double)count /
                Values.Length;
        }


        // =========================================================
        // PROBABILITY > TARGET
        // =========================================================

        public double ProbabilityGreaterThan(
            double target)
        {
            int count =
                Values.Count(
                    value =>
                        value > target);


            return
                (double)count /
                Values.Length;
        }


        // =========================================================
        // SPEARMAN CORRELATION
        // =========================================================

        private static double SpearmanCorrelation(
            double[] x,
            double[] y)
        {
            if (x.Length != y.Length ||
                x.Length < 2)
            {
                return 0;
            }


            double[] rankX =
                GetRanks(
                    x);


            double[] rankY =
                GetRanks(
                    y);


            return
                PearsonCorrelation(
                    rankX,
                    rankY);
        }


        // =========================================================
        // RANKING
        // =========================================================

        private static double[] GetRanks(
            double[] values)
        {
            int count =
                values.Length;


            var indexedValues =
                values
                    .Select(
                        (value, index) =>
                            new
                            {
                                Value =
                                    value,

                                OriginalIndex =
                                    index
                            })
                    .OrderBy(
                        item =>
                            item.Value)
                    .ToArray();


            double[] ranks =
                new double[count];


            int position =
                0;


            while (position < count)
            {
                int start =
                    position;


                double currentValue =
                    indexedValues[
                        position]
                        .Value;


                while (
                    position + 1 < count
                    &&
                    indexedValues[
                        position + 1]
                        .Value ==
                    currentValue)
                {
                    position++;
                }


                int end =
                    position;


                double averageRank =
                    (
                        (start + 1)
                        +
                        (end + 1)
                    )
                    /
                    2.0;


                for (
                    int i = start;
                    i <= end;
                    i++)
                {
                    int originalIndex =
                        indexedValues[i]
                            .OriginalIndex;


                    ranks[
                        originalIndex] =
                        averageRank;
                }


                position++;
            }


            return ranks;
        }


        // =========================================================
        // PEARSON ON RANKS
        // =========================================================

        private static double PearsonCorrelation(
            double[] x,
            double[] y)
        {
            double meanX =
                x.Average();


            double meanY =
                y.Average();


            double numerator =
                0;


            double sumSquareX =
                0;


            double sumSquareY =
                0;


            for (
                int i = 0;
                i < x.Length;
                i++)
            {
                double dx =
                    x[i] -
                    meanX;


                double dy =
                    y[i] -
                    meanY;


                numerator +=
                    dx *
                    dy;


                sumSquareX +=
                    dx *
                    dx;


                sumSquareY +=
                    dy *
                    dy;
            }


            double denominator =
                Math.Sqrt(
                    sumSquareX *
                    sumSquareY);


            if (denominator == 0)
            {
                return 0;
            }


            double correlation =
                numerator /
                denominator;


            if (correlation > 1)
            {
                correlation = 1;
            }


            if (correlation < -1)
            {
                correlation = -1;
            }


            return correlation;
        }


        // =========================================================
        // PERCENTILE
        // =========================================================

        private static double Percentile(
            double[] sortedValues,
            double percentile)
        {
            if (sortedValues.Length == 0)
            {
                return 0;
            }


            if (sortedValues.Length == 1)
            {
                return
                    sortedValues[0];
            }


            if (percentile <= 0)
            {
                return
                    sortedValues[0];
            }


            if (percentile >= 1)
            {
                return
                    sortedValues[
                        sortedValues.Length - 1];
            }


            double position =
                percentile *
                (sortedValues.Length - 1);


            int lowerIndex =
                (int)Math.Floor(
                    position);


            int upperIndex =
                (int)Math.Ceiling(
                    position);


            if (lowerIndex == upperIndex)
            {
                return
                    sortedValues[
                        lowerIndex];
            }


            double fraction =
                position -
                lowerIndex;


            double lowerValue =
                sortedValues[
                    lowerIndex];


            double upperValue =
                sortedValues[
                    upperIndex];


            return
                lowerValue
                +
                fraction *
                (
                    upperValue -
                    lowerValue
                );
        }
    }


    // =============================================================
    // COMPLETE SIMULATION RUN
    // =============================================================

    public class SimulationRunResult
    {
        public int Trials { get; }

        public List<ForecastRunResult> ForecastResults { get; }


        public SimulationRunResult(
            int trials,
            Dictionary<
                ForecastDefinition,
                double[]> forecastSamples,
            Dictionary<
                string,
                double[]> assumptionSamples)
        {
            if (trials <= 0)
            {
                throw new ArgumentException(
                    "Trials must be greater than zero.",
                    nameof(trials));
            }


            Trials =
                trials;


            ForecastResults =
                new List<ForecastRunResult>();


            foreach (
                KeyValuePair<
                    ForecastDefinition,
                    double[]> forecast
                in forecastSamples)
            {
                ForecastResults.Add(
                    new ForecastRunResult(
                        forecast.Key,
                        forecast.Value,
                        assumptionSamples));
            }
        }


        // =========================================================
        // FIND RESULT FOR A FORECAST
        // =========================================================

        public ForecastRunResult?
            FindForecastResult(
                string sheetName,
                string cellAddress)
        {
            return
                ForecastResults
                    .FirstOrDefault(
                        result =>
                            string.Equals(
                                result.Forecast.SheetName,
                                sheetName,
                                StringComparison.OrdinalIgnoreCase)
                            &&
                            string.Equals(
                                result.Forecast.CellAddress,
                                cellAddress,
                                StringComparison.OrdinalIgnoreCase));
        }
    }
}