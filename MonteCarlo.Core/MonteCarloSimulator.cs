namespace MonteCarlo.Core
{
    public static class MonteCarloSimulator
    {
        public static SimulationResult RunNormal(
            double mean,
            double standardDeviation,
            int trials)
        {
            if (trials <= 0)
                throw new ArgumentException(
                    "Trials must be greater than zero.",
                    nameof(trials));

            double[] samples = new double[trials];

            for (int i = 0; i < trials; i++)
            {
                samples[i] =
                    NormalDistribution.Sample(
                        mean,
                        standardDeviation);
            }

            return CalculateStatistics(samples);
        }


        public static SimulationResult RunTriangular(
            double minimum,
            double mostLikely,
            double maximum,
            int trials)
        {
            if (trials <= 0)
                throw new ArgumentException(
                    "Trials must be greater than zero.",
                    nameof(trials));

            double[] samples = new double[trials];

            for (int i = 0; i < trials; i++)
            {
                samples[i] =
                    TriangularDistribution.Sample(
                        minimum,
                        mostLikely,
                        maximum);
            }

            return CalculateStatistics(samples);
        }


        public static SimulationResult RunPert(
            double minimum,
            double mostLikely,
            double maximum,
            int trials)
        {
            if (trials <= 0)
                throw new ArgumentException(
                    "Trials must be greater than zero.",
                    nameof(trials));

            double[] samples = new double[trials];

            for (int i = 0; i < trials; i++)
            {
                samples[i] =
                    PertDistribution.Sample(
                        minimum,
                        mostLikely,
                        maximum);
            }

            return CalculateStatistics(samples);
        }


        private static SimulationResult CalculateStatistics(
            double[] samples)
        {
            Array.Sort(samples);

            double mean =
                samples.Average();

            double variance =
                samples
                    .Select(x =>
                        Math.Pow(x - mean, 2))
                    .Average();

            double standardDeviation =
                Math.Sqrt(variance);

            return new SimulationResult
            {
                Trials = samples.Length,

                Mean = mean,

                StandardDeviation =
                    standardDeviation,

                Minimum =
                    samples.First(),

                Maximum =
                    samples.Last(),

                P10 =
                    Percentile(
                        samples,
                        0.10),

                P50 =
                    Percentile(
                        samples,
                        0.50),

                P80 =
                    Percentile(
                        samples,
                        0.80),

                P90 =
                    Percentile(
                        samples,
                        0.90)
            };
        }


        private static double Percentile(
            double[] sortedSamples,
            double percentile)
        {
            if (sortedSamples == null ||
                sortedSamples.Length == 0)
            {
                throw new ArgumentException(
                    "Sample array cannot be empty.",
                    nameof(sortedSamples));
            }

            if (percentile < 0.0 ||
                percentile > 1.0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(percentile),
                    "Percentile must be between 0 and 1.");
            }

            double position =
                percentile *
                (sortedSamples.Length - 1);

            int lowerIndex =
                (int)Math.Floor(position);

            int upperIndex =
                (int)Math.Ceiling(position);

            if (lowerIndex == upperIndex)
            {
                return
                    sortedSamples[lowerIndex];
            }

            double weight =
                position - lowerIndex;

            return
                sortedSamples[lowerIndex] *
                (1.0 - weight)
                +
                sortedSamples[upperIndex] *
                weight;
        }
        public static SimulationResult RunProjectCost(
    double durationMin,
    double durationMostLikely,
    double durationMax,
    double rateMin,
    double rateMostLikely,
    double rateMax,
    int people,
    int trials)
        {
            if (people <= 0)
                throw new ArgumentException(
                    "People must be greater than zero.",
                    nameof(people));

            if (trials <= 0)
                throw new ArgumentException(
                    "Trials must be greater than zero.",
                    nameof(trials));

            double[] totalCosts = new double[trials];

            for (int i = 0; i < trials; i++)
            {
                double duration =
                    PertDistribution.Sample(
                        durationMin,
                        durationMostLikely,
                        durationMax);

                double dailyRate =
                    PertDistribution.Sample(
                        rateMin,
                        rateMostLikely,
                        rateMax);

                totalCosts[i] =
                    duration *
                    dailyRate *
                    people;
            }

            return CalculateStatistics(totalCosts);
        }
    }
}