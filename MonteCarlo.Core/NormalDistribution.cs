namespace MonteCarlo.Core
{
    public static class NormalDistribution
    {
        private static readonly Random _random = new();

        public static double Sample(double mean, double standardDeviation)
        {
            if (standardDeviation <= 0)
                throw new ArgumentException(
                    "Standard deviation must be greater than zero.",
                    nameof(standardDeviation));

            double u1 = 1.0 - _random.NextDouble();
            double u2 = 1.0 - _random.NextDouble();

            double standardNormal =
                Math.Sqrt(-2.0 * Math.Log(u1)) *
                Math.Cos(2.0 * Math.PI * u2);

            return mean + standardDeviation * standardNormal;
        }
    }
}