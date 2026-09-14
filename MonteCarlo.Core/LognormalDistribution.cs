namespace MonteCarlo.Core
{
    public static class LognormalDistribution
    {
        private static readonly Random _random = new();

        public static double Sample(
            double logMean,
            double logStandardDeviation)
        {
            if (logStandardDeviation <= 0)
            {
                throw new ArgumentException(
                    "Log standard deviation must be greater than zero.");
            }

            double normal =
                SampleStandardNormal();

            return
                Math.Exp(
                    logMean +
                    logStandardDeviation *
                    normal);
        }

        private static double SampleStandardNormal()
        {
            double u1 =
                1.0 -
                _random.NextDouble();

            double u2 =
                1.0 -
                _random.NextDouble();

            return
                Math.Sqrt(
                    -2.0 *
                    Math.Log(u1))
                *
                Math.Cos(
                    2.0 *
                    Math.PI *
                    u2);
        }
    }
}