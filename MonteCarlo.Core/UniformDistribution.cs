namespace MonteCarlo.Core
{
    public static class UniformDistribution
    {
        private static readonly Random _random = new();

        public static double Sample(
            double minimum,
            double maximum)
        {
            if (minimum >= maximum)
            {
                throw new ArgumentException(
                    "Minimum must be less than maximum.");
            }

            double u =
                _random.NextDouble();

            return
                minimum +
                u *
                (maximum - minimum);
        }
    }
}