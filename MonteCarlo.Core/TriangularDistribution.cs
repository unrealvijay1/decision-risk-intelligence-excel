namespace MonteCarlo.Core
{
    public static class TriangularDistribution
    {
        private static readonly Random _random = new();

        public static double Sample(
            double minimum,
            double mostLikely,
            double maximum)
        {
            if (minimum >= maximum)
                throw new ArgumentException(
                    "Minimum must be less than maximum.");

            if (mostLikely < minimum || mostLikely > maximum)
                throw new ArgumentException(
                    "Most likely value must be between minimum and maximum.");

            double u = _random.NextDouble();

            double c =
                (mostLikely - minimum) /
                (maximum - minimum);

            if (u < c)
            {
                return minimum +
                       Math.Sqrt(
                           u *
                           (maximum - minimum) *
                           (mostLikely - minimum));
            }
            else
            {
                return maximum -
                       Math.Sqrt(
                           (1.0 - u) *
                           (maximum - minimum) *
                           (maximum - mostLikely));
            }
        }
    }
}