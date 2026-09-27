namespace MonteCarlo.Core
{
    public static class PertDistribution
    {
        private static readonly Random _random = new();

        public static double Sample(
            double minimum,
            double mostLikely,
            double maximum,
            double lambda = 4.0, Random? random = null)
        {
            if (minimum >= maximum)
                throw new ArgumentException(
                    "Minimum must be less than maximum.");

            if (mostLikely < minimum || mostLikely > maximum)
                throw new ArgumentException(
                    "Most likely value must be between minimum and maximum.");

            if (lambda <= 0)
                throw new ArgumentException(
                    "Lambda must be greater than zero.");

            double range = maximum - minimum;

            double alpha =
                1.0 +
                lambda *
                (mostLikely - minimum) /
                range;

            double beta =
                1.0 +
                lambda *
                (maximum - mostLikely) /
                range;

            double betaSample =
                SampleBeta(alpha, beta, random: random);

            return minimum +
                   betaSample * range;
        }


        private static double SampleBeta(
            double alpha,
            double beta, Random? random = null)
        {
            double x =
                SampleGamma(alpha, random: random);

            double y =
                SampleGamma(beta, random: random);

            return x / (x + y);
        }


        private static double SampleGamma(
            double shape, Random? random = null)
        {
            if (shape < 1.0)
            {
                double u =
                    (random ?? _random).NextDouble();

                return SampleGamma(shape + 1.0, random: random) *
                       Math.Pow(
                           u,
                           1.0 / shape);
            }

            double d =
                shape - 1.0 / 3.0;

            double c =
                1.0 /
                Math.Sqrt(9.0 * d);

            while (true)
            {
                double x;
                double v;

                do
                {
                    x = SampleStandardNormal(random: random);

                    v = 1.0 + c * x;
                }
                while (v <= 0.0);

                v = v * v * v;

                double u =
                    (random ?? _random).NextDouble();

                if (u <
                    1.0 -
                    0.0331 *
                    x * x *
                    x * x)
                {
                    return d * v;
                }

                if (Math.Log(u) <
                    0.5 * x * x +
                    d *
                    (1.0 - v + Math.Log(v)))
                {
                    return d * v;
                }
            }
        }


        private static double SampleStandardNormal(Random? random = null)
        {
            double u1 =
                1.0 - (random ?? _random).NextDouble();

            double u2 =
                1.0 - (random ?? _random).NextDouble();

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