using System;

namespace MonteCarlo.Core
{
    public static class BetaDistribution
    {
        private static readonly Random Random =
            new Random();


        // =========================================================
        // SAMPLE STANDARD BETA
        //
        // Alpha > 0
        // Beta  > 0
        //
        // Output is between 0 and 1.
        // =========================================================

        public static double Sample(
            double alpha,
            double beta, Random? random = null)
        {
            if (alpha <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(alpha),
                    "Alpha must be greater than zero.");
            }


            if (beta <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(beta),
                    "Beta must be greater than zero.");
            }


            double x =
                SampleGamma(
                    alpha, random: random);


            double y =
                SampleGamma(
                    beta, random: random);


            return
                x /
                (x + y);
        }


        // =========================================================
        // SAMPLE SCALED BETA
        //
        // Maps Beta distribution from [0,1] to [min,max].
        // =========================================================

        public static double Sample(
            double minimum,
            double maximum,
            double alpha,
            double beta, Random? random = null)
        {
            if (maximum <= minimum)
            {
                throw new ArgumentException(
                    "Maximum must be greater than minimum.");
            }


            double standardBeta =
                Sample(
                    alpha,
                    beta, random: random);


            return
                minimum
                +
                standardBeta
                *
                (
                    maximum -
                    minimum
                );
        }


        // =========================================================
        // GAMMA SAMPLER
        //
        // Used internally to generate Beta samples.
        // =========================================================

        private static double SampleGamma(
            double shape, Random? random = null)
        {
            if (shape <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(shape));
            }


            // -----------------------------------------------------
            // Shape < 1
            // -----------------------------------------------------

            if (shape < 1.0)
            {
                double u =
                    NextUniform(random: random);


                return
                    SampleGamma(
                        shape + 1.0, random: random)
                    *
                    Math.Pow(
                        u,
                        1.0 / shape);
            }


            // -----------------------------------------------------
            // Marsaglia and Tsang method
            // -----------------------------------------------------

            double d =
                shape -
                1.0 / 3.0;


            double c =
                1.0 /
                Math.Sqrt(
                    9.0 * d);


            while (true)
            {
                double x =
                    SampleStandardNormal(random: random);


                double v =
                    1.0 +
                    c * x;


                if (v <= 0)
                {
                    continue;
                }


                v =
                    v *
                    v *
                    v;


                double u =
                    NextUniform(random: random);


                if (
                    u <
                    1.0 -
                    0.0331 *
                    x *
                    x *
                    x *
                    x)
                {
                    return
                        d * v;
                }


                if (
                    Math.Log(u)
                    <
                    0.5 *
                    x *
                    x
                    +
                    d *
                    (
                        1.0 -
                        v +
                        Math.Log(v)
                    ))
                {
                    return
                        d * v;
                }
            }
        }


        // =========================================================
        // STANDARD NORMAL
        // =========================================================

        private static double SampleStandardNormal(Random? random = null)
        {
            double u1 =
                NextUniform(random: random);


            double u2 =
                NextUniform(random: random);


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


        // =========================================================
        // UNIFORM (0,1)
        // =========================================================

        private static double NextUniform(Random? random = null)
        {
            double value;


            do
            {
                value =
                    (random ?? Random).NextDouble();
            }
            while (value <= 0.0);


            return value;
        }
    }
}