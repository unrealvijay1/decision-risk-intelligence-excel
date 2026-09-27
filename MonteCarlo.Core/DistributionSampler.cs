using System;

namespace MonteCarlo.Core
{
    public enum DistributionKind
    {
        Normal,
        Triangular,
        Pert,
        Uniform,
        Lognormal,
        Beta
    }


    public static class DistributionSampler
    {
        // =========================================================
        // SAMPLE
        // =========================================================

        public static double Sample(
            DistributionKind distribution,
            double parameter1,
            double parameter2,
            double parameter3 = 0,
            double parameter4 = 0, Random? random = null)
        {
            switch (distribution)
            {
                case DistributionKind.Normal:

                    return
                        NormalDistribution.Sample(
                            parameter1,
                            parameter2, random: random);


                case DistributionKind.Triangular:

                    return
                        TriangularDistribution.Sample(
                            parameter1,
                            parameter2,
                            parameter3, random: random);


                case DistributionKind.Pert:

                    return
                        PertDistribution.Sample(
                            parameter1,
                            parameter2,
                            parameter3, random: random);


                case DistributionKind.Uniform:

                    return
                        UniformDistribution.Sample(
                            parameter1,
                            parameter2, random: random);


                case DistributionKind.Lognormal:

                    return
                        LognormalDistribution.Sample(
                            parameter1,
                            parameter2, random: random);


                case DistributionKind.Beta:

                    return
                        BetaDistribution.Sample(
                            parameter1,
                            parameter2,
                            parameter3,
                            parameter4, random: random);


                default:

                    throw new ArgumentOutOfRangeException(
                        nameof(distribution),
                        distribution,
                        "Unsupported distribution.");
            }
        }
    }
}