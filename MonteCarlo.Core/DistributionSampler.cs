using System;

namespace MonteCarlo.Core
{
    public enum DistributionKind
    {
        Normal,
        Triangular,
        Pert,
        Uniform,
        Lognormal
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
            double parameter3 = 0)
        {
            switch (distribution)
            {
                case DistributionKind.Normal:

                    return
                        NormalDistribution.Sample(
                            parameter1,
                            parameter2);


                case DistributionKind.Triangular:

                    return
                        TriangularDistribution.Sample(
                            parameter1,
                            parameter2,
                            parameter3);


                case DistributionKind.Pert:

                    return
                        PertDistribution.Sample(
                            parameter1,
                            parameter2,
                            parameter3);


                case DistributionKind.Uniform:

                    return
                        UniformDistribution.Sample(
                            parameter1,
                            parameter2);


                case DistributionKind.Lognormal:

                    return
                        LognormalDistribution.Sample(
                            parameter1,
                            parameter2);


                default:

                    throw new ArgumentOutOfRangeException(
                        nameof(distribution),
                        distribution,
                        "Unsupported distribution.");
            }
        }
    }
}