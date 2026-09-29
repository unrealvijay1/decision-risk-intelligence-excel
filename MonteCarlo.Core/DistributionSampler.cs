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
        Beta,
        Exponential, Poisson, Binomial, Discrete, Weibull, Gamma, Bernoulli, TruncatedNormal
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
            double parameter4 = 0, Random? random = null, DiscreteTable? table = null)
        {
            switch (distribution)
            {
                case DistributionKind.Exponential or DistributionKind.Poisson or DistributionKind.Binomial or
                    DistributionKind.Discrete or DistributionKind.Weibull or DistributionKind.Gamma or
                    DistributionKind.Bernoulli or DistributionKind.TruncatedNormal:
                    return ExtendedDistributions.Sample(distribution, parameter1, parameter2, parameter3, parameter4, random ?? Random.Shared, table);
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
