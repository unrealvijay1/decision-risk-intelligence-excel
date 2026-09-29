namespace MonteCarlo.Core;

public static class DistributionCatalog
{
    public static string Name(DistributionKind kind) => kind switch { DistributionKind.Pert => "PERT", DistributionKind.TruncatedNormal => "Truncated Normal", _ => kind.ToString() };
    public static string[] Parameters(DistributionKind kind) => kind switch
    {
        DistributionKind.Normal => ["Mean", "Standard deviation"],
        DistributionKind.Lognormal => ["Log mean", "Log standard deviation"],
        DistributionKind.Uniform => ["Minimum", "Maximum"],
        DistributionKind.Triangular or DistributionKind.Pert => ["Minimum", "Most likely", "Maximum"],
        DistributionKind.Beta => ["Minimum", "Maximum", "Alpha", "Beta"],
        DistributionKind.Exponential => ["Mean"], DistributionKind.Poisson => ["Lambda"],
        DistributionKind.Binomial => ["Trials (integer)", "Probability"],
        DistributionKind.Weibull => ["Shape", "Scale"], DistributionKind.Gamma => ["Shape", "Scale"],
        DistributionKind.Bernoulli => ["Probability"],
        DistributionKind.TruncatedNormal => ["Mean", "Standard deviation", "Lower bound", "Upper bound"],
        _ => []
    };
    public static double[] Defaults(DistributionKind kind) => kind switch
    {
        DistributionKind.Normal => [0, 1, 0, 0], DistributionKind.Lognormal => [0, .25, 0, 0],
        DistributionKind.Uniform => [0, 1, 0, 0], DistributionKind.Triangular or DistributionKind.Pert => [0, .5, 1, 0],
        DistributionKind.Beta => [0, 1, 2, 2], DistributionKind.Exponential => [1, 0, 0, 0],
        DistributionKind.Poisson => [5, 0, 0, 0], DistributionKind.Binomial => [10, .5, 0, 0],
        DistributionKind.Weibull or DistributionKind.Gamma => [2, 1, 0, 0],
        DistributionKind.Bernoulli => [.5, 0, 0, 0], DistributionKind.TruncatedNormal => [0, 1, -2, 2], _ => [0, 0, 0, 0]
    };
    public static bool IsDiscrete(DistributionKind kind) => kind is DistributionKind.Discrete or DistributionKind.Poisson or DistributionKind.Binomial or DistributionKind.Bernoulli;
}
