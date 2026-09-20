namespace MonteCarlo.Core;

/// <summary>
/// Empirical comparisons retain IEEE double semantics: NaN never matches,
/// infinities compare normally, and every sample remains in the denominator.
/// An empty sample set returns NaN, matching zero divided by zero.
/// </summary>
public static class EmpiricalProbability
{
    public static double LessThanOrEqual(IReadOnlyList<double> samples, double target) =>
        Count(samples, value => value <= target);

    public static double GreaterThan(IReadOnlyList<double> samples, double target) =>
        Count(samples, value => value > target);

    public static double GreaterThanOrEqual(IReadOnlyList<double> samples, double target) =>
        Count(samples, value => value >= target);

    private static double Count(IReadOnlyList<double> samples, Func<double, bool> matches)
    {
        ArgumentNullException.ThrowIfNull(samples);
        int count = 0;
        foreach (double value in samples)
            if (matches(value)) count++;
        return (double)count / samples.Count;
    }
}
