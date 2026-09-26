namespace MonteCarlo.Core;

/// <summary>One right-continuous ECDF jump, including all samples equal to X.</summary>
public readonly record struct CumulativePoint(double X, double ProbabilityBefore, double Probability);

public sealed class CumulativeProbabilityResult
{
    public IReadOnlyList<CumulativePoint> Points { get; init; } = Array.Empty<CumulativePoint>();
    public string? ValidationMessage { get; init; }
}

public static class CumulativeProbability
{
    /// <summary>Sorts a copy. Invalid samples invalidate the entire curve, never its denominator.</summary>
    public static CumulativeProbabilityResult Build(IReadOnlyList<double> samples)
    {
        ArgumentNullException.ThrowIfNull(samples);
        if (samples.Count == 0) return new();
        if (samples.Any(value => !double.IsFinite(value)))
            return new() { ValidationMessage = "Cumulative probability unavailable: results contain non-finite values." };

        double[] sorted = samples.OrderBy(value => value).ToArray();
        var points = new List<CumulativePoint>();
        int i = 0;
        while (i < sorted.Length)
        {
            int start = i;
            double value = sorted[i++];
            while (i < sorted.Length && sorted[i] == value) i++;
            points.Add(new(value, (double)start / sorted.Length, (double)i / sorted.Length));
        }
        return new() { Points = points.AsReadOnly() };
    }
}
