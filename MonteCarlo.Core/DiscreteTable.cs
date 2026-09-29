namespace MonteCarlo.Core;

public sealed record DiscreteOutcome(double Outcome, double Probability);

/// <summary>Immutable, sorted probability table. Totals within 1e-10 of one are normalized.</summary>
public sealed class DiscreteTable
{
    public const double TotalTolerance = 1e-10;
    public IReadOnlyList<DiscreteOutcome> Outcomes { get; }
    private readonly double[] cumulative;
    public DiscreteTable(IEnumerable<DiscreteOutcome> outcomes)
    {
        ArgumentNullException.ThrowIfNull(outcomes);
        var rows = outcomes.Take(101).ToArray();
        if (rows.Length < 2 || rows.Length > 100) throw new ArgumentException("Enter between 2 and 100 outcome/probability rows.");
        if (rows.Any(r => r == null || !double.IsFinite(r.Outcome) || !double.IsFinite(r.Probability) || r.Probability < 0))
            throw new ArgumentException("Outcomes must be finite numbers; probabilities must be finite and nonnegative.");
        if (rows.Select(r => r.Outcome).Distinct().Count() != rows.Length) throw new ArgumentException("Each outcome must be unique.");
        double sum = rows.Sum(r => r.Probability);
        if (Math.Abs(sum - 1) > TotalTolerance) throw new ArgumentException($"Probabilities total {sum:G17}; they must sum to 1 (tolerance {TotalTolerance:G}).");
        rows = rows.OrderBy(r => r.Outcome).Select(r => r with { Probability = r.Probability / sum }).ToArray();
        Outcomes = Array.AsReadOnly(rows);
        cumulative = new double[rows.Length];
        double total = 0;
        for (int i = 0; i < rows.Length; i++) cumulative[i] = Math.Min(1, total += rows[i].Probability);
        cumulative[^1] = 1;
    }
    public double Mass(double x) => Outcomes.FirstOrDefault(r => r.Outcome == x)?.Probability ?? 0;
    public double Cdf(double x)
    {
        for (int i = Outcomes.Count - 1; i >= 0; i--) if (x >= Outcomes[i].Outcome) return cumulative[i];
        return 0;
    }
    public double Quantile(double probability)
    {
        if (!double.IsFinite(probability) || probability < 0 || probability > 1) throw new ArgumentException("Probability must be between zero and one.");
        if (probability == 1) return Outcomes.Last(r => r.Probability > 0).Outcome;
        for (int i = 0; i < Outcomes.Count; i++)
            if (Outcomes[i].Probability > 0 && probability <= cumulative[i]) return Outcomes[i].Outcome;
        return Outcomes.Last(r => r.Probability > 0).Outcome;
    }
}
