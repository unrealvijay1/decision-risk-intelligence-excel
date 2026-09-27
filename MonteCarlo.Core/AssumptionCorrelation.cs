namespace MonteCarlo.Core;

public sealed record AssumptionCorrelation(string AssumptionA, string AssumptionB, double Correlation);

/// <summary>Gaussian dependence, not calibrated Pearson correlation of transformed marginals.</summary>
public sealed class GaussianDependence
{
    private readonly double[,] lower;
    private readonly double[] independent;
    public IReadOnlyList<string> Ids { get; }
    public GaussianDependence(IReadOnlyList<string> available, IReadOnlyList<AssumptionCorrelation> relationships)
    {
        var known = available.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var pairs = new HashSet<(string, string)>();
        foreach (var r in relationships)
        {
            if (r == null || string.IsNullOrWhiteSpace(r.AssumptionA) || string.IsNullOrWhiteSpace(r.AssumptionB) ||
                !known.Contains(r.AssumptionA) || !known.Contains(r.AssumptionB))
                throw new ArgumentException("A correlated assumption is missing. Open Assumption Correlations and remove or replace its relationship.");
            string a = r.AssumptionA.ToUpperInvariant(), b = r.AssumptionB.ToUpperInvariant();
            if (a == b) throw new ArgumentException("Choose two different assumptions for a correlation.");
            if (!pairs.Add(string.CompareOrdinal(a, b) < 0 ? (a, b) : (b, a)))
                throw new ArgumentException("Each pair of assumptions can have only one correlation (including reversed pairs).");
            if (!double.IsFinite(r.Correlation) || Math.Abs(r.Correlation) > 1)
                throw new ArgumentException("Enter a finite correlation between -1 and +1.");
        }
        var participating = relationships.Where(r => r.Correlation != 0)
            .SelectMany(r => new[] { r.AssumptionA, r.AssumptionB }).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Ids = available.Where(participating.Contains).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        int n = Ids.Count;
        var index = Ids.Select((id, i) => (id, i)).ToDictionary(x => x.id, x => x.i, StringComparer.OrdinalIgnoreCase);
        var matrix = new double[n, n];
        for (int i = 0; i < n; i++) matrix[i, i] = 1;
        foreach (var r in relationships.Where(r => r.Correlation != 0))
            matrix[index[r.AssumptionA], index[r.AssumptionB]] = matrix[index[r.AssumptionB], index[r.AssumptionA]] = r.Correlation;
        lower = new double[n, n]; independent = new double[n];
        // Semidefinite Cholesky permits exact +/-1 and singular valid matrices. Never repairs coefficients.
        const double tolerance = 1e-12;
        for (int i = 0; i < n; i++)
        for (int j = 0; j <= i; j++)
        {
            double residual = matrix[i, j];
            for (int k = 0; k < j; k++) residual -= lower[i, k] * lower[j, k];
            if (i == j)
            {
                if (residual < -tolerance) throw Inconsistent();
                lower[i, j] = Math.Sqrt(Math.Max(0, residual));
            }
            else if (lower[j, j] > 0) lower[i, j] = residual / lower[j, j];
            else if (Math.Abs(residual) > tolerance) throw Inconsistent();
        }
    }
    private static ArgumentException Inconsistent() => new("The combination of correlations is inconsistent. Adjust the relationships in Assumption Correlations before running.");
    public void NextGaussian(Random random, double[] destination)
    {
        if (destination.Length != Ids.Count) throw new ArgumentException("Incorrect correlation vector size.");
        for (int i = 0; i < independent.Length; i++) independent[i] = NormalDistribution.Sample(0, 1, random);
        for (int i = 0; i < independent.Length; i++)
        {
            double value = 0;
            for (int j = 0; j <= i; j++) value += lower[i, j] * independent[j];
            destination[i] = value;
        }
    }
}
