namespace MonteCarlo.Core;

public readonly record struct DensityPoint(double X, double Density);

public sealed class DistributionPreviewResult
{
    public string? ValidationMessage { get; init; }
    public bool IsValid => ValidationMessage == null;
    public double MinimumX { get; init; }
    public double MaximumX { get; init; }
    public IReadOnlyList<DensityPoint> Points { get; init; } = Array.Empty<DensityPoint>();
}

/// <summary>Deterministic analytical densities for assumption previews. No Excel state is used.</summary>
public static class DistributionPreview
{
    private const int PointCount = 1025;
    private const double SqrtTwoPi = 2.5066282746310005;
    private static readonly double[] Lanczos =
    {
        0.99999999999980993, 676.5203681218851, -1259.1392167224028,
        771.32342877765313, -176.61502916214059, 12.507343278686905,
        -0.13857109526572012, 9.984369578019572e-6, 1.5056327351493116e-7
    };

    public static DistributionPreviewResult Generate(
        DistributionKind kind, double p1, double p2, double p3 = 0, double p4 = 0)
    {
        string? error = Validate(kind, p1, p2, p3, p4);
        if (error != null) return Invalid(error);

        double min;
        double max;
        switch (kind)
        {
            case DistributionKind.Normal:
                min = p1 - 4 * p2;
                max = p1 + 4 * p2;
                break;
            case DistributionKind.Lognormal:
                min = Math.Exp(p1 - 4 * p2);
                max = Math.Exp(p1 + 4 * p2);
                break;
            case DistributionKind.Triangular:
            case DistributionKind.Pert:
                min = p1;
                max = p3;
                break;
            default:
                min = p1;
                max = p2;
                break;
        }

        double span = max - min;
        if (!double.IsFinite(min) || !double.IsFinite(max) ||
            !double.IsFinite(span) || span <= 0)
            return Invalid("Parameters do not produce a finite plotting range.");

        var points = new DensityPoint[PointCount];
        for (int i = 0; i < PointCount; i++)
        {
            double fraction = (double)i / (PointCount - 1);
            // Beta can have infinite endpoint density when a shape is below one.
            // Evaluate just inside support without altering the analytical value.
            if (kind == DistributionKind.Beta)
                fraction = Math.Clamp(fraction, 1e-6, 1 - 1e-6);
            double x = min + span * fraction;
            double density = Density(kind, x, p1, p2, p3, p4);
            if (!double.IsFinite(x) || !double.IsFinite(density) || density < 0)
                return Invalid("Parameters produce density values outside the finite numeric range.");
            points[i] = new DensityPoint(x, density);
        }

        return new DistributionPreviewResult
        {
            MinimumX = min,
            MaximumX = max,
            Points = points
        };
    }

    public static string? Validate(DistributionKind kind, double p1, double p2,
        double p3 = 0, double p4 = 0)
    {
        if (!double.IsFinite(p1) || !double.IsFinite(p2) ||
            !double.IsFinite(p3) || !double.IsFinite(p4))
            return "Enter finite numeric parameters.";

        return kind switch
        {
            DistributionKind.Normal when p2 <= 0 => "Standard deviation must be greater than zero.",
            DistributionKind.Lognormal when p2 <= 0 => "Log standard deviation must be greater than zero.",
            DistributionKind.Uniform or DistributionKind.Beta when p1 >= p2 =>
                "Minimum must be less than maximum.",
            DistributionKind.Triangular or DistributionKind.Pert when p1 >= p3 =>
                "Minimum must be less than maximum.",
            DistributionKind.Triangular or DistributionKind.Pert when p2 < p1 || p2 > p3 =>
                "Most likely value must be between minimum and maximum.",
            DistributionKind.Beta when p3 <= 0 || p4 <= 0 =>
                "Alpha and Beta must be greater than zero.",
            DistributionKind.Normal or DistributionKind.Lognormal or
            DistributionKind.Uniform or DistributionKind.Triangular or
            DistributionKind.Pert or DistributionKind.Beta => null,
            _ => "Unsupported distribution."
        };
    }

    private static DistributionPreviewResult Invalid(string message) =>
        new() { ValidationMessage = message };

    private static double Density(DistributionKind kind, double x,
        double p1, double p2, double p3, double p4)
    {
        switch (kind)
        {
            case DistributionKind.Normal:
                double z = (x - p1) / p2;
                return Math.Exp(-0.5 * z * z) / (p2 * SqrtTwoPi);
            case DistributionKind.Lognormal:
                if (x <= 0) return 0;
                double logZ = (Math.Log(x) - p1) / p2;
                return Math.Exp(-0.5 * logZ * logZ) / (x * p2 * SqrtTwoPi);
            case DistributionKind.Uniform:
                return 1 / (p2 - p1);
            case DistributionKind.Triangular:
                if (x < p1 || x > p3) return 0;
                if (x == p2) return 2 / (p3 - p1);
                return x < p2
                    ? 2 * (x - p1) / ((p3 - p1) * (p2 - p1))
                    : 2 * (p3 - x) / ((p3 - p1) * (p3 - p2));
            case DistributionKind.Pert:
                double a = 1 + 4 * (p2 - p1) / (p3 - p1);
                double b = 1 + 4 * (p3 - p2) / (p3 - p1);
                return ScaledBetaDensity(x, p1, p3, a, b);
            case DistributionKind.Beta:
                return ScaledBetaDensity(x, p1, p2, p3, p4);
            default:
                return 0;
        }
    }

    private static double ScaledBetaDensity(double x, double min, double max,
        double alpha, double beta)
    {
        double span = max - min;
        double u = (x - min) / span;
        if (u <= 0) return alpha == 1 ? beta / span : 0;
        if (u >= 1) return beta == 1 ? alpha / span : 0;
        double logDensity = (alpha - 1) * Math.Log(u) +
            (beta - 1) * Math.Log(1 - u) -
            (LogGamma(alpha) + LogGamma(beta) - LogGamma(alpha + beta)) -
            Math.Log(span);
        return Math.Exp(logDensity);
    }

    private static double LogGamma(double value)
    {
        if (value < 0.5)
            return Math.Log(Math.PI) - Math.Log(Math.Sin(Math.PI * value)) -
                LogGamma(1 - value);
        double z = value - 1;
        double sum = Lanczos[0];
        for (int i = 1; i < Lanczos.Length; i++) sum += Lanczos[i] / (z + i);
        double t = z + 7.5;
        return 0.5 * Math.Log(2 * Math.PI) + (z + 0.5) * Math.Log(t) - t + Math.Log(sum);
    }
}
