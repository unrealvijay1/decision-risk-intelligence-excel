namespace MonteCarlo.Core;

/// <summary>Inverse marginal transform used only for correlated sampling. Independent samplers remain unchanged.</summary>
public static class DistributionQuantile
{
    public static double FromGaussian(DistributionKind kind, double z, double p1, double p2, double p3 = 0, double p4 = 0)
    {
        // Avoid roundoff to CDF endpoints for unbounded marginals; inverse Phi(Phi(z)) is exactly z.
        if (kind == DistributionKind.Normal) return p1 + p2 * z;
        if (kind == DistributionKind.Lognormal) return Math.Exp(p1 + p2 * z);
        return Inverse(kind, DistributionFitter.StandardNormalCDF(z), p1, p2, p3, p4);
    }

    public static double Inverse(DistributionKind kind, double probability, double p1, double p2, double p3 = 0, double p4 = 0)
    {
        if (DistributionPreview.Validate(kind, p1, p2, p3, p4) is string error) throw new ArgumentException(error);
        if (!double.IsFinite(probability) || probability < 0 || probability > 1) throw new ArgumentException("Probability must be between zero and one.");
        double u = probability;
        switch (kind)
        {
            case DistributionKind.Normal:
            case DistributionKind.Lognormal:
                double z;
                if (u == 0) z = double.NegativeInfinity;
                else if (u == 1) z = double.PositiveInfinity;
                else
                {
                    double lo = -40, hi = 40;
                    for (int i = 0; i < 64; i++) { double mid = (lo + hi) / 2; if (DistributionFitter.StandardNormalCDF(mid) < u) lo = mid; else hi = mid; }
                    z = (lo + hi) / 2;
                }
                return kind == DistributionKind.Normal ? p1 + p2 * z : Math.Exp(p1 + p2 * z);
            case DistributionKind.Uniform: return (1 - u) * p1 + u * p2;
            case DistributionKind.Triangular:
                double fraction = (p2 - p1) / (p3 - p1);
                return u < fraction ? p1 + (p3 - p1) * Math.Sqrt(u * fraction)
                    : p3 - (p3 - p1) * Math.Sqrt((1 - u) * (1 - fraction));
            case DistributionKind.Pert:
                return ScaleBeta(u, p1, p3, 1 + 4 * (p2 - p1) / (p3 - p1), 1 + 4 * (p3 - p2) / (p3 - p1));
            case DistributionKind.Beta: return ScaleBeta(u, p1, p2, p3, p4);
            default: throw new ArgumentException("Unsupported distribution.");
        }
    }
    private static double ScaleBeta(double u, double min, double max, double a, double b)
    {
        if (u == 0) return min;
        if (u == 1) return max;
        // Reflect upper tail; bisection in logit space resolves small shape parameters near endpoints.
        bool reflect = u > .5;
        if (reflect) { (a, b) = (b, a); u = 1 - u; }
        double lo = -745, hi = 37;
        for (int i = 0; i < 80; i++)
        {
            double mid = (lo + hi) / 2;
            double x = mid < 0 ? Math.Exp(mid) / (1 + Math.Exp(mid)) : 1 / (1 + Math.Exp(-mid));
            if (BetaCdf(x, a, b) < u) lo = mid; else hi = mid;
        }
        double t = (lo + hi) / 2;
        double q = t < 0 ? Math.Exp(t) / (1 + Math.Exp(t)) : 1 / (1 + Math.Exp(-t));
        if (reflect) q = 1 - q;
        return (1 - q) * min + q * max;
    }
    public static double BetaCdf(double x, double a, double b)
    {
        if (x <= 0) return 0;
        if (x >= 1) return 1;
        double logBeta = DistributionPreview.LogGamma(a) + DistributionPreview.LogGamma(b) - DistributionPreview.LogGamma(a + b);
        double factor = Math.Exp(a * Math.Log(x) + b * Math.Log(1 - x) - logBeta);
        double result = x < (a + 1) / (a + b + 2) ? factor * Fraction(x, a, b) / a : 1 - factor * Fraction(1 - x, b, a) / b;
        if (!double.IsFinite(result)) throw new ArgumentException("The Beta/PERT parameters exceed the supported numerical range for correlated sampling.");
        return Math.Clamp(result, 0, 1);
    }
    // Modified Lentz continued fraction for regularized incomplete beta, with explicit convergence failure.
    private static double Fraction(double x, double a, double b)
    {
        static double Nonzero(double v) => Math.Abs(v) < 1e-300 ? Math.CopySign(1e-300, v) : v;
        double c = 1, d = 1 / Nonzero(1 - (a + b) * x / (a + 1)), h = d;
        for (int m = 1; m <= 1000; m++)
        {
            double aa = m * (b - m) * x / ((a + 2 * m - 1) * (a + 2 * m));
            d = 1 / Nonzero(1 + aa * d); c = Nonzero(1 + aa / c); h *= d * c;
            aa = -(a + m) * (a + b + m) * x / ((a + 2 * m) * (a + 2 * m + 1));
            d = 1 / Nonzero(1 + aa * d); c = Nonzero(1 + aa / c);
            double delta = d * c; h *= delta;
            if (Math.Abs(delta - 1) < 3e-14) return h;
        }
        throw new ArgumentException("The Beta/PERT inverse calculation did not converge. Review the correlated assumption parameters.");
    }
}
