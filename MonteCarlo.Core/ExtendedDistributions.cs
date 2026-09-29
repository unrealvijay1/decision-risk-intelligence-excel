namespace MonteCarlo.Core;

/// <summary>Authoritative mathematics for the eight additional marginals. Existing samplers are untouched.</summary>
public static class ExtendedDistributions
{
    private const int IterationLimit = 10000;
    public static string? Validate(DistributionKind kind, double a, double b, double c, double d, DiscreteTable? table = null)
    {
        if (!double.IsFinite(a) || !double.IsFinite(b) || !double.IsFinite(c) || !double.IsFinite(d)) return "Enter finite numeric parameters.";
        if (kind == DistributionKind.Discrete) return table == null ? "Enter a valid outcome/probability table (2–100 rows)." : null;
        if (kind == DistributionKind.Exponential) return a > 0 && a <= double.MaxValue / 1024 ? null : "Mean must be positive and no greater than MaxValue/1024; rescale the units.";
        if (kind == DistributionKind.Poisson) return a > 0 && a <= 1_000_000 ? null : "Lambda must be greater than zero and at most 1,000,000.";
        if (kind == DistributionKind.Binomial) return a >= 1 && a <= 1_000_000 && a == Math.Truncate(a) && b >= 0 && b <= 1 ? null : "Trials must be an integer from 1 to 1,000,000; probability must be between 0 and 1.";
        if (kind == DistributionKind.Bernoulli) return a >= 0 && a <= 1 ? null : "Probability must be between 0 and 1.";
        if (kind is DistributionKind.Gamma or DistributionKind.Weibull)
        {
            if (a < .01 || a > 1_000_000 || b <= 0) return "Shape must be between 0.01 and 1,000,000 and scale must be positive; more extreme shapes are not numerically supported.";
            double logLow = Math.Log(b) + (kind == DistributionKind.Weibull ? Math.Log(1e-16) / a : Math.Min(0, Math.Log(1e-16) / a));
            double logHigh = Math.Log(b) + (kind == DistributionKind.Weibull ? Math.Log(40) / a : Math.Log(a + 50 * Math.Sqrt(a) + 100));
            if (logHigh > 700 || logLow < -740) return "Shape and scale exceed the supported numeric range in the tails. Rescale the units or choose a less extreme shape.";
            return null;
        }
        if (kind == DistributionKind.TruncatedNormal)
        {
            if (b <= 0 || c >= d) return "Standard deviation must be positive and lower bound must be less than upper bound.";
            double lo = (c - a) / b, hi = (d - a) / b;
            if (!double.IsFinite(lo) || !double.IsFinite(hi) || Math.Abs(lo) > 10000 || Math.Abs(hi) > 10000 ||
                Math.BitIncrement(lo) >= hi || Math.BitIncrement(c) >= d)
                return "Bounds must contain representable interior values and lie within 10,000 standard deviations of the mean. Rescale or widen the interval.";
            return null;
        }
        return "Unsupported distribution.";
    }
    private static void Check(DistributionKind kind, double a, double b, double c, double d, DiscreteTable? table)
    { if (Validate(kind, a, b, c, d, table) is { } error) throw new ArgumentException(error); }

    public static double Probability(DistributionKind kind, double x, double a, double b = 0, double c = 0, double d = 0, DiscreteTable? table = null)
    {
        Check(kind, a, b, c, d, table);
        if (double.IsNaN(x)) throw new ArgumentException("Value must not be NaN.");
        switch (kind)
        {
            case DistributionKind.Discrete: return table!.Mass(x);
            case DistributionKind.Bernoulli: return x == 0 ? 1 - a : x == 1 ? a : 0;
            case DistributionKind.Binomial:
                if (x < 0 || x > a || x != Math.Truncate(x)) return 0;
                if (b == 0) return x == 0 ? 1 : 0;
                if (b == 1) return x == a ? 1 : 0;
                return Math.Exp(DistributionPreview.LogGamma(a + 1) - DistributionPreview.LogGamma(x + 1) - DistributionPreview.LogGamma(a - x + 1) + x * Math.Log(b) + (a - x) * Log1p(-b));
            case DistributionKind.Poisson:
                if (!double.IsFinite(x) || x < 0 || x != Math.Truncate(x)) return 0;
                // Chernoff bound below the smallest representable mass avoids overflow in log-gamma.
                if (x > a && -a + x * (1 + Math.Log(a) - Math.Log(x)) < -745) return 0;
                return Math.Exp(-a + x * Math.Log(a) - DistributionPreview.LogGamma(x + 1));
            case DistributionKind.Exponential: return x < 0 ? 0 : Math.Exp(-x / a) / a;
            case DistributionKind.Weibull:
                if (x < 0 || double.IsPositiveInfinity(x)) return 0;
                if (x == 0) return a < 1 ? double.PositiveInfinity : a == 1 ? 1 / b : 0;
                double logRatio = Math.Log(x) - Math.Log(b);
                return Math.Exp(Math.Log(a) - Math.Log(b) + (a - 1) * logRatio - Math.Exp(a * logRatio));
            case DistributionKind.Gamma:
                if (x < 0 || double.IsPositiveInfinity(x)) return 0;
                if (x == 0) return a < 1 ? double.PositiveInfinity : a == 1 ? 1 / b : 0;
                return Math.Exp((a - 1) * (Math.Log(x) - Math.Log(b)) - x / b - DistributionPreview.LogGamma(a) - Math.Log(b));
            case DistributionKind.TruncatedNormal:
                if (x < c || x > d) return 0;
                double z = (x - a) / b;
                return Math.Exp(-.5 * z * z - .9189385332046727 - Math.Log(b) - NormalLogMass((c - a) / b, (d - a) / b));
            default: throw new ArgumentException("Unsupported distribution.");
        }
    }
    public static double Cdf(DistributionKind kind, double x, double a, double b = 0, double c = 0, double d = 0, DiscreteTable? table = null)
    {
        Check(kind, a, b, c, d, table);
        if (double.IsNaN(x)) throw new ArgumentException("Value must not be NaN.");
        switch (kind)
        {
            case DistributionKind.Discrete: return table!.Cdf(x);
            case DistributionKind.Bernoulli: return x < 0 ? 0 : x < 1 ? 1 - a : 1;
            case DistributionKind.Binomial:
                if (x < 0) return 0; if (x >= a || b == 0) return 1; if (b == 1) return 0;
                if (x < 1) return Math.Exp(a * Log1p(-b));
                return b < .5 ? 1 - DistributionQuantile.BetaCdf(b, Math.Floor(x) + 1, a - Math.Floor(x))
                    : DistributionQuantile.BetaCdf(1 - b, a - Math.Floor(x), Math.Floor(x) + 1);
            case DistributionKind.Poisson:
                if (x < 0) return 0; if (double.IsPositiveInfinity(x)) return 1;
                // The omitted survival probability is less than half an ulp at one.
                double next = Math.Floor(x) + 1;
                if (next > a && -a + next * (1 + Math.Log(a) - Math.Log(next)) < -38) return 1;
                return GammaRatio(Math.Floor(x) + 1, a, upper: true);
            case DistributionKind.Exponential: return x <= 0 ? 0 : -Expm1(-x / a);
            case DistributionKind.Weibull: return x <= 0 ? 0 : -Expm1(-Math.Pow(x / b, a));
            case DistributionKind.Gamma: return x <= 0 ? 0 : GammaRatio(a, x / b);
            case DistributionKind.TruncatedNormal:
                if (x <= c) return 0; if (x >= d) return 1;
                double lo = (c - a) / b, hi = (d - a) / b, z = (x - a) / b;
                return Math.Clamp(Math.Exp(NormalLogMass(lo, z) - NormalLogMass(lo, hi)), 0, 1);
            default: throw new ArgumentException("Unsupported distribution.");
        }
    }
    public static double Quantile(DistributionKind kind, double u, double a, double b = 0, double c = 0, double d = 0, DiscreteTable? table = null)
    {
        Check(kind, a, b, c, d, table);
        if (!double.IsFinite(u) || u < 0 || u > 1) throw new ArgumentException("Probability must be between zero and one.");
        switch (kind)
        {
            case DistributionKind.Discrete: return table!.Quantile(u);
            case DistributionKind.Bernoulli: return a == 1 ? 1 : u <= 1 - a ? 0 : 1;
            case DistributionKind.Exponential: return -a * Log1p(-u);
            case DistributionKind.Weibull: return b * Math.Pow(-Log1p(-u), 1 / a);
            case DistributionKind.Binomial:
            case DistributionKind.Poisson:
                if (kind == DistributionKind.Binomial && b == 1) return a;
                if (u == 0 || kind == DistributionKind.Binomial && b == 0) return 0;
                if (u == 1) return kind == DistributionKind.Binomial ? a : double.PositiveInfinity;
                int low = 0, high = kind == DistributionKind.Binomial ? (int)a : (int)Math.Ceiling(a + 16 * Math.Sqrt(a) + 100);
                while (low < high) { int mid = low + (high - low) / 2; if (Cdf(kind, mid, a, b) >= u) high = mid; else low = mid + 1; }
                return low;
            case DistributionKind.Gamma:
                if (u == 0) return 0; if (u == 1) return double.PositiveInfinity;
                // Log-space inversion also resolves small-shape lower tails.
                double l = -744, h = Math.Log(a + 50 * Math.Sqrt(a) + 100);
                for (int i = 0; i < 80; i++) { double m = (l + h) / 2; if (GammaRatio(a, Math.Exp(m)) < u) l = m; else h = m; }
                return b * Math.Exp((l + h) / 2);
            case DistributionKind.TruncatedNormal:
                if (u == 0) return c; if (u == 1) return d;
                double left = c, right = d;
                for (int i = 0; i < 80; i++)
                {
                    double mid = .5 * left + .5 * right;
                    if (mid == left || mid == right) break;
                    if (Cdf(kind, mid, a, b, c, d) < u) left = mid; else right = mid;
                }
                return .5 * left + .5 * right;
            default: throw new ArgumentException("Unsupported distribution.");
        }
    }
    public static double Sample(DistributionKind kind, double a, double b, double c, double d, Random random, DiscreteTable? table = null)
    {
        Check(kind, a, b, c, d, table);
        return kind switch
        {
            DistributionKind.Exponential => -a * Math.Log(Open(random)),
            DistributionKind.Weibull => b * Math.Pow(-Math.Log(Open(random)), 1 / a),
            DistributionKind.Gamma => b * GammaSample(a, random),
            DistributionKind.Poisson => PoissonSample(a, random),
            DistributionKind.Binomial => BinomialSample((int)a, b, random),
            DistributionKind.Bernoulli => random.NextDouble() < a ? 1 : 0,
            DistributionKind.Discrete => table!.Quantile(Open(random)),
            DistributionKind.TruncatedNormal => a + b * TruncatedSample((c - a) / b, (d - a) / b, random),
            _ => throw new ArgumentException("Unsupported distribution.")
        };
    }
    private static double Open(Random random)
    {
        for (int i = 0; i < IterationLimit; i++) { double u = random.NextDouble(); if (u > 0 && u < 1) return u; }
        throw new ArgumentException("Random generator did not produce an interior probability.");
    }
    // Marsaglia–Tsang gamma, with an explicit iteration bound. Original Beta/PERT RNG paths are unchanged.
    private static double GammaSample(double shape, Random random)
    {
        if (shape < 1) return GammaSample(shape + 1, random) * Math.Pow(Open(random), 1 / shape);
        double d = shape - 1d / 3, c = 1 / Math.Sqrt(9 * d);
        for (int i = 0; i < IterationLimit; i++)
        {
            double z = Math.Sqrt(-2 * Math.Log(Open(random))) * Math.Cos(2 * Math.PI * Open(random));
            double v = 1 + c * z; if (v <= 0) continue; v = v * v * v;
            double u = Open(random);
            if (u < 1 - .0331 * z * z * z * z || Math.Log(u) < .5 * z * z + d * (1 - v + Math.Log(v))) return d * v;
        }
        throw new ArgumentException("Gamma sampling did not converge; review the parameters.");
    }
    private static double PoissonSample(double lambda, Random random)
    {
        if (lambda < 10)
        {
            double product = 1, limit = Math.Exp(-lambda);
            for (int k = 0; k < IterationLimit; k++) { product *= Open(random); if (product <= limit) return k; }
        }
        else
        {
            // Hoermann's transformed rejection (PTRS); no normal approximation.
            double b = .931 + 2.53 * Math.Sqrt(lambda), a = -.059 + .02483 * b;
            double inverseAlpha = 1.1239 + 1.1328 / (b - 3.4), squeeze = .9277 - 3.6224 / (b - 2);
            for (int i = 0; i < IterationLimit; i++)
            {
                double u = Open(random) - .5, v = Open(random), us = .5 - Math.Abs(u);
                double k = Math.Floor((2 * a / us + b) * u + lambda + .43);
                if (k < 0 || us < .013 && v > us) continue;
                if (us >= .07 && v <= squeeze || Math.Log(v * inverseAlpha / (a / (us * us) + b)) <= -lambda + k * Math.Log(lambda) - DistributionPreview.LogGamma(k + 1)) return k;
            }
        }
        throw new ArgumentException("Poisson sampling did not converge; review the parameters.");
    }
    private static double BinomialSample(int n, double p, Random random)
    {
        if (p == 0) return 0; if (p == 1) return n;
        // Condition on a beta-distributed order statistic, reducing the remaining trial count.
        int successes = 0;
        for (int iteration = 0; n > 32 && iteration < 128; iteration++)
        {
            int m = Math.Min(n, 1 + (int)(n * p));
            double x = GammaSample(m, random), y = GammaSample(n - m + 1, random), split = x / (x + y);
            if (split <= p) { successes += m; n -= m; p = (p - split) / (1 - split); }
            else { n = m - 1; p /= split; }
        }
        if (n > 32) throw new ArgumentException("Binomial sampling did not converge; review the parameters.");
        for (int i = 0; i < n; i++) if (random.NextDouble() < p) successes++;
        return successes;
    }
    private static double TruncatedSample(double lo, double hi, Random random)
    {
        if (hi <= 0) return -TruncatedSample(-hi, -lo, random);
        double closest = Math.Max(0, lo);
        for (int i = 0; i < IterationLimit; i++)
        {
            double z;
            if (hi - lo < 2 / (1 + closest))
            {
                z = lo + Open(random) * (hi - lo);
                if (Math.Log(Open(random)) <= -.5 * (z - closest) * (z + closest)) return z;
            }
            else if (lo > 0)
            {
                double rate = .5 * (lo + Math.Sqrt(lo * lo + 4));
                z = lo - Math.Log(Open(random)) / rate;
                if (z <= hi && Math.Log(Open(random)) <= -.5 * (z - rate) * (z - rate)) return z;
            }
            else
            {
                z = Math.Sqrt(-2 * Math.Log(Open(random))) * Math.Cos(2 * Math.PI * Open(random));
                if (z >= lo && z <= hi) return z;
            }
        }
        throw new ArgumentException("Truncated Normal sampling did not converge; review the bounds.");
    }

    // Regularized incomplete gamma: convergent series below a+1, Lentz fraction above it.
    internal static double GammaRatio(double a, double x, bool upper = false)
    {
        if (x <= 0) return upper ? 1 : 0;
        if (double.IsPositiveInfinity(x)) return upper ? 0 : 1;
        double logFactor = a * Math.Log(x) - x - DistributionPreview.LogGamma(a);
        if (x < a + 1)
        {
            double term = 1 / a, sum = term;
            for (int i = 1; i <= IterationLimit; i++)
            {
                term *= x / (a + i); sum += term;
                if (Math.Abs(term) < Math.Abs(sum) * 3e-15)
                { double logP = logFactor + Math.Log(sum); return Math.Clamp(upper ? -Expm1(logP) : Math.Exp(logP), 0, 1); }
            }
        }
        else
        {
            double logQ = logFactor + GammaLogFraction(a, x);
            return Math.Clamp(upper ? Math.Exp(logQ) : -Expm1(logQ), 0, 1);
        }
        throw new ArgumentException("Incomplete Gamma calculation did not converge; reduce the parameter scale.");
    }
    private static double GammaLogFraction(double a, double x)
    {
        static double Nonzero(double value) => Math.Abs(value) < 1e-300 ? Math.CopySign(1e-300, value) : value;
        double b = x + 1 - a, c = 1e300, d = 1 / Nonzero(b), h = d;
        for (int i = 1; i <= IterationLimit; i++)
        {
            double an = -i * (i - a); b += 2;
            d = 1 / Nonzero(an * d + b); c = Nonzero(b + an / c);
            double delta = d * c; h *= delta;
            if (Math.Abs(delta - 1) < 3e-15) return Math.Log(h);
        }
        throw new ArgumentException("Incomplete Gamma calculation did not converge.");
    }
    private static double LogNormalTail(double x)
    {
        if (x < 2) return Math.Log(.5 * GammaRatio(.5, .5 * x * x, upper: true));
        double t = .5 * x * x;
        return Math.Log(.5) + .5 * Math.Log(t) - t - DistributionPreview.LogGamma(.5) + GammaLogFraction(.5, t);
    }
    private static double NormalLogMass(double lo, double hi)
    {
        if (lo >= hi) return double.NegativeInfinity;
        if ((hi - lo) * (1 + Math.Max(Math.Abs(lo), Math.Abs(hi))) < .1)
        {
            // Eight-point Gaussian quadrature of a smooth relative density avoids tail subtraction.
            ReadOnlySpan<double> nodes = [.1834346424956498, .5255324099163290, .7966664774136267, .9602898564975363];
            ReadOnlySpan<double> weights = [.3626837833783620, .3137066458778873, .2223810344533745, .1012285362903763];
            double mid = .5 * lo + .5 * hi, half = .5 * (hi - lo), sum = 0;
            for (int i = 0; i < nodes.Length; i++)
            { double delta = half * nodes[i]; sum += weights[i] * (Math.Exp(-delta * mid - .5 * delta * delta) + Math.Exp(delta * mid - .5 * delta * delta)); }
            return Math.Log(half) + Math.Log(sum) - .5 * mid * mid - .9189385332046727;
        }
        if (hi <= 0) return NormalLogMass(-hi, -lo);
        if (lo >= 0) { double first = LogNormalTail(lo); return first + Math.Log(-Expm1(LogNormalTail(hi) - first)); }
        return Math.Log(1 - Math.Exp(LogNormalTail(-lo)) - Math.Exp(LogNormalTail(hi)));
    }
    internal static double Log1p(double x)
    {
        if (Math.Abs(x) > 1e-4) return Math.Log(1 + x);
        double sum = 0, term = x;
        for (int k = 1; k <= 12; k++) { sum += term / k; term *= -x; }
        return sum;
    }
    internal static double Expm1(double x)
    {
        if (Math.Abs(x) > 1e-5) return Math.Exp(x) - 1;
        double sum = x, term = x;
        for (int k = 2; k <= 10; k++) { term *= x / k; sum += term; }
        return sum;
    }
}
