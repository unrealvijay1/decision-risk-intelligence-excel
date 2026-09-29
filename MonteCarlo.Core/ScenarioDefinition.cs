namespace MonteCarlo.Core;

public enum ScenarioAdjustmentType { PercentageChange }

public sealed record ScenarioAdjustment(string AssumptionId, ScenarioAdjustmentType Type, double Value);
public sealed record ScenarioDefinition(string Id, string Name, IReadOnlyList<ScenarioAdjustment> Adjustments);

public static class ScenarioValidation
{
    public static void Validate(IReadOnlyList<ScenarioDefinition> scenarios)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var scenario in scenarios)
        {
            if (scenario == null || string.IsNullOrWhiteSpace(scenario.Id) || !ids.Add(scenario.Id))
                throw new ArgumentException("Each scenario must have a unique identity.");
            if (string.IsNullOrWhiteSpace(scenario.Name) || scenario.Name.Length > 100 ||
                !names.Add(scenario.Name.Trim()))
                throw new ArgumentException("Use unique scenario names of 1–100 characters.");
            if (scenario.Adjustments == null) throw new ArgumentException("Scenario adjustments are missing.");
            var assumptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var adjustment in scenario.Adjustments)
            {
                if (adjustment == null || string.IsNullOrWhiteSpace(adjustment.AssumptionId) || !assumptions.Add(adjustment.AssumptionId))
                    throw new ArgumentException($"{scenario.Name}: select each assumption only once.");
                if (adjustment.Type != ScenarioAdjustmentType.PercentageChange)
                    throw new ArgumentException($"{scenario.Name}: this adjustment type is not supported.");
                ScenarioTransformation.Factor(adjustment.Value);
            }
        }
    }
}

public sealed record ScenarioDistribution(DistributionKind Kind, double P1, double P2, double P3 = 0, double P4 = 0,
    double? Constant = null, DiscreteTable? ProbabilityTable = null, double SampleScale = 1);

/// <summary>Percentage means Y = (1 + percentage/100) X, not a percentage of every parameter.</summary>
public static class ScenarioTransformation
{
    public static double Factor(double percentage)
    {
        if (!double.IsFinite(percentage) || percentage < -100)
            throw new ArgumentException("Enter a finite percentage of -100 or greater. Values below -100% are not supported.");
        return 1 + percentage / 100;
    }

    public static ScenarioDistribution Apply(ScenarioDistribution source, double percentage)
    {
        double factor = Factor(percentage);
        if (source.Constant is double constant)
        {
            if (!double.IsFinite(constant)) throw new ArgumentException("The constant must be finite.");
            double scaled = constant * factor;
            if (!double.IsFinite(scaled)) throw new ArgumentException("The percentage produces a constant outside the numeric range.");
            return source with { Constant = scaled };
        }
        string? error = DistributionPreview.Validate(source.Kind, source.P1, source.P2, source.P3, source.P4, source.ProbabilityTable);
        if (error != null) throw new ArgumentException(error);
        if (factor == 0) return source with { Constant = 0 };
        if (source.Kind >= DistributionKind.Exponential)
        {
            double scale = source.SampleScale * factor;
            if (!double.IsFinite(scale) || scale < 0) throw new ArgumentException("The percentage produces an invalid sample scale.");
            // Count-family scaling is a value transform, never a change in success probability.
            foreach (double value in source.ProbabilityTable?.Outcomes.Select(r => r.Outcome) ?? new[] { source.P1, source.P2, source.P3, source.P4 })
                if (!double.IsFinite(value * scale)) throw new ArgumentException("The percentage produces values outside the numeric range.");
            return source with { SampleScale = scale };
        }
        var result = source.Kind switch
        {
            DistributionKind.Normal => source with { P1 = source.P1 * factor, P2 = source.P2 * factor },
            DistributionKind.Lognormal => source with { P1 = source.P1 + Math.Log(factor) },
            DistributionKind.Uniform or DistributionKind.Beta => source with { P1 = source.P1 * factor, P2 = source.P2 * factor },
            DistributionKind.Triangular or DistributionKind.Pert => source with
                { P1 = source.P1 * factor, P2 = source.P2 * factor, P3 = source.P3 * factor },
            _ => throw new ArgumentException("Unsupported distribution.")
        };
        error = DistributionPreview.Validate(result.Kind, result.P1, result.P2, result.P3, result.P4);
        if (error != null) throw new ArgumentException("The percentage produces invalid parameters: " + error);
        if (result.Kind is DistributionKind.Uniform or DistributionKind.Beta && !double.IsFinite(result.P2 - result.P1) ||
            result.Kind is DistributionKind.Triangular or DistributionKind.Pert && !double.IsFinite(result.P3 - result.P1) ||
            result.Kind == DistributionKind.Lognormal && !double.IsFinite(Math.Exp(result.P1)))
            throw new ArgumentException("The percentage produces a distribution outside the numeric range.");
        return result;
    }

    // Sampling the unchanged source consumes identical RNG draws, including at -100%.
    // This also avoids PERT shape-rounding changing rejection-sampler consumption.
    public static double ScaleSample(double sample, double percentage) => sample * Factor(percentage);
}
