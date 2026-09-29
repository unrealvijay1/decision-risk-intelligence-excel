using System.Text.Json;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public static class DistributionDataPersistence
{
    public const string Header = "DistributionDataV1";
    private sealed record Payload(int Version, DiscreteOutcome[] Outcomes);
    public static string? Serialize(AssumptionDefinition assumption)
    {
        if (assumption.Distribution != DistributionType.Discrete) return null;
        if (assumption.ProbabilityTable == null) throw new ArgumentException("Discrete assumption is missing its probability table.");
        return JsonSerializer.Serialize(new Payload(1, assumption.ProbabilityTable.Outcomes.ToArray()));
    }
    public static DiscreteTable? Deserialize(DistributionType kind, object? header, object? value)
    {
        if (kind != DistributionType.Discrete && value == null) return null;
        if (!Equals(header, Header) || value is not string json || kind != DistributionType.Discrete)
            throw new ArgumentException("Saved distribution data is missing or unsupported. Redefine the assumption's probability table.");
        try
        {
            var payload = JsonSerializer.Deserialize<Payload>(json);
            if (payload?.Version != 1 || payload.Outcomes == null) throw new ArgumentException("Unsupported probability-table version.");
            return new DiscreteTable(payload.Outcomes);
        }
        catch (Exception ex) when (ex is JsonException or ArgumentException)
        { throw new ArgumentException("Saved Discrete probability table is invalid: " + ex.Message, ex); }
    }
}
