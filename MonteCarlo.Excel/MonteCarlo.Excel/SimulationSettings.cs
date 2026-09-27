using System.Globalization;

namespace MonteCarlo.Excel;

public enum SimulationSeedMode { Automatic, Fixed }
public sealed record SimulationSettings(int TrialCount = 10000, SimulationSeedMode SeedMode = SimulationSeedMode.Automatic, int? FixedSeed = null)
{
    public const int MinimumTrials = 1;
    public const int MaximumTrials = 1000000;
    public static readonly IReadOnlyList<int> Presets = Array.AsReadOnly(new[] { 1000, 5000, 10000, 50000 });
    public static string? ValidateTrials(int trials) => trials < MinimumTrials || trials > MaximumTrials
        ? $"Enter a whole number of trials from {MinimumTrials:N0} to {MaximumTrials:N0}." : null;
    public string? Validate() => ValidateTrials(TrialCount) ??
        (!Enum.IsDefined(SeedMode) || (SeedMode == SimulationSeedMode.Fixed && !FixedSeed.HasValue)
            ? "Enter an integer fixed seed from -2147483648 to 2147483647, or select Automatic." : null);
    public static bool TryParse(string trials, SimulationSeedMode mode, string seed, out SimulationSettings settings, out string? error)
    {
        settings = new();
        if (!int.TryParse(trials, NumberStyles.Integer, CultureInfo.CurrentCulture, out int count))
        { error = ValidateTrials(0); return false; }
        int? fixedSeed = null;
        if (mode == SimulationSeedMode.Fixed)
        {
            if (!int.TryParse(seed, NumberStyles.Integer, CultureInfo.CurrentCulture, out int parsed))
            { error = "Enter an integer fixed seed from -2147483648 to 2147483647."; return false; }
            fixedSeed = parsed;
        }
        settings = new(count, mode, fixedSeed);
        error = settings.Validate(); return error == null;
    }
}
