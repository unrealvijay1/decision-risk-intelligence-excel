namespace MonteCarlo.Core;

public enum ChartValueRange
{
    Invalid,
    BelowRange,
    InRange,
    AboveRange
}

/// <summary>A numeric axis classification; only in-range values have a position.</summary>
public readonly record struct ChartValuePosition
{
    public ChartValueRange Range { get; }
    public double? NormalizedPosition { get; }

    private ChartValuePosition(ChartValueRange range, double? position = null)
    {
        Range = range;
        NormalizedPosition = position;
    }

    public static ChartValuePosition Calculate(double? value, double min, double max)
    {
        if (!value.HasValue || !double.IsFinite(value.Value) ||
            !double.IsFinite(min) || !double.IsFinite(max) || max <= min)
            return new(ChartValueRange.Invalid);

        double number = value.Value;
        if (number < min) return new(ChartValueRange.BelowRange);
        if (number > max) return new(ChartValueRange.AboveRange);
        if (number == min) return new(ChartValueRange.InRange, 0);
        if (number == max) return new(ChartValueRange.InRange, 1);

        double span = max - min;
        // Halving before subtraction keeps opposite-sign finite extremes finite.
        double ratio = double.IsFinite(span)
            ? (number - min) / span
            : (number / 2 - min / 2) / (max / 2 - min / 2);
        return new(ChartValueRange.InRange, Math.Clamp(ratio, 0, 1));
    }
}
