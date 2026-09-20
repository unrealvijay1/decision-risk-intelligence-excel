using MonteCarlo.Core;

namespace MonteCarlo.Core.Tests;

public class ChartValuePositionTests
{
    [Theory]
    [InlineData(10, 10, 30, 0)]
    [InlineData(30, 10, 30, 1)]
    [InlineData(20, 10, 30, 0.5)]
    [InlineData(14, 10, 30, 0.2)]
    [InlineData(-15, -30, -10, 0.75)]
    [InlineData(0d, -10, 30, 0.25)]
    [InlineData(50, 0, 100, 0.5)]
    [InlineData(80, 0, 100, 0.8)]
    [InlineData(0d, -double.MaxValue, double.MaxValue, 0.5)]
    [InlineData(double.MaxValue / 2, -double.MaxValue, double.MaxValue, 0.75)]
    [InlineData(-double.MaxValue, -double.MaxValue, double.MaxValue, 0)]
    [InlineData(double.MaxValue, -double.MaxValue, double.MaxValue, 1)]
    [InlineData(double.Epsilon, 0, 2 * double.Epsilon, 0.5)]
    public void InRangeValuesHaveExpectedPosition(double value, double min, double max, double expected)
    {
        var result = ChartValuePosition.Calculate(value, min, max);
        Assert.Equal(ChartValueRange.InRange, result.Range);
        Assert.NotNull(result.NormalizedPosition);
        Assert.InRange(result.NormalizedPosition.Value, 0, 1);
        Assert.Equal(expected, result.NormalizedPosition.Value, 12);
    }

    [Theory]
    [InlineData(-1, ChartValueRange.BelowRange)]
    [InlineData(101, ChartValueRange.AboveRange)]
    public void OutsideValuesHaveNoPosition(double value, ChartValueRange expected)
    {
        var result = ChartValuePosition.Calculate(value, 0, 100);
        Assert.Equal(expected, result.Range);
        Assert.Null(result.NormalizedPosition);
    }

    [Theory]
    [InlineData(null, 0, 1)]
    [InlineData(double.NaN, 0, 1)]
    [InlineData(double.PositiveInfinity, 0, 1)]
    [InlineData(double.NegativeInfinity, 0, 1)]
    [InlineData(0d, 1, 1)]
    [InlineData(0d, 2, 1)]
    [InlineData(0d, double.NaN, 1)]
    [InlineData(0d, 0, double.NaN)]
    [InlineData(0d, double.NegativeInfinity, 1)]
    [InlineData(0d, double.PositiveInfinity, 1)]
    [InlineData(0d, 0, double.PositiveInfinity)]
    [InlineData(0d, 0, double.NegativeInfinity)]
    public void InvalidInputsHaveNoPosition(double? value, double min, double max)
    {
        var result = ChartValuePosition.Calculate(value, min, max);
        Assert.Equal(ChartValueRange.Invalid, result.Range);
        Assert.Null(result.NormalizedPosition);
    }
}
