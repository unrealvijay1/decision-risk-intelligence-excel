using System.Runtime.InteropServices;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

public class SimulationValidationTests
{
    [Theory]
    [InlineData(DistributionType.Normal, 0, 1, 0, 0)]
    [InlineData(DistributionType.Lognormal, -1, 1, 0, 0)]
    [InlineData(DistributionType.Uniform, -2, 2, 0, 0)]
    [InlineData(DistributionType.Triangular, 0, 0, 10, 0)]
    [InlineData(DistributionType.Pert, 0, 10, 10, 0)]
    [InlineData(DistributionType.Beta, 0, 1, 2, 3)]
    public void ValidDistributions(DistributionType kind, double p1, double p2, double p3, double p4)
    {
        var result = new ValidationResult();
        SimulationValidation.ValidateParameters(result, "Input!A1", kind, [p1, p2, p3, p4]);
        Assert.True(result.IsValid);
    }

    [Theory]
    [InlineData(DistributionType.Normal, 0, 0, 0, 0)]
    [InlineData(DistributionType.Lognormal, 1, -1, 0, 0)]
    [InlineData(DistributionType.Uniform, 2, 1, 0, 0)]
    [InlineData(DistributionType.Triangular, 0, 11, 10, 0)]
    [InlineData(DistributionType.Pert, 0, -1, 10, 0)]
    [InlineData(DistributionType.Beta, 0, 1, 0, 3)]
    [InlineData(DistributionType.Beta, 0, 1, 2, -1)]
    [InlineData((DistributionType)123, 0, 1, 0, 0)]
    public void InvalidDistributions(DistributionType kind, double p1, double p2, double p3, double p4)
    {
        var result = new ValidationResult();
        SimulationValidation.ValidateParameters(result, "Input!A1", kind, [p1, p2, p3, p4]);
        Assert.False(result.IsValid);
        Assert.All(result.Errors, error => Assert.NotEmpty(error.SuggestedCorrection));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("abc")]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(true)]
    public void InvalidOrMissingParameterIsNotConvertedToZero(object? value)
    {
        var result = new ValidationResult();
        SimulationValidation.ValidateParameters(result, "Input!A1", DistributionType.Normal, [value, 1]);
        Assert.False(result.IsValid);
        Assert.Contains("Mean", result.UserMessage);
    }

    [Fact]
    public void MissingFourthParameterIsRequiredOnlyForBeta()
    {
        var valid = new ValidationResult();
        SimulationValidation.ValidateParameters(valid, "Legacy", DistributionType.Pert, [0, 5, 10]);
        Assert.True(valid.IsValid);
        var invalid = new ValidationResult();
        SimulationValidation.ValidateParameters(invalid, "Legacy", DistributionType.Beta, [0, 1, 2]);
        Assert.Contains("Beta", invalid.UserMessage);
    }

    [Theory]
    [InlineData("#VALUE!")]
    [InlineData("#N/A")]
    [InlineData("#DIV/0!")]
    [InlineData("#REF!")]
    [InlineData("#NUM!")]
    [InlineData("#NAME?")]
    [InlineData("#NULL!")]
    [InlineData("#SPILL!")]
    [InlineData("#CALC!")]
    public void ExcelErrorsHaveLocationAndCorrection(string error)
    {
        var result = new ValidationResult();
        SimulationValidation.ValidateValue(result, "Forecast Cost (Budget!B5)", new ExcelCellError(error));
        Assert.False(result.IsValid);
        Assert.Contains(error, result.UserMessage);
        Assert.Contains("Budget!B5", result.UserMessage);
        Assert.Contains("Correct", result.UserMessage);
    }

    [Fact]
    public void ComErrorsAreNotNumbersButOrdinaryErrorCodeNumbersAre()
    {
        Assert.False(SimulationValidation.TryNumber(new ErrorWrapper(unchecked((int)0x800A07FA)), out _));
        Assert.False(SimulationValidation.TryNumber(unchecked((int)0x800A07FA), out _));
        Assert.True(SimulationValidation.TryNumber(2042, out double value));
        Assert.Equal(2042, value);
    }

    [Fact]
    public void MissingModelAndTrialSettingsAreReportedTogether()
    {
        var result = SimulationValidation.ValidateModel(0, [], []);
        Assert.Equal(3, result.Errors.Count);
    }

    [Fact]
    public void MissingReferencesAreReported()
    {
        var result = SimulationValidation.ValidateModel(1,
            [new() { Distribution = DistributionType.Normal, Parameter2 = 1 }], [new()]);
        Assert.Equal(2, result.Errors.Count);
    }

    [Fact]
    public void UnexpectedExceptionIsKeptOutOfUserMessage()
    {
        var result = new SimulationExecutionResult { DiagnosticException = new COMException("Secret internal stack details") };
        Assert.DoesNotContain("Secret", result.UserMessage);
        Assert.DoesNotContain("COM", result.UserMessage);
        Assert.NotNull(result.DiagnosticException);
    }
}
