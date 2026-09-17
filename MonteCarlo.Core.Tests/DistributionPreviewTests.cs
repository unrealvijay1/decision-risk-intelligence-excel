using MonteCarlo.Core;

namespace MonteCarlo.Core.Tests;

public class DistributionPreviewTests
{
    private static DensityPoint Nearest(DistributionPreviewResult result, double x) =>
        result.Points.MinBy(point => Math.Abs(point.X - x));

    public static TheoryData<DistributionKind, double, double, double, double, double> Cases => new()
    {
        { DistributionKind.Normal, 10, 2, 0, 0, 0.9999 },
        { DistributionKind.Triangular, 0, 3, 10, 0, 1.0 },
        { DistributionKind.Pert, 0, 3, 10, 0, 1.0 },
        { DistributionKind.Uniform, 2, 8, 0, 0, 1.0 },
        { DistributionKind.Lognormal, 1, 0.35, 0, 0, 0.9999 },
        { DistributionKind.Beta, 2, 8, 2, 3, 1.0 }
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void DensitiesAreFiniteNonnegativeAndIntegrateApproximatelyToExpectedMass(
        DistributionKind kind, double p1, double p2, double p3, double p4,
        double expectedArea)
    {
        // Runs in the platform-neutral Core test project without Excel.
        var result = DistributionPreview.Generate(kind, p1, p2, p3, p4);
        Assert.True(result.IsValid, result.ValidationMessage);
        Assert.Equal(1025, result.Points.Count);
        Assert.All(result.Points, point =>
        {
            Assert.True(double.IsFinite(point.X));
            Assert.True(double.IsFinite(point.Density));
            Assert.True(point.Density >= 0);
        });

        double area = 0;
        for (int i = 1; i < result.Points.Count; i++)
        {
            var previous = result.Points[i - 1];
            var current = result.Points[i];
            area += (current.X - previous.X) *
                (current.Density + previous.Density) / 2;
        }
        Assert.InRange(area, expectedArea - 0.015, expectedArea + 0.015);
    }

    [Theory]
    [InlineData(DistributionKind.Normal, 0, 0, 0, 0)]
    [InlineData(DistributionKind.Triangular, 0, 11, 10, 0)]
    [InlineData(DistributionKind.Pert, 0, -1, 10, 0)]
    [InlineData(DistributionKind.Uniform, 5, 5, 0, 0)]
    [InlineData(DistributionKind.Lognormal, 0, -1, 0, 0)]
    [InlineData(DistributionKind.Beta, 0, 1, 0, 2)]
    [InlineData(DistributionKind.Beta, 0, 1, 2, -1)]
    [InlineData(DistributionKind.Normal, double.NaN, 1, 0, 0)]
    public void InvalidParametersReturnMessage(DistributionKind kind,
        double p1, double p2, double p3, double p4)
    {
        var result = DistributionPreview.Generate(kind, p1, p2, p3, p4);
        Assert.False(result.IsValid);
        Assert.False(string.IsNullOrWhiteSpace(result.ValidationMessage));
        Assert.Empty(result.Points);
    }

    [Fact]
    public void SingularBetaShapesRemainFiniteWithoutEndpointEvaluation()
    {
        var result = DistributionPreview.Generate(DistributionKind.Beta, 0, 1, 0.5, 0.5);
        Assert.True(result.IsValid, result.ValidationMessage);
        Assert.True(result.Points[0].X > result.MinimumX);
        Assert.True(result.Points[^1].X < result.MaximumX);
        Assert.All(result.Points, point => Assert.True(double.IsFinite(point.Density)));
    }

    [Fact]
    public void NormalPeakAndAxisUseSpecifiedMeanAndStandardDeviation()
    {
        var result = DistributionPreview.Generate(DistributionKind.Normal, 12, 3);
        Assert.True(result.IsValid);
        Assert.Equal(0, result.MinimumX, 10);
        Assert.Equal(24, result.MaximumX, 10);
        Assert.InRange(result.Points.MaxBy(p => p.Density).X, 11.98, 12.02);
        Assert.InRange(Nearest(result, 12).Density, 0.132, 0.134);
    }

    [Fact]
    public void UniformDensityIsConstantOnSpecifiedSupport()
    {
        var result = DistributionPreview.Generate(DistributionKind.Uniform, 3, 11);
        Assert.True(result.IsValid);
        Assert.Equal(3, result.MinimumX);
        Assert.Equal(11, result.MaximumX);
        Assert.All(result.Points, p => Assert.Equal(0.125, p.Density, 10));
    }

    [Fact]
    public void TriangularPeakTracksAsymmetricMostLikelyValue()
    {
        var result = DistributionPreview.Generate(DistributionKind.Triangular, 0, 7, 10);
        Assert.True(result.IsValid);
        Assert.Equal(0, result.MinimumX);
        Assert.Equal(10, result.MaximumX);
        Assert.InRange(result.Points.MaxBy(p => p.Density).X, 6.98, 7.02);
        Assert.True(Nearest(result, 7).Density > Nearest(result, 3).Density);
    }

    [Fact]
    public void PertModeTracksAsymmetricMostLikelyValue()
    {
        var result = DistributionPreview.Generate(DistributionKind.Pert, 0, 7, 10);
        Assert.True(result.IsValid);
        Assert.Equal(0, result.MinimumX);
        Assert.Equal(10, result.MaximumX);
        Assert.InRange(result.Points.MaxBy(p => p.Density).X, 6.95, 7.05);
        Assert.True(Nearest(result, 8).Density > Nearest(result, 2).Density);
    }

    [Fact]
    public void LognormalAxisAndPeakUseLogScaleParameters()
    {
        const double mu = 1;
        const double sigma = 0.35;
        var result = DistributionPreview.Generate(DistributionKind.Lognormal, mu, sigma);
        Assert.True(result.IsValid);
        Assert.Equal(Math.Exp(mu - 4 * sigma), result.MinimumX, 10);
        Assert.Equal(Math.Exp(mu + 4 * sigma), result.MaximumX, 10);
        double expectedMode = Math.Exp(mu - sigma * sigma);
        Assert.InRange(result.Points.MaxBy(p => p.Density).X,
            expectedMode - 0.02, expectedMode + 0.02);
        Assert.True(Nearest(result, Math.Exp(mu - sigma)).Density >
                    Nearest(result, Math.Exp(mu + sigma)).Density);
    }

    [Fact]
    public void BetaShapeParametersAreNotSwapped()
    {
        var result = DistributionPreview.Generate(DistributionKind.Beta, 2, 12, 2, 5);
        Assert.True(result.IsValid);
        Assert.Equal(2, result.MinimumX);
        Assert.Equal(12, result.MaximumX);
        Assert.InRange(result.Points.MaxBy(p => p.Density).X, 3.9, 4.1);
        Assert.True(Nearest(result, 4.5).Density > Nearest(result, 9.5).Density);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void EveryNonFiniteParameterIsRejected(double invalid)
    {
        Assert.False(DistributionPreview.Generate(DistributionKind.Normal, invalid, 1).IsValid);
        Assert.False(DistributionPreview.Generate(DistributionKind.Normal, 0, invalid).IsValid);
        Assert.False(DistributionPreview.Generate(DistributionKind.Beta, 0, 1, invalid, 2).IsValid);
        Assert.False(DistributionPreview.Generate(DistributionKind.Beta, 0, 1, 2, invalid).IsValid);
    }
}
