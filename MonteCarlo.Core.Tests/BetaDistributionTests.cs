using System;
using Xunit;

namespace MonteCarlo.Core.Tests
{
    public class BetaDistributionTests
    {
        // =========================================================
        // STANDARD BETA RANGE
        // =========================================================

        [Fact]
        public void StandardBeta_ProducesValuesBetweenZeroAndOne()
        {
            const int samples =
                50000;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                double value =
                    BetaDistribution.Sample(
                        2.0,
                        5.0);


                Assert.InRange(
                    value,
                    0.0,
                    1.0);
            }
        }


        // =========================================================
        // STANDARD BETA MEAN
        //
        // Mean = alpha / (alpha + beta)
        // =========================================================

        [Fact]
        public void StandardBeta_ProducesExpectedMean()
        {
            const int samples =
                100000;


            const double alpha =
                2.0;


            const double beta =
                5.0;


            double sum =
                0;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                sum +=
                    BetaDistribution.Sample(
                        alpha,
                        beta);
            }


            double actualMean =
                sum /
                samples;


            double expectedMean =
                alpha /
                (
                    alpha +
                    beta
                );


            Assert.InRange(
                actualMean,
                expectedMean - 0.01,
                expectedMean + 0.01);
        }


        // =========================================================
        // SCALED BETA RANGE
        // =========================================================

        [Fact]
        public void ScaledBeta_ProducesValuesInsideSpecifiedRange()
        {
            const int samples =
                50000;


            const double minimum =
                70.0;


            const double maximum =
                100.0;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                double value =
                    BetaDistribution.Sample(
                        minimum,
                        maximum,
                        5.0,
                        2.0);


                Assert.InRange(
                    value,
                    minimum,
                    maximum);
            }
        }


        // =========================================================
        // SCALED BETA MEAN
        //
        // Mean =
        // min + (max-min) * alpha/(alpha+beta)
        // =========================================================

        [Fact]
        public void ScaledBeta_ProducesExpectedMean()
        {
            const int samples =
                100000;


            const double minimum =
                70.0;


            const double maximum =
                100.0;


            const double alpha =
                5.0;


            const double beta =
                2.0;


            double sum =
                0;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                sum +=
                    BetaDistribution.Sample(
                        minimum,
                        maximum,
                        alpha,
                        beta);
            }


            double actualMean =
                sum /
                samples;


            double expectedMean =
                minimum
                +
                (
                    maximum -
                    minimum
                )
                *
                alpha
                /
                (
                    alpha +
                    beta
                );


            Assert.InRange(
                actualMean,
                expectedMean - 0.20,
                expectedMean + 0.20);
        }


        // =========================================================
        // INVALID ALPHA
        // =========================================================

        [Fact]
        public void StandardBeta_Throws_WhenAlphaIsZeroOrNegative()
        {
            Assert.Throws<
                ArgumentOutOfRangeException>(
                () =>
                    BetaDistribution.Sample(
                        0.0,
                        2.0));


            Assert.Throws<
                ArgumentOutOfRangeException>(
                () =>
                    BetaDistribution.Sample(
                        -1.0,
                        2.0));
        }


        // =========================================================
        // INVALID BETA
        // =========================================================

        [Fact]
        public void StandardBeta_Throws_WhenBetaIsZeroOrNegative()
        {
            Assert.Throws<
                ArgumentOutOfRangeException>(
                () =>
                    BetaDistribution.Sample(
                        2.0,
                        0.0));


            Assert.Throws<
                ArgumentOutOfRangeException>(
                () =>
                    BetaDistribution.Sample(
                        2.0,
                        -1.0));
        }


        // =========================================================
        // INVALID RANGE
        // =========================================================

        [Fact]
        public void ScaledBeta_Throws_WhenMaximumIsNotGreaterThanMinimum()
        {
            Assert.Throws<
                ArgumentException>(
                () =>
                    BetaDistribution.Sample(
                        100.0,
                        100.0,
                        2.0,
                        5.0));


            Assert.Throws<
                ArgumentException>(
                () =>
                    BetaDistribution.Sample(
                        100.0,
                        90.0,
                        2.0,
                        5.0));
        }
    }
}