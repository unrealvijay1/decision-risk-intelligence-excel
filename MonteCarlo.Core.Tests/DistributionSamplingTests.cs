using System;
using Xunit;

namespace MonteCarlo.Core.Tests
{
    public class DistributionSamplingTests
    {
        // =========================================================
        // NORMAL
        // =========================================================

        [Fact]
        public void NormalDistribution_ProducesExpectedMeanAndStandardDeviation()
        {
            const int samples =
                100000;

            const double expectedMean =
                100.0;

            const double expectedSd =
                20.0;


            double sum =
                0;

            double sumSquares =
                0;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                double value =
                    NormalDistribution.Sample(
                        expectedMean,
                        expectedSd);


                sum +=
                    value;


                sumSquares +=
                    value *
                    value;
            }


            double actualMean =
                sum /
                samples;


            double variance =
                sumSquares /
                samples
                -
                actualMean *
                actualMean;


            double actualSd =
                Math.Sqrt(
                    variance);


            Assert.InRange(
                actualMean,
                99.0,
                101.0);


            Assert.InRange(
                actualSd,
                19.0,
                21.0);
        }


        // =========================================================
        // UNIFORM
        // =========================================================

        [Fact]
        public void UniformDistribution_ProducesValuesInsideRange()
        {
            const int samples =
                50000;

            const double minimum =
                10.0;

            const double maximum =
                30.0;


            double sum =
                0;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                double value =
                    UniformDistribution.Sample(
                        minimum,
                        maximum);


                Assert.InRange(
                    value,
                    minimum,
                    maximum);


                sum +=
                    value;
            }


            double actualMean =
                sum /
                samples;


            double expectedMean =
                (
                    minimum +
                    maximum
                )
                /
                2.0;


            Assert.InRange(
                actualMean,
                expectedMean - 0.25,
                expectedMean + 0.25);
        }


        // =========================================================
        // TRIANGULAR
        // =========================================================

        [Fact]
        public void TriangularDistribution_ProducesExpectedMean()
        {
            const int samples =
                100000;

            const double minimum =
                10.0;

            const double mode =
                20.0;

            const double maximum =
                40.0;


            double sum =
                0;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                double value =
                    TriangularDistribution.Sample(
                        minimum,
                        mode,
                        maximum);


                Assert.InRange(
                    value,
                    minimum,
                    maximum);


                sum +=
                    value;
            }


            double actualMean =
                sum /
                samples;


            double expectedMean =
                (
                    minimum +
                    mode +
                    maximum
                )
                /
                3.0;


            Assert.InRange(
                actualMean,
                expectedMean - 0.35,
                expectedMean + 0.35);
        }


        // =========================================================
        // PERT
        // =========================================================

        [Fact]
        public void PertDistribution_ProducesExpectedMean()
        {
            const int samples =
                100000;

            const double minimum =
                10.0;

            const double mostLikely =
                20.0;

            const double maximum =
                40.0;


            double sum =
                0;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                double value =
                    PertDistribution.Sample(
                        minimum,
                        mostLikely,
                        maximum);


                Assert.InRange(
                    value,
                    minimum,
                    maximum);


                sum +=
                    value;
            }


            // Standard PERT expected value:
            //
            // (Min + 4*MostLikely + Max) / 6

            double expectedMean =
                (
                    minimum
                    +
                    4.0 *
                    mostLikely
                    +
                    maximum
                )
                /
                6.0;


            double actualMean =
                sum /
                samples;


            Assert.InRange(
                actualMean,
                expectedMean - 0.35,
                expectedMean + 0.35);
        }


        // =========================================================
        // LOGNORMAL
        // =========================================================

        [Fact]
        public void LognormalDistribution_ProducesExpectedMean()
        {
            const int samples =
                100000;

            const double logMean =
                9.0;

            const double logSd =
                0.4;


            double sum =
                0;


            for (
                int i = 0;
                i < samples;
                i++)
            {
                double value =
                    LognormalDistribution.Sample(
                        logMean,
                        logSd);


                Assert.True(
                    value > 0);


                sum +=
                    value;
            }


            double actualMean =
                sum /
                samples;


            double expectedMean =
                Math.Exp(
                    logMean
                    +
                    logSd *
                    logSd /
                    2.0);


            double tolerance =
                expectedMean *
                0.03;


            Assert.InRange(
                actualMean,
                expectedMean - tolerance,
                expectedMean + tolerance);
        }
    }
}