using System;
using System.Linq;
using Xunit;

namespace MonteCarlo.Core.Tests
{
    public class DistributionFittingTests
    {
        // =========================================================
        // NORMAL DATA
        // =========================================================

        [Fact]
        public void DistributionFitter_RecognizesNormalData()
        {
            Random random =
                new Random(12345);


            double[] data =
                new double[5000];


            for (
                int i = 0;
                i < data.Length;
                i++)
            {
                data[i] =
                    SampleNormal(
                        random,
                        100.0,
                        15.0);
            }


            var results =
                DistributionFitter.FitAll(
                    data);


            DistributionFitResult best =
                results.First();


            Assert.Equal(
                FittedDistributionType.Normal,
                best.Distribution);


            Assert.InRange(
                best.Parameter1,
                98.5,
                101.5);


            Assert.InRange(
                best.Parameter2,
                14.0,
                16.0);
        }


        // =========================================================
        // LOGNORMAL DATA
        // =========================================================

        [Fact]
        public void DistributionFitter_RecognizesLognormalData()
        {
            Random random =
                new Random(22222);


            const double logMean =
                8.5;


            const double logSd =
                0.45;


            double[] data =
                new double[5000];


            for (
                int i = 0;
                i < data.Length;
                i++)
            {
                double normal =
                    SampleNormal(
                        random,
                        logMean,
                        logSd);


                data[i] =
                    Math.Exp(
                        normal);
            }


            var results =
                DistributionFitter.FitAll(
                    data);


            DistributionFitResult best =
                results.First();


            Assert.Equal(
                FittedDistributionType.Lognormal,
                best.Distribution);


            Assert.InRange(
                best.Parameter1,
                8.40,
                8.60);


            Assert.InRange(
                best.Parameter2,
                0.40,
                0.50);
        }


        // =========================================================
        // UNIFORM DATA
        // =========================================================

        [Fact]
        public void DistributionFitter_RecognizesUniformData()
        {
            Random random =
                new Random(33333);


            const double minimum =
                10.0;


            const double maximum =
                50.0;


            double[] data =
                new double[5000];


            for (
                int i = 0;
                i < data.Length;
                i++)
            {
                data[i] =
                    minimum
                    +
                    random.NextDouble()
                    *
                    (
                        maximum -
                        minimum
                    );
            }


            var results =
                DistributionFitter.FitAll(
                    data);


            DistributionFitResult best =
                results.First();


            Assert.Equal(
                FittedDistributionType.Uniform,
                best.Distribution);


            Assert.InRange(
                best.Parameter1,
                9.5,
                11.0);


            Assert.InRange(
                best.Parameter2,
                49.0,
                50.5);
        }


        // =========================================================
        // TRIANGULAR DATA
        // =========================================================

        [Fact]
        public void DistributionFitter_RecognizesTriangularData()
        {
            Random random =
                new Random(44444);


            const double minimum =
                10.0;


            const double mode =
                25.0;


            const double maximum =
                60.0;


            double[] data =
                new double[5000];


            for (
                int i = 0;
                i < data.Length;
                i++)
            {
                data[i] =
                    SampleTriangular(
                        random,
                        minimum,
                        mode,
                        maximum);
            }


            var results =
                DistributionFitter.FitAll(
                    data);


            DistributionFitResult triangular =
                results.First(
                    x =>
                        x.Distribution ==
                        FittedDistributionType.Triangular);


            Assert.InRange(
                triangular.Parameter1,
                9.5,
                11.5);


            Assert.InRange(
                triangular.Parameter2,
                22.0,
                28.0);


            Assert.InRange(
                triangular.Parameter3,
                58.0,
                60.5);
        }


        // =========================================================
        // NEGATIVE DATA
        //
        // LOGNORMAL MUST NOT BE OFFERED
        // =========================================================

        [Fact]
        public void DistributionFitter_DoesNotFitLognormal_WhenDataContainsNegativeValues()
        {
            double[] data =
            {
                -10,
                -5,
                -2,
                0,
                3,
                6,
                8,
                10,
                12,
                15
            };


            var results =
                DistributionFitter.FitAll(
                    data);


            Assert.DoesNotContain(
                results,
                x =>
                    x.Distribution ==
                    FittedDistributionType.Lognormal);
        }


        // =========================================================
        // TOO FEW OBSERVATIONS
        // =========================================================

        [Fact]
        public void DistributionFitter_Throws_WhenTooFewObservations()
        {
            double[] data =
            {
                1,
                2,
                3,
                4
            };


            Assert.Throws<ArgumentException>(
                () =>
                    DistributionFitter.FitAll(
                        data));
        }


        // =========================================================
        // NO VARIATION
        // =========================================================

        [Fact]
        public void DistributionFitter_Throws_WhenDataHasNoVariation()
        {
            double[] data =
            {
                10,
                10,
                10,
                10,
                10,
                10
            };


            Assert.Throws<ArgumentException>(
                () =>
                    DistributionFitter.FitAll(
                        data));
        }


        // =========================================================
        // HELPER: NORMAL SAMPLING
        // =========================================================

        private static double SampleNormal(
            Random random,
            double mean,
            double standardDeviation)
        {
            double u1 =
                1.0 -
                random.NextDouble();


            double u2 =
                1.0 -
                random.NextDouble();


            double standardNormal =
                Math.Sqrt(
                    -2.0 *
                    Math.Log(
                        u1))
                *
                Math.Cos(
                    2.0 *
                    Math.PI *
                    u2);


            return
                mean
                +
                standardDeviation
                *
                standardNormal;
        }


        // =========================================================
        // HELPER: TRIANGULAR SAMPLING
        // =========================================================

        private static double SampleTriangular(
            Random random,
            double minimum,
            double mode,
            double maximum)
        {
            double u =
                random.NextDouble();


            double threshold =
                (
                    mode -
                    minimum
                )
                /
                (
                    maximum -
                    minimum
                );


            if (u < threshold)
            {
                return
                    minimum
                    +
                    Math.Sqrt(
                        u
                        *
                        (
                            maximum -
                            minimum
                        )
                        *
                        (
                            mode -
                            minimum
                        ));
            }


            return
                maximum
                -
                Math.Sqrt(
                    (
                        1.0 -
                        u
                    )
                    *
                    (
                        maximum -
                        minimum
                    )
                    *
                    (
                        maximum -
                        mode
                    ));
        }
    }
}