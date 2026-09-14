using System;
using System.Collections.Generic;
using System.Linq;

namespace MonteCarlo.Core
{
    // =============================================================
    // DISTRIBUTIONS THAT CAN CURRENTLY BE FITTED
    // =============================================================

    public enum FittedDistributionType
    {
        Normal,
        Lognormal,
        Uniform,
        Triangular
    }


    // =============================================================
    // RESULT FROM FITTING ONE DISTRIBUTION
    // =============================================================

    public class DistributionFitResult
    {
        public FittedDistributionType Distribution { get; set; }

        public double Parameter1 { get; set; }

        public double Parameter2 { get; set; }

        public double Parameter3 { get; set; }

        public double LogLikelihood { get; set; }

        public double AIC { get; set; }

        public double KSStatistic { get; set; }

        public int Rank { get; set; }


        public string ParameterDescription
        {
            get
            {
                switch (Distribution)
                {
                    case FittedDistributionType.Normal:

                        return
                            $"Mean = {Parameter1:N4}, " +
                            $"Std Dev = {Parameter2:N4}";


                    case FittedDistributionType.Lognormal:

                        return
                            $"Log Mean = {Parameter1:N4}, " +
                            $"Log Std Dev = {Parameter2:N4}";


                    case FittedDistributionType.Uniform:

                        return
                            $"Min = {Parameter1:N4}, " +
                            $"Max = {Parameter2:N4}";


                    case FittedDistributionType.Triangular:

                        return
                            $"Min = {Parameter1:N4}, " +
                            $"Mode = {Parameter2:N4}, " +
                            $"Max = {Parameter3:N4}";


                    default:

                        return "";
                }
            }
        }
    }


    // =============================================================
    // DISTRIBUTION FITTER
    // =============================================================

    public static class DistributionFitter
    {
        // =========================================================
        // FIT ALL SUPPORTED DISTRIBUTIONS
        // =========================================================

        public static List<DistributionFitResult> FitAll(
            IEnumerable<double> inputData)
        {
            if (inputData == null)
            {
                throw new ArgumentNullException(
                    nameof(inputData));
            }


            double[] data =
                inputData
                    .Where(
                        x =>
                            !double.IsNaN(x)
                            &&
                            !double.IsInfinity(x))
                    .ToArray();


            if (data.Length < 5)
            {
                throw new ArgumentException(
                    "At least 5 valid numeric observations are required.");
            }


            double min =
                data.Min();

            double max =
                data.Max();


            if (min == max)
            {
                throw new ArgumentException(
                    "The selected data has no variation.");
            }


            List<DistributionFitResult> results =
                new List<DistributionFitResult>();


            // -----------------------------------------------------
            // NORMAL
            // -----------------------------------------------------

            TryAdd(
                results,
                () => FitNormal(data));


            // -----------------------------------------------------
            // LOGNORMAL
            //
            // Only valid when every observation is > 0.
            // -----------------------------------------------------

            if (data.All(x => x > 0))
            {
                TryAdd(
                    results,
                    () => FitLognormal(data));
            }


            // -----------------------------------------------------
            // UNIFORM
            // -----------------------------------------------------

            TryAdd(
                results,
                () => FitUniform(data));


            // -----------------------------------------------------
            // TRIANGULAR
            // -----------------------------------------------------

            TryAdd(
                results,
                () => FitTriangular(data));


            if (results.Count == 0)
            {
                throw new InvalidOperationException(
                    "No distribution could be fitted to the selected data.");
            }


            // -----------------------------------------------------
            // RANK
            //
            // Lower AIC is better.
            //
            // KS is used as the secondary tie-breaker.
            // -----------------------------------------------------

            List<DistributionFitResult> ranked =
                results
                    .OrderBy(
                        x => x.AIC)
                    .ThenBy(
                        x => x.KSStatistic)
                    .ToList();


            for (
                int i = 0;
                i < ranked.Count;
                i++)
            {
                ranked[i].Rank =
                    i + 1;
            }


            return ranked;
        }


        // =========================================================
        // NORMAL
        //
        // Parameter1 = mean
        // Parameter2 = standard deviation
        // =========================================================

        private static DistributionFitResult FitNormal(
            double[] data)
        {
            int n =
                data.Length;


            double mean =
                data.Average();


            // Maximum-likelihood standard deviation uses n,
            // not n - 1.
            double variance =
                data
                    .Select(
                        x =>
                            (x - mean) *
                            (x - mean))
                    .Sum()
                /
                n;


            double standardDeviation =
                Math.Sqrt(
                    variance);


            if (standardDeviation <= 0)
            {
                throw new InvalidOperationException(
                    "Normal distribution standard deviation is zero.");
            }


            double logLikelihood =
                0;


            double logConstant =
                -Math.Log(
                    standardDeviation *
                    Math.Sqrt(
                        2.0 *
                        Math.PI));


            for (
                int i = 0;
                i < n;
                i++)
            {
                double z =
                    (data[i] - mean) /
                    standardDeviation;


                logLikelihood +=
                    logConstant
                    -
                    0.5 *
                    z *
                    z;
            }


            int parameterCount =
                2;


            double aic =
                CalculateAIC(
                    parameterCount,
                    logLikelihood);


            double ks =
                CalculateKS(
                    data,
                    x =>
                        NormalCDF(
                            x,
                            mean,
                            standardDeviation));


            return
                new DistributionFitResult
                {
                    Distribution =
                        FittedDistributionType.Normal,

                    Parameter1 =
                        mean,

                    Parameter2 =
                        standardDeviation,

                    Parameter3 =
                        0,

                    LogLikelihood =
                        logLikelihood,

                    AIC =
                        aic,

                    KSStatistic =
                        ks
                };
        }


        // =========================================================
        // LOGNORMAL
        //
        // Parameter1 = mean of ln(X)
        // Parameter2 = standard deviation of ln(X)
        //
        // Requires all values > 0.
        // =========================================================

        private static DistributionFitResult FitLognormal(
            double[] data)
        {
            if (data.Any(x => x <= 0))
            {
                throw new InvalidOperationException(
                    "Lognormal distribution requires positive values.");
            }


            int n =
                data.Length;


            double[] logData =
                data
                    .Select(x => Math.Log(x))
                    .ToArray();


            double logMean =
                logData.Average();


            double logVariance =
                logData
                    .Select(
                        x =>
                            (x - logMean) *
                            (x - logMean))
                    .Sum()
                /
                n;


            double logStandardDeviation =
                Math.Sqrt(
                    logVariance);


            if (logStandardDeviation <= 0)
            {
                throw new InvalidOperationException(
                    "Lognormal standard deviation is zero.");
            }


            double logLikelihood =
                0;


            double constant =
                Math.Log(
                    logStandardDeviation *
                    Math.Sqrt(
                        2.0 *
                        Math.PI));


            for (
                int i = 0;
                i < n;
                i++)
            {
                double x =
                    data[i];


                double logX =
                    Math.Log(
                        x);


                double z =
                    (logX - logMean) /
                    logStandardDeviation;


                logLikelihood +=
                    -Math.Log(x)
                    -
                    constant
                    -
                    0.5 *
                    z *
                    z;
            }


            int parameterCount =
                2;


            double aic =
                CalculateAIC(
                    parameterCount,
                    logLikelihood);


            double ks =
                CalculateKS(
                    data,
                    x =>
                    {
                        if (x <= 0)
                        {
                            return 0;
                        }


                        return
                            StandardNormalCDF(
                                (
                                    Math.Log(x)
                                    -
                                    logMean
                                )
                                /
                                logStandardDeviation);
                    });


            return
                new DistributionFitResult
                {
                    Distribution =
                        FittedDistributionType.Lognormal,

                    Parameter1 =
                        logMean,

                    Parameter2 =
                        logStandardDeviation,

                    Parameter3 =
                        0,

                    LogLikelihood =
                        logLikelihood,

                    AIC =
                        aic,

                    KSStatistic =
                        ks
                };
        }


        // =========================================================
        // UNIFORM
        //
        // Parameter1 = minimum
        // Parameter2 = maximum
        // =========================================================

        private static DistributionFitResult FitUniform(
            double[] data)
        {
            int n =
                data.Length;


            double minimum =
                data.Min();


            double maximum =
                data.Max();


            double range =
                maximum -
                minimum;


            if (range <= 0)
            {
                throw new InvalidOperationException(
                    "Uniform distribution range is zero.");
            }


            double logLikelihood =
                -n *
                Math.Log(
                    range);


            // Treat min and max as estimated parameters.
            int parameterCount =
                2;


            double aic =
                CalculateAIC(
                    parameterCount,
                    logLikelihood);


            double ks =
                CalculateKS(
                    data,
                    x =>
                    {
                        if (x <= minimum)
                        {
                            return 0;
                        }


                        if (x >= maximum)
                        {
                            return 1;
                        }


                        return
                            (x - minimum)
                            /
                            range;
                    });


            return
                new DistributionFitResult
                {
                    Distribution =
                        FittedDistributionType.Uniform,

                    Parameter1 =
                        minimum,

                    Parameter2 =
                        maximum,

                    Parameter3 =
                        0,

                    LogLikelihood =
                        logLikelihood,

                    AIC =
                        aic,

                    KSStatistic =
                        ks
                };
        }


        // =========================================================
        // TRIANGULAR
        //
        // Parameter1 = minimum
        // Parameter2 = mode
        // Parameter3 = maximum
        //
        // For V1:
        // min/max are sample bounds.
        // Mode is estimated by maximum likelihood using
        // a one-dimensional grid search.
        // =========================================================

        private static DistributionFitResult FitTriangular(
            double[] data)
        {
            double minimum =
                data.Min();


            double maximum =
                data.Max();


            double range =
                maximum -
                minimum;


            if (range <= 0)
            {
                throw new InvalidOperationException(
                    "Triangular distribution range is zero.");
            }


            // Avoid exactly using the endpoints as mode.
            double epsilon =
                range *
                0.000001;


            double lowerMode =
                minimum +
                epsilon;


            double upperMode =
                maximum -
                epsilon;


            double bestMode =
                data.Average();


            bestMode =
                Math.Max(
                    lowerMode,
                    Math.Min(
                        upperMode,
                        bestMode));


            double bestLogLikelihood =
                double.NegativeInfinity;


            // V1 grid search.
            //
            // 1000 candidate mode values gives a sufficiently
            // precise fit for the Excel add-in at this stage.
            const int searchSteps =
                1000;


            for (
                int step = 0;
                step <= searchSteps;
                step++)
            {
                double fraction =
                    (double)step /
                    searchSteps;


                double mode =
                    lowerMode
                    +
                    fraction *
                    (
                        upperMode -
                        lowerMode
                    );


                double logLikelihood =
                    TriangularLogLikelihood(
                        data,
                        minimum,
                        mode,
                        maximum);


                if (
                    logLikelihood >
                    bestLogLikelihood)
                {
                    bestLogLikelihood =
                        logLikelihood;


                    bestMode =
                        mode;
                }
            }


            // min, mode, max
            int parameterCount =
                3;


            double aic =
                CalculateAIC(
                    parameterCount,
                    bestLogLikelihood);


            double ks =
                CalculateKS(
                    data,
                    x =>
                        TriangularCDF(
                            x,
                            minimum,
                            bestMode,
                            maximum));


            return
                new DistributionFitResult
                {
                    Distribution =
                        FittedDistributionType.Triangular,

                    Parameter1 =
                        minimum,

                    Parameter2 =
                        bestMode,

                    Parameter3 =
                        maximum,

                    LogLikelihood =
                        bestLogLikelihood,

                    AIC =
                        aic,

                    KSStatistic =
                        ks
                };
        }


        // =========================================================
        // TRIANGULAR LOG LIKELIHOOD
        // =========================================================

        private static double TriangularLogLikelihood(
            double[] data,
            double minimum,
            double mode,
            double maximum)
        {
            if (
                minimum >= maximum
                ||
                mode <= minimum
                ||
                mode >= maximum)
            {
                return
                    double.NegativeInfinity;
            }


            double logLikelihood =
                0;


            foreach (
                double x
                in data)
            {
                double density;


                if (x < minimum ||
                    x > maximum)
                {
                    return
                        double.NegativeInfinity;
                }


                if (x <= mode)
                {
                    density =
                        2.0 *
                        (x - minimum)
                        /
                        (
                            (maximum - minimum)
                            *
                            (mode - minimum)
                        );
                }
                else
                {
                    density =
                        2.0 *
                        (maximum - x)
                        /
                        (
                            (maximum - minimum)
                            *
                            (maximum - mode)
                        );
                }


                // Sample min/max have density zero exactly at
                // the endpoints. Use a tiny floor to keep the
                // numerical likelihood finite.
                density =
                    Math.Max(
                        density,
                        1e-300);


                logLikelihood +=
                    Math.Log(
                        density);
            }


            return
                logLikelihood;
        }


        // =========================================================
        // TRIANGULAR CDF
        // =========================================================

        private static double TriangularCDF(
            double x,
            double minimum,
            double mode,
            double maximum)
        {
            if (x <= minimum)
            {
                return 0;
            }


            if (x >= maximum)
            {
                return 1;
            }


            if (x <= mode)
            {
                return
                    (
                        (x - minimum)
                        *
                        (x - minimum)
                    )
                    /
                    (
                        (maximum - minimum)
                        *
                        (mode - minimum)
                    );
            }


            return
                1.0
                -
                (
                    (maximum - x)
                    *
                    (maximum - x)
                )
                /
                (
                    (maximum - minimum)
                    *
                    (maximum - mode)
                );
        }


        // =========================================================
        // KOLMOGOROV-SMIRNOV STATISTIC
        //
        // Smaller = better fit.
        // =========================================================

        private static double CalculateKS(
            double[] data,
            Func<double, double> cdf)
        {
            double[] sorted =
                data
                    .OrderBy(
                        x => x)
                    .ToArray();


            int n =
                sorted.Length;


            double maxDifference =
                0;


            for (
                int i = 0;
                i < n;
                i++)
            {
                double theoretical =
                    cdf(
                        sorted[i]);


                theoretical =
                    Math.Max(
                        0,
                        Math.Min(
                            1,
                            theoretical));


                double empiricalBefore =
                    (double)i /
                    n;


                double empiricalAfter =
                    (double)(i + 1) /
                    n;


                double difference1 =
                    Math.Abs(
                        theoretical -
                        empiricalBefore);


                double difference2 =
                    Math.Abs(
                        empiricalAfter -
                        theoretical);


                maxDifference =
                    Math.Max(
                        maxDifference,
                        Math.Max(
                            difference1,
                            difference2));
            }


            return
                maxDifference;
        }


        // =========================================================
        // AIC
        //
        // Lower = better.
        // =========================================================

        private static double CalculateAIC(
            int parameterCount,
            double logLikelihood)
        {
            return
                2.0 *
                parameterCount
                -
                2.0 *
                logLikelihood;
        }


        // =========================================================
        // NORMAL CDF
        // =========================================================

        private static double NormalCDF(
            double x,
            double mean,
            double standardDeviation)
        {
            if (standardDeviation <= 0)
            {
                return
                    x < mean
                        ? 0
                        : 1;
            }


            double z =
                (x - mean)
                /
                standardDeviation;


            return
                StandardNormalCDF(
                    z);
        }


        // =========================================================
        // STANDARD NORMAL CDF
        // =========================================================

        private static double StandardNormalCDF(
            double x)
        {
            return
                0.5 *
                (
                    1.0 +
                    Erf(
                        x /
                        Math.Sqrt(2.0))
                );
        }


        // =========================================================
        // ERROR FUNCTION APPROXIMATION
        // =========================================================

        private static double Erf(
            double x)
        {
            // Abramowitz and Stegun approximation.

            double sign =
                x < 0
                    ? -1.0
                    : 1.0;


            x =
                Math.Abs(
                    x);


            const double a1 =
                0.254829592;

            const double a2 =
                -0.284496736;

            const double a3 =
                1.421413741;

            const double a4 =
                -1.453152027;

            const double a5 =
                1.061405429;

            const double p =
                0.3275911;


            double t =
                1.0 /
                (
                    1.0 +
                    p *
                    x
                );


            double y =
                1.0
                -
                (
                    (
                        (
                            (
                                (
                                    a5 *
                                    t
                                    +
                                    a4
                                )
                                *
                                t
                                +
                                a3
                            )
                            *
                            t
                            +
                            a2
                        )
                        *
                        t
                        +
                        a1
                    )
                    *
                    t
                    *
                    Math.Exp(
                        -x *
                        x)
                );


            return
                sign *
                y;
        }


        // =========================================================
        // SAFE FIT
        // =========================================================

        private static void TryAdd(
            List<DistributionFitResult> results,
            Func<DistributionFitResult> fitter)
        {
            try
            {
                DistributionFitResult result =
                    fitter();


                if (
                    !double.IsNaN(
                        result.AIC)
                    &&
                    !double.IsInfinity(
                        result.AIC)
                    &&
                    !double.IsNaN(
                        result.KSStatistic)
                    &&
                    !double.IsInfinity(
                        result.KSStatistic))
                {
                    results.Add(
                        result);
                }
            }
            catch
            {
                // One failed candidate should not stop fitting
                // the remaining distributions.
            }
        }
    }
}