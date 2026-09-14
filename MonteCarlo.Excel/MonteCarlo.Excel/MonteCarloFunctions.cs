using ExcelDna.Integration;

namespace MonteCarlo.Excel
{
    public static class MonteCarloFunctions
    {
        // ---------------------------------------------------------
        // BASIC TEST FUNCTION
        // ---------------------------------------------------------

        [ExcelFunction(
            Name = "MC.ADD",
            Description = "Adds two numbers")]
        public static double Add(
            double a,
            double b)
        {
            return a + b;
        }


        // ---------------------------------------------------------
        // NORMAL DISTRIBUTION
        // ---------------------------------------------------------

        [ExcelFunction(
            Name = "MC.NORMAL",
            Description =
                "Returns a random value from a normal distribution",
            IsVolatile = true)]
        public static double Normal(
            double mean,
            double standardDeviation)
        {
            return
                MonteCarlo.Core
                    .NormalDistribution
                    .Sample(
                        mean,
                        standardDeviation);
        }


        // ---------------------------------------------------------
        // TRIANGULAR DISTRIBUTION
        // ---------------------------------------------------------

        [ExcelFunction(
            Name = "MC.TRIANGULAR",
            Description =
                "Returns a random value from a triangular distribution",
            IsVolatile = true)]
        public static double Triangular(
            double minimum,
            double mostLikely,
            double maximum)
        {
            return
                MonteCarlo.Core
                    .TriangularDistribution
                    .Sample(
                        minimum,
                        mostLikely,
                        maximum);
        }


        // ---------------------------------------------------------
        // PERT DISTRIBUTION
        // ---------------------------------------------------------

        [ExcelFunction(
            Name = "MC.PERT",
            Description =
                "Returns a random value from a PERT distribution",
            IsVolatile = true)]
        public static double Pert(
            double minimum,
            double mostLikely,
            double maximum)
        {
            return
                MonteCarlo.Core
                    .PertDistribution
                    .Sample(
                        minimum,
                        mostLikely,
                        maximum);
        }


        // ---------------------------------------------------------
        // NORMAL SIMULATION
        // ---------------------------------------------------------

        [ExcelFunction(
            Name = "MC.SIMULATE.NORMAL",
            Description =
                "Runs a Monte Carlo simulation using a normal distribution",
            IsVolatile = true)]
        public static object[,] SimulateNormal(
            double mean,
            double standardDeviation,
            int trials)
        {
            var result =
                MonteCarlo.Core
                    .MonteCarloSimulator
                    .RunNormal(
                        mean,
                        standardDeviation,
                        trials);

            return
                FormatSimulationResult(
                    result);
        }


        // ---------------------------------------------------------
        // TRIANGULAR SIMULATION
        // ---------------------------------------------------------

        [ExcelFunction(
            Name = "MC.SIMULATE.TRIANGULAR",
            Description =
                "Runs a Monte Carlo simulation using a triangular distribution",
            IsVolatile = true)]
        public static object[,] SimulateTriangular(
            double minimum,
            double mostLikely,
            double maximum,
            int trials)
        {
            var result =
                MonteCarlo.Core
                    .MonteCarloSimulator
                    .RunTriangular(
                        minimum,
                        mostLikely,
                        maximum,
                        trials);

            return
                FormatSimulationResult(
                    result);
        }


        // ---------------------------------------------------------
        // PERT SIMULATION
        // ---------------------------------------------------------

        [ExcelFunction(
            Name = "MC.SIMULATE.PERT",
            Description =
                "Runs a Monte Carlo simulation using a PERT distribution",
            IsVolatile = true)]
        public static object[,] SimulatePert(
            double minimum,
            double mostLikely,
            double maximum,
            int trials)
        {
            var result =
                MonteCarlo.Core
                    .MonteCarloSimulator
                    .RunPert(
                        minimum,
                        mostLikely,
                        maximum,
                        trials);

            return
                FormatSimulationResult(
                    result);
        }


        // ---------------------------------------------------------
        // RESULT FORMATTER
        // ---------------------------------------------------------

        private static object[,] FormatSimulationResult(
            MonteCarlo.Core.SimulationResult result)
        {
            return new object[,]
            {
                { "Metric", "Value" },
                { "Trials", result.Trials },
                { "Mean", result.Mean },
                { "Std Dev", result.StandardDeviation },
                { "Minimum", result.Minimum },
                { "Maximum", result.Maximum },
                { "P10", result.P10 },
                { "P50", result.P50 },
                { "P80", result.P80 },
                { "P90", result.P90 }
            };
        }
        [ExcelFunction(
    Name = "MC.PROJECT.COST",
    Description =
        "Runs a Monte Carlo simulation of total project cost",
    IsVolatile = true)]
        public static object[,] ProjectCost(
    double durationMin,
    double durationMostLikely,
    double durationMax,
    double rateMin,
    double rateMostLikely,
    double rateMax,
    int people,
    int trials)
        {
            var result =
                MonteCarlo.Core
                    .MonteCarloSimulator
                    .RunProjectCost(
                        durationMin,
                        durationMostLikely,
                        durationMax,
                        rateMin,
                        rateMostLikely,
                        rateMax,
                        people,
                        trials);

            return FormatSimulationResult(result);
        }
       
    }
}