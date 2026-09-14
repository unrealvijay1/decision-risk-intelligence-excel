namespace MonteCarlo.Core
{
    public class SimulationResult
    {
        public int Trials { get; set; }

        public double Mean { get; set; }

        public double StandardDeviation { get; set; }

        public double Minimum { get; set; }

        public double Maximum { get; set; }

        public double P10 { get; set; }

        public double P50 { get; set; }

        public double P80 { get; set; }

        public double P90 { get; set; }
    }
}