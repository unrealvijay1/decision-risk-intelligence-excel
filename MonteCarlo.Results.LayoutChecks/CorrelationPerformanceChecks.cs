using MonteCarlo.Core;
using System.Diagnostics;

internal static class CorrelationPerformanceChecks
{
    public static void Run()
    {
        const int trials = 50000;
        foreach (var kind in new[] { DistributionKind.Normal, DistributionKind.Pert, DistributionKind.Beta })
        foreach (bool correlated in new[] { false, true })
        {
            double[] parameters = kind switch { DistributionKind.Normal => [100, 15, 0, 0], DistributionKind.Pert => [0, 5, 10, 0], _ => [0, 10, 2, 5] };
            var sampler = new GaussianDependence(["a", "b"], [new("a", "b", .8)]);
            var z = new double[2]; var random = new Random(42); double sum = 0;
            var clock = Stopwatch.StartNew();
            for (int i = 0; i < trials; i++)
            {
                if (correlated) sampler.NextGaussian(random, z);
                for (int j = 0; j < 2; j++)
                    sum += correlated ? DistributionQuantile.FromGaussian(kind, z[j], parameters[0], parameters[1], parameters[2], parameters[3])
                        : DistributionSampler.Sample(kind, parameters[0], parameters[1], parameters[2], parameters[3], random);
            }
            Console.WriteLine($"PERFORMANCE {kind} correlated={correlated}: {trials:N0} trials / two inputs, {clock.Elapsed.TotalSeconds:F3}s; checksum {sum:F2}");
        }
    }
}
