using MonteCarlo.Core;

namespace MonteCarlo.Excel;

public static class SimulationExecution
{
    public static SimulationExecutionResult Run(int trials, IReadOnlyList<AssumptionDefinition> assumptions,
        IReadOnlyList<ForecastDefinition> forecasts, ISimulationWorkbook workbook, ValidationResult? savedValidation = null)
    {
        var outcome = new SimulationExecutionResult();
        if (savedValidation != null) outcome.Validation.Errors.AddRange(savedValidation.Errors);
        outcome.Validation.Errors.AddRange(SimulationValidation.ValidateModel(trials, assumptions, forecasts).Errors);
        if (!outcome.Validation.IsValid) return outcome;
        var inputs = new Dictionary<AssumptionDefinition, ISimulationCell>();
        var outputs = new Dictionary<ForecastDefinition, ISimulationCell>();
        var identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (var assumption in assumptions)
            {
                string location = SimulationValidation.Location("Assumption", assumption.Name, assumption.SheetName, assumption.CellAddress);
                var cell = Resolve(assumption.SheetName, assumption.CellAddress, location);
                if (cell == null) continue;
                if (!identities.Add(cell.Identity))
                    outcome.Validation.Add(location, "Multiple assumptions use the same input cell.", "Keep only one assumption for this cell.");
                if (cell.InputProblem != null)
                    outcome.Validation.Add(location, cell.InputProblem, "Choose a single editable input cell outside merged, array, or spill ranges.");
                SimulationValidation.ValidateValue(outcome.Validation, location, cell.ReadValue());
                inputs[assumption] = cell;
            }
            foreach (var forecast in forecasts)
            {
                string location = SimulationValidation.Location("Forecast", forecast.Name, forecast.SheetName, forecast.CellAddress);
                var cell = Resolve(forecast.SheetName, forecast.CellAddress, location);
                if (cell == null) continue;
                SimulationValidation.ValidateValue(outcome.Validation, location, cell.ReadValue());
                outputs[forecast] = cell;
            }
        }
        catch (Exception ex)
        {
            outcome.DiagnosticException = ex;
            return outcome;
        }
        if (!outcome.Validation.IsValid) return outcome;

        ISimulationCell? Resolve(string sheet, string address, string location)
        {
            try { return workbook.Resolve(sheet, address); }
            catch (SimulationCellAccessException)
            {
                outcome.Validation.Add(location, "The worksheet or single-cell reference cannot be used.",
                    "Check that the worksheet exists and redefine this item using one cell in this workbook.");
                return null;
            }
        }

        // Capture every original before changing any cell or application setting.
        var originals = new List<(ISimulationCell Cell, SimulationCellContent Content)>();
        bool screen;
        bool events;
        object? status;
        var assumptionSamples = new Dictionary<string, double[]>();
        var forecastSamples = new Dictionary<ForecastDefinition, double[]>();
        var names = new Dictionary<AssumptionDefinition, string>();
        try
        {
            foreach (var assumption in assumptions)
            {
                originals.Add((inputs[assumption], inputs[assumption].Capture()));
                string desiredName = GetAssumptionDisplayName(assumption);
                string name = desiredName;
                int suffix = 2;
                while (names.Values.Any(x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase)))
                    name = $"{desiredName} ({suffix++})";
                names[assumption] = name;
                assumptionSamples[name] = new double[trials];
            }
            foreach (var forecast in forecasts) forecastSamples[forecast] = new double[trials];
            screen = workbook.ScreenUpdating;
            events = workbook.EnableEvents;
            status = workbook.StatusBar;
        }
        catch (Exception ex)
        {
            outcome.DiagnosticException = ex;
            return outcome;
        }

        try
        {
            workbook.ScreenUpdating = false;
            workbook.EnableEvents = false;
            int progressInterval = Math.Max(1, trials / 100);
            bool stop = false;
            for (int i = 0; i < trials && !stop; i++)
            {
                foreach (var assumption in assumptions)
                {
                    double sample = GenerateSample(assumption);
                    if (!double.IsFinite(sample))
                    {
                        outcome.Validation.Add(inputs[assumption].Identity,
                            $"The assumption produced a value outside the supported numeric range on trial {i + 1}.",
                            "Review the distribution parameters and reduce their scale.");
                        stop = true;
                        break;
                    }
                    assumptionSamples[names[assumption]][i] = sample;
                    inputs[assumption].WriteSample(sample);
                }
                if (stop) break;
                workbook.Calculate();
                foreach (var forecast in forecasts)
                {
                    object? value = outputs[forecast].ReadValue();
                    SimulationValidation.ValidateValue(outcome.Validation,
                        $"Forecast {forecast.Name} ({outputs[forecast].Identity}), trial {i + 1}", value);
                    if (!outcome.Validation.IsValid) { stop = true; break; }
                    SimulationValidation.TryNumber(value, out double number);
                    forecastSamples[forecast][i] = number;
                }
                if (i % progressInterval == 0 || i == trials - 1)
                    workbook.StatusBar = $"Monte Carlo Simulation: {(int)((i + 1) * 100.0 / trials)}% ({i + 1:N0}/{trials:N0})";
            }
        }
        catch (Exception ex) { outcome.DiagnosticException = ex; }
        finally
        {
            // A failed restoration must not prevent the remaining cells/settings from being attempted.
            foreach (var original in originals)
                Restore(original.Cell.Identity, () => original.Cell.Restore(original.Content));
            Restore("workbook recalculation", workbook.Calculate);
            Restore("Screen Updating", () => workbook.ScreenUpdating = screen);
            Restore("Enable Events", () => workbook.EnableEvents = events);
            Restore("Status Bar", () => workbook.StatusBar = status);
        }
        if (outcome.Validation.IsValid && outcome.DiagnosticException == null && outcome.RestorationFailures.Count == 0)
        {
            try { outcome.Result = new SimulationRunResult(trials, forecastSamples, assumptionSamples); }
            catch (Exception ex) { outcome.DiagnosticException = ex; }
        }
        return outcome;

        void Restore(string location, Action restore)
        {
            Exception? last = null;
            for (int attempt = 0; attempt < 2; attempt++)
            {
                try { restore(); return; }
                catch (Exception ex) { last = ex; }
            }
            outcome.RestorationFailures.Add(location);
            outcome.DiagnosticException = outcome.DiagnosticException == null ? last
                : new AggregateException(outcome.DiagnosticException, last!);
        }
    }

        private static string GetAssumptionDisplayName(
            AssumptionDefinition assumption)
        {
            if (
                !string.IsNullOrWhiteSpace(
                    assumption.Name))
            {
                return
                    assumption.Name;
            }


            return
                $"{assumption.SheetName}!" +
                $"{assumption.CellAddress}";
        }


        // =========================================================
        // DISTRIBUTION SAMPLING
        // =========================================================

        private static double GenerateSample(
            AssumptionDefinition assumption)
        {
            MonteCarlo.Core.DistributionKind distribution =
                assumption.Distribution switch
                {
                    DistributionType.Normal =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Normal,

                    DistributionType.Triangular =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Triangular,

                    DistributionType.Pert =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Pert,

                    DistributionType.Uniform =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Uniform,

                    DistributionType.Lognormal =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Lognormal,

                    DistributionType.Beta =>
                        MonteCarlo.Core
                            .DistributionKind
                            .Beta,

                    _ =>
                        throw new InvalidOperationException(
                            $"Unsupported distribution: " +
                            $"{assumption.Distribution}")
                };


            return
                MonteCarlo.Core
                    .DistributionSampler
                    .Sample(
                        distribution,
                        assumption.Parameter1,
                        assumption.Parameter2,
                        assumption.Parameter3,
                        assumption.Parameter4);
        }
}
