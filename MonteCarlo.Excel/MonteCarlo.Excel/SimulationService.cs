using ExcelDna.Integration;

namespace MonteCarlo.Excel;

public static class SimulationService
{
    // Retain the original API for callers; the Ribbon uses the structured outcome.
    public static SimulationRunResult Run(int trials)
    {
        var outcome = TryRun(trials);
        return outcome.Succeeded ? outcome.Result! : throw new SimulationRunException(outcome);
    }

    public static SimulationExecutionResult TryRun() => TryRunCore(null);
    public static SimulationExecutionResult TryRun(int trials) => TryRunCore(new SimulationSettings(trials));
    private static SimulationExecutionResult TryRunCore(SimulationSettings? requested)
    {
        try
        {
            var validation = WorkbookPersistence.LoadModelForSimulation();
            var settings = requested ?? WorkbookPersistence.LoadSimulationSettings(validation);
            if (!validation.IsValid)
            {
                var invalid = new SimulationExecutionResult();
                invalid.Validation.Errors.AddRange(validation.Errors);
                return invalid;
            }
            dynamic app = ExcelDnaUtil.Application;
            dynamic workbook = app.ActiveWorkbook;
            if (workbook == null)
            {
                var invalid = new SimulationExecutionResult();
                invalid.Validation.Add("Workbook", "No workbook is open.", "Open the workbook containing your simulation model.");
                return invalid;
            }
            return SimulationExecution.Run(settings, SimulationModel.Assumptions.ToArray(),
                SimulationModel.Forecasts.ToArray(), new ExcelSimulationWorkbook(app, workbook), validation,
                correlations: WorkbookPersistence.LoadCorrelations((object)workbook));
        }
        catch (Exception ex) { return new SimulationExecutionResult { DiagnosticException = ex }; }
    }
}
