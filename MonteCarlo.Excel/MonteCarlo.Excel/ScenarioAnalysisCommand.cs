using ExcelDna.Integration;
using MonteCarlo.Core;

namespace MonteCarlo.Excel;

internal static class ScenarioAnalysisCommand
{
    public static void Show()
    {
        dynamic app = ExcelDnaUtil.Application;
        object source = app.ActiveWorkbook ?? throw new InvalidOperationException("Open your model workbook first.");
        var originalInputs = SimulationModel.Assumptions.ToArray();
        var originalOutputs = SimulationModel.Forecasts.ToArray();
        AssumptionDefinition[] inputs;
        ForecastDefinition[] outputs;
        SimulationSettings settings;
        try
        {
            var validation = WorkbookPersistence.LoadModelForSimulation(restoreHighlights: false);
            settings = WorkbookPersistence.LoadSimulationSettings(validation);
            if (!validation.IsValid) throw new ArgumentException(validation.UserMessage);
            inputs = SimulationModel.Assumptions.Select(ScenarioSimulationService.Clone).ToArray();
            outputs = SimulationModel.Forecasts.Select(ScenarioSimulationService.Clone).ToArray();
        }
        finally
        {
            SimulationModel.Clear(); SimulationModel.Assumptions.AddRange(originalInputs); SimulationModel.Forecasts.AddRange(originalOutputs);
        }
        var correlations = WorkbookPersistence.LoadCorrelations(source);
        var analysis = WorkbookPersistence.LoadScenarioAnalysis(source);
        var options = new ScenarioCalculationOptions((bool)app.Iteration, (int)app.MaxIterations, (double)app.MaxChange);
        using var form = new ScenarioAnalysisForm(inputs, outputs, analysis,
            edited => WorkbookPersistence.SaveScenarioAnalysis(source, edited, inputs, outputs),
            (selected, token, progress) => RunAsync(source, options, settings, inputs, outputs, selected, correlations, token, progress));
        form.ShowDialog();
    }

    private static Task<ScenarioComparisonResult> RunAsync(object source, ScenarioCalculationOptions options, SimulationSettings settings,
        AssumptionDefinition[] inputs, ForecastDefinition[] outputs, ScenarioAnalysisDefinition selected,
        IReadOnlyList<AssumptionCorrelation> correlations, CancellationToken token, IProgress<string> progress)
    {
        token.ThrowIfCancellationRequested();
        progress.Report("Creating a disposable workbook copy…");
        var copy = ScenarioExcelSandbox.Prepare(source); // Source COM stays on Excel's UI thread.
        var completion = new TaskCompletionSource<ScenarioComparisonResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var worker = new Thread(() =>
        {
            try
            {
                ScenarioComparisonResult result;
                using (copy)
                    result = ScenarioSimulationService.Run(settings, inputs, outputs, selected,
                        () => ScenarioExcelSandbox.Open(copy.Path, options), token, progress.Report, correlations);
                completion.SetResult(result); // Report success only after workbook/process/file cleanup.
            }
            catch (OperationCanceledException) { completion.SetCanceled(token); }
            catch (Exception ex) { completion.SetException(ex); }
        }) { IsBackground = true, Name = "Monte Carlo Scenario Analysis" };
        try { worker.SetApartmentState(ApartmentState.STA); worker.Start(); }
        catch { copy.Dispose(); throw; }
        return completion.Task;
    }
}
