using ExcelDna.Integration;
using MonteCarlo.Core;
using MonteCarlo.Excel.Licensing;
namespace MonteCarlo.Excel;

internal static class OptimizerCommand
{
    public static void Show()
    {
        LicenseService.EnsureAccess(); dynamic app = ExcelDnaUtil.Application;
        object source = app.ActiveWorkbook ?? throw new InvalidOperationException("Open your model workbook first.");
        var oldInputs = SimulationModel.Assumptions.ToArray(); var oldOutputs = SimulationModel.Forecasts.ToArray();
        AssumptionDefinition[] inputs; ForecastDefinition[] outputs;
        try
        {
            var validation = WorkbookPersistence.LoadModelForSimulation(restoreHighlights: false);
            if (!validation.IsValid) throw new ArgumentException(validation.UserMessage);
            inputs = SimulationModel.Assumptions.Select(ScenarioSimulationService.Clone).ToArray(); outputs = SimulationModel.Forecasts.Select(ScenarioSimulationService.Clone).ToArray();
        }
        finally { SimulationModel.Clear(); SimulationModel.Assumptions.AddRange(oldInputs); SimulationModel.Forecasts.AddRange(oldOutputs); }
        var correlations = WorkbookPersistence.LoadCorrelations(source);
        var sourceAdapter = new ExcelSimulationWorkbook(app, source);
        DecisionVariableSelection Read(DecisionVariable v)
        {
            var resolved = WorkbookPersistence.ResolveOptimizer(source, new(new("selection"), [v], [], new())).Variables[0];
            return DecisionVariableSelection.Read(sourceAdapter, resolved.Sheet, resolved.Cell);
        }
        DecisionVariableSelection? Pick()
        {
            object selected;
            try { selected = app.InputBox(Prompt: "Select exactly one numeric decision-variable cell in this workbook.", Title: "Decision Variable", Type: 8); }
            catch (System.Runtime.InteropServices.COMException) { return null; }
            return DecisionVariableSelection.ReadRange(source, (object)app, selected, sourceAdapter);
        }
        void ValidateVariables(IReadOnlyList<DecisionVariable> variables)
        {
            // Validate the draft using existing range/step and cell rules without requiring objective/constraints.
            var draft = new OptimizationProblem(new("selection"), variables, [], new());
            draft.Validate(); OptimizationService.ValidateCells(draft, sourceAdapter, inputs, outputs);
        }
        DecisionVariableSelection ReadConstraint(OptimizationConstraint c)
        {
            var resolved = WorkbookPersistence.ResolveOptimizer(source, new(new("selection"), [], [c], new())).Constraints[0];
            return ConstraintCellSelection.Read(sourceAdapter, resolved.Sheet, resolved.Cell);
        }
        DecisionVariableSelection? PickConstraint()
        {
            object selected;
            try { selected = app.InputBox(Prompt: "Select one numeric value or formula cell in this workbook.", Title: "Optimizer Constraint", Type: 8); }
            catch (System.Runtime.InteropServices.COMException) { return null; }
            return ConstraintCellSelection.ReadRange(source, (object)app, selected, sourceAdapter);
        }
        OptimizationProblem Save(OptimizationProblem p)
        {
            if ((bool)((dynamic)source).ReadOnly) throw new InvalidOperationException("The source workbook is read-only. Open an editable copy to save Optimizer settings.");
            p.Validate(inputs.Length == 0); OptimizationService.ValidateMode(p, inputs.Length == 0); OptimizationService.ValidateCells(p, sourceAdapter, inputs, outputs);
            OptimizationService.ValidateConstraintCells(p.Constraints, sourceAdapter);
            return WorkbookPersistence.SaveOptimizer(source, p, outputs);
        }
        using var form = new OptimizerForm(outputs, WorkbookPersistence.LoadOptimizer(source), Save,
            (p, token, progress) => RunLive(source, p, inputs, outputs, correlations, token, progress),
            r =>
            {
                if ((bool)((dynamic)source).ReadOnly) throw new InvalidOperationException("The source workbook is read-only. Open an editable copy before applying values.");
                var resolved = WorkbookPersistence.ResolveOptimizer(source, r.Problem);
                OptimizationService.Apply(r with { Problem = resolved }, sourceAdapter, inputs, outputs);
            }, r => OptimizerExporter.Export(source, r), Pick, Read, ValidateVariables, PickConstraint, ReadConstraint,
            c => OptimizationService.ValidateConstraintCells([c], sourceAdapter), deterministic: inputs.Length == 0);
        form.ShowDialog();
    }
    private static Task<OptimizationResult> RunLive(object source, OptimizationProblem problem,
        AssumptionDefinition[] inputs, ForecastDefinition[] outputs, IReadOnlyList<AssumptionCorrelation> correlations,
        CancellationToken token, IProgress<OptimizationProgress> progress)
    {
        dynamic app = ExcelDnaUtil.Application;
        bool interactive = app.Interactive, events = app.EnableEvents;
        using var liveCancellation = CancellationFor(token);
        token = liveCancellation.Token;
        using var live = new LiveRunForm("Optimizer — live model", liveCancellation, true);
        live.Show();
        try
        {
            if ((bool)((dynamic)source).ReadOnly) throw new InvalidOperationException("Open an editable workbook to run live optimization.");
            app.Interactive = false; app.EnableEvents = false;
            token.ThrowIfCancellationRequested();
            var result = OptimizationService.Run(problem, inputs, outputs,
                () => new SourceWorkbook(new ExcelSimulationWorkbook(app, source)), token,
                p => { progress.Report(p); live.Evaluation(p, problem, outputs); }, correlations,
                liveProgress: live.Trial, candidateStarted: x => live.CandidateStarted(x, problem, inputs.Length == 0));
            return Task.FromResult(result);
        }
        finally { live.Finish(); try { app.EnableEvents = events; } finally { app.Interactive = interactive; } }
    }
    private static CancellationTokenSource CancellationFor(CancellationToken token) => CancellationTokenSource.CreateLinkedTokenSource(token);
    private sealed class SourceWorkbook(ISimulationWorkbook workbook) : IScenarioSandbox
    {
        public ISimulationWorkbook Workbook => workbook;
        public void Dispose() { }
    }
}
