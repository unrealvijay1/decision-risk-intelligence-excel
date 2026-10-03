using MonteCarlo.Core;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

public sealed class OptimizationTests
{
    public static ForecastDefinition Forecast() => new() { Name = "Y", SheetName = "Model", CellAddress = "B1", CellLink = "forecast" };
    public static AssumptionDefinition Assumption() => new() { Name = "Noise", SheetName = "Model", CellAddress = "A1", Distribution = DistributionType.Uniform, Parameter1 = 0, Parameter2 = 1 };
    public static OptimizationProblem Problem(bool maximize = true) => new(new("forecast", Maximize: maximize),
        [new("x", "X", "Model", "C1", 0, 100, 1)], [], new(200, 10, 456));
    public sealed class Cell(string id, double initial) : ISimulationCell
    {
        public string Identity => id;
        public object? Value = initial;
        public string? InputProblem { get; set; }
        public bool Formula;
        public int Writes;
        public int? FailWrite;
        public bool FailRestore;
        public object? ReadValue() => Value;
        public SimulationCellContent Capture() => new(Value, Formula);
        public void WriteSample(double value) { Writes++; Value = value; Formula = false; if (Writes == FailWrite) throw new InvalidOperationException("write failure"); }
        public void Restore(SimulationCellContent content) { if (FailRestore) throw new InvalidOperationException("restore failure"); Value = content.Value; Formula = content.IsFormula; }
    }
    public sealed class Book : ISimulationWorkbook, IScenarioSandbox
    {
        public Dictionary<string, Cell> Cells = new() { ["A1"] = new("A1", .5), ["B1"] = new("B1", 100), ["C1"] = new("C1", 50), ["C2"] = new("C2", 2) };
        public Func<Book, double> Function = b => 2 * Convert.ToDouble(b.Cells["C1"].Value);
        public int? FailCalculation;
        public int Calculations;
        public bool Disposed;
        public Action? OnCalculate;
        public ISimulationWorkbook Workbook => this;
        public ISimulationCell Resolve(string sheet, string address) => sheet == "Model" && Cells.TryGetValue(address.Replace("$", ""), out var c) ? c : throw new SimulationCellAccessException("Missing cell " + address);
        public bool ScreenUpdating { get; set; } = true;
        public bool EnableEvents { get; set; } = true;
        public object? StatusBar { get; set; } = false;
        public void Calculate() { Calculations++; if (Calculations == FailCalculation) throw new InvalidOperationException("calculation failure"); Cells["B1"].Value = Function(this); OnCalculate?.Invoke(); }
        public void Dispose() => Disposed = true;
    }
    private static OptimizationResult Run(OptimizationProblem p, Book b, CancellationToken token = default, Action<OptimizationProgress>? progress = null) =>
        OptimizationService.Run(p, [Assumption()], [Forecast()], () => b, token, progress);

    [Theory] [InlineData(true, 100)] [InlineData(false, 0)]
    public void KnownLinearOptimumAndRestoration(bool maximize, double expected)
    {
        var b = new Book(); b.Cells["C1"].Formula = true; var r = Run(Problem(maximize), b);
        Assert.Equal(expected, r.Best.Values[0]); Assert.Equal(expected * 2, r.Best.Objective); Assert.Equal(101, r.Evaluations);
        Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Cells["C1"].Formula); Assert.Equal(.5, b.Cells["A1"].Value); Assert.True(b.Disposed);
    }
    [Fact] public void MultiVariableKnownOptimum()
    {
        var b = new Book { Function = b => 2 * Convert.ToDouble(b.Cells["C1"].Value) - Convert.ToDouble(b.Cells["C2"].Value) };
        var p = Problem() with { Variables = [new("x", "X", "Model", "C1", 0, 10, 2), new("z", "Z", "Model", "C2", 0, 8, 2)] };
        var r = Run(p, b); Assert.Equal(new double[] { 10, 0 }, r.Best.Values); Assert.Equal(20, r.Best.Objective); Assert.Equal(2d, b.Cells["C2"].Value);
    }
    [Theory] [InlineData(OptimizationMetric.Mean, 2)] [InlineData(OptimizationMetric.Median, 2)]
    [InlineData(OptimizationMetric.Probability, 66.66666666666667)]
    public void ObjectiveStatisticsInclusive(OptimizationMetric metric, double expected)
    {
        var f = new ForecastRunResult(Forecast(), [1, 2, 3], new());
        Assert.Equal(expected, OptimizationService.Statistic(f, metric, 2, TargetDirection.AtOrAbove), 10);
        if (metric == OptimizationMetric.Probability) Assert.Equal(expected, OptimizationService.Statistic(f, metric, 2, TargetDirection.AtOrBelow), 10);
    }
    [Theory] [InlineData(OptimizationMetric.Mean, true)] [InlineData(OptimizationMetric.Mean, false)]
    [InlineData(OptimizationMetric.Median, true)] [InlineData(OptimizationMetric.Median, false)]
    [InlineData(OptimizationMetric.Probability, true)] [InlineData(OptimizationMetric.Probability, false)]
    public void AllObjectivesRespectDirection(OptimizationMetric metric, bool maximize)
    {
        var p = Problem(maximize) with { Objective = new("forecast", metric, maximize, 100) }; var r = Run(p, new());
        Assert.Equal(metric == OptimizationMetric.Probability ? maximize ? 100 : 0 : maximize ? 200 : 0, r.Best.Objective);
    }
    [Fact] public void BoundsAndStepsIncludingUnalignedMaximum()
    {
        var p = Problem() with { Variables = [new("x", "X", "Model", "C1", -5, 10, 4)] };
        var seen = new List<double>(); var r = Run(p, new(), progress: s => seen.Add(s.Best.Values[0]));
        Assert.Equal(7, r.Best.Values[0]); Assert.All(seen, x => { Assert.InRange(x, -5, 10); Assert.Equal(0, (x + 5) % 4); });
    }
    [Theory] [InlineData(OptimizationConstraintKind.DecisionVariable, 40)] [InlineData(OptimizationConstraintKind.Mean, 80)] [InlineData(OptimizationConstraintKind.Median, 80)]
    public void ConstraintsSelectKnownFeasibleOptimum(OptimizationConstraintKind kind, double requirement)
    {
        var p = Problem() with { Constraints = [new(kind, kind == OptimizationConstraintKind.DecisionVariable ? "x" : "forecast", TargetDirection.AtOrBelow, requirement)] };
        var r = Run(p, new()); Assert.Equal(40, r.Best.Values[0]); Assert.True(r.Best.Feasible); Assert.True(r.Best.Constraints[0].Met);
    }
    [Fact] public void ProbabilityConstraintInclusivePercent()
    {
        var p = Problem(false) with { Constraints = [new(OptimizationConstraintKind.Probability, "forecast", TargetDirection.AtOrAbove, 100, 60)] };
        var r = Run(p, new()); Assert.Equal(30, r.Best.Values[0]); Assert.Equal(100, r.Best.Constraints[0].Actual);
    }
    [Fact] public void NoFeasibleCandidateCannotApply()
    {
        var p = Problem() with { Constraints = [new(OptimizationConstraintKind.Mean, "forecast", TargetDirection.AtOrAbove, 1000)] };
        var b = new Book(); var r = Run(p, b); Assert.False(r.Best.Feasible); Assert.Equal(100, r.Best.Values[0]);
        Assert.Throws<InvalidOperationException>(() => OptimizationService.Apply(r, b, [Assumption()], [Forecast()])); Assert.Equal(50d, b.Cells["C1"].Value);
    }
    [Fact] public void CommonRandomNumbersAndSeedReplay()
    {
        var streams = new List<double[]>(); var b = new Book { Function = b => Convert.ToDouble(b.Cells["C1"].Value) + Convert.ToDouble(b.Cells["A1"].Value) };
        var p = Problem() with { Settings = new(12, 30, 789) };
        var recorder = new RecordingAlgorithm(candidate => streams.Add(candidate.Run.ForecastResults[0].Values.Select(x => x - candidate.Values[0]).ToArray()));
        var r = OptimizationService.Run(p, [Assumption()], [Forecast()], () => b, algorithm: recorder);
        Assert.Equal(r.Evaluations, streams.Count); Assert.Equal(r.Evaluations, r.History.Count);
        Assert.True(streams[0].Distinct().Count() > 1, "Assumptions must vary between trials, not stay at their workbook value.");
        Assert.All(streams.SelectMany(s => s), value => Assert.InRange(value, 0, 1));
        Assert.All(streams, stream => { for (int i = 0; i < stream.Length; i++) Assert.Equal(streams[0][i], stream[i], 11); });
        var other = Run(p, new() { Function = b.Function }); Assert.Equal(r.Best.Values, other.Best.Values); Assert.Equal(r.Best.Objective, other.Best.Objective);
        var different = Run(p with { Settings = p.Settings with { Seed = 790 } }, new() { Function = b.Function }); Assert.NotEqual(r.Best.Objective, different.Best.Objective);
    }
    private sealed class RecordingAlgorithm(Action<OptimizationCandidate> record) : IOptimizationAlgorithm
    {
        public OptimizationCandidate Search(OptimizationProblem p, double[] initial, Func<double[], OptimizationCandidate> evaluate,
            CancellationToken cancellation, Action<OptimizationProgress>? progress, out int evaluations) =>
            new AutomaticOptimizationAlgorithm().Search(p, initial, values => { var c = evaluate(values); record(c); return c; }, cancellation, progress, out evaluations);
    }
    [Fact] public void NonExhaustiveQuadraticSearchImprovesAndFindsInteriorSolution()
    {
        var b = new Book { Function = b => -Math.Pow(Convert.ToDouble(b.Cells["C1"].Value) - 37, 2) - Math.Pow(Convert.ToDouble(b.Cells["C2"].Value) - 83, 2) };
        var p = Problem() with { Variables = [.. Problem().Variables, new("z", "Z", "Model", "C2", 0, 100, 1)], Settings = new(300, 2, 456) };
        var r = Run(p, b); Assert.Equal(new double[] { 37, 83 }, r.Best.Values); Assert.True(r.Evaluations <= 300);
        Assert.Equal(r.History.Count, r.History.Select(h => string.Join("|", h.Values)).Distinct().Count());
    }
    [Fact] public void CancellationDuringTrialRestoresAndDisposes()
    {
        using var cts = new CancellationTokenSource(); var b = new Book(); b.OnCalculate = () => { if (b.Calculations == 4) cts.Cancel(); };
        Assert.Throws<OperationCanceledException>(() => Run(Problem(), b, cts.Token)); Assert.Equal(50d, b.Cells["C1"].Value); Assert.Equal(.5, b.Cells["A1"].Value); Assert.True(b.Disposed);
    }
    [Fact] public void CancellationBeforeRunDoesNotCreateSandbox()
    {
        using var cts = new CancellationTokenSource(); cts.Cancel(); bool opened = false;
        Assert.Throws<OperationCanceledException>(() => OptimizationService.Run(Problem(), [Assumption()], [Forecast()], () => { opened = true; return new Book(); }, cts.Token)); Assert.False(opened);
    }
    [Theory] [InlineData(1)] [InlineData(3)]
    public void CalculationFailureRestores(int calculation)
    {
        var b = new Book { FailCalculation = calculation }; Assert.ThrowsAny<Exception>(() => Run(Problem(), b)); Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Disposed);
    }
    [Fact] public void CandidateWriteFailureRestores()
    { var b = new Book(); b.Cells["C1"].FailWrite = 1; Assert.Throws<InvalidOperationException>(() => Run(Problem(), b)); Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Disposed); }
    [Fact] public void ProgressExceptionRestores()
    { var b = new Book(); Assert.Throws<InvalidOperationException>(() => Run(Problem(), b, progress: _ => throw new InvalidOperationException())); Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Disposed); }
    [Fact] public void RestoreFailureIsVisibleAndDisposes()
    { var b = new Book(); b.Cells["C1"].FailRestore = true; Assert.Throws<AggregateException>(() => Run(Problem(), b)); Assert.True(b.Disposed); }
    [Fact] public void CloseWithoutApplyingKeepsOriginalAndApplyIsExplicit()
    {
        var b = new Book(); var r = Run(Problem(), b); Assert.Equal(50d, b.Cells["C1"].Value);
        OptimizationService.Apply(r, b, [Assumption()], [Forecast()]); Assert.Equal(100d, b.Cells["C1"].Value); Assert.True(b.EnableEvents);
    }
    [Fact] public void ApplyFailureRollsBackAllVariablesAndEvents()
    {
        var b = new Book(); var p = Problem() with { Variables = [.. Problem().Variables, new("z", "Z", "Model", "C2", 0, 4, 1)] };
        var r = Run(p, b); b.Cells["C2"].FailWrite = b.Cells["C2"].Writes + 1;
        Assert.Throws<InvalidOperationException>(() => OptimizationService.Apply(r, b, [Assumption()], [Forecast()])); Assert.Equal(50d, b.Cells["C1"].Value); Assert.Equal(2d, b.Cells["C2"].Value); Assert.True(b.EnableEvents);
    }
    [Fact] public void ApplyRejectsChangedInputBeforeWriting()
    { var b = new Book(); var r = Run(Problem(), b); b.Cells["C1"].Value = 99d; int writes = b.Cells["C1"].Writes; Assert.Throws<InvalidOperationException>(() => OptimizationService.Apply(r, b, [Assumption()], [Forecast()])); Assert.Equal(writes, b.Cells["C1"].Writes); }
    [Theory] [InlineData("merged")] [InlineData("array")] [InlineData("spill")] [InlineData("protected")]
    public void UnsafeCellsRejectedBeforeMutation(string problem)
    { var b = new Book(); b.Cells["C1"].InputProblem = problem; Assert.Throws<ArgumentException>(() => Run(Problem(), b)); Assert.Equal(0, b.Cells["C1"].Writes); Assert.True(b.Disposed); }
    [Theory] [InlineData(null)] [InlineData("text")] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)]
    public void NonnumericRejected(object? value)
    { var b = new Book(); b.Cells["C1"].Value = value; Assert.Throws<ArgumentException>(() => Run(Problem(), b)); Assert.Equal(0, b.Cells["C1"].Writes); }
    [Theory] [InlineData("A1")] [InlineData("B1")] [InlineData("#REF!")] [InlineData("C1:C2")]
    public void InvalidOrReservedReferenceRejected(string cell)
    { var p = Problem() with { Variables = [Problem().Variables[0] with { Cell = cell }] }; Assert.ThrowsAny<Exception>(() => Run(p, new())); }
    [Fact] public void DuplicatePhysicalCellsRejected()
    { var p = Problem() with { Variables = [.. Problem().Variables, new("z", "Z", "Model", "$C$1", 0, 100, 1)] }; Assert.Throws<ArgumentException>(() => Run(p, new())); }
    [Theory] [InlineData(0, 0, 1)] [InlineData(10, 0, 1)] [InlineData(0, 10, 0)] [InlineData(0, 10, -1)] [InlineData(0, 10, 11)] [InlineData(double.NaN, 10, 1)] [InlineData(0, double.PositiveInfinity, 1)]
    public void InvalidBoundsStepRejected(double min, double max, double step)
    { var p = Problem() with { Variables = [Problem().Variables[0] with { Minimum = min, Maximum = max, Step = step }] }; Assert.Throws<ArgumentException>(p.Validate); }
    [Fact] public void MissingForecastAndVariablesRejected()
    { Assert.Throws<ArgumentException>(() => (Problem() with { Variables = [] }).Validate()); Assert.Throws<ArgumentException>(() => Run(Problem() with { Objective = new("missing") }, new())); }
    [Theory] [InlineData(-1)] [InlineData(101)] [InlineData(double.NaN)]
    public void InvalidProbabilityRequirementRejected(double value)
    { Assert.Throws<ArgumentException>(() => (Problem() with { Constraints = [new(OptimizationConstraintKind.Probability, "forecast", TargetDirection.AtOrAbove, value)] }).Validate()); }
    [Fact] public void ContradictoryDecisionConstraintsRejected()
    { Assert.Throws<ArgumentException>(() => (Problem() with { Constraints = [new(OptimizationConstraintKind.DecisionVariable, "x", TargetDirection.AtOrAbove, 50), new(OptimizationConstraintKind.DecisionVariable, "x", TargetDirection.AtOrBelow, 40)] }).Validate()); }
    [Fact] public void FeasibilityAndViolationPrecedeObjective()
    {
        var r = Run(Problem(), new()); var good = r.Best; var bad = good with { Objective = 9999, Constraints = [new(new(OptimizationConstraintKind.Mean, "forecast", TargetDirection.AtOrAbove, 1), 0, false, 1)] };
        Assert.True(AutomaticOptimizationAlgorithm.Better(good, bad, true)); Assert.False(AutomaticOptimizationAlgorithm.Better(bad, good, true)); Assert.True(AutomaticOptimizationAlgorithm.Better(bad with { Constraints = [bad.Constraints[0] with { Violation = .5 }] }, bad, false));
    }
    [Fact] public void ExportContainsValuesStatisticsSettingsAndText()
    {
        var r = Run(Problem() with { Variables = [Problem().Variables[0] with { Name = "=unsafe label" }] }, new()); var table = OptimizerExporter.Table(r);
        Assert.Equal("MONTE CARLO OPTIMIZER", table[0, 0]); Assert.Contains("=unsafe label", table.Cast<object?>()); Assert.Contains("Common seed", table.Cast<object?>()); Assert.Contains("Standard deviation", table.Cast<object?>()); Assert.Contains(100d, table.Cast<object?>());
    }
    [Theory] [InlineData(0, 10)] [InlineData(100001, 10)] [InlineData(10, 0)] [InlineData(10, 1000001)]
    public void InvalidEvaluationSettingsDoNotOpenSandbox(int evaluations, int trials)
    {
        bool opened = false; var p = Problem() with { Settings = new(evaluations, trials, 0) };
        Assert.Throws<ArgumentException>(() => OptimizationService.Run(p, [Assumption()], [Forecast()], () => { opened = true; return new Book(); })); Assert.False(opened);
    }
    [Theory] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)]
    public void InvalidObjectiveTargetRejected(double target)
    { Assert.Throws<ArgumentException>(() => (Problem() with { Objective = new("forecast", OptimizationMetric.Probability, Target: target) }).Validate()); }
    [Fact] public void UnknownObjectiveAndConstraintEnumsRejected()
    {
        Assert.Throws<ArgumentException>(() => (Problem() with { Objective = new("forecast", (OptimizationMetric)99) }).Validate());
        Assert.Throws<ArgumentException>(() => (Problem() with { Constraints = [new((OptimizationConstraintKind)99, "forecast", TargetDirection.AtOrAbove, 1)] }).Validate());
    }
    [Fact] public void MissingForecastConstraintRejectedBeforeSandbox()
    {
        bool opened = false; var p = Problem() with { Constraints = [new(OptimizationConstraintKind.Mean, "deleted", TargetDirection.AtOrAbove, 1)] };
        Assert.Throws<ArgumentException>(() => OptimizationService.Run(p, [Assumption()], [Forecast()], () => { opened = true; return new Book(); })); Assert.False(opened);
    }
    [Fact] public void SimulationConfigurationErrorRejectedBeforeSandbox()
    {
        bool opened = false; var input = Assumption(); input.Parameter2 = -1;
        Assert.Throws<ArgumentException>(() => OptimizationService.Run(Problem(), [input], [Forecast()], () => { opened = true; return new Book(); })); Assert.False(opened);
    }
    [Fact] public void CancellationOnLastEvaluationIsHonored()
    {
        var b = new Book(); using var cts = new CancellationTokenSource(); var p = Problem() with { Settings = new(1, 2, 0) };
        Assert.Throws<OperationCanceledException>(() => Run(p, b, cts.Token, _ => cts.Cancel())); Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Disposed);
    }
    [Fact] public void ApplyRejectsChangedContentEvenWhenNumericValueUnchanged()
    {
        var b = new Book(); b.Cells["C1"].Formula = true; var r = Run(Problem(), b); b.Cells["C1"].Formula = false;
        Assert.Throws<InvalidOperationException>(() => OptimizationService.Apply(r, b, [Assumption()], [Forecast()])); Assert.Equal(50d, b.Cells["C1"].Value);
    }
}
