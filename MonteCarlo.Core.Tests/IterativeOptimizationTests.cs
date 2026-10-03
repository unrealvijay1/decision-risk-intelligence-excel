using MonteCarlo.Excel;
using MonteCarlo.Core;
namespace MonteCarlo.Core.Tests;

public sealed class IterativeOptimizationTests
{
    public static OptimizationConstraint Linear(double rhs, OptimizationConstraintOperator op, params (string Id, double Weight)[] terms) =>
        new(OptimizationConstraintKind.Linear, "linear", TargetDirection.AtOrAbove, rhs, Operator: op,
            Terms: terms.Select(x => new OptimizationLinearTerm(x.Id, x.Weight)).ToArray());
    public static (OptimizationTests.Book Book, OptimizationProblem Problem, AssumptionDefinition[] Inputs, ForecastDefinition Forecast) Portfolio(int budget = 250, bool cap = false)
    {
        var b = new OptimizationTests.Book();
        b.Cells["C1"].Value = .3; b.Cells["C2"].Value = .4; b.Cells["C3"] = new("C3", .3);
        b.Cells["A1"].Value = .08; b.Cells["A2"] = new("A2", .12); b.Cells["A3"] = new("A3", .07);
        b.Cells["D1"] = new("D1", 1);
        b.Function = book =>
        {
            double g = Convert.ToDouble(book.Cells["C1"].Value), s = Convert.ToDouble(book.Cells["C2"].Value), f = Convert.ToDouble(book.Cells["C3"].Value);
            book.Cells["D1"].Value = g + s + f;
            return 10000 * (g * Convert.ToDouble(book.Cells["A1"].Value) + s * Convert.ToDouble(book.Cells["A2"].Value) + f * Convert.ToDouble(book.Cells["A3"].Value));
        };
        var vars = new DecisionVariable[] { new("g", "Gold Allocation", "Model", "C1", 0, 1, .01), new("s", "Shares Allocation", "Model", "C2", 0, 1, .01), new("f", "FD Allocation", "Model", "C3", 0, 1, .01) };
        var constraints = new List<OptimizationConstraint> { Linear(1, OptimizationConstraintOperator.Equal, ("g", 1), ("s", 1), ("f", 1)), CellConstraintTests.Cell("D1", 1, OptimizationConstraintOperator.Equal, "Total Allocation") };
        if (cap) constraints.Add(new(OptimizationConstraintKind.DecisionVariable, "s", TargetDirection.AtOrBelow, .6));
        var inputs = new[] { .08, .12, .07 }.Select((mean, i) => new AssumptionDefinition { Name = "Return " + i, SheetName = "Model", CellAddress = "A" + (i + 1), Distribution = DistributionType.Normal, Parameter1 = mean, Parameter2 = .02 }).ToArray();
        return (b, new(new("forecast"), vars, constraints, new(budget, 1000, 42)), inputs, OptimizationTests.Forecast());
    }
    [Theory] [InlineData(false, 0, 1, 0, 1200)] [InlineData(true, .4, .6, 0, 1040)]
    public void PortfolioAcceptance(bool cap, double g, double s, double f, double mean)
    {
        var (b, p, inputs, output) = Portfolio(cap: cap); var r = OptimizationService.Run(p, inputs, [output], () => b);
        Assert.Equal(new[] { g, s, f }, r.Best.Values); Assert.InRange(r.Best.Objective, mean - 25, mean + 25); Assert.NotNull(r.BestFeasible);
        Assert.Equal(250, r.Evaluations); Assert.Equal(250, r.MonteCarloEvaluations); Assert.Equal(0, r.RejectedBeforeSimulation);
        Assert.Equal(250000, b.Cells["A1"].Writes); Assert.True(r.Best.Run.ForecastResults[0].StandardDeviation > 100);
        Assert.All(r.History, h => { Assert.True(h.Feasible); Assert.InRange(Math.Abs(h.Values.Sum() - 1), 0, 1e-9); Assert.All(h.Values, x => Assert.Equal(Math.Round(x / .01), x / .01, 8)); });
        Assert.Equal(.3, b.Cells["C1"].Value); Assert.Equal(.4, b.Cells["C2"].Value); Assert.Equal(.3, b.Cells["C3"].Value);
        Assert.Equal(.08, b.Cells["A1"].Value); Assert.Equal(.12, b.Cells["A2"].Value); Assert.Equal(.07, b.Cells["A3"].Value);
        var other = Portfolio(cap: cap); var replay = OptimizationService.Run(other.Problem, other.Inputs, [other.Forecast], () => other.Book);
        Assert.Equal(r.Best.Objective, replay.Best.Objective); Assert.Equal(r.History.Select(h => string.Join("|", h.Values)), replay.History.Select(h => string.Join("|", h.Values)));
    }
    [Fact] public void DeterministicConstraintRejectsBeforeMonteCarloAndReportsNoOptimum()
    {
        var b = new OptimizationTests.Book(); var p = OptimizationTests.Problem() with { Constraints = [CellConstraintTests.Cell("C2", 1)] };
        var r = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => b);
        Assert.Equal(r.Evaluations, r.RejectedBeforeSimulation); Assert.Equal(0, r.MonteCarloEvaluations); Assert.Equal(0, b.Cells["A1"].Writes);
        Assert.Null(r.BestFeasible); Assert.All(r.History, h => { Assert.False(h.Evaluated); Assert.True(double.IsNaN(h.Objective)); Assert.Null(h.BestFeasibleObjective); });
        Assert.Contains("No feasible solution found", OptimizerExporter.Table(r).Cast<object?>());
        Assert.Throws<InvalidOperationException>(() => OptimizationService.Apply(r, b, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()]));
        Assert.Equal(50d, b.Cells["C1"].Value);
    }
    [Fact] public void SimulationDependentConstraintsStillRequireDistribution()
    {
        var b = new OptimizationTests.Book(); var p = OptimizationTests.Problem() with { Constraints = [new(OptimizationConstraintKind.Mean, "forecast", TargetDirection.AtOrAbove, 1000)] };
        var r = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => b);
        Assert.Null(r.BestFeasible); Assert.Equal(0, r.RejectedBeforeSimulation); Assert.Equal(r.Evaluations, r.MonteCarloEvaluations);
        Assert.All(r.History, h => Assert.True(h.Evaluated));
    }
    [Fact] public void IterativeInteriorSearchUsesEliteRefinementAndGlobalExploration()
    {
        var b = new OptimizationTests.Book { Function = x => -Math.Pow(Convert.ToDouble(x.Cells["C1"].Value) - 37, 2) - Math.Pow(Convert.ToDouble(x.Cells["C2"].Value) - 83, 2) };
        var p = OptimizationTests.Problem() with { Variables = [.. OptimizationTests.Problem().Variables, new("z", "Z", "Model", "C2", 0, 100, 1)], Settings = new(300, 1000, 456) };
        var r = OptimizationService.Run(p, [], [OptimizationTests.Forecast()], () => b);
        Assert.Equal(new double[] { 37, 83 }, r.Best.Values); Assert.Equal(0, r.Best.Objective);
        Assert.Contains(r.History, h => h.Phase == "Global exploration");
        var local = r.History.Where(h => h.Phase == "Elite neighborhood").ToArray();
        Assert.Contains(local, h => h.RadiusSteps == 1); Assert.Contains(local, h => h.RadiusSteps >= 25);
        Assert.True(local.Average(h => h.Values.Zip(r.Best.Values, (a, z) => Math.Abs(a - z)).Sum()) < r.History.Where(h => h.Phase == "Global exploration").Average(h => h.Values.Zip(r.Best.Values, (a, z) => Math.Abs(a - z)).Sum()));
        Assert.True(r.History[0].BestFeasibleObjective < r.History[^1].BestFeasibleObjective);
        Assert.All(r.History.Zip(r.History.Skip(1)), pair => Assert.True(pair.Second.BestFeasibleObjective >= pair.First.BestFeasibleObjective));
    }
    [Theory] [InlineData(1)] [InlineData(42)] [InlineData(99)]
    public void WeightedEqualityAndBoundsStayOnGrid(int seed)
    {
        var p = OptimizationTests.Problem() with { Variables = [new("x", "X", "Model", "C1", 0, 25, 1), new("z", "Z", "Model", "C2", 0, 50, 1)],
            Constraints = [Linear(50, OptimizationConstraintOperator.Equal, ("x", 2), ("z", 1))] };
        var space = new OptimizationFeasibleSpace(p); var random = new Random(seed);
        for (int i = 0; i < 100; i++)
        { var values = space.Repair(space.Random(random), random); Assert.Equal(50, 2 * values[0] + values[1], 8); Assert.Equal(Math.Round(values[0]), values[0]); Assert.Equal(Math.Round(values[1]), values[1]); }
    }
    [Fact] public void EqualityToleranceDoesNotRequireExactFloatingPointSum()
    {
        var c = Linear(.3, OptimizationConstraintOperator.Equal, ("x", 1), ("z", 1));
        Assert.True(OptimizationService.EvaluateConstraint(c, .1 + .2).Met); Assert.False(OptimizationService.EvaluateConstraint(c, .30001).Met);
    }
    [Fact] public void ExpressionSupportsNamesCoefficientsAndNegativeTerms()
    {
        var vars = Portfolio().Problem.Variables;
        var terms = OptimizationFeasibleSpace.Parse("2*Gold Allocation + Shares Allocation - FD Allocation", vars);
        var c = Linear(1, OptimizationConstraintOperator.Equal) with { Terms = terms };
        Assert.Equal(new double[] { 2, 1, -1 }, terms.Select(t => t.Coefficient));
        Assert.Equal(terms, OptimizationFeasibleSpace.Parse(OptimizationFeasibleSpace.Expression(c, vars), vars));
    }
    [Theory] [InlineData("missing + X")] [InlineData("2*X*X")] [InlineData("X/2")] [InlineData("X - X")]
    public void InvalidExpressionsReject(string expression) => Assert.ThrowsAny<Exception>(() => OptimizationFeasibleSpace.Parse(expression, OptimizationTests.Problem().Variables));
    [Fact] public void CancellationInRejectedSearchRestores()
    {
        using var cts = new CancellationTokenSource(); var b = new OptimizationTests.Book(); b.Cells["C1"].Formula = true;
        var p = OptimizationTests.Problem() with { Constraints = [CellConstraintTests.Cell("C2", 1)] };
        Assert.Throws<OperationCanceledException>(() => OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => b, cts.Token, _ => cts.Cancel()));
        Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Cells["C1"].Formula); Assert.True(b.Disposed);
    }
    [Fact] public void ApplyAndClosePreserveFeasiblePortfolio()
    {
        var (b, p, _, f) = Portfolio(cap: true); var r = OptimizationService.Run(p, [], [f], () => b);
        Assert.Equal(.3, b.Cells["C1"].Value); OptimizationService.Apply(r, b, [], [f]);
        Assert.Equal(new[] { .4, .6, 0 }, r.Best.Values); Assert.Equal(1d, b.Cells["D1"].Value); Assert.Equal(1040d, Convert.ToDouble(b.Cells["B1"].Value), 8);
    }
    [Fact] public void CandidateAndTrialBudgetsAreSeparate()
    {
        var b = new OptimizationTests.Book(); var p = OptimizationTests.Problem() with { Settings = new(5, 1000, 42) };
        var r = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => b);
        Assert.Equal(5, r.Evaluations); Assert.Equal(5000, b.Cells["A1"].Writes); Assert.Equal(1000, r.Best.Run.Trials);
    }
    private sealed class FormulaCell(string formula) : ISimulationCell
    {
        public string Identity => "D1"; public string? InputProblem => null;
        public object? ReadValue() => 1d; public SimulationCellContent Capture() => new(formula, true);
        public void WriteSample(double value) => throw new Exception("Constraint write");
        public void Restore(SimulationCellContent content) => throw new Exception("Constraint restore");
    }
    [Theory] [InlineData("=C1+C2+C3")] [InlineData("=SUM(C1:C3)")] [InlineData("=sum($C$1,$C$2,$C$3)")]
    public void SimpleWorksheetEqualityProvidesGeneralFeasibleGuide(string formula)
    {
        var p = Portfolio().Problem; var c = CellConstraintTests.Cell("D1", 1, OptimizationConstraintOperator.Equal);
        var guide = OptimizationFeasibleSpace.Guide(c, new FormulaCell(formula), p.Variables)!;
        Assert.NotNull(guide); Assert.Equal(new double[] { 1, 1, 1 }, guide.Terms!.Select(t => t.Coefficient));
        var space = new OptimizationFeasibleSpace(p with { Constraints = [c] }, [guide]); var rng = new Random(42);
        for (int i = 0; i < 100; i++) Assert.Equal(1, space.Repair(space.Random(rng), rng).Sum(), 9);
    }
    [Theory] [InlineData("=2*SUM(C1:C3)")] [InlineData("=SUM(C1:C3)*2")]
    [InlineData("=C1*C2+C3")] [InlineData("=INDIRECT(\"C1\")")] [InlineData("=A1+C2+C3")]
    public void UnsupportedOrUncertainFormulaDoesNotPretendToBeLinear(string formula) =>
        Assert.Null(OptimizationFeasibleSpace.Guide(CellConstraintTests.Cell("D1", 1, OptimizationConstraintOperator.Equal), new FormulaCell(formula), Portfolio().Problem.Variables));
    [Fact] public void ImpossibleGridEqualityReturnsNoFeasibleSolutionWithoutTrials()
    {
        var b = new OptimizationTests.Book(); var p = OptimizationTests.Problem() with { Constraints = [Linear(.5, OptimizationConstraintOperator.Equal, ("x", 1))] };
        var r = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => b);
        Assert.Null(r.BestFeasible); Assert.Equal(r.Evaluations, r.RejectedBeforeSimulation); Assert.Equal(0, b.Cells["A1"].Writes);
    }
    [Theory] [InlineData(OptimizationConstraintOperator.LessThanOrEqual)] [InlineData(OptimizationConstraintOperator.GreaterThanOrEqual)]
    public void LinearInequalityGuidesCandidateGeneration(OptimizationConstraintOperator op)
    {
        var p = Portfolio().Problem with { Constraints = [Linear(.6, op, ("g", 1), ("s", 1))] };
        var space = new OptimizationFeasibleSpace(p); var rng = new Random(42);
        for (int i = 0; i < 100; i++) { var values = space.Repair(space.Random(rng), rng); Assert.True(OptimizationService.EvaluateConstraint(p.Constraints[0], values[0] + values[1]).Met); }
    }
    [Fact] public void ProgressIdentifiesCurrentAndBestFeasibleAndCounters()
    {
        var (b, p, inputs, f) = Portfolio(5); OptimizationProgress? last = null;
        var r = OptimizationService.Run(p, inputs, [f], () => b, progress: x => last = x);
        Assert.NotNull(last!.Current); Assert.NotNull(last.BestFeasible); Assert.Equal(r.MonteCarloEvaluations, last.MonteCarloEvaluations);
        var text = OptimizationProgressText.Format(last, p, false);
        Assert.Contains("Current Candidate", text); Assert.Contains("Best Feasible Candidate", text); Assert.Contains("1,000", text);
    }
    [Fact] public void MultipleIndependentEqualitiesRepair()
    {
        var p = Portfolio().Problem with { Constraints = [Linear(.2, OptimizationConstraintOperator.Equal, ("g", 1)), Linear(.8, OptimizationConstraintOperator.Equal, ("s", 1), ("f", 1))] };
        var space = new OptimizationFeasibleSpace(p); var rng = new Random(42);
        for (int i = 0; i < 100; i++) { var v = space.Repair(space.Random(rng), rng); Assert.Equal(.2, v[0], 9); Assert.Equal(.8, v[1] + v[2], 9); }
    }
    [Fact] public void FailureDuringConstrainedSimulationRestoresAllInputs()
    {
        var (b, p, inputs, f) = Portfolio(5); b.Cells["C1"].Formula = true;
        Assert.Throws<SimulationRunException>(() => OptimizationService.Run(p, inputs, [f], () => b, liveProgress: (_, _) => throw new InvalidOperationException("preview")));
        Assert.Equal(.3, b.Cells["C1"].Value); Assert.True(b.Cells["C1"].Formula); Assert.Equal(.08, b.Cells["A1"].Value);
    }
    [Fact] public void PortfolioCommonRandomNumbersReuseAllThreeUnderlyingReturns()
    {
        var (b, p, inputs, f) = Portfolio(5); var scenarios = new List<double[]>();
        var r = OptimizationService.Run(p, inputs, [f], () => b, liveProgress: (n, _) =>
        { if (n == 1) scenarios.Add(inputs.Select(a => Convert.ToDouble(b.Cells[a.CellAddress].Value)).ToArray()); });
        Assert.Equal(r.MonteCarloEvaluations, scenarios.Count); Assert.All(scenarios, x => Assert.Equal(scenarios[0], x));
        Assert.NotEqual(.08, scenarios[0][0]); Assert.NotEqual(.12, scenarios[0][1]); Assert.NotEqual(.07, scenarios[0][2]);
    }
    [Fact] public void CandidateStartFailureRestoresBeforeAnySimulation()
    {
        var b = new OptimizationTests.Book(); b.Cells["C1"].Formula = true;
        Assert.Throws<InvalidOperationException>(() => OptimizationService.Run(OptimizationTests.Problem(), [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => b,
            candidateStarted: x => { Assert.Equal(1, x.Number); throw new InvalidOperationException("Starting preview failed"); }));
        Assert.Equal(50d, b.Cells["C1"].Value); Assert.True(b.Cells["C1"].Formula); Assert.Equal(0, b.Cells["A1"].Writes);
    }
    [Theory] [InlineData(1e-10)] [InlineData(-2e20)]
    public void ScientificCoefficientsSurviveEditorExpressionRoundTrip(double coefficient)
    {
        var p = OptimizationTests.Problem(); var c = Linear(1, OptimizationConstraintOperator.Equal, ("x", coefficient));
        var parsed = OptimizationFeasibleSpace.Parse(OptimizationFeasibleSpace.Expression(c, p.Variables), p.Variables);
        Assert.Equal(c.Terms, parsed);
    }
}
