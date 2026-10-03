using MonteCarlo.Core;
using MonteCarlo.Excel;

namespace MonteCarlo.Core.Tests;

public sealed class CellConstraintTests
{
    public static OptimizationConstraint Cell(string address, double rhs, OptimizationConstraintOperator op = OptimizationConstraintOperator.LessThanOrEqual, string name = "Capacity") =>
        new(OptimizationConstraintKind.Cell, "cell:" + address, TargetDirection.AtOrAbove, rhs, Sheet: "Model", Cell: address, Name: name, Operator: op);
    public static (OptimizationTests.Book Book, OptimizationProblem Problem, ForecastDefinition Forecast) ProductMix()
    {
        var book = new OptimizationTests.Book();
        foreach (var (address, value) in new[] { ("B2", 10d), ("B3", 10d), ("B5", 13000d), ("B6", 60d), ("B7", 50d) })
            book.Cells[address] = new(address, value) { Formula = address is "B5" or "B6" or "B7" };
        book.Function = b =>
        {
            double a = Convert.ToDouble(b.Cells["B2"].Value), z = Convert.ToDouble(b.Cells["B3"].Value);
            b.Cells["B5"].Value = a * 500 + z * 800;
            b.Cells["B6"].Value = a * 2 + z * 4; b.Cells["B7"].Value = a * 3 + z * 2;
            return a * 500 + z * 800;
        };
        var forecast = new ForecastDefinition { Name = "Total Profit", SheetName = "Model", CellAddress = "B5", CellLink = "profit" };
        var p = new OptimizationProblem(new("profit"),
            [new("a", "Qty A", "Model", "B2", 0, 120, 10), new("b", "Qty B", "Model", "B3", 0, 120, 10)],
            [Cell("B6", 400, name: "Labor Used"), Cell("B7", 360, name: "Material Used"),
             Cell("B2", 0, OptimizationConstraintOperator.GreaterThanOrEqual, "Nonnegative A"), new(OptimizationConstraintKind.DecisionVariable, "b", TargetDirection.AtOrAbove, 0)], new(200, 2, 42));
        return (book, p, forecast);
    }
    [Fact] public void ProductMixReadsRecalculatedUnregisteredFormulasAndFindsKnownOptimum()
    {
        var (book, p, f) = ProductMix();
        var result = OptimizationService.Run(p, [OptimizationTests.Assumption()], [f], () => book);
        Assert.Equal(new double[] { 80, 60 }, result.Best.Values); Assert.Equal(88000, result.Best.Objective);
        Assert.True(result.Best.Feasible); Assert.All(result.Best.Constraints, c => Assert.True(c.Met));
        Assert.Equal(400, result.Best.Constraints[0].Actual); Assert.Equal(360, result.Best.Constraints[1].Actual);
        Assert.Contains(result.History, h => !h.Feasible); Assert.True(result.Best.Objective > 13000);
        Assert.Single(result.Best.Run.ForecastResults); Assert.Equal("B5", result.Best.Run.ForecastResults[0].Forecast.CellAddress);
        foreach (string address in new[] { "B5", "B6", "B7" }) { Assert.Equal(0, book.Cells[address].Writes); Assert.True(book.Cells[address].Formula); }
        Assert.Equal(10d, book.Cells["B2"].Value); Assert.Equal(10d, book.Cells["B3"].Value);
        Assert.Equal(60d, book.Cells["B6"].Value); Assert.Equal(50d, book.Cells["B7"].Value);
    }
    [Theory] [InlineData(OptimizationConstraintOperator.LessThanOrEqual, 2, true)]
    [InlineData(OptimizationConstraintOperator.GreaterThanOrEqual, 3, false)]
    [InlineData(OptimizationConstraintOperator.Equal, 2, true)]
    [InlineData(OptimizationConstraintOperator.Equal, 3, false)]
    public void ConstantCellAndOperators(OptimizationConstraintOperator op, double rhs, bool expected)
    {
        var book = new OptimizationTests.Book(); book.Cells["C2"].InputProblem = "Protected output";
        var p = OptimizationTests.Problem() with { Constraints = [Cell("C2", rhs, op)] };
        var result = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => book);
        Assert.Equal(expected, result.Best.Feasible); Assert.Equal(2, result.Best.Constraints[0].Actual); Assert.Equal(0, book.Cells["C2"].Writes);
    }
    [Fact] public void EqualityUsesDocumentedToleranceWithoutRelaxingInequalities()
    {
        var c = Cell("C2", 10, OptimizationConstraintOperator.Equal);
        Assert.True(OptimizationService.EvaluateConstraint(c, 10 + 5e-9).Met);
        Assert.False(OptimizationService.EvaluateConstraint(c, 10 + 2e-8).Met);
        Assert.False(OptimizationService.EvaluateConstraint(c with { Operator = OptimizationConstraintOperator.LessThanOrEqual }, 10 + 1e-10).Met);
        Assert.True(OptimizationService.EvaluateConstraint(c with { Operator = OptimizationConstraintOperator.GreaterThanOrEqual }, 10).Met);
    }
    [Fact] public void CellConstraintsCanReadDecisionAndForecastCellsWhileLegacyStatisticsRemainSupported()
    {
        var book = new OptimizationTests.Book { Function = b => 2 * Convert.ToDouble(b.Cells["C1"].Value) + Convert.ToDouble(b.Cells["A1"].Value) };
        var p = OptimizationTests.Problem() with { Constraints = [Cell("C1", 40), Cell("B1", 80.5), new(OptimizationConstraintKind.Mean, "forecast", TargetDirection.AtOrBelow, 81)] };
        var result = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => book);
        Assert.Equal(40, result.Best.Values[0]); Assert.True(result.Best.Feasible);
        Assert.Equal(80.5, result.Best.Constraints[1].Actual); Assert.InRange(result.Best.Constraints[2].Actual, 80, 81);
    }
    [Fact] public void EqualityDecisionConstraintKeepsGridValidation()
    {
        var c = new OptimizationConstraint(OptimizationConstraintKind.DecisionVariable, "x", TargetDirection.AtOrAbove, 30, Operator: OptimizationConstraintOperator.Equal);
        var p = OptimizationTests.Problem() with { Constraints = [c] };
        var result = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => new OptimizationTests.Book());
        Assert.Equal(30, result.Best.Values[0]); Assert.True(result.Best.Feasible);
        Assert.Throws<ArgumentException>(() => (p with { Constraints = [c with { Value = 30.5 }] }).Validate());
    }
    [Theory] [InlineData("C1:C2")] [InlineData("#REF!")] [InlineData("XFE1")] [InlineData("A1048577")] [InlineData("A0")] [InlineData("A1,B1")]
    public void InvalidReferencesRejectBeforeWrites(string address)
    {
        var book = new OptimizationTests.Book(); var p = OptimizationTests.Problem() with { Constraints = [Cell(address, 1)] };
        var ex = Assert.Throws<ArgumentException>(() => OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => book));
        Assert.Contains("Constraint", ex.Message); Assert.All(book.Cells.Values, x => Assert.Equal(0, x.Writes));
    }
    [Fact] public void MissingWorksheetRejectsBeforeWritesWithLocation()
    {
        var book = new OptimizationTests.Book(); var p = OptimizationTests.Problem() with { Constraints = [Cell("C2", 1) with { Sheet = "Missing" }] };
        var ex = Assert.Throws<ArgumentException>(() => OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => book));
        Assert.Contains("Missing!C2", ex.Message); Assert.All(book.Cells.Values, x => Assert.Equal(0, x.Writes));
    }
    [Theory] [InlineData(null)] [InlineData("12")] [InlineData("text")] [InlineData(true)] [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)]
    public void InvalidCurrentValuesRejectBeforeWrites(object? value)
    {
        var book = new OptimizationTests.Book(); book.Cells["C2"].Value = value;
        var p = OptimizationTests.Problem() with { Constraints = [Cell("C2", 1)] };
        var ex = Assert.Throws<ArgumentException>(() => OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => book));
        Assert.Contains("Capacity (Model!C2)", ex.Message); Assert.All(book.Cells.Values, x => Assert.Equal(0, x.Writes));
    }
    [Fact] public void ExcelErrorsAndErrorHresultsRejectButOrdinary2042IsNumeric()
    {
        Assert.Throws<ArgumentException>(() => ConstraintCellSelection.Number(new ExcelCellError("#DIV/0!"), "Labor"));
        Assert.Throws<ArgumentException>(() => ConstraintCellSelection.Number(unchecked((int)0x800A07FA), "Labor"));
        Assert.Throws<ArgumentException>(() => ConstraintCellSelection.Number(new System.Runtime.InteropServices.ErrorWrapper(2042), "Labor"));
        Assert.Equal(2042, ConstraintCellSelection.Number(2042, "Labor"));
    }
    [Fact] public void AFormulaBecomingAnErrorAbortsAndRestoresDecisionFormula()
    {
        var book = new OptimizationTests.Book(); book.Cells["C1"].Formula = true; book.Cells["C2"].Formula = true;
        book.Function = b => { double x = Convert.ToDouble(b.Cells["C1"].Value); b.Cells["C2"].Value = x == 1 ? new ExcelCellError("#DIV/0!") : 2d; return x * 2; };
        var p = OptimizationTests.Problem() with { Variables = [new("x", "X", "Model", "C1", 0, 2, 1)], Constraints = [Cell("C2", 10)] };
        var ex = Assert.Throws<ArgumentException>(() => OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => book));
        Assert.Contains("#DIV/0!", ex.Message); Assert.Equal(50d, book.Cells["C1"].Value); Assert.True(book.Cells["C1"].Formula);
        Assert.Equal(0, book.Cells["C2"].Writes); Assert.True(book.Cells["C2"].Formula); Assert.Equal(2d, book.Cells["C2"].Value);
    }
    [Fact] public void CancellationWithCalculatedConstraintsRestoresAllDecisionInputs()
    {
        var (book, p, f) = ProductMix(); using var cts = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => OptimizationService.Run(p, [OptimizationTests.Assumption()], [f], () => book, cts.Token, _ => cts.Cancel()));
        Assert.Equal(10d, book.Cells["B2"].Value); Assert.Equal(10d, book.Cells["B3"].Value);
        Assert.Equal(60d, book.Cells["B6"].Value); Assert.True(book.Cells["B6"].Formula); Assert.Equal(0, book.Cells["B6"].Writes);
    }
    [Fact] public void ExplicitApplyRecalculatesCellConstraintsAndKeepsFeasibleProductMix()
    {
        var (book, p, f) = ProductMix(); var result = OptimizationService.Run(p, [OptimizationTests.Assumption()], [f], () => book);
        OptimizationService.Apply(result, book, [OptimizationTests.Assumption()], [f]);
        Assert.Equal(80d, book.Cells["B2"].Value); Assert.Equal(60d, book.Cells["B3"].Value);
        Assert.Equal(400d, book.Cells["B6"].Value); Assert.Equal(360d, book.Cells["B7"].Value);
        Assert.Equal(0, book.Cells["B6"].Writes); Assert.True(book.Cells["B6"].Formula);
    }
    [Fact] public void ChangedConstantConstraintDuringApplyRollsBackDecisions()
    {
        var book = new OptimizationTests.Book(); var p = OptimizationTests.Problem() with { Constraints = [Cell("C2", 2)] };
        var result = OptimizationService.Run(p, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()], () => book);
        book.Cells["C2"].Value = 3d;
        Assert.Throws<InvalidOperationException>(() => OptimizationService.Apply(result, book, [OptimizationTests.Assumption()], [OptimizationTests.Forecast()]));
        Assert.Equal(50d, book.Cells["C1"].Value); Assert.Equal(3d, book.Cells["C2"].Value); Assert.Equal(0, book.Cells["C2"].Writes);
    }
    [Theory] [InlineData("Sheet1!b6", "Sheet1", "$B$6")] [InlineData("'My Sheet'!$B$6", "My Sheet", "$B$6")]
    [InlineData("'O''Brien!'!XFD1048576", "O'Brien!", "$XFD$1048576")]
    public void TypedReferencesRoundTripQuotedSheetNames(string text, string sheet, string cell)
    { Assert.Equal((sheet, cell), ConstraintCellSelection.Parse(text)); Assert.Equal((sheet, cell), ConstraintCellSelection.Parse(ConstraintCellSelection.Reference(sheet, cell))); }
    [Theory] [InlineData("B6")] [InlineData("Sheet!B6:B7")] [InlineData("[Other.xlsx]Sheet!B6")]
    [InlineData("'[Other.xlsx]Sheet'!B6")] [InlineData("Sheet!#REF!")] [InlineData("'Bad'Sheet'!A1")]
    public void TypedInvalidExternalAndMultipleReferencesAreRejected(string text) => Assert.Throws<ArgumentException>(() => ConstraintCellSelection.Parse(text));
    [Fact] public void PickerAcceptsFormulaOutputWithoutEditableInputRequirement()
    {
        var source = new PersistenceWorkbook(); var sheet = source.Worksheets.Add(); sheet.Name = "Model"; sheet.Range["B2"].Value2 = "Labor Used";
        var book = new OptimizationTests.Book(); book.Cells["C2"].InputProblem = "Protected array output";
        var selected = ConstraintCellSelection.ReadRange(source, new PersistenceApplication(), sheet.Range["C2"], book)!;
        Assert.Equal("Labor Used", selected.SuggestedName); Assert.Equal(2, selected.CurrentValue); Assert.Equal(0, book.Cells["C2"].Writes); Assert.Empty(source.Names.Items);
        Assert.Null(ConstraintCellSelection.ReadRange(source, new object(), false, book));
    }
    [Fact] public void PickerRejectsForeignOrMultipleCells()
    {
        var source = new PersistenceWorkbook(); var foreign = new PersistenceWorkbook { Name = "Foreign.xlsx" }; var sheet = source.Worksheets.Add();
        Assert.Throws<ArgumentException>(() => ConstraintCellSelection.ReadRange(source, new object(), foreign.Worksheets.Add().Range["C2"], new OptimizationTests.Book()));
        Assert.Throws<ArgumentException>(() => ConstraintCellSelection.ReadRange(source, new object(), new DecisionVariableSelectionTests.MultipleRange(sheet), new OptimizationTests.Book()));
    }
}
