using MonteCarlo.Excel;
namespace MonteCarlo.Core.Tests;

public sealed class DecisionVariableSelectionTests
{
    [Theory] [InlineData("c1", "$C$1")] [InlineData("$C$1", "$C$1")] [InlineData("AB123", "$AB$123")]
    public void AbsoluteReference(string cell, string expected) => Assert.Equal(expected, DecisionVariableSelection.Absolute(cell));
    [Theory] [InlineData("C1:C2")] [InlineData("#REF!")] [InlineData("")]
    public void InvalidReference(string cell) => Assert.Throws<ArgumentException>(() => DecisionVariableSelection.Absolute(cell));
    [Fact] public void ReadSelectionCapturesReferenceValueAndLabelWithoutWriting()
    {
        var book = new OptimizationTests.Book(); var selected = DecisionVariableSelection.Read(book, "Model", "C1", " Selling Price ");
        Assert.Equal("Model!$C$1", selected.Display); Assert.Equal(50, selected.CurrentValue); Assert.Equal("Selling Price", selected.SuggestedName);
        Assert.Equal(0, book.Cells["C1"].Writes); Assert.Equal(0, book.Calculations);
    }
    [Theory] [InlineData(null)] [InlineData("")] [InlineData("text")] [InlineData(double.NaN)]
    public void InvalidValuesRejected(object? value)
    { var book = new OptimizationTests.Book(); book.Cells["C1"].Value = value; Assert.Throws<ArgumentException>(() => DecisionVariableSelection.Read(book, "Model", "C1")); Assert.Equal(0, book.Cells["C1"].Writes); }
    [Fact] public void ExcelErrorRejected()
    { var book = new OptimizationTests.Book(); book.Cells["C1"].Value = new ExcelCellError("#N/A"); Assert.Throws<ArgumentException>(() => DecisionVariableSelection.Read(book, "Model", "C1")); }
    [Theory] [InlineData("merged")] [InlineData("array")] [InlineData("spill")] [InlineData("protected")]
    public void UnsafeInputRejected(string reason)
    { var book = new OptimizationTests.Book(); book.Cells["C1"].InputProblem = reason; Assert.Throws<ArgumentException>(() => DecisionVariableSelection.Read(book, "Model", "C1")); }
    [Fact] public void UserNameNotOverwritten() => Assert.Equal("My price", DecisionVariableSelection.SuggestName("My price", "Selling Price"));
    [Theory] [InlineData(null)] [InlineData(123)] [InlineData(" ")] [InlineData("Line\nbreak")]
    public void UnreasonableLabelsDoNotSuggest(object? label) => Assert.Equal("", DecisionVariableSelection.SuggestName("", label));
    [Fact] public void OffGridCurrentValueIsAccepted()
    { var book = new OptimizationTests.Book(); book.Cells["C1"].Value = 103d; Assert.Equal(103, DecisionVariableSelection.Read(book, "Model", "C1").CurrentValue); }
    [Fact] public void RangeCaptureUsesLeftLabelAndDoesNotMutateWorkbook()
    {
        var book = new PersistenceWorkbook(); var sheet = book.Worksheets.Add(); sheet.Name = "Model"; sheet.Range["B1"].Value2 = "Selling Price"; sheet.Range["C1"].Value2 = 50d;
        var selected = DecisionVariableSelection.ReadRange(book, new PersistenceApplication(), sheet.Range["C1"], new OptimizationTests.Book());
        Assert.Equal("Model!$C$1", selected!.Display); Assert.Equal("Selling Price", selected.SuggestedName); Assert.Equal(50d, sheet.Range["C1"].Value2); Assert.Empty(book.Names.Items);
    }
    [Fact] public void ForeignWorkbookRejected()
    {
        var book = new PersistenceWorkbook { Name = "Other.xlsx" }; var cell = book.Worksheets.Add().Range["C1"];
        Assert.Throws<ArgumentException>(() => DecisionVariableSelection.ReadRange(new PersistenceWorkbook(), new PersistenceApplication(), cell, new OptimizationTests.Book()));
    }
    [Fact] public void PickerCancellationIsNoOp() => Assert.Null(DecisionVariableSelection.ReadRange(new object(), new object(), false, new OptimizationTests.Book()));
    public sealed class MultipleRange(PersistenceSheet sheet)
    { public PersistenceSheet Worksheet => sheet; public PersistenceCount Cells => new(2); public PersistenceCount Areas => new(1); }
    [Fact] public void MultipleCellsRejected()
    { var book = new PersistenceWorkbook(); Assert.Throws<ArgumentException>(() => DecisionVariableSelection.ReadRange(book, new PersistenceApplication(), new MultipleRange(book.Worksheets.Add()), new OptimizationTests.Book())); }
}

