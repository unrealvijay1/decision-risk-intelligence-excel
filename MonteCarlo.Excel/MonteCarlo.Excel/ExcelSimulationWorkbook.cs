using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace MonteCarlo.Excel;

internal sealed class ExcelSimulationWorkbook : ISimulationWorkbook
{
    private readonly dynamic application;
    private readonly dynamic workbook;
    private readonly Action? calculate;
    public ExcelSimulationWorkbook(object application, object workbook, Action? calculate = null)
    {
        this.application = application;
        this.workbook = workbook;
        this.calculate = calculate;
    }

    public bool ScreenUpdating { get => application.ScreenUpdating; set => application.ScreenUpdating = value; }
    public bool EnableEvents { get => application.EnableEvents; set => application.EnableEvents = value; }
    public object? StatusBar { get => application.StatusBar; set => application.StatusBar = value; }
    public void Calculate()
    {
        if (calculate != null) calculate();
        else application.Calculate();
    }

    public ISimulationCell Resolve(string sheet, string address)
    {
        try
        {
            dynamic worksheet = workbook.Worksheets[sheet];
            dynamic cell = worksheet.Range[address];
            if (Convert.ToDouble(cell.Cells.CountLarge) != 1 || Convert.ToInt32(cell.Areas.Count) != 1 ||
                !string.Equals((string)cell.Worksheet.Name, sheet, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals((string)cell.Worksheet.Parent.Name, (string)workbook.Name, StringComparison.OrdinalIgnoreCase))
                throw new SimulationCellAccessException("Select one cell in the configured worksheet.");
            return new ExcelSimulationCell(application, cell);
        }
        catch (COMException ex) { throw new SimulationCellAccessException("The cell reference is unavailable.", ex); }
    }

    internal static object? ReadCellValue(dynamic application, dynamic cell)
    {
        // Pass the Range itself to Excel, before VT_ERROR can be mistaken for a number.
        if ((bool)application.WorksheetFunction.IsError(cell))
            return new ExcelCellError(Convert.ToString(cell.Text) ?? "(formula error)");
        return cell.Value2;
    }

    private sealed class ExcelSimulationCell : ISimulationCell
    {
        private readonly dynamic application;
        private readonly dynamic cell;
        public ExcelSimulationCell(object application, object cell) { this.application = application; this.cell = cell; }
        public string Identity => $"{cell.Worksheet.Name}!{cell.Address[false, false]}";
        public string? InputProblem
        {
            get
            {
                if ((bool)cell.MergeCells || (bool)cell.HasArray)
                    return "The input belongs to a merged cell or array formula.";
                try
                {
                    if ((bool)cell.HasSpill) return "The input belongs to a spilled formula range.";
                }
                catch (RuntimeBinderException) { } // Older Excel has no dynamic-array properties.
                catch (COMException ex) when (MissingMember(ex)) { }
                if ((bool)cell.Worksheet.ProtectContents && (bool)cell.Locked)
                    return "The input cell is protected against editing.";
                return null;
            }
        }
        public object? ReadValue() => ReadCellValue(application, cell);
        public SimulationCellContent Capture()
        {
            if (!(bool)cell.HasFormula) return new((object?)cell.Value2);
            try { return new((object?)cell.Formula2, true, true); }
            catch (RuntimeBinderException) { return new((object?)cell.Formula, true); }
            catch (COMException ex) when (MissingMember(ex)) { return new((object?)cell.Formula, true); }
        }
        public void WriteSample(double value) => cell.Value2 = value;
        public void Restore(SimulationCellContent content)
        {
            if (content.UsesFormula2) cell.Formula2 = content.Value;
            else if (content.IsFormula) cell.Formula = content.Value;
            else cell.Value2 = content.Value;
        }
        private static bool MissingMember(COMException ex) =>
            ex.HResult == unchecked((int)0x80020003) || ex.HResult == unchecked((int)0x80020006);
    }
}
