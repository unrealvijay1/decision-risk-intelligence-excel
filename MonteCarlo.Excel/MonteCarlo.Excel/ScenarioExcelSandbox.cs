using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace MonteCarlo.Excel;

public sealed record ScenarioCalculationOptions(bool Iteration, int MaxIterations, double MaxChange);

/// <summary>One private Excel application and workbook for the entire comparison. No source COM objects cross threads.</summary>
public sealed class ScenarioExcelSandbox : IScenarioSandbox
{
    private readonly List<object> owned = new();
    private dynamic? application;
    private dynamic? book;
    private bool disposed;
    public ISimulationWorkbook Workbook { get; private set; } = null!;
    private dynamic Own(object value) { owned.Add(value); return value; }
    private static void Release(object? value)
    { if (OperatingSystem.IsWindows() && value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

    public static ScenarioTemporaryFile Prepare(object sourceWorkbook)
    {
        dynamic source = sourceWorkbook;
        if ((bool)source.HasVBProject)
            throw new InvalidOperationException("Scenario Analysis V1 does not support workbooks with VBA. Macros cannot run in the disposable copy.");
        object? connections = null;
        try
        {
            connections = source.Connections;
            if ((int)((dynamic)connections).Count != 0 || source.LinkSources(1) != null || source.LinkSources(2) != null)
                throw new InvalidOperationException("Scenario Analysis V1 requires a self-contained workbook without external links or data connections.");
        }
        finally { Release(connections); }
        string name = (string)source.Name;
        if (!System.IO.Path.HasExtension(name))
            name += (int)source.FileFormat switch { 50 => ".xlsb", 52 => ".xlsm", 56 => ".xls", _ => ".xlsx" };
        return ScenarioTemporaryFile.Create(name, path => source.SaveCopyAs(path));
    }

    public static ScenarioExcelSandbox Open(string path, ScenarioCalculationOptions options, Func<object>? createApplication = null)
    {
        var sandbox = new ScenarioExcelSandbox();
        try
        {
            if (createApplication == null && !OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("Scenario Analysis requires Windows Excel.");
            sandbox.application = sandbox.Own(createApplication != null ? createApplication() :
                Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Microsoft Excel is not installed."))!);
            dynamic app = sandbox.application;
            app.Visible = false; app.DisplayAlerts = false; app.EnableEvents = false;
            app.AutomationSecurity = 3; // ForceDisable: do not execute copied VBA.
            app.AskToUpdateLinks = false;
            dynamic books = sandbox.Own(app.Workbooks);
            sandbox.book = sandbox.Own(books.Open(path, UpdateLinks: 0, ReadOnly: false,
                IgnoreReadOnlyRecommended: true, AddToMru: false));
            dynamic book = sandbox.book;
            if (book.LinkSources(1) != null || book.LinkSources(2) != null)
                throw new InvalidOperationException("The disposable copy contains external references. Scenario Analysis cannot safely evaluate this model.");
            app.Calculation = -4135; // Manual between trials; the existing engine explicitly recalculates.
            app.Iteration = options.Iteration; app.MaxIterations = options.MaxIterations; app.MaxChange = options.MaxChange;
            sandbox.Workbook = new SandboxWorkbook(sandbox, app, book);
            return sandbox;
        }
        catch { sandbox.Dispose(); throw; }
    }

    public void Dispose()
    {
        if (disposed) return;
        disposed = true;
        var failures = new List<Exception>();
        void Attempt(Action action)
        {
            try { action(); }
            catch { try { action(); } catch (Exception ex) { failures.Add(ex); } }
        }
        if (book != null) Attempt(() => book.Close(SaveChanges: false));
        if (application != null) Attempt(() => application.Quit());
        for (int i = owned.Count - 1; i >= 0; i--)
        { try { Release(owned[i]); } catch (Exception ex) { failures.Add(ex); } }
        owned.Clear(); book = null; application = null;
        if (failures.Count > 0) throw new AggregateException("Excel could not fully close the scenario sandbox. Check for an unresponsive Excel process.", failures);
    }

    // Cache owned ranges/functions once, avoiding transient COM wrappers in the trial loop.
    private sealed class SandboxWorkbook : ISimulationWorkbook
    {
        private readonly ScenarioExcelSandbox owner;
        private readonly dynamic app, sheets, functions;
        private readonly Dictionary<string, ISimulationCell> cells = new(StringComparer.OrdinalIgnoreCase);
        public SandboxWorkbook(ScenarioExcelSandbox owner, object app, object book)
        {
            this.owner = owner; this.app = app;
            sheets = owner.Own(((dynamic)book).Worksheets);
            functions = owner.Own(((dynamic)app).WorksheetFunction);
        }
        public bool ScreenUpdating { get => app.ScreenUpdating; set => app.ScreenUpdating = value; }
        public bool EnableEvents { get => app.EnableEvents; set => app.EnableEvents = value; }
        public object? StatusBar { get => app.StatusBar; set => app.StatusBar = value; }
        public void Calculate() => app.Calculate(); // Only this private instance can be recalculated.
        public ISimulationCell Resolve(string sheet, string address)
        {
            string key = sheet + "!" + address;
            if (cells.TryGetValue(key, out var existing)) return existing;
            try
            {
                dynamic worksheet = owner.Own(sheets[sheet]);
                dynamic range = owner.Own(worksheet.Range[address]);
                dynamic rangeCells = owner.Own(range.Cells), areas = owner.Own(range.Areas);
                if (Convert.ToDouble(rangeCells.CountLarge) != 1 || (int)areas.Count != 1)
                    throw new SimulationCellAccessException("Select a single input/output cell.");
                var cell = new SandboxCell(range, worksheet, functions, key);
                cells.Add(key, cell);
                return cell;
            }
            catch (COMException ex) { throw new SimulationCellAccessException("The copied cell could not be resolved.", ex); }
        }
    }
    private sealed class SandboxCell(dynamic range, dynamic sheet, dynamic functions, string identity) : ISimulationCell
    {
        public string Identity => identity;
        public string? InputProblem
        {
            get
            {
                if ((bool)range.MergeCells || (bool)range.HasArray) return "The input belongs to a merged cell or array formula.";
                try { if ((bool)range.HasSpill) return "The input belongs to a spilled formula range."; }
                catch (RuntimeBinderException) { }
                catch (COMException ex) when (MissingMember(ex)) { }
                return (bool)sheet.ProtectContents && (bool)range.Locked ? "The input is protected." : null;
            }
        }
        public object? ReadValue() => (bool)functions.IsError(range) ? new ExcelCellError(Convert.ToString(range.Text) ?? "formula error") : range.Value2;
        public SimulationCellContent Capture()
        {
            if (!(bool)range.HasFormula) return new((object?)range.Value2);
            try { return new((object?)range.Formula2, true, true); }
            catch (RuntimeBinderException) { return new((object?)range.Formula, true); }
            catch (COMException ex) when (MissingMember(ex)) { return new((object?)range.Formula, true); }
        }
        public void WriteSample(double value) => range.Value2 = value;
        public void Restore(SimulationCellContent content)
        {
            if (content.UsesFormula2) range.Formula2 = content.Value;
            else if (content.IsFormula) range.Formula = content.Value;
            else range.Value2 = content.Value;
        }
        private static bool MissingMember(COMException ex) => ex.HResult == unchecked((int)0x80020003) || ex.HResult == unchecked((int)0x80020006);
    }
}
