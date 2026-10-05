using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml.Linq;
using ExcelDna.Integration;

namespace MonteCarlo.Excel;

// Excel-DNA initializes its TraceSource before loading ExternalLibrary assemblies.
// AutoOpen is too late: registration diagnostics occur before that callback.
internal static class AddInDiagnostics
{
    private static TraceListener? applicationListener;
    // Required before AutoOpen for registration diagnostics in this add-in host.
    #pragma warning disable CA2255
    [ModuleInitializer]
    #pragma warning restore CA2255
    internal static void Initialize()
    {
        // This accessor is internal in the pinned Excel-DNA 1.9 API. Keep the live
        // startup check when upgrading Excel-DNA; do not use process-wide registry/env settings.
        var accessor = typeof(ExcelIntegration).GetMethod("GetIntegrationTraceSource", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new MissingMethodException("Excel-DNA diagnostic TraceSource accessor changed.");
        if (accessor.Invoke(null, null) is not TraceSource source) return; // Non-Excel tooling.

        foreach (TraceListener listener in source.Listeners.Cast<TraceListener>().ToArray())
            if (listener is ExcelDna.Logging.LogDisplayTraceListener) source.Listeners.Remove(listener);

        bool verbose = Environment.GetEnvironmentVariable("MONTECARLO_DIAGNOSTICS_VERBOSE") == "1";
        Exception? configurationError = null;
        try
        {
            string config = ExcelDnaUtil.XllPath + ".config";
            if (File.Exists(config))
                verbose |= (string?)XDocument.Load(config).Root?.Element("monteCarloDiagnostics")?.Attribute("verbose") == "true";
        }
        catch (Exception ex) { configurationError = ex; }

        source.Switch.Level = verbose ? SourceLevels.Verbose : SourceLevels.Warning;
        try
        {
            string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MonteCarlo", "Logs");
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, $"ExcelDna-{DateTime.UtcNow:yyyyMMdd}-{Environment.ProcessId}.log");
            var listener = new TextWriterTraceListener(new StreamWriter(path, append: true) { AutoFlush = true }, "MonteCarloFile")
                { TraceOutputOptions = TraceOptions.DateTime };
            source.Listeners.Add(listener);
            Trace.Listeners.Add(listener); // Existing application initialization/restoration Trace errors.
            applicationListener = listener;
            listener.WriteLine($"{DateTime.UtcNow:O} Monte Carlo diagnostics initialized; explicit exports; mode={(verbose ? "Verbose" : "Production")}; XLL={ExcelDnaUtil.XllPath}");
            if (configurationError != null) Trace.TraceError("Monte Carlo diagnostic configuration could not be read: {0}", configurationError);
        }
        catch (Exception ex)
        {
            // A failed log destination must not prevent Excel loading. Default debugger
            // logging remains available; this is restricted to diagnostic I/O failures.
            Trace.TraceError("Monte Carlo diagnostic file could not be opened: {0}", ex);
        }
    }

    internal static void Shutdown()
    {
        // Excel-DNA owns closing the source/listener after AutoClose. Do not leave
        // a closed listener in the shared application Trace collection on reload.
        if (applicationListener != null) Trace.Listeners.Remove(applicationListener);
        applicationListener = null;
    }
}
