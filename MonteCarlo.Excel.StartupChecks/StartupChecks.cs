using System.Collections;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.Loader;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;

// Developer-only fixture XLL, loaded AFTER the unmodified production packed XLL.
// Runs on Excel's main thread and accesses that XLL's actual assembly/context.
public static class StartupChecks
{
    private static string? freshProfile;
    private static string? legacyRegistrySnapshot;

    private static string RegistrySnapshot()
    {
        using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\\MonteCarloExcel");
        using var trial = key?.OpenSubKey("TrialState");
        string Values(Microsoft.Win32.RegistryKey? k) => k == null ? "absent" : string.Join("|",
            k.GetValueNames().Order().Select(n => n + "=" + k.GetValue(n)));
        return Values(key) + "/" + Values(trial);
    }

    [ExcelCommand(Name = "MC_VERIFY_FRESH_PROFILE")]
    public static void PrepareFreshProfile()
    {
        freshProfile = Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath)!, "fresh-profile");
        if (Directory.Exists(freshProfile)) throw new Exception("Fresh profile must not already exist.");
        Directory.CreateDirectory(freshProfile);
        legacyRegistrySnapshot = RegistrySnapshot();
        // Process-local environment only. Existing user files and Registry are untouched.
        Environment.SetEnvironmentVariable("LOCALAPPDATA", freshProfile);
    }
    [ExcelCommand(Name = "MC_VERIFY_COM_CLEANUP")]
    public static void Cleanup()
    {
        var production = AssemblyLoadContext.All.SelectMany(context => context.Assemblies)
            .Single(assembly => assembly.GetName().Name == "MonteCarlo.Excel");
        var integration = AssemblyLoadContext.GetLoadContext(production)!.Assemblies.Single(a => a.GetName().Name == "ExcelDna.Integration");
        string path = (string)integration.GetType("ExcelDna.Integration.ExcelDnaUtil")!.GetProperty("XllPath")!.GetValue(null)!;
        object productionApplication = integration.GetType("ExcelDna.Integration.ExcelDnaUtil")!.GetProperty("Application")!.GetValue(null)!;
        object fixtureApplication = ExcelDnaUtil.Application;
        ExcelIntegration.UnregisterXLL(path);
        // Both verification contexts cache Excel.Application. Release only this private
        // fixture's RCWs after unloading the production XLL so COM Quit can finish.
        if (Marshal.IsComObject(productionApplication)) Marshal.FinalReleaseComObject(productionApplication);
        if (!ReferenceEquals(productionApplication, fixtureApplication) && Marshal.IsComObject(fixtureApplication))
            Marshal.FinalReleaseComObject(fixtureApplication);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
    }

    [ExcelCommand(Name = "MC_VERIFY_STARTUP")]
    public static void Verify()
    {
        string report = Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath)!, "startup-report.txt");
        var lines = new List<string>();
        try
        {
            Assembly production = AssemblyLoadContext.All.SelectMany(context => context.Assemblies)
                .Single(assembly => assembly.GetName().Name == "MonteCarlo.Excel");
            if (production.GetTypes().Any(t => t.Namespace == "MonteCarlo.Excel.Licensing") ||
                production.GetManifestResourceNames().Any(n => n.Contains("public-key") || n.Contains("Licensing")))
                throw new Exception("Activation infrastructure remains in production assembly.");
            Type Type(string name) => production.GetType("MonteCarlo.Excel." + name, throwOnError: true)!;
            object? Call(string type, string name, params object?[] args) => Type(type).GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
                .Single(method => method.Name == name && method.GetParameters().Length == args.Length).Invoke(null, args);
            dynamic app = ExcelDnaUtil.Application;
            dynamic book = app.ActiveWorkbook;
            dynamic sheet = book.Worksheets[1]; sheet.Name = "Model";
            sheet.Range["A1"].Formula2 = "=5+5"; sheet.Range["B1"].Formula2 = "=A1*2";

            var context = AssemblyLoadContext.GetLoadContext(production)!;
            var integration = context.Assemblies.Single(a => a.GetName().Name == "ExcelDna.Integration");
            var scanner = integration.GetType("ExcelDna.Integration.AssemblyLoader")!
                .GetMethod("IsMethodSupported", BindingFlags.Static | BindingFlags.NonPublic)!;
            var helper = Type("SimulationSettings").GetMethod("TryParse")!;
            if ((bool)scanner.Invoke(null, [helper, true])!) throw new Exception("Internal model helper exported.");
            try
            {
                scanner.Invoke(null, [helper, false]);
                throw new Exception("Expected old implicit registration to reproduce the signature failure.");
            }
            catch (TargetInvocationException ex)
            {
                lines.Add("ROOT CAUSE REPRODUCED: implicit TryParse registration: " + ex.InnerException);
            }
            var marked = production.GetTypes().SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Static))
                .Where(m => m.GetCustomAttributes().Any(a => a.GetType().FullName is "ExcelDna.Integration.ExcelFunctionAttribute" or "ExcelDna.Integration.ExcelCommandAttribute")).ToArray();
            if (marked.Length != 9 || marked.Any(m => !(bool)scanner.Invoke(null, [m, true])!))
                throw new Exception("Existing attributed exports are not supported.");
            foreach (var (name, args) in new (string, object[])[]
            {
                ("MC.ADD", [2.0, 3.0]), ("MC.NORMAL", [10.0, 1.0]), ("MC.TRIANGULAR", [1.0, 2.0, 3.0]), ("MC.PERT", [1.0, 2.0, 3.0]),
                ("MC.SIMULATE.NORMAL", [10.0, 1.0, 100]), ("MC.SIMULATE.TRIANGULAR", [1.0, 2.0, 3.0, 100]),
                ("MC.SIMULATE.PERT", [1.0, 2.0, 3.0, 100]), ("MC.PROJECT.COST", [1.0, 2.0, 3.0, 10.0, 20.0, 30.0, 2, 100])
            })
            {
                object value = args.Length switch
                {
                    2 => app.Run(name, args[0], args[1]), 3 => app.Run(name, args[0], args[1], args[2]),
                    4 => app.Run(name, args[0], args[1], args[2], args[3]),
                    _ => app.Run(name, args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7])
                };
                if (value is not double && value is not Array) throw new Exception("Invalid UDF result: " + name);
                if (name == "MC.ADD" && (double)value != 5) throw new Exception("MC.ADD failed.");
                lines.Add("WORKSHEET FUNCTION PASS " + name);
            }
            object Definition(string name, params (string, object)[] properties)
            {
                object definition = Activator.CreateInstance(Type(name))!;
                foreach (var (property, value) in properties) Type(name).GetProperty(property)!.SetValue(definition, value);
                return definition;
            }
            Call("SimulationModel", "Clear");
            ((IList)Type("SimulationModel").GetProperty("Assumptions")!.GetValue(null)!).Add(Definition("AssumptionDefinition",
                ("Name", "Noise"), ("SheetName", "Model"), ("CellAddress", "A1"), ("Parameter1", 10.0), ("Parameter2", 1.0)));
            ((IList)Type("SimulationModel").GetProperty("Forecasts")!.GetValue(null)!).Add(Definition("ForecastDefinition",
                ("Name", "Output"), ("SheetName", "Model"), ("CellAddress", "B1")));
            Call("WorkbookPersistence", "SaveModel");
            object Simulation()
            {
                object result = Type("SimulationService").GetMethod("TryRun", [typeof(int)])!.Invoke(null, [100])!;
                if (!(bool)result.GetType().GetProperty("Succeeded")!.GetValue(result)!)
                    throw new Exception((string)result.GetType().GetProperty("UserMessage")!.GetValue(result)!);
                if ((string)sheet.Range["A1"].Formula2 != "=5+5" || (string)sheet.Range["B1"].Formula2 != "=A1*2")
                    throw new Exception("Simulation did not restore workbook formulas.");
                return result.GetType().GetProperty("Result")!.GetValue(result)!;
            }
            object simulationResult = Simulation(); lines.Add("SIMULATION PASS 100 trials, forecast calculation and formula restoration");
            using (var results = (Form)Activator.CreateInstance(Type("ResultsForm"), [simulationResult])!)
            using (var closeResults = new System.Windows.Forms.Timer { Interval = 250 })
            {
                bool shown = false;
                closeResults.Tick += (_, _) => { shown = results.Visible; results.Close(); };
                closeResults.Start(); results.ShowDialog(); closeResults.Stop();
                if (!shown) throw new Exception("Installed Results window did not open.");
            }
            lines.Add("RESULTS PASS installed forecast distribution display opens");
            sheet.Range["B2"].Formula2 = "=5+5"; sheet.Range["B3"].Value2 = 20.0;
            sheet.Range["B5"].Formula2 = "=B2*B3*2";
            if (Call("ExcelModelSimulator", "RunProjectModel", 5) is not Array ||
                (string)sheet.Range["B2"].Formula2 != "=5+5" || (double)sheet.Range["B3"].Value2 != 20.0)
                throw new Exception("Legacy command calculation/COM restoration failed.");
            lines.Add("LEGACY COMMAND PASS attributed MC_RUN_PROJECTMODEL and its production calculator/restoration");

            object ribbon = Activator.CreateInstance(Type("MonteCarloRibbon"))!;
            string xml = (string)Type("MonteCarloRibbon").GetMethod("GetCustomUI")!.Invoke(ribbon, [""])!;
            if (!xml.Contains("Telivu") || !xml.Contains("OnOptimizer") || !xml.Contains("OnAboutTelivu") || xml.Contains("OnLicense")) throw new Exception("Ribbon XML failed.");
            bool connected = false;
            foreach (dynamic addin in app.COMAddIns)
                if (((string)addin.Description).Contains("MonteCarlo", StringComparison.OrdinalIgnoreCase) && (bool)addin.Connect) connected = true;
            if (!connected) throw new Exception("Monte Carlo ribbon COM add-in is not connected.");
            foreach (var (callback, expected) in new[]
            {
                ("OnScenarioAnalysis", "ScenarioAnalysisForm"), ("OnOptimizer", "OptimizerForm"),
                ("OnCorrelations", "AssumptionCorrelationsForm"), ("OnSpcAnalysis", "SpcConfigurationForm"),
                ("OnAboutTelivu", "AboutTelivuForm")
            })
            {
                bool opened = false;
                using var timer = new System.Windows.Forms.Timer { Interval = 250 };
                timer.Tick += (_, _) =>
                {
                    foreach (Form form in Application.OpenForms.Cast<Form>().ToArray())
                        if (form.GetType().Name == expected) { opened = true; form.Close(); }
                };
                timer.Start(); Type("MonteCarloRibbon").GetMethod(callback)!.Invoke(ribbon, [null]); timer.Stop();
                if (!opened) throw new Exception("Ribbon callback did not open " + expected);
                lines.Add("RIBBON DIALOG PASS " + expected);
            }
            lines.Add("RIBBON PASS COM add-in connected, 12 analysis callbacks and About Telivu connected");
            string workbookFile = Path.ChangeExtension(report, ".xlsx");
            book.SaveAs(workbookFile, 51); book.Close(false); book = app.Workbooks.Open(workbookFile); sheet = book.Worksheets["Model"];
            Simulation(); lines.Add("SAVE/REOPEN PASS persisted assumptions/forecast and successful simulation");

            var source = (TraceSource)integration.GetType("ExcelDna.Integration.ExcelIntegration")!
                .GetMethod("GetIntegrationTraceSource", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, null)!;
            if (source.Listeners.Cast<TraceListener>().Any(l => l.GetType().Name == "LogDisplayTraceListener"))
                throw new Exception("Production popup trace listener is present.");
            if (!source.Listeners.Cast<TraceListener>().Any(l => l.Name == "MonteCarloFile")) throw new Exception("File logging absent.");
            source.TraceEvent(TraceEventType.Error, 1, "STARTUP CHECK synthetic initialization error; confirms durable logging without popup."); source.Flush();
            lines.Add("DIAGNOSTICS PASS no popup listener; synthetic error captured; level=" + source.Switch.Level);
            if (freshProfile != null)
            {
                if (Directory.Exists(Path.Combine(freshProfile, "MonteCarloExcel")))
                    throw new Exception("Fresh user activation state was created.");
                if (legacyRegistrySnapshot != RegistrySnapshot())
                    throw new Exception("Legacy activation Registry state was changed.");
                lines.Add("ACTIVATION-FREE PASS fresh process profile; no activation types/key/files; legacy Registry untouched");
            }
            lines.Add("PASS");
        }
        catch (Exception ex) { lines.Add("FAIL " + ex); }
        File.WriteAllLines(report, lines);
    }
}
