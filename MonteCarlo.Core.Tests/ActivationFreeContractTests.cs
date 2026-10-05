using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MonteCarlo.Core.Tests;

public class ActivationFreeContractTests
{
    private static string Root
    {
        get
        {
            for (var d = new DirectoryInfo(AppContext.BaseDirectory); d != null; d = d.Parent)
                if (File.Exists(Path.Combine(d.FullName, "PROJECT_CONTEXT.md"))) return d.FullName;
            throw new DirectoryNotFoundException("Repository source required.");
        }
    }
    private static string App => Path.Combine(Root, "MonteCarlo.Excel", "MonteCarlo.Excel");

    [Fact]
    public void RuntimeHasNoActivationArchitectureOrStateAccess()
    {
        foreach (var file in Directory.EnumerateFiles(App, "*.cs", SearchOption.AllDirectories)
            .Where(f => !f.Split(Path.DirectorySeparatorChar).Any(p => p is "obj" or "bin")))
        {
            string source = File.ReadAllText(file);
            Assert.DoesNotMatch(@"\b(LicenseService|TrialService|LicenseStorage|LocalLicenseValidator|LicenseInfo|LicenseForm|ActivationForm|DevelopmentMode|EnsureAccess)\b", source);
            Assert.DoesNotContain("MonteCarloExcel", source); // Former activation storage namespace.
            Assert.DoesNotContain("TrialState", source);
        }
    }

    [Fact]
    public void FunctionalRibbonRetainsCommandsAndReplacesActivationWithHelp()
    {
        string ribbon = File.ReadAllText(Path.Combine(App, "MonteCarloRibbon.cs"));
        var callbacks = Regex.Matches(ribbon, @"onAction='([^']+)'").Select(m => m.Groups[1].Value).ToArray();
        Assert.Equal(13, callbacks.Length);
        Assert.Contains("OnAboutTelivu", callbacks);
        Assert.DoesNotContain("OnLicense", callbacks);
        Assert.Contains("label='Telivu'", ribbon);
        Assert.Contains("id='HelpGroup'", ribbon);
        Assert.Contains("label='About Telivu'", ribbon);
        foreach (string callback in new[] { "OnDefineAssumption", "OnDefineForecast", "OnCorrelations", "OnModelManager", "OnClearCellDefinition", "OnRunSimulation", "OnScenarioAnalysis", "OnOptimizer", "OnSimulationSettings", "OnReloadModel", "OnSpcAnalysis", "OnClearModel" })
            Assert.Contains(callback, callbacks);
    }

    [Fact]
    public void SolutionAndRuntimeDoNotReferenceGeneratorOrVerificationKey()
    {
        string solution = File.ReadAllText(Path.Combine(Root, "MonteCarlo.Excel", "MonteCarlo.Excel.slnx"));
        Assert.DoesNotContain("LicenseGenerator", solution);
        var project = XDocument.Load(Path.Combine(App, "MonteCarlo.Excel.csproj"));
        Assert.DoesNotContain(project.Descendants("EmbeddedResource"), e => ((string?)e.Attribute("Include") ?? "").Contains("Licensing"));
        Assert.False(File.Exists(Path.Combine(Root, "MonteCarlo.LicenseGenerator", "MonteCarlo.LicenseGenerator.csproj")));
    }

    [Fact]
    public void InstallerRetainsRegistrationAndIgnoresLegacyActivationData()
    {
        string installer = File.ReadAllText(Path.Combine(Root, "Installer", "MonteCarloForExcel.iss"));
        Assert.Contains("9B4D8E57-7B33-4C68-9D1D-91D5E4D8F001", installer);
        Assert.DoesNotContain("TrialState", installer);
        Assert.DoesNotContain("license.dat", installer);
        Assert.DoesNotContain("trial.dat", installer);
        Assert.Contains("No activation is required.", installer);
        string build = File.ReadAllText(Path.Combine(Root, "build-installer.ps1"));
        Assert.DoesNotContain("LicensingChecks", build);
    }
}
