using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace MonteCarlo.Core.Tests;

public class ExcelDnaExportContractTests
{
    private static string Root
    {
        get
        {
            for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory != null; directory = directory.Parent)
                if (File.Exists(Path.Combine(directory.FullName, "PROJECT_CONTEXT.md"))) return directory.FullName;
            throw new DirectoryNotFoundException("Repository source needed for export contract checks.");
        }
    }
    private static string Source(string file) => File.ReadAllText(Path.Combine(Root, "MonteCarlo.Excel", "MonteCarlo.Excel", file));

    [Fact]
    public void ApplicationAssemblyRequiresExplicitExports()
    {
        var project = XDocument.Parse(Source("MonteCarlo.Excel.csproj"));
        Assert.Equal("true", project.Descendants("ExcelAddInExplicitExports").Single().Value);
        Assert.DoesNotContain("ExcelFunction", Source("SimulationSettings.cs"));
        Assert.DoesNotContain("ExcelCommand", Source("SimulationSettings.cs"));
    }

    [Fact]
    public void ExistingWorksheetExportNamesArePreserved()
    {
        var exports = Regex.Matches(Source("MonteCarloFunctions.cs"), @"\[ExcelFunction\(\s*Name\s*=\s*""([^""]+)""")
            .Select(match => match.Groups[1].Value).Order().ToArray();
        Assert.Equal(new[] { "MC.ADD", "MC.NORMAL", "MC.PERT", "MC.PROJECT.COST", "MC.SIMULATE.NORMAL", "MC.SIMULATE.PERT", "MC.SIMULATE.TRIANGULAR", "MC.TRIANGULAR" }, exports);
        Assert.Matches(@"\[ExcelCommand\(\s*Name\s*=\s*""MC_RUN_PROJECTMODEL""", Source("ExcelModelCommands.cs"));
    }

    [Fact]
    public void RibbonCallbacksRemainDiscoverableIndependentlyOfWorksheetExports()
    {
        string ribbon = Source("MonteCarloRibbon.cs");
        var callbacks = Regex.Matches(ribbon, @"onAction='([^']+)'").Select(match => match.Groups[1].Value).ToArray();
        Assert.Equal(13, callbacks.Length);
        foreach (string callback in callbacks) Assert.Contains("public void " + callback, ribbon);
        Assert.Contains(": ExcelRibbon", ribbon);
        Assert.Contains(": IExcelAddIn", Source("MonteCarloAddIn.cs"));
    }

    [Fact]
    public void CleanStartupDoesNotDependOnAnExternalOfficePia()
    {
        var project = XDocument.Parse(Source("MonteCarlo.Excel.csproj"));
        var reference = project.Descendants("Reference").Single(e => (string?)e.Attribute("Include") == "Microsoft.Office.Interop.Excel");
        Assert.Equal("true", reference.Element("EmbedInteropTypes")?.Value);
        Assert.Equal("false", reference.Element("Private")?.Value);
        var package = project.Descendants("PackageReference").Single(e => (string?)e.Attribute("Include") == "Microsoft.Office.Interop.Excel");
        Assert.Equal("compile;runtime", (string?)package.Attribute("ExcludeAssets"));
    }

    [Fact]
    public void ShippedConfigurationIsQuietAndInstallerRetainsItsName()
    {
        var config = XDocument.Parse(Source("App.config"));
        Assert.Equal("false", (string?)config.Root?.Element("monteCarloDiagnostics")?.Attribute("verbose"));
        Assert.Null(config.Root?.Element("system.diagnostics")); // Not supported by .NET 10.
        string installer = File.ReadAllText(Path.Combine(Root, "Installer", "MonteCarloForExcel.iss"));
        Assert.Contains("Source: \"{#SourceXll}.config\"", installer);
        Assert.Contains("DestName: \"{#MyAppFileName}.config\"", installer);
    }
}
