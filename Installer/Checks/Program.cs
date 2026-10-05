using System.Reflection;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;

if (args.Length != 3) throw new ArgumentException("Usage: Checks <extracted Excel.dll> <repository root> <version>");
string dll = args[0], root = args[1], version = args[2];
using var pe = new PEReader(File.OpenRead(dll));
var md = pe.GetMetadataReader();
if (md.GetAssemblyDefinition().Version.ToString() != version + ".0") throw new Exception("Assembly version mismatch.");
var names = md.TypeDefinitions.Select(h => md.GetTypeDefinition(h)).Select(t => md.GetString(t.Namespace) + "." + md.GetString(t.Name)).ToHashSet();
if (names.Any(name => name.StartsWith("MonteCarlo.Excel.Licensing.", StringComparison.Ordinal)))
    throw new Exception("Obsolete activation architecture remains in the runtime.");
foreach (string name in new[] { "MonteCarloRibbon", "MonteCarloAddIn", "SimulationService", "ScenarioSimulationService", "OptimizationService", "AssumptionCorrelationsForm", "SpcConfigurationForm", "SpcResultsForm", "AboutTelivuForm" })
    if (!names.Contains("MonteCarlo.Excel." + name)) throw new Exception("Missing product type " + name);
var resourceNames = new List<string>();
foreach (var handle in md.ManifestResources) {
    var resource = md.GetManifestResource(handle);
    string name = md.GetString(resource.Name); resourceNames.Add(name);
    if (!resource.Implementation.IsNil) throw new Exception("External resource: " + name);
    int rva = pe.PEHeaders.CorHeader!.ResourcesDirectory.RelativeVirtualAddress;
    var reader = pe.GetSectionData(rva).GetReader((int)resource.Offset, pe.GetSectionData(rva).Length - (int)resource.Offset);
    byte[] bytes = reader.ReadBytes(reader.ReadInt32());
    if (name.Contains("private", StringComparison.OrdinalIgnoreCase) || name.Contains("Licensing", StringComparison.OrdinalIgnoreCase)
        || name.Contains("public-key", StringComparison.OrdinalIgnoreCase) || name.Contains("LicenseGenerator", StringComparison.OrdinalIgnoreCase))
        throw new Exception("Prohibited resource " + name);
}
if (resourceNames.Count(n => n.StartsWith("MonteCarlo.Excel.Icons.", StringComparison.Ordinal) && n.EndsWith(".png")) != 12) throw new Exception("Ribbon icons missing.");
var refs = md.AssemblyReferences.Select(h => md.GetString(md.GetAssemblyReference(h).Name)).ToArray();
if (refs.Any(n => n.Contains("LicenseGenerator") || n == "Microsoft.Office.Interop.Excel" || n == "office")) throw new Exception("Developer/PIA dependency in runtime.");
Console.WriteLine("PASS: activation-free runtime/About Telivu/12 icons/assembly version/no generator or Office PIA dependencies.");
