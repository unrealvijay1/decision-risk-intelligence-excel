namespace MonteCarlo.Excel;

/// <summary>Owns only a newly created temporary directory, never a user's workbook path.</summary>
public sealed class ScenarioTemporaryFile : IDisposable
{
    private readonly string directory;
    public string Path { get; }
    private bool disposed;
    private ScenarioTemporaryFile(string name)
    {
        directory = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MonteCarloScenario-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        Path = System.IO.Path.Combine(directory, System.IO.Path.GetFileName(name));
    }
    public static ScenarioTemporaryFile Create(string name, Action<string> saveCopy)
    {
        if (string.IsNullOrWhiteSpace(name) || System.IO.Path.GetFileName(name) != name)
            throw new ArgumentException("Invalid temporary workbook name.");
        var file = new ScenarioTemporaryFile(name);
        try { saveCopy(file.Path); return file; }
        catch { file.Dispose(); throw; }
    }
    public void Dispose()
    {
        if (disposed) return;
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                File.Delete(Path);
                Directory.Delete(directory); // Never recursively delete an unknown directory.
                disposed = true;
                return;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                if (attempt == 2) throw new IOException($"Could not remove the temporary scenario workbook: {Path}", ex);
                Thread.Sleep(100);
            }
        }
    }
}
