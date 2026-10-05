param(
    [ValidateSet('Debug','Release')] [string]$Configuration = 'Release',
    [switch]$VerboseDiagnostics,
    [string]$ProductionPath,
    [switch]$UseInstalledCopy,
    [switch]$FreshProfile
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot
$production = if ($ProductionPath) { $ProductionPath } else { Join-Path $repo "artifacts/diagnostics-$($Configuration.ToLower())/MonteCarlo.Excel-AddIn64-packed.xll" }
$fixture = Join-Path $repo 'artifacts/diagnostics-fixture/MonteCarlo.Excel.StartupChecks-AddIn64-packed.xll'
$runName = 'diagnostics-live-' + $Configuration.ToLower() + $(if ($VerboseDiagnostics) { '-verbose' } else { '' }) + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
$directory = Join-Path $repo ('artifacts/' + $runName)
[void][IO.Directory]::CreateDirectory($directory)
$source = if ($UseInstalledCopy) { $production } else { Join-Path $directory 'MonteCarlo.Excel-AddIn64-packed.xll' }
$probe = Join-Path $directory 'MonteCarlo.Excel.StartupChecks-AddIn64-packed.xll'
if (-not $UseInstalledCopy) {
    Copy-Item -LiteralPath $production -Destination $source
    Copy-Item -LiteralPath ($production + '.config') -Destination ($source + '.config')
}
Copy-Item -LiteralPath $fixture -Destination $probe
if ($VerboseDiagnostics) {
    if ($UseInstalledCopy) { throw 'Verbose test must not overwrite installed logging preferences.' }
    [IO.File]::WriteAllText(($source + '.config'), '<configuration><monteCarloDiagnostics verbose="true" /></configuration>')
}
$report = Join-Path $directory 'startup-report.txt'
if (Test-Path -LiteralPath $report) { Remove-Item -LiteralPath $report }
Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class StartupWindows {
 public delegate bool WindowCallback(IntPtr window, IntPtr parameter);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
 [DllImport("user32.dll")] static extern bool EnumWindows(WindowCallback callback, IntPtr parameter);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window, StringBuilder text, int length);
 public static bool HasDiagnosticWindow(uint process) {
  bool found=false;
  EnumWindows((window, parameter) => { uint pid; GetWindowThreadProcessId(window,out pid);
   var text=new StringBuilder(512); GetWindowText(window,text,512);
   if(pid==process && text.ToString().Contains("Diagnostic Display")) found=true;
   return true;
  },IntPtr.Zero);
  return found;
 }
}
'@
$existing = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$ownedCom = [Collections.Generic.List[object]]::new()
$app = $null
$ownedPid = 0
$ownedStart = $null
try {
    $app = New-Object -ComObject Excel.Application
    [uint32]$excelPid = 0
    [void][StartupWindows]::GetWindowThreadProcessId([IntPtr]$app.Hwnd, [ref]$excelPid)
    if ($existing -contains [int]$excelPid) { throw 'Refusing to operate on an existing Excel process.' }
    $ownedPid = [int]$excelPid
    $ownedStart = (Get-Process -Id $ownedPid).StartTime
    [pscustomobject]@{pid=$ownedPid;start=(Get-Process -Id $ownedPid).StartTime;existing=$existing;production=$source} |
        ConvertTo-Json | Set-Content (Join-Path $directory 'owned-excel.json')
    $ownedCom.Add($app)
    $app.Visible = $false; $app.DisplayAlerts = $false
    $books = $app.Workbooks; $ownedCom.Add($books)
    $book = $books.Add(); $ownedCom.Add($book)
    if ($FreshProfile) {
        if (-not $app.RegisterXLL($probe)) { throw 'Developer verification XLL registration failed.' }
        $app.Run('MC_VERIFY_FRESH_PROFILE')
    }
    if (-not $app.RegisterXLL($source)) { throw 'Production XLL registration failed.' }
    if ([StartupWindows]::HasDiagnosticWindow($excelPid)) { throw 'Diagnostic Display appeared during clean production XLL startup.' }
    Write-Output "PASS $Configuration clean Excel startup: no Diagnostic Display"
    if (-not $FreshProfile -and -not $app.RegisterXLL($probe)) { throw 'Developer verification XLL registration failed.' }
    $app.Run('MC_VERIFY_STARTUP')
    if ([StartupWindows]::HasDiagnosticWindow($excelPid)) { throw 'Diagnostic Display appeared during verification/error logging.' }
    $content = [IO.File]::ReadAllText($report)
    Write-Output $content
    if ($content -notmatch '(?m)^PASS\r?$') { throw 'In-Excel verification failed.' }
    # Diagnostics deliberately uses Windows' known-folder API, not the process
    # LOCALAPPDATA override used to exercise an empty activation-state environment.
    $logDirectory = Join-Path $env:LOCALAPPDATA 'MonteCarlo/Logs'
    $log = Get-ChildItem -LiteralPath $logDirectory -Filter "ExcelDna-*-$ownedPid.log" | Select-Object -Last 1
    if (-not $log) { throw 'Production diagnostic log not found.' }
    $stream = [IO.FileStream]::new($log.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $reader = [IO.StreamReader]::new($stream)
    try { $logText = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ($logText -match 'could not be processed|No objects loaded') { throw 'Underlying initialization/registration warning remains.' }
    if ($logText -notmatch 'synthetic initialization error') { throw 'Genuine error-level logging was not captured.' }
    if ($VerboseDiagnostics -and $logText -notmatch 'Processing type') { throw 'Verbose configuration did not capture registration details.' }
    [IO.File]::WriteAllText((Join-Path $directory 'excel-dna.log'), $logText)
    [IO.File]::WriteAllText((Join-Path $repo ('artifacts/diagnostics-live-' + $Configuration.ToLower() + $(if ($VerboseDiagnostics) { '-verbose' } else { '' }) + '-latest.txt')), $directory)
    Write-Output "PASS durable diagnostics, no SimulationSettings registration warning; verbose=$VerboseDiagnostics"
} finally {
    if ($ownedPid -ne 0) {
        try {
            $books.Close()
            try { [void]$app.Run('MC_VERIFY_COM_CLEANUP') } catch { Write-Warning "Fixture COM cleanup unavailable: $_" }
            $app.Quit()
        } finally {
            for ($index = $ownedCom.Count - 1; $index -ge 0; $index--) {
                if ([Runtime.InteropServices.Marshal]::IsComObject($ownedCom[$index])) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ownedCom[$index]) }
            }
            $ownedCom.Clear()
            $app = $null; $book = $null; $books = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
        }
        $process = Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($process -and -not $process.WaitForExit(10000)) {
            # Some Excel-DNA/COM verification contexts retain RCWs after Quit. This
            # fallback is exclusive to the new, workbook-closed fixture process.
            $process = Get-Process -Id $ownedPid -ErrorAction Stop
            if (($existing -contains $ownedPid) -or ($process.StartTime -ne $ownedStart)) {
                throw 'Refusing cleanup: private Excel process identity changed.'
            }
            Stop-Process -Id $ownedPid -Force
            [IO.File]::WriteAllText((Join-Path $directory 'cleanup.txt'), 'COM Quit retained this newly created private fixture process; owned PID/start identity checked before forced cleanup. No customer process was operated on.')
            Write-Output 'PASS private fixture cleanup after COM Quit retained references; customer Excel instances untouched'
        } else { Write-Output 'PASS owned Excel process exited; customer Excel instances untouched' }
    } elseif ($app) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) }
}
