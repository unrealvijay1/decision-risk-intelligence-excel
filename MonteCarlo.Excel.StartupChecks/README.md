# Excel-DNA startup verification

This developer-only fixture is never included in the customer add-in or installer.
It loads after the real packed production XLL in a private Excel instance and executes
checks on Excel's main thread. It reproduces the old implicit-export signature exception
without registering the offending helper, verifies all eight worksheet functions and the
legacy command calculator, opens five real ribbon dialogs plus Results, runs/restores simulation,
saves/reopens an `.xlsx`, and emits a synthetic error to verify durable quiet logging.
The runner rejects pre-existing Excel PIDs and quits only its owned instance.
Explicit production XLL unregister and cached application RCW release precede Quit.
Some verification contexts still retain their private process. The runner records PID/start
identity and all pre-existing PIDs, and may terminate only that newly created, workbook-closed
fixture process after Quit times out. `cleanup.txt` records this fallback; it does not
establish graceful COM shutdown. No customer Excel process is operated on.

To verify an installer-deployed XLL directly, pass `-ProductionPath <installed XLL>` and
`-UseInstalledCopy`. That leaves its logging configuration intact. Alternatively run
`./build-installer.ps1 -LiveExcelChecks`; installer fixtures use isolated AppId/HKCU paths.

Build the solution to isolated `artifacts/diagnostics-debug` and
`artifacts/diagnostics-release` publish directories, then build this project:

```powershell
dotnet build MonteCarlo.Excel/MonteCarlo.Excel.slnx -c Debug -p:ExcelDnaPublishPath=C:\montecarlo\artifacts\diagnostics-debug
dotnet build MonteCarlo.Excel/MonteCarlo.Excel.slnx -c Release -p:ExcelDnaPublishPath=C:\montecarlo\artifacts\diagnostics-release
dotnet build MonteCarlo.Excel.StartupChecks/MonteCarlo.Excel.StartupChecks.csproj -c Release -p:ExcelDnaPublishPath=C:\montecarlo\artifacts\diagnostics-fixture
pwsh -STA -File MonteCarlo.Excel.StartupChecks/Verify-Startup.ps1 -Configuration Release
pwsh -STA -File MonteCarlo.Excel.StartupChecks/Verify-Startup.ps1 -Configuration Debug
pwsh -STA -File MonteCarlo.Excel.StartupChecks/Verify-Startup.ps1 -Configuration Release -VerboseDiagnostics
pwsh -File MonteCarlo.Excel.StartupChecks/Verify-Packages.ps1
```

Reports, synthetic workbooks and copied diagnostic logs are under
`artifacts/diagnostics-live-*`. The runner uses `RegisterXLL` to exercise the genuine
loading/registration path; it does not change the customer's persistent AddIns/OPEN settings.
Installed 32-bit Excel and actual monitor DPI still require separate manual checks.

Production logging is **Warning** for the add-in's Excel-DNA TraceSource, with its
popup listener removed. File logs include errors from existing application Trace calls
and use `%LOCALAPPDATA%\MonteCarlo\Logs\ExcelDna-YYYYMMDD-PID.log` (UTC date).
The default debugger listener remains. Diagnostic I/O failure falls back to Trace/debugger
logging without failing application initialization. Logs have no automatic retention policy.

For verbose file logging, set `verbose="true"` in the shipped `<XLL filename>.config`:

```xml
<configuration>
  <monteCarloDiagnostics verbose="true" />
</configuration>
```

Alternatively launch Excel with `MONTECARLO_DIAGNOSTICS_VERBOSE=1` in its environment.
Restart Excel after changing the setting. Verbose mode never enables the popup.
This XML is read by **AddInDiagnostics**, not .NET's `system.diagnostics` machinery,
which Excel-DNA does not support on .NET 6+. The package embeds a copy and ships matching
x86/x64 `.xll.config` sidecars; missing sidecars retain quiet production defaults.
The installer renames the sidecar with the installed XLL.

The bootstrap uses the pinned Excel-DNA 1.9 internal TraceSource accessor, before AutoOpen;
repeat these checks when upgrading Excel-DNA. Native/runtime failures before the application
assembly executes remain outside this managed logging boundary and require loader debugging.
For those, Excel-DNA's documented environment settings can be set in a dedicated developer
launch process. Do not change shared Excel-DNA registry settings for production users.

## Activation-free startup

Use -FreshProfile with Verify-Startup.ps1. The developer fixture prepares an empty process-local LOCALAPPDATA before loading the production XLL, rejects any production activation types/key resources, opens About Telivu, runs analysis/save-reopen, checks no activation folder was created and verifies legacy Registry values were not changed. Real user activation files/Registry are never cleared. This is a private-process/profile test, not a new Windows-user or clean-VM certification.
