# Production installer

The existing Inno Setup installer retains AppId `{9B4D8E57-7B33-4C68-9D1D-91D5E4D8F001}`
and per-user installation under LocalAppData/Programs. This preserves existing upgrades,
HKCU registration ownership without requiring elevation for the add-in.
Only a missing Microsoft Desktop Runtime prerequisite requires administrator approval.

## Activation-free product

The add-in has no license key, account, trial, expiry, activation service or access-tier architecture. Help → About Telivu replaces the former licensing control and shows the actual assembly version. No support-payment button exists without a real provider. The software license remains an independent owner decision; see OPEN_SOURCE_LICENSE_AUDIT.md.

Old activation files and Registry values are ignored and deliberately left untouched by startup and installation. The installer retains its stable AppId, per-user path and OPEN registration ownership. Workbook data and logging storage are independent and unchanged.

## Build and release gates

Use `./build-installer.ps1` from the repository root; add `-LiveExcelChecks` to verify the
installer-deployed x64 XLL in a private Excel instance. See parameter help and generated
release manifest/report. `Directory.Build.props` is the authoritative product version.
Every build uses a new isolated root, cleans it, builds Release, runs the complete suite,
extracts/verifies both packed XLLs and the absence of activation types/resources, checks a strict payload allowlist,
validates bundled Microsoft runtime SHA-512/Authenticode signatures, then compiles Inno Setup.
It stops on failure. Inno Setup 7.1 (the verified compiler) and the .NET 10 SDK are developer
build requirements, never customer requirements. The customer only receives the Setup EXE.

Microsoft Desktop Runtime 10.0.12 x86/x64 installers are pinned in runtime-manifest.json
against Microsoft's release metadata. Builds can download them or use the verified cache.
Update the manifest and installer minimum runtime together for future security releases.
Both x86 and x64 `InstalledVersions/<architecture>/sharedfx/Microsoft.WindowsDesktop.App`
keys are read from the **32-bit HKLM registry view**. Their architecture qualifier, rather
than the registry view, distinguishes the runtime. Read-only original-Inno-function probes
in `Test-RuntimeDetection.ps1` verify both actual registrations during every release build.
The build does not install these runtimes on the developer machine. The installed Inno 7
compiler reports non-commercial use; upstream requests commercial-user support but does not strictly require purchase. See the dependency audit for the actual terms. A different compatible compiler can be supplied with `-IsccPath` and verified.

Code-signing hooks verify binary publisher identity. A supplied signing script receives
`-Stage Assemblies|Xll|Installer -Paths <files>`; it must use your real code-signing identity
and fail on error. Assemblies are signed before repacking; XLLs and Setup are signed afterward.
No fake certificate is generated. Unsigned builds are clearly identified; signing and
clean-machine acceptance remain required release decisions. A signed uninstaller can later
be configured with Inno Setup's SignTool/SignedUninstaller options.

## Architecture and registration

Setup discovers EXCEL.EXE through App Paths and Office install roots in both registry views,
then reads the executable's PE machine header. Windows bitness never chooses the payload.
Conflicting/missing paths require an explicit EXCEL.EXE selection (or `/EXCELPATH=...` in
silent deployment); native ARM64 Excel is rejected because only x86/x64 XLLs are built.
The selected Excel file version determines the HKCU Office/<major>.0/Excel/Options key.
Only the matching XLL/config is installed, using the original `MonteCarloForExcel.xll` name.

Registration retains the existing `/R "<XLL>"` OPEN/OPENn approach. Setup finds an owned
entry or a free slot, checks writes, removes only exact owned duplicates/previous paths,
and records the slot/command/options key. Uninstall checks current data matches the recorded
installed XLL command before deleting; other add-ins are preserved. Excel-running checks
block install/update/uninstall; setup does not force close Excel or request file replacement
after a reboot. Runtime prerequisites are installed only if the selected architecture lacks
the pinned Desktop Runtime patch or newer in the same major/minor line.

Registration runs in PrepareToInstall so a failed write fails Setup rather than merely
warning after installation. Snapshots restore previous owned entries/metadata on write
failure or cancelled installation. Legacy InstallLocation is captured before Inno updates
it. A changed install path carries logging preferences and removes the obsolete owned XLL.
No licensing folder or parent licensing registry key is deleted.
`Test-ProductionBlock.ps1` probes the real Setup while Excel is already running; it checks
nonzero exit, the close-Excel message, identical registration/uninstall snapshots and no
installed XLL. It never operates on those Excel processes. The production 0.2.0 probe passed.

The suite includes four activation-free source/ribbon/build contracts. The 23 obsolete RSA/DPAPI algorithm fixture checks are removed; packed runtime checks now reject activation types/keys.
The 34 installer assertions exercise actual fixture PE bitness and installed file hashes,
fresh/reinstall/legacy migration/duplicates/other add-ins/uninstall/replaced slots, missing
Excel/prerequisite failures and injected registration-write rollback. AppId/registry paths
are isolated; WMI/runtime-install behavior is mocked there. License/trial sentinels establish
installer non-interference, not real paid-code activation through the production key.
Live installed x64 checks cover fresh-profile startup, UDFs, ribbon dialogs including About Telivu/Results, simulation,
save/reopen and quiet/file diagnostics. The COM harness can retain its private Excel after
Quit; only the captured new PID/start time is eligible for cleanup, recorded in cleanup.txt.
No customer Excel is closed. This harness fallback does not prove graceful COM shutdown.

## Clean-machine acceptance checklist

Use a disposable Windows 10 22H2+/11 VM with desktop Excel and no developer tools.
Test both Excel architectures, including x86 Excel on x64 Windows, missing runtime/elevation,
no Excel, conflicting detection and cancelled selection. Verify fresh install, ribbon,
About Telivu, no activation prompts/state creation, all analysis dialogs, simulation/Results, real model save/reopen and production logs/no diagnostic popup.
Upgrade an activated 0.1.0 installation, reinstall the same version, uninstall (no stale
loading/missing-XLL warning), then reinstall and confirm analysis works without using any old activation state, with config retained.
Verify a running Excel workbook blocks all mutations and other OPEN registrations survive.
Check silent deployment failures and unsigned/signed installer UX. Record each actual result;
automated fixture tests do not certify a clean machine, GUI startup or live x86 Excel.
