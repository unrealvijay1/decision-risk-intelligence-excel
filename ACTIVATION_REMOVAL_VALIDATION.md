# Activation removal validation

Reconciled 2026-10-06. The original implementation validation below records the previously preserved download; see the current release publication status immediately below. Application: https://github.com/unrealvijay1/decision-risk-intelligence-excel (develop). Official website: https://unrealvijay1.github.io/telivu/ . Website/release repository: https://github.com/unrealvijay1/telivu (main).

## Current activation-free release publication

Owner authorized the website download transition. Published `v0.2.1-activation-free.1` in the Telivu release repository with the exact previously validated installer, public filename `TelivuForExcel-Setup-0.2.1.exe`; assembly version remains 0.2.1. GitHub asset metadata and a complete browser-button download confirm 117,421,206 bytes and SHA-256 `d296fa49e7244483d3a2446e129bd036bff241cbe366830f5f2aab2d380fc048`. Historical v0.2.1 is unchanged. Website commit `3105e2a` points all download buttons to the new release and updates the activation disclosure. Static validation passes 11 pages/287 references; browser regression passes all six widths and both root/project routes with zero JavaScript errors/overflow. Legal open-source license approval, signing, clean Windows VM and real x86 Excel acceptance remain pending. The release is explicitly an unsigned prerelease.

## Implementation

Removed the activation architecture and all five analysis access boundaries; no always-licensed bypass remains. Simulation, Correlations, Scenario Analysis, Optimizer and SPC are unrestricted. Assumptions/distributions/fitting, forecasts, settings, target analysis, sensitivity, export and workbook persistence retain their technical validation. Telivu ribbon keeps 12 functional commands and replaces licensing with Help → About Telivu. The About dialog reports assembly version 0.2.1, no subscription/activation and pending legal license approval, with real website/source/issues links and no donation button.

The generator is removed from source/solution/build; the verification key and activation forms/state code are removed. Old user files/Registry and ignored key backups are untouched. Installer keeps AppId, file names, registration, architecture/runtime detection and data preservation; its guide has no activation step.

Some required installer/startup/diagnostic infrastructure was already present but uncommitted at task start. It is included in the source commit so its tested packaging/validation is reproducible; existing working logic is preserved. The unrelated untracked application-repository Pages workflow remains local; the official website deployment uses the existing separate telivu main workflow.

## Validation results

| Check | Result and evidence |
| --- | --- |
| Full Debug/Release solution builds | PASS; both final configurations build/pack successfully. Existing CS8600/CS8603 warnings remain; removed licensing CS0162 disappears |
| Automated unit suite | PASS, **908** in each configuration, zero failed/skipped |
| Tests added/retained | Four ActivationFreeContractTests protect runtime/source state access, 12 functional ribbon commands/About, solution/resources and installer invariants. Five existing explicit-export/packaging contracts retained |
| Tests removed | 23 standalone RSA/DPAPI algorithm-only fixture checks removed with obsolete Installer/LicensingChecks. They were separate from the 904-unit baseline; unit count increases by four |
| Tests rewritten | Packed checker rejects activation namespace/key/resources instead of requiring them; startup opens About and checks absence of state creation; scenario fixture ends its dedicated STA before exit assertion; legacy-data installer sentinels remain |
| Native layouts | PASS full Debug/Release harness at 100%, 125%, 150%, including new About cases, 14 distributions, results, scenarios, correlations, optimizer/constraints, SPC and progress layouts |
| x86/x64 packing | PASS both Debug/Release: PE architecture, explicit exports, embedded/sidecar config and final assembly hashes. Release installer checker also confirms no activation types/resources and 12 icons |
| Installer | PASS isolated build, signed/hash-pinned runtime cache, both runtime-detection probes and 34 lifecycle assertions; source registration/other add-ins/legacy state preserved |
| Local installer artifact | dist/MonteCarloForExcel-Setup-0.2.1.exe; 117421206 bytes, SHA-256 d296fa49e7244483d3a2446e129bd036bff241cbe366830f5f2aab2d380fc048; unsigned; never uploaded |
| Actual Excel smoke tests | PASS production packed Debug x64 and installer-deployed Release x64: Excel-DNA loads, eight UDFs/legacy command, ribbon connects, all analysis dialogs/About/Results, simulation/restoration and disk save/reopen |
| Fresh activation environment | PASS empty process-local LOCALAPPDATA input, absence of all packed activation types/key, no activation folder created and actual legacy Registry snapshots unchanged. This is not a new Windows account/VM; Windows-known-folder diagnostics remain normal |
| Major workflow integration | PASS existing private Excel distributions, scenarios, optimizer and SPC fixtures: source preservation, correlations/replay, forecast sensitivities, reports/export, explicit optimizer Apply, persistence/save-reopen, cancellation/failure and owned process exit. All 14 distribution/fitting maths are covered by unit/UI checks; eight appended distributions additionally disk-tested |
| Fixture cleanup | Startup COM/DNA fixtures required their documented PID/start-time-guarded cleanup after Quit retained references. No customer processes operated on. Scenario harness now exits naturally after dedicated STA teardown; no production COM semantics changed |
| Website | PASS 11 pages, 299 local references, desktop/laptop/tablet/mobile (1920/1440/1366/768/390/375), keyboard nav/FAQ, reduced motion, image fallback; zero JS errors/overflow |
| Placeholders/paid marketing | PASS no TODO/TBD/coming-soon/dummy buttons, pricing/Enterprise pages or fake donations in published HTML. Old installer disclosure is intentionally retained per owner's explicit instruction |
| Existing Download Flow | **PASS** actual hero-button click → original GitHub v0.2.1 URL → normal GitHub release-assets redirect → complete existing installer transfer |
| Download Destination Verified | **YES**, 117427571 bytes, SHA-256 7ea29b75f44f4f64181a4c1c062ef6d575eddee004512d6a97194a48d1cf480d. site/js/site.js and site/site-config.js are byte-identical to the previously deployed versions |
| Dependency audit | OPEN_SOURCE_LICENSE_AUDIT.md inventories all 17 restored NuGet dependencies, platform/compiler/runtime/artwork and required notices. Apache-2.0 is not applied: Office interop terms and artwork/contributor rights remain unresolved |
| Donation provider | **Donation provider setup required**. No payment URL or button exists |
| Git diff --check | PASS |

## Remaining owner decisions and manual acceptance

Approve the legal license after Office interop/artwork ownership is resolved. Activation-free public release publication is now authorized and completed as described above. Historical v0.2.1 retains its original trial/activation asset; release history was preserved. Clean Windows VM installation/normal persistent startup on both actual Excel architectures, real x86 Excel, per-monitor DPI, signing and arbitrary customer models remain manual acceptance gates. The website now selects the activation-free installer; current source/local installer transition is implemented and validated.

## Files removed

- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/ActivationForm.cs
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseForm.cs
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseInfo.cs
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseService.cs
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseStorage.cs
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LocalLicenseValidator.cs
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/TrialService.cs
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/ActivationForm.resx
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseForm.resx
- MonteCarlo.Excel/MonteCarlo.Excel/Licensing/public-key.pem
- MonteCarlo.LicenseGenerator/Program.cs
- MonteCarlo.LicenseGenerator/MonteCarlo.LicenseGenerator.csproj
- Installer/LicensingChecks/Program.cs
- Installer/LicensingChecks/LicensingChecks.csproj
- Installer/LicensingChecks/Prepare-Fixture.ps1

## Publication

Application implementation commit `2a29450` pushed to develop; website commit `d9403f7` pushed to main. GitHub Pages run `37358337618` completed successfully. All 11 published pages responded successfully with the retained current-installer disclosure. Published download JavaScript matches after Git line-ending normalization; configuration differs only by the existing build-time siteUrl substitution. Download URL and handler are unchanged.

## Source commit inventory (2a29450; 79 files)

```text
M	.gitignore
A	ACTIVATION_REMOVAL_VALIDATION.md
A	CONTRIBUTING.md
A	Directory.Build.props
A	Installer/Checks/Checks.csproj
A	Installer/Checks/Program.cs
A	Installer/CustomerGuide.txt
M	Installer/MonteCarloForExcel.iss
A	Installer/README.md
A	Installer/Test-Installer.ps1
A	Installer/Test-ProductionBlock.ps1
A	Installer/Test-RuntimeDetection.ps1
A	Installer/Verify-Payload.ps1
A	Installer/runtime-manifest.json
A	MonteCarlo.Core.Tests/ActivationFreeContractTests.cs
A	MonteCarlo.Core.Tests/ExcelDnaExportContractTests.cs
A	MonteCarlo.Excel.StartupChecks/MonteCarlo.Excel.StartupChecks.csproj
A	MonteCarlo.Excel.StartupChecks/README.md
A	MonteCarlo.Excel.StartupChecks/StartupChecks.cs
A	MonteCarlo.Excel.StartupChecks/Verify-Packages.ps1
A	MonteCarlo.Excel.StartupChecks/Verify-Startup.ps1
M	MonteCarlo.Excel/MonteCarlo.Excel.slnx
A	MonteCarlo.Excel/MonteCarlo.Excel/AboutTelivuForm.cs
A	MonteCarlo.Excel/MonteCarlo.Excel/AddInDiagnostics.cs
A	MonteCarlo.Excel/MonteCarlo.Excel/App.config
M	MonteCarlo.Excel/MonteCarlo.Excel/ExcelModelSimulator.cs
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/ActivationForm.cs
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/ActivationForm.resx
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseForm.cs
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseForm.resx
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseInfo.cs
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseService.cs
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LicenseStorage.cs
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/LocalLicenseValidator.cs
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/TrialService.cs
D	MonteCarlo.Excel/MonteCarlo.Excel/Licensing/public-key.pem
M	MonteCarlo.Excel/MonteCarlo.Excel/MonteCarlo.Excel.csproj
M	MonteCarlo.Excel/MonteCarlo.Excel/MonteCarloAddIn.cs
M	MonteCarlo.Excel/MonteCarlo.Excel/MonteCarloRibbon.cs
M	MonteCarlo.Excel/MonteCarlo.Excel/OptimizerCommand.cs
M	MonteCarlo.Excel/MonteCarlo.Excel/ScenarioAnalysisCommand.cs
M	MonteCarlo.Excel/MonteCarlo.Excel/SpcAnalysisCommand.cs
D	MonteCarlo.LicenseGenerator/MonteCarlo.LicenseGenerator.csproj
D	MonteCarlo.LicenseGenerator/Program.cs
A	MonteCarlo.Results.LayoutChecks/AboutTelivuLayoutChecks.cs
M	MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj
M	MonteCarlo.Results.LayoutChecks/Program.cs
M	MonteCarlo.Results.LayoutChecks/ScenarioLiveChecks.cs
A	OPEN_SOURCE_LICENSE_AUDIT.md
M	PROJECT_CONTEXT.md
A	README.md
A	build-installer.ps1
A	site/.gitignore
A	site/.nojekyll
A	site/CONTRIBUTING.md
A	site/README.md
A	site/assets/brand/mark.svg
A	site/assets/brand/social.png
A	site/assets/product/optimizer.webp
A	site/assets/product/results.webp
A	site/css/site.css
A	site/index.html
A	site/js/site.js
A	site/license/index.html
A	site/privacy/index.html
A	site/resources/demo/index.html
A	site/resources/documentation/index.html
A	site/resources/example-models/index.html
A	site/resources/getting-started/index.html
A	site/resources/index.html
A	site/resources/release-notes/index.html
A	site/resources/tutorials/index.html
A	site/robots.txt
A	site/scripts/browser-check.mjs
A	site/scripts/build.py
A	site/scripts/check.py
A	site/scripts/create-pages.py
A	site/site-config.js
A	site/support/index.html
```
