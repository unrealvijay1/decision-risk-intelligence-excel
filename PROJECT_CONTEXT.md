# Monte Carlo for Excel --- Project Context

**Purpose:** Permanent handover/context document for continuing
development of the Monte Carlo for Excel product across future coding
sessions.

**Last updated:** 26 September 2026\
**Current installer version:** 0.1.0\
**Primary platform:** Windows / Microsoft Excel 64-bit\
**Technology:** C# / .NET 10 / Excel-DNA

------------------------------------------------------------------------

## 1. Product Overview

Monte Carlo for Excel is an Excel add-in for performing Monte Carlo
simulation directly inside Microsoft Excel.

The goal is to evolve it into a commercially distributable Excel product
with a simple workflow:

1.  Define model assumptions in Excel.
2.  Define forecast/output cells.
3.  Run Monte Carlo simulation.
4.  Review simulation results.
5.  Export/report results.
6.  Use a 30-day trial initially.
7.  Activate a paid Professional or Enterprise license using an offline
    signed license code.

The product is currently being developed locally in Visual Studio and is
moving toward packaged commercial distribution.

------------------------------------------------------------------------

## 2. Solution / Repository Structure

The Git repository root is intended to be:

`C:\montecarlo\`

Important projects/components include:

-   **MonteCarlo.Core** --- core simulation/domain logic.
-   **MonteCarlo.Excel** --- Excel-DNA add-in, Ribbon, UI, Excel
    integration and licensing.
-   **MonteCarlo.Core.Tests** --- platform-independent automated tests.
-   **MonteCarlo.LicenseGenerator** --- private utility used by the
    developer to generate signed customer licenses.
-   **Installer** --- Inno Setup installer scripts and generated
    installer output.

The LicenseGenerator must remain independent from the Excel add-in. The
Excel project must never reference or distribute the LicenseGenerator's
private signing key.

------------------------------------------------------------------------

## 3. Current Product Capabilities

The product has already progressed beyond a proof of concept.
Current/implemented areas include:

-   Excel-DNA based Excel add-in.
-   Custom **Monte Carlo** Ribbon.
-   32×32 Ribbon images/icons.
-   Define Assumption workflow.
-   Define Forecast workflow.
-   Monte Carlo simulation execution.
-   Simulation results/dashboard.
-   Forecast target marker and P50/P80 value labels with overlap handling.
-   Target success/failure visualization with explicit success-direction
    selection and empirical success probability (see section 26).
-   Report exporter.
-   Distribution-related functionality/framework.
-   Licensing UI.
-   30-day local trial.
-   Offline signed customer licenses.
-   Professional and Enterprise license types.
-   Packed Excel-DNA XLL release artifact.
-   Initial Inno Setup installer.

Correlation between assumptions is intentionally deferred for a later
development milestone.

------------------------------------------------------------------------

## 4. Ribbon / Excel UI

The add-in has a dedicated **Monte Carlo** Ribbon tab.

Ribbon commands include functionality such as:

-   Define Assumption
-   Define Forecast
-   Run Simulation
-   Results/reporting functions
-   Licensing

Custom 32×32 images have been added to improve the Ribbon presentation.

When adding future features, preserve the existing Ribbon organization
and avoid overcrowding it. New commands should be grouped according to
the user workflow rather than simply appended.

------------------------------------------------------------------------

## 5. Licensing Architecture

### Current licensing strategy

Licensing is deliberately local/offline for Phase 1.

Normal customer flow:

`Install → 30-day Trial → Purchase → Receive signed license → Activate → Professional/Enterprise`

There is currently no required licensing server or internet activation.

### License types

The licensing model supports:

-   **Development**
-   **Trial**
-   **Professional**
-   **Enterprise**

Development mode exists for development/testing but should be disabled
when testing or distributing the customer experience.

### Main licensing classes

Inside `MonteCarlo.Excel\Licensing`, the important classes include:

-   `LicenseInfo.cs`
-   `LicenseService.cs`
-   `LicenseStorage.cs`
-   `LocalLicenseValidator.cs`
-   `TrialService.cs`
-   `ActivationForm.cs`
-   `LicenseForm.cs`

------------------------------------------------------------------------

## 6. Signed Offline Licenses

Customer licenses use asymmetric RSA signing.

The conceptual format is:

`MC1.<base64url-payload>.<base64url-signature>`

The signed payload contains information such as:

-   License ID
-   Customer name
-   Edition
-   Issued date
-   Expiry date

### Critical security rule

**The private RSA key must NEVER be distributed with the Excel add-in or
committed to Git.**

Architecture:

`LicenseGenerator + PRIVATE KEY → signs customer license`

`Excel Add-in + PUBLIC KEY → verifies customer license`

The add-in can therefore verify a legitimate license offline without
containing the secret needed to generate one.

### Case sensitivity

Signed license codes are **case-sensitive**.

Do not call:

`ToUpperInvariant()`

or use an activation textbox configured with:

`CharacterCasing.Upper`

The activation UI was specifically corrected to use normal character
casing because uppercasing the Base64URL payload/signature invalidates
the RSA signature.

------------------------------------------------------------------------

## 7. License Generator

The private license-generation utility is:

`MonteCarlo.LicenseGenerator`

It generates customer licenses using the private RSA key.

The generator asks for information such as:

-   Customer name
-   Professional or Enterprise edition
-   Expiry date

It generates a unique license ID and signed `MC1...` license code.

Generated customer license files are stored under a local `Licenses`
output directory.

The LicenseGenerator and private signing material are **developer-only
assets** and must never be included in the customer installer.

------------------------------------------------------------------------

## 8. Key Management

The RSA key pair consists of:

-   `private-key.pem`
-   `public-key.pem`

### Private key

`private-key.pem` is the master signing credential.

Rules:

-   Never distribute it.
-   Never place it inside `MonteCarlo.Excel`.
-   Never commit it to Git.
-   Maintain a secure backup outside the repository.
-   Loss of the production private key would prevent issuance of new
    licenses compatible with that signing identity.

### Public key

The public key is safe to distribute.

A copy is embedded in `MonteCarlo.Excel` as an **Embedded Resource** and
is used by `LocalLicenseValidator` to verify customer license
signatures.

------------------------------------------------------------------------

## 9. Git Ignore / Sensitive Files

The repository `.gitignore` is located at:

`C:\montecarlo\.gitignore`

Licensing-related exclusions were added so sensitive/generated data is
not committed.

The intended exclusions include:

``` gitignore
# Monte Carlo licensing
# Never commit signing keys
**/Keys/
**/private-key.pem

# Never commit generated customer licenses
**/Licenses/
```

Before every release or significant licensing change, verify in Git
Changes that no private key or generated customer license is staged.

------------------------------------------------------------------------

## 10. Trial Architecture

The product provides a **30-day trial**.

Trial handling was separated from `LicenseService` into:

`TrialService.cs`

### Trial protection

The trial is not stored only as an obvious plaintext date.

The current implementation uses:

-   Windows DPAPI (`ProtectedData`)
-   Current-user protection
-   Encrypted trial state
-   Local file copy
-   Windows Registry copy
-   Start date
-   Last-run date
-   Trial ID
-   Basic clock rollback detection
-   Conservative reconciliation of duplicate trial state

The local trial state is designed so deleting only the local `trial.dat`
file does not simply restart the trial; the second state copy can
restore it.

### Important limitation

Pure offline trial protection cannot be made impossible to bypass by a
determined user with sufficient access to the machine. The objective is
reasonable commercial protection, not DRM-level security.

### Deactivation behavior

Deactivating a Professional/Enterprise license removes the locally
stored paid license.

It does **not** create a new 30-day trial.

The user falls back to the state of the original trial. If that trial
has already expired, it remains expired.

------------------------------------------------------------------------

## 11. License UI

Two main Windows Forms are used:

### `LicenseForm`

Displays:

-   License type
-   Status
-   Licensed To
-   Expiry
-   Trial days remaining where applicable
-   Activate License button
-   Deactivate button
-   Close button

The UI was adjusted so **Activate License** remains on one line and the
bottom controls are properly spaced.

### `ActivationForm`

Allows the customer to paste the full signed license code.

The textbox was enlarged because RSA-signed licenses are long.

Critical setting:

`CharacterCasing = CharacterCasing.Normal`

Do not change it back to uppercase.

------------------------------------------------------------------------

## 12. LicenseService Responsibilities

`LicenseService` is the high-level licensing coordinator.

Current intended responsibilities:

1.  Development override when explicitly enabled.
2.  Check locally saved signed customer license.
3.  Validate it through `LocalLicenseValidator`.
4.  Return Professional/Enterprise license, including expired state.
5.  If no customer license exists, delegate to `TrialService`.
6.  Activate and persist valid signed licenses.
7.  Expose access checks.
8.  Deactivate locally stored paid licenses.

Trial persistence implementation itself should remain in `TrialService`,
not duplicated inside `LicenseService`.

------------------------------------------------------------------------

## 13. Release Build

Commercial/test packaging should use a **Release** build rather than
Debug.

Current release/publish path:

`C:\montecarlo\MonteCarlo.Excel\MonteCarlo.Excel\bin\Release\net10.0-windows\publish\`

The important Excel-DNA artifact is:

`MonteCarlo.Excel-AddIn64-packed.xll`

The unpacked XLL was tested and failed when moved by itself because it
required the companion `.dna` file. The **packed XLL was then tested and
worked standalone**.

Therefore the packed XLL is the current customer distribution artifact.

------------------------------------------------------------------------

## 14. Excel-DNA Distribution

Current target is primarily **64-bit Excel**.

The packed XLL contains the required Excel-DNA/add-in content and is
intended to avoid distributing a loose collection of DLLs and `.dna`
files.

The installer renames the distributed artifact to a cleaner
customer-facing name:

`MonteCarloForExcel.xll`

Do not distribute:

-   `.pdb` debugging symbols
-   LicenseGenerator
-   private signing keys
-   generated customer licenses
-   unnecessary source/build files

------------------------------------------------------------------------

## 15. Installer

Installer technology:

**Inno Setup 7, 64-bit edition**

Installer source location:

`C:\montecarlo\Installer\MonteCarloForExcel.iss`

Installer output location:

`C:\montecarlo\Installer\Output\`

Current installer version:

`0.1.0`

Expected installer filename:

`MonteCarloForExcel-Setup-0.1.0.exe`

### Current source XLL

The installer currently points to:

`C:\montecarlo\MonteCarlo.Excel\MonteCarlo.Excel\bin\Release\net10.0-windows\publish\MonteCarlo.Excel-AddIn64-packed.xll`

### Installation strategy

Current installer is designed as a per-user installation under Local
AppData rather than requiring Program Files/admin installation.

Conceptually:

`Setup.exe → install packed XLL → register with Excel → user opens Excel → Monte Carlo Ribbon loads`

------------------------------------------------------------------------

## 16. Excel Registration

The installer registers the XLL through Excel's user Options registry
area.

The important registry area currently used is:

`HKCU\Software\Microsoft\Office\16.0\Excel\Options`

Excel may already contain add-in startup values named:

-   `OPEN`
-   `OPEN1`
-   `OPEN2`
-   `OPEN3`
-   etc.

The installer must **not overwrite another add-in's OPEN entry**.

Current installer logic:

1.  Look for an existing registration pointing to our XLL.
2.  If not present, find the first unused `OPEN`, `OPEN1`, `OPEN2`, etc.
3.  Register Monte Carlo there.
4.  Record which OPEN value was used.
5.  During uninstall, verify that the value still points to our XLL
    before deleting it.

This safety behavior should be preserved in future installer revisions.

------------------------------------------------------------------------

## 17. Installer Decisions

The installer currently:

-   Uses a stable `AppId` that should remain unchanged across upgrades
    of the same product.
-   Uses per-user installation.
-   Installs the 64-bit packed XLL.
-   Registers the XLL for Excel startup.
-   Supports uninstall.
-   Does not create a desktop shortcut for the XLL.
-   Should not delete trial/licensing state merely because the
    application is uninstalled.

The customer should interact with the product through Excel, not by
double-clicking a desktop shortcut to an `.xll`.

------------------------------------------------------------------------

## 18. Versioning

Current development installer version:

`0.1.0`

Future releases should increment versions deliberately, for example:

-   0.1.0 --- first installer/testing build
-   0.2.0 --- meaningful feature addition
-   0.2.1 --- bug fix
-   1.0.0 --- first production commercial release

Keep the Inno Setup `AppId` stable across upgrades.

------------------------------------------------------------------------

## 19. Git Workflow

Development is maintained in Git.

Useful milestone commit messages already used/recommended during
development include concepts such as:

-   Add licensing framework, trial mode and activation UI
-   Implement signed offline licensing and customer activation
-   Refactor licensing with signed licenses and protected trial service
-   Complete offline licensing, protected trial and license UI
-   Prepare packed Excel-DNA release build

Continue committing at logical milestones rather than waiting until many
unrelated changes accumulate.

------------------------------------------------------------------------

## 20. Current Development Status

Major completed milestones:

-   [x] Core Monte Carlo Excel add-in
-   [x] Ribbon UI
-   [x] Ribbon icons
-   [x] Assumption definition
-   [x] Forecast definition
-   [x] Simulation execution
-   [x] Results/dashboard
-   [x] Forecast Target Marker (see section 26)
-   [x] P50/P80 marker-label UX enhancement (see section 26)
-   [x] Target Success / Failure Regions (see section 26)
-   [x] Report exporter
-   [x] Licensing framework
-   [x] 30-day trial
-   [x] RSA key pair generation
-   [x] Customer license generator
-   [x] Signed offline license validation
-   [x] Professional/Enterprise activation
-   [x] License UI
-   [x] Protected local trial
-   [x] Release build
-   [x] Packed 64-bit XLL
-   [x] Initial Inno Setup installer compilation

------------------------------------------------------------------------

## 21. Immediate Next Steps

Recommended sequence from the current point:

1.  Test the generated installer end-to-end.
2.  Close Excel before installation.
3.  Install `MonteCarloForExcel-Setup-0.1.0.exe`.
4.  Open Excel normally.
5.  Confirm the Monte Carlo Ribbon loads automatically.
6.  Test Define Assumption.
7.  Test Define Forecast.
8.  Run a simulation.
9.  Test results/report export.
10. Verify trial/license screen.
11. Test signed Professional activation.
12. Close/reopen Excel and verify license persistence.
13. Uninstall from Windows Installed Apps.
14. Reopen Excel and verify the Ribbon is gone.
15. Verify uninstall/reinstall does not reset the original trial.
16. Test upgrade/version handling.
17. Test on a clean Windows machine/VM without Visual Studio.
18. Prepare code signing for commercial distribution.

------------------------------------------------------------------------

## 22. Future Product Roadmap

Potential future work includes:

-   Correlation between assumptions.
-   Additional probability distributions.
-   Improved distribution-selection UX.
-   More simulation diagnostics.
-   Enhanced charts/results.
-   More reporting/export options.
-   Model validation.
-   Sensitivity analysis enhancements.
-   Better error handling and user guidance.
-   Product documentation/help.
-   Expanded automated regression coverage.
-   Installer upgrade handling.
-   32-bit Excel support if commercially required.
-   Code signing.
-   Commercial publisher identity.
-   Optional future online licensing/device activation.
-   License revocation/renewal infrastructure if online licensing is
    later introduced.

Do not implement online licensing merely because it is possible. The
current Phase-1 product intentionally uses offline signed licenses.

------------------------------------------------------------------------

## 23. Rules for Future Coding Sessions

When continuing this project with ChatGPT or another developer:

1.  Read this file first.
2.  Use the **current source files** as the source of truth for
    implementation details.
3.  Do not recreate licensing architecture from scratch.
4.  Never expose or commit the private RSA key.
5.  Preserve signed-license case sensitivity.
6.  Preserve trial state across ordinary uninstall/reinstall.
7.  Use the packed Release XLL for distribution.
8.  Do not overwrite another Excel add-in's `OPEN` registry entry.
9.  Keep `LicenseService`, `LocalLicenseValidator`, and `TrialService`
    responsibilities separated.
10. Update this document whenever an architectural decision materially
    changes.

When asking ChatGPT to add a feature, provide this file plus the
relevant current `.cs` files.

Example request:

> Continue my Monte Carlo for Excel project using PROJECT_CONTEXT.md as
> the project context. I now want to add correlation between
> assumptions. Here are the current simulation and assumption files.
> Preserve existing licensing and installer behavior.

------------------------------------------------------------------------

## 24. Source of Truth

This document is a **context/handover document**, not a substitute for
source control.

Priority when resolving discrepancies:

1.  Current working source code.
2.  Git history.
3.  This `PROJECT_CONTEXT.md`.
4.  Previous chat history.

Update this file after significant architectural changes so it remains
useful as a future handover document.

------------------------------------------------------------------------

## 25. Distribution Preview

The Define/Edit Assumption form previews all six supported assumption
distributions from the unsaved values in its controls. The deterministic
analytical density and parameter validation are implemented in
`MonteCarlo.Core/DistributionPreview.cs`. The form only reads its controls,
displays inline validation, and renders the returned points. Previewing
does not sample, modify Excel cells, or save the simulation model.

Normal and Lognormal previews cover four standard deviations on either
side of the mean in their respective spaces. Bounded distributions use
their support. Beta densities with singular endpoints are evaluated just
inside the support; visual peak clipping is confined to the form renderer.

------------------------------------------------------------------------

## 26. Forecast Target and Success / Failure Regions

### Architecture and relevant files

`ResultsForm` renders the histogram with custom WinForms/GDI+. Core's
`ChartValuePosition` supplies shared range classification/normalization;
`TargetDirection` defines success direction; `EmpiricalProbability` counts
sample comparisons. `ForecastRunResult` delegates probability calculations
to that helper. Simulation generation, statistics, and percentiles are unchanged.

- `MonteCarlo.Excel/MonteCarlo.Excel/`: `ResultsForm.cs`, `SimulationRunResult.cs`.
- `MonteCarlo.Core/`: `ChartValuePosition.cs`, `TargetDirection.cs`, `EmpiricalProbability.cs`.
- `MonteCarlo.Core.Tests/`: `ChartValuePositionTests.cs`, `EmpiricalProbabilityTests.cs`.

### Session state and probability semantics

Target and direction are ResultsForm session state, not persisted in
`ForecastDefinition` or workbook storage. Rendering uses the accepted,
calculated target. Pending edits or missing/invalid/non-finite targets clear
the marker, regions, success percentage, and legend. Forecast refresh resets
the target to P80 rounded to two decimals; this can differ from exact P80.

Direction is explicit, initially unspecified in each new results window,
and retained independently per forecast within that window. Never infer it
from forecast name, worksheet, distribution, target value, or formatting.
Switching forecasts restores direction but retains the target reset above.
Without direction, the marker and existing probabilities still work;
no success claim or shading is shown.

- `AtOrBelow`: success is outcome `<= target`; Miss is `> target`.
- `AtOrAbove`: success is outcome `>= target`; Miss is `< target`.
- `P(<= target)` is inclusive lower-side probability; existing `P(> target)`
  remains strict upper-side probability; `P(>= target)` is inclusive upper-side.
- Never implement `P(>= target)` as `1 - P(<= target)`: that gives `P(> target)`
  and loses equality mass. Probabilities use matching samples / total samples.

### Rendering, edge cases, and limits

The in-range target and percentile markers share the same X-axis transformation.
P50/P80 labels include full culture-aware numeric values. P50, P80, and Target
use separate compact header rows with measured labels constrained to chart bounds.
Ticks associate in-range labels with exact X coordinates; rows/ticks resolve
presentation overlap without shifting coincident or nearby marker positions.
Percentiles remain thin/dashed; Target remains orange/solid. Redundant percentile
X-axis values are omitted; minimum/maximum labels remain.

Subtle plot-background Success/Miss regions split at the exact target coordinate,
including inside a bin. Bins are not recolored; bars, P50/P80, target marker,
and labels render above shading. A text legend identifies both sides.
Shading represents value regions, not probability mass; numerical success
probability comes from simulation samples, never histogram geometry.

- Out-of-range targets retain an annotation, with no target line, clamping,
  or axis expansion. The whole plot represents Success/Miss according to
  direction (0%/100% for finite samples).
- Endpoint equality follows inclusive comparisons even when its visual region
  has zero width.
- Constant outcomes retain the blank histogram; empirical success can still display.
- Future persistence requires an explicit backward-compatible model/workbook change.
- Worksheet currency/date/percentage formatting is not implemented; values use
  the existing culture-aware numeric formatting.
- UI state/layout integration is primarily manually verified, not automated UI-tested.

### Validation baseline

As of September 2026: **88 automated tests passing**, complete solution build successful,
and manual Excel verification completed for percentile/target markers and both success directions.

## 27. Cumulative Probability (S-Curve)

The forecast results chart has a Distribution/Cumulative selector. Distribution
is the initial default and its existing renderer is unchanged. ResultsForm stores
the last selected mode in a static, session-only field, so subsequent results
dialogs retain the selection. Each successful simulation still opens a new modal
dialog with a new SimulationRunResult; there is no in-place rerun workflow.
Switching views only invalidates the chart and never calls the simulation engine.

Core's CumulativeProbability.Build sorts a copy of the supplied outcomes and
groups duplicates into one jump per distinct value. Each point includes the
probability immediately before the jump and the inclusive probability at that
value. ResultsForm draws a right-continuous empirical step curve with a 0–100%
Y-axis, caching the data per ForecastRunResult within the dialog. Empty results
have no curve; any non-finite sample makes the whole curve unavailable with an
inline explanation. Samples are never silently removed from the denominator.

The cumulative chart reuses ForecastRunResult.P50/P80 exactly and shares the
existing marker renderer. It does not calculate percentiles. Interpolated
percentile values may not intersect an empirical step at exactly 50%/80%.
Constant/single outcomes get a padded range. Cumulative plotting includes finite
out-of-range targets in its axis range; Distribution retains its original
out-of-range annotation behavior. Numeric formatting remains unchanged (no new
Excel date/currency-format support).

Target confidence reuses ProbabilityLessThanOrEqual for AtOrBelow and
ProbabilityGreaterThanOrEqual for AtOrAbove. With unspecified direction, the
chart explicitly labels P(X <= target) as cumulative probability, not success
confidence. Existing success/miss regions are reused. When currentTarget is
absent, the cumulative chart shows its curve and percentile markers without a
target line or chart probability label. The existing P80 default-target and
Calculate workflows remain unchanged.

Validation: complete suite passed before edits (88 tests), after the Core stage
(102 tests), and after integration (107 tests; 19 new cases, no failures/skips).
The complete solution build succeeded with three warnings in unchanged files
(LicenseService, AssumptionForm, WorkbookPersistence). Tests link the non-UI
forecast result/model source files to exercise actual percentile/probability
behavior without requiring Excel. Manual Excel UI verification is still required:
switching without a simulation, target/no-target states, both directions,
coincident markers, constant outcomes, out-of-range targets, display scaling,
and a new run reopening in the selected mode with fresh results.

## 28. Confidence to Target

Probability Analysis now contains compact Target → Confidence and Confidence →
Target tabs within its existing footprint. The existing percentile panel and
P50/P80 chart benchmarks remain unchanged. Neither calculation-tab changes nor
chart-view changes invoke a simulation.

ForecastRunResult.GetTargetForConfidence accepts a percentage from 0 through 100
and an optional TargetDirection. It sorts a copy and calls the same private
Percentile method used for P10/P50/P80/P90. AtOrBelow uses C/100; AtOrAbove uses
1 - C/100; unspecified direction uses the lower-tail percentile C/100 without
claiming success confidence. Invalid confidence, non-finite samples, and
unrepresentable interpolation results are rejected; no samples are filtered.
The percentile algorithm and EmpiricalProbability remain unchanged.

ResultsForm tracks manual versus confidence-derived target origin and accepted
requested confidence. A successful reverse calculation updates the shared target,
both probability displays, success regions, and chart. The full target is retained
internally; the editable target textbox uses round-trip numeric formatting, while
result labels use the existing display formatting. Invalid confidence submissions
leave the accepted calculation intact. Tab changes retain accepted state. Direction
changes recompute a confidence-derived target from accepted confidence (not pending
textbox edits), but retain a manually entered target. Failed reverse recalculation
restores the prior direction. Forecast changes clear reverse state and retain the
existing rounded-P80 default-target workflow.

The reverse result distinguishes requested confidence, required target, and actual
sample probability. These may differ for interpolated quantiles, discrete samples,
and duplicates. The S-Curve remains an increasing lower-tail CDF. Its dashed orange
horizontal guide, intersection dot, and percentage label always represent actual
P(X <= target), even when success means X >= target. Inclusive upper-tail success
continues to use ProbabilityGreaterThanOrEqual. No requested-probability guide is
drawn at a misleading curve intersection.

Validation for this enhancement: baseline 107 tests; calculation stage 130 tests;
final suite 139 tests (32 new cases), all passing with no failures/skips. Complete
solution build succeeds with the same three warnings in unchanged files noted
above. Manual Excel verification remains required for both calculation modes,
73%/87.5%, both directions and switching, 0%/100%, invalid input preserving accepted
confidence, guide/marker overlaps, constants, out-of-range targets, forecast changes,
fresh simulations, chart-view session persistence, and Windows display scaling.

## 29. Decision-Focused Results Layout

ResultsForm retains its 1120 x 830 window and existing controls/handlers, arranged
with nested WinForms TableLayoutPanels. The top contains the forecast selector
and decision summary: dominant actual success probability, accepted target, P50,
and P80. The main chart is on the left and Decision tools on the right. Sensitivity
and permanently visible, borderless Details occupy the supporting lower row.
Details contains Trials, Mean, Std Dev, Minimum, Maximum, P10, and P90. The separate
statistics/percentiles presentation is removed; all model calculations remain.

RefreshSuccessDisplay is the shared target-dependent presentation path. It clears
all contextual labels before rebuilding the summary, requested confidence, and
legends from accepted state. Missing target displays 'No target defined' and an
em dash instead of a percentage. An accepted target without direction displays
'Select success direction', with no claim of success. Forecast changes retain
the rounded-P80 workflow, reset reverse-calculation state, and refresh landmarks,
Details, and forecast tooltips. Full target precision and direction-change behavior
are unchanged.

Actual success probability has one dominant summary presentation. The calculation
tabs no longer repeat that result; the reverse tab shows requested confidence and
'Target at ... confidence'. Unspecified direction uses 'cumulative probability'.
The redundant cumulative-chart footer is removed; the actual probability guide
remains labelled '≤ target', including for AtOrAbove success. Existing inclusive
and strict comparison values remain as subdued context below Decision tools.
Tooltips expose long forecast names and summary/detail values when clipped.

Validation: complete baseline and final suites each pass 139 tests, with zero
failures/skips. No new calculation or pixel-coordinate tests were needed. Full
solution build succeeds with the three existing warnings noted above. Runtime
presentation smoke checks cover unspecified direction, target clearing without
stale comparison labels, confidence-created targets, and direction recalculation.
Offscreen rendering did not provide a usable preview; manual Excel visual checks
remain required for hierarchy, both chart/calculation modes, both directions,
missing target, forecast switching, long names/large values, marker overlap,
Details, Sensitivity, Export Report, and Windows 100%/125%/150% scaling.

## 30. Simulation Validation and User-Friendly Errors

The Ribbon entry remains OnRunSimulation. After licensing and the settings dialog,
it calls SimulationService.TryRun. This loads the saved configuration with
WorkbookPersistence.LoadModelForSimulation and invokes SimulationExecution only
after validating it. The existing SimulationService.Run signature remains available
as an exception-based compatibility wrapper. ResultsForm, distributions, percentiles,
target semantics, sensitivity calculations, and result persistence are unchanged.
At this stage targets remained ResultsForm session state. Section 31 records the
subsequent addition of optional workbook target settings and their validation.

ValidationResult holds ValidationError records with location, user message,
suggested correction, and severity. Critical errors block execution before any
sample writes or Excel setting changes. The rules include:

- A positive trial count, at least one assumption, and at least one forecast.
- Required finite numeric distribution parameters; zero remains a valid mean or
  bound and is never used as a substitute for missing required saved parameters.
- Existing DistributionPreview.Validate rules: positive Normal/Lognormal standard
  deviation; ordered Uniform/Beta bounds; ordered Triangular/PERT bounds with mode
  inside them; positive Beta shape parameters; supported distributions only.
- Existing worksheets and single-cell references in the configured workbook/sheet,
  finite numeric cell values, and no Excel error values. Excel's IsError is checked
  on the Range before converting Value2, avoiding confusion between error codes and
  ordinary numeric values. Runtime forecast checks include cell and trial context.
- No duplicate physical assumption cells, protected locked inputs, merged inputs,
  array-formula inputs, or spilled-formula inputs.
- Invalid/unknown saved item types, missing references, and malformed parameters
  are reported rather than silently skipped for simulation. Blank rows do not
  truncate simulation validation. Both existing saved layouts remain readable;
  the older layout's absent Parameter4 remains valid for non-Beta distributions.
  The normal LoadModel and SaveModel schema are retained.

SimulationExecution contains the existing sampling loop behind small
ISimulationWorkbook/ISimulationCell interfaces. ExcelSimulationWorkbook pins the
workbook and resolved cell objects for the run. All originals and settings are
captured before mutation. Ordinary formulas are captured/restored as Formula2
(Formula on older Excel without that property); plain values use Value2. Array
and spill inputs are rejected because changing their single-cell representation
cannot safely preserve the entire range.

ScreenUpdating, EnableEvents, sample writes, calculations, and progress updates
are inside a try/finally cleanup boundary. Cleanup attempts every original cell,
recalculation, and every saved application setting independently, retrying a failed
operation once. A persistent restoration failure prevents success and identifies
the affected cells/settings with instructions to check before saving. No code can
guarantee restoration if Excel closes or permanently refuses writes; such a failure
is explicitly reported rather than hidden. Original exceptions and cleanup errors
are retained on the structured outcome.

The older MC_RUN_PROJECTMODEL command also uses this execution boundary with its
same fixed PERT inputs, output-cell-only recalculation, and existing report/percentile
calculations. Both simulation entry points show safe user messages and send technical
exceptions to System.Diagnostics.Trace. There was no durable logging subsystem;
Trace output requires an attached/configured listener and no new log store is added.

Tests link the non-COM validation/execution files into the existing test project and
use fake workbook/cell adapters to exercise validation gates, formula/value capture,
partial writes, runtime formula errors, setup failures, restoration retries, failed
cell/settings cleanup, and preservation of original settings. Baseline: 139 passing
tests. Final: 200 passing tests, 0 failures/skips (61 new cases).

The full solution and both packed XLLs build in artifacts/validation-build because
Excel has the normal Debug packed XLL locked. Three existing warnings remain in
LicenseService, AssumptionForm, and WorkbookPersistence. Manual Excel verification
is still required for actual COM error marshaling, valid old/new workbooks, missing
worksheets/ranges, protected/merged/array/spill inputs, formula restoration after a
runtime failure, and both simulation entry points. No live workbook was modified
as part of automated validation.


## 31. Workbook Model Persistence Review and Confirmed Gaps

The existing `__MonteCarloConfig` Very Hidden worksheet remains the persistence
store. No replacement serialization format or speculative feature structures were
introduced. Assumptions, forecasts, friendly names, all six distribution types,
and all four distribution parameters were already persisted. Both the legacy
three-parameter/name-in-column-8 layout and the four-parameter/name-in-column-9
layout remain readable. Classifications, user-configured correlations, scenarios,
and decision variables are not current product features; their persistence is
Not Applicable / Future. Sensitivity correlations are calculated results, not
saved correlation settings.

The lifecycle is: Define/Edit commands update the model and call SaveModel;
normal Excel Save writes metadata with the workbook; workbook open/activation
reloads it; SimulationService.TryRun reloads and validates again before simulation.
The add-in now uses the Excel interop event delegate types, logs registration
failures through Trace, and attempts the initial load independently of registration.
Normal restore uses the existing structured validation path and surfaces affected
saved rows with correction guidance instead of silently skipping damaged entries.
Unexpected restoration failures produce a safe message and Trace diagnostics.

Confirmed gaps addressed:

- Optional column 10 (`CellLink`) stores a hidden workbook-scoped name per input or
  output (`_MC_Cell_` plus a GUID). Excel maintains these direct single-cell
  references across worksheet renames and structural moves. Restore resolves the
  tracked cell and refreshes the sheet/address used by simulation. Broken names,
  deleted cells/sheets, invalid ranges, or references to another workbook produce
  validation errors; broken tracked links never fall back to stale coordinates.
  Existing legacy sheet/address references still load. Tracking starts when those
  definitions are next saved; an earlier rename cannot safely be inferred.
- Optional columns 11–14 store TargetConfigured, Target, TargetDirection and
  RequestedConfidence. Null settings retain the existing P80 default; explicit
  empty target settings preserve the no-target state. Accepted manual targets,
  inclusive direction, and requested confidence restore without rounding or
  changing percentile/probability calculations. Requested confidence remains
  separate from actual success probability on a subsequent run. Restored targets
  are fixed saved values; they are not recomputed from a new sample automatically.
- With explicit user approval, Results restores and saves these settings when
  accepted targets/directions change or the target is cleared. Draft numeric text
  is not an accepted target. Updates touch only the matching forecast row and pin
  the source workbook, including after Export Report activates another workbook.
  Forecast rename/redefinition preserves valid target settings and tracked links.
- Saved targets must be finite; direction must be supported; optional requested
  confidence must be finite, 0–100, and accompanied by a target. Malformed saved
  settings are reported and block simulation via the existing validation gate.
- Text metadata columns use text formatting so names cannot become formulas.
  Model metadata remains in the Very Hidden sheet and hidden direct cell names;
  no user calculation cells are changed by saving/restoring the configuration.

The user still needs to save the workbook normally to retain edits on disk.
Charts, percentiles, probability/target algorithms, sensitivity, reports and the
simulation mutation/restoration boundary are unchanged. Old add-in versions can
read the original columns but may discard new optional fields when they save.

Automated coverage links the actual WorkbookPersistence, ExcelSimulationWorkbook
and SimulationService into Core.Tests behind a test-only Excel-DNA application
boundary and in-memory workbook objects. Tests exercise every distribution and
parameter, old layouts and upgrades, target/clear/direction/requested-confidence
round trips, multiple forecasts, renamed/moved/broken/deleted links, corrupted
metadata, workbook isolation, and source-workbook updates after report activation.
A save/reopen/simulate test recreates workbook objects, clears/poisons in-memory
model state, calls the production simulation entry point, verifies samples use
restored parameters and verifies input restoration. These are adapter/lifecycle
unit tests, not a real Excel file/COM save-close-reopen test.

Validation: baseline 200 passed; final 238 passed, 0 failed/skipped (38 new cases).
Complete Release solution build and Excel-DNA packaging succeeded in
`artifacts/persistence-release`; both 32-bit and 64-bit packed XLLs were generated.
Whitespace checks passed. Existing warnings remain in LicenseService (CS0162),
AssumptionForm (CS8600), and WorkbookPersistence (CS8603); the latter is also emitted
when its source is compiled into the test project.

Still requires real Excel verification: save/close/reopen with old and new workbooks,
workbook event subscriptions and switching, persisted target/cleared-target and
confidence displays, sheet rename/move/delete and cell insertion/deletion handling,
normal numeric 2042 versus actual Excel errors, hidden name/text-metadata behavior,
protected/read-only save failures, and Export Report followed by a target update.
No live workbook was modified for these automated tests. The generated packed XLLs
were built and inspected, not loaded into Excel during this review.
