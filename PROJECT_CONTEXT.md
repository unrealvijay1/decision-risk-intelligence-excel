# Monte Carlo for Excel --- Project Context

**Purpose:** Permanent handover/context document for continuing
development of the Monte Carlo for Excel product across future coding
sessions.

**Last updated:** 14 September 2026\
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
-   **MonteCarlo.Tests** --- automated tests where applicable.
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
-   Automated regression tests.
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
