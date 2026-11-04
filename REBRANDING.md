# MCsim rebranding and compatibility

MCsim is the product name. The repository default branch is `develop`; this change
belongs on `rebrand/telivu-to-mcsim`. No repository rename, release publication,
website deployment or merge is part of this change.

## Scope assessed

| Area | Result |
| --- | --- |
| Runtime UI | Ribbon tab, About dialog, callback/type names and product metadata use MCsim. |
| Calculation source | Core mathematics, simulation/optimization/scenario/SPC execution, seeds and restoration are unchanged. |
| Excel interface | Existing `MC.*` worksheet functions, macro commands, namespaces, assembly names and resource identities remain stable. |
| User data | Hidden config sheet, workbook names, JSON envelopes, logs and legacy activation state retain existing identities; no data migration or deletion. |
| Installer | Display name is MCsim; new executable is `MCsim-Setup-0.2.1.exe`, with `MCsim-0.2.1-manifest.json`. Build and lifecycle/runtime probes use the same new name. |
| Upgrade identity | AppId, setup mutex, registry ownership key, default installation directory and installed `MonteCarloForExcel.xll`/sidecar remain stable. The product is an Excel add-in, not a standalone application EXE. |
| Website | Page titles, navigation, descriptions, copyright, configuration global, social card and browser assertions use MCsim. Script cache keys change together to prevent mixing old/new configuration globals. Existing diagrams and analysis screenshots remain applicable. |
| Distribution links | The real `telivu` website/release repository, hosted project path, historical asset names and hashes are retained. No nonexistent MCsim release URL is substituted. The website explains that the available download predates the rename. |
| GitHub Actions | Adds MCsim build/test/website validation for develop and PRs. The pre-existing untracked `pages.yml` is left untouched and excluded from this change. Website publication remains in its separate repository. |
| Documentation | Current instructions and project handoff use MCsim. Historical implementation/validation reports retain their original names and evidence. License status and ownership requirements are unchanged. |

## Verification

Run the full xUnit suite, Debug/Release solution builds, developer fixture builds,
native layout checks, website build/link/browser checks and installer build checks.
Installer lifecycle tests use isolated registration/process/runtime fixtures; they
do not certify clean-machine deployment. Private Excel startup tests and real x86
Excel/clean Windows deployment remain distinct from unit and build validation.

Do not overwrite historical public release assets. Publishing a newly built MCsim
installer and updating the shared download configuration require a separate release.
