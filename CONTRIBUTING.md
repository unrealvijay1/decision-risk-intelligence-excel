# Contributing to Telivu

## Bugs and features

Use [GitHub issues](https://github.com/unrealvijay1/telivu/issues). Include Windows/Excel versions and Excel architecture, module, expected/actual behavior and reproducible steps with synthetic data. Do not attach confidential workbooks or private credentials. For feature requests explain the decision workflow and a small example before proposing UI changes.

## Development setup

Application development requires Windows, .NET 10 SDK and a checkout containing MonteCarlo.Excel/MonteCarlo.Excel.slnx. Desktop Excel is needed only for live COM integration checks. Installer builds require Inno Setup; the current default compiler is C:\Program Files\Inno Setup 7\ISCC.exe. Python 3.10+ builds the website. Read AGENTS.md and PROJECT_CONTEXT.md before implementation. The public telivu repository presently contains the website/release distribution; application source is available in the existing decision-risk-intelligence-excel repository; legal license approval is pending owner review.

From the application repository root:

```powershell
dotnet restore MonteCarlo.Excel/MonteCarlo.Excel.slnx
dotnet build MonteCarlo.Excel/MonteCarlo.Excel.slnx -c Debug --no-restore
dotnet build MonteCarlo.Excel/MonteCarlo.Excel.slnx -c Release --no-restore
dotnet test MonteCarlo.Core.Tests -c Debug --no-restore
dotnet test MonteCarlo.Core.Tests -c Release --no-restore
dotnet run --project MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj
./build-installer.ps1
git diff --check
```

Use an isolated ExcelDnaPublishPath when Excel locks the usual output; never close a user's Excel instance or replace a loaded XLL merely to build. Installer validation isolates registration state. Read the fixture READMEs before running live checks; those checks create private Excel instances and never certify arbitrary customer workbooks or a clean Windows VM.

## Website development

```powershell
python site/scripts/build.py --site-url https://unrealvijay1.github.io/telivu/
python site/scripts/check.py site/_build
python -m http.server 8765 --bind 127.0.0.1 --directory site/_build
```

Browser checks use node site/scripts/browser-check.mjs with an existing Playwright installation (PLAYWRIGHT_PATH) and optional browser executable (BROWSER_PATH). All download buttons must retain the existing data-download/config mechanism and published asset. Never create a new release as part of a documentation change.

## Pull requests

Explain the problem, resulting behavior, relevant validation and limits. Preserve worksheet function names, workbook formats, input restoration and installer ownership. Add meaningful regression coverage for changed behavior. Update PROJECT_CONTEXT.md when future sessions need the new facts. Avoid generated binaries, keys, customer workbooks and unrelated formatting. Approval of a software license is an owner decision; no contribution automatically resolves existing ownership concerns.
