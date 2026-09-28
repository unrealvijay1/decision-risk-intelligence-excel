# SPC V1 delivery

SPC is available under **Process Analysis → SPC Analysis**, independently of simulation
and scenario analysis. No existing simulation, forecasting, scenario transformation or
correlation mathematics were changed. No dependencies were added. No commit or push was made.

## Architecture and agreed decisions

- `MonteCarlo.Core` owns calculation, immutable result contracts and independent signal detection.
- Excel adapters own strict range reading, tracked references, configuration persistence and export.
- Dedicated configuration/results forms use WinForms/GDI+ and existing range-picker conventions.
- The user confirmed conventional constants **2.66 / 3.268**, rejection of zero-variation baselines,
  and persisted fixed limits with an explicit **Establish / replace baseline** action.
- **Analyze** reuses saved limits. This includes “All observations,” source edits and newly selected
  extended ranges. Baseline option changes take effect only through explicit replacement.
- One configuration per workbook uses the existing very-hidden `__MonteCarloConfig` sheet,
  versioned `SpcAnalysisV1` JSON in AB:AC and hidden `_MC_SPC_` names. Normal Save is required
  for disk persistence. Model saves and Clear Model preserve SPC, including unreadable blocks.

## Statistical behavior

For the selected baseline only:

- Mean = average baseline observations.
- MR(i) = absolute difference between successive observations; the first MR is absent.
- Average MR includes only pairs wholly within the baseline.
- Individuals center = mean; limits = mean ± **2.66 × average MR**.
- MR center = average MR; UCL = **3.268 × average MR**; LCL = **0**.

The complete chronological sequence is evaluated against these fixed limits. A subsequent
MR may use the last baseline value. At least 20 baseline observations are required, with an
explicit reliability caveat. Gaps, Excel errors, non-finite values, zero variation and
unrepresentable ranges/limits are rejected. Negative Individuals limits are not clamped.

Configurable rules detect strictly out-of-limit Individuals/MR points, eight Individuals
strictly on one side of the mean, and six strictly increasing/decreasing Individuals.
Equality interrupts runs/trends. Only the first qualifying endpoint of an uninterrupted
run/trend episode is emitted; different rules may overlap. Signals retain rule, chart,
inclusive observation range, endpoint, direction, actual and reference value.

Business targets are inclusive ≤ or ≥ comparisons, independent of statistical signals.
Results include observed attainment count/percentage, baseline mean and latest attainment.
Insights are deterministic and make no claims of normality, capability, operational cause,
or proven stability. No capability indices or AI functionality were introduced.

## Verification

| Check | Result |
| --- | --- |
| Existing baseline before implementation | 504 passing xUnit cases |
| Final Debug unit suite | **555 passed, 0 failed, 0 skipped** |
| Final Release unit suite | **555 passed, 0 failed, 0 skipped** |
| Added unit coverage | 51 cases: independent reference datasets, baseline boundaries/fixed limits, numerical edges, rules, gaps, target independence, persistence and export tables |
| Debug and Release solution builds | Passed; only the three existing warning locations remain |
| Packed x86/x64 XLL generation | Passed in both configurations |
| Independent package inspection | PE machines 0x14C / 0x8664 and embedded DNA/Core/Excel compressed resources verified for all four XLLs |
| Debug and Release layout harness | SPC, existing Results, Scenario and Correlations checks pass at 100/125/150% geometry scaling |
| SPC visual inspection | Full form and chart previews inspected; summary sizing preserves plot space under scaling |
| Private real Excel integration | Passed, including owned process exit |
| Whitespace check | `git diff --check` passed |

Live SPC checks exercise the actual workbook service: source extension with unchanged limits,
source formulas, strict Excel error handling, readable dates from narrow columns, literal
text labels, tracked name reuse and rename/insert/delete, unique exports, both editable
native charts, blank first MR, literal-text safety and save/close/reopen of fixed statistics.
The synthetic workbook uses a private Excel instance and does not attach to user Excel.

Packed artifacts:

| Configuration | x86 bytes | x64 bytes | Directory |
| --- | ---: | ---: | --- |
| Debug | 1,561,600 | 1,571,840 | `artifacts/spc-debug/` |
| Release | 1,555,456 | 1,565,696 | `artifacts/spc-release/` |

The files are `MonteCarlo.Excel-AddIn-packed.xll` and
`MonteCarlo.Excel-AddIn64-packed.xll`. Builds used isolated output directories without
replacing a loaded XLL. Existing warnings are LicenseService CS0162, AssumptionForm CS8600
and WorkbookPersistence CS8603. The persistence warning also appears in its linked test build.

## Files added

| Area | Files |
| --- | --- |
| Core | `MonteCarlo.Core/SpcAnalysis.cs`, `SpcSignalDetector.cs`, `XmRCalculator.cs` |
| Excel integration | `MonteCarlo.Excel/MonteCarlo.Excel/SpcAnalysisCommand.cs`, `SpcAnalysisService.cs`, `SpcDataReader.cs`, `SpcPersistence.cs`, `SpcExporter.cs` |
| UI | `MonteCarlo.Excel/MonteCarlo.Excel/SpcConfigurationForm.cs`, `SpcResultsForm.cs`, `SpcChartPanel.cs` |
| Unit tests | `MonteCarlo.Core.Tests/XmRCalculatorTests.cs`, `SpcSignalTests.cs`, `SpcIntegrationTests.cs` |
| Layout/live tests | `MonteCarlo.Results.LayoutChecks/SpcLayoutChecks.cs`, `SpcLiveChecks.cs` |
| Delivery reference | `SPC_V1_IMPLEMENTATION.md` |

Files listed without a repeated directory share the directory of the first file in their row.

## Files modified

- `MonteCarlo.Excel/MonteCarlo.Excel/MonteCarloRibbon.cs`: separate ribbon group and callback.
- `MonteCarlo.Excel/MonteCarlo.Excel/WorkbookPersistence.cs`: preserve SPC during save/clear.
- `MonteCarlo.Core.Tests/MonteCarlo.Core.Tests.csproj`: link SPC Excel adapters for testing.
- `MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj`: compile SPC forms.
- `MonteCarlo.Results.LayoutChecks/Program.cs`: run SPC layout checks and optional live checks.
- `MonteCarlo.Results.LayoutChecks/README.md`: validation commands and scope.
- `PROJECT_CONTEXT.md`: reconciled architecture, workflows, persistence, constraints and verification.

## Limitations and validation boundaries

- V1 accepts a single contiguous row or column, with a 100,000-observation ceiling, and one saved
  SPC configuration per workbook. Selected baseline periods use inclusive observation positions.
- The baseline is a statistics snapshot; historical source values are not stored. Re-select
  data/label ranges to include appended cells outside the tracked range. Changed baseline length
  requires deliberate replacement. Superseded hidden names may remain, as with existing tracking.
- Native Excel charts highlight signal endpoints and baseline start/end markers. They do not
  reproduce the on-screen baseline background shading or every participating Individuals point.
  This distinction is stated in the report; both native charts remain editable.
- Export failure can leave a partial newly created report; existing worksheets are never overwritten.
- Geometry/rendering tests are not actual per-monitor DPI tests. Interactive ribbon clicking,
  range-picker interaction, customer-specific protected workbooks, clean-machine installation,
  and live x86 Excel were not certified. The callback/ribbon binding was inspected and built;
  range reading/tracking and the full analysis service were tested through real Excel COM.

Run the optional live check with:

```powershell
dotnet run --project MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj -c Release --no-restore -- --live-spc
```
