# SPC V1 — Wheeler's Four Tests

SPC remains independent of Monte Carlo and Scenario Analysis. This implementation adds
Wheeler tests 3 and 4, preserves the optional six-point trend, and changes no sampling,
simulation, scenario, control-limit or fixed-baseline calculations. No dependencies were
added. No commit or push was made.

## Statistical and signal decisions

- Individuals limits remain mean ± **2.66 × baseline average MR**; MR UCL remains
  **3.268 × baseline average MR**, LCL 0.
- Sigma = **baseline average MR / 1.128**. Zones are baseline mean ± sigma and ± 2 × sigma.
  They are derived from saved statistics, never recomputed from appended/current values.
  Explicit baseline replacement updates both limits and zones.
- Rule 3 detects at least two of three strictly above the upper 2σ threshold or strictly
  below the lower 2σ threshold. Rule 4 detects at least four of five beyond the same 1σ
  threshold. The remaining observation can be anywhere; equality never qualifies.
- Every complete consecutive window is evaluated. Each rule/direction emits its first
  qualifying window, suppressing consecutive qualifying windows until a nonqualifying
  window ends the episode. The next episode can signal again. Different rules remain
  independent. Existing run/trend first-endpoint behavior and MR detection are unchanged.
- Wheeler rules use IDs 1–4; the supplementary trend now uses ID 5 (previously 3).
  Signals are not persisted, so renumbering does not migrate or invalidate saved baselines.
- Signals retain direction, one-based window bounds/endpoint, threshold, qualifying indices
  and actual window values. The endpoint itself need not qualify. Missing/non-finite data
  remain rejected rather than allowing any window to bridge a gap.

## Configuration, persistence and presentation

The configuration form groups the first four controls under **Wheeler's Four Tests** and
Six-point trend under **Supplementary Rules**. All five default to enabled and can be
selected independently. The existing scrollable form keeps every control reachable.

The existing **SpcAnalysisV1** JSON envelope gains two optional rule fields, `TwoOfThree`
and `FourOfFive`. Missing fields initialize true while all three older selections survive.
All five selections round-trip. Saved baseline fields and tracked names are unchanged;
zone properties are derived and excluded from JSON. Normal Excel Save remains necessary.

Individuals charts draw red rings around qualifying blue data points. A scrollable details
grid shows each rule's name/ID, direction, window, qualifying observations, threshold and
actual values. Multiple rules remain separate entries. Deterministic insights describe
upper/lower zone signals and multiple signals without claiming an operational cause.

Export creates a unique worksheet and preserves the existing summary, target evaluation,
control limits, aligned data and two editable native charts. Each qualifying observation
retains every applicable rule ID and description, including full window context. Four zone
columns are appended at S:V. Native charts highlight qualifying observations; they do not
reproduce the on-screen baseline shading. Source cells and previous reports remain intact.

## Validation

| Check | Result |
| --- | --- |
| Baseline before changes | 555 xUnit cases passed |
| Final Debug suite | **582 passed, 0 failed, 0 skipped** |
| Final Release suite | **582 passed, 0 failed, 0 skipped** |
| Added coverage | 27 cases for both zones/directions, strict equality, insufficient/opposite-side points, all qualifying points, nonqualifying endpoints, episodes, disabling, gaps, fixed/replaced zones, legacy/five-rule persistence and multi-rule exports |
| Debug/Release solution builds | Passed; only existing warning locations |
| x86/x64 packed XLLs | Generated in both configurations; PE architecture and embedded packed resource names checked |
| Debug/Release layout harness | SPC, Results, Scenario and Correlations pass at 100/125/150% geometry scaling |
| SPC UI review | Rendered charts/forms inspected; rule scroll access and complete signal details checked |
| Private Excel SPC integration | Passed: zone detection/export, fixed baseline, source/formula preservation, strict errors, text/date labels, native charts, unique reports, tracked references, save/reopen and owned process exit |
| Private Excel Scenario integration | Passed: success, cancellation, failure, source isolation and owned process cleanup |

Packages are in `artifacts/wheeler-debug/` and `artifacts/wheeler-release/`:
`MonteCarlo.Excel-AddIn-packed.xll` (x86) and `MonteCarlo.Excel-AddIn64-packed.xll` (x64).
Existing warnings remain LicenseService CS0162, AssumptionForm CS8600 and
WorkbookPersistence CS8603 (also linked into unit tests).

## Files changed

| Area | Files (relative to repository root) |
| --- | --- |
| Core | `MonteCarlo.Core/SpcAnalysis.cs`, `MonteCarlo.Core/SpcSignalDetector.cs` |
| UI/export | `MonteCarlo.Excel/MonteCarlo.Excel/SpcConfigurationForm.cs`, `SpcResultsForm.cs`, `SpcChartPanel.cs`, `SpcExporter.cs` (same directory) |
| Unit tests | `MonteCarlo.Core.Tests/SpcZoneTests.cs` (new), `SpcSignalTests.cs`, `SpcIntegrationTests.cs`, `XmRCalculatorTests.cs` (same directory) |
| Layout/live checks | `MonteCarlo.Results.LayoutChecks/SpcLayoutChecks.cs`, `SpcLiveChecks.cs` (same directory) |
| Documentation | `PROJECT_CONTEXT.md`, `SPC_V1_IMPLEMENTATION.md` |

## Remaining limitations

- One configuration per workbook; one contiguous row/column with 20–100,000 observations.
  The baseline stores statistics, not historical source values. Re-select ranges to include
  appended cells; changed baseline length requires explicit replacement.
- Episode deduplication retains the first qualifying window, not every continuing window.
- Export failure may leave a partial new report. Superseded tracked names can remain.
- Geometry/render checks do not certify actual per-monitor DPI. Interactive ribbon/range
  picker operation, arbitrary customer/protected workbooks, clean-machine installation,
  code signing and live x86 Excel remain outside this validation.
