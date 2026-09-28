# Monte Carlo for Excel — Project Context

Last reconciled with source and tests: **2026-09-28**.

This is the persistent project handoff and source of truth for future Codex sessions.
Read it before implementation. Verify relevant source before changing behavior; reconcile
discrepancies rather than relying on previous chat history. After every implementation
task, update affected sections when future sessions need the information, remove obsolete
claims, and verify this document against the final code. Keep current decisions and
capabilities, not a change log, temporary debugging details or build transcripts.

## Architecture and repository map

Windows Excel add-in using C#/.NET 10, Excel-DNA 1.9 and WinForms/GDI+.
Primary customer platform is **64-bit Excel**; builds also produce an x86 XLL.

| Location (relative to repository root) | Responsibility |
| --- | --- |
| `MonteCarlo.Core/` | Sampling, fitting/preview, probability/CDF, chart helpers, scenario rules and independent SPC/XmR mathematics; `net10.0` |
| `MonteCarlo.Excel/MonteCarlo.Excel/` | Ribbon, model/UI, Excel COM adapters, persistence, execution, reporting and licensing; `net10.0-windows` |
| `MonteCarlo.Excel/MonteCarlo.Excel.slnx` | Main solution |
| `MonteCarlo.Core.Tests/` | xUnit; links production non-UI Excel source against workbook doubles |
| `MonteCarlo.Results.LayoutChecks/` | Separate Windows-only hidden WinForms layout/rendering harness |
| `MonteCarlo.LicenseGenerator/` | Developer-only offline license signing utility |
| `Installer/MonteCarloForExcel.iss` | Per-user Inno Setup installer |

`SimulationModel` holds static in-memory assumption/forecast lists for the active workbook.
`WorkbookPersistence` reloads them on workbook open/activation, explicit Reload Model and
before simulation. `SimulationService` coordinates restore/settings and execution;
`SimulationExecution` owns validation and mutation/restoration through `ISimulationWorkbook`
and `ISimulationCell`. `ExcelSimulationWorkbook` implements the Excel COM boundary.
Results use `SimulationRunResult` with one `ForecastRunResult` per forecast. Each successful
Ribbon run opens a fresh modal `ResultsForm`; chart/calculator actions do not rerun simulation.

## Implemented model workflows

- Ribbon groups: Model Setup (Define Assumption, Define Forecast, Model Manager, Correlations, Clear
  Cell Definition); Simulation (Run Simulation, Simulation Settings, Scenario Analysis, Reload Model);
  Process Analysis (SPC Analysis); Maintenance (Clear Model); Product (License). Export Report is in Results.
- Six distributions: Normal (mean, SD), Lognormal (log mean, log SD), Uniform (min, max),
  Triangular and PERT (min, most likely, max), Beta (min, max, alpha, beta). Preserve
  parameter order across UI, sampling and persistence.
- Define/Edit Assumption provides analytical previews of unsaved parameters with inline
  validation. Previewing neither samples nor writes Excel. Historical-data fitting
  supports Normal, Lognormal, Uniform and Triangular, ranked by AIC then KS; fitting
  does not support all six sampling distributions.
- Model Manager supports viewing, navigation, editing and confirmed deletion.
  `ForecastEditForm` uses detached `ForecastEditDraft` for name, explicit target,
  direction, default P80 or No Target. Cancel/invalid drafts do not apply changes.
  Unchanged target/direction retain requested confidence; changes clear its provenance.
- `ModelDefinitionDeletion` is shared by Model Manager and Clear Cell Definition.
  The Ribbon action accepts one cell only, identifies the definition in confirmation,
  and preserves values/formulas. Ordinary cells are an informational no-op; invalid
  selections are rejected. Lookup loads without highlighting and restores prior
  in-memory lists on cancellation/no-op. Confirmed deletion saves remaining definitions.
- Assumption Correlations supports model-level input relationships. Classifications and
  decision-variable management are not implemented.

## Workbook persistence and cell highlighting

Storage is the **Very Hidden `__MonteCarloConfig` worksheet**, with headers in row 1
and definition rows starting at row 2. Normal Excel Save is required for disk persistence;
writing configuration does not automatically save the workbook file.

| Columns | Current meaning |
| --- | --- |
| A:D (1–4) | Type, Sheet, Cell, Distribution |
| E:H (5–8) | Parameter1–Parameter4 |
| I (9) | Friendly Name |
| J (10) | CellLink: hidden workbook-scoped `_MC_Cell_` GUID name |
| K:N (11–14) | TargetConfigured, Target, TargetDirection, RequestedConfidence |
| O (15) | `OriginalFillV1`: original Interior snapshot as JSON |
| Q:T (17–20), row 1 only | `SimulationSettingsV1`, trial count, seed mode, optional fixed seed |
| V:W (22–23) | V1=`ScenarioDefinitionsV2`, W1=JSON chunk count; V2 onward holds 30,000-character chunks; version V1 remains readable |
| Y:Z (25–26) | Y1=`AssumptionCorrelationsV1`, Z1=chunk count; Y2 onward holds 30,000-character JSON chunks |
| AB:AC (28–29) | AB1=`SpcAnalysisV1`, AC1=chunk count; AB2 onward holds 30,000-character JSON chunks for SPC configuration, tracked ranges and fixed baseline statistics |

The legacy three-parameter/name-in-H and four-parameter/name-in-I layouts remain readable.
Optional metadata is backward compatible. Text metadata is written as text to avoid formula
interpretation. Invalid saved definitions/targets and broken references surface structured
validation errors and block simulation. Event registration/restore failures have safe
messages and Trace diagnostics; initial restore is independent of event registration.

Hidden names track direct single-cell references through sheet renames and structural
edits. Restore resolves current sheet/address. Broken/deleted/cross-workbook links never
fall back to stale coordinates. Legacy coordinate-only definitions acquire links on save;
renames before tracking cannot safely be inferred. Highlight resolution uses these names
on load/save, not continuous cell-movement monitoring.

`ModelCellHighlight` centralizes light blue Assumption and light green Forecast fills.
Create/edit/save/restore apply the appropriate fill, capturing original Interior before
first application. Repeated edits retain that snapshot. Deletion restores it and removes
metadata; no-fill, solid, theme/tint, pattern and gradient fills are represented. Only
Interior is intentionally changed, never values, formulas, number formats, font, borders,
alignment or protection. Legacy rows without snapshots capture their current fill before
highlighting. Formatting failures are best-effort and traced: unavailable/protected/deleted
cells can prevent restoration, and an original fill never captured cannot be reconstructed.

Definition saves preserve the separate settings, scenario and SPC blocks verbatim, even if malformed.
Clear Model restores definition fills and removes definitions, retaining
explicitly saved simulation settings, scenarios and SPC. Without any of those blocks it deletes the config
sheet. Deleted definitions lose their saved CellLink metadata, but the current code does
not remove the corresponding hidden workbook names; unused names can remain. Older add-in
versions may discard newer optional metadata when saving.

### Persisted forecast targets

- `ForecastDefinition.TargetSettings == null` means the historical P80 default.
  A non-null `ForecastTargetSettings` with null Target means explicit **No Target**.
- Accepted Results targets, direction changes and target clearing update only the
  matching saved forecast row. Draft numeric edits are not accepted targets.
  Results pins the source workbook so later active-workbook changes cannot redirect
  these updates. Save failure reports a warning.
- Targets must be finite; directions supported; requested confidence, if present, must
  be finite, 0–100 and accompanied by a target. Requested confidence is provenance,
  not a promise of actual success probability. Saved targets remain fixed on a new run;
  they are not automatically recomputed from fresh samples.
- Preserve explicit `object?` boxing of nullable Target/RequestedConfidence at the
  dynamic COM `Value2` boundary. Passing `Nullable<double>` directly caused a live
  clear/redefine failure. Empty values must remain null, not placeholder numbers.

## Simulation settings, execution and safety

`SimulationSettings` is a workbook-specific immutable record: TrialCount, SeedMode
(Automatic/Fixed), FixedSeed. Defaults: **10,000 trials, Automatic**. Presets are
1,000/5,000/10,000/50,000; custom count is 1–1,000,000. Fixed seeds accept signed 32-bit
integers. UI and execution share validation. Run Simulation uses saved settings without
reopening the dialog. Missing settings use defaults; malformed settings block normal
execution. The older `TryRun(int)` API remains supported with explicit trials/Automatic.

Each execution creates one `Random(actualSeed)` and passes it through all six samplers
and nested helpers. Automatic seeds are generated once per run. Results retain configured
settings and actual seed without changing Automatic to Fixed. Same-seed replay is tested;
volatile Excel formulas, external data and runtime changes are outside this guarantee.

Before mutation, validation checks definitions/parameters, single-cell references,
duplicate assumption cells, editable inputs outside merged/array/spill ranges, and finite
numeric input/output values. The adapter identifies actual Excel errors before conversion;
ordinary numeric 2042 must not become `#N/A`. Every original input value/formula and saved
application setting is captured first. Formula2 is preferred, with supported fallback
to Formula. During trials, inputs are sampled, Excel recalculates and forecasts are validated.

`finally` independently restores every input, recalculates, and restores ScreenUpdating,
EnableEvents and StatusBar, retrying each failed operation once. Persistent restoration
failures prevent success and identify what needs checking before saving. Excel closing
or permanently refusing writes cannot be recovered reliably. Errors use structured
outcomes/user messages and `System.Diagnostics.Trace`; no durable logging store exists.
The legacy `MC_RUN_PROJECTMODEL` command shares this boundary while retaining fixed PERT
inputs and output-cell-only recalculation.

Statistics use population SD and linearly interpolated percentiles at `p * (n - 1)`.
Sensitivity is Spearman rank correlation with average ranks for ties, ordered by absolute
correlation. Preserve these semantics when changing presentation.

## Assumption Correlations

Model Setup → Correlations opens `AssumptionCorrelationsForm`: add/edit/delete pairs using
names plus references and a **Correlation** coefficient in [-1,+1]. Self-pairs, reversed
duplicates, missing assumptions and non-finite/out-of-range coefficients are rejected.
Zero relationships are omitted on save. Missing references remain visible for correction.

The coefficient means **latent Gaussian dependence**, not the Pearson correlation of final
transformed samples. There is no calibration to force final Pearson correlation. The editor
explains that observed correlation depends on the marginal distributions. `GaussianDependence`
constructs a symmetric unit-diagonal matrix with unspecified pairs zero and performs
semidefinite Cholesky once per engine run. Valid singular matrices (including ±1) are
supported; inconsistent combinations are rejected before input writes. Numerical tolerance
is 1e-12; coefficients are never adjusted to repair an invalid matrix.

`DistributionQuantile` maps correlated Gaussian draws through the configured marginals:
Normal/Lognormal use the equivalent direct normal transform; Uniform/Triangular use closed
forms; Beta/PERT use incomplete-beta continued fractions and inverse search. It reuses the
existing normal CDF and log-gamma helpers. Floating-point endpoint limits apply; numerical
nonconvergence/range failures stop the run with restoration. All six distributions are tested.
No-correlation and zero-only models retain the exact independent seeded sequence. Unlinked
inputs retain their existing samplers; correlated vectors share the run's seeded Random.
Adding relationships changes random consumption, so replay requires the same full configuration.

`CorrelationPersistence` stores only pairs of tracked assumption IDs and coefficients in
the independent versioned block. Legacy definitions gain links on correlation save. Normal
definition saves preserve relationships and remove those whose assumptions were removed;
Clear Model removes correlations while retaining its existing settings/scenario behavior.
Missing references produce actionable validation, never silent independent execution.
Unreadable correlation blocks survive unrelated saves and block execution. Old workbooks
without the block load normally. Normal Excel Save is required for disk persistence.

Normal and scenario runs pass model-level correlations to `SimulationExecution`. Scenario
runs retain the same dependence and common draws, applying their existing marginal scaling
after sampling; zero-scale scenarios remain constants. Sensitivity still uses actual sampled
inputs and the unchanged Spearman method. With correlated inputs, sensitivity reflects shared
movement and must not be interpreted as isolated causal contribution. Results UI is unchanged.

Independent sampling retains its fast path. Correlated Beta/PERT inverse transforms are more
expensive; the optional harness `--correlation-performance` measures sampling without Excel
recalculation, so its timings are not end-to-end workbook forecasts.

## Results UI and calculation semantics

`ResultsForm` is **1120 × 830**, with exactly three primary tabs: **Forecast** (default),
**Sensitivity**, **Statistics**. One forecast selector is shared above them; Export Report
and Close are in a shared footer. The agreed visual-polish pass is complete: preserve this
structure, size, chart prominence and **34% controls / 66% chart** Forecast workspace
unless a new request calls for redesign or live testing finds a functional regression.

- Forecast left: dominant actual probability, caption **Probability of achieving target**,
  target direction/value, Decision Tools, Success when selector and compact
  **Target → Probability** / **Probability → Target** calculator tabs. Each calculator
  displays its result directly; P50/P80 are not headline KPI cards.
- Forecast right: large chart with **Distribution / Cumulative Probability** selector.
  Chart mode is static session-only state retained across Results windows. Forecast
  switching refreshes all tabs; chart/calculator tab changes do not simulate.
- Sensitivity: its own large tornado chart with existing correlation ordering.
- Statistics: Distribution Statistics (Mean, Std Dev, Minimum, Maximum, P10, P50, P80,
  P90) and Simulation Run (Trials, Seed Mode, actual Seed).
- Preserve content-driven row sizing and positive chart plot space under scaling.
  Do not reintroduce shared AutoSize rows whose scaled supporting controls starve the
  chart, or percent-height compression of calculator/content rows.

### Targets and calculators

Success direction is explicit and never inferred: AtOrBelow means `X <= target`;
AtOrAbove means `X >= target`. Probability is matching samples / all samples.
`P(X >= target)` is **not** `1 - P(X <= target)` because equality mass matters.
Without direction, lower-tail probability can be calculated but no success claim or
Success/Miss shading is shown. Missing target clears target-dependent presentation.

When TargetSettings is null, the default target still parses P80 formatted to **two
decimal places**. It can differ from the exact P80 marker. Saved/manual/reverse-calculated
accepted targets retain double precision. `SetTargetDisplay` caches the exact value
separately from formatted text; recalculating unchanged text uses that value. Actual
user edits clear the cache and accepted display; blank text persists explicit No Target.
Forecast refresh restores saved target state.

Probability → Target accepts 0–100 and uses the same percentile interpolation as statistics:
lower-tail `C/100` for AtOrBelow or unspecified direction, `1 - C/100` for AtOrAbove.
Requested probability can differ from actual empirical success for duplicate/discrete
values and interpolation. Invalid reverse submissions retain accepted state. Direction
changes recompute confidence-derived targets using accepted confidence, not pending
textbox edits; manual targets remain fixed. Failed reverse recalculation restores the
prior direction. Target → Probability accepts a manual target and clears previous
requested-confidence provenance.

### Chart behavior

Both GDI+ renderers reuse exact stored P50/P80 and shared coordinate helpers. Percentiles
are thin/dashed; Target is orange/solid. Labels include formatted values and target
direction, use separate staggered rows, clamp within chart bounds, and connect vertically
to the exact marker X coordinate. Labels must not shift actual marker positions. Numeric
formatting is culture-aware; worksheet currency/date/percentage formats are not imported.

Success/Miss background regions split at the target, not histogram-bin edges. Bars/curve
remain above shading. Labels sit inside sufficiently wide regions near the baseline;
there is no separate legend strip. Shaded area is not probability mass.

Distribution uses 15–30 bins. Out-of-range targets retain an annotation without a target
line or expanded axis; whole-plot shading follows direction. Constant outcomes retain
the blank histogram while empirical probability remains available.

Cumulative is a cached, right-continuous empirical step CDF (0–100%) built from a sorted
copy, grouping duplicate outcomes. Non-finite data makes the curve unavailable rather
than silently removing samples. Constants get a padded range; finite out-of-range targets
expand the axis. Its orange guide/dot always represents actual `P(X <= target)`, including
when success means `X >= target`. Interpolated P50/P80 need not intersect the step curve
at exactly 50%/80%.

`ReportExporter` creates a uniquely named `MonteCarlo_Report_...` sheet in the active
workbook with run information, assumption definitions, forecast statistics and sensitivity.
It is a separate report implementation, not a capture of the Results UI.

## Scenario Analysis

Sensitivity answers “What drives the outcome?” Scenario Analysis answers “What happens
under different assumption scenarios compared with a user-selected baseline?” The baseline
is the chosen scenario, not the unchanged workbook model.

The separate Scenario Analysis Ribbon command opens `ScenarioAnalysisForm`; normal
ResultsForm is unchanged. Users create/rename/delete named scenarios, check which to run,
edit percentage changes for multiple assumptions, or clear an adjustment to No change.
Names are unique ignoring case/outer whitespace. Empty scenarios are allowed. Users
explicitly select one user-defined scenario as baseline; its adjustments execute normally.
There is no synthetic baseline. Selection is stored by ID, survives rename, clears on
deletion and is always included in the run selection. Missing baseline blocks execution.
Save definitions writes metadata;
Run saves definitions before execution. Normal Excel Save is still required for disk storage.

Core owns `ScenarioDefinition`, extensible `ScenarioAdjustmentType`, validation and
`ScenarioTransformation`. V1 exposes only PercentageChange with business meaning
`Y = (1 + percentage / 100) X`:

- Normal: scale mean and SD.
- Uniform: scale bounds; Triangular/PERT: scale min/mode/max, retaining relative shape.
- Beta: scale bounds, retain alpha/beta.
- Lognormal: add log(factor) to log mean, retain log SD.
- Exactly -100%: explicit zero constant. Below -100% is rejected rather than reflecting
  distributions/signs. Non-finite percentages, invalid source/transformed parameters and
  representational overflow are rejected; sample-time overflow also stops execution.
  Core supports constant transformation, but no new constant assumption type/UI was added.

`ScenarioSimulationService` clones input/forecast definitions, validates all selected
plans before creating a sandbox, and calls the existing `SimulationExecution` for the chosen
baseline first, then each other selected scenario exactly once. It never assigns global model lists. The engine's optional sample
transform and cancellation token leave normal callers unchanged. Scenario runs sample
the original distribution, then scale the sample using Core mathematics; this is equivalent
to the validated transformed distribution and preserves exact common RNG consumption,
including PERT/Beta rejection sampling and zero-scale cases. Do not both scale parameters
and scale samples. One fixed seed is shared across comparisons: the saved fixed seed or
one newly generated automatic seed. Normal simulation settings are not overwritten.

`ScenarioAnalysisCommand` captures a read-only model snapshot (no highlight restore) and
restores prior in-memory lists. `ScenarioExcelSandbox.Prepare` uses SaveCopyAs once per
analysis. A dedicated STA worker opens that file in a private hidden Excel application;
only the copied workbook receives trial values and recalculation. The complete workbook
copy preserves formulas, names and sheet references. Iteration settings are copied;
calculation is explicit/manual between trials. The engine restores inputs/formulas and
recalculates after every run before the same sandbox is reused. N selected scenarios cost
N simulations plus one file copy/Excel startup/cleanup; no copy per scenario or extra baseline run.

Owned COM objects are cached and released; cleanup closes without saving, quits the owned
Excel instance and deletes the unique temporary file/directory. Cleanup also runs after
setup failure, cancellation and execution failure, and failures are reported rather than
claimed successful. Cancellation is cooperative between trials; an in-flight COM call or
unresponsive Excel can delay it. The original Excel instance is never used for trial
calculation. Source workbook cells, formulas and normal settings remain unchanged; only
explicit scenario-definition saves and legacy tracking metadata writes affect the source.

V1 requires a self-contained workbook: VBA projects, external links and data connections
are rejected before copying. Macros are disabled in the private instance. External/add-in
UDF availability and volatile/iterative models can limit comparability; the private instance
does not promise the user's loaded add-in environment. No forced process termination is
used if Excel refuses to quit; cleanup failure must remain visible.

The Analysis Question tab selects one global mode. **Probability → Value** uses a separate
0–100% probability and direction per forecast, held constant across scenarios. **Value → Probability** uses a target
and direction per forecast, held constant across scenarios. Both modes retain their own
direction settings when switching, independently of normal Results targets. Execution
requires complete settings for every current forecast. Existing ForecastRunResult quantile
and inclusive probability methods provide the calculations with unchanged semantics.
Probability → Value asks “At a specified probability, what value does each scenario produce?”
Value → Probability asks “What probability does each scenario have of meeting a specified target?”

`ScenarioPersistence` stores a chunked JSON **ScenarioDefinitionsV2** envelope containing
scenarios, baseline ID, mode and per-forecast questions including each probability. The older
common probability remains a compatibility fallback for questions without their own value;
the editor saves explicit per-forecast probabilities. This is an additive V2 field. V1 scenario
lists migrate without an invented baseline or question settings; users must configure
these before running. Incomplete drafts may be saved. Results are never persisted.
Adjustments and forecast questions use CellLink identities; legacy coordinate identities
upgrade on save without changing normal parameters, targets or fills. Definition saves
and Clear Model preserve the block. Missing definitions remain visible for removal and
block affected execution; invalid/unknown versions block scenario editing while normal
simulation remains available.

`ScenarioComparisonForm` has six columns: Scenario, Forecast, Mean, Δ Mean, and either
Probability Value/Δ Value or Probability/Δ Probability (percentage points).
Each forecast is compared only with its matching baseline result. Baseline deltas are
em dashes; a subtle fill and star identify baseline rows. Friendly forecast names include
cell references. The top identifies baseline, question, trials and common seed; separate
scrollable forecast context lists each probability/direction or target/direction. Scenario comparison
has no P50/P80 or generic Success/Target/Direction columns. Normal Results, Statistics
and Sensitivity retain their existing presentation and calculations.

## Process Analysis / SPC V1

The separate **Process Analysis → SPC Analysis** ribbon action opens `SpcConfigurationForm`;
`SpcResultsForm` is independent of the simulation and scenario windows. `SpcAnalysisCommand`
pins the source workbook, uses the existing hidden-form Excel `InputBox(Type: 8)` selection
pattern and delegates to `SpcAnalysisService`. No simulation, forecasting, correlation or
scenario mathematics are involved. One saved SPC configuration/baseline is supported per workbook.

`SpcDataReader` accepts one contiguous row or column of 20–100,000 observations in existing
order. Blank/text/error/non-finite observations are rejected with their index; nothing is
silently dropped and no moving range bridges a gap. Excel error detection distinguishes
error HRESULTs from legitimate numbers. Optional labels must have matching length; Excel
dates remain readable even when narrow source columns display hashes. Cross-workbook and
multi-area selections are rejected. Data/range selection is read-only until an explicit
analysis action writes metadata.

Core owns immutable configuration, baseline and signal records, `XmRCalculator` and the
separate `SpcSignalDetector`. Baselines are all observations, first N or an inclusive selected
contiguous observation period, with at least 20 observations. V1 explicitly uses the agreed
conventional factors **2.66** and **3.268**: Individuals mean ± 2.66 × average successive
baseline moving range; MR center = average moving range, UCL = 3.268 × that average, LCL = 0.
Individuals LCL is not clamped to zero. Only pairs wholly within the baseline establish MR̄.
The first observation has null MR; subsequent MR can use the last baseline observation.
Scaled compensated means avoid sum overflow; non-finite ranges/limits, unrepresentable limits
and zero-variation baselines are rejected with actionable guidance.

**Establish / replace baseline** deliberately computes and saves limits. **Analyze** reuses
the saved statistics, including after source values change or a longer source range is
selected. Even “All observations” freezes at establishment. Baseline-option edits apply only
when replacing it. Hidden `_MC_SPC_` workbook names track data, labels and baseline through
renames/structural edits. A changed baseline length or broken link blocks reuse; no stale
coordinate fallback exists. Unchanged names are reused; superseded names can remain, like
existing cell tracking. Re-select the data/label ranges to include appended observations
outside their tracked range. Limits are snapshots; historical source values are not persisted.

Rules are individually configurable: strictly outside limits on both charts; eight
Individuals observations strictly above/below center; six strictly increasing/decreasing
Individuals observations. Equality interrupts runs/trends. Run/trend episodes report their
first qualifying endpoint only; different rules may overlap. All chronological observations
are evaluated against fixed limits. Signals carry rule/chart, one-based endpoint and inclusive
range, direction, actual and reference values. Baseline investigation warnings require a
signal wholly inside the baseline; a subsequent transition MR is not a baseline signal.

Results show the summary, aligned Individuals/MR charts, scrollable deterministic insights
and signal details. Baseline shading/boundaries, blue center, red dashed limits and orange
target are distinct. Chart reference values are separate from axis ticks to avoid overlap.
Content-sized summary and proportional chart/insight rows preserve chart space on resizing.
Targets use inclusive ≤ or ≥ and report observed count/percentage, mean and latest attainment
independently of signal status. No normality assumptions, capability indices, causal claims,
AI explanations or SPC-to-simulation integration are introduced. No signals is not proof of
stability; 20 observations is a practical minimum, not a reliability guarantee.

Explicit analysis writes `SpcAnalysisV1` in the existing very-hidden config sheet; ordinary
Excel Save is still needed for disk persistence. Normal model save/Clear Model preserve this
block, including unreadable blocks. Unreadable SPC settings block SPC opening, not simulation.
`SpcExporter` creates a unique new source-workbook worksheet with configuration/target/limits,
all aligned observations and signal descriptions, and two editable native Excel charts.
User strings are text, never formulas. The first MR stays blank. Native charts highlight
signal endpoints (red circles) and baseline endpoints (green diamonds); the on-screen charts
also highlight participating signal ranges and shade the baseline. Native charts do not
reproduce that background shading. Export failure can leave a partial new report, never an
overwritten source/report sheet.

## Licensing and distribution decisions

Phase 1 intentionally uses offline licensing: a 30-day trial and RSA-signed Professional/
Enterprise licenses in `MC1.<base64url-payload>.<base64url-signature>` form. `LicenseService`
coordinates access/storage, `LocalLicenseValidator` verifies signatures, and `TrialService`
owns trial state. DevelopmentMode is currently **false**; retain this for customer testing/
distribution. Codes are case-sensitive; activation uses normal casing and trims only outer
whitespace.

Trial state is DPAPI CurrentUser-protected in a local file and Registry, with start/last-run
dates, trial ID, rollback checks and conservative reconciliation. Deleting one copy does
not restart the trial. Deactivation removes the paid license without resetting the original
trial. Ordinary uninstall/reinstall must retain state. Offline protection is not tamper-proof
against a determined local user.

Only the public key is embedded in the add-in. The LicenseGenerator, private signing key
and generated customer licenses are developer-only and must never be distributed or
committed. `.gitignore` excludes `**/Keys/`, `**/private-key.pem`, `**/Licenses/`. Keep signing
key backups outside the repository; Excel must never reference the generator/private key.

Installer version is **0.1.0**, publisher Vijay. Inno Setup installs the packed x64 Release
XLL as `MonteCarloForExcel.xll` under LocalAppData/Programs, without admin rights or a desktop
shortcut. Preserve its stable AppId across upgrades. Excel registration uses HKCU
Office/16.0/Excel/Options: reuse our entry or find a free OPEN/OPENn value, never overwrite
another add-in; uninstall verifies ownership before deletion. Installer paths currently
assume `C:\montecarlo`.

Distribute the packed Release XLL, not an unpacked XLL alone, symbols, signing assets or
development files. Customer installation/upgrade/clean-machine validation and code signing
remain release work; generated x86 packaging does not establish live x86 support.

## Verification and working constraints

Current suite: **555 xUnit cases**,
zero failures/skips in Debug and Release; full solution builds and x86/x64 packed-XLL
generation passed in both configurations. Require both configurations for changes to
execution/persistence. The Windows harness passes six existing Results cases and six
Scenario editor/comparison cases (both analysis modes) at 100/125/150% geometry scaling.
Three correlation-editor cases cover add/edit/delete and help/grid/footer bounds at the same scales.
Three SPC configuration/results cases cover rule defaults, both chart paint paths, footer
bounds and resizing at 100/125/150% geometry scaling. SPC adds 51 unit cases covering reference
statistics, fixed baselines, signals, rejected gaps/edge cases, target independence, export
tables and backward-compatible persistence.

Run from repository root as appropriate to the change:

```powershell
dotnet test MonteCarlo.Core.Tests --no-restore
dotnet run --project MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj --no-restore
dotnet build MonteCarlo.Excel/MonteCarlo.Excel.slnx -c Release --no-restore
git diff --check
```

Restore dependencies first on a fresh checkout. Tests cover sampling/fitting/preview,
probability/CDF/quantiles, validation/restoration failures, old/new persistence layouts,
tracked references, highlighting/original fills, edit/delete/recreation, nullable COM
call-site serialization, settings, seeded replay, scenario transformations/comparisons,
explicit adjusted baselines, both decision modes/deltas, V1 migration/V2 question round trips,
correlation validation/PSD matrices, latent dependence and all marginal distributions,
seeded correlated runs, correlation persistence/deletion/reference tracking,
missing references, source isolation, cancellation and owned
resource cleanup. Save/reopen unit tests use in-memory
adapters, not real Excel files or COM.

The separate layout harness compiles production ResultsForm with a workbook double,
creates hidden native handles and stubs Export. Automatic/Fixed cases use 100/125/150%
geometry scaling and matching bitmap DPI. Assertions cover tabs, shared selector/footer,
statistics/calculator bounds, forecast/direction changes, exact target retention, positive
plot/bar geometry, both paint paths and marker placement/labels. These are presence/geometry
checks, not pixel-perfect or per-monitor DPI tests. `RESULTS_LAYOUT_PREVIEW=1` optionally
produces local previews; hidden native input text may not paint there. See its README.

Opt-in real Excel integration check (Windows/Excel required): append `-- --live-scenarios`
to the harness command. It creates only private Excel instances and a synthetic workbook,
verifying correlated inputs, named/cross-sheet formulas, source values/formulas/format/Saved state, one copy
per comparison, success/cancellation/failure cleanup, file deletion and owned process exit.
It passed locally; Excel startup/shutdown latency can vary substantially and the bounded
process-exit assertion can time out before eventual exit. This does not certify arbitrary
customer workbooks or real-monitor DPI.

Release publish directory:
`MonteCarlo.Excel/MonteCarlo.Excel/bin/Release/net10.0-windows/publish/`.
Packed outputs: `MonteCarlo.Excel-AddIn-packed.xll` (x86) and
`MonteCarlo.Excel-AddIn64-packed.xll` (x64). When Excel locks the normal XLL, use an isolated
output, e.g. `-p:ExcelDnaPublishPath=C:\montecarlo\artifacts\validation-build`. Do not delete/
replace a loaded XLL or close Excel merely for validation. Existing warnings: LicenseService
CS0162, AssumptionForm CS8600 and WorkbookPersistence CS8603 (also linked into tests).

Live Excel checks remain necessary for COM errors, protected/read-only failures, formula
restoration, workbook events/switching, save/close/reopen, original-fill restoration,
reference moves/deletions, clear/redefine, seeded replay and Export followed by target
updates. UI verification must include actual display scaling, both charts, close/edge
markers, constant/out-of-range cases, both calculators/directions and native control text.
Automated checks do not establish completion of these live checks.

`-- --live-spc` runs the SPC workbook service in a private Excel STA instance with a synthetic
workbook: strict Excel errors, text/date labels, appended observations with fixed limits,
formula/source preservation, name reuse, native charts, unique exports, text safety, tracked
rename/insert/delete, save/reopen and owned process exit. It does not certify interactive
ribbon clicks/range-picker operation, arbitrary customer workbooks or actual per-monitor DPI.
