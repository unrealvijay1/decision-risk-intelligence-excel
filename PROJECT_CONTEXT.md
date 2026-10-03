# Monte Carlo for Excel — Project Context

Last reconciled with source and tests: **2026-10-03**.

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

- Ribbon groups: Model Setup (Define Assumption, Define Forecast, Correlations, Model Manager, Clear
  Cell Definition); Simulation (Run Simulation, Scenario Analysis, Optimizer, Simulation Settings, Reload Model);
  Process Analysis (SPC Analysis); Maintenance (Clear Model); Product (License). Export Report is in Results.
- All 13 ribbon commands are native large buttons. Excel owns label wrapping, alignment and
  narrow-window group collapse; there are no fixed-width spacers or custom panels. All existing
  command IDs, callbacks and implementations are preserved. License retains `imageMso=FileProperties`.
- Twelve supplied PNGs in `MonteCarlo.Excel/MonteCarlo.Excel/Icons/` are embedded in the Excel
  assembly. `GetRibbonImage` resolves exact manifest names, clones 32×32 artwork, and fits wider
  images proportionally into a transparent 32×32 canvas without upscaling. New filenames with
  spaces use explicit `LogicalName` metadata. The supplied clear-cell filename is actually
  `Clear Cell Defition.png` (sic); do not substitute the spelling from the request. Packed XLLs
  carry these resources inside the compressed Excel assembly and require no external PNG folder.
- Fourteen distributions. Original enum IDs/order remain Normal=0, Triangular=1, PERT=2,
  Uniform=3, Lognormal=4, Beta=5. Their parameter order and seeded sampling paths are unchanged:
  Normal (mean, SD), Lognormal (log mean, log SD), Uniform (min, max), Triangular/PERT
  (min, most likely, max), Beta (min, max, alpha, beta). Appended IDs 6–13 are Exponential
  (mean), Poisson (lambda), Binomial (trials, probability), Discrete (separate probability
  table), Weibull (shape, scale), Gamma (shape, scale), Bernoulli (probability), Truncated
  Normal (mean, SD, lower, upper). Scalar positions correspond to Parameter1–Parameter4.
- Define/Edit Assumption provides analytical previews of unsaved parameters with inline
  validation. Previewing neither samples nor writes Excel. Historical-data fitting
  supports Normal, Lognormal, Uniform and Triangular, ranked by AIC then KS; fitting
  remains limited to those four families; the form explicitly says so. Discrete uses a
  dedicated outcome/probability grid, Add Row/Delete Row, and a live total. The form adjusts
  its preview/footer to the available client height at scaled or constrained window sizes.
- Model Manager supports viewing, navigation, editing and confirmed deletion.
  `ForecastEditForm` uses detached `ForecastEditDraft` for name, explicit target,
  direction, default P80 or No Target. Cancel/invalid drafts do not apply changes.
  Unchanged target/direction retain requested confidence; changes clear its provenance.
- `ModelDefinitionDeletion` is shared by Model Manager and Clear Cell Definition.
  The Ribbon action accepts one cell only, identifies the definition in confirmation,
  and preserves values/formulas. Ordinary cells are an informational no-op; invalid
  selections are rejected. Lookup loads without highlighting and restores prior
  in-memory lists on cancellation/no-op. Confirmed deletion saves remaining definitions.
- Assumption Correlations supports model-level input relationships. Optimizer manages its own
  controllable decision variables; general model classifications are not implemented.

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
| P (16) | `DistributionDataV1`: per-assumption version-1 JSON `{Version,Outcomes:[{Outcome,Probability}]}` for Discrete only |
| Q:T (17–20), row 1 only | `SimulationSettingsV1`, trial count, seed mode, optional fixed seed |
| V:W (22–23) | V1=`ScenarioDefinitionsV2`, W1=JSON chunk count; V2 onward holds 30,000-character chunks; version V1 remains readable |
| Y:Z (25–26) | Y1=`AssumptionCorrelationsV1`, Z1=chunk count; Y2 onward holds 30,000-character JSON chunks |
| AB:AC (28–29) | AB1=`SpcAnalysisV1`, AC1=chunk count; AB2 onward holds 30,000-character JSON chunks for SPC configuration, tracked ranges and fixed baseline statistics |
| AD:AE (30–31) | AD1=`OptimizerV1`, AE1=chunk count; AD2 onward holds 30,000-character JSON chunks for optimizer configuration only |

The legacy three-parameter/name-in-H and four-parameter/name-in-I layouts remain readable.
Optional metadata is backward compatible. Text metadata is written as text to avoid formula
interpretation. Invalid saved definitions/targets and broken references surface structured
validation errors and block simulation. Event registration/restore failures have safe
messages and Trace diagnostics; initial restore is independent of event registration.
Missing P data is valid for legacy scalar distributions. Discrete requires a valid versioned
table; unsupported/malformed tables block execution with a row-specific correction message.
Up to 100 rows fit in one Excel text cell. Definition saves preserve settings, scenarios,
correlations, SPC and Optimizer blocks; the existing scalar columns retain their meanings.

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

Definition saves preserve the separate settings, scenario, SPC and Optimizer blocks verbatim, even if malformed.
Clear Model restores definition fills and removes definitions, retaining
explicitly saved simulation settings, scenarios, SPC and Optimizer. Without any of those blocks it deletes the config
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

Each execution creates one `Random(actualSeed)` and passes it through all fourteen samplers
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

### Additional distribution mathematics and limits

`ExtendedDistributions` is the authoritative Core implementation for the eight new families;
`DistributionSampler`, `DistributionQuantile`, and `DistributionPreview` delegate to it.
`DistributionCatalog` supplies labels/defaults; immutable `DiscreteTable` sorts paired rows,
requires 2–100 finite unique outcomes and nonnegative probabilities, accepts totals within
1e-10 of one, and normalizes accepted totals. Zero-mass rows never sample. Discrete quantiles
return the smallest outcome reaching the requested cumulative mass, with supported endpoints.
Bernoulli/Binomial p=0/1 are deterministic. New continuous previews show central 99.8% with
an explicit tail note; discrete previews use PMF bars at actual outcome coordinates. Large
count previews show at most 2,001 central integer outcomes with an omission note, never a
continuous approximation. Previewing never consumes RNG or writes workbook state.

Exponential/Weibull use inverse transforms. Poisson uses small-lambda product sampling and
PTRS rejection; Binomial uses beta order-statistic conditioning instead of n draws; Gamma
uses Marsaglia–Tsang with shape boosting. Truncated Normal uses conditional rejection with
tail tilting/narrow-interval envelopes, not clipping. All rejection loops are bounded.
Incomplete-gamma series/continued fractions, incomplete beta, log-space Gamma inverse and
stable Normal interval masses support CDFs/quantiles. Narrow Normal masses use Gaussian
quadrature; same-side tails use log survival probabilities. No normal approximation is
substituted for a count distribution.

Numerical limits are explicit validation errors: Poisson lambda <=1,000,000; Binomial integer
trials 1–1,000,000; Gamma/Weibull shape 0.01–1,000,000 and positive scale, additionally rejecting
shape/scale tails whose estimated logs fall outside [-740,700] (so not every shape in that
interval is supported). Exponential mean is positive and <=Double.MaxValue/1024. Truncated
Normal requires finite standardized bounds within ±10,000 SD and representable interior
values in both original and standardized units. Floating-point resolution still limits
extreme-tail/narrow-bound accuracy. Overflow/nonconvergence stops execution with restoration.

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
nonconvergence/range failures stop the run with restoration. All fourteen distributions are
tested. The eight additions delegate to Core quantiles, clamping normal-CDF probabilities
to [1e-16, BitDecrement(1)] to avoid infinite endpoint samples. Discrete transforms create
ties; observed Pearson coefficients differ from latent coefficients and some final outcome
correlations are unattainable. There is no independent fallback for configured dependence.
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
- Forecast right: large distribution chart; the CDF chart and view selector are removed. Probability calculators remain.
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

Core CDF/quantile mathematics remain available for sampling and probability calculations; Results exposes the distribution chart only.

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
- The eight additions retain their source parameters/table and record a sample scale in
  `ScenarioDistribution`; execution samples the unchanged source then scales once. Count
  or Bernoulli scaling never changes success probability and never rounds scaled outcomes.
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

## Monte Carlo Optimizer V1

The Simulation ribbon's **Optimizer** command uses the supplied `Optimizer.png` through
the existing embedded-resource/32×32 callback. `OptimizerCommand` pins the source workbook,
loads detached model definitions without highlighting, and opens `OptimizerForm`.
Configuration has Objective/Settings, Decision Variables and Constraints tabs with explicit
Add/Edit/Delete, Save Configuration, Run, progress and cooperative Cancel controls.
`DecisionVariableForm` is a compact dedicated editor: editable name, a read-only combined
absolute `Worksheet!$Cell` reference with Select, informational Current Value, and Search
Range Minimum/Maximum/Step.
The Cancel/OK footer has its own content-sized row outside a scrollable editor body,
so larger system fonts or constrained client heights cannot push the actions below the window.
Select uses the existing hidden-form Excel `InputBox(Type: 8)`
pattern, hides both editor and parent Optimizer while picking, and restores them in `finally`.
`DecisionVariableSelection` captures exactly one source-workbook cell, reads numeric values
through the production adapter and rejects blank/text/error/merged/array/spill/protected
inputs before acceptance. Only the immediate left cell supplies a conservative short text
name suggestion, and an entered name is never overwritten. Selection/cancellation performs
no cell or metadata writes. OK reuses existing bounds/steps, duplicate and reserved-cell
validation. Current values need not lie on the step grid. Editing resolves the existing
tracked cell, displays its current value and stored range, and preserves its link through
renames/moves; replacing the cell clears that link so explicit configuration Save creates
one for the replacement. A broken reference requires reselection. Optimizer mathematics,
candidate evaluation and restoration are unchanged by this editor.
`OptimizerConstraintForm` defaults to Excel cell (value/formula), with editable qualified
reference, Select, optional display name, operator and constant RHS. The existing hidden
Excel InputBox pattern and read-only current-value/left-label suggestion are reused.
Quoted sheet names/escaped apostrophes are supported; only source-workbook single A1 cells
within Excel bounds are accepted. Existing variable and forecast-statistic modes remain
selectable/editable. Blank/text/Excel errors/nonfinite, missing/deleted/external and multi-cell
references fail with constraint name/location. Protected, array and spill numeric outputs
are valid because constraints are read-only. Picking/cancelling writes no model cells or names;
editing the same resolved cell retains its tracked link; replacing it creates a new link on Save.
The footer is outside the scrollable body so scaled/constrained layouts retain OK/Cancel.

Cell constraints are validated before mutation, then read afresh immediately after candidate
decision writes and Excel recalculation, before Monte Carlo trials. They evaluate the current
worksheet value with the original/restored assumption values, not a sampled statistic or
last trial. For stochastic requirements, use the existing Forecast Mean/P50/Probability
constraints. Cell/decision/linear constraints reject candidates before simulation; forecast
Mean/P50/Probability constraints use completed trial statistics. Constraint evaluation
never writes constraint cells; a cell may still be written by its separately configured
Decision Variable/Assumption role under the existing engine's restore guarantees.

`Optimization.cs` contains UI-independent problem/objective/variable/constraint/settings,
candidate/result/progress/evaluation records, `IOptimizationAlgorithm`, Automatic search and
`OptimizationService`. Candidate evaluation uses `IOptimizationCandidateEvaluator`: zero assumptions select
`DeterministicCandidateEvaluator`; one or more select `MonteCarloCandidateEvaluator`, which calls
the existing `SimulationExecution`, preserving normal
simulation, correlations, scenarios, SPC and Results behavior.
`OptimizationFeasibleSpace` owns bounded grid generation, deterministic linear repair,
conservative worksheet-formula guides and shared progress text; search does not own sampling.

Objectives retain the existing configured-Forecast workflow: maximize/minimize Mean, P50
or inclusive target probability (≥/≤). A configured objective Forecast is required. The Optimizer
automatically supports deterministic optimization with zero assumptions and stochastic Monte Carlo
optimization with one or more assumptions. Deterministic candidates use one recalculation and one
finite forecast reading: Mean/P50 (and other percentiles) equal that value, SD is zero and samples
contain one value. No simulation engine/trial loop or fake assumption is used. Probability-specific
objectives/constraints are rejected before mutation; the deterministic editor omits those choices,
and result/export target probabilities are unavailable. Normal Run Simulation still requires an
assumption; its validation is unchanged. Monte Carlo probability objectives and requirements
use **0–100 percent**, not fractions. Optimizer constraints may reference any valid numeric
Excel cell, including calculated/formula cells. Constraint cells are recalculated and
evaluated for every candidate solution and do not need to be configured as Forecasts.
Existing decision-variable and forecast Mean/P50/Probability constraints remain supported.
The editor also accepts explicit linear decision expressions (e.g. `2*X + Y`) and numeric RHS,
with <=, >= or =. Persisted terms use stable variable IDs and finite nonzero coefficients.
Use unique names without expression operators in the UI; advanced callers can use stable IDs.
Decimal and scientific-notation coefficients round-trip through expression editing.
Generation tightens decision-variable bounds and projects linear residuals onto permitted
grid points in at most 32 repair passes. It supports simple weighted equalities, inequalities
and independent equalities; coupled/discrete systems are not a general integer/LP solver.
Raw constraints referencing a decision cell or simple same-sheet A1 additions/subtractions,
numeric coefficient products, or a whole `SUM` of direct decision cells/single-column ranges
provide equivalent run-local linear guides. No portfolio/sum-to-one rule is hard-coded.
Unsupported formulas, unknown inputs, nested functions, constants and indirect/external
references stay black-box checks. Every candidate is verified against actual recalculated
cells before simulation, even if a guide was used. Grid repair may fail to find an existing
feasible solution for difficult systems; never imply that failed search proves infeasibility.
All constraint kinds accept <=, >= or = against a finite numeric RHS. Inequalities retain
exact inclusive comparisons; equality accepts abs(actual-RHS) <= 1e-9 * max(1, abs(RHS)).
Feasible candidates outrank infeasible ones; infeasible candidates compare
sum of positive violations normalized by `max(1, abs(requirement))`, then objective. Ties retain
the earlier candidate. Best Feasible is separate from the diagnostic infeasible candidate.
No feasible result shows no optimum/objective; Apply is disabled. Rejected candidates have
no objective/distribution (history Objective is NaN, Evaluated=false); Monte Carlo never runs
for them. Diagnostics identify failed constraints, attempts and pre-simulation rejections.

Variables are finite numeric single-cell inputs on a grid anchored at Minimum with positive
Step and Minimum < Maximum. Maximum is included only when step-aligned (floating-point
rounding is bounded). V1 supports up to 100 variables and 10¹² steps per variable; Step must
fit the range and produce distinguishable values. Duplicate physical cells, assumption/
forecast overlaps, merged/array/spill cells and locked protected inputs are rejected before
mutation. Numeric formula inputs are allowed and their exact formula content is captured.
Decision constraints with no permitted grid point are rejected; statistical infeasibility
is determined by candidate evaluation rather than presumed from the workbook formula.

Automatic exhausts small lattices (after repair/deduplication). Larger problems initialize
from the original point, constrained corners and diverse seeded feasible-grid attempts.
It retains six lightweight elites using feasibility-first ranking. Most subsequent attempts
poll the current best; some poll another elite. Coordinate radii sweep from coarse to one
grid step, repeatedly refining around updated solutions; every eighth attempt explores
globally. Linear repair creates compensating changes in other variables when required.
It is derivative-free, noisy-objective compatible, seeded/reproducible and cancellable.
Stop on the unique-candidate budget, exhausted small lattice, cancellation, error, or
40 × budget generation attempts (duplicates/repair can exhaust the space earlier).
This is a bounded heuristic, not a global-optimality guarantee. Defaults are
**250 candidate evaluations, 1,000 trials, seed 12345**;
accepted settings are 1–100,000 evaluations and, in Monte Carlo mode, 1–1,000,000 trials.
The UI says Maximum Candidate Evaluations and explains that decisions and uncertainty are
separate loops: 250 candidates × 1,000 trials is at most 250,000 Monte Carlo trials.
The pre-run computational estimate updates as settings change; rejected candidates lower
actual trial consumption. Five evaluations still means at most five combinations, not 5,000.
The UI shows Evaluation Mode. Deterministic trials display N/A; retained trial settings are not used.
The search seed stays editable for deterministic exploration; Monte Carlo shows trials and seed.
Mode is detected from current definitions, not persisted; the OptimizerV1 envelope is unchanged.
The candidate-search PRNG is separate from simulation. Every evaluation restarts the
production Monte Carlo random stream at the same fixed seed in stochastic mode, including correlations.
Assumptions are sampled from their configured distributions on every trial. Fixed seeds
reuse common random draws across candidates; sampled inputs and calculated outputs vary
within each evaluation. Structured run history stores candidate values/scores, feasibility,
constraint actuals, rejection/evaluation flag, best-feasible-so-far, search phase and radius;
full trial samples are retained for the best candidate only. Run results are not persisted.
Deterministic live previews show the single calculated forecast per candidate while convergence
continues by evaluation; ScreenUpdating is enabled for those previews and restored afterward.

The interactive Optimizer runs synchronously on Excel's STA against the pinned source
workbook, temporarily writing decision and assumption cells. Excel interaction and events
are disabled while running; the live progress window pumps UI messages at bounded preview
updates and supports cancellation. Every decision restoration is attempted in `finally`;
in stochastic mode, the simulation engine independently restores assumptions, recalculates forecasts and
restores ScreenUpdating/events/status between evaluations. Outer interaction/event settings
restore on success, failure and cancellation. Restoration failures are reported. Source
formulas/values restore but Excel may remain dirty; the Saved flag is not forced. Apply
Optimal Values is the explicit action that retains optimized decisions after the run.
No temporary workbook or worker Excel process is used by the interactive Optimizer.
Scenario Analysis keeps its private sandbox. OptimizationService retains its sandbox factory
for noninteractive callers/tests. An in-flight Excel COM calculation can delay cancellation.
Workbook VBA events are suppressed, but volatile formulas, UDF side effects and external
refreshes are outside deterministic replay/restoration guarantees.

LiveRunForm displays selectable forecast histograms/frequency curves during both normal
simulation and optimization, plus a best-objective step chart by evaluation for Optimizer.
Targets use saved forecast values for mean/median (legacy P80 defaults use a fixed initial-evaluation P80 reference); probability objectives use a matching
saved requested confidence (percent), never a forecast value on a percent axis. Explicit No Target has no target line. The best line follows feasibility-first
selection; no infeasible objective is plotted as an optimum. The scrollable progress summary
shows current candidate/objective, best feasible candidate/objective, constraint status,
attempts, trials per candidate, feasible count, rejection count and completed MC evaluations.
SimulationExecution's optional
observer receives detached, evenly spaced previews of at most 4,096 completed samples per
forecast on first/last trials and roughly every 100 ms. Previewing consumes no RNG; default
noninteractive execution still suppresses ScreenUpdating. Interactive runs enable worksheet
redraw, and completed trial values/formulas are visible behind the progress window. Curves
represent empirical results, not an assumed Normal fit. Excel editing is blocked until
restoration completes; move the progress window aside to see the worksheet cells.

`OptimizerPersistence` stores the independent chunked `OptimizerV1` envelope in AD:AE.
Cell (ID 4) and Linear (ID 5) are appended kinds; optional Operator defaults to FromDirection, preserving
historical V1 JSON and direction semantics without migration. Linear terms are optional in
old JSON and retained in the same envelope for new constraints. Cell constraints persist their
source-workbook sheet/cell, friendly name, RHS/operator and hidden `_MC_Cell_` reference in
the same envelope. Tracking follows sheet rename/row-column movement; deleted/foreign links
become `#REF!` and never silently rebind. Normal SaveModel and Clear Model preserve even
unreadable optimizer blocks. Decision cells
use hidden `_MC_Cell_` workbook names; rename/move resolves current addresses, broken names
become `#REF!` without stale fallback. Forecast objectives/constraints use existing tracked
forecast identities; legacy forecast coordinates upgrade on optimizer save. Unknown/invalid
envelopes block Optimizer only. Normal Excel Save is required for disk persistence.

`OptimizerResultsForm` shows the best observed objective, evaluations/trials/seed/duration
and separate Decision Variables, Constraints and Forecast Results tabs (Mean/P50/P80/SD/
applicable target probability). **Apply Optimal Values** retains optimized decision values after temporary live-run restoration;
it re-resolves tracked decisions, rejects changed values or formulas, validates all cells
before writing, suppresses source events and rolls back captured contents on failure.
Apply also recalculates and rechecks raw cell constraints; changed capacities/invalid formulas
or an infeasible candidate trigger rollback and require rerunning Optimizer.
Close/Cancel/Export do not apply values. `OptimizerExporter` creates a unique report worksheet
with objective, original/optimized/change, bounds/steps, constraints/results/statistics and
settings; strings are text, numbers remain numeric. Export failure may leave a partial new
report, never an overwritten existing sheet. Source read-only workbooks cannot save, run live optimization or apply.

Future extension points are new `IOptimizationAlgorithm` implementations and richer domain
variable/objective records. Integer/binary/categorical types, multiple objectives/Pareto,
sensitivity near the optimum, robust objectives and scenario-specific
optimization are deferred. Numeric step=1 can express an integer-valued grid but V1 has no
separate typed-variable UI. Volatile Excel functions can weaken common-random-number
comparability; stochastic estimates need independent confirmation for important decisions.

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

**Wheeler's Four Tests** are independently selectable: (1) strictly outside control limits
on both charts, (2) eight Individuals observations strictly above/below center, (3) at least
two of three strictly beyond the same 2σ threshold, and (4) at least four of five strictly
beyond the same 1σ threshold. **Supplementary Rules** contains the six-point strictly
increasing/decreasing trend (rule ID 5; previously 3). All five default to enabled. Sigma is
saved baseline MR̄ / **1.128**; zone thresholds are saved mean ± sigma and ± 2 × sigma.
These derived properties are not separately persisted and never replace the 2.66/3.268 limits.
Zone tests apply only to Individuals; the remaining window observation can be anywhere.
Equality does not qualify. Missing/non-finite data remain rejected before detection.

All observations are evaluated chronologically. Run/trend episodes report their first
qualifying endpoint only. Zone episodes report the first qualifying window per rule and
direction; consecutive qualifying windows are suppressed until a nonqualifying window
resets that episode. Different rules never suppress each other. Signals retain one-based
window bounds, endpoint, direction, threshold, qualifying indices and actual window values.
Highlights/export membership use qualifying observations, which need not include the window
endpoint. Baseline investigation requires a signal wholly inside the baseline; a subsequent
transition MR is not a baseline signal. Insights describe rule/direction and multiple signals
without inferring operational causes.

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
The additive `TwoOfThree`/`FourOfFive` rule fields retain the `SpcAnalysisV1` envelope;
missing fields default to true while all three existing selections and saved limits survive.
`SpcExporter` creates a unique new source-workbook worksheet with configuration/target/limits,
all aligned observations and signal descriptions, and two editable native Excel charts.
User strings are text, never formulas. The first MR stays blank. Native charts highlight
qualifying signal observations (red rings) and baseline endpoints (green diamonds); the
on-screen charts keep blue data points visible inside red rings and shade the baseline.
The scrollable details grid retains every rule, qualifying indices, threshold and actual
window values; exports retain all applicable identifiers/descriptions per participating
observation and include four zone columns (S:V). Native charts do not
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

Current suite: **899 xUnit cases** (including 36 constraint-aware iterative-search regressions),
zero failures/skips in Debug and Release; full solution builds and x86/x64 packed-XLL
generation passed in both configurations. Require both configurations for changes to
execution/persistence. The Windows harness passes six existing Results cases and six
Scenario editor/comparison cases (both analysis modes) at 100/125/150% geometry scaling.
Three correlation-editor cases cover add/edit/delete and help/grid/footer bounds at the same scales.
The distribution harness adds 42 form cases (14 families at each scale), covering selector,
parameter visibility/labels, valid previews, table add/delete/total errors, editing/saving,
and control bounds. Rendered Discrete/Truncated Normal previews are in
`artifacts/distribution-layout/`. Distribution tests include one million samples each for
large Poisson/Binomial/Gamma parameters, analytic references, tails, singular/mixed
correlations, common-seed scenarios with a selected adjusted baseline, disk-format parsing,
tracked reference/edit/delete round trips and inclusive target mass with tied sensitivities.
Three SPC configuration/results cases cover five rule defaults/selections, scroll access,
control overlap, multi-rule details, both chart paint paths, footer bounds and resizing at
100/125/150% geometry scaling. SPC has 78 unit cases covering reference statistics, fixed
baselines/zones, strict same-side rules, episode deduplication, rejected gaps/edge cases,
target independence, multi-rule export tables and legacy/five-rule persistence. Debug and
Release layouts and rendered previews were checked; private Excel SPC and Scenario checks
passed, including owned process exit. Current optimizer packages are in
`artifacts/iterative-optimizer-debug/` and `artifacts/iterative-optimizer-release/` (the previous Release
x64 XLL can remain loaded; do not overwrite it or close Excel merely to build).
All four XLLs were extracted and verified: correct x86/x64 PE headers, all 12 PNG hashes,
transparent 32×32 image callback results, five ordered groups and 13 callback mappings
(the previous 12 unchanged, plus Optimizer);
packed Core/Excel SHA-256 hashes match their final compiled assemblies.
A private Excel instance previously accepted the ribbon Release x64 XLL. Native ribbon interaction, narrow-window
collapse and actual 100/125/150% Windows DPI rendering remain manual checks; no approved mockup
image was supplied, so the implementation follows the requested textual layout.

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
creates native handles and links the production report exporter. Automatic/Fixed cases use 100/125/150%
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

Optimizer engine/persistence tests cover all objective directions, inclusive probability/statistic/
decision constraints, known single/multiple/interior optima, bounds/steps, feasible-first
ranking, per-candidate common streams, seed replay, cancellation (including the last
evaluation), source apply validation/rollback, restoration failures, report data and tracked
configuration save/reopen/move/delete/legacy upgrade. Debug and Release full layout harnesses
pass at 100/125/150% geometry scaling, including three optimizer cases and rendered result
previews under `artifacts/optimizer-layout/`. `-- --optimizer` runs only optimizer layouts;
`-- --live-optimizer` runs synthetic private Excel source/sandbox integration checks for
source formulas/value/format/Saved preservation for sandbox callers, one sandbox per run, success/replay/cancel/
failure, temporary-file cleanup, explicit apply, report export and disk save/reopen with
tracked rename/delete and owned process exit. It also verifies actual source trial values, forecast recalculation and restoration after success/cancellation/chart failure using the production COM adapter. LiveRunLayoutChecks renders both progress charts at 100/125/150% including constant/extreme samples. The synthetic source fixture runs in its own
STA apartment, which ends before process-exit assertions. Actual-monitor DPI, interactive ribbon/range editing, loaded x86
Excel and arbitrary customer models remain manual checks.
The decision-variable editor adds 26 selection cases (absolute coordinates/current values,
left-label suggestion/name retention, invalid/unsafe/foreign/multiple cells, cancellation
and no configuration writes). Native form checks also cover new/edit, off-grid current
values, picker cancellation, name suggestion and tracked link retention after rename.
Decision-variable previews are under `artifacts/optimizer-layout/`. Interactive Excel
InputBox picking/cancellation, workbook switching and actual display DPI remain manual checks.

`-- --constraints` exercises the dedicated constraint editor at 100/125/150% with typed/picked
cells, cancellation, equality, name/link retention, legacy kinds and resized/larger-font footer
checks. Previews are under `artifacts/constraint-layout/`; it is also part of the full layout run.
`-- --live-cell-constraints` uses an owned private Excel STA and real product-mix formulas:
Qty A=80, Qty B=60 gives profit=88,000, labor=400 and material=360. Only profit is a Forecast;
protected formula constraints stay read-only. Success/cancel/chart failure and candidate
formula errors preserve input/output formulas; Apply, disk save/reopen and named-reference
rename/row insertion/deletion pass with owned process exit. The 44 unit cases include legacy
V1 JSON with missing optional fields, equality tolerance, invalid references/values/Excel
errors, recalculation, mixed constraint kinds, Apply rollback and persistence. Interactive
Excel picking/workbook switching, live x86 and actual-monitor DPI remain manual checks.
The same private Excel check runs a zero-assumption price/demand/revenue workbook and finds
Price=100, Demand=500, Revenue=50,000; single-sample live updates, success/cancel/preview-failure
formula restoration, disk reopen and Apply pass. The 23 deterministic-mode regressions cover
single/multiple variables, maximize/minimize Mean/P50, constraints, probability rejection,
restoration/cancellation/Apply, persistence and unchanged normal simulation validation.
Optimizer layout checks cover both modes, deterministic probability-option omission,
linear-expression editing, candidate/trial help, computational estimates and progress counters.
`-- --live-iterative-optimizer` uses an owned private Excel workbook with three Normal returns
(means .08/.12/.07, SD .02), capital 10,000 and three allocation decisions on a .01 grid.
Worksheet SUM equality and shares-cell cap guide generation. Unit acceptance at 250 × 1,000
finds 0/1/0 (mean near 1,200) and .4/.6/0 (near 1,040); the seeded deterministic interior fixture
finds 37/83 with objective zero. Tests verify CRN across all three underlying return draws,
pre-simulation rejection, weighted/independent equalities, strict inequalities, tolerance,
grid/bounds, elite/global/refinement history, no-feasible diagnostics and linear persistence.
Real Excel acceptance finds the same allocations: uncapped mean 1,199.957 and capped mean
1,041.920 (seed42, 1,000 trials). Benchmark source restoration, cancellation, preview failure,
five pre-simulation rejections with zero MC runs, disk save/reopen, Apply and owned process
exit pass without touching customer instances.
The live harness also measures 50/250/500 × 1,000 against the production Excel COM adapter
with manual calculation and no rendering overhead, recording CSV in
`artifacts/iterative-optimizer-benchmarks.csv`. Excel access/recalculation dominates runtime;
timings depend on workbook, machine and calculation mode. Customer live rendering can cost more.
Current-candidate notification precedes simulation and remains visible alongside a labelled
partial-sample objective preview during trial updates. Preview failure/cancellation follows
the same restoration boundary and consumes no search/sampling RNG.

`-- --live-distributions` uses a synthetic .xlsx and private Excel STA instance: all eight
new families save/reopen from disk, independent fixed-seed runs replay, mixed
Exponential/Poisson correlations run through production loading/execution, all eight
forecasts/sensitivities are produced, input formulas/number formats restore, and report
parameter descriptions contain the new parameters/table. The owned Excel process exits;
user instances are not used. Full report export interaction, live x86 loading, actual
per-monitor DPI and arbitrary customer workbooks remain manual checks. The harness's
scenario unit tests cover all eight families; live distribution checking is not a full
interactive scenario/report certification.

`-- --live-spc` runs the SPC workbook service in a private Excel STA instance with a synthetic
workbook: strict Excel errors, text/date labels, appended observations with fixed limits/zones,
formula/source preservation, name reuse, native charts, unique exports, text safety, tracked
rename/insert/delete, zone signals/multiple identifiers and thresholds in export, save/reopen
and owned process exit. It does not certify interactive
ribbon clicks/range-picker operation, arbitrary customer workbooks or actual per-monitor DPI.
