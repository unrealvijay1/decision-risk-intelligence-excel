# Results tab layout and rendering checks

Run on Windows:

    dotnet run --project MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj

This hidden WinForms harness compiles the production ResultsForm source against the existing workbook double. It creates native control handles but never opens or modifies Excel. Export is stubbed and not executed.

Six cases cover Automatic and Fixed results at 100%, 125%, and 150% geometry scaling with matching bitmap DPI. Assertions cover:

- Exactly three primary tabs, Forecast default, one shared forecast selector and shared footer.
- P50/P80 in Statistics, probability emphasis in Forecast, and sensitivity ownership/space.
- Nonempty samples/bins, positive plot/bar dimensions, in-range marker positions.
- Production Distribution paint output (CDF chart removed; probability calculators retained), with chart and marker presence checks rather than pixel-perfect snapshots.
- All distribution statistics and run metadata, long seeds, label bounds and footer separation.
- Forecast switching, direction changes, target/confidence calculator handlers and numeric marker labels.

Set RESULTS_LAYOUT_PREVIEW=1 to save an optional local DrawToBitmap preview under artifacts/results-tabs-preview. Hidden native controls may not paint their text in that preview.

These checks do not replace live Excel verification or actual per-monitor DPI testing. Desktop working-area limits intentionally remain in effect to expose space starvation under scaling.

Optimizer adds three cases at 100%, 125% and 150% geometry scaling, covering six objective
options, target activation, configuration save, variable/constraint grids, all result tabs,
export/close callbacks and editor/footer bounds. Previews are written to
`artifacts/optimizer-layout/`. Append `-- --optimizer` to run just these cases.
Append `-- --live-optimizer` for opt-in private Excel integration: one sandbox per run,
source formulas/values/formats/Saved preservation, fixed-seed replay, cancellation/failure,
copy deletion, explicit apply, report export, actual disk save/reopen and tracked rename/delete.
The synthetic source runs in a dedicated STA apartment that ends before process-exit
assertions; all owned Excel instances must exit. Existing user instances are never used.

Forecast refinement checks also verify the 34/66 two-column workspace, calculator terminology and result bounds, and exact target retention when formatted text is recalculated. Success/Miss labels are integrated into chart regions; disconnected marker leader ticks are removed.

Visual-polish assertions verify the nonduplicated probability caption and production marker-label bounds at left, middle, and right anchors, including coincident markers in separate rows. Chart dimensions remain unchanged.

Scenario Analysis adds six cases (both analysis modes) at 100%, 125%, and 150% scaling,
exercising explicit baseline selection, independent forecast probabilities and mode switching with retained questions, adjustment
save/remove handlers, six decision columns, neutral baseline deltas, forecast context,
and grid/footer bounds. The comparison supports horizontal scrolling.

Assumption Correlations adds three cases at the same scales, checking add/edit/delete,
the three-column configuration grid, explanatory help and footer bounds.
Append `-- --correlation-performance` for a sampling-only timing comparison of independent
and correlated Normal/PERT/Beta pairs (50,000 trials, no Excel recalculation).

Optional real Excel scenario integration checks (separate from the default layout run):

    dotnet run --project MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj --no-restore -- --live-scenarios

This opt-in path creates private hidden Excel instances and a synthetic workbook. It
does not attach to existing Excel processes. It exercises the production copy/sandbox
and engine paths, correlated assumptions, an adjusted user-selected baseline, named ranges, cross-sheet formulas, source values/formulas/format/Saved
state, one copy per comparison, completion/cancellation/failure, temporary-file deletion
and owned process exit. It requires Windows and an installed Excel; it is not part of
the platform-independent xUnit count or a substitute for customer-workbook testing.
# SPC validation

The default run includes SPC configuration/results checks at 100%, 125% and 150%
geometry scaling, both chart paint paths and resize/footer checks. Set
`RESULTS_LAYOUT_PREVIEW=1` to save SPC form and chart previews alongside the executable.
Hidden native controls may omit their text in bitmap previews; these checks do not
replace actual display-DPI and interactive range-picker testing.

Run `dotnet run --project MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj --no-restore -- --live-spc`
from the repository root for a private Excel instance and synthetic workbook. It checks
the workbook analysis service, frozen limits after extending the source, strict errors,
text/date labels, source formulas, tracked references, native chart export, save/reopen
and process cleanup. It never attaches to an existing Excel process.

LiveRunLayoutChecks verifies simulation/optimizer progress charts at 100/125/150% scale, target lines, constant/extreme values and saves previews under artifacts/live-progress-layout. --live-optimizer also checks real source trial writes and restoration after success/cancellation/chart failure.

Use -- --constraints for the new constraint editor alone; the default run also includes it.
Checks cover typed/picked cells, cancellation, equality, tracked name retention, all legacy
constraint types and footer bounds at 100/125/150%, including larger-font/constrained-height
cases. Previews are under artifacts/constraint-layout.

Use -- --live-cell-constraints for a separate owned Excel instance with real product-mix
formulas. It verifies the known 80/60 optimum (profit 88,000, labor 400, material 360),
protected unregistered formula constraints, candidate recalculation/errors, cancellation,
restoration, explicit Apply, disk save/reopen and tracked rename/row insert/delete. User
Excel instances are untouched. Interactive picking and actual-monitor DPI remain manual.
The same check also runs a zero-assumption price/demand/revenue workbook: Price=100,
Demand=500, Revenue=50,000, one forecast sample per candidate, live recalculation,
success/cancel/preview-failure formula restoration, disk reopen and explicit Apply.
Optimizer layouts include both detected modes, deterministic N/A trials, four value
objectives and omission of probability constraints at 100/125/150% geometry scaling.
They also cover the linear expression editor, candidate/trial terminology and computation estimate.

Use `-- --live-iterative-optimizer` for opt-in real Excel portfolio acceptance and performance.
It uses an owned private instance, three Normal(.08/.12/.07, SD .02) inputs, capital 10,000,
allocation step .01, worksheet `SUM` equality and a shares-cell cap. It runs the real engine
at 50/250/500 candidates × 1,000 trials and writes `artifacts/iterative-optimizer-benchmarks.csv`.
The fixture uses manual calculation and no chart rendering, so timing is explicitly for
that configuration, not arbitrary customer models. This check can take a long time because
every trial writes three Excel cells and recalculates. It also verifies cancellation,
failure/rejection restoration, disk reopen and Apply; it never touches user Excel instances.
