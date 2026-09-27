# Results tab layout and rendering checks

Run on Windows:

    dotnet run --project MonteCarlo.Results.LayoutChecks/MonteCarlo.Results.LayoutChecks.csproj

This hidden WinForms harness compiles the production ResultsForm source against the existing workbook double. It creates native control handles but never opens or modifies Excel. Export is stubbed and not executed.

Six cases cover Automatic and Fixed results at 100%, 125%, and 150% geometry scaling with matching bitmap DPI. Assertions cover:

- Exactly three primary tabs, Forecast default, one shared forecast selector and shared footer.
- P50/P80 in Statistics, probability emphasis in Forecast, and sensitivity ownership/space.
- Nonempty samples/bins, positive plot/bar dimensions, in-range marker positions.
- Production Distribution and Cumulative paint output, with chart and marker presence checks rather than pixel-perfect snapshots.
- All distribution statistics and run metadata, long seeds, label bounds and footer separation.
- Forecast switching, direction changes, target/confidence calculator handlers and numeric marker labels.

Set RESULTS_LAYOUT_PREVIEW=1 to save an optional local DrawToBitmap preview under artifacts/results-tabs-preview. Hidden native controls may not paint their text in that preview.

These checks do not replace live Excel verification or actual per-monitor DPI testing. Desktop working-area limits intentionally remain in effect to expose space starvation under scaling.

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
