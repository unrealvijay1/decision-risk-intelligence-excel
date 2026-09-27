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
