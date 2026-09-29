# Ribbon redesign implementation

## Final layout

All commands are native large buttons, in this order:

1. **Model Setup:** Define Assumption → Define Forecast → Correlations → Model Manager → Clear Cell Definition.
2. **Simulation:** Run Simulation → Scenario Analysis → Simulation Settings → Reload Model.
3. **Process Analysis:** SPC Analysis.
4. **Maintenance:** Clear Model.
5. **Product:** License.

Excel retains control of fonts, backgrounds, separators, natural label wrapping and group
collapse at reduced widths. No spacers, artificial label padding or fixed-width containers
were introduced. All 12 command IDs, callbacks and command implementations are unchanged.

## Icon mapping

All PNG paths below are relative to `MonteCarlo.Excel/MonteCarlo.Excel/Icons/`.

| Command | Resource |
| --- | --- |
| Define Assumption | `assumption.png` |
| Define Forecast | `forecast.png` |
| Correlations | `Correlations.png` |
| Model Manager | `model_manager.png` |
| Clear Cell Definition | `Clear Cell Defition.png` |
| Run Simulation | `run_simulation.png` |
| Scenario Analysis | `Scenario Analysis.png` |
| Simulation Settings | `Simulation Settings.png` |
| Reload Model | `reload_model.png` |
| SPC Analysis | `SPC Analysis.png` |
| Clear Model | `clear_model.png` |
| License | Existing Office `FileProperties` icon |

The clear-cell filename retains the supplied spelling. Four supplied images are wider than
32 pixels: clear-cell 37×32, correlations 38×32, settings 36×32 and SPC 48×32. The loader fits
these proportionally into transparent 32×32 canvases using high-quality downsampling. Other
images remain at their native 32×32 size. No source PNGs were modified or regenerated.

`GetRibbonImage` retains the existing manifest-stream approach. The project embeds all 11
PNGs in the Excel assembly, with explicit logical names for new files containing spaces.
Excel-DNA packs that assembly into both XLL architectures; external PNG files are unnecessary.

## Files changed in this task

- `MonteCarlo.Excel/MonteCarlo.Excel/MonteCarloRibbon.cs`: XML layout and image mappings/sizing only.
- `MonteCarlo.Excel/MonteCarlo.Excel/MonteCarlo.Excel.csproj`: embed the five additional PNGs.
- Five user-supplied PNGs are present as untracked source assets: `Correlations.png`,
  `Clear Cell Defition.png`, `Scenario Analysis.png`, `Simulation Settings.png`, `SPC Analysis.png`.
- `PROJECT_CONTEXT.md`: current ribbon/resource/validation handoff.
- `RIBBON_UI_IMPLEMENTATION.md`: this report.

Earlier uncommitted SPC changes were preserved. No analytical, persistence, settings or
licensing implementation changed in this task. No commit or push was made.

## Validation

| Check | Result |
| --- | --- |
| Baseline | 582 tests passed |
| Final Debug suite | 582 passed, zero failed/skipped |
| Final Release suite | 582 passed, zero failed/skipped |
| Complete Debug and Release builds | Passed; only existing warning locations |
| Packed x86/x64 XLL generation | Passed in both configurations |
| Packed-resource verification | Extracted the Excel assembly from all four XLLs; all 11 embedded PNG SHA-256 hashes match supplied files |
| Actual image callbacks | All 11 render visible transparent 32×32 bitmaps from extracted in-memory assemblies, without external image loading |
| XML/callback validation | Five ordered groups, 12 large controls, original action mappings and method bodies, existing License icon |
| Live Excel | Private Excel accepted the Release x64 packed XLL; blank workbook closed without saving |

Packages: `artifacts/ribbon-debug/` and `artifacts/ribbon-release/`, each containing
`MonteCarlo.Excel-AddIn-packed.xll` (x86) and `MonteCarlo.Excel-AddIn64-packed.xll` (x64).
The local resource verification script is `artifacts/verify-ribbon.ps1`.
Existing warnings remain LicenseService CS0162, AssumptionForm CS8600 and WorkbookPersistence CS8603.

## Deviations and remaining manual validation

No approved mockup image was included or found in the repository; the explicit textual
layout was implemented. Native Excel controls determine exact spacing and label wrapping.
The four non-square supplied images are proportionally reduced rather than stretched.

Interactive ribbon behavior, actual Windows display scaling at 100%, 125% and 150%,
wide/narrow Excel windows and collapsed-group command access have not been inspected.
Live x86 Excel was not tested. Check these manually before distributing the redesign.
