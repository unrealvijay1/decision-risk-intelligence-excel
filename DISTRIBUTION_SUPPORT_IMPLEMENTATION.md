# Phase 1 distribution support

Implemented Exponential, Poisson, Binomial, custom Discrete, Weibull, Gamma, Bernoulli,
and Truncated Normal throughout the applicable add-in workflows. The original six
identifiers, parameter order, fitting algorithms and independent sampling paths remain
unchanged. No commit or push was performed.

## Implementation

| Area | Files and behavior |
| --- | --- |
| Core | `ExtendedDistributions.cs` implements validation, PDF/PMF, CDF, quantiles and bounded sampling; `DiscreteTable.cs` validates immutable paired rows; `DistributionCatalog.cs` supplies names, labels and defaults. Existing sampler, preview and quantile entry points delegate to these implementations. |
| Assumption UI | `AssumptionForm.cs` exposes all 14 families, relevant fields, inline analytical previews, a dedicated Discrete grid with add/delete/total validation, and an explicit four-family fitting note. PMFs use bars at actual outcome coordinates. Footer/preview sizing accommodates a constrained scaled window. |
| Model integration | `SimulationModel.cs`, `SimulationValidation.cs`, `SimulationExecution.cs`, Ribbon, Model Manager and report descriptions carry the immutable table and added enum values. Existing Results statistics, percentile interpolation, target inclusivity and tied-rank sensitivity remain unchanged. |
| Correlations | `DistributionQuantile.cs` supports all eight marginals, including count/table quantiles and safe finite normal-CDF endpoints. Configured latent Gaussian dependence retains PSD validation and never falls back to independent sampling. Discrete ties can change observed Pearson correlation; final correlations are not guaranteed or calibrated. |
| Scenarios | `ScenarioDefinition.cs` and `ScenarioSimulationService.cs` preserve source parameters/table and apply the percentage to sampled values once, including noninteger scaled counts. Common draws, selected adjusted baselines and -100% zero outcomes remain intact. |
| Persistence | `DistributionDataPersistence.cs` and `WorkbookPersistence.cs` use column P, header `DistributionDataV1`, with versioned JSON for Discrete. The four scalar columns, original cell fills, tracked references, settings, scenarios, correlations and SPC blocks retain their meanings. Old workbooks remain readable; invalid table data blocks execution with a correction message. Excel Save remains required for disk persistence. |

Parameter order is Exponential(mean), Poisson(lambda), Binomial(trials, probability),
Weibull(shape, scale), Gamma(shape, scale), Bernoulli(probability), and Truncated Normal
(mean, SD, lower, upper). Discrete stores 2–100 finite unique outcomes and nonnegative
probabilities, normalizing totals accepted within 1e-10 of one. It supports negative,
noninteger and zero-mass rows and preserves pairs when sorting.

Poisson uses PTRS for large lambda; Binomial conditions on beta order statistics;
Gamma uses Marsaglia–Tsang. Truncated Normal samples the conditional distribution with
tail/narrow-interval methods, never clipping. Mathematical references include
[NumPy's distribution implementations](https://raw.githubusercontent.com/numpy/numpy/main/numpy/random/src/distributions/distributions.c)
and [NIST's incomplete-gamma continued fractions](https://dlmf.nist.gov/8.9).

## Verification

| Check | Result |
| --- | --- |
| Baseline | 582 xUnit cases passed before implementation. |
| Final Debug / Release | 691 passed in each configuration, zero failures/skips; 109 added cases. |
| Full solution builds | Debug and Release succeeded. Existing nullable/licensing warnings remain on clean compilation. |
| Numerical/performance coverage | Analytical references, inverse definitions, endpoints, rare events, narrow/extreme truncated tails, moments and deterministic cases. One million large-parameter independent samples each for Poisson, Binomial and Gamma pass the 45-second per-case guard. |
| Integration coverage | All eight families exercise seeded execution, restoration, singular/mixed correlations, common-draw scenarios, selected baseline, zero scale, unrounded outcomes, save/reload, reference moves, editing and deletion. Discrete result tests check inclusive target mass and tied sensitivities. |
| Layout | Full Debug/Release harness passed at 100/125/150% geometry scaling: 42 added distribution form cases plus existing Results, scenario, correlation and SPC checks. Discrete/Truncated Normal renders were inspected. |
| Live Excel | Private synthetic .xlsx: all eight disk round trips, seeded replay, mixed Exponential/Poisson correlation, eight forecast results, formula/number-format restoration and report parameter descriptions passed. Private STA Excel exited successfully; user Excel instances were not used. |
| Packed XLLs | Debug/Release x86/x64 generated in `artifacts/distributions-debug/` and `artifacts/distributions-release/`. PE architecture verified; compressed Core and Excel assemblies extracted and SHA-256 matched to final builds. All 11 embedded PNGs, image callbacks, 12 Ribbon actions and group order also verified. |
| Repository | `git diff --check` passed; `PROJECT_CONTEXT.md` reconciled with implementation and test counts. |

## Limits and manual verification

Poisson lambda and Binomial trials are limited to 1,000,000. Gamma/Weibull shapes must
be between 0.01 and 1,000,000, with additional shape/scale tail-range validation; some
combinations within those nominal bounds are rejected. Exponential mean must not exceed
Double.MaxValue/1024. Truncated Normal bounds must contain representable interior values
and be within ±10,000 standardized units. Extreme floating-point resolution and iterative
quantile cost still apply; the million-trial performance tests cover independent sampling,
not every correlated parameter combination. Errors are explicit rather than substituting
an approximate family.

Continuous previews show central 99.8% with a tail note. Large count previews show at
most 2,001 central outcomes with an omission note. Historical fitting remains Normal,
Lognormal, Uniform and Triangular only.

Actual per-monitor DPI, interactive ribbon/range-selection behavior, live x86 XLL loading,
full report-export interaction and arbitrary customer workbooks remain manual checks.
The live harness verifies report descriptions, not a completed interactive report export.
All eight new scenario families are covered by automated workbook-adapter tests; the
new live check is not a full interactive scenario certification.
