# Telivu

**Open-Source Decision Intelligence for Excel**

Simulation • Optimization • Scenarios • Sensitivity • SPC

**Free. Open Source. No Subscription. No Activation.**

The current application source has no activation, license key, trial, expiry or edition checks. The same toolkit serves individuals and organizations. Voluntary donations/sponsorship may support development; services such as training, consulting, implementation and support can fund the project without restricting the core product.

[Website](https://unrealvijay1.github.io/telivu/) · [Download](https://unrealvijay1.github.io/telivu/#download) · [Source](https://github.com/unrealvijay1/decision-risk-intelligence-excel) · [Documentation](https://unrealvijay1.github.io/telivu/resources/documentation/) · [Issues](https://github.com/unrealvijay1/decision-risk-intelligence-excel/issues) · [Contributing](CONTRIBUTING.md) · [Support information](https://unrealvijay1.github.io/telivu/support/)

## Distribution and legal status

The public download still serves the unchanged unsigned 0.2.1 installer with legacy trial/activation. This development build removes activation; a new public release requires separate authorization. No release asset or download mechanism is changed by this task.

Apache-2.0 is the preferred candidate, but no legal license has been applied. [The dependency audit](OPEN_SOURCE_LICENSE_AUDIT.md) identifies unresolved Office interop terms and artwork ownership for owner review. Until resolved and approved, open-source positioning expresses the strategy and does not grant source redistribution or commercial-use rights.

Donation provider setup required; no donation button or payment collection exists.

## Develop and validate

See [CONTRIBUTING](CONTRIBUTING.md), [installer instructions](Installer/README.md), [startup fixture](MonteCarlo.Excel.StartupChecks/README.md) and [layout harness](MonteCarlo.Results.LayoutChecks/README.md). The product is a C#/.NET 10 Windows Excel-DNA add-in; builds produce packed x86/x64 XLLs. The ribbon includes About Telivu under Help. Old user activation files/Registry values are ignored and never deleted by startup or installation.
