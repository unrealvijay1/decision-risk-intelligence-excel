# Telivu

**Open-Source Decision Intelligence for Excel**

Simulation • Optimization • Scenarios • Sensitivity • SPC

**Free. Open Source. No Subscription. No Activation.**

Telivu's current development code runs the same analytical toolkit for everyone without activation or access tiers. The existing public 0.2.1 installer retains its legacy trial/activation until an activation-free release is authorized. This repository hosts the official static website and existing installer distribution; application source is available in the existing decision-risk-intelligence-excel repository; legal open-source license approval remains pending owner review.

[Website](https://unrealvijay1.github.io/telivu/) · [Download](https://unrealvijay1.github.io/telivu/#download) · [Source](https://github.com/unrealvijay1/decision-risk-intelligence-excel) · [Documentation](https://unrealvijay1.github.io/telivu/resources/documentation/) · [Issues](https://github.com/unrealvijay1/decision-risk-intelligence-excel/issues) · [Contributing](https://github.com/unrealvijay1/telivu/blob/main/CONTRIBUTING.md) · [Support information](https://unrealvijay1.github.io/telivu/support/)

Apache-2.0 is preferred but not applied: Office interop and artwork rights need owner review. No legal usage/redistribution grant is made here. Donation provider setup required; no donation collection endpoint exists.

## Website development

```powershell
python site/scripts/build.py --site-url https://unrealvijay1.github.io/telivu/
python site/scripts/check.py site/_build
```

Static HTML/CSS/JavaScript with no framework, backend, analytics or external fonts. The existing main-branch Pages workflow builds only site/_build. Every download button uses the unchanged data-download handler in site/js/site.js and verified asset configuration in site/site-config.js. Keep the original release URL and asset name. No new release or installer upload is needed for website edits. Product images are real development captures with synthetic data.
