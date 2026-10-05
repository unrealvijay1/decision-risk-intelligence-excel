# Telivu

**Open-Source Decision Intelligence for Excel**

Simulation • Optimization • Scenarios • Sensitivity • SPC

**Free. Open Source. No Subscription. No Activation.**

The current activation-free installer runs without product keys, trial limits, expiry or paid feature gates. Its assembly version remains 0.2.1. The historical v0.2.1 release remains available separately. Open-source license approval remains pending.

[Website](https://unrealvijay1.github.io/telivu/) · [Download](https://unrealvijay1.github.io/telivu/#download) · [Source](https://github.com/unrealvijay1/decision-risk-intelligence-excel) · [Documentation](https://unrealvijay1.github.io/telivu/resources/documentation/) · [Issues](https://github.com/unrealvijay1/decision-risk-intelligence-excel/issues) · [Contributing](https://github.com/unrealvijay1/telivu/blob/main/CONTRIBUTING.md) · [Support information](https://unrealvijay1.github.io/telivu/support/)

Apache-2.0 is preferred but not applied: Office interop and artwork rights need owner review. No legal usage/redistribution grant is made here. Donation provider setup required; no donation collection endpoint exists.

## Website development

```powershell
python site/scripts/build.py --site-url https://unrealvijay1.github.io/telivu/
python site/scripts/check.py site/_build
```

Static HTML/CSS/JavaScript with no framework, backend, analytics or external fonts. The existing main-branch Pages workflow builds only site/_build. Every download button uses the unchanged data-download handler in site/js/site.js and verified asset configuration in site/site-config.js. Keep the original release URL and asset name. No new release or installer upload is needed for website edits. Product images are real development captures with synthetic data.
