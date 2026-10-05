# Contributing to the Telivu website

Report bugs or request features through https://github.com/unrealvijay1/telivu/issues. Include the page URL, browser/viewport, expected result and steps to reproduce. For application issues include Windows/Excel architecture and synthetic models only.

Clone the existing public repository and use Python 3.10+ from its root:

```powershell
python site/scripts/build.py --site-url https://unrealvijay1.github.io/telivu/
python site/scripts/check.py site/_build
python -m http.server 8765 --bind 127.0.0.1 --directory site/_build
git diff --check
```

Edit the source HTML/CSS/JavaScript, not site/_build. Optional browser QA uses node site/scripts/browser-check.mjs with an existing Playwright installation (PLAYWRIGHT_PATH) and browser executable (BROWSER_PATH). Verify desktop/tablet/mobile navigation, keyboard use, assets and all links.

Pull requests should explain the problem, final behavior and checks. Preserve the existing data-download/config mechanism, asset URL, Pages workflow and release history. Use real product captures and implemented capabilities. Do not add placeholders, paid tiers, fake donation destinations, generated binaries, credentials or confidential data. The application is maintained separately with .NET 10/Excel-DNA; its source is publicly available and legal license approval remains pending, so application build commands do not apply to this website-only checkout.
