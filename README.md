export_shopify_channel_performance

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

Run the script (ensure `.env` contains `SHOPIFY_STORE_SLUG`, `SHOPIFY_EMAIL`, `SHOPIFY_PASSWORD`):


Playwright (alternative, more robust automation):

 - Install Python dependencies including Playwright:

```bash
python -m pip install -r requirements.txt
```

 - Install Playwright browser binaries (required once):

```bash
python -m playwright install chromium
```

 - Run the Playwright runner (created as `playwright_runner.py`) or run the notebook cell that installs Playwright and runs the export.

Notes:
- Playwright auto-manages Chromium and is often more reliable on Windows than Selenium+ChromeDriver.
- If you prefer the notebook, run the Playwright cells near the end of `report_analysis.ipynb`.
```bash
python export_shopify_channel_performance.py
```

Notes:
- `webdriver-manager` will download the ChromeDriver automatically.
- The notebook contains fallback pip-installs for missing packages; prefer using the `requirements.txt` in reproducible environments.

# Shopify_project
This project uses Shopify’s web analytics to generate a weekly sales report that highlights performance across different channels. The report will help stakeholders gain clearer insights into overall sales, conversion rates, user retention, and related metrics.
