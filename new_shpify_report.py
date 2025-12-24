# ===============================================
# file: playwright_runner.py
# Shopify Channel Performance (weekly) -> Summary -> Google Sheets
# ===============================================

import os
import sys
import time
import json
import asyncio
import argparse
from pathlib import Path
from datetime import datetime, timedelta, date
import pandas as pd

# --- Windows stability for Playwright
if sys.platform.startswith("win"):
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

# --- deps (install if missing)
def ensure(pkg: str):
    import importlib
    try:
        importlib.import_module(pkg)
    except ModuleNotFoundError:
        import subprocess
        print(f"{pkg} not found; installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

ensure("python-dotenv")
ensure("playwright")
ensure("pandas")
ensure("google-api-python-client")
ensure("google-auth")
ensure("google-auth-oauthlib")
ensure("google-auth-httplib2")

from dotenv import load_dotenv
import pandas as pd
from playwright.async_api import async_playwright, TimeoutError as PWTimeoutError

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Ensure Playwright Chromium exists (idempotent)
def ensure_playwright_browsers():
    import subprocess
    try:
        subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
    except Exception as e:
        print('Failed to install Playwright Chromium automatically:', e)
        print('Run manually: python -m playwright install chromium')

ensure_playwright_browsers()

load_dotenv()

# -------------------------
# Date helpers
# -------------------------
def parse_ymd(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()

def ymd(d: date) -> str:
    return d.isoformat()

def week_start_for(d: date, week_start: int = 0) -> date:
    # week_start: Mon=0 ... Sun=6
    delta = (d.weekday() - week_start) % 7
    return d - timedelta(days=delta)

def iter_weeks(since: date, until: date, week_start: int = 0):
    """
    Yields inclusive (start,end) ranges aligned to week_start,
    but clipped to SINCE/UNTIL.
    """
    cur = since
    while cur <= until:
        anchor = week_start_for(cur, week_start)
        week_end = anchor + timedelta(days=6)
        end = min(week_end, until)
        yield cur, end
        cur = end + timedelta(days=1)

def format_date_range(start_ymd: str, end_ymd: str) -> str:
    sdt = datetime.strptime(start_ymd, "%Y-%m-%d")
    edt = datetime.strptime(end_ymd, "%Y-%m-%d")
    return f"{sdt.month}/{sdt.day}-{edt.month}/{edt.day}"

# -------------------------
# Args + env config
# -------------------------
def get_date_range_from_args_or_env():
    parser = argparse.ArgumentParser()
    parser.add_argument("--since", type=str, default=None)
    parser.add_argument("--until", type=str, default=None)
    args = parser.parse_args()

    today = datetime.now().date()
    since = args.since or os.getenv("SINCE") or "2025-09-01"
    until = args.until or os.getenv("UNTIL") or ymd(today)
    return since, until

STORE_SLUG = os.getenv("SHOPIFY_STORE_SLUG")
SHOPIFY_EMAIL = os.getenv("SHOPIFY_EMAIL")
SHOPIFY_PASSWORD = os.getenv("SHOPIFY_PASSWORD")
AUTO_LOGIN = os.getenv("AUTO_LOGIN", "0").lower() in ("1", "true", "yes")

DOWNLOAD_DIR = Path(os.getenv("DOWNLOAD_DIR", "./downloads")).resolve()
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

SINCE, UNTIL = get_date_range_from_args_or_env()
COUNTRY = os.getenv("COUNTRY", "US")

HEADLESS = os.getenv("CHROME_HEADLESS", "0").lower() not in ("0", "false", "no")
STATE_FILE = Path(os.getenv("PLAYWRIGHT_STATE_FILE", "playwright_storage_state.json")).resolve()

UPLOAD_TO_SHEET = os.getenv("UPLOAD_TO_SHEET", "0").lower() in ("1", "true", "yes")
SHEET_ID = os.getenv("SHEET_ID")
SHEET_NAME = os.getenv("SHEET_NAME", "summary_df")
SHEET_MODE = os.getenv("SHEET_MODE", "append").strip().lower()  # append | overwrite

# Week starts Monday by default
WEEK_START = int(os.getenv("WEEK_START", "0"))

CREDENTIALS_JSON = Path(os.getenv("GOOGLE_CREDENTIALS_JSON", "credentials.json")).resolve()
TOKEN_JSON = Path(os.getenv("GOOGLE_TOKEN_JSON", "token.json")).resolve()

# -------------------------
# Google Sheets helpers (OAuth installed app)
# -------------------------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def get_sheets_service():
    creds = None
    if TOKEN_JSON.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_JSON), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_JSON.exists():
                raise FileNotFoundError(
                    f"Missing {CREDENTIALS_JSON}. Save your OAuth JSON as 'credentials.json' next to this script."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_JSON), SCOPES)
            creds = flow.run_local_server(port=0)

        TOKEN_JSON.write_text(creds.to_json(), encoding="utf-8")

    return build("sheets", "v4", credentials=creds)

def ensure_tab(service, spreadsheet_id: str, sheet_name: str):
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    existing = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if sheet_name in existing:
        return
    req = {"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=req).execute()

def upload_df_to_sheet(df: pd.DataFrame, spreadsheet_id: str, sheet_name: str, mode: str = "append"):
    if not spreadsheet_id:
        raise ValueError("SHEET_ID missing in .env")

    service = get_sheets_service()
    ensure_tab(service, spreadsheet_id, sheet_name)

    header_and_rows = [df.columns.tolist()] + df.astype(object).where(pd.notnull(df), "").values.tolist()

    if mode == "overwrite":
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A:ZZ", body={}
        ).execute()
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": header_and_rows},
        ).execute()
        print(f"✅ Uploaded (overwrite) to tab '{sheet_name}'")
        return

    # append
    first_cell = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A1:A1"
    ).execute().get("values", [])

    if not first_cell:
        # empty sheet => include header
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            body={"values": header_and_rows},
        ).execute()
        print(f"✅ Uploaded (new tab) to '{sheet_name}'")
    else:
        # existing header => append only values
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": df.astype(object).where(pd.notnull(df), "").values.tolist()},
        ).execute()
        print(f"✅ Appended {len(df)} row(s) to '{sheet_name}'")

# -------------------------
# Summarization logic
# -------------------------
def to_number(x) -> float:
    if x is None:
        return 0.0
    s = str(x).strip()
    if s in ("", "—", "-", "nan", "None"):
        return 0.0
    s = s.replace("$", "").replace(",", "")
    try:
        return float(s)
    except Exception:
        return 0.0

def safe_lower(x) -> str:
    if x is None:
        return ""
    # pandas NaN
    if isinstance(x, float) and pd.isna(x):
        return ""
    return str(x).strip().lower()

def bucket_row(ref_platform, channel, typ) -> str:
    rp = safe_lower(ref_platform)
    ch = safe_lower(channel)
    ty = safe_lower(typ)

    if ch == "direct" or rp == "direct":
        return "Direct Website Sales (Organic)"
    if ch == "google" and ty == "paid":
        return "Google ads (Sales)"
    if ch == "google" and ty == "organic":
        return "Google Search (Organic Sales)"
    if ch == "attentive" or rp == "attentive":
        return "Attentive SMS (Sales)"
    if ch == "privy" or rp == "privy":
        return "Privey Email Marketing (Sales)"
    if ch == "activecampaign" or rp == "activecampaign":
        return "ActiveCampaign (Sales)"

    return "Other Channel Sales MISC"

def build_misc_notes(df_other: pd.DataFrame, ch_col: str, ty_col: str) -> str:
    if df_other.empty:
        return ""
    tmp = df_other.copy()
    tmp["_name"] = tmp.apply(lambda r: f"{str(r[ch_col]).strip()} ({str(r[ty_col]).strip()})", axis=1)
    grouped = (
        tmp.groupby("_name", as_index=False)["_sales"].sum()
        .sort_values("_sales", ascending=False)
    )
    parts = []
    for _, r in grouped.head(12).iterrows():
        if float(r["_sales"]) > 0:
            parts.append(f"{r['_name']} ${float(r['_sales']):,.2f}")
    return " | ".join(parts)

def summarize_channel_csv_to_weekly_row(csv_path: Path, start: date, end: date) -> pd.DataFrame:
    df = pd.read_csv(csv_path)

    # ✅ remove fully blank rows
    df = df.dropna(how="all")
    df = df[(df.astype(str).apply(lambda x: x.str.strip()).ne("").any(axis=1))]

    colmap = {c.lower().strip(): c for c in df.columns}
    def col(name_lower):
        return colmap.get(name_lower, None)

    rp_col = col("referring platform")
    ch_col = col("channel")
    ty_col = col("type")
    sales_col = col("sales")
    cost_col = col("cost")

    if not (rp_col and ch_col and ty_col and sales_col):
        raise ValueError(f"CSV missing required columns. Found: {list(df.columns)}")

    df["_sales"] = df[sales_col].apply(to_number)
    df["_cost"] = df[cost_col].apply(to_number) if cost_col else 0.0

    df["_bucket"] = df.apply(lambda r: bucket_row(r[rp_col], r[ch_col], r[ty_col]), axis=1)

    agg = df.groupby("_bucket", as_index=False).agg({"_sales": "sum", "_cost": "sum"})
    sales_by = {row["_bucket"]: float(row["_sales"]) for _, row in agg.iterrows()}
    cost_by  = {row["_bucket"]: float(row["_cost"])  for _, row in agg.iterrows()}

    other_df = df[df["_bucket"] == "Other Channel Sales MISC"].copy()
    misc_notes = build_misc_notes(other_df, ch_col, ty_col)

    tot_sales = float(df["_sales"].sum())
    tot_cost  = float(df["_cost"].sum())
    gpm = ((tot_sales - tot_cost) / tot_sales) if tot_sales else 0.0

    start_str = start.isoformat()
    end_str = end.isoformat()

    row = {
        "Month": start.strftime("%B"),
        "Dates/ Week": format_date_range(start_str, end_str),

        "Direct Website Sales (Organic)": round(sales_by.get("Direct Website Sales (Organic)", 0.0), 2),
        "Google ads (Sales)": round(sales_by.get("Google ads (Sales)", 0.0), 2),
        "Google Search (Organic Sales)": round(sales_by.get("Google Search (Organic Sales)", 0.0), 2),
        "Attentive SMS (Sales)": round(sales_by.get("Attentive SMS (Sales)", 0.0), 2),
        "Privey Email Marketing (Sales)": round(sales_by.get("Privey Email Marketing (Sales)", 0.0), 2),
        "ActiveCampaign (Sales)": round(sales_by.get("ActiveCampaign (Sales)", 0.0), 2),
        "Other Channel Sales MISC": round(sales_by.get("Other Channel Sales MISC", 0.0), 2),

        "Tot Sales": round(tot_sales, 2),

        "Google ads (Cost)": round(cost_by.get("Google ads (Sales)", 0.0), 2),
        "Privey Email Marketing (Cost)": round(cost_by.get("Privey Email Marketing (Sales)", 0.0), 2),
        "Attentive SMS (Cost)": round(cost_by.get("Attentive SMS (Sales)", 0.0), 2),
        "Total Cost": round(tot_cost, 2),

        "GPM": round(gpm, 4),
        "MISC.": misc_notes,

        "Upload_Date": datetime.now().strftime("%Y-%m-%d"),
        "Range_Start": start_str,
        "Range_End": end_str,
        "Country": COUNTRY,
    }

    ordered_cols = [
        "Month",
        "Dates/ Week",
        "Direct Website Sales (Organic)",
        "Google ads (Sales)",
        "Google Search (Organic Sales)",
        "Attentive SMS (Sales)",
        "Privey Email Marketing (Sales)",
        "ActiveCampaign (Sales)",
        "Other Channel Sales MISC",
        "Tot Sales",
        "Google ads (Cost)",
        "Privey Email Marketing (Cost)",
        "Attentive SMS (Cost)",
        "Total Cost",
        "GPM",
        "MISC.",
        "Upload_Date",
        "Range_Start",
        "Range_End",
        "Country",
    ]

    return pd.DataFrame([[row.get(c, "") for c in ordered_cols]], columns=ordered_cols)

# -------------------------
# Shopify automation helpers
# -------------------------
def build_report_url(store_slug: str, since_ymd: str, until_ymd: str, country: str) -> str:
    return (
        f"https://admin.shopify.com/store/{store_slug}/marketing/reports/channels"
        f"?attributionModel=last_click_non_direct"
        f"&since={since_ymd}&until={until_ymd}"
        f"&sortColumn=sessions&sortDirection=desc"
        f"&country={country}"
    )

async def auto_login_shopify(page):
    if not SHOPIFY_EMAIL or not SHOPIFY_PASSWORD:
        raise ValueError("AUTO_LOGIN=1 but SHOPIFY_EMAIL/SHOPIFY_PASSWORD missing in .env")

    await page.goto(f"https://admin.shopify.com/store/{STORE_SLUG}", timeout=60000)
    await page.wait_for_load_state("domcontentloaded")

    # Email
    try:
        await page.fill("#account_email", SHOPIFY_EMAIL, timeout=15000)
        await page.click("button[name='commit']", timeout=15000)
    except Exception:
        pass

    # Password
    try:
        await page.fill("#account_password", SHOPIFY_PASSWORD, timeout=15000)
        await page.click("button[name='commit']", timeout=15000)
    except Exception:
        pass

    # 2FA if present
    try:
        otp_selector = "input[name='two_factor_code'], input[name='otp']"
        await page.wait_for_selector(otp_selector, timeout=8000)
        otp = os.getenv("SHOPIFY_OTP", "").strip() or input("Enter Shopify 2FA code: ").strip()
        await page.fill(otp_selector, otp)
        await page.click("button[name='commit']", timeout=15000)
    except Exception:
        pass

    await page.wait_for_selector("nav[aria-label='Primary'], #AppFrameMain", timeout=60000)

async def click_export_flow(page):
    await page.wait_for_load_state("domcontentloaded")
    await page.wait_for_timeout(2000)

    # Wait for a stable "report" UI area (fallback: any button)
    await page.wait_for_selector("button", timeout=60000)

    # Try direct Export
    for sel in [
        "button:has-text('Export')",
        "button[aria-label='Export']",
        "[role='button']:has-text('Export')",
    ]:
        try:
            await page.click(sel, timeout=8000)
            break
        except Exception:
            pass
    else:
        # Try overflow menu then Export
        for sel in [
            "button[aria-label='More actions']",
            "button[aria-haspopup='menu']",
        ]:
            try:
                await page.click(sel, timeout=8000)
                break
            except Exception:
                pass
        await page.click("text=Export", timeout=15000)

    await page.wait_for_timeout(800)

    # Choose CSV if present
    for sel in [
        "text=CSV",
        "button:has-text('CSV')",
        "[role='menuitem']:has-text('CSV')",
        "label:has-text('CSV')",
    ]:
        try:
            await page.click(sel, timeout=3000)
            break
        except Exception:
            pass

    # Confirm export in dialog (second Export click)
    for sel in [
        "button:has-text('Export')",
        "button[aria-label='Export']",
        "text=Export",
    ]:
        try:
            await page.click(sel, timeout=10000)
            break
        except Exception:
            pass

# -------------------------
# Main
# -------------------------
async def run():
    if not STORE_SLUG:
        raise ValueError("Missing SHOPIFY_STORE_SLUG in .env")

    since_d = parse_ymd(SINCE)
    until_d = parse_ymd(UNTIL)
    if until_d < since_d:
        raise ValueError("UNTIL must be >= SINCE")

    print(f"DOWNLOAD_DIR: {DOWNLOAD_DIR}")
    print(f"Range: {since_d} -> {until_d} (inclusive)")
    print(f"SHEET_NAME: {SHEET_NAME} | MODE: {SHEET_MODE} | UPLOAD_TO_SHEET: {UPLOAD_TO_SHEET}")

    overwrite_first = (SHEET_MODE == "overwrite")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=HEADLESS,
            args=["--disable-dev-shm-usage", "--no-sandbox", "--disable-setuid-sandbox"]
        )

        if STATE_FILE.exists():
            context = await browser.new_context(storage_state=str(STATE_FILE), accept_downloads=True)
            print(f"Loaded storage state from {STATE_FILE}")
        else:
            context = await browser.new_context(accept_downloads=True)

        page = await context.new_page()

        try:
            await page.goto(f"https://admin.shopify.com/store/{STORE_SLUG}", timeout=60000)

            if not STATE_FILE.exists():
                if AUTO_LOGIN:
                    print("AUTO_LOGIN enabled: attempting login...")
                    await auto_login_shopify(page)
                else:
                    print("\nFIRST RUN: Please log in to Shopify in the opened browser, then press Enter here.\n")
                    input("Press Enter after Shopify login is complete... ")

                await context.storage_state(path=str(STATE_FILE))
                print(f"✅ Saved Shopify session to: {STATE_FILE}")

            all_summary_rows = []

            for start, end in iter_weeks(since_d, until_d, week_start=WEEK_START):
                try:
                    # existing weekly export + summary + upload code
                    ...
                except Exception as e:
                    print(f"⚠️ Week {start} to {end} failed: {e}")
                    continue
                since_ymd = ymd(start)
                until_ymd = ymd(end)
                report_url = build_report_url(STORE_SLUG, since_ymd, until_ymd, COUNTRY)

                print(f"\n--- Exporting week {since_ymd} to {until_ymd} ---")
                await page.goto(report_url, timeout=60000)

                async with page.expect_download(timeout=180000) as dl_info:
                    await click_export_flow(page)

                download = await dl_info.value
                stamp = time.strftime("%Y%m%d_%H%M%S")
                raw_path = DOWNLOAD_DIR / f"shopify_channel_perf_{COUNTRY}_{since_ymd}_{until_ymd}_{stamp}.csv"
                await download.save_as(str(raw_path))
                print(f"✅ Downloaded raw: {raw_path}")

                summary_df = summarize_channel_csv_to_weekly_row(raw_path, start, end)
                print("SUMMARY ROW PREVIEW:")
                print(summary_df.to_string(index=False))

                summary_path = raw_path.with_name(raw_path.stem + "_SUMMARY.csv")
                summary_df.to_csv(summary_path, index=False)
                print(f"✅ Saved summary: {summary_path}")

                all_summary_rows.append(summary_df)

                if UPLOAD_TO_SHEET:
                    mode = "overwrite" if overwrite_first else "append"
                    upload_df_to_sheet(summary_df, SHEET_ID, SHEET_NAME, mode=mode)
                    print(f"✅ Uploaded summary row to spreadsheet {SHEET_ID} tab '{SHEET_NAME}' (mode={mode})")
                    overwrite_first = False

            if all_summary_rows:
                combined = pd.concat(all_summary_rows, ignore_index=True)
                combined_path = DOWNLOAD_DIR / f"shopify_weekly_summary_{COUNTRY}_{SINCE}_{UNTIL}_{time.strftime('%Y%m%d_%H%M%S')}.csv"
                combined.to_csv(combined_path, index=False)
                print(f"\n✅ Combined summary saved: {combined_path}")

        finally:
            await context.close()
            await browser.close()

if __name__ == "__main__":
    asyncio.run(run())
