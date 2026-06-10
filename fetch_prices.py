"""
Portfolio Price Fetcher
=======================
Fetches daily closing prices for your portfolio via InsightSentry (RapidAPI)
and saves them to a CSV file that can be imported into Excel via Power Query.

On each run the script:
  1. Reads your asset list from assets.csv
  2. Checks prices.csv to find the latest date already saved per ticker
  3. Fetches only new prices from the API (nothing already in the file)
  4. Appends the new prices and saves the updated file

Author: Built with Claude
Dependencies: pip install requests pandas python-dotenv
"""

import argparse
import logging
import os
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import requests
from dotenv import load_dotenv


# ===========================================================================
# STEP 1 — Load settings from the .env file
# ===========================================================================
# The .env file keeps your API key out of this script (safer, easier to manage).

load_dotenv()

# Your RapidAPI key — set this in the .env file, not here.
RAPIDAPI_KEY = os.getenv("RAPIDAPI_KEY")

# Rate limit — how many seconds to wait between API calls.
#   Your current (deprecated) plan: 1000 req/hour = 1 request per 3.6 seconds
#   If moved to the newer plan:      10 req/minute = 1 request per 6.0 seconds
# To change this, update RATE_LIMIT_SECONDS in your .env file.
RATE_LIMIT_SECONDS = float(os.getenv("RATE_LIMIT_SECONDS", "3.6"))

# How many daily bars to fetch per ticker.
#   10  = last 10 trading days (normal daily top-up runs)
#   750 = roughly 3 years of history (first-time setup / full refresh)
# Override on the command line with: --bars 750
DEFAULT_BARS = int(os.getenv("DEFAULT_BARS", "10"))

# API connection details — these don't need changing.
RAPIDAPI_HOST = "insightsentry.p.rapidapi.com"
BASE_URL = f"https://{RAPIDAPI_HOST}/v3/symbols"

# Retry settings — if a call fails temporarily, try again up to this many times.
MAX_RETRIES = 3
RETRY_WAIT_SECONDS = 10


# ===========================================================================
# STEP 2 — Set up logging
# ===========================================================================
# The script writes progress messages to both the screen and a log file
# so you can always see what happened on the last run.

def setup_logging(log_file: str) -> None:
    """Configure logging to write to both the screen and a log file."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(),  # also print to the screen
        ],
    )


# ===========================================================================
# STEP 3 — Read the assets input file
# ===========================================================================
# assets.csv is the file you maintain in Excel.
# Required columns : ticker, asset_type
# Optional column  : description (a friendly label for your reference)

def load_assets(input_csv: str) -> list[dict]:
    """Load the list of assets from the input CSV."""

    if not Path(input_csv).exists():
        raise FileNotFoundError(
            f"Cannot find the assets file: '{input_csv}'\n"
            f"Please make sure it is in the same folder as this script."
        )

    df = pd.read_csv(input_csv, dtype=str)

    # Tidy up column names — strip spaces and make lowercase so the CSV works
    # even if someone accidentally adds a space in a heading.
    df.columns = [c.strip().lower() for c in df.columns]

    required = {"ticker", "asset_type"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            f"Your assets.csv is missing these required columns: {missing}\n"
            f"Please check the column headings in your file."
        )

    # Strip whitespace from all values and drop rows with no ticker.
    df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
    df = df[df["ticker"].notna() & (df["ticker"] != "")]

    assets = df.to_dict(orient="records")
    logging.info(f"Loaded {len(assets)} asset(s) from '{input_csv}'")
    return assets


# ===========================================================================
# STEP 4 — Read the existing prices file (if it exists)
# ===========================================================================

def load_existing_prices(output_csv: str) -> pd.DataFrame:
    """
    Load prices already saved from previous runs.
    If the file doesn't exist yet, return an empty table — it will be
    created automatically at the end of this run.
    """
    path = Path(output_csv)

    if not path.exists():
        logging.info(
            f"No existing prices file found at '{output_csv}'. "
            f"A new file will be created on this run."
        )
        return pd.DataFrame(columns=["ticker", "excel_ticker", "date", "close", "currency"])

    df = pd.read_csv(output_csv)
    # Robustly parse the date column regardless of format
    df["date"] = pd.to_datetime(df["date"], dayfirst=False, errors="coerce")
    logging.info(f"Loaded {len(df)} existing price rows from '{output_csv}'")
    return df


def get_latest_dates(df: pd.DataFrame) -> dict[str, date]:
    """
    Find the most recent date already saved for each ticker.
    Returns a dictionary like: {"XLON:EPRA": date(2024, 3, 15), ...}
    """
    if df.empty:
        return {}
    latest = df.groupby("ticker")["date"].max()
    result = {}
    for ticker, d in latest.items():
        try:
            # Handle both datetime objects and strings
            if hasattr(d, "date"):
                result[ticker] = d.date()
            else:
                result[ticker] = pd.to_datetime(d, dayfirst=False).date()
        except Exception:
            pass
    return result


# ===========================================================================
# STEP 5 — Call the InsightSentry API
# ===========================================================================

def build_headers() -> dict:
    """Build the HTTP headers needed to authenticate with RapidAPI."""
    if not RAPIDAPI_KEY:
        raise EnvironmentError(
            "No API key found!\n"
            "Please add your RapidAPI key to the .env file:\n"
            "  RAPIDAPI_KEY=your_key_here"
        )
    return {
        "X-RapidAPI-Key":  RAPIDAPI_KEY,
        "X-RapidAPI-Host": RAPIDAPI_HOST,
    }


def fetch_prices_for_ticker(ticker: str, num_bars: int, currency_code: str = "", excel_ticker: str = "") -> list[dict]:
    """
    Fetch the most recent daily closing prices for one ticker.

    Parameters:
        ticker   : the ticker code, e.g. "LSE:EPRA"
        num_bars : how many trading days of data to request

    Returns a list of rows, e.g.:
        [{"ticker": "LSE:EPRA", "date": date(2024,3,15),
          "close": 12.34, "currency": "GBP"}, ...]
    """
    # Build the URL manually, telling urllib to leave the colon in "XLON:EPRA"
    # untouched. Without this, requests would encode it as "XLON%3AEPRA"
    # which InsightSentry rejects with a 400 error.
    encoded_ticker = quote(ticker, safe=":")
    url = f"{BASE_URL}/{encoded_ticker}/series"

    params = {
        "bar_type":    "day",      # daily bars (not minute/hour)
        "bar_interval": 1,         # 1-day interval
        "dp":          num_bars,   # number of bars to return (InsightSentry uses "dp")
        "badj":        "true",     # adjust prices for any stock splits
        "dadj":        "false",    # do NOT adjust for dividends (raw close)
        "extended":    "true",     # include extended hours data
        "long_poll":   "false",    # don't wait for new data
    }

    headers = build_headers()

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.get(url, headers=headers, params=params, timeout=30)

            # --- Handle specific HTTP status codes ---

            if response.status_code == 429:
                # Too Many Requests — slow down and retry
                wait = RETRY_WAIT_SECONDS * attempt
                logging.warning(
                    f"  [{ticker}] Rate limit hit (429). "
                    f"Waiting {wait}s before retry {attempt}/{MAX_RETRIES}..."
                )
                time.sleep(wait)
                continue

            if response.status_code == 401:
                # Authentication failed
                raise EnvironmentError(
                    "API authentication failed (401 error).\n"
                    "Please check your RAPIDAPI_KEY in the .env file is correct."
                )

            if response.status_code == 404:
                # Ticker not recognised by the API
                logging.warning(
                    f"  [{ticker}] Not found (404). "
                    f"Please check the ticker code is correct."
                )
                return []

            # Any other HTTP error (500, 503 etc.) raises an exception
            response.raise_for_status()

            # --- Parse and return the JSON response ---
            data = response.json()
            return parse_api_response(ticker, data, currency_code, excel_ticker)

        except requests.exceptions.Timeout:
            logging.warning(
                f"  [{ticker}] Request timed out "
                f"(attempt {attempt}/{MAX_RETRIES})"
            )
        except requests.exceptions.ConnectionError:
            logging.warning(
                f"  [{ticker}] Connection error "
                f"(attempt {attempt}/{MAX_RETRIES}). "
                f"Please check your internet connection."
            )
        except requests.exceptions.HTTPError as e:
            logging.error(f"  [{ticker}] HTTP error: {e}")
            return []

        # Wait before retrying
        if attempt < MAX_RETRIES:
            logging.info(f"  [{ticker}] Retrying in {RETRY_WAIT_SECONDS}s...")
            time.sleep(RETRY_WAIT_SECONDS)

    logging.error(
        f"  [{ticker}] Failed after {MAX_RETRIES} attempts. Skipping."
    )
    return []


def parse_api_response(ticker: str, data: dict, currency_code: str = "", excel_ticker: str = "") -> list[dict]:
    """
    Convert the raw JSON from InsightSentry into a clean list of price rows.

    The API returns a response like:
        {
          "code": "LSE:EPRA",
          "bar_type": "1D",
          "series": [
            {"time": 1710460800, "open": 12.1, "high": 12.5,
             "low": 11.9, "close": 12.3, "volume": 50000},
            ...
          ]
        }

    Where:
        "time"  = Unix timestamp (seconds since 1 Jan 1970 UTC)
        "close" = closing price
    Note: currency is not in the response — we look it up from assets.csv
    """
    rows = []

    # The actual response uses full field names: "time", "open", "high",
    # "low", "close", "volume" — and currency is not included in this endpoint.
    # We pass currency through from the assets.csv instead.
    currency = currency_code  # passed in from the caller

    series = data.get("series")

    if not series:
        logging.warning(f"  [{ticker}] Response contained no price data.")
        return []

    for bar in series:
        try:
            timestamp   = bar.get("time")   # "time" not "t"
            close_price = bar.get("close")  # "close" not "c"

            if timestamp is None or close_price is None:
                continue  # skip any incomplete bars

            # Convert Unix timestamp to a plain calendar date
            bar_date = datetime.fromtimestamp(timestamp, timezone.utc).date()

            rows.append({
                "ticker":       ticker,
                "excel_ticker": excel_ticker,
                "date":         bar_date,
                "close":        round(float(close_price), 6),
                "currency":     currency,
            })

        except (KeyError, ValueError, TypeError) as e:
            # If one bar is malformed, skip it and carry on with the rest
            logging.debug(f"  [{ticker}] Skipping malformed bar: {e}")
            continue

    return rows


# ===========================================================================
# STEP 6 — Save the updated prices file
# ===========================================================================

def save_prices(df: pd.DataFrame, output_csv: str) -> None:
    """Sort the combined prices and save to CSV."""
    df = df.sort_values(["ticker", "date"])
    # Write dates as YYYY-MM-DD text — clean for Power Query to parse
    df["date"] = pd.to_datetime(df["date"]).dt.strftime("%Y-%m-%d")
    df.to_csv(output_csv, index=False)
    logging.info(f"Saved {len(df)} total rows to '{output_csv}'")


# ===========================================================================
# STEP 7 — Main orchestration
# ===========================================================================
# This is the function that ties all the steps above together.

def main(input_csv: str, output_csv: str, log_file: str, num_bars: int) -> None:

    setup_logging(log_file)

    logging.info("=" * 60)
    logging.info("Portfolio Price Fetcher — starting run")
    logging.info(f"  Assets file  : {input_csv}")
    logging.info(f"  Output file  : {output_csv}")
    logging.info(f"  Bars to fetch: {num_bars} trading days per ticker")
    logging.info(f"  Rate limit   : one API call every {RATE_LIMIT_SECONDS}s")
    logging.info("=" * 60)

    # Load the asset list and any prices already saved
    assets        = load_assets(input_csv)
    existing_df   = load_existing_prices(output_csv)
    latest_dates  = get_latest_dates(existing_df)

    today         = date.today()
    all_new_rows  = []
    skipped_count = 0
    error_count   = 0

    for asset in assets:
        ticker        = asset["ticker"]
        excel_ticker  = asset.get("excel_ticker", ticker)  # fallback to ticker if not set
        asset_type    = asset.get("asset_type", "")
        description   = asset.get("description", "")
        currency_code = asset.get("currency_code", "")  # optional column in assets.csv

        # Build a friendly display label for the log
        label = ticker + (f" ({description})" if description else "")

        # --- Check if this ticker is already up to date ---
        latest = latest_dates.get(ticker)
        if latest and latest >= today - timedelta(days=1):
            logging.info(
                f"  {label} — already up to date (latest: {latest}). Skipping."
            )
            skipped_count += 1
            continue

        if latest:
            logging.info(
                f"  {label} [{asset_type}] — fetching {num_bars} bars "
                f"(will filter to dates after {latest})"
            )
        else:
            logging.info(
                f"  {label} [{asset_type}] — no existing data, "
                f"fetching {num_bars} bars"
            )

        # --- Call the API ---
        rows = fetch_prices_for_ticker(ticker, num_bars, currency_code, excel_ticker)

        if not rows:
            error_count += 1
        else:
            # Filter out any dates we already have in the file
            if latest:
                rows = [r for r in rows if r["date"] > latest]

            if rows:
                all_new_rows.extend(rows)
                most_recent = max(r["date"] for r in rows)
                logging.info(
                    f"  {label} — {len(rows)} new row(s) added "
                    f"(most recent: {most_recent})"
                )
            else:
                logging.info(f"  {label} — already up to date, nothing new to add.")
                skipped_count += 1

        # Wait between calls to stay within the API rate limit
        time.sleep(RATE_LIMIT_SECONDS)

    # --- Combine new rows with existing data and save ---
    if all_new_rows:
        new_df   = pd.DataFrame(all_new_rows)
        new_df["date"] = pd.to_datetime(new_df["date"])

        # Stack old + new rows, then remove any accidental duplicates
        # (same ticker on the same date)
        combined = pd.concat([existing_df, new_df], ignore_index=True)
        combined = combined.drop_duplicates(subset=["ticker", "date"], keep="last")

        save_prices(combined, output_csv)
    else:
        logging.info("No new price data to add — output file is unchanged.")

    # --- Print a run summary ---
    logging.info("-" * 60)
    logging.info("Run complete.")
    logging.info(f"  New price rows added : {len(all_new_rows)}")
    logging.info(f"  Tickers skipped      : {skipped_count} (already up to date)")
    logging.info(f"  Tickers with errors  : {error_count}")
    logging.info("=" * 60)


# ===========================================================================
# Entry point
# ===========================================================================
# Python runs this block when you execute the script directly.
# The argparse section lets you pass optional arguments on the command line.

if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Fetch portfolio closing prices via InsightSentry (RapidAPI)"
    )
    parser.add_argument(
        "--input", default="assets.csv",
        help="Path to your assets CSV (default: assets.csv)"
    )
    parser.add_argument(
        "--output", default="prices.csv",
        help="Path to the output prices CSV (default: prices.csv)"
    )
    parser.add_argument(
        "--log", default="fetch.log",
        help="Path to the log file (default: fetch.log)"
    )
    parser.add_argument(
        "--bars", type=int, default=DEFAULT_BARS,
        help=(
            f"Number of daily bars to fetch per ticker "
            f"(default: {DEFAULT_BARS}). "
            f"Use --bars 750 for an initial backfill of ~3 years."
        )
    )

    args = parser.parse_args()
    main(args.input, args.output, args.log, args.bars)
