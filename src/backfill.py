#!/usr/bin/env python3
"""One-shot backfill: download the 3 most recent puzzles per outlet.

Tries dates going back from today. Skips dates that fail (no puzzle
published) and tabs that already exist.
"""

import logging
import os
import sys
import time

from dotenv import load_dotenv

load_dotenv()

from puzzle_downloader import download_puzzle
from sheet_formatter import create_crossword_sheet
from script_deployer import deploy_navigation_script
from sheet_formatter import sort_and_prune_tabs

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
logger = logging.getLogger(__name__)

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
OUTLETS = [
    o.strip()
    for o in os.environ.get("PUZZLE_OUTLETS", "wp,usa,uni").split(",")
    if o.strip()
]
TARGET_COUNT = 3       # how many puzzles per outlet
MAX_DAYS_BACK = 14     # how far back to search


def _with_retry(fn, max_attempts=3, initial_backoff=10):
    backoff = initial_backoff
    for attempt in range(1, max_attempts + 1):
        try:
            return fn()
        except Exception as exc:
            if attempt == max_attempts:
                raise
            logger.warning(f"Attempt {attempt}/{max_attempts} failed: {exc}. Retrying in {backoff}s…")
            time.sleep(backoff)
            backoff *= 2


def main():
    from datetime import date, timedelta

    today = date.today()
    sheet_configs = {}

    for outlet in OUTLETS:
        logger.info(f"=== Backfilling {outlet} (target: {TARGET_COUNT} most recent) ===")
        found = 0
        for days_back in range(0, MAX_DAYS_BACK):
            if found >= TARGET_COUNT:
                break
            dt = today - timedelta(days=days_back)
            date_str = dt.isoformat()
            logger.info(f"  Trying {outlet} {date_str}...")
            try:
                puzzle = download_puzzle(outlet, puzzle_date=date_str)
            except Exception as exc:
                logger.info(f"  No puzzle for {outlet} on {date_str}: {exc}")
                continue

            try:
                result = _with_retry(
                    lambda p=puzzle, s=SPREADSHEET_ID: create_crossword_sheet(s, p)
                )
                from datetime import datetime as _dt
                _short = _dt.strptime(puzzle['date'], "%Y-%m-%d").strftime("%y/%m/%d")
                sheet_name = f"{_short} - {puzzle['outlet']}"
                sheet_configs[sheet_name] = puzzle

                if result is None:
                    logger.info(f"  {date_str} already exists — skipping (counts toward {TARGET_COUNT})")
                else:
                    logger.info(f"  ✓ Created {sheet_name}")
                found += 1
            except Exception:
                logger.exception(f"  ✗ Failed to create sheet for {outlet} {date_str}")

        logger.info(f"  {outlet}: found {found}/{TARGET_COUNT}")

    # Deploy config for any new tabs
    if sheet_configs:
        try:
            _with_retry(lambda: deploy_navigation_script(SPREADSHEET_ID, sheet_configs))
            logger.info("Config deployed")
        except Exception:
            logger.exception("Failed to deploy config (non-fatal)")

    # Sort and prune
    try:
        _with_retry(lambda: sort_and_prune_tabs(SPREADSHEET_ID))
        logger.info("Tabs sorted and pruned")
    except Exception:
        logger.exception("Failed to sort/prune (non-fatal)")

    logger.info("Backfill complete!")


if __name__ == "__main__":
    main()
