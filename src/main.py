#!/usr/bin/env python3
"""Daily crossword → Google Sheets automation.

Downloads crossword puzzles via xword-dl, parses the .puz files,
and creates formatted crossword grids in a Google Sheet.

Configuration (environment variables):
    GOOGLE_CREDENTIALS  – Service-account JSON key (required)
    SPREADSHEET_ID      – Target Google Sheet ID (required)
    PUZZLE_OUTLETS      – Comma-separated outlet codes (default: "usa")
"""

import logging
import os
import sys
import time

from dotenv import load_dotenv

# Load .env when running locally; no-op in CI where env vars are set directly
load_dotenv()

from puzzle_downloader import download_puzzle
from sheet_formatter import create_crossword_sheet, sort_and_prune_tabs
from script_deployer import deploy_navigation_script

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
logger = logging.getLogger(__name__)


def _with_retry(fn, max_attempts=3, initial_backoff=10):
    """Call fn(), retrying on any exception with exponential backoff."""
    backoff = initial_backoff
    for attempt in range(1, max_attempts + 1):
        try:
            return fn()
        except Exception as exc:
            if attempt == max_attempts:
                raise
            logger.warning(
                f"Attempt {attempt}/{max_attempts} failed: {exc}. "
                f"Retrying in {backoff}s…"
            )
            time.sleep(backoff)
            backoff *= 2


def main():
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "")
    if not spreadsheet_id:
        logger.error("SPREADSHEET_ID environment variable is required")
        sys.exit(1)

    outlets = [
        o.strip()
        for o in os.environ.get("PUZZLE_OUTLETS", "usa").split(",")
        if o.strip()
    ]

    logger.info(f"Outlets to process: {outlets}")

    sheet_configs = {}  # sheet_name -> grid layout, collected for script deploy
    success = 0
    for outlet in outlets:
        logger.info(f"--- {outlet} ---")
        try:
            # download_puzzle has its own multi-strategy fallback chain;
            # we still wrap it in a retry to handle transient network errors
            # that could make every strategy fail on a single attempt.
            puzzle = _with_retry(
                lambda o=outlet: download_puzzle(o),
                max_attempts=2,
                initial_backoff=15,
            )
            result = _with_retry(lambda p=puzzle, s=spreadsheet_id: create_crossword_sheet(s, p))

            # Tab name uses YY/MM/DD (matches sheet_formatter)
            from datetime import datetime as _dt
            _short = _dt.strptime(puzzle['date'], "%Y-%m-%d").strftime("%y/%m/%d")
            sheet_name = f"{_short} - {puzzle['outlet']}"
            sheet_configs[sheet_name] = puzzle  # full puzzle data for script deployer

            if result is None:
                logger.info(f"Skipped {outlet} (already exists)")
                success += 1
                continue

            logger.info(f"✓ {outlet} done")
            success += 1
        except Exception:
            logger.exception(f"✗ Failed to process {outlet}")

    logger.info(f"Finished: {success}/{len(outlets)} outlets succeeded")

    if sheet_configs:
        try:
            _with_retry(lambda: deploy_navigation_script(spreadsheet_id, sheet_configs))
            logger.info("Navigation script deployed")
        except Exception:
            logger.exception("Failed to deploy navigation script (non-fatal)")

    # Sort tabs by date (newest first) and prune old ones
    try:
        _with_retry(lambda: sort_and_prune_tabs(spreadsheet_id))
    except Exception:
        logger.exception("Failed to sort/prune tabs (non-fatal)")

    if success == 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
