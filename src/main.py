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

from puzzle_downloader import download_puzzle
from sheet_formatter import create_crossword_sheet
from script_deployer import deploy_script_config

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
logger = logging.getLogger(__name__)


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

    success = 0
    for outlet in outlets:
        logger.info(f"--- {outlet} ---")
        try:
            puzzle = download_puzzle(outlet)
            result = create_crossword_sheet(spreadsheet_id, puzzle)
            if result is None:
                logger.info(f"Skipped {outlet} (already exists)")
                success += 1
                continue

            worksheet, script_config = result
            deploy_script_config(spreadsheet_id, script_config)
            logger.info(f"✓ {outlet} done")
            success += 1
        except Exception:
            logger.exception(f"✗ Failed to process {outlet}")

    logger.info(f"Finished: {success}/{len(outlets)} outlets succeeded")
    if success == 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
