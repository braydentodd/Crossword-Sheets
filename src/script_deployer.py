"""Deploy the Apps Script to the bound spreadsheet.

Google Sheets doesn't allow creating Apps Script via the Sheets API,
so this module writes the PUZZLE_CONFIGS data into a hidden config sheet
that the Apps Script reads at runtime. This avoids needing the
Apps Script API (which requires OAuth consent screens).

The Apps Script itself is a one-time manual paste — but the config
data is refreshed automatically on every run.
"""

import json
import logging

import gspread

logger = logging.getLogger(__name__)

CONFIG_SHEET_NAME = "_crossword_config"


def deploy_script_config(spreadsheet_id, script_config, client=None):
    """Write the script config for one puzzle to the hidden config sheet.

    Args:
        spreadsheet_id: Google Sheets spreadsheet ID.
        script_config: dict from sheet_formatter.create_crossword_sheet().
        client: Optional pre-authenticated gspread client.
    """
    if client is None:
        from sheet_formatter import _get_client
        client = _get_client()

    spreadsheet = client.open_by_key(spreadsheet_id)

    # Get or create the hidden config sheet
    try:
        config_ws = spreadsheet.worksheet(CONFIG_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        config_ws = spreadsheet.add_worksheet(
            title=CONFIG_SHEET_NAME, rows=1000, cols=2,
        )
        # Hide this sheet from casual view
        spreadsheet.batch_update({
            "requests": [{
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": config_ws.id,
                        "hidden": True,
                    },
                    "fields": "hidden",
                }
            }]
        })
        logger.info(f"Created hidden config sheet '{CONFIG_SHEET_NAME}'")

    sheet_name = script_config["sheetName"]

    # Read existing rows to find if this sheet already has a config entry
    all_values = config_ws.get_all_values()
    target_row = None
    for i, row in enumerate(all_values):
        if row and row[0] == sheet_name:
            target_row = i + 1  # 1-indexed
            break

    config_json = json.dumps(script_config, separators=(",", ":"))

    if target_row:
        config_ws.update(
            [[sheet_name, config_json]],
            f"A{target_row}:B{target_row}",
            value_input_option="RAW",
        )
    else:
        config_ws.append_row(
            [sheet_name, config_json],
            value_input_option="RAW",
        )

    logger.info(f"Wrote script config for '{sheet_name}'")
