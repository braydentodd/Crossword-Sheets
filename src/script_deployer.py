"""Write per-sheet grid config to the hidden _xw_config sheet.

The Apps Script API does not support service accounts, so the Apps Script
must be installed in the spreadsheet manually once:

  1. Open the spreadsheet in your browser.
  2. Extensions → Apps Script.
  3. Delete the default code, paste the contents of src/apps_script/Code.gs.
  4. Save (Ctrl+S / Cmd+S).

After that one-time step, the script reads _xw_config automatically for
every new puzzle tab — no further action needed.
"""

import json
import logging
import os

import gspread
from google.oauth2.service_account import Credentials

import config as C

logger = logging.getLogger(__name__)

_SCRIPT_ID_FILE = os.path.join(os.path.dirname(__file__), "..", ".script_id")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _get_credentials():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        return Credentials.from_service_account_info(
            json.loads(creds_json), scopes=SCOPES
        )
    return Credentials.from_service_account_file("credentials.json", scopes=SCOPES)


def _load_script_source():
    script_dir = os.path.join(os.path.dirname(__file__), "apps_script")
    with open(os.path.join(script_dir, "Code.gs")) as f:
        code = f.read()
    with open(os.path.join(script_dir, "appsscript.json")) as f:
        manifest = f.read()
    return code, manifest


def _load_cached_script_id():
    if os.path.exists(_SCRIPT_ID_FILE):
        with open(_SCRIPT_ID_FILE) as f:
            sid = f.read().strip()
            return sid or None
    return None


def _save_script_id(script_id):
    with open(_SCRIPT_ID_FILE, "w") as f:
        f.write(script_id)


# ---------------------------------------------------------------------------
# Per-sheet config builder
# ---------------------------------------------------------------------------

def _build_sheet_config(puzzle):
    """Build the full config dict for one puzzle sheet.

    Includes grid dimensions, column positions, word-membership map, and
    clue row positions — everything Code.gs needs at runtime.
    """
    w    = puzzle["width"]
    h    = puzzle["height"]
    grid = puzzle["grid"]

    total_grid_cols = C.GRID_START_COL + w * 2
    across_num_col  = total_grid_cols + C.CLUE_GAP_COLS
    across_text_col = across_num_col + 1
    down_num_col    = across_text_col + 1
    down_text_col   = down_num_col + 1
    clue_start_row  = C.HEADER_ROW + C.CLUE_HEADER_ROWS

    cell_map   = {}
    word_cells = {}

    for clue in puzzle["across_clues"]:
        num    = clue["num"]
        length = clue["length"]
        for r, row in enumerate(grid):
            for c_idx, cell in enumerate(row):
                if cell["number"] == num:
                    cells = [
                        [r, c_idx + off]
                        for off in range(length)
                        if c_idx + off < w and not grid[r][c_idx + off]["is_black"]
                    ]
                    word_cells[f"A{num}"] = cells
                    for coord in cells:
                        cell_map.setdefault(f"{coord[0]},{coord[1]}", {})["a"] = num
                    break
            else:
                continue
            break

    for clue in puzzle["down_clues"]:
        num    = clue["num"]
        length = clue["length"]
        for r, row in enumerate(grid):
            for c_idx, cell in enumerate(row):
                if cell["number"] == num:
                    cells = [
                        [r + off, c_idx]
                        for off in range(length)
                        if r + off < h and not grid[r + off][c_idx]["is_black"]
                    ]
                    word_cells[f"D{num}"] = cells
                    for coord in cells:
                        cell_map.setdefault(f"{coord[0]},{coord[1]}", {})["d"] = num
                    break
            else:
                continue
            break

    # clueRows: word_key → [0-indexed sheet row, 0-indexed sheet col]
    clue_rows = {}
    for i, clue in enumerate(puzzle["across_clues"]):
        clue_rows[f"A{clue['num']}"] = [clue_start_row + i, across_text_col]
    for i, clue in enumerate(puzzle["down_clues"]):
        clue_rows[f"D{clue['num']}"] = [clue_start_row + i, down_text_col]

    # blackCells: "r,c" → 1  (logical grid coords, for fast JS lookup)
    black_cells = {
        f"{r},{c_idx}": 1
        for r, row in enumerate(grid)
        for c_idx, cell in enumerate(row)
        if cell["is_black"]
    }

    # solutions: "r,c" → letter  (for Check / Reveal in Apps Script)
    solutions = {
        f"{r},{c_idx}": cell["solution"]
        for r, row in enumerate(grid)
        for c_idx, cell in enumerate(row)
        if not cell["is_black"] and cell.get("solution")
    }

    max_clue_rows = max(len(puzzle["across_clues"]), len(puzzle["down_clues"]))

    return {
        "gridStartRow":  C.GRID_START_ROW,
        "gridStartCol":  C.GRID_START_COL,
        "gridWidth":     w,
        "gridHeight":    h,
        "cellMap":       cell_map,
        "wordCells":     word_cells,
        "clueRows":      clue_rows,
        "blackCells":    black_cells,
        "solutions":     solutions,
        "acrossNumCol":  across_num_col,
        "acrossTextCol": across_text_col,
        "downNumCol":    down_num_col,
        "downTextCol":   down_text_col,
        "clueStartRow":  clue_start_row,
        "maxClueRows":   max_clue_rows,
    }


# ---------------------------------------------------------------------------
# _xw_config sheet  (one row per puzzle tab: [sheetName, jsonConfig])
# ---------------------------------------------------------------------------

def _write_xw_config(spreadsheet, sheet_puzzles):
    """Write per-sheet grid config to hidden _xw_config sheet.

    Merges with any existing rows so previously-created puzzle tabs
    (not processed in this run) keep their config.
    """
    try:
        cfg_sheet = spreadsheet.worksheet("_xw_config")
    except gspread.exceptions.WorksheetNotFound:
        cfg_sheet = spreadsheet.add_worksheet("_xw_config", rows=50, cols=2)
        spreadsheet.batch_update({"requests": [{
            "updateSheetProperties": {
                "properties": {"sheetId": cfg_sheet.id, "hidden": True},
                "fields": "hidden",
            }
        }]})

    # Read existing rows; discard legacy single-cell format (starts with '{')
    existing = {}
    for row in cfg_sheet.get_all_values():
        if len(row) >= 2 and row[0] and not row[0].startswith("{"):
            existing[row[0]] = row[1]

    # Update with freshly-built configs
    for sheet_name, puzzle in sheet_puzzles.items():
        cfg = _build_sheet_config(puzzle)
        existing[sheet_name] = json.dumps(cfg, separators=(",", ":"))

    rows = [[name, cfg_json] for name, cfg_json in existing.items()]
    if rows:
        cfg_sheet.clear()
        cfg_sheet.update(rows, "A1", value_input_option="RAW")

    logger.info(f"Updated _xw_config sheet ({len(existing)} sheet(s))")


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def deploy_navigation_script(spreadsheet_id, sheet_puzzles):
    """Write per-sheet grid config to the hidden _xw_config sheet.

    The Apps Script reads this at runtime for every puzzle tab, so this is
    the only programmatic step needed after the one-time manual script install.
    """
    creds = _get_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)
    _write_xw_config(spreadsheet, sheet_puzzles)
    logger.info("_xw_config updated — Apps Script will read this at runtime")
