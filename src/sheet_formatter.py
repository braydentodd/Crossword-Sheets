"""Format and populate a Google Sheet with a crossword puzzle grid."""

import json
import logging
import os

import gspread
from google.oauth2.service_account import Credentials

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Layout constants (0-indexed)
GRID_START_ROW = 3  # Row 4 in the sheet (rows 1-3: title, author, spacer)
GRID_START_COL = 1  # Column B (column A is a narrow gutter)
CELL_SIZE_PX = 36  # Square cell dimension
CLUE_GAP_COLS = 2  # Gap between grid and clue list
CLUE_COL_WIDTH_PX = 360  # Width of the clue column


# ---------------------------------------------------------------------------
# Google Auth
# ---------------------------------------------------------------------------


def _get_client():
    """Authenticate with Google and return a gspread client.

    Reads credentials from the GOOGLE_CREDENTIALS env var (JSON string)
    or falls back to a local credentials.json file.
    """
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(creds)


# ---------------------------------------------------------------------------
# Cell helpers
# ---------------------------------------------------------------------------

_BLACK = {"red": 0.12, "green": 0.12, "blue": 0.12}
_WHITE = {"red": 1, "green": 1, "blue": 1}
_BORDER_COLOR = {"red": 0, "green": 0, "blue": 0}


def _border(width=1):
    return {"style": "SOLID", "width": width, "color": _BORDER_COLOR}


def _all_borders(width=1):
    b = _border(width)
    return {"top": b, "bottom": b, "left": b, "right": b}


# ---------------------------------------------------------------------------
# Clue lookup & notes
# ---------------------------------------------------------------------------


def _build_clue_lookup(puzzle_data):
    """Map clue number -> {'across': text, 'down': text}."""
    lookup = {}
    for c in puzzle_data["across_clues"]:
        lookup.setdefault(c["num"], {})["across"] = f"{c['clue']} ({c['length']})"
    for c in puzzle_data["down_clues"]:
        lookup.setdefault(c["num"], {})["down"] = f"{c['clue']} ({c['length']})"
    return lookup


def _cell_note(number, clue_lookup):
    """Build a hover-note string for a numbered cell."""
    info = clue_lookup.get(number)
    if not info:
        return None
    parts = []
    if "across" in info:
        parts.append(f"{number}-Across: {info['across']}")
    if "down" in info:
        parts.append(f"{number}-Down: {info['down']}")
    return "\n".join(parts) if parts else None


# ---------------------------------------------------------------------------
# Grid row data (values + formatting for one Sheets API call)
# ---------------------------------------------------------------------------


def _build_grid_rows(puzzle_data, clue_lookup):
    """Return a list of RowData dicts for the entire crossword grid."""
    rows = []
    for grid_row in puzzle_data["grid"]:
        cells = []
        for cell in grid_row:
            if cell["is_black"]:
                cells.append(
                    {
                        "userEnteredFormat": {
                            "backgroundColor": _BLACK,
                            "borders": _all_borders(1),
                        }
                    }
                )
            else:
                cd = {
                    "userEnteredFormat": {
                        "backgroundColor": _WHITE,
                        "borders": _all_borders(1),
                        "verticalAlignment": "TOP",
                        "horizontalAlignment": "LEFT",
                        "textFormat": {
                            "fontSize": 7,
                            "bold": True,
                            "foregroundColor": {
                                "red": 0.25,
                                "green": 0.25,
                                "blue": 0.25,
                            },
                        },
                    }
                }
                if cell["number"] is not None:
                    cd["userEnteredValue"] = {"stringValue": str(cell["number"])}
                    note = _cell_note(cell["number"], clue_lookup)
                    if note:
                        cd["note"] = note
                cells.append(cd)
        rows.append({"values": cells})
    return rows


# ---------------------------------------------------------------------------
# Clue text
# ---------------------------------------------------------------------------


def _build_clue_lines(puzzle_data):
    """Return a flat list of strings for the clue column."""
    lines = ["ACROSS", ""]
    for c in puzzle_data["across_clues"]:
        lines.append(f"{c['num']}. {c['clue']} ({c['length']})")
    lines.extend(["", "", "DOWN", ""])
    for c in puzzle_data["down_clues"]:
        lines.append(f"{c['num']}. {c['clue']} ({c['length']})")
    return lines


# ---------------------------------------------------------------------------
# Formatting batch requests
# ---------------------------------------------------------------------------


def _formatting_requests(sheet_id, puzzle_data, clue_lines):
    """Build the list of Sheets API batchUpdate requests."""
    w = puzzle_data["width"]
    h = puzzle_data["height"]
    clue_col = GRID_START_COL + w + CLUE_GAP_COLS

    reqs = []

    # Hide default gridlines
    reqs.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"hideGridlines": True},
                },
                "fields": "gridProperties.hideGridlines",
            }
        }
    )

    # Tab colour
    reqs.append(
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "tabColor": {"red": 0.18, "green": 0.38, "blue": 0.75},
                },
                "fields": "tabColor",
            }
        }
    )

    # Column A gutter
    reqs.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "properties": {"pixelSize": 20},
                "fields": "pixelSize",
            }
        }
    )

    # Grid column widths
    reqs.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": GRID_START_COL,
                    "endIndex": GRID_START_COL + w,
                },
                "properties": {"pixelSize": CELL_SIZE_PX},
                "fields": "pixelSize",
            }
        }
    )

    # Grid row heights
    reqs.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": GRID_START_ROW,
                    "endIndex": GRID_START_ROW + h,
                },
                "properties": {"pixelSize": CELL_SIZE_PX},
                "fields": "pixelSize",
            }
        }
    )

    # Clue column width
    reqs.append(
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": clue_col,
                    "endIndex": clue_col + 1,
                },
                "properties": {"pixelSize": CLUE_COL_WIDTH_PX},
                "fields": "pixelSize",
            }
        }
    )

    # --- Title (row 0) ---
    reqs.append(
        {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": GRID_START_COL,
                    "endColumnIndex": GRID_START_COL + w,
                },
                "mergeType": "MERGE_ALL",
            }
        }
    )
    reqs.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": GRID_START_COL,
                    "endColumnIndex": GRID_START_COL + w,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "fontSize": 14,
                            "bold": True,
                            "foregroundColor": {
                                "red": 0.1,
                                "green": 0.1,
                                "blue": 0.1,
                            },
                        },
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(textFormat,verticalAlignment)",
            }
        }
    )

    # --- Author (row 1) ---
    reqs.append(
        {
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": GRID_START_COL,
                    "endColumnIndex": GRID_START_COL + w,
                },
                "mergeType": "MERGE_ALL",
            }
        }
    )
    reqs.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": GRID_START_COL,
                    "endColumnIndex": GRID_START_COL + w,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "fontSize": 10,
                            "italic": True,
                            "foregroundColor": {
                                "red": 0.4,
                                "green": 0.4,
                                "blue": 0.4,
                            },
                        },
                        "verticalAlignment": "MIDDLE",
                    }
                },
                "fields": "userEnteredFormat(textFormat,verticalAlignment)",
            }
        }
    )

    # --- Clue base formatting (word wrap + font) ---
    reqs.append(
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": GRID_START_ROW,
                    "endRowIndex": GRID_START_ROW + len(clue_lines),
                    "startColumnIndex": clue_col,
                    "endColumnIndex": clue_col + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "fontSize": 9,
                            "foregroundColor": {
                                "red": 0.2,
                                "green": 0.2,
                                "blue": 0.2,
                            },
                        },
                        "wrapStrategy": "WRAP",
                    }
                },
                "fields": "userEnteredFormat(textFormat,wrapStrategy)",
            }
        }
    )

    # --- Clue section headers (ACROSS / DOWN) bold ---
    across_hdr_row = GRID_START_ROW  # first line of clue_lines
    down_hdr_row = GRID_START_ROW + len(puzzle_data["across_clues"]) + 4

    for hdr_row in [across_hdr_row, down_hdr_row]:
        reqs.append(
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": hdr_row,
                        "endRowIndex": hdr_row + 1,
                        "startColumnIndex": clue_col,
                        "endColumnIndex": clue_col + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": {
                                "fontSize": 12,
                                "bold": True,
                                "foregroundColor": {
                                    "red": 0.1,
                                    "green": 0.1,
                                    "blue": 0.1,
                                },
                            },
                        }
                    },
                    "fields": "userEnteredFormat.textFormat",
                }
            }
        )

    return reqs


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------


def create_crossword_sheet(spreadsheet_id, puzzle_data):
    """Create a new worksheet tab with the crossword puzzle.

    Args:
        spreadsheet_id: Google Sheets spreadsheet ID.
        puzzle_data: dict returned by puzzle_downloader.parse_puzzle().

    Returns:
        The created gspread Worksheet, or None if it already existed.
    """
    client = _get_client()
    spreadsheet = client.open_by_key(spreadsheet_id)

    w = puzzle_data["width"]
    h = puzzle_data["height"]
    clue_col = GRID_START_COL + w + CLUE_GAP_COLS

    clue_lines = _build_clue_lines(puzzle_data)
    total_rows = max(GRID_START_ROW + h + 5, len(clue_lines) + GRID_START_ROW + 5)
    total_cols = clue_col + 5

    sheet_name = f"{puzzle_data['date']} {puzzle_data['outlet']}"

    # Skip if this puzzle tab already exists
    try:
        spreadsheet.worksheet(sheet_name)
        logger.info(f"Worksheet '{sheet_name}' already exists — skipping")
        return None
    except gspread.exceptions.WorksheetNotFound:
        pass

    worksheet = spreadsheet.add_worksheet(
        title=sheet_name, rows=total_rows, cols=total_cols
    )
    sheet_id = worksheet.id

    # ---- Write plain values ------------------------------------------------
    # Title & author
    worksheet.update(
        [[puzzle_data["title"]], [f"By {puzzle_data['author']}"]],
        f"B1:B2",
        value_input_option="RAW",
    )

    # Clues
    clue_cell = gspread.utils.rowcol_to_a1(GRID_START_ROW + 1, clue_col + 1)
    worksheet.update(
        [[line] for line in clue_lines],
        clue_cell,
        value_input_option="RAW",
    )

    # ---- Build & execute batch formatting ----------------------------------
    clue_lookup = _build_clue_lookup(puzzle_data)
    grid_rows = _build_grid_rows(puzzle_data, clue_lookup)

    requests = _formatting_requests(sheet_id, puzzle_data, clue_lines)

    # Grid cells — values + formatting + notes in one request
    requests.append(
        {
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": GRID_START_ROW,
                    "endRowIndex": GRID_START_ROW + h,
                    "startColumnIndex": GRID_START_COL,
                    "endColumnIndex": GRID_START_COL + w,
                },
                "rows": grid_rows,
                "fields": "userEnteredValue,userEnteredFormat,note",
            }
        }
    )

    spreadsheet.batch_update({"requests": requests})

    logger.info(f"Created worksheet '{sheet_name}'")
    return worksheet
