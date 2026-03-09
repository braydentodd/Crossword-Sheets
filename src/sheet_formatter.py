"""Format and populate a Google Sheet with a crossword puzzle grid.

Layout per logical crossword cell (2 sheet columns):
  ┌─────────┬──────────────────┐
  │companion│      main        │  ← one sheet row
  │  15 px  │     20 px        │
  │ clue #  │  letter entry    │
  └─────────┴──────────────────┘
  Together they form a ~35×35 visual square.

All visual constants are imported from config.py.
"""

import json
import logging
import os

import gspread
from google.oauth2.service_account import Credentials

import config as C

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ---------------------------------------------------------------------------
# Google Auth
# ---------------------------------------------------------------------------

def _get_client():
    """Authenticate with Google and return a gspread client."""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(creds)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _border(weight=C.BORDER_WEIGHT, style=C.BORDER_STYLE, color=None):
    return {"style": style, "width": weight, "color": color or C.BLACK}


def _no_border():
    return {"style": "NONE", "width": 0}


def _cell_pair_borders():
    """Borders for a companion+main pair.

    companion: top, bottom, left solid — right NONE (interior seam hidden).
    main:      top, bottom, right solid — left NONE (interior seam hidden).
    """
    b = _border()
    companion = {"top": b, "bottom": b, "left": b, "right": _no_border()}
    main = {"top": b, "bottom": b, "left": _no_border(), "right": b}
    return companion, main


def _black_cell_borders():
    """Borders for a fully-black companion+main pair."""
    b = _border()
    companion = {"top": b, "bottom": b, "left": b, "right": _no_border()}
    main = {"top": b, "bottom": b, "left": _no_border(), "right": b}
    return companion, main


def _sheet_col(grid_col):
    """Convert a logical grid column index to the sheet column of the companion cell."""
    return C.GRID_START_COL + grid_col * 2


def _sheet_col_main(grid_col):
    """Convert a logical grid column index to the sheet column of the main cell."""
    return C.GRID_START_COL + grid_col * 2 + 1


# ---------------------------------------------------------------------------
# Word membership map  (used by Apps Script for highlight)
# ---------------------------------------------------------------------------

def _build_word_map(puzzle_data):
    """Build across/down word membership for every white cell.

    Returns a dict keyed by "row,col" (grid coords) -> {
        "across": {"clueNum": N, "cells": [[r,c], ...]},
        "down":   {"clueNum": N, "cells": [[r,c], ...]},
    }
    """
    w = puzzle_data["width"]
    grid = puzzle_data["grid"]
    word_map = {}

    # Across words
    for clue in puzzle_data["across_clues"]:
        num = clue["num"]
        length = clue["length"]
        for r, row in enumerate(grid):
            for c_idx, cell in enumerate(row):
                if cell["number"] == num:
                    cells = []
                    for offset in range(length):
                        cc = c_idx + offset
                        if cc < w and not grid[r][cc]["is_black"]:
                            cells.append([r, cc])
                    for coord in cells:
                        key = f"{coord[0]},{coord[1]}"
                        word_map.setdefault(key, {})
                        word_map[key]["across"] = {"clueNum": num, "cells": cells}
                    break
            else:
                continue
            break

    # Down words
    h = puzzle_data["height"]
    for clue in puzzle_data["down_clues"]:
        num = clue["num"]
        length = clue["length"]
        for r, row in enumerate(grid):
            for c_idx, cell in enumerate(row):
                if cell["number"] == num:
                    cells = []
                    for offset in range(length):
                        rr = r + offset
                        if rr < h and not grid[rr][c_idx]["is_black"]:
                            cells.append([rr, c_idx])
                    for coord in cells:
                        key = f"{coord[0]},{coord[1]}"
                        word_map.setdefault(key, {})
                        word_map[key]["down"] = {"clueNum": num, "cells": cells}
                    break
            else:
                continue
            break

    return word_map


def _build_clue_row_index(puzzle_data, clue_start_row):
    """Map clue identifiers to their sheet row for highlighting.

    Returns {"A<num>": sheetRow, "D<num>": sheetRow, ...}
    """
    idx = {}
    row = clue_start_row       # "ACROSS" header
    row += 1                   # blank line after header
    for c in puzzle_data["across_clues"]:
        idx[f"A{c['num']}"] = row
        row += 1
    row += 1  # blank separator
    row += 1  # "DOWN" header
    row += 1  # blank line after header
    for c in puzzle_data["down_clues"]:
        idx[f"D{c['num']}"] = row
        row += 1
    return idx


# ---------------------------------------------------------------------------
# Grid rows (updateCells payload)
# ---------------------------------------------------------------------------

def _build_grid_rows(puzzle_data):
    """Return RowData dicts for the crossword grid (companion + main per cell)."""
    comp_borders, main_borders = _cell_pair_borders()
    black_comp_b, black_main_b = _black_cell_borders()
    rows = []

    for grid_row in puzzle_data["grid"]:
        cells = []
        for cell in grid_row:
            if cell["is_black"]:
                # Companion — solid black
                cells.append({
                    "userEnteredFormat": {
                        "backgroundColor": C.BLACK,
                        "borders": black_comp_b,
                    }
                })
                # Main — solid black
                cells.append({
                    "userEnteredFormat": {
                        "backgroundColor": C.BLACK,
                        "borders": black_main_b,
                    }
                })
            else:
                # --- Companion cell (clue number) ---
                comp = {
                    "userEnteredFormat": {
                        "backgroundColor": C.WHITE,
                        "borders": comp_borders,
                        "verticalAlignment": C.COMPANION_V_ALIGN,
                        "horizontalAlignment": C.COMPANION_H_ALIGN,
                        "textFormat": {
                            "fontFamily": C.FONT_FAMILY,
                            "fontSize": C.COMPANION_FONT_SIZE,
                            "bold": C.COMPANION_BOLD,
                            "foregroundColor": C.BLACK,
                        },
                    }
                }
                if cell["number"] is not None:
                    comp["userEnteredValue"] = {"stringValue": str(cell["number"])}
                cells.append(comp)

                # --- Main cell (letter entry) ---
                main = {
                    "userEnteredFormat": {
                        "backgroundColor": C.WHITE,
                        "borders": main_borders,
                        "verticalAlignment": C.MAIN_V_ALIGN,
                        "horizontalAlignment": C.MAIN_H_ALIGN,
                        "textFormat": {
                            "fontFamily": C.FONT_FAMILY,
                            "fontSize": C.MAIN_FONT_SIZE,
                            "bold": C.MAIN_BOLD,
                            "foregroundColor": C.BLACK,
                        },
                    }
                }
                cells.append(main)

        rows.append({"values": cells})
    return rows


# ---------------------------------------------------------------------------
# Clue panel data
# ---------------------------------------------------------------------------

def _build_clue_data(puzzle_data):
    """Return list of (num_str, text) tuples including section headers."""
    rows = []
    rows.append(("", "ACROSS"))
    rows.append(("", ""))
    for c in puzzle_data["across_clues"]:
        rows.append((str(c["num"]), f"{c['clue']} ({c['length']})"))
    rows.append(("", ""))
    rows.append(("", "DOWN"))
    rows.append(("", ""))
    for c in puzzle_data["down_clues"]:
        rows.append((str(c["num"]), f"{c['clue']} ({c['length']})"))
    return rows


def _clue_section_header_rows(puzzle_data):
    """Return 0-based offsets within the clue data for ACROSS and DOWN headers."""
    across_offset = 0
    down_offset = 2 + len(puzzle_data["across_clues"]) + 1
    return [across_offset, down_offset]


# ---------------------------------------------------------------------------
# Formatting requests
# ---------------------------------------------------------------------------

def _dimension_requests(sheet_id, puzzle_data, clue_data):
    """Column widths, row heights, gridline hiding."""
    w = puzzle_data["width"]
    h = puzzle_data["height"]
    total_sheet_cols = C.GRID_START_COL + w * 2
    clue_num_col = total_sheet_cols + C.CLUE_GAP_COLS
    clue_text_col = clue_num_col + 1

    reqs = []

    # Hide default gridlines
    reqs.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"hideGridlines": True},
            },
            "fields": "gridProperties.hideGridlines",
        }
    })

    # Tab color
    reqs.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "tabColor": C.TAB_COLOR},
            "fields": "tabColor",
        }
    })

    # Column A gutter
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": C.GUTTER_COL_PX},
            "fields": "pixelSize",
        }
    })

    # Companion columns widths
    for gc in range(w):
        col = _sheet_col(gc)
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": col, "endIndex": col + 1},
                "properties": {"pixelSize": C.COMPANION_COL_PX},
                "fields": "pixelSize",
            }
        })

    # Main columns widths
    for gc in range(w):
        col = _sheet_col_main(gc)
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": col, "endIndex": col + 1},
                "properties": {"pixelSize": C.MAIN_COL_PX},
                "fields": "pixelSize",
            }
        })

    # Grid row heights
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": C.GRID_START_ROW,
                      "endIndex": C.GRID_START_ROW + h},
            "properties": {"pixelSize": C.CELL_HEIGHT_PX},
            "fields": "pixelSize",
        }
    })

    # Header row height
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": C.HEADER_ROW, "endIndex": C.HEADER_ROW + 1},
            "properties": {"pixelSize": C.HEADER_ROW_HEIGHT_PX},
            "fields": "pixelSize",
        }
    })

    # Spacer row height
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": C.SPACER_ROW, "endIndex": C.SPACER_ROW + 1},
            "properties": {"pixelSize": C.SPACER_ROW_HEIGHT_PX},
            "fields": "pixelSize",
        }
    })

    # Gap columns between grid and clue panel
    for g in range(C.CLUE_GAP_COLS):
        col = total_sheet_cols + g
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": col, "endIndex": col + 1},
                "properties": {"pixelSize": 10},
                "fields": "pixelSize",
            }
        })

    # Clue number column width
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": clue_num_col, "endIndex": clue_num_col + 1},
            "properties": {"pixelSize": C.CLUE_NUM_COL_PX},
            "fields": "pixelSize",
        }
    })

    # Clue text column width
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": clue_text_col, "endIndex": clue_text_col + 1},
            "properties": {"pixelSize": C.CLUE_TEXT_COL_PX},
            "fields": "pixelSize",
        }
    })

    # Clue row heights
    clue_start = C.GRID_START_ROW + C.CLUE_PANEL_START_ROW_OFFSET
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": clue_start,
                      "endIndex": clue_start + len(clue_data)},
            "properties": {"pixelSize": C.CLUE_ROW_HEIGHT_PX},
            "fields": "pixelSize",
        }
    })

    return reqs


def _header_requests(sheet_id, puzzle_data):
    """Merged header bar spanning the full grid width."""
    w = puzzle_data["width"]
    start_col = C.GRID_START_COL
    end_col = C.GRID_START_COL + w * 2

    reqs = []

    reqs.append({
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": C.HEADER_ROW,
                "endRowIndex": C.HEADER_ROW + 1,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col,
            },
            "mergeType": "MERGE_ALL",
        }
    })

    reqs.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": C.HEADER_ROW,
                "endRowIndex": C.HEADER_ROW + 1,
                "startColumnIndex": start_col,
                "endColumnIndex": end_col,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": C.HEADER_BG,
                    "verticalAlignment": C.HEADER_V_ALIGN,
                    "horizontalAlignment": C.HEADER_H_ALIGN,
                    "textFormat": {
                        "fontFamily": C.FONT_FAMILY,
                        "fontSize": C.HEADER_FONT_SIZE,
                        "bold": C.HEADER_BOLD,
                        "underline": C.HEADER_UNDERLINE,
                        "foregroundColor": C.HEADER_FG,
                    },
                }
            },
            "fields": "userEnteredFormat",
        }
    })

    return reqs


def _clue_panel_requests(sheet_id, puzzle_data, clue_data):
    """Formatting for the 2-column clue panel."""
    w = puzzle_data["width"]
    total_sheet_cols = C.GRID_START_COL + w * 2
    clue_num_col = total_sheet_cols + C.CLUE_GAP_COLS
    clue_text_col = clue_num_col + 1
    clue_start_row = C.GRID_START_ROW + C.CLUE_PANEL_START_ROW_OFFSET

    reqs = []
    section_hdr_offsets = _clue_section_header_rows(puzzle_data)

    # Base formatting for all clue rows (both columns)
    for col in [clue_num_col, clue_text_col]:
        reqs.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": clue_start_row,
                    "endRowIndex": clue_start_row + len(clue_data),
                    "startColumnIndex": col,
                    "endColumnIndex": col + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": C.CLUE_BG,
                        "textFormat": {
                            "fontFamily": C.FONT_FAMILY,
                            "fontSize": C.CLUE_FONT_SIZE,
                            "bold": C.CLUE_FONT_BOLD,
                            "foregroundColor": C.CLUE_FG,
                        },
                        "verticalAlignment": "MIDDLE",
                        "wrapStrategy": "WRAP",
                        "borders": {
                            "bottom": _border(C.CLUE_BORDER_WEIGHT,
                                              C.CLUE_BORDER_STYLE,
                                              C.CLUE_BORDER_COLOR),
                        },
                    }
                },
                "fields": "userEnteredFormat",
            }
        })

    # Number column: right-align + bold
    reqs.append({
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": clue_start_row,
                "endRowIndex": clue_start_row + len(clue_data),
                "startColumnIndex": clue_num_col,
                "endColumnIndex": clue_num_col + 1,
            },
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "RIGHT",
                    "textFormat": {"bold": True},
                }
            },
            "fields": "userEnteredFormat(horizontalAlignment,textFormat.bold)",
        }
    })

    # Section headers (ACROSS / DOWN): black bg, white bold underline, merge 2 cols
    for offset in section_hdr_offsets:
        row = clue_start_row + offset
        reqs.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row,
                    "endRowIndex": row + 1,
                    "startColumnIndex": clue_num_col,
                    "endColumnIndex": clue_text_col + 1,
                },
                "mergeType": "MERGE_ALL",
            }
        })
        reqs.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row,
                    "endRowIndex": row + 1,
                    "startColumnIndex": clue_num_col,
                    "endColumnIndex": clue_text_col + 1,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": C.CLUE_HEADER_BG,
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE",
                        "textFormat": {
                            "fontFamily": C.FONT_FAMILY,
                            "fontSize": C.CLUE_HEADER_FONT_SIZE,
                            "bold": C.CLUE_HEADER_BOLD,
                            "underline": C.CLUE_HEADER_UNDERLINE,
                            "foregroundColor": C.CLUE_HEADER_FG,
                        },
                    }
                },
                "fields": "userEnteredFormat",
            }
        })

    # Outer border on the clue panel
    b = _border(C.CLUE_BORDER_WEIGHT, C.CLUE_BORDER_STYLE, C.CLUE_BORDER_COLOR)
    reqs.append({
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": clue_start_row,
                "endRowIndex": clue_start_row + len(clue_data),
                "startColumnIndex": clue_num_col,
                "endColumnIndex": clue_text_col + 1,
            },
            "left": b, "right": b, "top": b, "bottom": b,
        }
    })

    return reqs


# ---------------------------------------------------------------------------
# Protection (companion cells + black squares uneditable)
# ---------------------------------------------------------------------------

def _protection_requests(sheet_id, puzzle_data):
    """Protect companion columns and black squares from editing."""
    w = puzzle_data["width"]
    h = puzzle_data["height"]
    reqs = []

    # Protect all companion columns
    for gc in range(w):
        col = _sheet_col(gc)
        reqs.append({
            "addProtectedRange": {
                "protectedRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": C.GRID_START_ROW,
                        "endRowIndex": C.GRID_START_ROW + h,
                        "startColumnIndex": col,
                        "endColumnIndex": col + 1,
                    },
                    "description": f"Companion col {gc}",
                    "warningOnly": True,
                }
            }
        })

    # Protect black main cells
    for r, grid_row in enumerate(puzzle_data["grid"]):
        for gc, cell in enumerate(grid_row):
            if cell["is_black"]:
                main_col = _sheet_col_main(gc)
                reqs.append({
                    "addProtectedRange": {
                        "protectedRange": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": C.GRID_START_ROW + r,
                                "endRowIndex": C.GRID_START_ROW + r + 1,
                                "startColumnIndex": main_col,
                                "endColumnIndex": main_col + 1,
                            },
                            "description": f"Black cell ({r},{gc})",
                            "warningOnly": True,
                        }
                    }
                })

    return reqs


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def create_crossword_sheet(spreadsheet_id, puzzle_data):
    """Create a new worksheet tab with the crossword puzzle.

    Returns (worksheet, script_config) or None if tab already existed.
    """
    client = _get_client()
    spreadsheet = client.open_by_key(spreadsheet_id)

    w = puzzle_data["width"]
    h = puzzle_data["height"]

    total_grid_cols = w * 2
    clue_num_col = C.GRID_START_COL + total_grid_cols + C.CLUE_GAP_COLS
    clue_text_col = clue_num_col + 1

    clue_data = _build_clue_data(puzzle_data)
    clue_start_row = C.GRID_START_ROW + C.CLUE_PANEL_START_ROW_OFFSET

    total_rows = max(C.GRID_START_ROW + h + 5, clue_start_row + len(clue_data) + 5)
    total_cols = clue_text_col + 3

    sheet_name = f"{puzzle_data['date']} {puzzle_data['outlet']}"

    # Skip if tab already exists
    try:
        spreadsheet.worksheet(sheet_name)
        logger.info(f"Worksheet '{sheet_name}' already exists — skipping")
        return None
    except gspread.exceptions.WorksheetNotFound:
        pass

    worksheet = spreadsheet.add_worksheet(
        title=sheet_name, rows=total_rows, cols=total_cols,
    )
    sheet_id = worksheet.id

    # ---- Header value -------------------------------------------------------
    header_cell = gspread.utils.rowcol_to_a1(C.HEADER_ROW + 1, C.GRID_START_COL + 1)
    header_text = (
        f"{puzzle_data['outlet']} — {puzzle_data['date']} — {puzzle_data['author']}"
    )
    worksheet.update([[header_text]], header_cell, value_input_option="RAW")

    # ---- Clue values (2-column) ---------------------------------------------
    clue_values = [[num_str, text] for num_str, text in clue_data]
    clue_origin = gspread.utils.rowcol_to_a1(clue_start_row + 1, clue_num_col + 1)
    worksheet.update(clue_values, clue_origin, value_input_option="RAW")

    # ---- Batch formatting ---------------------------------------------------
    requests = []
    requests.extend(_dimension_requests(sheet_id, puzzle_data, clue_data))
    requests.extend(_header_requests(sheet_id, puzzle_data))
    requests.extend(_clue_panel_requests(sheet_id, puzzle_data, clue_data))
    requests.extend(_protection_requests(sheet_id, puzzle_data))

    grid_rows = _build_grid_rows(puzzle_data)
    requests.append({
        "updateCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": C.GRID_START_ROW,
                "endRowIndex": C.GRID_START_ROW + h,
                "startColumnIndex": C.GRID_START_COL,
                "endColumnIndex": C.GRID_START_COL + w * 2,
            },
            "rows": grid_rows,
            "fields": "userEnteredValue,userEnteredFormat",
        }
    })

    spreadsheet.batch_update({"requests": requests})

    # ---- Build script config for Apps Script --------------------------------
    word_map = _build_word_map(puzzle_data)
    clue_row_index = _build_clue_row_index(puzzle_data, clue_start_row)

    script_config = {
        "gridStartRow": C.GRID_START_ROW,
        "gridStartCol": C.GRID_START_COL,
        "gridWidth": w,
        "gridHeight": h,
        "clueNumCol": clue_num_col,
        "clueTextCol": clue_text_col,
        "clueStartRow": clue_start_row,
        "clueDataLen": len(clue_data),
        "highlightWord": _color_to_hex(C.HIGHLIGHT_WORD),
        "highlightClue": _color_to_hex(C.HIGHLIGHT_CLUE),
        "whiteHex": "#ffffff",
        "clueBgHex": _color_to_hex(C.CLUE_BG),
        "wordMap": word_map,
        "clueRowIndex": clue_row_index,
        "sheetName": sheet_name,
    }

    logger.info(f"Created worksheet '{sheet_name}'")

    return worksheet, script_config


def _color_to_hex(c):
    """Convert a Google Sheets RGB float dict to a hex string."""
    r = int(c["red"] * 255)
    g = int(c["green"] * 255)
    b = int(c["blue"] * 255)
    return f"#{r:02x}{g:02x}{b:02x}"
