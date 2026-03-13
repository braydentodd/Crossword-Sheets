"""Format and populate a Google Sheet with a crossword puzzle grid.

Layout per logical crossword cell (2 sheet columns, 2 sheet rows merged):
  ┌─────────┬──────────────────┐
  │companion│      main        │  ← 2 sheet rows merged (18px each = 36px)
  │  15 px  │     21 px        │
  │ clue #  │  letter entry    │
  └─────────┴──────────────────┘

Clue panel: side-by-side ACROSS and DOWN columns to the right of the grid.
No Apps Script — pure Google Sheets API formatting.

All visual constants are imported from config.py.
"""

import json
import logging
import os
import re
from collections import defaultdict
from datetime import datetime
from itertools import groupby

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
    """Authenticate with Google and return (gspread_client, service_email)."""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_info = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    service_email = getattr(creds, 'service_account_email', '') or ''
    return gspread.authorize(creds), service_email


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


def _sheet_row(grid_row):
    """Convert a logical grid row index to the first of its two sheet rows."""
    return C.GRID_START_ROW + grid_row * 2


# ---------------------------------------------------------------------------
# Grid rows (updateCells payload — 2 sheet rows per logical row)
# ---------------------------------------------------------------------------

def _build_grid_rows(puzzle_data):
    """Return RowData dicts for the crossword grid.

    Each logical grid row produces TWO sheet rows (they will be merged later).
    Only the first sheet row carries cell values/formatting; the second is empty
    so the merge inherits the first row's format.
    """
    comp_borders, main_borders = _cell_pair_borders()
    black_comp_b, black_main_b = _black_cell_borders()
    rows = []

    for grid_row in puzzle_data["grid"]:
        # --- First sub-row: carries values + formatting ---
        cells_top = []
        for cell in grid_row:
            if cell["is_black"]:
                cells_top.append({
                    "userEnteredFormat": {
                        "backgroundColor": C.BLACK,
                        "borders": black_comp_b,
                    }
                })
                cells_top.append({
                    "userEnteredFormat": {
                        "backgroundColor": C.BLACK,
                        "borders": black_main_b,
                    }
                })
            else:
                comp = {
                    "userEnteredFormat": {
                        "backgroundColor": C.WHITE,
                        "borders": comp_borders,
                        "verticalAlignment": C.COMPANION_V_ALIGN,
                        "horizontalAlignment": C.COMPANION_H_ALIGN,
                        "textFormat": {
                            "fontFamily": C.COMPANION_FONT_FAMILY,
                            "fontSize": C.COMPANION_FONT_SIZE,
                            "bold": C.COMPANION_BOLD,
                            "foregroundColor": C.BLACK,
                        },
                    }
                }
                if cell["number"] is not None:
                    comp["userEnteredValue"] = {"numberValue": cell["number"]}
                cells_top.append(comp)

                main = {
                    "userEnteredFormat": {
                        "backgroundColor": C.WHITE,
                        "borders": main_borders,
                        "verticalAlignment": C.MAIN_V_ALIGN,
                        "horizontalAlignment": C.MAIN_H_ALIGN,
                        "textFormat": {
                            "fontFamily": C.MAIN_FONT_FAMILY,
                            "fontSize": C.MAIN_FONT_SIZE,
                            "bold": C.MAIN_BOLD,
                            "foregroundColor": C.BLACK,
                        },
                    }
                }
                cells_top.append(main)

        rows.append({"values": cells_top})

        # --- Second sub-row: empty (will be merged with first) ---
        cells_bot = []
        for cell in grid_row:
            if cell["is_black"]:
                cells_bot.append({
                    "userEnteredFormat": {
                        "backgroundColor": C.BLACK,
                        "borders": black_comp_b,
                    }
                })
                cells_bot.append({
                    "userEnteredFormat": {
                        "backgroundColor": C.BLACK,
                        "borders": black_main_b,
                    }
                })
            else:
                cells_bot.append({
                    "userEnteredFormat": {
                        "backgroundColor": C.WHITE,
                        "borders": comp_borders,
                    }
                })
                cells_bot.append({
                    "userEnteredFormat": {
                        "backgroundColor": C.WHITE,
                        "borders": main_borders,
                    }
                })
        rows.append({"values": cells_bot})

    return rows


# ---------------------------------------------------------------------------
# Merge requests (2 sheet rows per logical grid row)
# ---------------------------------------------------------------------------

def _grid_merge_requests(sheet_id, puzzle_data):
    """Merge every companion and main cell across their 2 sub-rows."""
    w = puzzle_data["width"]
    h = puzzle_data["height"]
    reqs = []

    for gr in range(h):
        sr = _sheet_row(gr)
        for gc in range(w):
            # Merge companion cell (2 rows)
            comp_col = _sheet_col(gc)
            reqs.append({
                "mergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": sr,
                        "endRowIndex": sr + 2,
                        "startColumnIndex": comp_col,
                        "endColumnIndex": comp_col + 1,
                    },
                    "mergeType": "MERGE_ALL",
                }
            })
            # Merge main cell (2 rows)
            main_col = _sheet_col_main(gc)
            reqs.append({
                "mergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": sr,
                        "endRowIndex": sr + 2,
                        "startColumnIndex": main_col,
                        "endColumnIndex": main_col + 1,
                    },
                    "mergeType": "MERGE_ALL",
                }
            })

    return reqs


# ---------------------------------------------------------------------------
# Clue data (side-by-side ACROSS + DOWN)
# ---------------------------------------------------------------------------

def _build_across_clues(puzzle_data):
    """Return list of (num_str, text) tuples for ACROSS clues only."""
    return [(str(c["num"]), c["clue"]) for c in puzzle_data["across_clues"]]


def _build_down_clues(puzzle_data):
    """Return list of (num_str, text) tuples for DOWN clues only."""
    return [(str(c["num"]), c["clue"]) for c in puzzle_data["down_clues"]]


# ---------------------------------------------------------------------------
# Formatting requests
# ---------------------------------------------------------------------------

def _dimension_requests(sheet_id, puzzle_data, across_clues, down_clues):
    """Column widths, row heights, gridline hiding, column freezing."""
    w = puzzle_data["width"]
    h = puzzle_data["height"]
    total_grid_cols = C.GRID_START_COL + w * 2
    grid_sheet_rows = h * 2  # 2 sub-rows per logical row

    # Clue column positions
    across_num_col = total_grid_cols + C.CLUE_GAP_COLS
    across_text_col = across_num_col + 1
    down_num_col = across_text_col + 1
    down_text_col = down_num_col + 1

    max_clue_rows = max(len(across_clues), len(down_clues))
    clue_start_row = C.HEADER_ROW + C.CLUE_HEADER_ROWS  # clues start after header

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

    # Tab color — keyed by outlet name
    tab_color = C.OUTLET_TAB_COLORS.get(puzzle_data["outlet"], C.BLACK)
    reqs.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "tabColor": tab_color},
            "fields": "tabColor",
        }
    })

    # Freeze all grid columns (gutter + grid companion/main pairs)
    reqs.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenColumnCount": total_grid_cols},
            },
            "fields": "gridProperties.frozenColumnCount",
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

    # Companion column widths
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

    # Main column widths
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

    # Header row heights (rows 0 and 1 — each 18px, merged later)
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": C.HEADER_ROW,
                      "endIndex": C.HEADER_ROW + C.HEADER_ROWS},
            "properties": {"pixelSize": C.HEADER_ROW_HEIGHT_PX},
            "fields": "pixelSize",
        }
    })

    # Grid sub-row heights (all 18px each)
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": C.GRID_START_ROW,
                      "endIndex": C.GRID_START_ROW + grid_sheet_rows},
            "properties": {"pixelSize": C.ROW_HEIGHT_PX},
            "fields": "pixelSize",
        }
    })

    # Gap column(s) between grid and clue panel
    for g in range(C.CLUE_GAP_COLS):
        col = total_grid_cols + g
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": col, "endIndex": col + 1},
                "properties": {"pixelSize": 10},
                "fields": "pixelSize",
            }
        })

    # Clue number column widths (ACROSS + DOWN)
    for col in [across_num_col, down_num_col]:
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": col, "endIndex": col + 1},
                "properties": {"pixelSize": C.CLUE_NUM_COL_PX},
                "fields": "pixelSize",
            }
        })

    # Clue text column widths (initial — auto-resized after content is committed)
    for col in [across_text_col, down_text_col]:
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": col, "endIndex": col + 1},
                "properties": {"pixelSize": C.CLUE_TEXT_COL_PX},
                "fields": "pixelSize",
            }
        })

    # Notepad: gap column + single notepad column
    notepad_gap_col = down_text_col + 1
    notepad_col = notepad_gap_col + 1

    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": notepad_gap_col, "endIndex": notepad_gap_col + 1},
            "properties": {"pixelSize": 10},
            "fields": "pixelSize",
        }
    })
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": notepad_col, "endIndex": notepad_col + 1},
            "properties": {"pixelSize": C.NOTEPAD_COL_PX},
            "fields": "pixelSize",
        }
    })

    # Clue row heights
    reqs.append({
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": clue_start_row,
                      "endIndex": clue_start_row + max_clue_rows},
            "properties": {"pixelSize": C.CLUE_ROW_HEIGHT_PX},
            "fields": "pixelSize",
        }
    })

    return reqs


def _header_requests(sheet_id, puzzle_data):
    """Merged header bar spanning the full grid width, 2 rows tall."""
    w = puzzle_data["width"]
    start_col = C.GRID_START_COL
    end_col = C.GRID_START_COL + w * 2

    reqs = []

    # Merge header across 2 rows and full grid width
    reqs.append({
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": C.HEADER_ROW,
                "endRowIndex": C.HEADER_ROW + C.HEADER_ROWS,
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
                "endRowIndex": C.HEADER_ROW + C.HEADER_ROWS,
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


def _clue_panel_requests(sheet_id, puzzle_data, across_clues, down_clues):
    """Formatting for side-by-side ACROSS + DOWN clue panels."""
    w = puzzle_data["width"]
    total_grid_cols = C.GRID_START_COL + w * 2

    across_num_col = total_grid_cols + C.CLUE_GAP_COLS
    across_text_col = across_num_col + 1
    down_num_col = across_text_col + 1
    down_text_col = down_num_col + 1
    notepad_col = down_text_col + 2       # +1 gap col, +1 = notepad col

    max_clue_rows = max(len(across_clues), len(down_clues))
    clue_start_row = C.HEADER_ROW + C.CLUE_HEADER_ROWS  # clues begin after header

    reqs = []

    # --- Section headers: merged across num+text cols, 2 rows, at row 0 ---
    for num_col, text_col, label in [
        (across_num_col, across_text_col, "ACROSS"),
        (down_num_col, down_text_col, "DOWN"),
        (notepad_col, notepad_col, "NOTEPAD"),
    ]:
        reqs.append({
            "mergeCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": C.HEADER_ROW,
                    "endRowIndex": C.HEADER_ROW + C.CLUE_HEADER_ROWS,
                    "startColumnIndex": num_col,
                    "endColumnIndex": text_col + 1,
                },
                "mergeType": "MERGE_ALL",
            }
        })
        reqs.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": C.HEADER_ROW,
                    "endRowIndex": C.HEADER_ROW + C.CLUE_HEADER_ROWS,
                    "startColumnIndex": num_col,
                    "endColumnIndex": text_col + 1,
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

    # --- Base formatting for all clue/notepad rows ---
    for num_col, text_col in [
        (across_num_col, across_text_col),
        (down_num_col, down_text_col),
        (notepad_col, notepad_col),
    ]:
        for col in [num_col, text_col]:
            reqs.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": clue_start_row,
                        "endRowIndex": clue_start_row + max_clue_rows,
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
                            "wrapStrategy": "CLIP",
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

        # Number columns: right-align + bold
        reqs.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": clue_start_row,
                    "endRowIndex": clue_start_row + max_clue_rows,
                    "startColumnIndex": num_col,
                    "endColumnIndex": num_col + 1,
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

        # Outer border on each clue section
        b = _border(C.CLUE_BORDER_WEIGHT, C.CLUE_BORDER_STYLE, C.CLUE_BORDER_COLOR)
        reqs.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": clue_start_row,
                    "endRowIndex": clue_start_row + max_clue_rows,
                    "startColumnIndex": num_col,
                    "endColumnIndex": text_col + 1,
                },
                "left": b, "right": b, "top": b, "bottom": b,
            }
        })

    return reqs


# ---------------------------------------------------------------------------
# Data validation (1-character limit on main cells)
# ---------------------------------------------------------------------------

def _validation_requests(sheet_id, puzzle_data):
    """Apply a 1-character data validation to every white main cell.

    Only white main cells get validation (LEN<=1, strict).
    Companion cells and black squares are blocked via sheet protection
    instead, to avoid the red 'Invalid' badge that =FALSE causes on
    cells that already contain content (e.g. clue numbers).
    """
    reqs = []
    for r, grid_row in enumerate(puzzle_data["grid"]):
        for gc, cell in enumerate(grid_row):
            if not cell["is_black"]:
                sr       = _sheet_row(r)
                main_col = _sheet_col_main(gc)
                a1 = gspread.utils.rowcol_to_a1(sr + 1, main_col + 1)  # 1-indexed
                reqs.append({
                    "setDataValidation": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex":    sr,
                            "endRowIndex":      sr + 2,
                            "startColumnIndex": main_col,
                            "endColumnIndex":   main_col + 1,
                        },
                        "rule": {
                            "condition": {
                                "type": "CUSTOM_FORMULA",
                                "values": [{"userEnteredValue": f"=LEN({a1})<=1"}],
                            },
                            "showCustomUi": False,
                            "strict": True,
                        },
                    }
                })
    return reqs


# ---------------------------------------------------------------------------
# Protection (companion cells + black squares fully locked)
# ---------------------------------------------------------------------------

def _protection_requests(sheet_id, puzzle_data, editor_email=None):
    """Protect companion columns and black squares — full lock, not warning-only.

    Sets explicit editors (only the service account) so that all other
    users — including spreadsheet editors — are blocked.
    """
    w = puzzle_data["width"]
    h = puzzle_data["height"]
    grid_sheet_rows = h * 2
    reqs = []

    # Protect all companion columns (full grid height)
    for gc in range(w):
        col = _sheet_col(gc)
        reqs.append({
            "addProtectedRange": {
                "protectedRange": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": C.GRID_START_ROW,
                        "endRowIndex": C.GRID_START_ROW + grid_sheet_rows,
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
                sr = _sheet_row(r)
                reqs.append({
                    "addProtectedRange": {
                        "protectedRange": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": sr,
                                "endRowIndex": sr + 2,
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

    Returns the worksheet object, or None if the tab already existed.
    """
    client, service_email = _get_client()
    spreadsheet = client.open_by_key(spreadsheet_id)

    w = puzzle_data["width"]
    h = puzzle_data["height"]

    grid_sheet_rows = h * 2  # 2 sub-rows per logical row
    total_grid_cols = C.GRID_START_COL + w * 2

    # Clue column positions (side-by-side ACROSS + DOWN)
    across_num_col = total_grid_cols + C.CLUE_GAP_COLS
    across_text_col = across_num_col + 1
    down_num_col = across_text_col + 1
    down_text_col = down_num_col + 1

    across_clues = _build_across_clues(puzzle_data)
    down_clues = _build_down_clues(puzzle_data)
    max_clue_rows = max(len(across_clues), len(down_clues))
    clue_start_row = C.HEADER_ROW + C.CLUE_HEADER_ROWS

    notepad_col = down_text_col + 2       # +1 gap col, +1 = notepad col

    total_rows = max(C.GRID_START_ROW + grid_sheet_rows + 5,
                     clue_start_row + max_clue_rows + 5)
    total_cols = notepad_col + 1

    # Convert YYYY-MM-DD → YY/MM/DD for the tab name
    raw_date = puzzle_data['date']  # e.g. "2026-03-11"
    short_date = datetime.strptime(raw_date, "%Y-%m-%d").strftime("%y/%m/%d")
    sheet_name = f"{short_date} - {puzzle_data['outlet']}"

    # Skip if tab already exists (possibly renamed with a status indicator like ✅ or ❌)
    for ws in spreadsheet.worksheets():
        base = re.sub(r'\s[✅❌]$', '', ws.title)
        if base == sheet_name:
            logger.info(f"Worksheet '{sheet_name}' already exists (as '{ws.title}') — skipping")
            return None

    worksheet = spreadsheet.add_worksheet(
        title=sheet_name, rows=total_rows, cols=total_cols,
    )
    sheet_id = worksheet.id

    # ---- Header value -------------------------------------------------------
    header_cell = gspread.utils.rowcol_to_a1(C.HEADER_ROW + 1, C.GRID_START_COL + 1)
    title  = puzzle_data.get("title", "")
    author = puzzle_data.get("author", "Unknown")
    if title:
        header_text = f"{title} by {author}"
    else:
        header_text = f"{puzzle_data['outlet']} ({puzzle_data['date']}) by {author}"
    worksheet.update([[header_text]], header_cell, value_input_option="RAW")

    # ---- Clue header values (ACROSS / DOWN at row 0) ------------------------
    across_hdr_cell = gspread.utils.rowcol_to_a1(C.HEADER_ROW + 1, across_num_col + 1)
    worksheet.update([["ACROSS"]], across_hdr_cell, value_input_option="RAW")

    down_hdr_cell = gspread.utils.rowcol_to_a1(C.HEADER_ROW + 1, down_num_col + 1)
    worksheet.update([["DOWN"]], down_hdr_cell, value_input_option="RAW")

    notepad_hdr_cell = gspread.utils.rowcol_to_a1(C.HEADER_ROW + 1, notepad_col + 1)
    worksheet.update([["NOTEPAD"]], notepad_hdr_cell, value_input_option="RAW")

    # ---- Clue values (side-by-side) -----------------------------------------
    # Build a combined grid: across on the left pair, down on the right pair
    clue_rows = []
    for i in range(max_clue_rows):
        row = []
        if i < len(across_clues):
            row.extend([across_clues[i][0], across_clues[i][1]])
        else:
            row.extend(["", ""])
        if i < len(down_clues):
            row.extend([down_clues[i][0], down_clues[i][1]])
        else:
            row.extend(["", ""])
        clue_rows.append(row)

    if clue_rows:
        clue_origin = gspread.utils.rowcol_to_a1(clue_start_row + 1,
                                                  across_num_col + 1)
        worksheet.update(clue_rows, clue_origin, value_input_option="RAW")

    # ---- Batch formatting ---------------------------------------------------
    requests = []
    requests.extend(_dimension_requests(sheet_id, puzzle_data,
                                        across_clues, down_clues))
    requests.extend(_header_requests(sheet_id, puzzle_data))
    requests.extend(_clue_panel_requests(sheet_id, puzzle_data,
                                         across_clues, down_clues))
    requests.extend(_validation_requests(sheet_id, puzzle_data))

    requests.extend(_protection_requests(sheet_id, puzzle_data, service_email))

    # Grid cell data (2 sheet rows per logical row)
    grid_rows = _build_grid_rows(puzzle_data)
    requests.append({
        "updateCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": C.GRID_START_ROW,
                "endRowIndex": C.GRID_START_ROW + grid_sheet_rows,
                "startColumnIndex": C.GRID_START_COL,
                "endColumnIndex": C.GRID_START_COL + w * 2,
            },
            "rows": grid_rows,
            "fields": "userEnteredValue,userEnteredFormat",
        }
    })

    # Grid cell merges (each companion + main spans 2 sub-rows)
    requests.extend(_grid_merge_requests(sheet_id, puzzle_data))

    spreadsheet.batch_update({"requests": requests})

    # Auto-resize clue text columns in a separate call so the content is fully
    # committed before the resize measures rendered widths.
    spreadsheet.batch_update({"requests": [
        {"autoResizeDimensions": {"dimensions": {
            "sheetId": sheet_id, "dimension": "COLUMNS",
            "startIndex": across_text_col, "endIndex": across_text_col + 1,
        }}},
        {"autoResizeDimensions": {"dimensions": {
            "sheetId": sheet_id, "dimension": "COLUMNS",
            "startIndex": down_text_col, "endIndex": down_text_col + 1,
        }}},
    ]})

    logger.info(f"Created worksheet '{sheet_name}'")

    return worksheet


# ---------------------------------------------------------------------------
# Tab management (sort by date + prune old tabs)
# ---------------------------------------------------------------------------

def sort_and_prune_tabs(spreadsheet_id, max_per_outlet=7):
    """Sort puzzle tabs by outlet then date (newest first) and keep only max_per_outlet per outlet.

    Outlet order: WP → LA Times → USA Today → Universal → SDP → Daily Beast → Walrus (others last).
    """
    OUTLET_ORDER = {
        "Washington Post Sunday": 0,
        "Washington Post": 1,
        "LA Times": 2,
        "USA Today": 3,
        "Universal": 4,
        "Simply Daily Puzzles": 5,
        "The Daily Beast": 6,
        "The Walrus": 7,
    }

    client, _ = _get_client()
    spreadsheet = client.open_by_key(spreadsheet_id)

    worksheets = spreadsheet.worksheets()

    puzzle_tabs = []   # (date_str, outlet, worksheet)
    system_tabs = []   # _xw_config, etc.

    for ws in worksheets:
        name = ws.title
        if name.startswith("_"):
            system_tabs.append(ws)
            continue
        # Strip status indicators for parsing
        base = re.sub(r'\s[✅❌]$', '', name)
        parts = base.split(" - ", 1)
        if len(parts) == 2 and re.match(r'^\d{2}/\d{2}/\d{2}$', parts[0]):
            puzzle_tabs.append((parts[0], parts[1], ws))
        else:
            system_tabs.append(ws)

    if not puzzle_tabs:
        return

    # Group by outlet
    by_outlet = defaultdict(list)
    for date_str, outlet, ws in puzzle_tabs:
        by_outlet[outlet].append((date_str, ws))

    to_delete = []
    to_keep = []   # (date_str, outlet, ws)
    for outlet, tabs in by_outlet.items():
        tabs.sort(key=lambda x: x[0], reverse=True)  # newest first
        to_keep.extend([(d, outlet, w) for d, w in tabs[:max_per_outlet]])
        to_delete.extend(tabs[max_per_outlet:])

    # Collect base names for _xw_config cleanup before deleting
    deleted_base_names = set()
    for _, ws in to_delete:
        deleted_base_names.add(re.sub(r'\s[✅❌]$', '', ws.title))

    for _, ws in to_delete:
        logger.info(f"Pruning old tab: {ws.title}")
        spreadsheet.del_worksheet(ws)

    # Reorder: grouped by outlet (WP → USA Today → Universal), newest-first within each
    ordered_tabs = []
    to_keep.sort(key=lambda x: OUTLET_ORDER.get(x[1], 99))
    for _key, grp in groupby(to_keep, key=lambda x: x[1]):
        items = sorted(grp, key=lambda x: x[0], reverse=True)
        ordered_tabs.extend(items)
    ordered = [ws for _, _, ws in ordered_tabs] + system_tabs
    if len(ordered) > 1:
        spreadsheet.reorder_worksheets(ordered)

    # Clean up _xw_config entries for deleted tabs
    if deleted_base_names:
        try:
            cfg_sheet = spreadsheet.worksheet("_xw_config")
            existing = cfg_sheet.get_all_values()
            rows = [r for r in existing if r[0] not in deleted_base_names]
            cfg_sheet.clear()
            if rows:
                cfg_sheet.update(rows, "A1", value_input_option="RAW")
        except Exception as e:
            logger.warning(f"Could not clean _xw_config: {e}")

    logger.info(f"Tab management: kept {len(to_keep)}, pruned {len(to_delete)}")
