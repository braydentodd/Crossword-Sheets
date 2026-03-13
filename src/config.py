"""Centralized configuration for Crossword-Sheets.

Every visual constant, color, font, and layout value lives here.
No magic numbers anywhere else.
"""

# ---------------------------------------------------------------------------
# Colors (Google Sheets RGB float format, 0.0–1.0)
# ---------------------------------------------------------------------------
BLACK = {"red": 0, "green": 0, "blue": 0}
WHITE = {"red": 1, "green": 1, "blue": 1}

# ---------------------------------------------------------------------------
# Font
# ---------------------------------------------------------------------------
FONT_FAMILY = "Times New Roman"        # Headers and clue panel text
COMPANION_FONT_FAMILY = "Playfair Display"  # Clue-number cells in the grid
MAIN_FONT_FAMILY = "Kalam"              # Main letter-entry cells

# ---------------------------------------------------------------------------
# Grid layout  (0-indexed sheet coordinates)
# ---------------------------------------------------------------------------
GUTTER_COL_PX = 10          # Column A – narrow left gutter
GRID_START_ROW = 2          # 0-indexed row where the grid begins (row 3)
GRID_START_COL = 1          # 0-indexed col where companion cols begin (col B)

# Each logical crossword cell = [companion col] + [main col]
# Each logical crossword row  = 2 sheet rows, merged vertically
COMPANION_COL_PX = 15       # Width of the number-label column
MAIN_COL_PX = 21            # Width of the letter-entry column
ROW_HEIGHT_PX = 18          # Height of each sub-row (2 per logical cell)

# ---------------------------------------------------------------------------
# Grid cell text
# ---------------------------------------------------------------------------
COMPANION_FONT_SIZE = 7     # Clue-number superscript in companion cell
COMPANION_BOLD = False
COMPANION_V_ALIGN = "TOP"
COMPANION_H_ALIGN = "LEFT"

MAIN_FONT_SIZE = 15         # Letter entry in main cell
MAIN_BOLD = True
MAIN_V_ALIGN = "MIDDLE"
MAIN_H_ALIGN = "CENTER"

# ---------------------------------------------------------------------------
# Borders
# ---------------------------------------------------------------------------
BORDER_WEIGHT = 2
BORDER_STYLE = "SOLID"

# ---------------------------------------------------------------------------
# Header (merged rows 0-1 above the grid, spanning the grid columns)
# ---------------------------------------------------------------------------
HEADER_ROW = 0              # 0-indexed
HEADER_ROWS = 2             # Header spans 2 rows (merged)
HEADER_FONT_SIZE = 12
HEADER_BOLD = True
HEADER_UNDERLINE = True
HEADER_FG = WHITE
HEADER_BG = BLACK
HEADER_V_ALIGN = "MIDDLE"
HEADER_H_ALIGN = "CENTER"
HEADER_ROW_HEIGHT_PX = 18   # Each sub-row of the header

# ---------------------------------------------------------------------------
# Tab color  (per outlet; falls back to black for unknown outlets)
# ---------------------------------------------------------------------------
OUTLET_TAB_COLORS = {
    "Washington Post":        {"red": 0,    "green": 0,    "blue": 0   },  # black
    "Washington Post Sunday":  {"red": 0,    "green": 0,    "blue": 0   },  # black
    "USA Today":       {"red": 0.0,  "green": 0.45, "blue": 0.8 },  # blue
    "Universal":       {"red": 1.0,  "green": 0.5,  "blue": 0.0 },  # orange
    "LA Times":        {"red": 1.0,  "green": 1.0,  "blue": 1.0 },  # white
    "Simply Daily Puzzles": {"red": 0.2, "green": 0.7, "blue": 0.2},  # green
    "The Daily Beast": {"red": 0.8,  "green": 0.0,  "blue": 0.0 },  # red
}

# ---------------------------------------------------------------------------
# Clue panel (side-by-side ACROSS + DOWN, to the right of the grid)
# ---------------------------------------------------------------------------
CLUE_GAP_COLS = 1           # Empty columns between grid and clue panel
CLUE_NUM_COL_PX = 30        # Width of each clue-number column
CLUE_TEXT_COL_PX = 250      # Initial width (auto-resized after content is committed)

CLUE_HEADER_FONT_SIZE = 12
CLUE_HEADER_BOLD = True
CLUE_HEADER_UNDERLINE = False
CLUE_HEADER_FG = WHITE
CLUE_HEADER_BG = BLACK
CLUE_HEADER_ROWS = 2        # ACROSS / DOWN headers span 2 rows (merged)

CLUE_FONT_SIZE = 10
CLUE_FONT_BOLD = False
CLUE_FG = BLACK
CLUE_BG = WHITE

CLUE_ROW_HEIGHT_PX = 18     # Clue rows match the sub-row height
CLUE_BORDER_WEIGHT = 1
CLUE_BORDER_STYLE = "SOLID"
CLUE_BORDER_COLOR = BLACK

# ---------------------------------------------------------------------------
# Notepad (single column to the right of the clue panels)
# ---------------------------------------------------------------------------
NOTEPAD_COL_PX = 600        # Width of the notepad column