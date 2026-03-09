"""Centralized configuration for Crossword-Sheets.

Every visual constant, color, font, and layout value lives here.
No magic numbers anywhere else.
"""

# ---------------------------------------------------------------------------
# Colors (Google Sheets RGB float format, 0.0–1.0)
# ---------------------------------------------------------------------------
BLACK = {"red": 0, "green": 0, "blue": 0}
WHITE = {"red": 1, "green": 1, "blue": 1}

HIGHLIGHT_WORD = {"red": 1, "green": 0.95, "blue": 0.6}       # light gold
HIGHLIGHT_CLUE = {"red": 0.85, "green": 0.92, "blue": 1}      # light blue

# ---------------------------------------------------------------------------
# Font
# ---------------------------------------------------------------------------
FONT_FAMILY = "Roboto Serif"

# ---------------------------------------------------------------------------
# Grid layout  (0-indexed sheet coordinates)
# ---------------------------------------------------------------------------
GUTTER_COL_PX = 10          # Column A – narrow left gutter
GRID_START_ROW = 2          # 0-indexed row where the grid begins (row 3)
GRID_START_COL = 1          # 0-indexed col where companion cols begin (col B)

# Each logical crossword cell = [companion col] + [main col]
COMPANION_COL_PX = 15       # Width of the number-label column
MAIN_COL_PX = 20            # Width of the letter-entry column
CELL_HEIGHT_PX = 35         # Row height (companion + main share the same row)

# ---------------------------------------------------------------------------
# Grid cell text
# ---------------------------------------------------------------------------
COMPANION_FONT_SIZE = 7     # Clue-number superscript in companion cell
COMPANION_BOLD = False
COMPANION_V_ALIGN = "MIDDLE"
COMPANION_H_ALIGN = "CENTER"

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
# Header (merged row above the grid)
# ---------------------------------------------------------------------------
HEADER_ROW = 0              # 0-indexed
HEADER_FONT_SIZE = 15
HEADER_BOLD = True
HEADER_UNDERLINE = True
HEADER_FG = WHITE
HEADER_BG = BLACK
HEADER_V_ALIGN = "MIDDLE"
HEADER_H_ALIGN = "CENTER"
HEADER_ROW_HEIGHT_PX = 40

# Spacer row between header and grid
SPACER_ROW = 1
SPACER_ROW_HEIGHT_PX = 8

# ---------------------------------------------------------------------------
# Tab color
# ---------------------------------------------------------------------------
TAB_COLOR = BLACK

# ---------------------------------------------------------------------------
# Clue panel (to the right of the grid)
# ---------------------------------------------------------------------------
CLUE_GAP_COLS = 1           # Empty columns between grid and clue panel
CLUE_NUM_COL_PX = 35        # Width of the clue-number column
CLUE_TEXT_COL_PX = 320      # Width of the clue-text column

CLUE_HEADER_FONT_SIZE = 12
CLUE_HEADER_BOLD = True
CLUE_HEADER_UNDERLINE = True
CLUE_HEADER_FG = WHITE
CLUE_HEADER_BG = BLACK

CLUE_FONT_SIZE = 10
CLUE_FONT_BOLD = False
CLUE_FG = BLACK
CLUE_BG = WHITE

CLUE_ROW_HEIGHT_PX = 22     # Compact row height for clue list
CLUE_BORDER_WEIGHT = 1
CLUE_BORDER_STYLE = "SOLID"
CLUE_BORDER_COLOR = BLACK

# Frozen header row count inside the clue panel (for scroll illusion)
CLUE_PANEL_START_ROW_OFFSET = 0  # Relative to GRID_START_ROW
