"""Download crossword puzzles using xword-dl and parse .puz files."""

import logging
import os
import subprocess
import tempfile
from datetime import date

import puz

logger = logging.getLogger(__name__)

# Human-readable names for supported free outlets
OUTLET_NAMES = {
    "usa": "USA Today",
    "uni": "Universal",
    "lat": "LA Times",
    "latm": "LA Times Mini",
    "nd": "Newsday",
    "pop": "Daily Pop",
    "atl": "The Atlantic",
    "wp": "Washington Post",
    "db": "The Daily Beast",
    "vox": "Vox",
    "sdp": "Simply Daily Puzzles",
    "sdpc": "Simply Daily Cryptic",
    "sdpq": "Simply Daily Quick",
    "vult": "Vulture",
    "wal": "The Walrus",
}


def download_puzzle(outlet, puzzle_date=None):
    """Download a crossword puzzle and return structured data.

    Args:
        outlet: Short code for the puzzle outlet (e.g. 'usa', 'uni').
        puzzle_date: Optional date string to fetch a specific puzzle.

    Returns:
        dict with puzzle data (grid, clues, metadata) or raises RuntimeError.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = os.path.join(tmpdir, "puzzle.puz")

        if puzzle_date:
            cmd = ["xword-dl", outlet, "-d", puzzle_date, "-o", output_path]
        else:
            cmd = ["xword-dl", outlet, "--latest", "-o", output_path]

        logger.info(f"Running: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        if result.returncode != 0:
            raise RuntimeError(
                f"xword-dl failed for '{outlet}': {result.stderr.strip()}"
            )

        # xword-dl may use a different filename than specified
        if not os.path.exists(output_path):
            puz_files = [f for f in os.listdir(tmpdir) if f.endswith(".puz")]
            if puz_files:
                output_path = os.path.join(tmpdir, puz_files[0])
            else:
                raise RuntimeError(f"No .puz file generated for '{outlet}'")

        return parse_puzzle(output_path, outlet)


def parse_puzzle(filepath, outlet):
    """Parse a .puz file into a structured dict.

    Args:
        filepath: Path to the .puz file.
        outlet: Outlet short code (for naming).

    Returns:
        dict with keys: title, author, width, height, grid, across_clues,
        down_clues, outlet, date.
    """
    p = puz.read(filepath)
    numbering = p.clue_numbering()

    width = p.width
    height = p.height

    # Build the 2D grid
    grid = []
    for row in range(height):
        grid_row = []
        for col in range(width):
            idx = row * width + col
            char = p.solution[idx]
            is_black = char == "."
            grid_row.append(
                {
                    "solution": char if not is_black else None,
                    "is_black": is_black,
                    "number": None,
                }
            )
        grid.append(grid_row)

    # Populate clue numbers on the grid and collect clues
    across_clues = []
    for clue_info in numbering.across:
        r = clue_info["cell"] // width
        c = clue_info["cell"] % width
        grid[r][c]["number"] = clue_info["num"]
        across_clues.append(
            {
                "num": clue_info["num"],
                "clue": clue_info["clue"],
                "answer": clue_info["answer"],
                "length": len(clue_info["answer"]),
            }
        )

    down_clues = []
    for clue_info in numbering.down:
        r = clue_info["cell"] // width
        c = clue_info["cell"] % width
        # Only set number if not already set by an across clue
        if grid[r][c]["number"] is None:
            grid[r][c]["number"] = clue_info["num"]
        down_clues.append(
            {
                "num": clue_info["num"],
                "clue": clue_info["clue"],
                "answer": clue_info["answer"],
                "length": len(clue_info["answer"]),
            }
        )

    outlet_name = OUTLET_NAMES.get(outlet, outlet.upper())

    return {
        "title": p.title or f"{outlet_name} Crossword",
        "author": p.author or "Unknown",
        "width": width,
        "height": height,
        "grid": grid,
        "across_clues": across_clues,
        "down_clues": down_clues,
        "outlet": outlet_name,
        "date": date.today().strftime("%Y-%m-%d"),
    }
