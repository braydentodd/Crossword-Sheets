"""Download crossword puzzles using xword-dl and parse .puz files."""

import base64
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date, datetime

import puz
import requests

logger = logging.getLogger(__name__)


def _xword_dl_cmd():
    """Locate the xword-dl executable.

    Prefers the one in the same bin directory as the running Python
    interpreter (i.e. inside the active venv), then falls back to PATH.
    """
    venv_bin = os.path.join(os.path.dirname(sys.executable))
    candidate = os.path.join(venv_bin, "xword-dl")
    if os.path.isfile(candidate):
        return candidate
    on_path = shutil.which("xword-dl")
    if on_path:
        return on_path
    raise FileNotFoundError(
        "xword-dl not found. Install it with: pip install xword-dl"
    )

# Human-readable names for supported free outlets
OUTLET_NAMES = {
    "usa": "USA Today",
    "uni": "Universal",
    "lat": "LA Times",
    "latm": "LA Times Mini",
    "nd": "Newsday",
    "pop": "Daily Pop",
    "atl": "The Atlantic",
    "wp": "Washington Post Sunday",
    "wpd": "Washington Post",
    "wpm": "WaPo Mini",
    "db": "The Daily Beast",
    "vox": "Vox",
    "sdp": "Simply Daily Puzzles",
    "sdpc": "Simply Daily Cryptic",
    "sdpq": "Simply Daily Quick",
    "vult": "Vulture",
    "wal": "The Walrus",
}

# ---------------------------------------------------------------------------
# AmuseLabs direct-download support
#
# Some AmuseLabs-based outlets (db, wal, lat) have a broken --latest scraper
# in xword-dl because the picker page no longer includes a "puzzles" array.
# Instead, the puzzle list is in "streakInfo" and puzzle IDs can be random
# hex strings (db, wal) or date-based (lat = tca{YYMMDD}).
#
# Our strategy:
#   1. Fetch the date-picker page.
#   2. Parse <script id="params"> → extract streakInfo (puzzle list) and
#      rawsps → base64 decode → loadToken.
#   3. For --latest: use puzzleId from streakInfo[0].
#      For -d DATE: match by publicationTime.
#   4. Build the full puzzle URL with loadToken and hand it to xword-dl.
# ---------------------------------------------------------------------------
_AMUSELABS_OUTLETS = {
    "db":  ("cdn3.amuselabs.com/tdb", "tdb"),
    "wal": ("cdn2.amuselabs.com/pmm", "walrus-weekly-crossword"),
    "lat": ("lat.amuselabs.com/lat",  "latimes"),
}


def _amuselabs_fetch_picker(host, set_name):
    """Fetch and parse the AmuseLabs date-picker page.

    Returns:
        (streak_info, load_token) where streak_info is a list of puzzle
        entries and load_token is the JWT for puzzle URLs.
    """
    from bs4 import BeautifulSoup, Tag

    picker_url = f"https://{host}/date-picker?set={set_name}"
    r = requests.get(
        picker_url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=15,
    )
    r.raise_for_status()

    soup = BeautifulSoup(r.text, "html.parser")
    params_tag = soup.find("script", id="params")
    if not isinstance(params_tag, Tag) or not params_tag.string:
        raise RuntimeError("No <script id='params'> on AmuseLabs picker page")

    params = json.loads(params_tag.string)

    # Extract puzzle list from streakInfo (new format) or puzzles (legacy)
    streak_info = params.get("streakInfo", [])
    puzzles_legacy = params.get("puzzles", [])

    if not streak_info and not puzzles_legacy:
        raise RuntimeError("No puzzle data in AmuseLabs picker page")

    # Extract loadToken from rawsps
    rawsps = params.get("rawsps", "")
    load_token = ""
    if rawsps:
        sps_data = json.loads(base64.b64decode(rawsps))
        load_token = sps_data.get("loadToken", "")

    return streak_info, puzzles_legacy, load_token


def _amuselabs_find_puzzle(outlet, puzzle_date_str=None):
    """Find puzzle ID and build the download URL.

    Args:
        outlet: Short outlet code (e.g. 'db', 'wal', 'lat').
        puzzle_date_str: Optional ISO date string (YYYY-MM-DD).
                         If None, returns the latest puzzle.

    Returns:
        (puzzle_url, puzzle_title) tuple.
    """
    if outlet not in _AMUSELABS_OUTLETS:
        raise RuntimeError(f"No AmuseLabs config for outlet '{outlet}'")

    host, set_name = _AMUSELABS_OUTLETS[outlet]
    streak_info, puzzles_legacy, load_token = _amuselabs_fetch_picker(host, set_name)

    puzzle_id = None
    puzzle_title = ""

    if streak_info:
        if puzzle_date_str:
            # Match by publicationTime (epoch ms) — compare dates
            target_date = datetime.strptime(puzzle_date_str, "%Y-%m-%d").date()
            for entry in streak_info:
                pd = entry.get("puzzleDetails", {})
                pub_ms = pd.get("publicationTime", 0)
                if pub_ms:
                    pub_date = datetime.utcfromtimestamp(pub_ms / 1000).date()
                    if pub_date == target_date:
                        puzzle_id = pd.get("puzzleId")
                        puzzle_title = pd.get("title", "")
                        break
            if not puzzle_id:
                raise RuntimeError(
                    f"No AmuseLabs puzzle for {outlet} on {puzzle_date_str}"
                )
        else:
            # Latest = first entry
            pd = streak_info[0].get("puzzleDetails", {})
            puzzle_id = pd.get("puzzleId")
            puzzle_title = pd.get("title", "")
    elif puzzles_legacy:
        # Fallback to legacy format (list of {id, ...} dicts)
        if puzzle_date_str:
            puzzle_id = puzzles_legacy[0].get("id")
        else:
            puzzle_id = puzzles_legacy[0].get("id")
        puzzle_title = ""

    if not puzzle_id:
        raise RuntimeError(f"Could not determine puzzle ID for {outlet}")

    puzzle_url = f"https://{host}/crossword?id={puzzle_id}&set={set_name}"
    if load_token:
        puzzle_url += f"&loadToken={load_token}"

    logger.info(
        f"AmuseLabs resolved {outlet} → id={puzzle_id} title='{puzzle_title}'"
    )
    return puzzle_url, puzzle_title


def _amuselabs_download(outlet, tmpdir, puzzle_date_str=None):
    """Download a .puz file from AmuseLabs.

    Args:
        outlet: Short outlet code (e.g. 'db', 'wal', 'lat').
        tmpdir: Directory to save the .puz file in.
        puzzle_date_str: Optional ISO date string (YYYY-MM-DD).

    Returns:
        Path to the saved .puz file, or raises RuntimeError.
    """
    puzzle_url, puzzle_title = _amuselabs_find_puzzle(outlet, puzzle_date_str)

    output_path = os.path.join(tmpdir, "puzzle.puz")
    cmd = [_xword_dl_cmd(), puzzle_url, "-o", output_path]
    logger.info("AmuseLabs download: xword-dl <url> -o puzzle.puz")
    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=120, cwd=tmpdir,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"xword-dl failed for AmuseLabs URL ({outlet}): "
            f"{result.stderr.strip()}"
        )
    # xword-dl may save with its own filename
    if not os.path.isfile(output_path):
        puz_files = [f for f in os.listdir(tmpdir) if f.endswith(".puz")]
        if not puz_files:
            raise RuntimeError(f"No .puz file from AmuseLabs download for '{outlet}'")
        output_path = os.path.join(tmpdir, puz_files[0])

    return output_path


# ---------------------------------------------------------------------------
# Washington Post direct API support
#
# xword-dl only supports the Sunday WaPo crossword ("wp" outlet). The WaPo
# API also has daily and mini endpoints that return the same JSON format.
# We handle "wpd" (daily) and "wpm" (mini) by fetching the JSON directly
# and using xword-dl's WaPo parser to build the .puz file.
# ---------------------------------------------------------------------------
_WAPO_API_BASE = "https://games-service-prod.site.aws.wapo.pub/crossword/levels"
_WAPO_CUSTOM_OUTLETS = {
    "wpd": "daily",
    "wpm": "mini",
}


def _wapo_download(outlet, tmpdir, puzzle_date_str=None):
    """Download a .puz file from the WaPo games API.

    Args:
        outlet: 'wpd' (daily) or 'wpm' (mini).
        tmpdir: Directory to save the .puz file in.
        puzzle_date_str: Optional ISO date (YYYY-MM-DD). Defaults to today.

    Returns:
        Path to the saved .puz file, or raises RuntimeError.
    """
    from xword_dl.downloader.wapodownloader import WaPoDownloader

    level = _WAPO_CUSTOM_OUTLETS[outlet]
    dt = datetime.strptime(puzzle_date_str, "%Y-%m-%d") if puzzle_date_str else datetime.now()
    url_date = dt.strftime("%Y/%m/%d")
    api_url = f"{_WAPO_API_BASE}/{level}/{url_date}"

    logger.info(f"WaPo API: fetching {api_url}")
    r = requests.get(api_url, timeout=15)
    r.raise_for_status()
    if not r.text or len(r.text) < 20:
        raise RuntimeError(f"No WaPo {level} puzzle available for {url_date}")

    xw_data = r.json()

    # Use xword-dl's WaPo parser to build the .puz object
    dl = WaPoDownloader()
    dl.date = dt
    puzzle_obj = dl.parse_xword(xw_data)

    # .puz format uses latin-1 encoding; replace problematic Unicode chars
    for attr in ("title", "author", "copyright", "notes"):
        val = getattr(puzzle_obj, attr, "") or ""
        val = val.replace("\u2019", "'").replace("\u2018", "'")
        val = val.replace("\u201c", '"').replace("\u201d", '"')
        val = val.replace("\u2014", "--").replace("\u2013", "-")
        val = val.encode("latin-1", errors="replace").decode("latin-1")
        setattr(puzzle_obj, attr, val)

    # Also sanitize clues
    sanitized_clues = []
    for clue in puzzle_obj.clues:
        clue = clue.replace("\u2019", "'").replace("\u2018", "'")
        clue = clue.replace("\u201c", '"').replace("\u201d", '"')
        clue = clue.replace("\u2014", "--").replace("\u2013", "-")
        clue = clue.encode("latin-1", errors="replace").decode("latin-1")
        sanitized_clues.append(clue)
    puzzle_obj.clues = sanitized_clues

    output_path = os.path.join(tmpdir, "puzzle.puz")
    puzzle_obj.save(output_path)
    logger.info(f"WaPo {level} puzzle saved: {puzzle_obj.title}")
    return output_path


def _xword_dl_run(cmd, tmpdir):
    """Run an xword-dl command and return the .puz filepath, or raise."""
    logger.info(f"Running: {' '.join(cmd)}")
    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=120, cwd=tmpdir,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"xword-dl failed: {result.stderr.strip()}"
        )
    puz_files = [f for f in os.listdir(tmpdir) if f.endswith(".puz")]
    if not puz_files:
        raise RuntimeError("No .puz file generated")
    return os.path.join(tmpdir, puz_files[0])


def download_puzzle(outlet, puzzle_date=None):
    """Download a crossword puzzle and return structured data.

    Uses a multi-strategy fallback chain:
      1. If puzzle_date given: xword-dl -d DATE
      2. Else: xword-dl --latest
      3. If (1) or (2) fail and outlet has AmuseLabs config: direct URL download
      4. If --latest failed: retry with xword-dl -d TODAY

    Args:
        outlet: Short code for the puzzle outlet (e.g. 'usa', 'uni').
        puzzle_date: Optional ISO date string (YYYY-MM-DD) for a specific puzzle.

    Returns:
        dict with puzzle data (grid, clues, metadata) or raises RuntimeError.
    """
    today_str = date.today().strftime("%Y-%m-%d")
    target_date = puzzle_date or today_str
    errors = []

    with tempfile.TemporaryDirectory() as tmpdir:
        filepath = None
        used_date = puzzle_date  # track which date we actually got

        # --- Strategy 1: Explicit date via -d flag ---
        if puzzle_date:
            try:
                output_path = os.path.join(tmpdir, "puzzle.puz")
                cmd = [_xword_dl_cmd(), outlet, "-d", puzzle_date, "-o", output_path]
                filepath = _xword_dl_run(cmd, tmpdir)
                used_date = puzzle_date
            except Exception as exc:
                errors.append(f"-d {puzzle_date}: {exc}")
                logger.debug(f"Strategy '-d DATE' failed for {outlet}: {exc}")

        # --- Strategy 2: --latest (only when no explicit date) ---
        if not filepath and not puzzle_date:
            try:
                cmd = [_xword_dl_cmd(), outlet, "--latest"]
                filepath = _xword_dl_run(cmd, tmpdir)
                used_date = None  # will extract from filename
            except Exception as exc:
                errors.append(f"--latest: {exc}")
                logger.debug(f"Strategy '--latest' failed for {outlet}: {exc}")

        # --- Strategy 3a: WaPo direct API (wpd/wpm) ---
        if not filepath and outlet in _WAPO_CUSTOM_OUTLETS:
            try:
                filepath = _wapo_download(outlet, tmpdir, puzzle_date_str=target_date)
                used_date = target_date
                logger.info(f"WaPo direct API succeeded for {outlet}")
            except Exception as exc:
                errors.append(f"WaPo API: {exc}")
                logger.debug(f"Strategy 'WaPo API' failed for {outlet}: {exc}")

        # --- Strategy 3b: AmuseLabs picker-based download ---
        if not filepath and outlet in _AMUSELABS_OUTLETS:
            try:
                filepath = _amuselabs_download(
                    outlet, tmpdir, puzzle_date_str=puzzle_date,
                )
                used_date = puzzle_date or target_date
                logger.info(f"AmuseLabs download succeeded for {outlet}")
            except Exception as exc:
                errors.append(f"AmuseLabs: {exc}")
                logger.debug(f"Strategy 'AmuseLabs' failed for {outlet}: {exc}")

        # --- Strategy 4: Fallback -d TODAY (if --latest failed) ---
        if not filepath and not puzzle_date:
            try:
                output_path = os.path.join(tmpdir, "puzzle.puz")
                cmd = [_xword_dl_cmd(), outlet, "-d", today_str, "-o", output_path]
                filepath = _xword_dl_run(cmd, tmpdir)
                used_date = today_str
            except Exception as exc:
                errors.append(f"-d {today_str} (fallback): {exc}")
                logger.debug(f"Strategy '-d TODAY fallback' failed for {outlet}: {exc}")

        # --- All strategies exhausted ---
        if not filepath:
            raise RuntimeError(
                f"All download strategies failed for '{outlet}':\n"
                + "\n".join(f"  • {e}" for e in errors)
            )

        # Extract date from filename if we don't have one
        file_date = used_date
        if not file_date:
            filename = os.path.basename(filepath)
            m = re.search(r'(\d{8})', filename)
            if m:
                try:
                    file_date = datetime.strptime(m.group(1), "%Y%m%d").strftime("%Y-%m-%d")
                except ValueError:
                    pass

        return parse_puzzle(filepath, outlet, file_date=file_date)


def parse_puzzle(filepath, outlet, file_date=None):
    """Parse a .puz file into a structured dict.

    Args:
        filepath: Path to the .puz file.
        outlet: Outlet short code (for naming).
        file_date: Optional YYYY-MM-DD date string extracted from filename.

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
                "length": clue_info["len"],
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
                "length": clue_info["len"],
            }
        )

    outlet_name = OUTLET_NAMES.get(outlet, outlet.upper())

    # --- Puzzle date ---
    # Priority: file_date (from filename/explicit) → copyright → today
    puzzle_date = file_date
    if not puzzle_date:
        m = re.search(r'(\w+ \d+, \d{4})', p.copyright or '')
        if m:
            try:
                puzzle_date = datetime.strptime(m.group(1), "%B %d, %Y").strftime("%Y-%m-%d")
            except ValueError:
                pass
    if not puzzle_date:
        puzzle_date = date.today().strftime("%Y-%m-%d")

    # --- Title ---
    title = (p.title or "").strip()

    # --- Author (strip leading "By ") ---
    author = (p.author or "Unknown").strip()
    if author.lower().startswith("by "):
        author = author[3:].strip()

    return {
        "title": title,
        "author": author,
        "width": width,
        "height": height,
        "grid": grid,
        "across_clues": across_clues,
        "down_clues": down_clues,
        "outlet": outlet_name,
        "date": puzzle_date,
    }
