# Crossword-Sheets

A free, automated daily pipeline that downloads crossword puzzles and renders
them as formatted grids in Google Sheets — powered by GitHub Actions and
[xword-dl](https://github.com/thisisparker/xword-dl).

Each day a new tab appears in your Google Sheet with:

- ✅ Numbered grid with black squares and borders
- ✅ Across / Down clue list alongside the grid
- ✅ Hover notes on numbered cells showing the clue text
- ✅ Clean formatting (hidden gridlines, square cells, styled headers)

---

## How It Works

```
GitHub Actions (cron) ──► xword-dl (download .puz) ──► Python (parse) ──► Google Sheets API (render)
```

1. A scheduled GitHub Action runs daily at **12:00 UTC**.
2. `xword-dl` downloads the latest puzzle from each configured outlet.
3. The `.puz` file is parsed with `puzpy` to extract the grid and clues.
4. The Google Sheets API creates a new worksheet tab with the formatted crossword.

---

## Quick Start

### 1. Google Cloud Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/) and **create a project** (or use an existing one).
2. Enable the **Google Sheets API**:
   - Navigate to **APIs & Services → Library**
   - Search for "Google Sheets API" and click **Enable**
3. Enable the **Google Drive API** (same steps — search for "Google Drive API").
4. Create a **Service Account**:
   - Go to **APIs & Services → Credentials**
   - Click **Create Credentials → Service Account**
   - Give it a name (e.g. `crossword-bot`) and click through
5. Create a **JSON key** for the service account:
   - Click on the service account → **Keys** tab → **Add Key → Create new key → JSON**
   - Download the JSON file — you'll need its contents in step 3

### 2. Google Sheet Setup

1. Create a new Google Sheet at [sheets.google.com](https://sheets.google.com).
2. **Share the sheet** with the service account's email address (found in the JSON
   key file as `client_email`, looks like `name@project.iam.gserviceaccount.com`).
   Give it **Editor** access.
3. Copy the **Spreadsheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit
   ```

### 3. GitHub Repository Setup

In your repo go to **Settings → Secrets and variables → Actions** and add:

| Type       | Name                 | Value                                        |
| ---------- | -------------------- | -------------------------------------------- |
| **Secret** | `GOOGLE_CREDENTIALS` | Entire contents of the service-account JSON  |
| **Secret** | `SPREADSHEET_ID`     | The spreadsheet ID from step 2               |

Optionally, add a **Variable** (not secret):

| Type         | Name              | Value                              |
| ------------ | ----------------- | ---------------------------------- |
| **Variable** | `PUZZLE_OUTLETS`  | Comma-separated outlets, e.g. `usa,uni` |

### 4. Run It

- **Manually**: Go to **Actions → Daily Crossword → Run workflow**.
- **Automatically**: The cron runs every day at 12:00 UTC.

---

## Supported Free Outlets

These outlets work without authentication:

| Outlet             | Code   |
| ------------------ | ------ |
| USA Today          | `usa`  |
| Universal          | `uni`  |
| LA Times           | `lat`  |
| LA Times Mini      | `latm` |
| Newsday            | `nd`   |
| Daily Pop          | `pop`  |
| The Atlantic       | `atl`  |
| Washington Post    | `wp`   |
| The Daily Beast    | `db`   |
| Vox                | `vox`  |
| Simply Daily       | `sdp`  |
| Vulture            | `vult` |
| The Walrus         | `wal`  |

> **Note**: Some outlets (like NYT) require authentication. See the
> [xword-dl README](https://github.com/thisisparker/xword-dl) for details.

---

## Local Development

```bash
# Clone and install
git clone https://github.com/braydentodd/Crossword-Sheets.git
cd Crossword-Sheets
pip install -r requirements.txt

# Set environment variables
export GOOGLE_CREDENTIALS='{ ... }'   # or place credentials.json in the repo root
export SPREADSHEET_ID='your-sheet-id'
export PUZZLE_OUTLETS='usa'

# Run
python src/main.py
```

---

## Configuration

| Env Variable         | Required | Default | Description                            |
| -------------------- | -------- | ------- | -------------------------------------- |
| `GOOGLE_CREDENTIALS` | Yes      | —       | Service-account JSON key               |
| `SPREADSHEET_ID`     | Yes      | —       | Target Google Sheet ID                 |
| `PUZZLE_OUTLETS`     | No       | `usa`   | Comma-separated outlet codes           |

---

## Project Structure

```
├── .github/workflows/daily-crossword.yml   # Cron workflow
├── src/
│   ├── main.py                 # Entry point
│   ├── puzzle_downloader.py    # xword-dl wrapper + .puz parser
│   └── sheet_formatter.py      # Google Sheets grid renderer
├── requirements.txt
└── README.md
```

---

## Cost

**$0.** GitHub Actions Free tier includes 2,000 minutes/month. Each run takes
~1–2 minutes. Google Sheets API is free for normal usage.

---

## Limitations

- Google Sheets supports up to **200 worksheet tabs**. At one puzzle/day that's
  ~6 months. Periodically delete old tabs or archive the sheet.
- Clue numbers in grid cells will be overwritten if you type a letter into
  them — the clue list to the right is the primary reference for solving.
- Not all `xword-dl` outlets publish every day; `--latest` grabs the most
  recent available puzzle.

---

## License

MIT
