/**
 * Crossword-Sheets — Apps Script (bound to the spreadsheet)
 *
 * Reads puzzle config from a hidden "_crossword_config" sheet that the
 * Python automation populates on each run. No manual config editing needed.
 *
 * Features:
 *  1. Click a main grid cell → highlight the across & down word + clues.
 *  2. Click a companion (number) cell → redirect focus to the main cell.
 *  3. Click anything else → clear previous highlights.
 *
 * ONE-TIME SETUP:
 *  1. Open the spreadsheet → Extensions → Apps Script.
 *  2. Paste this entire file into Code.gs (replace anything there).
 *  3. Click the clock icon (Triggers) → Add Trigger:
 *       Function: onSelectionChange
 *       Event type: On selection change
 *       (or run installTrigger() once from the editor)
 *  4. Save. Done — it auto-refreshes config from the hidden sheet.
 */

// ===== Trigger installer (run once from the script editor) =================
function installTrigger() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("onSelectionChange")
    .forSpreadsheet(ss)
    .onSelectionChange()
    .create();
}

// ===== Config cache (per-execution, avoids repeated sheet reads) ===========
var _cfgCache = {};

function _loadConfig(sheetName) {
  if (_cfgCache[sheetName] !== undefined) return _cfgCache[sheetName];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgSheet;
  try {
    cfgSheet = ss.getSheetByName("_crossword_config");
  } catch (_) {
    _cfgCache[sheetName] = null;
    return null;
  }
  if (!cfgSheet) { _cfgCache[sheetName] = null; return null; }

  var data = cfgSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === sheetName) {
      try {
        var cfg = JSON.parse(data[i][1]);
        _cfgCache[sheetName] = cfg;
        return cfg;
      } catch (e) {
        Logger.log("Config parse error for " + sheetName + ": " + e);
        _cfgCache[sheetName] = null;
        return null;
      }
    }
  }
  _cfgCache[sheetName] = null;
  return null;
}

// ===== Main event handler ==================================================
function onSelectionChange(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  var cfg = _loadConfig(sheet.getName());
  if (!cfg) return;

  var row = e.range.getRow() - 1;    // 0-indexed
  var col = e.range.getColumn() - 1;  // 0-indexed

  // Clear previous highlights
  _clearHighlights(sheet, cfg);

  // Check if inside grid
  var gridRowEnd = cfg.gridStartRow + cfg.gridHeight;
  var gridColEnd = cfg.gridStartCol + cfg.gridWidth * 2;

  if (row < cfg.gridStartRow || row >= gridRowEnd ||
      col < cfg.gridStartCol || col >= gridColEnd) {
    return;
  }

  var relCol = col - cfg.gridStartCol;
  var gridCol = Math.floor(relCol / 2);
  var isCompanion = (relCol % 2 === 0);

  // Companion click → redirect to main cell
  if (isCompanion) {
    var mainCol = cfg.gridStartCol + gridCol * 2 + 1;
    sheet.getRange(row + 1, mainCol + 1).activate();
    return;
  }

  // Main cell — look up word membership
  var gridRow = row - cfg.gridStartRow;
  var key = gridRow + "," + gridCol;
  var wordInfo = cfg.wordMap[key];
  if (!wordInfo) return;

  if (wordInfo.across) {
    _highlightWord(sheet, cfg, wordInfo.across.cells);
    _highlightClue(sheet, cfg, "A" + wordInfo.across.clueNum);
  }
  if (wordInfo.down) {
    _highlightWord(sheet, cfg, wordInfo.down.cells);
    _highlightClue(sheet, cfg, "D" + wordInfo.down.clueNum);
  }
}

// ===== Highlight helpers ===================================================
function _highlightWord(sheet, cfg, cells) {
  for (var i = 0; i < cells.length; i++) {
    var r = cells[i][0];
    var c = cells[i][1];
    var sheetRow = cfg.gridStartRow + r + 1;
    var compCol = cfg.gridStartCol + c * 2 + 1;
    var mainCol = compCol + 1;
    sheet.getRange(sheetRow, compCol).setBackground(cfg.highlightWord);
    sheet.getRange(sheetRow, mainCol).setBackground(cfg.highlightWord);
  }
}

function _highlightClue(sheet, cfg, clueKey) {
  var clueRow = cfg.clueRowIndex[clueKey];
  if (clueRow === undefined) return;
  var sheetRow = clueRow + 1;
  sheet.getRange(sheetRow, cfg.clueNumCol + 1).setBackground(cfg.highlightClue);
  sheet.getRange(sheetRow, cfg.clueTextCol + 1).setBackground(cfg.highlightClue);
}

function _clearHighlights(sheet, cfg) {
  // Clear grid cells
  for (var key in cfg.wordMap) {
    var parts = key.split(",");
    var r = parseInt(parts[0]);
    var c = parseInt(parts[1]);
    var sheetRow = cfg.gridStartRow + r + 1;
    var compCol = cfg.gridStartCol + c * 2 + 1;
    var mainCol = compCol + 1;
    sheet.getRange(sheetRow, compCol).setBackground(cfg.whiteHex);
    sheet.getRange(sheetRow, mainCol).setBackground(cfg.whiteHex);
  }
  // Clear clue panel
  for (var clueKey in cfg.clueRowIndex) {
    var clueRow = cfg.clueRowIndex[clueKey] + 1;
    sheet.getRange(clueRow, cfg.clueNumCol + 1).setBackground(cfg.clueBgHex);
    sheet.getRange(clueRow, cfg.clueTextCol + 1).setBackground(cfg.clueBgHex);
  }
}
