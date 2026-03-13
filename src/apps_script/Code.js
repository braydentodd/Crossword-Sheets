/**
 * Crossword helper -- container-bound Apps Script.
 *
 * Features:
 *   onOpen       - adds Crossword menu with Check / Reveal
 *   onEdit       - strikethrough clue when its word is fully filled
 *   checkPuzzle  - colors each main cell green (correct) or red (wrong/empty)
 *   revealPuzzle - fills in correct answers + colors green/red
 *
 * Config is stored in the hidden _xw_config sheet (one row per puzzle tab).
 * Run clearCache() after re-running main.py to flush stale config.
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

var COLOR_CORRECT = "#d9ead3";  // light green
var COLOR_WRONG   = "#f4cccc";  // light red
var COLOR_WHITE   = "#ffffff";
var COLOR_HIGHLIGHT = "#fff2cc"; // light yellow
var FONT_CROSSED  = "#cc0000";  // red text for completed clues
var FONT_DEFAULT  = "#000000";  // black text (normal)

// ---------------------------------------------------------------------------
// Config - CacheService so we only hit _xw_config once per 6 hours
// ---------------------------------------------------------------------------

function _getCfg(sheet) {
  var name  = _getBaseName(sheet.getName());
  var cache = CacheService.getDocumentCache();
  var key   = "xw_" + name;

  var cached = cache.get(key);
  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      // Validate that cached config has required fields added in later deploys.
      // If missing, the cache is stale — discard it and re-read from the sheet.
      if (parsed.acrossNumCol !== undefined && parsed.clueStartRow !== undefined) {
        return parsed;
      }
      cache.remove(key);
    } catch(ignore) {}
  }

  var cfgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_xw_config");
  if (!cfgSheet) return null;

  var data = cfgSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === name && data[i][1]) {
      var raw = String(data[i][1]);
      try { cache.put(key, raw, 21600); } catch(ignore) {}
      return JSON.parse(raw);
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Cell-type helpers
// ---------------------------------------------------------------------------

function _inGrid(r, c, cfg) {
  return r >= cfg.gridStartRow && r < cfg.gridStartRow + cfg.gridHeight * 2 &&
         c >= cfg.gridStartCol && c < cfg.gridStartCol + cfg.gridWidth  * 2;
}

function _isCompanionCol(c, cfg) {
  return (c - cfg.gridStartCol) % 2 === 0;
}

function _isBlackLogical(logR, logC, cfg) {
  return !!(cfg.blackCells && cfg.blackCells[logR + "," + logC]);
}

function _getBaseName(name) {
  if (name.length >= 2) {
    var tail = name.slice(-2);
    if (tail === " ✅" || tail === " ❌") return name.slice(0, -2);
  }
  return name;
}

function _restoreHighlights(sheet) {
  var cache = CacheService.getDocumentCache();
  var key = "xw_hl_" + _getBaseName(sheet.getName());
  var raw = cache.get(key);
  cache.remove(key);           // always clear, even if null
  var cfg = _getCfg(sheet);
  if (!cfg) return;

  var GSR = cfg.gridStartRow, GSC = cfg.gridStartCol;
  var GW  = cfg.gridWidth,    GH  = cfg.gridHeight;

  try {
    var gridDirty = false, clueDirty = false;

    var gridRange = sheet.getRange(GSR + 1, GSC + 1, GH * 2, GW * 2);
    var gridBgs   = gridRange.getBackgrounds();

    var clueRange = null, clueBgs = null;
    if (cfg.acrossNumCol !== undefined && cfg.clueStartRow !== undefined) {
      var maxCR = cfg.maxClueRows || 0;
      var csStart = cfg.acrossNumCol;
      var csEnd   = cfg.downTextCol;
      var csWidth = csEnd - csStart + 1;
      if (maxCR > 0) {
        clueRange = sheet.getRange(cfg.clueStartRow + 1, csStart + 1, maxCR, csWidth);
        clueBgs   = clueRange.getBackgrounds();
      }
    }

    // Restore known originals from cache
    if (raw) {
      var cells = JSON.parse(raw);
      for (var i = 0; i < cells.length; i++) {
        var p = cells[i];
        if (p.t === "g") {
          gridBgs[p.r][p.c] = p.bg;
          gridDirty = true;
        } else if (p.t === "c" && clueBgs) {
          clueBgs[p.r][p.c] = p.bg;
          clueDirty = true;
        }
      }
    }

    // Safety net: reset any remaining orphaned highlight cells to white.
    // This catches highlights stranded by cache eviction or race conditions.
    for (var gr = 0; gr < gridBgs.length; gr++) {
      for (var gc = 0; gc < gridBgs[gr].length; gc++) {
        if (gridBgs[gr][gc] === COLOR_HIGHLIGHT) {
          gridBgs[gr][gc] = COLOR_WHITE;
          gridDirty = true;
        }
      }
    }
    if (clueBgs) {
      for (var cr = 0; cr < clueBgs.length; cr++) {
        for (var cc = 0; cc < clueBgs[cr].length; cc++) {
          if (clueBgs[cr][cc] === COLOR_HIGHLIGHT) {
            clueBgs[cr][cc] = COLOR_WHITE;
            clueDirty = true;
          }
        }
      }
    }

    if (gridDirty) gridRange.setBackgrounds(gridBgs);
    if (clueDirty && clueRange) clueRange.setBackgrounds(clueBgs);
  } catch(ignore) {}
}

function _findClueWordKey(row, col, cfg) {
  if (!cfg.clueRows) return null;
  for (var wk in cfg.clueRows) {
    var pos = cfg.clueRows[wk];
    if (!pos) continue;
    if (pos[0] === row && (col === pos[1] || col === pos[1] - 1)) {
      return wk;
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Check Puzzle")
    .addItem("Check Puzzle", "checkPuzzle")
    .addToUi();
  ui.createMenu("Reveal Puzzle")
    .addItem("Reveal Puzzle", "revealPuzzle")
    .addToUi();
  ui.createMenu("Clear Colors")
    .addItem("Clear Colors", "clearColors")
    .addToUi();
  ui.createMenu("\u2699 Options")
    .addItem("Clear Cache", "clearCache")
    .addItem("Setup Highlight Trigger", "setupTriggers")
    .addItem("Test Highlight (current cell)", "testHighlight")
    .addToUi();
}

// ---------------------------------------------------------------------------
// Setup installable trigger for onSelectionChange (run once)
// ---------------------------------------------------------------------------

function setupTriggers() {
  // Remove existing onSelectionChange triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onSelectionChange") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Create installable trigger
  ScriptApp.newTrigger("onSelectionChange")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
  SpreadsheetApp.getActive().toast(
    "Selection-change trigger installed! Highlighting should now work on cell clicks.",
    "Done", 5
  );
}

// ---------------------------------------------------------------------------
// Test highlight - manually trigger highlighting for the active cell
// ---------------------------------------------------------------------------

function testHighlight() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell  = sheet.getActiveCell();
  var fakeEvent = { range: cell };
  try {
    _doSelectionChange(fakeEvent);
    SpreadsheetApp.getActive().toast("Highlight applied!", "OK", 3);
  } catch(err) {
    SpreadsheetApp.getUi().alert("Error in highlighting:\\n" + err.message + "\\n\\nStack:\\n" + err.stack);
  }
}

// ---------------------------------------------------------------------------
// Check Puzzle - color each cell green (correct) or red (wrong / empty)
// ---------------------------------------------------------------------------

function checkPuzzle() {
  var sheet = SpreadsheetApp.getActiveSheet();
  _restoreHighlights(sheet);
  var cfg   = _getCfg(sheet);
  if (!cfg || !cfg.solutions) {
    SpreadsheetApp.getUi().alert("No puzzle config found for this sheet.");
    return;
  }

  var GSR = cfg.gridStartRow, GSC = cfg.gridStartCol;
  var GW  = cfg.gridWidth,    GH  = cfg.gridHeight;

  var gridRange = sheet.getRange(GSR + 1, GSC + 1, GH * 2, GW * 2);
  var gridVals  = gridRange.getValues();
  var gridBgs   = gridRange.getBackgrounds();

  var allCorrect = true;

  for (var key in cfg.solutions) {
    var parts = key.split(",");
    var lr = parseInt(parts[0], 10), lc = parseInt(parts[1], 10);
    var subRow = lr * 2;
    var subCol = lc * 2 + 1;  // main column

    var entered = String(gridVals[subRow][subCol]).trim().toUpperCase();
    var correct = cfg.solutions[key].toUpperCase();

    var color = (entered === correct) ? COLOR_CORRECT : COLOR_WRONG;
    if (color === COLOR_WRONG) allCorrect = false;

    gridBgs[subRow][subCol - 1]     = color;  // companion cell
    gridBgs[subRow + 1][subCol - 1] = color;
    gridBgs[subRow][subCol]         = color;  // main cell
    gridBgs[subRow + 1][subCol]     = color;
  }

  gridRange.setBackgrounds(gridBgs);

  var baseName = _getBaseName(sheet.getName());
  sheet.setName(baseName + (allCorrect ? " ✅" : " ❌"));

  SpreadsheetApp.flush();
}

// ---------------------------------------------------------------------------
// Reveal Puzzle - fill in correct answers + color green/red
// ---------------------------------------------------------------------------

function revealPuzzle() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    "Reveal Puzzle",
    "This will fill in all answers. Are you sure?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var sheet = SpreadsheetApp.getActiveSheet();
  _restoreHighlights(sheet);
  var cfg   = _getCfg(sheet);
  if (!cfg || !cfg.solutions) {
    ui.alert("No puzzle config found for this sheet.");
    return;
  }

  var GSR = cfg.gridStartRow, GSC = cfg.gridStartCol;
  var GW  = cfg.gridWidth,    GH  = cfg.gridHeight;

  var gridRange = sheet.getRange(GSR + 1, GSC + 1, GH * 2, GW * 2);
  var gridVals  = gridRange.getValues();
  var gridBgs   = gridRange.getBackgrounds();

  var allCorrect = true;

  for (var key in cfg.solutions) {
    var parts = key.split(",");
    var lr = parseInt(parts[0], 10), lc = parseInt(parts[1], 10);
    var subRow = lr * 2;
    var subCol = lc * 2 + 1;

    var entered = String(gridVals[subRow][subCol]).trim().toUpperCase();
    var correct = cfg.solutions[key].toUpperCase();

    var color = (entered === correct) ? COLOR_CORRECT : COLOR_WRONG;
    if (color === COLOR_WRONG) allCorrect = false;

    gridBgs[subRow][subCol - 1]     = color;  // companion cell
    gridBgs[subRow + 1][subCol - 1] = color;
    gridBgs[subRow][subCol]         = color;  // main cell
    gridBgs[subRow + 1][subCol]     = color;

    // Fill in the correct answer
    gridVals[subRow][subCol] = correct;
  }

  gridRange.setValues(gridVals);
  gridRange.setBackgrounds(gridBgs);

  // Strike through ALL clues since the puzzle is now complete
  if (cfg.clueRows) {
    for (var wk in cfg.clueRows) {
      var pos = cfg.clueRows[wk];
      if (pos) {
        var clueCell = sheet.getRange(pos[0] + 1, pos[1] + 1);
        clueCell.setFontLine("line-through");
        clueCell.setFontColor(FONT_CROSSED);
      }
    }
  }

  var baseName = _getBaseName(sheet.getName());
  sheet.setName(baseName + (allCorrect ? " ✅" : " ❌"));

  SpreadsheetApp.flush();
}

// ---------------------------------------------------------------------------
// Clear Colors - remove all greens, reds, and yellows; flush highlight cache
//   so no process restores them until the user actively selects a cell
//   or clicks Check / Reveal again.
// ---------------------------------------------------------------------------

function clearColors() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cfg   = _getCfg(sheet);
  if (!cfg) return;

  var GSR = cfg.gridStartRow, GSC = cfg.gridStartCol;
  var GW  = cfg.gridWidth,    GH  = cfg.gridHeight;

  // 1. Flush the highlight cache so onSelectionChange won't restore old
  //    green/red originals on the next click.
  var cache = CacheService.getDocumentCache();
  var hlKey = "xw_hl_" + _getBaseName(sheet.getName());
  cache.remove(hlKey);

  // 2. Clear grid colours (greens, reds, yellows) -> white
  var gridRange = sheet.getRange(GSR + 1, GSC + 1, GH * 2, GW * 2);
  var gridBgs   = gridRange.getBackgrounds();
  var gridDirty = false;

  for (var gr = 0; gr < gridBgs.length; gr++) {
    for (var gc = 0; gc < gridBgs[gr].length; gc++) {
      var bg = gridBgs[gr][gc];
      if (bg === COLOR_CORRECT || bg === COLOR_WRONG || bg === COLOR_HIGHLIGHT) {
        gridBgs[gr][gc] = COLOR_WHITE;
        gridDirty = true;
      }
    }
  }
  if (gridDirty) gridRange.setBackgrounds(gridBgs);

  // 3. Clear clue yellows
  if (cfg.acrossNumCol !== undefined && cfg.clueStartRow !== undefined) {
    var maxCR   = cfg.maxClueRows || 0;
    var csStart = cfg.acrossNumCol;
    var csEnd   = cfg.downTextCol;
    var csWidth = csEnd - csStart + 1;
    if (maxCR > 0) {
      var clueRange = sheet.getRange(cfg.clueStartRow + 1, csStart + 1, maxCR, csWidth);
      var clueBgs   = clueRange.getBackgrounds();
      var clueDirty = false;
      for (var cr = 0; cr < clueBgs.length; cr++) {
        for (var cc = 0; cc < clueBgs[cr].length; cc++) {
          if (clueBgs[cr][cc] === COLOR_HIGHLIGHT) {
            clueBgs[cr][cc] = COLOR_WHITE;
            clueDirty = true;
          }
        }
      }
      if (clueDirty) clueRange.setBackgrounds(clueBgs);
    }
  }

  // 4. Remove check/reveal tab indicator (✅ / ❌)
  var baseName = _getBaseName(sheet.getName());
  if (baseName !== sheet.getName()) sheet.setName(baseName);

  SpreadsheetApp.flush();
}

// ---------------------------------------------------------------------------
// Cache / Debug helpers
// ---------------------------------------------------------------------------

function clearCache() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var cache = CacheService.getDocumentCache();
  ss.getSheets().forEach(function(s) {
    var name     = s.getName();
    var baseName = _getBaseName(name);
    cache.remove("xw_" + name);
    cache.remove("xw_" + baseName);
    cache.remove("xw_hl_" + name);
    cache.remove("xw_hl_" + baseName);
  });
  ss.toast("Cache cleared. Config will reload on next action.", "Done", 6);
}

function debugInfo() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var cfg   = _getCfg(sheet);
  if (!cfg) {
    Logger.log("No config found for sheet: " + sheet.getName());
  } else {
    Logger.log("Config OK: " + sheet.getName());
    Logger.log("Grid: " + cfg.gridWidth + "x" + cfg.gridHeight);
    Logger.log("cellMap:" + !!cfg.cellMap + " wordCells:" + !!cfg.wordCells +
               " clueRows:" + !!cfg.clueRows + " blackCells:" + !!cfg.blackCells +
               " solutions:" + !!cfg.solutions);
    if (cfg.solutions) Logger.log("Solution count: " + Object.keys(cfg.solutions).length);
  }
}

// ---------------------------------------------------------------------------
// onEdit - strikethrough clue when its word is fully filled, remove when not
// ---------------------------------------------------------------------------

function onEdit(e) {
  var sheet = e.range.getSheet();
  var cfg   = _getCfg(sheet);
  if (!cfg) return;

  var r = e.range.getRow()    - 1;
  var c = e.range.getColumn() - 1;

  if (!_inGrid(r, c, cfg)) return;
  if (_isCompanionCol(c, cfg)) return;

  var GSR  = cfg.gridStartRow;
  var GSC  = cfg.gridStartCol;
  var GW   = cfg.gridWidth;
  var GH   = cfg.gridHeight;
  var logR = Math.floor((r - GSR) / 2);
  var logC = Math.floor((c - GSC) / 2);

  if (_isBlackLogical(logR, logC, cfg)) return;

  var gridVals = sheet.getRange(GSR + 1, GSC + 1, GH * 2, GW * 2).getValues();

  if (!cfg.cellMap || !cfg.wordCells || !cfg.clueRows) return;

  var refs = cfg.cellMap[logR + "," + logC];
  if (!refs) return;

  var wordKeys = [];
  if (refs.a !== undefined) wordKeys.push("A" + refs.a);
  if (refs.d !== undefined) wordKeys.push("D" + refs.d);

  var strikeTasks = [];
  for (var wi = 0; wi < wordKeys.length; wi++) {
    var wk    = wordKeys[wi];
    var cells = cfg.wordCells[wk];
    if (!cells) continue;

    var complete = true;
    for (var ci = 0; ci < cells.length; ci++) {
      var lr = cells[ci][0], lc = cells[ci][1];
      if (!String(gridVals[lr * 2][lc * 2 + 1]).trim()) {
        complete = false;
        break;
      }
    }

    var pos = cfg.clueRows[wk];
    if (pos) strikeTasks.push({ pos: pos, complete: complete });
  }

  for (var si = 0; si < strikeTasks.length; si++) {
    var t = strikeTasks[si];
    var clueCell = sheet.getRange(t.pos[0] + 1, t.pos[1] + 1);
    clueCell.setFontLine(t.complete ? "line-through" : "none");
    clueCell.setFontColor(t.complete ? FONT_CROSSED : FONT_DEFAULT);
  }
}

// ---------------------------------------------------------------------------
// onSelectionChange -- bidirectional highlighting
//   Grid cell click  -> highlight word's grid cells + clue rows in yellow
//   Clue cell click  -> highlight word's grid cells + clue row in yellow
//   Preserves green/red backgrounds from Check/Reveal
// ---------------------------------------------------------------------------

function onSelectionChange(e) {
  try {
    _doSelectionChange(e);
  } catch(ignore) {}
}

function _doSelectionChange(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var cfg   = _getCfg(sheet);
  if (!cfg) return;
  if (!cfg.cellMap || !cfg.wordCells || !cfg.clueRows) return;
  if (cfg.acrossNumCol === undefined || cfg.clueStartRow === undefined) return;

  var r = e.range.getRow()    - 1;   // 0-indexed sheet row
  var c = e.range.getColumn() - 1;   // 0-indexed sheet col

  var GSR = cfg.gridStartRow, GSC = cfg.gridStartCol;
  var GW  = cfg.gridWidth,    GH  = cfg.gridHeight;
  var clueStartR   = cfg.clueStartRow;
  var maxClueRows  = cfg.maxClueRows || 0;
  var clueColStart = cfg.acrossNumCol;
  var clueColEnd   = cfg.downTextCol;
  var clueWidth    = clueColEnd - clueColStart + 1;

  // Bulk-read current backgrounds
  var gridRange = sheet.getRange(GSR + 1, GSC + 1, GH * 2, GW * 2);
  var gridBgs   = gridRange.getBackgrounds();

  var clueRange = null, clueBgs = null;
  if (maxClueRows > 0) {
    clueRange = sheet.getRange(clueStartR + 1, clueColStart + 1, maxClueRows, clueWidth);
    clueBgs   = clueRange.getBackgrounds();
  }

  // ---- Restore previous highlights ----
  var cache   = CacheService.getDocumentCache();
  var hlKey   = "xw_hl_" + _getBaseName(sheet.getName());
  var prevRaw = cache.get(hlKey);
  var gridDirty = false, clueDirty = false;

  if (prevRaw) {
    try {
      var prev = JSON.parse(prevRaw);
      for (var i = 0; i < prev.length; i++) {
        var p = prev[i];
        if (p.t === "g") {
          gridBgs[p.r][p.c] = p.bg;
          gridDirty = true;
        } else if (p.t === "c" && clueBgs) {
          clueBgs[p.r][p.c] = p.bg;
          clueDirty = true;
        }
      }
    } catch(ignore) {}
    cache.remove(hlKey);
  }

  // ---- Safety net: reset any orphaned highlight cells ----
  // Catches highlights stranded by cache eviction or rapid-fire events
  for (var _gr = 0; _gr < gridBgs.length; _gr++) {
    for (var _gc = 0; _gc < gridBgs[_gr].length; _gc++) {
      if (gridBgs[_gr][_gc] === COLOR_HIGHLIGHT) {
        gridBgs[_gr][_gc] = COLOR_WHITE;
        gridDirty = true;
      }
    }
  }
  if (clueBgs) {
    for (var _cr = 0; _cr < clueBgs.length; _cr++) {
      for (var _cc = 0; _cc < clueBgs[_cr].length; _cc++) {
        if (clueBgs[_cr][_cc] === COLOR_HIGHLIGHT) {
          clueBgs[_cr][_cc] = COLOR_WHITE;
          clueDirty = true;
        }
      }
    }
  }

  // ---- Determine word keys from the clicked cell ----
  var wordKeys = [];

  if (_inGrid(r, c, cfg)) {
    var logR = Math.floor((r - GSR) / 2);
    var logC = Math.floor((c - GSC) / 2);
    if (!_isBlackLogical(logR, logC, cfg)) {
      var refs = cfg.cellMap[logR + "," + logC];
      if (refs) {
        if (refs.a !== undefined) wordKeys.push("A" + refs.a);
        if (refs.d !== undefined) wordKeys.push("D" + refs.d);
      }
    }
  } else if (clueBgs &&
             c >= clueColStart && c <= clueColEnd &&
             r >= clueStartR   && r < clueStartR + maxClueRows) {
    var wk = _findClueWordKey(r, c, cfg);
    if (wk) wordKeys.push(wk);
  }

  // ---- Apply highlights for all word keys ----
  var newHl = [];
  var seen  = {};   // deduplicate intersection cells

  for (var wi = 0; wi < wordKeys.length; wi++) {
    var wk = wordKeys[wi];

    // Highlight grid cells (both sub-rows, companion + main)
    var cells = cfg.wordCells[wk];
    if (cells) {
      for (var ci = 0; ci < cells.length; ci++) {
        var lr = cells[ci][0], lc = cells[ci][1];
        var subRows = [lr * 2, lr * 2 + 1];
        var subCols = [lc * 2, lc * 2 + 1];   // companion, main
        for (var ri = 0; ri < subRows.length; ri++) {
          for (var cci = 0; cci < subCols.length; cci++) {
            var gr = subRows[ri], gc = subCols[cci];
            var gk = "g" + gr + "_" + gc;
            if (seen[gk]) continue;
            seen[gk] = true;
            var origG = gridBgs[gr][gc];
            if (origG === COLOR_HIGHLIGHT) origG = COLOR_WHITE;
            newHl.push({t: "g", r: gr, c: gc, bg: origG});
            gridBgs[gr][gc] = COLOR_HIGHLIGHT;
            gridDirty = true;
          }
        }
      }
    }

    // Highlight clue row (number + text cell)
    var pos = cfg.clueRows[wk];
    if (pos && clueBgs) {
      var cr  = pos[0] - clueStartR;
      var ccN = pos[1] - 1 - clueColStart;   // num col in clueBgs
      var ccT = pos[1] - clueColStart;        // text col in clueBgs
      if (cr >= 0 && cr < maxClueRows) {
        if (ccN >= 0 && ccN < clueWidth) {
          var ck1 = "c" + cr + "_" + ccN;
          if (!seen[ck1]) {
            seen[ck1] = true;
            var origC1 = clueBgs[cr][ccN];
            if (origC1 === COLOR_HIGHLIGHT) origC1 = COLOR_WHITE;
            newHl.push({t: "c", r: cr, c: ccN, bg: origC1});
            clueBgs[cr][ccN] = COLOR_HIGHLIGHT;
            clueDirty = true;
          }
        }
        if (ccT >= 0 && ccT < clueWidth) {
          var ck2 = "c" + cr + "_" + ccT;
          if (!seen[ck2]) {
            seen[ck2] = true;
            var origC2 = clueBgs[cr][ccT];
            if (origC2 === COLOR_HIGHLIGHT) origC2 = COLOR_WHITE;
            newHl.push({t: "c", r: cr, c: ccT, bg: origC2});
            clueBgs[cr][ccT] = COLOR_HIGHLIGHT;
            clueDirty = true;
          }
        }
      }
    }
  }

  // ---- Write back ----
  if (gridDirty) gridRange.setBackgrounds(gridBgs);
  if (clueDirty && clueRange) clueRange.setBackgrounds(clueBgs);

  // ---- Persist undo data ----
  if (newHl.length > 0) {
    cache.put(hlKey, JSON.stringify(newHl), 600);
  }
}
