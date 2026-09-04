/**
 * ============================================================
 * PUSHAAN BLENDING SYSTEM - CODE.gs (Backend Logic)
 * ============================================================
 * Paste this as a SECOND file in the same Apps Script project
 * (alongside Setup.gs). File name: Code.gs
 *
 * This contains all backend functions the web app UI calls via
 * google.script.run, plus the doGet() entry point.
 * ============================================================
 */

// ---------- CONFIG / CONSTANTS ----------

function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  return SpreadsheetApp.openById(id);
}

function getFolderId(category, blendType) {
  var key = category + "_" + blendType; // e.g. "BP_RR"
  return PropertiesService.getScriptProperties().getProperty(key);
}

function updateMasterFromTracker() {
  var ss = getSpreadsheet();
  var today = new Date();

  ["BP", "DUST"].forEach(function(cat) {
    var masterSheet = ss.getSheetByName(cat + "_MASTER");
    var trackerSheet = ss.getSheetByName(cat + "_TRACKER");

    var layout = detectLayout(masterSheet);
    var trackerData = trackerSheet.getDataRange().getValues();

    // Build remaining stock map from tracker
    // Tracker cols: A=Particulars, B=Grade, C=Invoice, D=Opening, E=Remaining
    var remainingMap = {};
    for (var i = 1; i < trackerData.length; i++) {
      var row = trackerData[i];
      if (!row[0]) continue;
      var key = (String(row[0]).trim() + "|" + String(row[1]).trim() + "|" + String(row[2]).trim()).toUpperCase();
      remainingMap[key] = parseFloat(row[4]) || 0;
    }

    // Update each variety row in master with remaining stock + today's date
    var masterLastRow = masterSheet.getLastRow();
    var masterData = masterSheet.getRange(1, 1, masterLastRow, 17).getValues();

    layout.varietyRows.forEach(function(rn) {
      var particulars = String(masterData[rn-1][COL.PARTICULARS-1] || '').trim();
      var grade = String(masterData[rn-1][COL.GRADE-1] || '').trim();
      var invoice = String(masterData[rn-1][COL.INV-1] || '').trim();
      var key = (particulars + "|" + grade + "|" + invoice).toUpperCase();

      if (remainingMap[key] !== undefined) {
        var remaining = remainingMap[key];
        // Only update if remaining > 0 (don't overwrite with 0 — variety may have been removed)
        if (remaining > 0) {
          masterSheet.getRange(rn, COL.BAGS).setValue(remaining);
          masterSheet.getRange(rn, COL.MD).setValue(today);
        }
      }
    });

    Logger.log(cat + "_MASTER updated from tracker at " + today);
  });
}

function installDailyTrigger() {
  // Remove any existing triggers for this function first
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "updateMasterFromTracker") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new 2 AM daily trigger
  ScriptApp.newTrigger("updateMasterFromTracker")
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();

  Logger.log("Daily 2 AM trigger installed for updateMasterFromTracker.");
}

function parseDateLocal(dateStr) {
  if (!dateStr) return new Date();
  if (dateStr instanceof Date) return dateStr;
  var parts = String(dateStr).split("-");
  if (parts.length === 3) {
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  }
  return new Date(dateStr);
}

/**
 * Recursively converts any Date objects to ISO strings so the result can
 * safely cross the google.script.run serialization boundary. Without this,
 * returning objects that contain Date values causes the ENTIRE return to
 * arrive as null on the client side.
 */
function sanitizeForClient(obj) {
  if (obj === null || obj === undefined) return obj;
  if (obj instanceof Date) {
    return Utilities.formatDate(obj, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  }
  if (Array.isArray(obj)) {
    return obj.map(function(x) { return sanitizeForClient(x); });
  }
  if (typeof obj === "object") {
    var out = {};
    for (var k in obj) {
      if (obj.hasOwnProperty(k)) out[k] = sanitizeForClient(obj[k]);
    }
    return out;
  }
  return obj;
}

// ============================================================
// DYNAMIC LAYOUT DETECTION + EXCLUSION RULES (per spec doc)
// ============================================================

/**
 * Returns true if a MASTER variety_name should be SKIPPED (not a real tea).
 * Per spec Section 2: ANY BLEND / ANY / MIX placeholders, and DUST-only
 * blend-category labels (PRATAPGARH, RR DUST, SHRI DUST, DUP DUST).
 */
function isExcludedVarietyName(name) {
  if (!name) return true;
  var n = String(name).trim().toUpperCase();
  if (n === "") return true;
  if (n.indexOf("ANY BLEND") === 0) return true;
  if (n.indexOf("ANY ") === 0) return true;
  if (n.indexOf("MIX") === 0) return true;
  if (n.indexOf("PRATAPGARH") === 0) return true;
  if (n.indexOf("RR DUST") === 0) return true;
  if (n.indexOf("SHRI DUST") === 0) return true;
  if (n.indexOf("DUP DUST") === 0) return true;
  if (n.indexOf("LOOSE SAMPLE") === 0) return true;
  if (n.indexOf("950 KG") === 0) return true;
  return false;
}

/**
 * Returns true if a BLEND_LINES variety_name should be SKIPPED.
 * Per spec Section 2: empty, PARTICULARS/TOTAL headers, ANY BLEND/ANY/MIX/LOOSE.
 */
function isExcludedBlendLineName(name) {
  if (!name) return true;
  var n = String(name).trim().toUpperCase();
  if (n === "") return true;
  if (n === "PARTICULARS" || n === "TOTAL") return true;
  if (n.indexOf("ANY BLEND") === 0) return true;
  if (n.indexOf("ANY ") === 0) return true;
  if (n.indexOf("MIX") === 0) return true;
  if (n.indexOf("LOOSE SAMPLE") === 0) return true;
  if (n.indexOf("950 KG") === 0) return true;
  return false;
}

/**
 * Dynamically scans a master sheet and returns its structural layout by
 * reading text labels, so the system never depends on hardcoded row numbers.
 * Returns { totalRow, looseSampleRow, leafRow, dividerRows[], firstVarietyRow,
 *           lastVarietyRow, anyBlendRows[], allRows[] }
 */
function detectLayout(sheet) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(1, 1, lastRow, 17).getValues();

  var layout = {
    totalRow: null, looseSampleRow: null, leafRow: null,
    dividerRows: [], anyBlendRows: [], varietyRows: [],
    firstVarietyRow: 6, lastVarietyRow: null
  };

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 1;
    if (rowNum < 6) continue; // header area

    // Structural labels (Loose Sample / 950 KG / Total / footer) live in COLUMN A
    // in the imported master; variety names live in COLUMN D (Particulars).
    var aCell = data[i][COL.MD - 1];           // column A
    var aName = aCell ? String(aCell).trim().toUpperCase() : "";
    var dCell = data[i][COL.PARTICULARS - 1];  // column D
    var dName = dCell ? String(dCell).trim().toUpperCase() : "";

    // --- TOTAL row: stop everything here. Look in BOTH col A and col D. ---
   if (aName.indexOf("TOTAL") === 0 || dName.indexOf("TOTAL") === 0) {
      layout.totalRow = rowNum;
      break;
    }
   

    // --- Leaf rows (labels in col A) ---
    if (aName.indexOf("LOOSE SAMPLE") === 0) { layout.looseSampleRow = rowNum; continue; }
    if (aName.indexOf("950 KG") === 0) { layout.leafRow = rowNum; continue; }

    // --- Any Blend / placeholder rows (names in col D) ---
    if (dName.indexOf("ANY BLEND") === 0 || dName.indexOf("ANY ") === 0 ||
        (dName.indexOf("MIX") === 0 && dName.length < 12) ||
        dName.indexOf("RR DUST") === 0 || dName.indexOf("SHRI DUST") === 0 ||
        dName.indexOf("DUP DUST") === 0 || dName.indexOf("PRATAPGARH") === 0) {
      layout.anyBlendRows.push(rowNum); continue;
    }

    // --- Divider / section header rows (e.g. "OF", "BP" in col A or col E with no variety name) ---
    if (dName === "") {
      var eCell = data[i][COL.GRADE - 1];
      // a short label in col A (like "OF") or col E, with no particulars = divider
      if ((aName && aName.length <= 4) || eCell) {
        layout.dividerRows.push(rowNum);
      }
      continue; // blank or divider, not a variety
    }

    // --- Otherwise it's a real variety row ---
    layout.varietyRows.push(rowNum);
    layout.lastVarietyRow = rowNum;
  }

  return layout;
}



// Column map (1-indexed) for master / blend sheets
var COL = {
  MD: 1, BAGS: 2, PLACE: 3, PARTICULARS: 4, GRADE: 5, INV: 6, KG_BAG: 7,
  NO_BAGS: 8, PCT: 9, KG: 10, OPEN: 11, SHORTAGE: 12, RATE: 13,
  AMOUNT: 14, CHAMACH: 15, CHAMCH_NAAP: 16, T_WT: 17
};


// ---------- WEB APP ENTRY POINT ----------

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Pushaan Blending System")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// MODULE: DASHBOARD / STOCK READ
// ============================================================

/**
 * Returns dashboard summary for both categories:
 * { BP: {total, critical, low, lastRefreshed}, DUST: {...} }
 */
function getDashboardSummary() {
  var ss = getSpreadsheet();
  var result = {};
  ["BP", "DUST"].forEach(function(cat) {
    var sheet = ss.getSheetByName(cat + "_TRACKER");
    var data = sheet.getDataRange().getValues();
    var total = 0, critical = 0, low = 0, lastUpdated = "";
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // skip empty rows
      total++;
      var remaining = parseFloat(row[4]) || 0; // Remaining Stock col E
      if (remaining <= 1) critical++;
      else if (remaining < 3) low++;
      if (row[16]) lastUpdated = row[16]; // Stock Updated On col Q
    }
    result[cat] = { total: total, critical: critical, low: low, lastUpdated: lastUpdated };
  });
  return sanitizeForClient(result);
}


/**
 * Returns full tracker data for a category, as array of row objects.
 */
function getStockTracker(category) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(category + "_TRACKER");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[headers[c]] = data[i][c];
    }
    rows.push(obj);
  }
  return sanitizeForClient(rows);
}


/**
 * Returns the master variety list for a category, formatted for the
 * "New Blend" screen - includes all variety rows, divider rows, and
 * Any Blend rows, with remaining stock computed live.
 */
/**
 * Returns varieties from the OPPOSITE category, for cross-category use
 * (e.g. pulling an OF/DUST tea into a BP blend). Each is flagged
 * isCrossCategory=true and carries its own category in itemCategory.
 */
function getCrossCategoryVarieties(currentCategory) {
  var other = (currentCategory === "BP") ? "DUST" : "BP";
  var list = getMasterVarieties(other); // already sanitized
  var filtered = [];
  list.forEach(function(v) {
    if (v.isDivider || v.isAnyBlend || v.isFixed) return; // only real teas
    if (!v.particulars) return;
    v.isCrossCategory = true;
    v.itemCategory = other;
    filtered.push(v);
  });
  return sanitizeForClient(filtered);
}

function getMasterVarieties(category) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(category + "_MASTER");
  var layout = detectLayout(sheet);
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(1, 1, lastRow, 17).getValues();
  var trackerMap = buildTrackerRemainingMap(category);

  var varieties = [], anyBlends = [], fixed = [];

  for (var idx = 0; idx < data.length; idx++) {
    var rowNum = idx + 1;
    if (rowNum < 6) continue;
    if (layout.totalRow && rowNum >= layout.totalRow) break;

    var row = data[idx];
    var particulars = row[COL.PARTICULARS - 1];
    var grade = row[COL.GRADE - 1];
    var invoice = row[COL.INV - 1];
    var nameU = particulars ? String(particulars).trim().toUpperCase() : "";

    var isDivider = layout.dividerRows.indexOf(rowNum) !== -1;
    var isAnyBlend = layout.anyBlendRows.indexOf(rowNum) !== -1;
    var isLeaf = (rowNum === layout.leafRow);
    var isLooseSample = (rowNum === layout.looseSampleRow);
    var isFixed = isLeaf || isLooseSample;

    if (isDivider) continue;
    if (!isAnyBlend && !isFixed && nameU === "") continue;

    var key = (particulars + "|" + grade + "|" + invoice).toUpperCase();
    var remaining = trackerMap[key];
    if (remaining === undefined) remaining = row[COL.BAGS - 1];

    var obj = {
      row: rowNum,
      particulars: particulars || "",
      grade: grade || "",
      invoice: invoice || "",
      placeOrLNo: row[COL.PLACE - 1],
      kgBag: row[COL.KG_BAG - 1],
      rate: row[COL.RATE - 1],
      chamach: row[COL.CHAMACH - 1],
      openingBags: row[COL.BAGS - 1],
      remaining: remaining,
      isDivider: false,
      isAnyBlend: isAnyBlend,
      isFixed: isFixed,
      isLeaf: isLeaf,
      isLooseSample: isLooseSample
    };

    if (isFixed) { fixed.push(obj); }
    else if (isAnyBlend) { anyBlends.push(obj); }
    else { varieties.push(obj); }
  }

  varieties.sort(function(a, b) {
    return String(a.particulars).trim().toUpperCase().localeCompare(
      String(b.particulars).trim().toUpperCase()
    );
  });

  return sanitizeForClient([].concat(varieties, anyBlends, fixed));
}

function buildTrackerRemainingMap(category) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(category + "_TRACKER");
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var key = (row[0] + "|" + row[1] + "|" + row[2]).toUpperCase();
    map[key] = parseFloat(row[4]) || 0; // Remaining Stock col E
  }
  return map;
}


// ============================================================
// MODULE: STOCK GUARD (validation before saving a blend)
// ============================================================

/**
 * Checks if the requested bags for each variety in a blend exceed
 * available remaining stock. Returns array of warnings (empty = OK).
 * items: [{particulars, grade, invoice, bagsRequested, category}]
 */
function checkStockAvailability(items) {
  var warnings = [];
  var trackerMaps = {
    BP: buildTrackerRemainingMap("BP"),
    DUST: buildTrackerRemainingMap("DUST")
  };

  items.forEach(function(item) {
    if (!item.particulars || !item.bagsRequested || item.bagsRequested <= 0) return;
    var cat = item.itemCategory || item.category; // the category this VARIETY belongs to (for cross-category)
    var map = trackerMaps[cat];
    var key = (item.particulars + "|" + item.grade + "|" + item.invoice).toUpperCase();
    var remaining = map[key];
    if (remaining === undefined) remaining = 0;

    if (item.bagsRequested > remaining) {
      warnings.push({
        particulars: item.particulars,
        grade: item.grade,
        invoice: item.invoice,
        requested: item.bagsRequested,
        available: remaining,
        message: "Only " + remaining + " bag(s) available for " + item.particulars + " (" + item.grade + ", Inv " + item.invoice + ") - you requested " + item.bagsRequested
      });
    } else if (remaining - item.bagsRequested <= 0) {
      warnings.push({
        particulars: item.particulars,
        grade: item.grade,
        invoice: item.invoice,
        requested: item.bagsRequested,
        available: remaining,
        message: item.particulars + " (" + item.grade + ", Inv " + item.invoice + ") will hit ZERO stock after this blend",
        isZeroWarning: true
      });
    }
  });

  return sanitizeForClient(warnings);
}


// ============================================================
// MODULE: BATCH NUMBER GENERATION
// ============================================================

/**
 * Returns the next batch number for a category+blendType, e.g. "RR-22"
 * Does NOT increment the counter - call incrementBatchCounter() on save.
 */
function getNextBatchNumber(category, blendType) {
  // Per spec Section 7: scan all existing blends of this category+type,
  // extract the numeric part after the prefix, take MAX, add 1.
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var data = sheet.getDataRange().getValues();

  var prefix = blendType + "-";
  var maxN = 0;
  var seen = {};
  for (var i = 1; i < data.length; i++) {
    var rowCat = data[i][2];      // Category
    var rowType = data[i][3];     // Blend_Type
    var batchNo = data[i][5];     // Batch_No
    if (rowCat !== category || rowType !== blendType || !batchNo) continue;
    if (seen[batchNo]) continue;
    seen[batchNo] = true;
    var m = String(batchNo).match(/-(\d+)$/);
    if (m) { var n = parseInt(m[1], 10); if (n > maxN) maxN = n; }
  }
  return String(prefix + (maxN + 1));
}

function incrementBatchCounter(category, blendType) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("CONFIG");
  var data = sheet.getDataRange().getValues();
  var key = category + "_" + blendType + "_COUNTER";
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      var next = (parseInt(data[i][1]) || 0) + 1;
      sheet.getRange(i + 1, 2).setValue(next);
      return next;
    }
  }
  return null;
}


// ============================================================
// MODULE: CREATE NEW BLEND
// ============================================================

/**
 * Main function called when "Save Blend" is tapped.
 * blendInfo: {
 *   category: "BP"|"DUST", blendType: "RR"|"DUP"|"SK"|"LOOSE",
 *   customer, date, basePrice (H2 value),
 *   items: [{ row, particulars, grade, invoice, placeOrLNo, kgBag,
 *             bags, rate, shortage, chamach, itemCategory }]
 * }
 * Returns { success, batchNo, sheetUrl, blendId, warnings }
 */
function createBlend(blendInfo) {
  var status = blendInfo.status || "Draft"; // "Draft" or "Finalized"

  // 1. Stock guard check - only meaningful for Finalized (Draft doesn't touch stock)
  var warnings = [];
  if (status === "Finalized") {
    warnings = checkStockAvailability(blendInfo.items.map(function(it) {
      return {
        particulars: it.particulars, grade: it.grade, invoice: it.invoice,
        bagsRequested: parseFloat(it.bags) || 0,
        itemCategory: it.itemCategory || blendInfo.category
      };
    }));

    var blockingWarnings = warnings.filter(function(w) { return !w.isZeroWarning; });
    if (blockingWarnings.length > 0 && !blendInfo.overrideWarnings) {
      return sanitizeForClient({ success: false, warnings: warnings });
    }
  }

  // 2. Generate batch number
  var batchNo = blendInfo.batchNo || getNextBatchNumber(blendInfo.category, blendInfo.blendType);

  // 3. Create the per-blend Google Sheet
  var blendId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss") + "_" + batchNo;
  var sheetUrl = createBlendSheet(blendInfo, batchNo, blendId, status);

  // 4. Log to BLEND_LOG (one row per item)
  logBlendToLog(blendInfo, batchNo, blendId, sheetUrl, status);

  // 5. Increment batch counter (only for Finalized - Drafts don't consume batch numbers
  //    permanently... but to keep numbering simple and avoid collisions, we increment
  //    on first save regardless. This matches "batch number assigned when blend created".)
  incrementBatchCounter(blendInfo.category, blendInfo.blendType);

  // 6. Refresh stock tracker (only Finalized rows affect it, but safe to always refresh)
  var refreshResult = refreshStockTracker();

   return sanitizeForClient({ success: true, batchNo: batchNo, sheetUrl: sheetUrl, blendId: blendId, status: status, warnings: warnings, zeroStock: refreshResult.zeroStock });
}


/**
 * Writes one row per blend item into BLEND_LOG.
 */
function logBlendToLog(blendInfo, batchNo, blendId, sheetUrl, status) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var rows = [];

  blendInfo.items.forEach(function(item) {
    var bags = parseFloat(item.bags) || 0;
    if (bags === 0) return;

    var kg = (parseFloat(item.kgBag) || 0) * bags;
    var shortage = parseFloat(item.shortage) || 0;
    var rate = parseFloat(item.rate) || 0;
    var itemCat = item.itemCategory || blendInfo.category;
    var amount = (itemCat === "BP") ? (kg + shortage) * rate : kg * rate;
    var chamach = parseFloat(item.chamach) || 0;
    var tWt = kg * chamach;

    var finalizedBags = (item.finalizedBags !== undefined && item.finalizedBags !== null && item.finalizedBags !== "")
      ? parseFloat(item.finalizedBags)
      : (status === "Finalized" ? bags : 0);

    var openCol = (status === "Finalized") ? (bags - finalizedBags) : 0;

    // Col B: finalization date (today) if Finalized, else creation date
    // Col V: original creation date (always preserved)
    var finalizationDate = status === "Finalized"
      ? Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
      : (blendInfo.date ? Utilities.formatDate(parseDateLocal(blendInfo.date), Session.getScriptTimeZone(), "yyyy-MM-dd") : "");
    var creationDate = blendInfo.date
      ? Utilities.formatDate(parseDateLocal(blendInfo.date), Session.getScriptTimeZone(), "yyyy-MM-dd")
      : "";

    rows.push([
      blendId,
      finalizationDate,
      item.itemCategory || blendInfo.category,
      blendInfo.blendType,
      blendInfo.customer || "",
      batchNo,
      blendInfo.basePrice,
      item.particulars,
      item.grade,
      item.invoice,
      bags,
      item.kgBag,
      kg,
      shortage,
      rate,
      amount,
      chamach,
      tWt,
      sheetUrl,
      status,
      finalizedBags,
      creationDate
    ]);

    item._openValue = openCol;
  });

  if (rows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function buildVarietyRowLookup(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 6) return {};
  var data = sheet.getRange(1, 1, lastRow, COL.INV).getValues();
  var lookup = {};
  for (var i = 5; i < data.length; i++) {
    var p = String(data[i][COL.PARTICULARS - 1] || '').trim().toUpperCase();
    var g = String(data[i][COL.GRADE - 1] || '').trim().toUpperCase();
    var inv = String(data[i][COL.INV - 1] || '').trim().toUpperCase();
    if (p && p !== 'PARTICULARS' && p !== 'TOTAL') {
      lookup[p + '|' + g + '|' + inv] = i + 1;
    }
  }
  return lookup;
}


function createBlendSheet(blendInfo, batchNo, blendId, status) {
  var ss = getSpreadsheet();
  var masterSheet = ss.getSheetByName(blendInfo.category + "_MASTER");
  var layout = detectLayout(masterSheet);

  var dateStr = Utilities.formatDate(parseDateLocal(blendInfo.date), Session.getScriptTimeZone(), "dd.MM.yyyy");
  var fileName = dateStr + " " + batchNo;

  // 1. Create a new spreadsheet for this blend.
  var newSs = SpreadsheetApp.create(fileName);

  // 2. NATIVELY COPY the master tab into the new spreadsheet. copyTo() preserves
  //    EVERYTHING: fonts, sizes, colors, borders, merges, column widths, row
  //    heights, number formats, alignment, and all formulas. This guarantees the
  //    blend sheet is a true replica of the current master.
  var copied = masterSheet.copyTo(newSs);
  copied.setName(blendInfo.blendType);

  // 3. Remove the default empty "Sheet1" that create() made.
  var defaultSheet = newSs.getSheetByName("Sheet1");
  if (defaultSheet) newSs.deleteSheet(defaultSheet);

  var newSheet = copied;

  // 3b. Replace the "MASTER FILE DON'T CHANGE" header text with the blend type.
  //     Runs on the COPIED blend sheet (never the master). Handles merged cells.
  try {
    var searchValues = newSheet.getDataRange().getDisplayValues();
    outerLoop:
    for (var sr = 0; sr < searchValues.length; sr++) {
      for (var sc = 0; sc < searchValues[sr].length; sc++) {
        var txt = String(searchValues[sr][sc] || "").trim().toUpperCase();
        if (txt.indexOf("MASTER FILE") !== -1) {
          var targetCell = newSheet.getRange(sr + 1, sc + 1);
          if (targetCell.isPartOfMerge()) {
            targetCell = targetCell.getMergedRanges()[0].getCell(1, 1);
          }
          targetCell.setValue(blendInfo.blendType);
          break outerLoop;
        }
      }
    }
  } catch (e) {
    Logger.log("Could not replace MASTER FILE header: " + e);
  }

  // 4. Clear out the No.Bags / Rate / Shortage / Chamach / Open inputs from all
  //    variety + Any Blend + leaf rows, so the blend starts from the master but
  //    with a clean slate (we then write in only what this blend uses).
  var clearRows = [].concat(layout.varietyRows, layout.anyBlendRows);
  if (layout.leafRow) clearRows.push(layout.leafRow);
  if (layout.looseSampleRow) clearRows.push(layout.looseSampleRow);
  clearRows.forEach(function(rn) {
    newSheet.getRange(rn, COL.NO_BAGS).setValue(0);
  });

  // 5. Write the actual bag values, rates, shortage, chamach for each item used.
  // 5. Write bag values — content-based lookup so row shifts don't cause misses
  var rowLookup5 = buildVarietyRowLookup(newSheet);
  blendInfo.items.forEach(function(item) {
    if (item.isCrossCategory) return; // handled by writeCrossCategoryRows

    // Fixed rows (leaf/loose sample) have no particulars — use item.row directly
    if (!item.particulars) {
  if (!item.row || typeof item.row !== 'number') return;
  newSheet.getRange(item.row, COL.NO_BAGS).setValue(parseFloat(item.bags) || 0);
  newSheet.getRange(item.row, COL.NO_BAGS).setNumberFormat("0.00");
  return;
}

    // All other varieties — find row by content match
    var key = String(item.particulars).trim().toUpperCase() + '|' +
              String(item.grade || '').trim().toUpperCase() + '|' +
              String(item.invoice || '').trim().toUpperCase();
    var rn = rowLookup5[key] || (typeof item.row === 'number' ? item.row : null);
    if (!rn) return;

    var bags = parseFloat(item.bags) || 0;
    newSheet.getRange(rn, COL.NO_BAGS).setValue(bags);
    if (item.rate !== undefined && item.rate !== '') newSheet.getRange(rn, COL.RATE).setValue(parseFloat(item.rate));
    if (item.shortage !== undefined && item.shortage !== '') newSheet.getRange(rn, COL.SHORTAGE).setValue(parseFloat(item.shortage));
    if (item.chamach !== undefined && item.chamach !== '') newSheet.getRange(rn, COL.CHAMACH).setValue(parseFloat(item.chamach));
    var finalizedBags = (item.finalizedBags !== undefined && item.finalizedBags !== null && item.finalizedBags !== '')
      ? parseFloat(item.finalizedBags) : (status === 'Finalized' ? bags : 0);
    newSheet.getRange(rn, COL.OPEN).setValue(status === 'Finalized' ? (bags - finalizedBags) : 0);
  });

  // 5b. CROSS-CATEGORY OF ROWS: write any cross-category (e.g. OF/DUST) teas used
  //     in this blend as real inserted rows in the sheet, so the sheet's own
  //     SUM totals include them and match the app. See writeCrossCategoryRows().
  writeCrossCategoryRows(newSheet, blendInfo, status);

  // 6. Update batch number, date, customer, base price (same cells as master layout).
  newSheet.getRange(4, 4).setValue(batchNo);                    // D4 = Batch No
  var blendDate = parseDateLocal(blendInfo.date);
newSheet.getRange(4, 5).setValue(blendDate);
newSheet.getRange(4, 5).setNumberFormat("dd/mm/yyyy");
  newSheet.getRange(2, 8).setValue(blendInfo.basePrice);        // H2 = base price
  if (blendInfo.customer) {
    newSheet.getRange(2, 3).setValue(blendInfo.customer);       // C2 = customer
  }

  // 7. Move the new file into the correct Drive folder (Category/BlendType/MonthYear).
  var folderId = getOrCreateMonthFolder(blendInfo.category, blendInfo.blendType, parseDateLocal(blendInfo.date));
  var folder = DriveApp.getFolderById(folderId);
  var file = DriveApp.getFileById(newSs.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return newSs.getUrl();
}

/**
 * Writes cross-category teas (e.g. OF/DUST teas used inside a BP blend) as real
 * rows in the blend sheet, inserted right AFTER the black "OF" divider band.
 *
 * Method: insert N blank rows after the OF band, copy an existing variety row
 * into each (for FORMATTING only - fonts, borders, colours, number formats),
 * then explicitly write every value AND every formula (KG, %, Amount, Chamch
 * Naap, T.Wt) referencing the correct row + the live Total row. Writing the
 * formulas explicitly (instead of trusting the copied ones) removes any doubt
 * about relative/absolute reference adjustment after the insert.
 *
 * Amount uses the DUST style (KG x Rate, no shortage added) for OF rows.
 *
 * Because the sheet's Total formulas use a CONTINUOUS range spanning the OF band
 * (e.g. =SUM(J7:J35)), inserting rows inside that range makes Google Sheets
 * auto-expand the SUM, so the total picks up the OF rows and matches the app.
 *
 * Does nothing if the blend has no cross-category items.
 */
function writeCrossCategoryRows(sheet, blendInfo, status) {
  var crossItems = (blendInfo.items || []).filter(function(it) {
    var isCross = it.isCrossCategory ||
      (it.itemCategory && it.itemCategory !== blendInfo.category);
    var bags = parseFloat(it.bags) || 0;
    return isCross && bags > 0;
  });
  if (crossItems.length === 0) return;

  // --- Find the black "OF" divider band row dynamically ---
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), COL.T_WT);
  var scan = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  var ofBandRow = null;
  for (var r = 0; r < scan.length; r++) {
    var rowNum = r + 1;
    if (rowNum < 6) continue;
    for (var c = 0; c < Math.min(scan[r].length, COL.GRADE); c++) { // cols A-E
      var txt = String(scan[r][c] || "").trim().toUpperCase();
      if (txt === "OF") { ofBandRow = rowNum; break; }
    }
    if (ofBandRow) break;
  }

  if (!ofBandRow) {
    var fbLayout = detectLayout(sheet);
    if (fbLayout.totalRow) {
      ofBandRow = fbLayout.totalRow - 1;
    } else {
      Logger.log("writeCrossCategoryRows: no OF band and no Total row found; skipping.");
      return;
    }
  }

  // --- Pick a source row to copy FORMATTING from (a real variety row above the band) ---
  var srcLayout = detectLayout(sheet);
  var sourceRow = null;
  for (var i = 0; i < srcLayout.varietyRows.length; i++) {
    if (srcLayout.varietyRows[i] < ofBandRow) sourceRow = srcLayout.varietyRows[i];
  }
  if (!sourceRow) sourceRow = srcLayout.varietyRows.length ? srcLayout.varietyRows[0] : 6;

  var n = crossItems.length;

  // Insert N blank rows AFTER the OF band. Everything below shifts down; the
  // continuous SUM ranges auto-expand to include the new rows.
  sheet.insertRowsAfter(ofBandRow, n);

  // After the insert, re-detect the layout so we know where the Total row is now.
  var newLayout = detectLayout(sheet);
  var totalRow = newLayout.totalRow; // shifted total row (e.g. 36 -> 38)

  var srcRange = sheet.getRange(sourceRow, 1, 1, COL.T_WT);

  for (var k = 0; k < n; k++) {
    var item = crossItems[k];
    var destRowNum = ofBandRow + 1 + k;
    var destRange = sheet.getRange(destRowNum, 1, 1, COL.T_WT);

    // Copy the source row for FORMATTING ONLY (so the OF row looks identical:
    // fonts, borders, backgrounds, number formats). Formulas are then overwritten
    // explicitly below, so we don't depend on copyTo's reference adjustment.
    srcRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    var bags = parseFloat(item.bags) || 0;
    var kgBag = parseFloat(item.kgBag) || 0;
    var rate = (item.rate !== undefined && item.rate !== "" && item.rate !== null) ? parseFloat(item.rate) : 0;
    var shortage = (item.shortage !== undefined && item.shortage !== "" && item.shortage !== null) ? parseFloat(item.shortage) : 0;
    var chamach = (item.chamach !== undefined && item.chamach !== "" && item.chamach !== null) ? parseFloat(item.chamach) : 0;

    // --- Input cells ---
    sheet.getRange(destRowNum, COL.MD).setValue(parseDateLocal(blendInfo.date));   // A = date
    sheet.getRange(destRowNum, COL.BAGS).setValue("");                        // B = opening stock (blank for OF)
    sheet.getRange(destRowNum, COL.PLACE).setValue(item.placeOrLNo || "");    // C = warehouse
    sheet.getRange(destRowNum, COL.PARTICULARS).setValue(item.particulars || "");
    sheet.getRange(destRowNum, COL.GRADE).setValue(item.grade || "");
    sheet.getRange(destRowNum, COL.INV).setValue(item.invoice || "");
    sheet.getRange(destRowNum, COL.KG_BAG).setValue(kgBag);
    sheet.getRange(destRowNum, COL.NO_BAGS).setValue(bags);
    sheet.getRange(destRowNum, COL.OPEN).setValue(0);
    sheet.getRange(destRowNum, COL.SHORTAGE).setValue(shortage);
    sheet.getRange(destRowNum, COL.RATE).setValue(rate);
    sheet.getRange(destRowNum, COL.CHAMACH).setValue(chamach);

    // --- Formulas (explicit, referencing this row + the live total row) ---
    sheet.getRange(destRowNum, COL.KG).setFormula("=G" + destRowNum + "*H" + destRowNum);
    if (totalRow) {
      sheet.getRange(destRowNum, COL.PCT).setFormula("=(J" + destRowNum + "/$J$" + totalRow + ")");
    }
    // Amount: DUST style -> KG x Rate (no shortage added).
    sheet.getRange(destRowNum, COL.AMOUNT).setFormula("=J" + destRowNum + "*M" + destRowNum);
    sheet.getRange(destRowNum, COL.CHAMCH_NAAP).setFormula("=O" + destRowNum + "*18.5");
    sheet.getRange(destRowNum, COL.T_WT).setFormula("=J" + destRowNum + "*O" + destRowNum);
  }
}

/**
 * Creates a new Google Sheet for this blend, formatted like the master
 * blend sheet, with bags filled in. Saves it in the correct Drive folder
 * (Category/BlendType/MonthYear). Returns the sheet URL.
 */


/**
 * Copies background colors, font colors, bold, column widths from
 * source sheet to target sheet for the given range.
 */
function copySheetFormatting(source, target, numRows, numCols) {
  var srcRange = source.getRange(1, 1, numRows, numCols);
  var backgrounds = srcRange.getBackgrounds();
  var fontColors = srcRange.getFontColors();
  var fontWeights = srcRange.getFontWeights();
  var numberFormats = srcRange.getNumberFormats();

  var tgtRange = target.getRange(1, 1, numRows, numCols);
  tgtRange.setBackgrounds(backgrounds);
  tgtRange.setFontColors(fontColors);
  tgtRange.setFontWeights(fontWeights);
  tgtRange.setNumberFormats(numberFormats);

  for (var c = 1; c <= numCols; c++) {
    target.setColumnWidth(c, source.getColumnWidth(c));
  }
  target.setFrozenRows(source.getFrozenRows());
}


/**
 * Gets or creates the Drive folder for Category/BlendType/MonthYear
 * e.g. BP/RR/JUNE 2026
 */
function getOrCreateMonthFolder(category, blendType, date) {
  var folderId = getFolderId(category, blendType); // e.g. BP_RR -> folder ID for BP/RR/
  var parentFolder = DriveApp.getFolderById(folderId);

  var monthName = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM yyyy").toUpperCase();

  var folders = parentFolder.getFoldersByName(monthName);
  if (folders.hasNext()) {
    return folders.next().getId();
  }
  var newFolder = parentFolder.createFolder(monthName);
  return newFolder.getId();
}


// ============================================================
// MODULE: ADD NEW VARIETY
// ============================================================

/**
 * Adds a new variety row to BP_MASTER or DUST_MASTER, just before
 * the divider row (so it stays within the Total row's SUM range).
 * varietyInfo: {category, particulars, grade, invoice, bags, placeOrLNo,
 *               kgBag, rate, chamach, stockDate}
 * Returns {success, row}
 */
function addNewVariety(varietyInfo) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(varietyInfo.category + "_MASTER");
  var layout = detectLayout(sheet);

  var insertBeforeRow;
  if (layout.varietyRows && layout.varietyRows.length > 0) {
    var newNameUpper = String(varietyInfo.particulars || '').trim().toUpperCase();
    var colDData = sheet.getRange(1, COL.PARTICULARS, sheet.getLastRow(), 1).getValues();
    insertBeforeRow = null;
    for (var vi = 0; vi < layout.varietyRows.length; vi++) {
      var rn = layout.varietyRows[vi];
      var existingName = String(colDData[rn - 1][0] || '').trim().toUpperCase();
      if (existingName > newNameUpper) { insertBeforeRow = rn; break; }
    }
    if (!insertBeforeRow) insertBeforeRow = layout.lastVarietyRow + 1;
  } else {
    var candidates = [].concat(layout.dividerRows, layout.anyBlendRows,
      [layout.looseSampleRow, layout.leafRow, layout.totalRow]).filter(function(x) { return x; });
    insertBeforeRow = Math.min.apply(null, candidates);
  }

  sheet.insertRowBefore(insertBeforeRow);
  var newRow = insertBeforeRow;

  var newLayout = detectLayout(sheet);
  var totalRow = newLayout.totalRow;

  sheet.getRange(newRow, COL.MD).setValue(varietyInfo.stockDate ? parseDateLocal(varietyInfo.stockDate) : new Date());
  sheet.getRange(newRow, COL.BAGS).setValue(parseFloat(varietyInfo.bags) || 0);
  sheet.getRange(newRow, COL.PLACE).setValue(varietyInfo.placeOrLNo || "");
  sheet.getRange(newRow, COL.PARTICULARS).setValue(varietyInfo.particulars);
  sheet.getRange(newRow, COL.GRADE).setValue(varietyInfo.grade);
  sheet.getRange(newRow, COL.INV).setValue(varietyInfo.invoice);
  sheet.getRange(newRow, COL.KG_BAG).setValue(parseFloat(varietyInfo.kgBag) || 0);
  sheet.getRange(newRow, COL.NO_BAGS).setValue(0);
  sheet.getRange(newRow, COL.RATE).setValue(parseFloat(varietyInfo.rate) || 0);
  sheet.getRange(newRow, COL.CHAMACH).setValue(parseFloat(varietyInfo.chamach) || 0);

  if (totalRow) {
    sheet.getRange(newRow, COL.PCT).setFormula("=(J" + newRow + "/$J$" + totalRow + ")");
  } else {
    // totalRow not found - use MATCH to find Total row dynamically
    sheet.getRange(newRow, COL.PCT).setFormula("=IFERROR(J" + newRow + "/INDIRECT(\"J\"&MATCH(\"Total\",A:A,0)),0)");
  }
  sheet.getRange(newRow, COL.KG).setFormula("=G" + newRow + "*H" + newRow);
  sheet.getRange(newRow, COL.AMOUNT).setFormula(
    varietyInfo.category === "BP"
      ? "=(J" + newRow + "+L" + newRow + ")*M" + newRow
      : "=J" + newRow + "*M" + newRow
  );
  sheet.getRange(newRow, COL.CHAMCH_NAAP).setFormula("=O" + newRow + "*18.5");
  sheet.getRange(newRow, COL.T_WT).setFormula("=J" + newRow + "*O" + newRow);

  if (newRow > 6) copyRowFormatting(sheet, newRow - 1, newRow, 17);

  restoreMasterFormulas();
refreshStockTracker();
return sanitizeForClient({ success: true, row: newRow });
}

function copyRowFormatting(sheet, fromRow, toRow, numCols) {
  var src = sheet.getRange(fromRow, 1, 1, numCols);
  var tgt = sheet.getRange(toRow, 1, 1, numCols);
  tgt.setBackgrounds(src.getBackgrounds());
  tgt.setFontColors(src.getFontColors());
  tgt.setFontWeights(src.getFontWeights());
  tgt.setNumberFormats(src.getNumberFormats());
}


// ============================================================
// MODULE: UPDATE OPENING STOCK (existing variety, new lot)
// ============================================================

/**
 * Updates the BAGS (opening stock) for an existing variety row,
 * and refreshes the M/D timestamp.
 */
function updateOpeningStock(category, row, newBags) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(category + "_MASTER");

  sheet.getRange(row, COL.BAGS).setValue(parseFloat(newBags) || 0);
  sheet.getRange(row, COL.MD).setValue(new Date());

  refreshStockTracker();
  return sanitizeForClient({ success: true });
}


// ============================================================
// MODULE: REFRESH STOCK TRACKER
// ============================================================

/**
 * Reads BP_MASTER + DUST_MASTER + BLEND_LOG, computes consumption per
 * variety (triple-key match: Particulars+Grade+Invoice), aggregates by
 * blend type, computes remaining stock, writes to BP_TRACKER/DUST_TRACKER.
 * Also detects cross-category usage (DUST variety used in BP blend, etc).
 *
 * Returns { success, alerts: [{category, particulars, grade, invoice, remaining, level}] }
 */
function refreshStockTracker() {
  var ss = getSpreadsheet();
  var blendLogSheet = ss.getSheetByName("BLEND_LOG");
  var logData = blendLogSheet.getDataRange().getValues();

  // Pre-parse all finalized blend lines into a flat list with normalized dates.
  // Each entry: { blendCategory, blendType, key, particulars, grade, invoice, bags, dateMs }
  var lines = [];
  for (var i = 1; i < logData.length; i++) {
    var row = logData[i];
    if (row[19] !== "Finalized") continue;          // only Finalized blends affect stock
    var particulars = row[7];
    if (isExcludedBlendLineName(particulars)) continue; // skip placeholder/header lines
    var grade = row[8], invoice = row[9];
    var bags = parseFloat(row[10]) || 0;
    if (bags <= 0) continue;
    var d = row[1] instanceof Date ? row[1] : parseDateLocal(row[1]);
    lines.push({
      blendCategory: row[2],            // the blend's own category (BP or DUST)
      blendType: row[3],
      key: (particulars + "|" + grade + "|" + invoice).toUpperCase(),
      particulars: particulars, grade: grade, invoice: invoice,
      bags: bags,
      dateMs: dateOnlyMs(d)
    });
  }

  var alerts = [];
  var zeroStock = [];

  ["BP", "DUST"].forEach(function(cat) {
    var masterSheet = ss.getSheetByName(cat + "_MASTER");
    var layout = detectLayout(masterSheet);
    var lastRow = masterSheet.getLastRow();
    var masterData = masterSheet.getRange(1, 1, lastRow, 17).getValues();

    var trackerRows = [];
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    masterData.forEach(function(row, idx) {
      var rowNum = idx + 1;
      if (rowNum < 6) return;
      var particulars = row[COL.PARTICULARS - 1];

      // EXCLUSION RULES (spec Section 2)
      if (isExcludedVarietyName(particulars)) return;
      var grade = row[COL.GRADE - 1];
      var invoice = row[COL.INV - 1];
      var opening = parseFloat(row[COL.BAGS - 1]) || 0;
      var stockDateRaw = row[COL.MD - 1];
      var stockDate = stockDateRaw instanceof Date ? stockDateRaw : (stockDateRaw ? new Date(stockDateRaw) : null);

      if (!stockDate || isNaN(stockDate.getTime())) return; // invalid stock_date -> skip
      if (opening <= 0) return;                              // zero/empty stock -> skip

      var stockDateMs = dateOnlyMs(stockDate);
      var key = (particulars + "|" + grade + "|" + invoice).toUpperCase();

      var c = { RR:0, DUP:0, SK:0, LOOSE:0, CROSS:0, RRD:[], DUPD:[], SKD:[], LOOSED:[], CROSSD:[] };

      // Scan all finalized blend lines for matches (date-filtered)
      lines.forEach(function(ln) {
        if (ln.key !== key) return;
        if (ln.dateMs < stockDateMs) return;  // DATE FILTER: ignore blends before this stock's date

        var dStr = Utilities.formatDate(new Date(ln.dateMs), Session.getScriptTimeZone(), "dd/MM/yyyy");

        if (ln.blendCategory === cat) {
          // same-category usage -> bucket by blend type
          if (ln.blendType === "RR") { c.RR += ln.bags; c.RRD.push(dStr); }
          else if (ln.blendType === "DUP") { c.DUP += ln.bags; c.DUPD.push(dStr); }
          else if (ln.blendType === "SK") { c.SK += ln.bags; c.SKD.push(dStr); }
          else if (ln.blendType === "LOOSE") { c.LOOSE += ln.bags; c.LOOSED.push(dStr); }
          else { c.SK += ln.bags; c.SKD.push(dStr); } // AM merges into SK
        } else {
          // cross-category usage (e.g. this BP variety used inside a DUST blend)
          c.CROSS += ln.bags; c.CROSSD.push(dStr);
        }
      });

      var totalUsed = c.RR + c.DUP + c.SK + c.LOOSE + c.CROSS;
      var remaining = opening - totalUsed;

      var level = remaining <= 1 ? "critical" : (remaining < 3 ? "low" : "ok");
      if (level !== "ok") {
        alerts.push({ category: cat, particulars: particulars, grade: grade, invoice: invoice, remaining: remaining, level: level });
      }
      if (remaining <= 0) {
        zeroStock.push({ category: cat, particulars: particulars, grade: grade, invoice: invoice, row: rowNum });
      }

      trackerRows.push([
        particulars, grade, invoice, opening, remaining,
        c.RR, c.RRD.join(", "),
        c.DUP, c.DUPD.join(", "),
        c.SK, c.SKD.join(", "),
        c.LOOSE, c.LOOSED.join(", "),
        c.CROSS, c.CROSSD.join(", "),
        totalUsed,
        nowStr
      ]);
    });

    var trackerSheet = ss.getSheetByName(cat + "_TRACKER");
    var tLast = trackerSheet.getLastRow();
    if (tLast > 1) trackerSheet.getRange(2, 1, tLast - 1, 17).clearContent();
    if (trackerRows.length > 0) {
      trackerSheet.getRange(2, 1, trackerRows.length, 17).setValues(trackerRows);
    }
  });

  return sanitizeForClient({ success: true, alerts: alerts, zeroStock: zeroStock });
}


/**
 * Returns the millisecond timestamp of a date with the time component stripped
 * (local midnight), so date-only comparisons work regardless of time-of-day.
 */
function dateOnlyMs(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return 0;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}



function getBlendHistory() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var data = sheet.getDataRange().getValues();

  var groups = {};
  var order = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var blendId = row[0];
    if (!groups[blendId]) {
      // Use col V (creation date) for display if available, else col B
      var displayDate = (row.length > 21 && row[21]) ? row[21] : row[1];
     groups[blendId] = {
        blendId: blendId,
        date: displayDate,
        finalizationDate: row[1],
        category: row[2],
        blendType: row[3],
        customer: row[4],
        batchNo: row[5],
        basePrice: row[6],
        sheetUrl: row[18],
        status: row[19] || "Draft",
        items: [],
        totalKg: 0
      };
      order.push(blendId);
    }
    groups[blendId].items.push({
      particulars: row[7], grade: row[8], invoice: row[9],
      bags: row[10], kg: row[12], rate: row[14], amount: row[15],
      finalizedBags: row[20],
      open: (row[19] === "Finalized") ? (parseFloat(row[10]) - parseFloat(row[20])) : 0
    });
    groups[blendId].totalKg += parseFloat(row[12]) || 0;
  }

  var result = order.map(function(id) { return groups[id]; });
  result.reverse();
  return sanitizeForClient(result);
}

/**
 * Returns full detail for one blend by Blend_ID.
 */
function getBlendDetail(blendId) {
  var history = getBlendHistory();
  for (var i = 0; i < history.length; i++) {
    if (history[i].blendId === blendId) return sanitizeForClient(history[i]);
  }
  return null;
}

function getBlendPrintData(blendId) {

  var history = getBlendHistory();
  var blend = null;

  for (var i = 0; i < history.length; i++) {
    if (history[i].blendId === blendId) {
      blend = history[i];
      break;
    }
  }

  if (!blend) throw new Error("Blend not found");
  if (!blend.sheetUrl) throw new Error("Blend sheet URL not found");

  var match = String(blend.sheetUrl).match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) throw new Error("Could not parse Google Sheet ID");

  var ssId = match[1];
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheets()[0];
  var gid = sheet.getSheetId();

  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 12);
  var scanData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  // Find the "S.R." row (matched loosely - ignores spaces/dots)
  var srRow = null;
  for (var r = 0; r < scanData.length; r++) {
    for (var c = 0; c < scanData[r].length; c++) {
      var raw = String(scanData[r][c] || '');
      var norm = raw.replace(/[\s\.]/g, '').toUpperCase();
      if (norm === 'SR') { srRow = r + 1; break; }
    }
    if (srRow) break;
  }

  // Fallback: a few rows past Total, if S.R. isn't found
  if (!srRow) {
    for (var r2 = 0; r2 < scanData.length; r2++) {
      var a = String(scanData[r2][0] || '').trim().toUpperCase();
      var d = String(scanData[r2][3] || '').trim().toUpperCase();
      if (a.indexOf('TOTAL') === 0 || d.indexOf('TOTAL') === 0) { srRow = r2 + 3; break; }
    }
  }
  if (!srRow) srRow = lastRow;
  if (srRow > lastRow) srRow = lastRow;

  var range = "A1:L" + srRow;

  // Native Google Sheets PDF export - preserves all colors, borders, fonts exactly
  var exportUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export"
    + "?format=pdf"
    + "&gid=" + gid
    + "&range=" + encodeURIComponent(range)
    + "&size=A4"
    + "&portrait=true"
    + "&fitw=true"
    + "&top_margin=0.3"
    + "&bottom_margin=0.3"
    + "&left_margin=0.3"
    + "&right_margin=0.3"
    + "&gridlines=false"
    + "&printtitle=false"
    + "&sheetnames=false"
    + "&pagenum=false"
    + "&attachment=false";

  return sanitizeForClient({
    success: true,
    srRow: srRow,
    printUrl: exportUrl
  });
}
// ============================================================
// MODULE: EDIT / UPDATE EXISTING BLEND
// ============================================================

/**
 * Returns the full blend data for a given Blend_ID, formatted exactly
 * like getMasterVarieties() output (with .bags, .rate, .shortage, .chamach
 * pre-filled from the saved blend), so the New Blend table can be
 * pre-populated for editing.
 */
function getBlendForEdit(blendId) {
  var ss = getSpreadsheet();
  var logSheet = ss.getSheetByName("BLEND_LOG");
  var data = logSheet.getDataRange().getValues();

  var blendMeta = null;
  var itemsByKey = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] !== blendId) continue;
    if (!blendMeta) {
      // Use col V (creation date) for the date picker, so editing keeps original date
      var creationDate = (row.length > 21 && row[21]) ? row[21] : row[1];
      blendMeta = {
        blendId: row[0],
        date: creationDate,
        category: row[2],
        blendType: row[3],
        customer: row[4],
        batchNo: row[5],
        basePrice: row[6],
        sheetUrl: row[18],
        status: row[19] || "Draft"
      };
    }
    var key = (row[7] + "|" + row[8] + "|" + row[9]).toUpperCase();
    itemsByKey[key] = {
      itemCategory: row[2],
      particulars: row[7], grade: row[8], invoice: row[9],
      bags: row[10], kgBag: row[11], shortage: row[13], rate: row[14], chamach: row[16],
      finalizedBags: row[20]
    };
  }

  if (!blendMeta) return null;

  var varieties = getMasterVarieties(blendMeta.category);
  var matchedKeys = {};

  varieties.forEach(function(v) {
    if (v.isDivider) return;
    var key = (v.particulars + "|" + v.grade + "|" + v.invoice).toUpperCase();
    var saved = itemsByKey[key];
    if (saved) {
      matchedKeys[key] = true;
      v.savedBags = saved.bags;
      v.savedRate = saved.rate;
      v.savedShortage = saved.shortage;
      v.savedChamach = saved.chamach;
      v.savedFinalizedBags = saved.finalizedBags;
      v.itemCategory = saved.itemCategory || blendMeta.category;
      if (blendMeta.status === "Finalized") {
        v.remaining = (parseFloat(v.remaining) || 0) + (parseFloat(saved.bags) || 0);
      }
    } else {
      v.savedBags = 0;
      v.savedFinalizedBags = 0;
      v.itemCategory = blendMeta.category;
    }
  });

  Object.keys(itemsByKey).forEach(function(key) {
    if (matchedKeys[key]) return;
    var saved = itemsByKey[key];
    varieties.push({
      row: null,
      particulars: saved.particulars,
      grade: saved.grade,
      invoice: saved.invoice,
      kgBag: saved.kgBag,
      rate: saved.rate,
      chamach: saved.chamach,
      remaining: (function() {
        var crossCat = saved.itemCategory;
        if (!crossCat) return 0;
        var crossMap = buildTrackerRemainingMap(crossCat);
        var crossKey = (saved.particulars + '|' + saved.grade + '|' + saved.invoice).toUpperCase();
        var r = crossMap[crossKey] !== undefined ? crossMap[crossKey] : 0;
        if (blendMeta.status === 'Finalized') r += parseFloat(saved.bags) || 0;
        return r;
      })(),
      isAnyBlend: false,
      isFixed: false,
      isCrossCategory: true,
      itemCategory: saved.itemCategory,
      savedBags: saved.bags,
      savedRate: saved.rate,
      savedShortage: saved.shortage,
      savedChamach: saved.chamach,
      savedFinalizedBags: saved.finalizedBags
    });
  });

  return sanitizeForClient({ meta: blendMeta, varieties: varieties });
}


/**
 * Updates an existing blend: overwrites the same per-blend Google Sheet
 * in place, removes old BLEND_LOG rows for this Blend_ID and re-inserts
 * updated rows, then refreshes the stock tracker.
 *
 * blendInfo: same shape as createBlend(), plus blendId (existing) and
 * sheetUrl (existing, to locate the file).
 */
function updateBlend(blendId, blendInfo) {
  var status = blendInfo.status || "Draft";

  // LOCK RULE: once a blend is Finalized it is frozen - no edits allowed.
  // Check the blend's CURRENT stored status in BLEND_LOG; if it's already
  // Finalized, reject the edit. (Finalizing a Draft is allowed; editing a
  // Draft is allowed; editing a Finalized blend is not.)
  var currentStatus = getBlendStoredStatus(blendId);
  if (currentStatus === "Finalized") {
    return sanitizeForClient({
      success: false,
      locked: true,
      message: "This blend is finalized and can no longer be edited."
    });
  }

  // 1. Stock guard - only meaningful for Finalized.
  //    (getBlendForEdit already adjusted v.remaining to add back this blend's
  //     previous finalized consumption, so the check is accurate)
  var warnings = [];
  if (status === "Finalized") {
    warnings = checkStockAvailability(blendInfo.items.map(function(it) {
      return {
        particulars: it.particulars, grade: it.grade, invoice: it.invoice,
        bagsRequested: parseFloat(it.bags) || 0,
        itemCategory: it.itemCategory || blendInfo.category
      };
    }));

    var blockingWarnings = warnings.filter(function(w) { return !w.isZeroWarning; });
    if (blockingWarnings.length > 0 && !blendInfo.overrideWarnings) {
      return sanitizeForClient({ success: false, warnings: warnings });
    }
  }

  // 2. Overwrite the existing per-blend sheet in place
 try {
  var oldIdMatch = blendInfo.sheetUrl && blendInfo.sheetUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (oldIdMatch) DriveApp.getFileById(oldIdMatch[1]).setTrashed(true);
} catch(e) { Logger.log('Could not trash old sheet: ' + e); }
var newSheetUrl = createBlendSheet(blendInfo, blendInfo.batchNo, blendId, status);
blendInfo.sheetUrl = newSheetUrl;

  // 3. Remove old BLEND_LOG rows for this Blend_ID, insert updated rows
  removeBlendLogRows(blendId);
  logBlendToLog(blendInfo, blendInfo.batchNo, blendId, blendInfo.sheetUrl, status);

  // 4. Refresh stock tracker (recalculates from scratch - changes in bags
  //    automatically reflected in remaining stock, only for Finalized rows)
 var refreshResult = refreshStockTracker();

   return sanitizeForClient({ success: true, batchNo: blendInfo.batchNo, sheetUrl: blendInfo.sheetUrl, blendId: blendId, status: status, warnings: warnings, zeroStock: refreshResult.zeroStock });
}


/**
 * Returns the stored Status of a blend ("Draft" / "Finalized") by reading
 * the first matching BLEND_LOG row for this Blend_ID. Returns null if not found.
 */
function getBlendStoredStatus(blendId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === blendId) {
      return data[i][19] || "Draft"; // Status column (T)
    }
  }
  return null;
}


/**
 * DELETE BLEND (same-day undo). Per spec:
 *  - Allowed only if the blend's date is TODAY.
 *  - Works on Draft and Finalized blends.
 *  - Removes the blend's BLEND_LOG rows + deletes (trashes) the blend's Sheet file.
 *  - Refreshes the stock tracker so any bags the blend consumed are restored
 *    (stock recalculates from BLEND_LOG; with the rows gone, consumption drops).
 * Returns { success, restored } or { success:false, message }.
 */
function deleteBlend(blendId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var data = sheet.getDataRange().getValues();

  // Find the blend's date + sheet URL from its first row
  var blendDate = null, sheetUrl = null, found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === blendId) {
      found = true;
      blendDate = data[i][1];       // Date (col B)
      sheetUrl = data[i][18];       // Sheet_URL (col S)
      break;
    }
  }
  if (!found) {
    return sanitizeForClient({ success: false, message: "Blend not found." });
  }

  // SAME-DAY RULE: the blend's date must equal today (date-only comparison)
  var d = blendDate instanceof Date ? blendDate : parseDateLocal(blendDate);
  var todayMs = dateOnlyMs(new Date());
  if (dateOnlyMs(d) !== todayMs) {
    return sanitizeForClient({
      success: false,
      message: "This blend can only be deleted on the day it was made."
    });
  }
// Restore any variety that was zero-stock-removed specifically because of this blend
   restoreVarietiesForBlend(blendId);
  // 1. Delete the blend's Google Sheet file (move to trash)
  if (sheetUrl) {
    try {
      var idMatch = String(sheetUrl).match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (idMatch) {
        DriveApp.getFileById(idMatch[1]).setTrashed(true);
      }
    } catch (e) {
      // file may already be gone; continue with log removal regardless
    }
  }

  // 2. Remove the blend's BLEND_LOG rows
  removeBlendLogRows(blendId);

  // 3. Refresh stock - bags the blend consumed are now restored automatically
  refreshStockTracker();

  return sanitizeForClient({ success: true });
}


/**
 * Removes all BLEND_LOG rows matching the given Blend_ID.
 */
function removeBlendLogRows(blendId) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var data = sheet.getDataRange().getValues();

  // Find row indices (1-indexed sheet rows) to delete, from bottom to top
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === blendId) {
      rowsToDelete.push(i + 1); // sheet row number
    }
  }
  // delete from bottom to top to avoid index shifting issues
  rowsToDelete.sort(function(a,b){ return b - a; });
  rowsToDelete.forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });
}


/**
 * Opens the existing per-blend sheet (by URL) and overwrites the bag,
 * rate, shortage, chamach, date, customer, base price values in place -
 * formulas and formatting remain untouched since we only touch input cells.
 */
function updateBlendSheetInPlace(blendInfo, blendId, status) {
  var ss = getSpreadsheet();
  var url = blendInfo.sheetUrl;
  var idMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!idMatch) throw new Error("Could not parse sheet ID from URL: " + url);

  var targetSs = SpreadsheetApp.openById(idMatch[1]);
  var targetSheet = targetSs.getSheets()[0]; // single-sheet blend file

  // Remove any cross-category (OF) rows written by a previous save, so we can
  // re-write them cleanly from the current item list. This keeps edits correct
  // even when OF teas are added/removed between saves.
  removeCrossCategoryRows(targetSheet);

var masterSheet = ss.getSheetByName(blendInfo.category + '_MASTER');
var clearLayout = detectLayout(masterSheet);
var clearRows = [].concat(clearLayout.varietyRows, clearLayout.anyBlendRows);
if (clearLayout.leafRow) clearRows.push(clearLayout.leafRow);
if (clearLayout.looseSampleRow) clearRows.push(clearLayout.looseSampleRow);
clearRows.forEach(function(rn) {
  targetSheet.getRange(rn, COL.NO_BAGS).setValue(0);
});

  // Overwrite bag/rate/shortage/chamach for each item
  var rowLookupU = buildVarietyRowLookup(targetSheet);
  blendInfo.items.forEach(function(item) {
    if (item.isCrossCategory) return;

   if (!item.particulars) {
  if (!item.row || typeof item.row !== 'number') return;
  targetSheet.getRange(item.row, COL.NO_BAGS).setValue(parseFloat(item.bags) || 0);
targetSheet.getRange(item.row, COL.NO_BAGS).setNumberFormat("0.00");
  return;
}

    var key = String(item.particulars).trim().toUpperCase() + '|' +
              String(item.grade || '').trim().toUpperCase() + '|' +
              String(item.invoice || '').trim().toUpperCase();
    var rn = rowLookupU[key] || (typeof item.row === 'number' ? item.row : null);
    if (!rn) return;

    var bags = parseFloat(item.bags) || 0;
    targetSheet.getRange(rn, COL.NO_BAGS).setValue(bags);
    if (item.rate !== undefined && item.rate !== '') targetSheet.getRange(rn, COL.RATE).setValue(parseFloat(item.rate));
    if (item.shortage !== undefined && item.shortage !== '') targetSheet.getRange(rn, COL.SHORTAGE).setValue(parseFloat(item.shortage));
    if (item.chamach !== undefined && item.chamach !== '') targetSheet.getRange(rn, COL.CHAMACH).setValue(parseFloat(item.chamach));
    var finalizedBags = (item.finalizedBags !== undefined && item.finalizedBags !== null && item.finalizedBags !== '')
      ? parseFloat(item.finalizedBags) : (status === 'Finalized' ? bags : 0);
    targetSheet.getRange(rn, COL.OPEN).setValue(status === 'Finalized' ? (bags - finalizedBags) : 0);
  });
  // Re-write cross-category OF rows for the current item list.
  writeCrossCategoryRows(targetSheet, blendInfo, status);

  // Update header info
  targetSheet.getRange(4, 4).setValue(blendInfo.batchNo);
  targetSheet.getRange(4, 5).setValue(parseDateLocal(blendInfo.date));
  targetSheet.getRange(2, 8).setValue(blendInfo.basePrice);
  if (blendInfo.customer) {
    targetSheet.getRange(2, 3).setValue(blendInfo.customer);
  }
}

/**
 * Removes rows previously written as cross-category OF rows. These are the rows
 * sitting between the black "OF" divider band and the next divider/ANY BLEND
 * section that contain a real Particulars value. (The single original blank row
 * after the OF band has no Particulars, so it is left untouched.)
 *
 * We identify OF-section data rows as: rows after the OF band, up to the next
 * black divider band or the first ANY BLEND row, that have a non-empty
 * Particulars cell. Deleting bottom-up keeps row indices stable.
 */
function removeCrossCategoryRows(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), COL.T_WT);
  var scan = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  // Find OF band row.
  var ofBandRow = null;
  for (var r = 0; r < scan.length; r++) {
    var rowNum = r + 1;
    if (rowNum < 6) continue;
    for (var c = 0; c < Math.min(scan[r].length, COL.GRADE); c++) {
      var txt = String(scan[r][c] || "").trim().toUpperCase();
      if (txt === "OF") { ofBandRow = rowNum; break; }
    }
    if (ofBandRow) break;
  }
  if (!ofBandRow) return;

  // Walk rows after the OF band. Stop at the first ANY BLEND row or a black
  // divider-like row (a short label in col A-E with no particulars that isn't
  // simply an empty spacer). Collect rows that have a Particulars value -> these
  // are OF data rows to remove.
  var rowsToDelete = [];
  for (var rr = ofBandRow + 1; rr <= scan.length; rr++) {
    var rowArr = scan[rr - 1];
    if (!rowArr) break;
    var particulars = String(rowArr[COL.PARTICULARS - 1] || "").trim();
    var pUpper = particulars.toUpperCase();

    // Stop conditions: reached the ANY BLEND section or the Total row.
    if (pUpper.indexOf("ANY BLEND") === 0 || pUpper.indexOf("ANY ") === 0 ||
        pUpper.indexOf("MIX") === 0 || pUpper.indexOf("TOTAL") === 0) {
      break;
    }
    // A black divider band (e.g. another "OF"/section) - stop.
    var colA = String(rowArr[COL.MD - 1] || "").trim().toUpperCase();
    if (!particulars && (colA === "OF" || colA.length > 0 && colA.length <= 3)) {
      // could be a spacer/divider; stop to be safe.
      break;
    }

    // An OF data row we previously wrote: has a real Particulars value.
    if (particulars) {
      rowsToDelete.push(rr);
    }
  }

  rowsToDelete.sort(function(a, b) { return b - a; }); // bottom-up
  rowsToDelete.forEach(function(rn) { sheet.deleteRow(rn); });
}


/**
 * Convenience wrapper: marks a blend as Finalized.
 * If the blend is currently a Draft, this is its first finalization -
 * the current bags become the Finalized_Bags baseline (Open/K = 0).
 * If it's already Finalized, this is a no-op status-wise (use updateBlend
 * for subsequent edits/top-ups instead).
 * blendInfo must include the current items/meta (same shape as updateBlend).
 */
function finalizeBlend(blendId, blendInfo) {
  blendInfo.status = "Finalized";
  // On first finalize, clear any stale finalizedBags so baseline = current bags
  blendInfo.items.forEach(function(item) {
    if (item.finalizedBags === undefined || item.finalizedBags === null || item.finalizedBags === "") {
      item.finalizedBags = item.bags;
    }
  });
  return updateBlend(blendId, blendInfo);
}


// ============================================================
// ONE-TIME MIGRATION: add Status + Finalized_Bags columns to
// an existing BLEND_LOG sheet (run this once if you set up the
// spreadsheet before this update).
// ============================================================
function migrateAddStatusColumns() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  if (headerRow.indexOf("Status") === -1) {
    sheet.getRange(1, lastCol + 1).setValue("Status");
    sheet.getRange(1, lastCol + 2).setValue("Finalized_Bags");
    Logger.log("Added Status and Finalized_Bags columns to BLEND_LOG.");
  } else {
    Logger.log("BLEND_LOG already has Status column - no migration needed.");
  }
}


// ============================================================
// ZERO-STOCK REMOVAL (spec Section 4, with user confirmation)
// ============================================================

function getOrCreateRemovedVarietiesSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("REMOVED_VARIETIES");
  if (!sheet) {
    sheet = ss.insertSheet("REMOVED_VARIETIES");
    sheet.getRange(1, 1, 1, 12).setValues([[
      "Removed_At", "Trigger_Blend_Id", "Category", "Particulars", "Grade", "Invoice",
      "PlaceOrLNo", "KgBag", "Rate", "Chamach", "OpeningBags", "StockDate"
    ]]);
  }
  return sheet;
}

/**
 * If deleting blendId previously caused any variety to be zero-stock-removed,
 * restore that variety back into its master sheet using the saved snapshot.
 */
function restoreVarietiesForBlend(blendId) {
  var removedSheet = getOrCreateRemovedVarietiesSheet();
  var data = removedSheet.getDataRange().getValues();
  var ss = getSpreadsheet();
  var rowsToDeleteFromLog = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[1] !== blendId) continue; // Trigger_Blend_Id

    var category = row[2];
    var masterSheet = ss.getSheetByName(category + "_MASTER");
    var layout = detectLayout(masterSheet);

    var insertBeforeRow;
    if (layout.lastVarietyRow) {
      insertBeforeRow = layout.lastVarietyRow + 1;
    } else {
      var candidates = [].concat(layout.dividerRows, layout.anyBlendRows,
        [layout.looseSampleRow, layout.leafRow, layout.totalRow]).filter(function(x){ return x; });
      insertBeforeRow = Math.min.apply(null, candidates);
    }

    masterSheet.insertRowBefore(insertBeforeRow);
    var newRow = insertBeforeRow;
    var newLayout = detectLayout(masterSheet);
    var totalRow = newLayout.totalRow;

    masterSheet.getRange(newRow, COL.MD).setValue(row[11] ? new Date(row[11]) : new Date());
    masterSheet.getRange(newRow, COL.BAGS).setValue(parseFloat(row[10]) || 0);
    masterSheet.getRange(newRow, COL.PLACE).setValue(row[6] || "");
    masterSheet.getRange(newRow, COL.PARTICULARS).setValue(row[3]);
    masterSheet.getRange(newRow, COL.GRADE).setValue(row[4]);
    masterSheet.getRange(newRow, COL.INV).setValue(row[5]);
    masterSheet.getRange(newRow, COL.KG_BAG).setValue(parseFloat(row[7]) || 0);
    masterSheet.getRange(newRow, COL.NO_BAGS).setValue(0);
    masterSheet.getRange(newRow, COL.RATE).setValue(parseFloat(row[8]) || 0);
    masterSheet.getRange(newRow, COL.CHAMACH).setValue(parseFloat(row[9]) || 0);

    masterSheet.getRange(newRow, COL.PCT).setFormula("=(J" + newRow + "/$J$" + totalRow + ")");
    masterSheet.getRange(newRow, COL.KG).setFormula("=G" + newRow + "*H" + newRow);
    masterSheet.getRange(newRow, COL.AMOUNT).setFormula(
      category === "BP"
        ? "=(J" + newRow + "+L" + newRow + ")*M" + newRow
        : "=J" + newRow + "*M" + newRow
    );
    masterSheet.getRange(newRow, COL.CHAMCH_NAAP).setFormula("=O" + newRow + "*18.5");
    masterSheet.getRange(newRow, COL.T_WT).setFormula("=J" + newRow + "*O" + newRow);

    if (newRow > 6) copyRowFormatting(masterSheet, newRow - 1, newRow, 17);

    rowsToDeleteFromLog.push(i + 1);
  }

  rowsToDeleteFromLog.sort(function(a, b) { return b - a; });
  rowsToDeleteFromLog.forEach(function(rn) { removedSheet.deleteRow(rn); });
}

/**
 * Deletes the given zero-stock variety rows from their master sheets.
 * Called by the app AFTER the user confirms the removal prompt.
 * zeroStockList: [{category, particulars, grade, invoice}] (row numbers are
 * re-resolved here by matching, because row numbers may have shifted).
 */
function removeZeroStockVarieties(zeroStockList, triggerBlendId) {
  var ss = getSpreadsheet();
  var removed = 0;
  var removedSheet = getOrCreateRemovedVarietiesSheet();
  var snapshotRows = [];

  ["BP", "DUST"].forEach(function(cat) {
    var sheet = ss.getSheetByName(cat + "_MASTER");
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(1, 1, lastRow, 17).getValues();

    var targets = zeroStockList.filter(function(z) { return z.category === cat; });
    if (targets.length === 0) return;

    var rowsToDelete = [];
    for (var i = 0; i < data.length; i++) {
      var rowNum = i + 1;
      if (rowNum < 6) continue;
      var p = String(data[i][COL.PARTICULARS - 1] || '').trim();
      var g = String(data[i][COL.GRADE - 1] || '').trim();
      var inv = String(data[i][COL.INV - 1] || '').trim();
      var key = (p + "|" + g + "|" + inv).toUpperCase();
      var match = targets.some(function(t) {
        return (String(t.particulars || '').trim() + "|" +
                String(t.grade || '').trim() + "|" +
                String(t.invoice || '').trim()).toUpperCase() === key;
      });
      if (match) {
        rowsToDelete.push(rowNum);
        snapshotRows.push([
          new Date(), triggerBlendId || "", cat,
          p, g, inv,
          data[i][COL.PLACE - 1],
          data[i][COL.KG_BAG - 1],
          data[i][COL.RATE - 1],
          data[i][COL.CHAMACH - 1],
          data[i][COL.BAGS - 1],
          data[i][COL.MD - 1]
        ]);
      }
    }

    rowsToDelete.sort(function(a, b) { return b - a; });
    rowsToDelete.forEach(function(rn) { sheet.deleteRow(rn); removed++; });
  });

  if (snapshotRows.length > 0) {
    var startRow = removedSheet.getLastRow() + 1;
    removedSheet.getRange(startRow, 1, snapshotRows.length, snapshotRows[0].length).setValues(snapshotRows);
  }

  refreshStockTracker();
  return sanitizeForClient({ success: true, removed: removed });
}
// ============================================================
// AUTO STOCK_DATE UPDATE (spec Section 6)
// ============================================================

/**
 * Installable onEdit trigger: when the BAGS column (col B) of a master sheet
 * is edited, set that row's M/D (stock_date, col A) to today.
 *
 * SETUP (one time): in Apps Script, run installOnEditTrigger() once to register
 * this. (A simple onEdit won't fire reliably for all cases, so we use an
 * installable trigger.)
 */
function onEditStockDate(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  if (name !== "BP_MASTER" && name !== "DUST_MASTER") return;

  var col = e.range.getColumn();
  var rowNum = e.range.getRow();
  if (col !== COL.BAGS) return;       // only react to BAGS column edits
  if (rowNum < 6) return;             // skip header area

  // Write today's date into the M/D column (col A) of the same row
  sheet.getRange(rowNum, COL.MD).setValue(new Date());
}

/**
 * Run this ONCE to install the onEdit trigger for stock_date auto-update.
 */
function installOnEditTrigger() {
  var ssId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  // Remove any existing duplicate triggers for this handler first
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "onEditStockDate") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onEditStockDate")
    .forSpreadsheet(ssId)
    .onEdit()
    .create();
  Logger.log("onEdit trigger installed for stock_date auto-update.");
}


// ============================================================
// MODULE: VARIETY DETAILS — edit, delete, sort
// ============================================================

function migrateAddCreationDateColumn() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headerRow.indexOf("Creation_Date") === -1) {
    sheet.getRange(1, lastCol + 1).setValue("Creation_Date");
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var existingDates = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      sheet.getRange(2, lastCol + 1, lastRow - 1, 1).setValues(existingDates);
    }
    Logger.log("Added Creation_Date column.");
  } else {
    Logger.log("Already exists.");
  }
}

function getVarietyMasterDetails(category, particulars, grade, invoice) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(category + "_MASTER");
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(1, 1, lastRow, 17).getValues();
  var pU = String(particulars || '').trim().toUpperCase();
  var gU = String(grade || '').trim().toUpperCase();
  var invU = String(invoice || '').trim().toUpperCase();
  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 1;
    if (rowNum < 6) continue;
    var p = String(data[i][COL.PARTICULARS - 1] || '').trim().toUpperCase();
    var g = String(data[i][COL.GRADE - 1] || '').trim().toUpperCase();
    var inv = String(data[i][COL.INV - 1] || '').trim().toUpperCase();
    if (p === pU && g === gU && inv === invU) {
      return sanitizeForClient({
        row: rowNum,
        rate: data[i][COL.RATE - 1],
        chamach: data[i][COL.CHAMACH - 1],
        place: data[i][COL.PLACE - 1],
        invoice: data[i][COL.INV - 1]
      });
    }
  }
  return null;
}

function updateVarietyDetails(category, row, details) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(category + "_MASTER");
  if (details.rate !== undefined && details.rate !== "") sheet.getRange(row, COL.RATE).setValue(parseFloat(details.rate));
  if (details.chamach !== undefined && details.chamach !== "") sheet.getRange(row, COL.CHAMACH).setValue(parseFloat(details.chamach));
  if (details.place !== undefined) sheet.getRange(row, COL.PLACE).setValue(details.place);
  if (details.invoice !== undefined && details.invoice !== "") sheet.getRange(row, COL.INV).setValue(details.invoice);
  if (details.particulars !== undefined && details.particulars !== "") sheet.getRange(row, COL.PARTICULARS).setValue(details.particulars);
  if (details.grade !== undefined && details.grade !== "") sheet.getRange(row, COL.GRADE).setValue(details.grade);  
  return sanitizeForClient({ success: true });
}

function deleteVarietyFromMaster(category, row) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(category + "_MASTER");
  sheet.deleteRow(row);
  refreshStockTracker();
  return sanitizeForClient({ success: true });
}

function restoreMasterFormulas() {
  var ss = getSpreadsheet();
  ["BP", "DUST"].forEach(function(cat) {
    var sheet = ss.getSheetByName(cat + "_MASTER");
    var layout = detectLayout(sheet);
    var totalRow = layout.totalRow;
    if (!totalRow) { Logger.log("No total row found for " + cat); return; }
    Logger.log(cat + " totalRow=" + totalRow + " varietyRows=" + layout.varietyRows.length);

    var allRows = [].concat(
      layout.varietyRows,
      layout.anyBlendRows,
      layout.leafRow ? [layout.leafRow] : [],
      layout.looseSampleRow ? [layout.looseSampleRow] : []
    );

    allRows.forEach(function(rn) {
      var isAnyBlend = layout.anyBlendRows.indexOf(rn) !== -1;
      var isFixed = (rn === layout.leafRow || rn === layout.looseSampleRow);

      // % formula
      sheet.getRange(rn, COL.PCT).setFormula(
        "=(J" + rn + "/$J$" + totalRow + ")"
      );

      // KG formula
      sheet.getRange(rn, COL.KG).setFormula("=G" + rn + "*H" + rn);

      // Amount formula
      if (isFixed) {
        sheet.getRange(rn, COL.AMOUNT).setFormula("=J" + rn + "*M" + rn);
      } else if (cat === "BP" && !isAnyBlend) {
        sheet.getRange(rn, COL.AMOUNT).setFormula("=(J" + rn + "+L" + rn + ")*M" + rn);
      } else {
        sheet.getRange(rn, COL.AMOUNT).setFormula("=J" + rn + "*M" + rn);
      }

      // Chamch Naap
      if (isAnyBlend || isFixed) {
        sheet.getRange(rn, COL.CHAMCH_NAAP).setFormula("=O" + rn + "/18.5");
      } else {
        sheet.getRange(rn, COL.CHAMCH_NAAP).setFormula("=O" + rn + "*18.5");
      }

      // T.Wt
      sheet.getRange(rn, COL.T_WT).setFormula("=J" + rn + "*O" + rn);
    });

    Logger.log(cat + " formulas restored for " + allRows.length + " rows.");

  if (totalRow && layout.firstVarietyRow) {
  var first = layout.firstVarietyRow;
  var last = totalRow - 1;
var leafStart = null;
if (layout.looseSampleRow) leafStart = layout.looseSampleRow;
if (layout.leafRow && (!leafStart || layout.leafRow < leafStart)) leafStart = layout.leafRow;
if (leafStart) {
  sheet.getRange(totalRow, COL.NO_BAGS).setFormula('=SUM(H' + first + ':H' + (leafStart - 1) + ')');
}
  sheet.getRange(totalRow, COL.KG).setFormula('=SUM(J' + first + ':J' + last + ')');
  sheet.getRange(totalRow, COL.OPEN).setFormula('=SUM(K' + first + ':K' + last + ')');
  sheet.getRange(totalRow, COL.SHORTAGE).setFormula('=SUM(L' + first + ':L' + last + ')');
  sheet.getRange(totalRow, COL.AMOUNT).setFormula('=SUM(N' + first + ':N' + last + ')');
  sheet.getRange(totalRow, COL.T_WT).setFormula('=SUM(Q' + first + ':Q' + last + ')');
}

  });
}
function sortMasterAlphabetically(category) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(category + "_MASTER");
  var layout = detectLayout(sheet);
  if (!layout.varietyRows || layout.varietyRows.length < 2) {
    return sanitizeForClient({ success: true, message: "Nothing to sort" });
  }

  // Read only the INPUT columns (not formula columns)
  // Input cols: A(MD), B(BAGS), C(PLACE), D(PARTICULARS), E(GRADE), F(INV), G(KG_BAG), H(NO_BAGS), L(SHORTAGE), M(RATE), O(CHAMACH)
  var inputCols = [COL.MD, COL.BAGS, COL.PLACE, COL.PARTICULARS, COL.GRADE, COL.INV, COL.KG_BAG, COL.NO_BAGS, COL.SHORTAGE, COL.RATE, COL.CHAMACH];

  var rowsData = layout.varietyRows.map(function(rn) {
    var vals = {};
    inputCols.forEach(function(c) {
      vals[c] = sheet.getRange(rn, c).getValue();
    });
    return {
      name: String(vals[COL.PARTICULARS] || '').trim().toUpperCase(),
      vals: vals
    };
  });

  rowsData.sort(function(a, b) { return a.name.localeCompare(b.name); });

  // Write back only input columns, leaving formula columns untouched
  rowsData.forEach(function(rd, idx) {
    var targetRow = layout.varietyRows[idx];
    inputCols.forEach(function(c) {
      sheet.getRange(targetRow, c).setValue(rd.vals[c]);
    });
  });

  refreshStockTracker();
  return sanitizeForClient({ success: true, message: "Sorted " + rowsData.length + " varieties" });
}

function fixLeafRowLabels() {
  var ss = getSpreadsheet();
  
  var bp = ss.getSheetByName("BP_MASTER");
  bp.getRange(32, 1).setValue("Loose Sample Tea");
  bp.getRange(33, 1).setValue("  950 KG BLEND  AND ABOVE FIX 7KG LEAF ");
  
  var dust = ss.getSheetByName("DUST_MASTER");
  dust.getRange(24, 1).setValue("Loose Sample Tea");
  dust.getRange(25, 1).setValue("  950 KG BLEND  AND ABOVE FIX 6KG LEAF ");
  
  Logger.log("Leaf row labels fixed.");
}

function sortBothMasters() {
  sortMasterAlphabetically("BP");
  sortMasterAlphabetically("DUST");
  Logger.log("Both masters sorted alphabetically.");
}


// ============================================================
// ONE-TIME: rebuild master sheets from latest Excel data with
// EXACT formatting (colors, fonts, sizes, alignment, borders,
// column widths, row heights). Run rebuildMasters() once.
// ============================================================
function rebuildMasters() {
  rebuildBPMaster();
  rebuildDUSTMaster();
  refreshStockTracker();
  Logger.log("Both masters rebuilt and tracker refreshed.");
}

function rebuildBPMaster() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BP_MASTER");
  if (!sheet) { sheet = ss.insertSheet("BP_MASTER"); }
  sheet.clear();
  var R = 35, C = 17;

  // Ensure enough rows/cols
  if (sheet.getMaxRows() < R) sheet.insertRowsAfter(sheet.getMaxRows(), R - sheet.getMaxRows());

  var values = [["STOCK \n09.06.2026","","Pushaan Blend Formula ","","BP","",""," DP","=Q4","","","","=N34/J34","=M1+5%*(M1)","=N1+30","","=N1+61"],["","","MASTER FILE DON’T CHANGE ","","","","",100,"YEAR","","Total Kg.","Katori","=Q34/J34","PAID ","WH","","DP2"],["","","","","","","","","","","","","","","","",""],["","","Batch No. ","NO BATCH ",new Date("2026-06-15T00:00:00"),"","","",2026,"","=J34","=Q34/J34","WT","","=H2-O1","=H2-P1","=H2-Q1"],["M/D","BAGS ","PLACE","Particulars ","BP/DUST","INV","Kg. Bag","No.  Bags ","%","KG","Open ","Shortage ","Rate ","AMOUNT  ","CHAMACH","Chamch Naap","T. Wt "],[new Date("2026-05-27T00:00:00"),2,16,"BHOLATAR","BP",645,30,0,"=(J6/$J$34)","=G6*H6","","",330,"=(J6+L6)*M6",3.6,"=O6*18.5","=J6*O6"],[new Date("2026-06-08T00:00:00"),14,32,"DUBORY ","BP",23,28,0,"=(J7/$J$34)","=G7*H7","","",212,"=(J7+L7)*M7",2.8,"=O7*18.5","=J7*O7"],[new Date("2026-06-15T00:00:00"),10,23,"FALLANGNI ","BP","49,60",35,0,"=(J8/$J$34)","=G8*H8","","",215,"=(J8+L8)*M8",2.8,"=O8*18.5","=J8*O8"],[new Date("2026-06-04T00:00:00"),9,23,"GOLAIJAAN ","BP1",15,30,"","=(J9/$J$34)","=G9*H9","","",190,"=(J9+L9)*M9",2.8,"=O9*18.5","=J9*O9"],[new Date("2026-06-08T00:00:00"),18,14,"HUKAJAN ","BP",150,36,"","=(J10/$J$34)","=G10*H10","","",215,"=(J10+L10)*M10",3,"=O10*18.5","=J10*O10"],[new Date("2026-06-01T00:00:00"),12,9,"KISHORIBAGH","BP",226,26,"","=(J11/$J$34)","=G11*H11","","",310,"=(J11+L11)*M11",3.3,"=O11*18.5","=J11*O11"],[new Date("2026-06-15T00:00:00"),20,"","KAMLA GOLD ","BP",375,26,"","=(J12/$J$34)","=G12*H12","","",290,"=(J12+L12)*M12",3.2,"=O12*18.5","=J12*O12"],[new Date("2026-06-15T00:00:00"),40,5,"LAADLI GOLD ","BP",341,25,"","=(J13/$J$34)","=G13*H13","","",232,"=(J13+L13)*M13",2.8,"=O13*18.5","=J13*O13"],[new Date("2026-06-04T00:00:00"),14,7,"NEEBO ","BP",214,37,"","=(J14/$J$34)","=G14*H14","","",200,"=(J14+L14)*M14",2.8,"=O14*18.5","=J14*O14"],[new Date("2026-06-04T00:00:00"),20,21,"PRISHNIBHA","BP1",167,35,"","=(J15/$J$34)","=G15*H15","","",230,"=(J15+L15)*M15",2.8,"=O15*18.5","=J15*O15"],[new Date("2026-05-08T00:00:00"),1,1,"POMI","BP",129,34,0,"=(J16/$J$34)","=G16*H16","","",230,"=(J16+L16)*M16",2.8,"=O16*18.5","=J16*O16"],[new Date("2026-06-04T00:00:00"),8,2,"PURI PREMIUM ","BP",55,38,"","=(J17/$J$34)","=G17*H17","","",315,"=(J17+L17)*M17",3.2,"=O17*18.5","=J17*O17"],[new Date("2026-06-15T00:00:00"),10,"K","RADHAJULI ","BOPSM",77,35,"","=(J18/$J$34)","=G18*H18","","",195,"=(J18+L18)*M18",2.9,"=O18*18.5","=J18*O18"],[new Date("2026-06-12T00:00:00"),20,"21A","RAJAPOOKRI ","BP",17,35,"","=(J19/$J$34)","=G19*H19","","",227,"=(J19+L19)*M19",3.15,"=O19*18.5","=J19*O19"],[new Date("2026-05-08T00:00:00"),4,8,"RANGDDHALI","BP",14,35,0,"=(J20/$J$34)","=G20*H20","","",235,"=(J20+L20)*M20",3,"=O20*18.5","=J20*O20"],[new Date("2026-05-25T00:00:00"),5,"13,14","RUGBY ","BP",122131,35,"","=(J21/$J$34)","=G21*H21","","",205,"=(J21+L21)*M21",2.75,"=O21*18.5","=J21*O21"],[new Date("2026-06-15T00:00:00"),24,"","SALBONI","BP","",38,"","=(J22/$J$34)","=G22*H22","","",190,"=(J22+L22)*M22",2.75,"=O22*18.5","=J22*O22"],[new Date("2026-06-08T00:00:00"),9,30,"TORSANA ","BP",672,26,"","=(J23/$J$34)","=G23*H23","","",220,"=(J23+L23)*M23",3.3,"=O23*18.5","=J23*O23"],["","","OF","","","","",0," ","","","","","",0,"=O24*18.5","=J24*O24"],[new Date("2026-05-29T00:00:00"),4,3,"MALIBELIA ","PF1",419,38,0,"=(J25/$J$34)","=G25*H25","","",228,"=(J25+L25)*M25",3,"=O25/18.5","=J25*O25"],["","","","","","","",0,"","","","","","","","",""],["","","","ANY BLEND DUP","BP","E-88",30,0,"=(J27/$J$34)","=G27*H27","","",276,"=(J27+L27)*M27",3,"=O27/18.5","=J27*O27"],["","","","ANY BLEND RR ","BP","F-26",30,0,"=(J28/$J$34)","=G28*H28","","",236,"=(J28+L28)*M28",3,"=O28/18.5","=J28*O28"],["","","","ANY BLEND SHRI","BP","F-28",30,0,"=(J29/$J$34)","=G29*H29","","",202,"=(J29+L29)*M29",2.9,"=O29/18.5","=J29*O29"],["","","","ANY BLEND JODHPUR","BP","E-18",30,0,"=(J30/$J$34)","=G30*H30","","",178,"=(J30+L30)*M30",2.85,"=O30/18.5","=J30*O30"],["","","","MIX  BLEND ","MIX ","",30,0,"=(J31/$J$34)","=G31*H31","","",180,"=(J31+L31)*M31",3,"=O31/18.5","=J31*O31"],["Loose Sample Tea","","","","","",7,"=H33","=(J32/$J$34)","=G32*H32","","",220,"=J32*M32",3,"=O32/18.5","=J32*O32"],["  950 KG BLEND  AND ABOVE FIX 7KG LEAF ","","","","","",7,0.01,"=(J33/$J$34)","=G33*H33",0,"",330,"=J33*M33",2,"=O33/18.5","=J33*O33"],["Total ","","","","","","","=SUM(H6:H31)","=J34/$J$34","=SUM(J6:J33)","=SUM(K6:K33)","=SUM(L6:L33)","","=SUM(N6:N33)","","","=SUM(Q6:Q33)"],["BLENDING AREA          DATE ","","","BLENDAR       IN              DATE "," BLENDAR    OUT              DATE ","","","TESTING","","C","T","","","","=J34/30","",""]];
  sheet.getRange(1,1,R,C).setValues(values);

  var rng = sheet.getRange(1,1,R,C);
  rng.setBackgrounds([["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#3C3B38","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#A5B6CB","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#A5B6CB","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#A5B6CB","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#000000","#000000","#000000","#000000","#000000","#000000","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#000000","#D9D9D9","#D9D9D9","#D9D9D9"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#FFFFFF","#D9D9D9"],["#FFFFFF","#000000","#000000","#000000","#000000","#000000","#000000","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#FFFFFF","#D9D9D9"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#FFFFFF","#D9D9D9"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#FFFFFF","#D9D9D9"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#FFFFFF","#D9D9D9"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#FFFFFF","#D9D9D9"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#D9D9D9","#D9D9D9","#FFFFFF","#D9D9D9"],["#92D050","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#92D050","#FFFF00","#FFFF00","#FFFF00","#92D050","#92D050","#FFFFFF","#FFFF00","#D9D9D9","#FFFFFF","#D9D9D9"],["#92D050","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#92D050","#FFFFFF","#FFFF00","#FFFF00","#92D050","#92D050","#92D050","#FFFF00","#D9D9D9","#FFFFFF","#D9D9D9"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFF00","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#000000","#FFFFFF","#000000","#000000","#000000","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"]]);
  rng.setFontColors([["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#BFBFBF","#BFBFBF","#C00000","#BFBFBF","#BFBFBF"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FF0000","#FF0000","#000000","#FF0000","#FF0000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FF0000","#FF0000","#000000","#000000","#000000","#000000","#000000"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#FFFFFF","#FFFFFF","#FFFFFF","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FF0000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FFFF00","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FF0000","#000000","#FF0000","#FF0000","#FF0000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FFFFFF","#000000","#FFFFFF","#FFFFFF","#FFFFFF","#000000","#000000","#000000","#000000","#000000"]]);
  rng.setFontWeights([["bold","normal","bold","normal","bold","normal","normal","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["normal","normal","bold","normal","normal","normal","normal","bold","bold","normal","bold","bold","bold","bold","bold","bold","bold"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","bold","bold","bold","normal","normal","normal","bold","normal","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","normal","normal","normal","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","normal","normal","normal","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","normal","normal","normal","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","bold","bold","normal","normal","bold","normal","bold","bold","bold","bold","bold","bold","normal","normal"]]);
  rng.setFontStyles([["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["italic","italic","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"]]);
  rng.setFontSizes([[16.0,11.0,16.0,11.0,20.0,11.0,11.0,22.0,12.0,16.0,16.0,14.0,16.0,14.0,14.0,14.0,14.0],[11.0,11.0,26.0,11.0,11.0,11.0,11.0,18.0,16.0,11.0,16.0,16.0,16.0,16.0,16.0,16.0,16.0],[11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,11.0,20.0,26.0,36.0,11.0,11.0,11.0,14.0,11.0,36.0,28.0,16.0,14.0,14.0,14.0,14.0],[18.0,16.0,20.0,16.0,26.0,18.0,18.0,18.0,14.0,14.0,14.0,14.0,16.0,12.0,14.0,11.0,14.0],[18.0,16.0,20.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,17.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,16.0,20.0,18.0,26.0,18.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,22.0,12.0,20.0,11.0,11.0,11.0,11.0,11.0,16.0,11.0],[16.0,16.0,16.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,16.0,16.0,16.0,11.0,11.0,16.0,11.0],[16.0,16.0,16.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,16.0,16.0,16.0,11.0,11.0,16.0,11.0],[16.0,16.0,16.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,16.0,16.0,16.0,11.0,11.0,16.0,11.0],[16.0,16.0,16.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,16.0,16.0,16.0,11.0,11.0,16.0,11.0],[16.0,16.0,16.0,18.0,26.0,16.0,18.0,22.0,12.0,20.0,16.0,16.0,16.0,11.0,11.0,16.0,11.0],[16.0,16.0,16.0,16.0,26.0,16.0,18.0,22.0,16.0,20.0,16.0,16.0,16.0,11.0,11.0,16.0,11.0],[12.0,11.0,11.0,11.0,11.0,11.0,18.0,22.0,12.0,20.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[18.0,11.0,11.0,11.0,11.0,11.0,18.0,22.0,12.0,50.0,12.0,12.0,16.0,11.0,11.0,11.0,11.0],[16.0,11.0,11.0,11.0,11.0,11.0,18.0,22.0,11.0,20.0,16.0,16.0,16.0,11.0,11.0,11.0,11.0],[16.0,11.0,11.0,16.0,16.0,11.0,11.0,18.0,11.0,14.0,14.0,11.0,16.0,11.0,11.0,11.0,11.0]]);
  rng.setHorizontalAlignments([["center","left","center","left","center","left","left","center","left","left","left","center","center","left","center","center","left"],["left","left","center","left","left","left","left","center","center","left","center","center","center","center","center","center","center"],["left","left","left","left","left","left","left","left","left","left","left","left","left","left","left","left","left"],["left","left","center","center","center","left","left","left","center","left","left","left","center","left","center","center","center"],["left","center","center","center","center","center","center","center","center","center","center","center","center","center","center","center","center"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","left","left","left","center","center","center","left","left","left","left","center","left","center","center","left"],["center","center","center","left","center","center","center","center","left","left","left","left","left","left","center","center","left"],["left","center","center","left","center","center","center","center","left","left","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","center","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","center","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","center","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","center","left","left","center","left","center","center","left"],["left","center","center","left","center","center","center","center","left","center","left","left","center","left","center","center","left"],["center","left","left","left","left","left","center","center","left","left","left","left","center","left","center","center","left"],["center","left","left","left","left","left","center","center","left","left","left","left","center","left","center","center","left"],["center","left","left","left","left","left","center","left","left","left","center","center","center","center","center","center","center"],["center","left","left","center","center","left","left","center","left","center","center","center","center","center","center","left","left"]]);
  rng.setVerticalAlignment("middle");
  rng.setBorder(true, true, true, true, true, true);

  // Column widths (px)
  var widths = [91,75,111,253,143,122,119,101,75,157,157,132,96,94,112,79,82];
  for (var c = 0; c < widths.length; c++) sheet.setColumnWidth(c+1, widths[c]);

  // Row heights (px)
  var heights = [36,20,27,69,67,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,35,34,28,26,26,26,26,26,44,59,44,76];
  for (var r = 0; r < heights.length; r++) sheet.setRowHeight(r+1, heights[r]);

  Logger.log("BP_MASTER rebuilt: " + R + " rows.");
}

function rebuildDUSTMaster() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("DUST_MASTER");
  if (!sheet) { sheet = ss.insertSheet("DUST_MASTER"); }
  sheet.clear();
  var R = 27, C = 17;

  // Ensure enough rows/cols
  if (sheet.getMaxRows() < R) sheet.insertRowsAfter(sheet.getMaxRows(), R - sheet.getMaxRows());

  var values = [["STOCK \n10.06.2026","","Pushaan Blend Formula ","","DUST","","","DP","=Q4","",""," ","=N26/J26","=M1+M1*5%","=N1+30","","=N1+61"],["","","MASTER FILE ","","","","",100,"Year ","","Total Kg.","Katori ","=Q26/J26","PAID","WH","DP1(+)","DP2"],["","","","","","","","","","","","","","","","",""],["","","BATCH NO.","NO BATCH ",new Date("2026-06-15T00:00:00"),"","","",2025,"","=J26","=Q26/J26","WT","","=H2-O1","=H2-P1","=H2-Q1"],["M/D","BAGS ","L.No.","Particulars ","Grade","Inv.","Kg. Bag","No.  Bags ","%","KG","Open ","Shortage ","Rate ","Amount ","CHAMACH","Chamch Naap","T. WT"],[new Date("2026-06-11T00:00:00"),10,33,"AASHVI","SOF",971,38,0,"=J6/$J$26","=G6*H6","","",174,"=J6*M6",2.8,"=O6*18.5","=J6*O6"],[new Date("2026-06-11T00:00:00"),14,3,"AABELI ","PF",59,40,"","=J7/$J$26","=G7*H7","","",235,"=J7*M7",3.1,"=O7*18.5","=J7*O7"],[new Date("2026-06-08T00:00:00"),20,31,"BHOLATAR","OF",590,30,0,"=J8/$J$26","=G8*H8","","",305,"=J8*M8",3.3,"=O8*18.5","=J8*O8"],[new Date("2026-05-08T00:00:00"),3,2,"CHARAIBARI ","OF",86,34,0,"=J9/$J$26","=G9*H9","","",201,"=J9*M9",2.8,"=O9*18.5","=J9*O9"],[new Date("2026-06-08T00:00:00"),6,22,"HUKAJAN","OF",152,38,0,"=J10/$J$26","=G10*H10","","",215,"=J10*M10",2.8,"=O10*18.5","=J10*O10"],[new Date("2026-05-05T00:00:00"),6,"21A","KISHORIBAG","OF",138,26,0,"=J11/$J$26","=G11*H11","","",330,"=J11*M11",3.3,"=O11*18.5","=J11*O11"],[new Date("2026-05-18T00:00:00"),1,4,"MADHAV","OF",1014,38,0,"=J12/$J$26","=G12*H12","","",235,"=J12*M12",3,"=O12*18.5","=J12*O12"],[new Date("2026-06-11T00:00:00"),8,22,"MANGALMURTI ","OF",381,26,"","=J13/$J$26","=G13*H13","","",245,"=J13*M13",3.3,"=O13*18.5","=J13*O13"],[new Date("2026-05-29T00:00:00"),3,22,"MALIBELIA ","PD1",404,40,0,"=J14/$J$26","=G14*H14","","",228,"=J14*M14",3,"=O14*18.5","=J14*O14"],[new Date("2026-05-27T00:00:00"),4,5,"MALIBELIA ","PF1",363,38,0,"=J15/$J$26","=G15*H15","","",228,"=J15*M15",3,"=O15*18.5","=J15*O15"],["","","","","","","",0,"=J16/$J$26","=G16*H16","","",250,"=J16*M16",3.1,"=O16*18.5","=J16*O16"],["BP","","","","","","",0,"=J17/$J$26","=G17*H17","","",0,"=J17*M17",0,"=O17*18.5","=J17*O17"],["","","","PRATAPGARH BLEND","OF","B-37",30,0,"=J18/$J$26","=G18*H18","","",237,"=J18*M18",3.4,"=O18*18.5","=J18*O18"],["","","","RR DUST","OF","C-67",30,0,"=J19/$J$26","=G19*H19","","",230,"=J19*M19",3.1,"=O19*18.5","=J19*O19"],["","","","SHRI DUST","PD","C-57",30,0,"=J20/$J$26","=G20*H20","","",193,"=J20*M20",2.9,"=O20*18.5","=J20*O20"],["","","","SHRI DUST","OF","C-71",30,0,"=J21/$J$26","=G21*H21","","",210,"=J21*M21",2.8,"=O21*18.5","=J21*O21"],["","","","DUP DUST","OF","C-23",30,0,"=J22/$J$26","=G22*H22","","",260,"=J22*M22",3.26,"=O22*18.5","=J22*O22"],["","",""," ANY Blend","MIX"," ",30,0,"=J23/$J$26","=G23*H23","","",180,"=J23*M23",2.8,"=O23*18.5","=J23*O23"],["Loose Sample Tea","","","","","",6,"=H25","=J24/$J$26","=G24*H24","","",220,"=J24*M24",2.8,"=O24*18.5","=J24*O24"],["  950 KG BLEND  AND ABOVE FIX 6KG LEAF ","","","","","",6,0.01,"=J25/$J$26","=G25*H25",0,"",290,"=J25*M25",1.5,"=O25*18.5","=J25*O25"],["Total ","","","","","","","=SUM(H6:H23)","=SUM(I6:I25)","=SUM(J6:J25)-L26","=SUM(K6:K25)-M26","=SUM(L6:L25)","","=SUM(N6:N25)","","","=SUM(Q6:Q25)"],["BLENDING AREA   DATE ","","","BLENDAR   IN  DATE "," BLENDAR  OUT  DATE ","","","TESTING BY","","C","T","","","","=J26/30","",""]];
  sheet.getRange(1,1,R,C).setValues(values);

  var rng = sheet.getRange(1,1,R,C);
  rng.setBackgrounds([["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#3C3B38","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFF00"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#000000","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#000000","#000000","#FFFF00","#FFFF00","#000000","#000000","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#92D050","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#92D050","#FFFFFF","#FFFF00","#FFFF00","#92D050","#92D050","#92D050","#FFFF00","#92D050","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFF00","#FFFF00","#FFFF00","#FFFF00","#FFFF00","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#000000","#FFFFFF","#000000","#000000","#000000","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"]]);
  rng.setFontColors([["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#BFBFBF","#BFBFBF","#BFBFBF","#BFBFBF","#BFBFBF"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FF0000","#FF0000","#000000","#FF0000","#FF0000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FF0000","#FF0000","#000000","#000000","#000000","#000000","#000000"],["#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF","#FFFFFF"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#FFFFFF","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FFFF00","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FF0000","#FF0000","#FF0000","#FF0000","#FF0000","#000000","#000000","#000000","#000000","#000000"],["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#FFFFFF","#000000","#FFFFFF","#FFFFFF","#FFFFFF","#000000","#000000","#000000","#000000","#000000"]]);
  rng.setFontWeights([["bold","normal","bold","normal","bold","normal","normal","bold","bold","bold","normal","bold","bold","bold","bold","bold","bold"],["normal","normal","bold","normal","normal","normal","normal","bold","bold","normal","bold","bold","bold","bold","bold","bold","bold"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","bold","bold","bold","normal","normal","normal","bold","normal","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","normal","normal","normal","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","normal","normal","normal","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","normal","normal","normal","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","normal","normal","normal","normal","bold","bold","bold","bold","bold","bold","bold","bold","bold","bold"],["bold","normal","normal","bold","bold","normal","normal","bold","normal","bold","bold","bold","bold","bold","bold","normal","normal"]]);
  rng.setFontStyles([["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["italic","italic","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","italic","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"],["normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal","normal"]]);
  rng.setFontSizes([[14.0,11.0,14.0,11.0,16.0,11.0,11.0,16.0,12.0,16.0,11.0,14.0,14.0,14.0,14.0,14.0,14.0],[11.0,11.0,24.0,11.0,11.0,11.0,11.0,16.0,16.0,11.0,16.0,16.0,16.0,11.0,16.0,16.0,16.0],[11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,11.0,18.0,28.0,36.0,11.0,11.0,11.0,14.0,11.0,22.0,18.0,14.0,14.0,14.0,14.0,14.0],[14.0,12.0,18.0,14.0,20.0,14.0,14.0,16.0,14.0,14.0,14.0,14.0,14.0,14.0,14.0,11.0,14.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[14.0,11.0,11.0,11.0,11.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,12.0,18.0,14.0,20.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[11.0,11.0,11.0,11.0,11.0,11.0,18.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[18.0,11.0,11.0,11.0,11.0,11.0,18.0,28.0,11.0,50.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0],[14.0,11.0,11.0,11.0,11.0,11.0,11.0,16.0,16.0,16.0,16.0,16.0,11.0,11.0,11.0,11.0,11.0],[16.0,11.0,11.0,16.0,16.0,11.0,11.0,16.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0,11.0]]);
  rng.setHorizontalAlignments([["center","left","center","left","center","left","left","center","left","center","left","center","left","left","center","center","left"],["left","left","center","left","left","left","left","center","center","left","center","center","right","center","center","center","center"],["left","left","left","left","left","left","left","left","left","left","left","left","left","left","left","left","left"],["left","left","center","center","center","left","left","left","center","left","left","left","center","center","center","center","center"],["center","center","center","center","center","center","center","center","center","center","center","center","center","center","center","center","center"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","left","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","left","left","left","left","left","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","center","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","center","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","center","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","center","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","center","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","center","center","center","center","center","center","center","left","center","left","left","left","left","center","center","left"],["center","left","left","left","left","left","center","center","left","center","left","left","left","left","center","center","left"],["center","left","left","left","left","left","center","center","left","center","left","left","left","left","center","center","left"],["center","left","left","left","left","left","left","center","center","center","center","center","center","center","center","center","center"],["center","left","left","center","center","left","left","center","left","center","center","center","center","center","center","left","left"]]);
  rng.setVerticalAlignment("middle");
  rng.setBorder(true, true, true, true, true, true);

  // Column widths (px)
  var widths = [84,69,106,229,85,80,73,101,67,157,119,90,81,94,112,85,70];
  for (var c = 0; c < widths.length; c++) sheet.setColumnWidth(c+1, widths[c]);

  // Row heights (px)
  var heights = [25,20,20,57,47,32,32,32,32,32,32,32,32,32,32,32,31,21,21,21,21,21,38,81,66,62,44];
  for (var r = 0; r < heights.length; r++) sheet.setRowHeight(r+1, heights[r]);

  Logger.log("DUST_MASTER rebuilt: " + R + " rows.");
}
function renameBlend(blendId, newBatchNo, newCustomer) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("BLEND_LOG");
  var data = sheet.getDataRange().getValues();
  var sheetUrl = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== blendId) continue;
    if (!sheetUrl) sheetUrl = data[i][18];
    if (newBatchNo) sheet.getRange(i + 1, 6).setValue(newBatchNo);
    if (newCustomer !== undefined) sheet.getRange(i + 1, 5).setValue(newCustomer);
  }
  if (sheetUrl && newBatchNo) {
    try {
      var idMatch = String(sheetUrl).match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (idMatch) {
        var file = DriveApp.getFileById(idMatch[1]);
        var d = file.getName().split(' ')[0]; // keep date prefix
        file.setName(d + ' ' + newBatchNo);
      }
    } catch(e) { Logger.log('Rename file error: ' + e); }
  }
  return sanitizeForClient({ success: true });
}

function debugDustLayout() {
  var sheet = getSpreadsheet().getSheetByName("DUST_MASTER");
  var L = detectLayout(sheet);
  Logger.log("totalRow: " + L.totalRow);
  Logger.log("varietyRows: " + JSON.stringify(L.varietyRows));
  Logger.log("dividerRows: " + JSON.stringify(L.dividerRows));
  Logger.log("anyBlendRows: " + JSON.stringify(L.anyBlendRows));
  Logger.log("leafRow: " + L.leafRow + "  looseSampleRow: " + L.looseSampleRow);
  Logger.log("lastRow: " + sheet.getLastRow());
}

function debugWhichSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  var ss = SpreadsheetApp.openById(id);
  Logger.log("SPREADSHEET_ID: " + id);
  Logger.log("Name: " + ss.getName());
  Logger.log("URL: " + ss.getUrl());
  var m = ss.getSheetByName("DUST_MASTER");
  Logger.log("DUST_MASTER lastRow: " + m.getLastRow());
  Logger.log("D9: " + m.getRange(9,4).getValue() + " | D10: " + m.getRange(10,4).getValue());
}
