/**
 * ============================================================
 * PUSHAAN BLENDING SYSTEM - SETUP SCRIPT (v2 - exact colors)
 * ============================================================
 * RUN THIS ONCE. It will:
 *  1. Create the Drive folder structure (BP/DUST -> RR/DUP/SK/LOOSE)
 *  2. Create the main spreadsheet "Pushaan Blending System"
 *  3. Build BP_MASTER and DUST_MASTER as exact formula + color replicas
 *  4. Build BP_TRACKER, DUST_TRACKER, BLEND_LOG, CONFIG sheets
 *
 * HOW TO RUN:
 *  - Open script.google.com -> New Project
 *  - Paste this entire file as "Setup.gs" (replace any existing content)
 *  - In the function dropdown at top, select "runFullSetup"
 *  - Click Run (authorize permissions if asked)
 *  - Check Execution Log for the Spreadsheet URL + Folder URL
 *
 * NOTE: If you already ran v1 of this script, running this again will
 * create a SECOND spreadsheet + folder. Either delete the old ones first,
 * or run this fresh and use the new URLs going forward.
 * ============================================================
 */

function runFullSetup() {
  Logger.log("Starting Pushaan Blending System setup...");

  var folders = createFolderStructure();
  Logger.log("Root folder: " + folders.root.getUrl());

  var ss = SpreadsheetApp.create("Pushaan Blending System");
  var ssFile = DriveApp.getFileById(ss.getId());
  folders.root.addFile(ssFile);
  DriveApp.getRootFolder().removeFile(ssFile);

  Logger.log("Spreadsheet created: " + ss.getUrl());

  buildBpMaster(ss);
  buildDustMaster(ss);
  buildTracker(ss, "BP_TRACKER");
  buildTracker(ss, "DUST_TRACKER");
  buildBlendLog(ss);
  buildConfig(ss);

  var defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }

  var props = PropertiesService.getScriptProperties();
  props.setProperty("SPREADSHEET_ID", ss.getId());
  props.setProperty("ROOT_FOLDER_ID", folders.root.getId());
  props.setProperty("BP_RR", folders.BP_RR.getId());
  props.setProperty("BP_DUP", folders.BP_DUP.getId());
  props.setProperty("BP_SK", folders.BP_SK.getId());
  props.setProperty("BP_LOOSE", folders.BP_LOOSE.getId());
  props.setProperty("DUST_RR", folders.DUST_RR.getId());
  props.setProperty("DUST_DUP", folders.DUST_DUP.getId());
  props.setProperty("DUST_SK", folders.DUST_SK.getId());
  props.setProperty("DUST_LOOSE", folders.DUST_LOOSE.getId());

  Logger.log("============================================");
  Logger.log("SETUP COMPLETE!");
  Logger.log("Spreadsheet URL: " + ss.getUrl());
  Logger.log("Drive Folder URL: " + folders.root.getUrl());
  Logger.log("============================================");
}


function createFolderStructure() {
  var root = DriveApp.createFolder("Pushaan Blending System");

  var bp = root.createFolder("BP");
  var bp_rr = bp.createFolder("RR");
  var bp_dup = bp.createFolder("DUP");
  var bp_sk = bp.createFolder("SK");
  var bp_loose = bp.createFolder("LOOSE");

  var dust = root.createFolder("DUST");
  var dust_rr = dust.createFolder("RR");
  var dust_dup = dust.createFolder("DUP");
  var dust_sk = dust.createFolder("SK");
  var dust_loose = dust.createFolder("LOOSE");

  return {
    root: root,
    BP: bp, BP_RR: bp_rr, BP_DUP: bp_dup, BP_SK: bp_sk, BP_LOOSE: bp_loose,
    DUST: dust, DUST_RR: dust_rr, DUST_DUP: dust_dup, DUST_SK: dust_sk, DUST_LOOSE: dust_loose
  };
}


function buildBpMaster(ss) {
  var sheet = ss.insertSheet("BP_MASTER");
  var data = [
  ['STOCK \n01.05.2026', null, 'Pushaan Blend Formula ', null, 'BP', null, null, ' DP', '=Q4', null, null, null, '=N30/J30', '=M1+5%*(M1)', '=N1+30', null, '=N1+61'], // row 1
  [null, null, 'MASTER FILE DON’T CHANGE ', null, null, null, null, 100, 'YEAR', null, 'Total Kg.', 'Katori', '=Q30/J30', 'PAID ', 'WH', null, 'DP2'], // row 2
  [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], // row 3
  [null, null, 'Batch No. ', 'NO BATCH ', '2026-05-20 00:00:00', null, null, null, 2026, null, '=J30', '=Q30/J30', 'WT', null, '=H2-O1', '=H2-P1', '=H2-Q1'], // row 4
  ['M/D', 'BAGS ', 'PLACE', 'Particulars ', 'BP/DUST', 'INV', 'Kg. Bag', 'No.  Bags ', '%', 'KG', 'Open ', 'Shortage ', 'Rate ', 'AMOUNT  ', 'CHAMACH', 'Chamch Naap', 'T. Wt '], // row 5
  ['2026-05-29 15:05:59', 28, 16, 'BHOLATAR', 'BP', 645, 30, 0, '=(J6/$J$30)', '=G6*H6', null, null, 330, '=(J6+L6)*M6', 3, '=O6*18.5', '=J6*O6'], // row 6
  ['2026-05-28 18:00:19', 15, 10, 'HASIMARA', 'BP', 520, 26, null, '=(J7/$J$30)', '=G7*H7', null, null, 265, '=(J7+L7)*M7', 3, '=O7*18.5', '=J7*O7'], // row 7
  ['2026-05-28 18:03:09', 16, 21, 'HAKATHIPUKRI', 'BP', 283, 30, 0, '=(J8/$J$30)', '=G8*H8', null, null, 300, '=(J8+L8)*M8', 3.3, '=O8*18.5', '=J8*O8'], // row 8
  ['2026-05-28 17:53:51', 5, 9, 'HOLLANGBARI', 'BP', 91, 35, null, '=(J9/$J$30)', '=G9*H9', null, null, 225, '=(J9+L9)*M9', 3, '=O9*18.5', '=J9*O9'], // row 9
  ['2026-05-29 15:05:59', 10, 22, 'KISHORIBAG ', 'BP1', 150, 35, null, '=(J10/$J$30)', '=G10*H10', null, null, 250, '=(J10+L10)*M10', 3, '=O10*18.5', '=J10*O10'], // row 10
  ['2026-05-29 15:05:58', 17, 21, 'MAIBELIA', 'BP1', 401, 36, null, '=(J11/$J$30)', '=G11*H11', null, null, 228, '=(J11+L11)*M11', 2.9, '=O11*18.5', '=J11*O11'], // row 11
  ['2026-05-29 14:56:19', 15, 30, 'MAIBELIA', 'BOPSM1', 417, 34, null, '=(J12/$J$30)', '=G12*H12', null, null, 228, '=(J12+L12)*M12', 2.8, '=O12*18.5', '=J12*O12'], // row 12
  ['2026-05-28 17:55:13', 8, 7, 'MADHAV', 'BP', 1015, 38, null, '=(J13/$J$30)', '=G13*H13', null, null, 235, '=(J13+L13)*M13', 3, '=O13*18.5', '=J13*O13'], // row 13
  ['2026-05-28 18:08:10', 11, 1, 'POMI', 'BP', 129, 34, null, '=(J14/$J$30)', '=G14*H14', null, null, 230, '=(J14+L14)*M14', 2.8, '=O14*18.5', '=J14*O14'], // row 14
  ['2026-05-28 17:55:56', 6, 4, 'PURI', 'BP', 52, 38, null, '=(J15/$J$30)', '=G15*H15', null, null, 318, '=(J15+L15)*M15', 3, '=O15*18.5', '=J15*O15'], // row 15
  ['2026-05-28 17:56:20', 15, 8, 'RANGDDHALI', 'BP', 14, 35, null, '=(J16/$J$30)', '=G16*H16', null, null, 235, '=(J16+L16)*M16', 3, '=O16*18.5', '=J16*O16'], // row 16
  ['2026-05-29 14:19:41', 20, 13, 'RUGBY ', 'BP', '113 , 122 , 131', 35, null, '=(J17/$J$30)', '=G17*H17', null, null, 205, '=(J17+L17)*M17', 2.75, '=O17*18.5', '=J17*O17'], // row 17
  ['2026-05-28 18:06:19', 7, 15, 'SUPRABHA', 'BP', 95, 35, null, '=(J18/$J$30)', '=G18*H18', null, null, 180, '=(J18+L18)*M18', 2.8, '=O18*18.5', '=J18*O18'], // row 18
  ['2026-05-28 18:05:33', 3, 12, 'THAWRAJAAN ', 'BP', 68, 35, null, '=(J19/$J$30)', '=G19*H19', null, null, 211, '=(J19+L19)*M19', 3, '=O19*18.5', '=J19*O19'], // row 19
  ['2026-05-28 18:05:33', null, 'OF', null, null, null, null, 0, '=(J20/$J$30)', null, null, null, null, null, 0, '=O20*18.5', '=J20*O20'], // row 20
  ['2026-05-28 18:15:54', 7, 31, 'KISHORIBAG', 'OF1', 152, 35, null, '=(J21/$J$30)', '=G21*H21', null, null, 250, '=J21*M21', 3, '=O21*18.5', '=J21*O21'], // row 21
  ['2026-05-28 18:09:57', null, null, null, null, null, null, null, '=(J22/$J$30)', null, null, null, null, null, null, null, null], // row 22
  ['2026-05-25 18:09:13', null, null, 'ANY BLEND DUP', 'BP', 'E-88', 30, 0, '=(J23/$J$30)', '=G23*H23', null, null, 276, '=(J23+L23)*M23', 3, '=O23/18.5', '=J23*O23'], // row 23
  ['2026-05-25 18:09:13', null, null, 'ANY BLEND RR ', 'BP', 'F-26', 30, 0, '=(J24/$J$30)', '=G24*H24', null, null, 236, '=(J24+L24)*M24', 3, '=O24/18.5', '=J24*O24'], // row 24
  [null, null, null, 'ANY BLEND SHRI', 'BP', 'F-28', 30, 0, '=(J25/$J$30)', '=G25*H25', null, null, 202, '=(J25+L25)*M25', 2.9, '=O25/18.5', '=J25*O25'], // row 25
  [null, null, null, 'ANY BLEND JODHPUR', 'BP', 'E-18', 30, 0, '=(J26/$J$30)', '=G26*H26', null, null, 178, '=(J26+L26)*M26', 2.85, '=O26/18.5', '=J26*O26'], // row 26
  [null, null, null, 'MIX  BLEND ', 'MIX ', null, 30, 0, '=(J27/$J$30)', '=G27*H27', null, null, 180, '=(J27+L27)*M27', 3, '=O27/18.5', '=J27*O27'], // row 27
  ['Loose Sample Tea', null, null, null, null, null, 7, '=H29', '=(J28/$J$30)', '=G28*H28', null, null, 220, '=J28*M28', 3, '=O28/18.5', '=J28*O28'], // row 28
  ['  950 KG BLEND  AND ABOVE FIX 7KG LEAF ', null, null, null, null, null, 7, 0.01, '=(J29/$J$30)', '=G29*H29', 0, null, 290, '=J29*M29', 2, '=O29/18.5', '=J29*O29'], // row 29
  ['Total ', null, null, null, null, null, null, '=SUM(H6:H27)', '=J30/$J$30', '=SUM(J6:J29)', '=SUM(K6:K29)', '=SUM(L6:L29)', null, '=SUM(N6:N29)', null, null, '=SUM(Q6:Q29)'], // row 30
  ['BLENDING AREA          DATE ', null, null, 'BLENDAR       IN              DATE ', ' BLENDAR    OUT              DATE ', null, null, 'TESTING', null, 'C', 'T', null, null, null, '=J30/30', null, null], // row 31
  [null, null, null, null, null, null, null, 'MOHIT', null, null, null, null, null, null, null, null, null], // row 32
];
  sheet.getRange(1, 1, data.length, 17).setValues(data);
  formatMasterSheet(sheet, "BP", 30);
  applyBpFills(sheet);
  applyBpFontColors(sheet);
  Logger.log("BP_MASTER built.");
}


function buildDustMaster(ss) {
  var sheet = ss.insertSheet("DUST_MASTER");
  var data = [
  [' \nSTOCK', null, 'Pushaan Blend Formula ', null, 'DUST', null, null, 'DP', '=Q4', null, null, ' ', '=N29/J29', '=M1+M1*5%', '=N1+30', null, '=N1+61'], // row 1
  [null, null, 'MASTER FILE ', null, null, null, null, 100, 'Year ', null, 'Total Kg.', 'Katori ', '=Q29/J29', 'PAID', 'WH', 'DP1(+)', 'DP2'], // row 2
  [null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null], // row 3
  ['=E4', null, 'BATCH NO.', 'NO BATCH ', '2026-05-29 00:00:00', null, null, null, 2026, null, '=J29', '=Q29/J29', 'WT', null, '=H2-O1', '=H2-P1', '=H2-Q1'], // row 4
  ['M/D', 'BAGS ', 'L.No.', 'Particulars ', 'Grade', 'Inv.', 'Kg. Bag', 'No.  Bags ', '%', 'KG', 'Open ', 'Shortage ', 'Rate ', 'Amount ', 'CHAMACH', 'Chamch Naap', 'T. WT'], // row 5
  ['2026-05-29 13:53:24', 2, 21, 'BEECH ', 'OF', 4, 38, 0, '=J6/$J$29', '=G6*H6', null, null, 320, '=J6*M6', 3.3, '=O6*18.5', '=J6*O6'], // row 6
  ['2026-05-29 13:53:26', 9, 2, 'CHARAIBARI ', 'OF', 86, 34, 0, '=J7/$J$29', '=G7*H7', null, null, 201, '=J7*M7', 2.8, '=O7*18.5', '=J7*O7'], // row 7
  ['2026-06-02 13:48:53', 24, '21A', 'KISHORIBAG', 'OF', 138, 26, 0, '=J8/$J$29', '=G8*H8', null, null, 330, '=J8*M8', 3.3, '=O8*18.5', '=J8*O8'], // row 8
  ['2026-05-29 13:43:04', 9, 31, 'KISHORIBAG', 'OF1', 152, 35, null, '=J9/$J$29', '=G9*H9', null, null, 250, '=J9*M9', 3, '=O9*18.5', '=J9*O9'], // row 9
  ['2026-05-29 13:49:29', 3, 33, 'MAAKAMAKHIYA', 'PD', 'UK74', 40, 0, '=J10/$J$29', '=G10*H10', null, null, 275, '=J10*M10', 3.1, '=O10*18.5', '=J10*O10'], // row 10
  ['2026-05-29 13:42:55', 12, 4, 'MADHAV', 'OF', 1014, 38, 0, '=J11/$J$29', '=G11*H11', null, null, 235, '=J11*M11', 3, '=O11*18.5', '=J11*O11'], // row 11
  ['2026-05-29 13:43:14', 24, 5, 'MALIBELIA ', 'PF1', 363, 39, null, '=J12/$J$29', '=G12*H12', null, null, 228, '=J12*M12', 3, '=O12*18.5', '=J12*O12'], // row 12
  ['2026-05-29 13:43:24', 7, 32, 'RUGBY', 'OF', 105, 35, 0, '=J13/$J$29', '=G13*H13', null, null, 205, '=J13*M13', 2.8, '=O13*18.5', '=J13*O13'], // row 13
  ['2026-06-02 13:37:25', 16, 22, 'MALIBELIA ', 'PD1', 404, 40, null, '=J14/$J$29', '=G14*H14', null, null, 228, '=J14*M14', 3, '=O14*18.5', '=J14*O14'], // row 14
  ['2026-06-02 13:37:25', 21, 3, 'MALIBELIA ', 'PF1', 419, 38, null, '=J15/$J$29', '=G15*H15', null, null, 228, '=J15*M15', 3, '=O15*18.5', '=J15*O15'], // row 15
  ['2026-06-02 13:37:29', 1, 2, 'SHRTHI', 'PD', 175, 38, 0, '=J16/$J$29', '=G16*H16', null, null, 195, '=J16*M16', 2.7, '=O16*18.5', '=J16*O16'], // row 16
  ['PD/PF ', null, null, null, null, null, null, 0, '=J17/$J$29', '=G17*H17', null, null, null, '=J17*M17', 0, '=O17*18.5', '=J17*O17'], // row 17
  ['2026-05-26 10:43:07', null, null, null, null, null, null, null, null, '=G18*H18', null, null, null, null, null, null, '=J18*O18'], // row 18
  ['2026-05-25 17:52:59', null, null, null, null, null, null, null, '=J19/$J$29', '=G19*H19', null, null, 250, '=J19*M19', 3.1, '=O19*18.5', '=J19*O19'], // row 19
  ['BP', null, null, null, null, null, null, 0, '=J20/$J$29', '=G20*H20', null, null, 0, '=J20*M20', 0, '=O20*18.5', '=J20*O20'], // row 20
  ['2026-05-25 17:52:59', null, null, 'PRATAPGARH BLEND', 'OF', 'B-37', 30, 0, '=J21/$J$29', '=G21*H21', null, null, 237, '=J21*M21', 3.4, '=O21*18.5', '=J21*O21'], // row 21
  ['2026-05-25 17:52:59', null, null, 'RR DUST', 'OF', 'C-67', 30, 0, '=J22/$J$29', '=G22*H22', null, null, 230, '=J22*M22', 3.1, '=O22*18.5', '=J22*O22'], // row 22
  ['2026-05-25 17:52:59', null, null, 'SHRI DUST', 'PD', 'C-57', 30, 0, '=J23/$J$29', '=G23*H23', null, null, 193, '=J23*M23', 2.9, '=O23*18.5', '=J23*O23'], // row 23
  ['2026-05-25 17:52:59', null, null, 'SHRI DUST', 'OF', 'C-71', 30, 0, '=J24/$J$29', '=G24*H24', null, null, 210, '=J24*M24', 2.8, '=O24*18.5', '=J24*O24'], // row 24
  ['2026-05-25 17:52:59', null, null, 'DUP DUST', 'OF', 'C-23', 30, 0, '=J25/$J$29', '=G25*H25', null, null, 260, '=J25*M25', 3.26, '=O25*18.5', '=J25*O25'], // row 25
  ['2026-05-25 17:52:59', null, null, ' ANY Blend', 'MIX', ' ', 30, 0, '=J26/$J$29', '=G26*H26', null, null, 180, '=J26*M26', 2.8, '=O26*18.5', '=J26*O26'], // row 26
  ['Loose Sample Tea', null, null, null, null, null, 6, '=H28', '=J27/$J$29', '=G27*H27', null, null, 220, '=J27*M27', 2.8, '=O27*18.5', '=J27*O27'], // row 27
  ['  950 KG BLEND  AND ABOVE FIX 6KG LEAF ', null, null, null, null, null, 6, 0.01, '=J28/$J$29', '=G28*H28', 0, null, 290, '=J28*M28', 1.5, '=O28*18.5', '=J28*O28'], // row 28
  ['Total ', null, null, null, null, null, null, '=SUM(H6:H26)', '=SUM(I8:I28)', '=SUM(J6:J28)-L29', '=SUM(K8:K28)-M29', '=SUM(L6:L28)', null, '=SUM(N8:N28)', null, null, '=SUM(Q8:Q28)'], // row 29
  ['BLENDING AREA   DATE ', null, null, 'BLENDAR   IN  DATE ', ' BLENDAR  OUT  DATE ', null, null, 'TESTING BY', null, 'C', 'T', null, null, null, '=J29/30', null, null], // row 30
  [null, null, null, null, null, null, null, 'MOHIT', null, null, null, null, null, null, null, null, null], // row 31
  [null, null, null, null, null, null, null, 'S. R. ', null, null, null, null, null, null, null, null, null], // row 32
];
  sheet.getRange(1, 1, data.length, 17).setValues(data);
  formatMasterSheet(sheet, "DUST", 29);
  applyDustFills(sheet);
  applyDustFontColors(sheet);
  Logger.log("DUST_MASTER built.");
}


/**
 * Base formatting applied to both masters before color-specific overrides:
 * - Freeze rows, column widths, bold rows, print area, date formats
 */
function formatMasterSheet(sheet, category, totalRow) {
  sheet.setFrozenRows(5);

  var widths = [95, 50, 55, 130, 70, 70, 60, 65, 55, 55, 50, 60, 55, 75, 60, 70, 60];
  for (var c = 0; c < widths.length; c++) {
    sheet.setColumnWidth(c + 1, widths[c]);
  }

  // Bold rows: header row 5, total row, footer label row
  sheet.getRange(5, 1, 1, 17).setFontWeight("bold");
  sheet.getRange(totalRow, 1, 1, 17).setFontWeight("bold");

  // Date format for M/D column (col A), rows 6 to totalRow-2
  sheet.getRange(6, 1, totalRow - 7, 1).setNumberFormat("dd/mm/yyyy hh:mm");

  // Set print area note + tab color
  sheet.setTabColor(category === "BP" ? "#4a86e8" : "#b45f06");

  sheet.getRange(totalRow + 5, 4).setValue("Print area: A1:L" + totalRow + " (one A4 page)");
  sheet.getRange(totalRow + 5, 4).setFontStyle("italic").setFontColor("#999999");
}


/**
 * Apply exact fill colors from BP_MASTER_FIXED.xlsm
 * - Dark header cell (I1)
 * - Yellow formula/highlight cells
 * - Black section-divider rows (with white bold text)
 * - Light gray formula-result columns (M-Q for variety rows)
 * - Green fixed rows (Loose Sample Tea / 950KG)
 */
function applyBpFills(sheet) {
  sheet.getRangeList(["I1"]).setBackground("#3C3B38");
  sheet.getRangeList(["O1", "I2", "K2", "L2", "M4", "Q4", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "M20", "B21", "H28", "I28", "J28", "N28", "I29", "J29", "N29", "H30", "I30", "J30", "K30", "L30"]).setBackground("#FFFF00");
  sheet.getRangeList(["A5", "B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5", "K5", "L5", "M5", "N5", "O5", "P5", "Q5", "B20", "C20", "D20", "E20", "F20", "G20", "N20", "B22", "H31", "J31", "K31", "L31"]).setBackground("#000000").setFontColor("#FFFFFF");
  sheet.getRangeList(["M6", "N6", "O6", "P6", "Q6", "M7", "N7", "O7", "P7", "Q7", "M8", "N8", "O8", "P8", "Q8", "M9", "N9", "O9", "P9", "Q9", "M10", "N10", "O10", "P10", "Q10", "M11", "N11", "O11", "P11", "Q11", "M12", "N12", "O12", "P12", "Q12", "M13", "N13", "O13", "P13", "Q13", "M14", "N14", "O14", "P14", "Q14", "M15", "N15", "O15", "P15", "Q15", "M16", "N16", "O16", "P16", "Q16", "M17", "N17", "O17", "P17", "Q17", "M18", "N18", "O18", "P18", "Q18", "M19", "N19", "O19", "P19", "Q19", "O20", "P20", "Q20", "N23", "O23", "Q23", "N24", "O24", "Q24", "N25", "O25", "Q25", "N26", "O26", "Q26", "N27", "O27", "Q27", "O28", "Q28", "O29", "Q29"]).setBackground("#D9D9D9");
  sheet.getRangeList(["A28", "G28", "K28", "L28", "A29", "G29", "K29", "L29", "M29"]).setBackground("#92D050");
}


/**
 * Apply exact fill colors from DUST_MASTER_FIXED.xlsm
 */
function applyDustFills(sheet) {
  sheet.getRangeList(["I1"]).setBackground("#3C3B38");
  sheet.getRangeList(["I2", "K2", "L2", "O2", "P2", "Q2", "M4", "B6", "I6", "J6", "B7", "I7", "J7", "B8", "I8", "J8", "B9", "I9", "J9", "B10", "I10", "J10", "B11", "I11", "J11", "B12", "I12", "J12", "B13", "I13", "J13", "B14", "I14", "J14", "B15", "I15", "J15", "B16", "I16", "J16", "I17", "J17", "N17", "B18", "I18", "J18", "B19", "I19", "J19", "I20", "J20", "I21", "J21", "I22", "J22", "I23", "J23", "I24", "J24", "I25", "J25", "I26", "J26", "H27", "I27", "J27", "I28", "J28", "N28", "H29", "I29", "J29", "K29", "L29"]).setBackground("#FFFF00");
  sheet.getRangeList(["A5", "B5", "C5", "D5", "E5", "F5", "G5", "H5", "I5", "J5", "K5", "L5", "M5", "N5", "O5", "P5", "Q5", "A17", "G17", "H17", "K17", "L17", "A20", "G20", "H20", "K20", "L20", "H30", "J30", "K30", "L30"]).setBackground("#000000").setFontColor("#FFFFFF");
  sheet.getRangeList(["A28", "G28", "K28", "L28", "M28", "O28"]).setBackground("#92D050");
}


/**
 * Apply exact font colors (red warning text, dark red, yellow-on-yellow) for BP
 */
function applyBpFontColors(sheet) {
  sheet.getRangeList(["O1"]).setFontColor("#C00000");
  sheet.getRangeList(["H2", "I2", "K2", "L2", "K4", "L4", "M20", "H30", "J30", "K30", "L30"]).setFontColor("#FF0000");
  sheet.getRangeList(["J28"]).setFontColor("#FFFF00");
}


/**
 * Apply exact font colors for DUST
 */
function applyDustFontColors(sheet) {
  sheet.getRangeList(["H2", "I2", "K2", "L2", "K4", "L4", "H29", "I29", "J29", "K29", "L29"]).setFontColor("#FF0000");
  sheet.getRangeList(["J27"]).setFontColor("#FFFF00");
}


function buildTracker(ss, name) {
  var sheet = ss.insertSheet(name);
  var category = name.indexOf("BP") === 0 ? "BP" : "DUST";
  var crossLabel = category === "BP" ? "Used in DUST" : "Used in BP";

  var headers = [
    "Tea Variety", "Grade", "Invoice", "Opening Stock", "Remaining Stock",
    "RR Used", "RR Dates", "DUP Used", "DUP Dates",
    "SK Used", "SK Dates", "LOOSE Used", "LOOSE Dates",
    crossLabel + " Used", crossLabel + " Dates",
    "Total Used", "Stock Updated On"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9d9d9");
  sheet.setFrozenRows(1);

  for (var c = 1; c <= headers.length; c++) {
    sheet.setColumnWidth(c, 100);
  }
  sheet.setColumnWidth(1, 150);

  sheet.setTabColor(category === "BP" ? "#4a86e8" : "#b45f06");
  Logger.log(name + " built.");
}


function buildBlendLog(ss) {
  var sheet = ss.insertSheet("BLEND_LOG");

  var headers = [
    "Blend_ID", "Date", "Category", "Blend_Type", "Customer", "Batch_No",
    "Base_Price", "Particulars", "Grade", "Invoice", "Bags_Used",
    "Kg_Bag", "KG", "Shortage", "Rate", "Amount", "CHAMACH", "T_Wt",
    "Sheet_URL"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9d9d9");
  sheet.setFrozenRows(1);

  for (var c = 1; c <= headers.length; c++) {
    sheet.setColumnWidth(c, 100);
  }

  sheet.setTabColor("#999999");
  Logger.log("BLEND_LOG built.");
}


function buildConfig(ss) {
  var sheet = ss.insertSheet("CONFIG");

  var data = [
    ["Key", "Value"],
    ["BP_RR_COUNTER", 0],
    ["BP_DUP_COUNTER", 0],
    ["BP_SK_COUNTER", 0],
    ["BP_LOOSE_COUNTER", 0],
    ["DUST_RR_COUNTER", 0],
    ["DUST_DUP_COUNTER", 0],
    ["DUST_SK_COUNTER", 0],
    ["DUST_LOOSE_COUNTER", 0]
  ];

  sheet.getRange(1, 1, data.length, 2).setValues(data);
  sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#d9d9d9");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setTabColor("#666666");
  Logger.log("CONFIG built.");
}
