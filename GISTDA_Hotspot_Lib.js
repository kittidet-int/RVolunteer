/*
Copyright (C) 2026 Kittidet Intharaksa

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

// ==========================================
// CONFIGURATION DEFAULTS
// ==========================================
var CONFIG = {
  SHEET_RAW: 'Daily',
  SHEET_COUNTRY: 'Semantic_Country',
  SHEET_PROVINCE: 'Semantic_Province_TH',
  SHEET_LANDUSE: 'Semantic_Landuse_TH',
};

// ==========================================
// 1. PUBLIC FUNCTIONS
// ==========================================

/**
 * @param {string} folderId - Google Drive's folder ID to store the data file
 * @param {boolean} isTestMode - If true, will be in Test mode
 * @return {Spreadsheet} - Return the Spreadsheet
 */
function executePipeline(folderId, isTestMode) {
  if (!folderId) {
    throw new Error("[GISTDA_Hotspot_Lib]: Folder ID cannot be null");
  }

  // Define the archive file name (T-2)
  const today = new Date();
  const archiveDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 2);
  const archiveDateStr = Utilities.formatDate(archiveDate, Session.getScriptTimeZone(), "yyyyMMdd");

  let masterFileName = "Hotspot_Data";
  let archiveFileName = `Hotspot_${archiveDateStr}`;

  if (isTestMode) {
    masterFileName = "Hotspot_Data_TEST";
    archiveFileName = `${masterFileName}_${archiveDateStr}`;
    Logger.log("[GISTDA_Hotspot_Lib] Running in TEST MODE");
  }

  Logger.log(`[GISTDA_Hotspot_Lib] Processing date: ${Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd")}`);

  // 1. Prepare File
  const ss = _prepareHotspotSpreadsheet(folderId, masterFileName, archiveFileName);

  // 2. Fetch API
  _fetchDataToSpreadsheet(ss);

  // 3. Generate Semantic Tables
  _generateSemanticTables(ss);

  return ss;
}

// ==========================================
// 2. INTERNAL FUNCTIONS (Private Helpers)
// ==========================================

function _prepareHotspotSpreadsheet(folderId, masterName, archiveFileName) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(masterName);
  let ss;

  if (files.hasNext()) {
    const file = files.next();
    ss = SpreadsheetApp.open(file);

    // Archive Logic
    const existingArchives = folder.getFilesByName(archiveFileName);
    // if (existingArchives.hasNext()) {
    //   existingArchives.next().setTrashed(true);
    // }
    file.makeCopy(archiveFileName, folder);

    // Clear Sheets
    let dailySheet = ss.getSheetByName(CONFIG.SHEET_RAW);
    if (!dailySheet) {
      dailySheet = ss.insertSheet(CONFIG.SHEET_RAW);
    }
    dailySheet.clear();

    // Cleanup
    const allSheets = ss.getSheets();
    const keepSheets = [CONFIG.SHEET_RAW];
    allSheets.forEach(sheet => {
      if (!keepSheets.includes(sheet.getName())) {
        ss.deleteSheet(sheet);
      }
    });
  } else {
    ss = SpreadsheetApp.create(masterName);
    const sheet = ss.getActiveSheet();
    sheet.setName(CONFIG.SHEET_RAW);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  }
  return ss;
}

function _fetchDataToSpreadsheet(ss) {
  const BASE_ENDPOINT = "https://api-gateway.gistda.or.th/api/2.0/resources/features/viirs/1day";
  const apiKey = PropertiesService.getScriptProperties().getProperty('GISTDA_KEY');

  if (!apiKey) {
    throw new Error("[GISTDA_Hotspot_Lib]: API Key is not found");
  }

  let allRows = [];
  let offset = 0;
  let hasMoreData = true;
  const LIMIT = 1000;
  const MAX_LIMIT = 100000;

  // Headers
  const headers = [
    "_createdAt", "_createdBy", "_id", "_updatedAt", "_updatedBy", "acq_date", "acq_time",
    "amphoe", "amphoe_t", "ap_code", "ap_en", "ap_idn", "ap_tn", "bright_ti4", "bright_ti5",
    "changwat", "confidence", "ct_en", "ct_tn", "f_alarm", "file_name", "frp", "hotspotid",
    "instrument", "latitude", "linkgmap", "longitude", "lu_code", "lu_hp", "lu_hp_name",
    "lu_name", "moo_1", "name_1", "province_t", "pv_code", "pv_en", "pv_idn", "pv_tn",
    "re_nesdb", "re_royin", "satellite", "scan", "tambol", "tambon_t", "tb_code", "tb_en",
    "tb_idn", "tb_tn", "th_date", "th_time", "timestamp", "track", "utm_e", "utm_n",
    "utm_zone", "v_angle", "v_direct", "v_dist", "version", "village"
  ];

  Logger.log("[GISTDA_Hotspot_Lib] Data retrieving starts...")

  do {
    const url = `${BASE_ENDPOINT}?api_key=${apiKey}&limit=${LIMIT}&offset=${offset}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      throw new Error(`[GISTDA_Hotspot_Lib]: API Error (${responseCode})`);
    }

    const json = JSON.parse(response.getContentText());
    const features = json.features || [];
    if (features.length === 0) {
      break;
    }

    const batchRows = features.map(item => {
      const p = item.properties;
      let cleanDate = p.th_date ? p.th_date.split('T')[0] : "";
      return [
        p._createdAt, p._createdBy, p._id, p._updatedAt, p._updatedBy, p.acq_date, p.acq_time,
        p.amphoe, p.amphoe_t, p.ap_code, p.ap_en, p.ap_idn, p.ap_tn, p.bright_ti4, p.bright_ti5,
        p.changwat, p.confidence, p.ct_en, p.ct_tn, p.f_alarm, p.file_name, p.frp, p.hotspotid,
        p.instrument, p.latitude, p.linkgmap, p.longitude, p.lu_code, p.lu_hp, p.lu_hp_name,
        p.lu_name, p.moo_1, p.name_1, p.province_t, p.pv_code, p.pv_en, p.pv_idn, p.pv_tn,
        p.re_nesdb, p.re_royin, p.satellite, p.scan, p.tambol, p.tambon_t, p.tb_code, p.tb_en,
        p.tb_idn, p.tb_tn, cleanDate, p.th_time, p.timestamp, p.track, p.utm_e, p.utm_n,
        p.utm_zone, p.v_angle, p.v_direct, p.v_dist, p.version, p.village
      ];
    });

    // Data concatenation
    allRows = allRows.concat(batchRows);
    Logger.log(`[GISTDA_Hotspot_Lib] Data retrieved: ${allRows.length} rows`);

    offset += LIMIT;

    // 1. Safety Break
    if (offset > MAX_LIMIT) {
      Logger.log(`[GISTDA_Hotspot_Lib] STOP: Data retrieving exceed the maximum limit (${MAXIMUM_LIMIT})`);
      break;
    }

    // 2. Logical Stop
    // - json.numberMatched : API tells the total records
    // - features.length < LIMIT : API doesn't tell the total records
    if ((json.numberMatched && offset >= json.numberMatched) || features.length < LIMIT) {
      hasMoreData = false;
      Logger.log(`[GISTDA_Hotspot_Lib] Data retrieving finished: ${json.numberMatched} rows`);
    }
  } while (hasMoreData);

  if (allRows.length > 0) {
    const sheet = ss.getSheetByName(CONFIG.SHEET_RAW);
    sheet.appendRow(headers);
    sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);

    // Formatting
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#D9EAD3");
    // Format Lat/Long & Numbers if needed
    sheet.getRange(2, 1, allRows.length, 2).setNumberFormat("0.000000000");

    Logger.log(`[GISTDA_Hotspot_Lib] Data saved: ${allRows.length} rows`);
  } else {
    Logger.log(`[GISTDA_Hotspot_Lib] Data not found`);
  }
}

// Convert index (0, 1, 2) to character (A, B, C)
function _getColumnLetter(index) {
  if (index < 0) {
    return null;
  }

  let temp, letter = '';
  while (index >= 0) {
    temp = (index) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    index = Math.floor((index) / 26) - 1;
  }
  return letter;
}

function _createSemanticSheet(ss, sheetName, queryString) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();
  sheet.getRange("A1").setFormula(queryString);
}

function _generateSemanticTables(ss) {
  Logger.log("[GISTDA_Hotspot_Lib] Semantic starts...");

  const sourceSheet = ss.getSheetByName(CONFIG.SHEET_RAW);
  if (!sourceSheet) {
    throw new Error("[GISTDA_Hotspot_Lib] Semantic: Data sheet not found");
  }

  // 1. Find mandatory columns
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const colCt = _getColumnLetter(headers.indexOf('ct_en'));   // Country
  const colPv = _getColumnLetter(headers.indexOf('pv_tn'));   // Province (TH)
  const colLu = _getColumnLetter(headers.indexOf('lu_name')); // Land Use

  if (!colCt || !colPv || !colLu) {
    throw new Error("[GISTDA_Hotspot_Lib] Semantic Error: Mandatory columns (ct_en, pv_tn, lu_name) not found");
  }

  const rawName = CONFIG.SHEET_RAW;

  // --- Semantic 1: Group by country (excludes China) ---
  const qCountry = `=QUERY('${rawName}'!A:ZZ, "SELECT ${colCt}, COUNT(${colCt}) WHERE ${colCt} IS NOT NULL AND ${colCt} != 'China' GROUP BY ${colCt} ORDER BY COUNT(${colCt}) DESC LABEL ${colCt} 'Country', COUNT(${colCt}) 'Hotspots'", 1)`;
  _createSemanticSheet(ss, CONFIG.SHEET_COUNTRY, qCountry);

  // --- Semantic 2: Group by province (only Thailand) ---
  const qProvince = `=QUERY('${rawName}'!A:ZZ, "SELECT ${colPv}, COUNT(${colPv}) WHERE ${colCt} = 'Thailand' AND ${colPv} IS NOT NULL GROUP BY ${colPv} ORDER BY COUNT(${colPv}) DESC LABEL ${colPv} 'Province (TH)', COUNT(${colPv}) 'Hotspots'", 1)`;
  _createSemanticSheet(ss, CONFIG.SHEET_PROVINCE, qProvince);

  // --- Semantic 3: Group by Land Use (only Thailand) ---
  const qLanduse = `=QUERY('${rawName}'!A:ZZ, "SELECT ${colLu}, COUNT(${colLu}) WHERE ${colCt} = 'Thailand' AND ${colLu} IS NOT NULL GROUP BY ${colLu} ORDER BY COUNT(${colLu}) DESC LABEL ${colLu} 'Land Use', COUNT(${colLu}) 'Hotspots'", 1)`;
  _createSemanticSheet(ss, CONFIG.SHEET_LANDUSE, qLanduse);

  Logger.log("[GISTDA_Hotspot_Lib] Semantic finished");
}
