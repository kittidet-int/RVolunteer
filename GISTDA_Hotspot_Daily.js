/**
 * Application: GISTDA Hotspot Daily
 * Description: A Google Apps Script to retreive VIIRS/MODIS hotspot data.
 * * Copyright (C) 2026  Kittidet Intharaksa
 * * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

// ==========================================
// 1. CONFIGURATION
// ==========================================
const CONFIG = {
  SHEET_DAILY: 'Daily',
  SHEET_COUNTRY: 'Semantic_Country',
  SHEET_PROVINCE: 'Semantic_Province_TH',
  SHEET_LANDUSE: 'Semantic_Landuse_TH',
};

// ==========================================
// 2. MAIN PROCESS
// ==========================================

function runDailyProcess() {
  const TARGET_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('TARGET_FOLDER_ID');

  if (!TARGET_FOLDER_ID) {
    throw new Error("Could not find Folder ID in properties");
  }

  const today = new Date();
  // Data date (T-1)
  // const dataDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
  // Archive date (T-2)
  const archiveDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 2);
  // const archiveDateStr = Utilities.formatDate(archiveDate, Session.getScriptTimeZone(), "yyyyMMdd");

  Logger.log(`[Start] Run Daily Process: ${Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd")}`);

  try {
    const spreadsheet = GISTDA_Hotspot_Lib.executePipeline(TARGET_FOLDER_ID);
    sendSemanticToLine(spreadsheet);

    Logger.log("[Success]");
  } catch (e) {
    throw new Error(e.message);
  }
}

function runTestProcess() {
  const TEST_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('TEST_FOLDER_ID');

  if (!TEST_FOLDER_ID) {
    throw new Error("Could not find Folder ID in properties");
  }

  Logger.log(`[TEST] Start...`);

  try {
    const spreadsheet = GISTDA_Hotspot_Lib.executePipeline(TEST_FOLDER_ID, true);
    sendSemanticToLine(spreadsheet, true);

    Logger.log("[TEST] Success");
  } catch (e) {
    throw Error(`[TEST] ${e.message}`);
  }
}

// ==========================================
// LINE message
// ==========================================

function getTopRankingData(ss, sheetName, limit) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, Math.min(limit, lastRow - 1), 2).getValues();

  return data;
}

function getTopRankingFlexMessage(ss, sheetName, title, dataName, limit) {
  const data = getTopRankingData(ss, sheetName, limit);

  let message = {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "box",
        "layout": "vertical",
        "contents": [
          {
            "type": "text",
            "text": title,
            "weight": "bold",
            "color": "#3C00F0"
          }
        ],
        "paddingStart": "8px",
        "paddingEnd": "8px"
      }
    ],
    "paddingEnd": "8px",
    "paddingStart": "0px",
    "paddingTop": "8px",
    "paddingBottom": "0px"
  };

  if (data) {
    message.contents.push({
      "type": "box",
      "layout": "horizontal",
      "contents": [
        {
          "type": "text",
          "text": "ลำดับ",
          "flex": 2
        },
        {
          "type": "text",
          "text": dataName,
          "flex": 6
        },
        {
          "type": "text",
          "text": "จำนวน",
          "flex": 2,
          "align": "end"
        }
      ],
      "backgroundColor": "#f0f0f0",
      "paddingStart": "8px",
      "paddingEnd": "8px"
    });

    data.forEach((row, index) => {
      message.contents.push({
        "type": "box",
        "layout": "horizontal",
        "contents": [
          {
            "type": "text",
            "text": `${index + 1}`,
            "flex": 2,
            "align": "end",
            "offsetEnd": "24px"
          },
          {
            "type": "text",
            "text": `${row[0]}`,
            "flex": 6
          },
          {
            "type": "text",
            "text": `${row[1].toLocaleString()}`,
            "flex": 2,
            "align": "end"
          }
        ],
        "paddingStart": "8px",
        "paddingEnd": "8px"
      });
    });
  }
  else {
    message.contents.push({
      "type": "box",
      "layout": "vertical",
      "contents": [
        {
          "type": "text",
          "text": "- ไม่พบข้อมูล -"
        }
      ],
      "paddingStart": "8px",
      "paddingEnd": "8px"
    });
  }

  message.contents.push({
    "type": "separator"
  });

  return message;
}

function sendSemanticToLine(ss, isTesting = false) {
  Logger.log("[Line] กำลังเตรียมข้อมูลส่งไลน์ (Messaging API)...");

  const today = new Date();
  const dataDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
  const dataDateStr = Utilities.formatDate(dataDate, Session.getScriptTimeZone(), "d MMM yyyy");

  const countryMessage = getTopRankingFlexMessage(ss, CONFIG.SHEET_COUNTRY, "ประเทศ สูงสุด 5 อันดับ", "ประเทศ", 5);
  const provinceMessage = getTopRankingFlexMessage(ss, CONFIG.SHEET_PROVINCE, "จังหวัด สูงสุด 5 อันดับ", "จังหวัด", 5);
  const landUseMessage = getTopRankingFlexMessage(ss, CONFIG.SHEET_LANDUSE, "ประเภทพื้นที่", "ประเภทพื้นที่", 99);

  let header = {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "รายงานจุดความร้อน",
        "weight": "bold"
      },
      {
        "type": "box",
        "layout": "horizontal",
        "contents": [
          {
            "type": "text",
            "text": "วันที่",
            "flex": 1
          },
          {
            "type": "text",
            "text": dataDateStr,
            "flex": 5,
            "weight": "bold",
            "color": "#0000FF"
          }
        ]
      }
    ],
    "paddingStart": "8px",
    "paddingEnd": "8px",
    "paddingTop": "10px",
    "paddingBottom": "8px"
  };

  let dataContents = [];

  // dataContents.push({
  //   "type": "box",
  //   "layout": "vertical",
  //   "contents": [
  //     {
  //       "type": "text",
  //       "text": "จุดความร้อน (Hotspot)",
  //       "weight": "bold"
  //     }
  //   ],
  //   "backgroundColor": "#FCC6BB",
  //   "paddingStart": "8px",
  //   "paddingEnd": "8px",
  //   "paddingTop": "4px",
  //   "paddingBottom": "4px"
  // });

  dataContents.push(countryMessage);
  dataContents.push(provinceMessage);
  dataContents.push(landUseMessage);

  let bodyContents = [
    {
      "type": "box",
      "layout": "vertical",
      "contents": dataContents,
      "paddingAll": "0px"
    }
  ];

  let body = {
    "type": "box",
    "layout": "vertical",
    "contents": bodyContents,
    "paddingTop": "8px",
    "paddingBottom": "0px",
    "paddingStart": "0px",
    "paddingEnd": "0px"
  };

  let footer = {
    "type": "box",
    "layout": "vertical",
    "contents": [
      {
        "type": "text",
        "text": "จัดทำรายงานโดย",
        "weight": "bold"
      },
      {
        "type": "text",
        "text": "นายกิตติเดช  อินทรักษา"
      },
      {
        "type": "text",
        "text": "ประจำกลุ่มงานจิตอาสาภัยพิบัติ"
      },
      {
        "type": "text",
        "text": "ศอญ.จอส.พระราชทาน ปี 2569"
      }
    ],
    "paddingStart": "8px",
    "paddingEnd": "8px"
  };

  let message = {
    "type": "bubble",
    "header": header,
    "body": body,
    "footer": footer,
    "styles": {
      "header": {
        "backgroundColor": "#B5FF70"
      }
    }
  };

  if (isTesting) {
    LINE_Utils_Lib.sendMessageToTest('flex', `รายงานจุดความร้อน ${dataDateStr}`, message);
  } else {
    LINE_Utils_Lib.sendMessageToGroup('flex', `รายงานจุดความร้อน ${dataDateStr}`, message);
  }
}
