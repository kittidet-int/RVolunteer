function importJSONtoSheet() {
  // 1. API
  var url = "http://air4thai.pcd.go.th/services/getNewAQI_JSON.php";

  // 2. Retrieve data from API
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText();
  var data = JSON.parse(json);

  if (!data || data.length === 0) {
    Logger.log("[PCD_AirQuality_Lib] Data now found");
    return;
  }

  // 3. Store data to sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  sheet.clear();

  // ดึง Header (ชื่อคอลัมน์) จากข้อมูลชุดแรก
  var headers = [
    "Station ID",
    "Station Name (TH)",
    "Station Name (EN)",
    "Station Area (TH)",
    "Date",
    "Time",
    "PM2.5 (Value)",
    "PM2.5 (AQI)",
    "PM2.5 (Color)"
  ];

  var rows = [headers];

  if (data.stations && data.stations.length > 0) {
    data.stations.forEach(function (st) {

      var l = st.AQILast;

      rows.push([
        st.stationID,
        st.nameTH,
        st.nameEN,
        st.areaTH,
        l.date,
        l.time,
        (l.PM25 ? Number(l.PM25.value) : ""),
        (l.PM25 ? l.PM25.aqi : ""),
        (l.PM25 ? l.PM25.color_id : "0")
      ]);
    });
  }

  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }

  Logger.log("[PCD_AirQuality_Lib] Finished");
}
