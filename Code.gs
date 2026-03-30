function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Return empty array if sheet has only headers or is empty
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
  
  // Dynamically find column indices based on header names
  var storeIdx    = headers.findIndex(function(h) { return h.indexOf("store") !== -1; });
  var brandIdx    = headers.findIndex(function(h) { return h.indexOf("brand") !== -1; });
  var natureIdx   = headers.findIndex(function(h) { return h.indexOf("nature") !== -1; });
  var positionIdx = headers.findIndex(function(h) { return h.indexOf("position") !== -1; });
  var notesIdx    = headers.findIndex(function(h) { return h.indexOf("notes") !== -1 || h.indexOf("remarks") !== -1; });

  var results = [];
  
  // Loop through rows and build the JSON array our frontend expects
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    if (row[storeIdx]) {
      results.push({
        store_name:      row[storeIdx] ? row[storeIdx].toString() : "",
        brand:           row[brandIdx] ? row[brandIdx].toString() : "",
        shelft_nature:   row[natureIdx] ? row[natureIdx].toString() : "",
        shelft_position: row[positionIdx] ? row[positionIdx].toString() : "",
        notes:           notesIdx !== -1 && row[notesIdx] ? row[notesIdx].toString() : ""
      });
    }
  }
  
  // We return a flat array mapping to the DB_DATA object.
  // The frontend selectStore() function handles grouping into 
  // Store -> Shelf Nature -> Brands -> Position efficiently.
  return ContentService.createTextOutput(JSON.stringify(results))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var storeName      = e.parameter.store_name || "";
  var brand          = e.parameter.brand || "";
  var shelftNature   = e.parameter.shelft_nature || "";
  var shelftPosition = e.parameter.shelft_position || "";
  var notes          = e.parameter.notes || "";
  
  // Append new data to the bottom of the Sheet
  sheet.appendRow([
    new Date(), // Timestamp
    storeName,
    brand,
    shelftNature,
    shelftPosition,
    notes
  ]);
  
  // Return success payload seamlessly
  return ContentService.createTextOutput(JSON.stringify({ "status": "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
