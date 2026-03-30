function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });
  
  var dateIdx     = headers.findIndex(function(h) { return h.indexOf("date") !== -1 || h.indexOf("time") !== -1; });
  var storeIdx    = headers.findIndex(function(h) { return h.indexOf("store") !== -1; });
  var brandIdx    = headers.findIndex(function(h) { return h.indexOf("brand") !== -1; });
  var natureIdx   = headers.findIndex(function(h) { return h.indexOf("nature") !== -1; });
  var positionIdx = headers.findIndex(function(h) { return h.indexOf("position") !== -1; });
  var notesIdx    = headers.findIndex(function(h) { return h.indexOf("notes") !== -1; });

  var results = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Only capture rows where there is at least a store name or brand
    if (storeIdx !== -1 && row[storeIdx]) {
      
      // Attempt to format date cleanly if it exists
      var dateStr = "";
      if (dateIdx !== -1 && row[dateIdx]) {
        var d = new Date(row[dateIdx]);
        if (!isNaN(d.getTime())) {
          dateStr = d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        } else {
          dateStr = row[dateIdx].toString(); // fallback
        }
      }

      results.push({
        date:            dateStr,
        store_name:      row[storeIdx] ? row[storeIdx].toString() : "",
        brand:           brandIdx !== -1 && row[brandIdx] ? row[brandIdx].toString() : "",
        shelft_nature:   natureIdx !== -1 && row[natureIdx] ? row[natureIdx].toString() : "",
        shelft_position: positionIdx !== -1 && row[positionIdx] ? row[positionIdx].toString() : "",
        notes:           notesIdx !== -1 && row[notesIdx] ? row[notesIdx].toString() : ""
      });
    }
  }
  
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
  
  // Appends exactly: Date, Store, Brand, Nature, Position, Notes
  sheet.appendRow([
    new Date(), 
    storeName,
    brand,
    shelftNature,
    shelftPosition,
    notes
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({ "status": "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
