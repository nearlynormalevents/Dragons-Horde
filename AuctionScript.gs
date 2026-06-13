// ================================================================
// SILENT AUCTION — Google Apps Script
// Deploy as a Web App (Anyone can access) and paste the URL
// into auction.html where it says YOUR_GOOGLE_APPS_SCRIPT_URL_HERE
// ================================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Open (or create) the sheet named "Bids"
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Bids");

    if (!sheet) {
      sheet = ss.insertSheet("Bids");
      // Write header row on first run
      sheet.appendRow([
        "Timestamp",
        "Bidder Name",
        "Phone Number",
        "Auction Item",
        "Minimum Bid ($)",
        "Bid Amount ($)"
      ]);

      // Style the header row
      var header = sheet.getRange(1, 1, 1, 6);
      header.setBackground("#1a1a2e");
      header.setFontColor("#ffffff");
      header.setFontWeight("bold");
      header.setFontSize(11);

      // Freeze header
      sheet.setFrozenRows(1);

      // Set column widths
      sheet.setColumnWidth(1, 160);  // Timestamp
      sheet.setColumnWidth(2, 160);  // Name
      sheet.setColumnWidth(3, 140);  // Phone
      sheet.setColumnWidth(4, 340);  // Item
      sheet.setColumnWidth(5, 120);  // Min Bid
      sheet.setColumnWidth(6, 120);  // Bid Amount
    }

    // Append the new bid
    sheet.appendRow([
      data.timestamp   || new Date().toLocaleString(),
      data.name        || "",
      data.phone       || "",
      data.item        || "",
      data.minBid      || "",
      data.bidAmount   || ""
    ]);

    // Auto-resize for readability
    sheet.autoResizeColumns(1, 6);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: GET handler so the script URL can be tested in a browser
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "alive", message: "Silent Auction Script is running." }))
    .setMimeType(ContentService.MimeType.JSON);
}
