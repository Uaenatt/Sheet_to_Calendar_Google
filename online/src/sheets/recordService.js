
class RecordService {
  constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  findSheetWithTimestamp() {
    const sheets = this.spreadsheet.getSheets(); // Get all sheets

    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const a1Value = sheet.getRange("A1").getValue(); // Get the value of A1

      if (a1Value === "時間戳記") {
        Logger.log("Sheet found: " + sheet.getName());
        return sheet; // Return the matching sheet
      }
    }

    Logger.log("No sheet found with '時間戳記' in A1.");
    return null; // Return null if no sheet matches
  }

  recordOrUpdateEventInfo(createTime, eventID) {
    const sheetName = "recordEventID";
    let sheet = this.spreadsheet.getSheetByName(sheetName);

    // Create the sheet if it doesn't exist and add headers in row 2
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet(sheetName);
      sheet.getRange("A2:B2")
        .setValues([["Create Time", "Event ID"]])
        .setFontWeight("bold");
      sheet.setFrozenRows(2); // Freeze including the empty first row
    }

    // Get data starting from row 3 (after empty row and header)
    const dataRange = sheet.getRange(3, 1, Math.max(1, sheet.getLastRow() - 2), 2);
    const values = dataRange.getValues();
    
    // Create a lookup map for better performance
    const createTimeColumn = values.map(row => row[0].toString());
    const rowIndex = createTimeColumn.indexOf(createTime.toString());

    if (rowIndex >= 0) { // Found existing entry
      // Update only the eventID in the found row (add 3 because we start data from row 3)
      sheet.getRange(rowIndex + 3, 2).setValue(eventID);
      Logger.log(`Updated eventID for createTime: ${createTime}`);
    } else {
      // Append new row if createTime not found
      const lastRow = Math.max(2, sheet.getLastRow());
      sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[createTime, eventID]]);
      Logger.log(`Added new record: createTime = ${createTime}, eventID = ${eventID}`);
    }

    // Optional: Autosize columns for better readability
    sheet.autoResizeColumns(1, 2);
  }

  getEventIDByCreateTime(createTime) {
    const sheetName = "recordEventID";
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    
    // If sheet doesn't exist, return 0
    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found`);
      return 0;
    }

    // Get data starting from row 3 (after empty row and header)
    const dataRange = sheet.getRange(3, 1, Math.max(1, sheet.getLastRow() - 2), 2);
    const values = dataRange.getValues();
    
    // Look for matching createTime
    const createTimeColumn = values.map(row => row[0].toString());
    const rowIndex = createTimeColumn.indexOf(createTime.toString());

    if (rowIndex >= 0) {
      const eventID = values[rowIndex][1];
      Logger.log(`Found eventID: ${eventID} for createTime: ${createTime}`);
      return eventID;
    }

    Logger.log(`No eventID found for createTime: ${createTime}`);
    return 0;
  }
}
