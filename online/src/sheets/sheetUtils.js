class SheetUtils {
    static logSheetContent() {
      const sheetName = "recordEventID";
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        Logger.log(`Sheet "${sheetName}" not found`);
        return;
      }
  
      const lastRow = sheet.getLastRow();
      const dataRange = sheet.getRange(1, 1, lastRow, 2);
      const values = dataRange.getValues();
      
      Logger.log("Sheet Content (recordEventID):");
      Logger.log("Row | Create Time | Event ID");
      Logger.log("-".repeat(30));
      
      values.forEach((row, index) => {
        Logger.log(`R${index + 1} | ${row[0]} | ${row[1]}`);
      });
      
      Logger.log("-".repeat(30));
      Logger.log(`Total Rows: ${lastRow}`);
    }
  }