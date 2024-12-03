class ErrorHandler {
    static handleError(error) {
      Logger.log(`Error: ${error.message}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `發生錯誤: ${error.message}`, 
        "錯誤", 
        5
      );
    }
  }