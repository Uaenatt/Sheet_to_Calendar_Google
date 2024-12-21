const errorHandlerState = {
    eventTitles: [],
    errorMessages: [],
    errorRow: [],
    errorColumn: [],
    errorStacks: []  // 新增錯誤堆疊追蹤
};

class ErrorHandler {
    constructor() {
        this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }

    static chr(n) {
        return String.fromCharCode(n);
    }

    static handleError(error, eventTitle, row, column) {
        Logger.log(`Error in event "${eventTitle}": ${error.message}`);
        Logger.log("Error stack: " + error.stack);  // 記錄錯誤堆疊

        errorHandlerState.eventTitles.push(eventTitle);
        errorHandlerState.errorMessages.push(error.message);
        errorHandlerState.errorRow.push(row + 1);
        errorHandlerState.errorColumn.push(this.chr(column + 65));
        errorHandlerState.errorStacks.push(error.stack);  // 儲存錯誤堆疊

        // 設定錯誤儲存格背景顏色
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const cell = sheet.getRange(row + 1, column + 1); // +2 因為要跳過標題列，+1 因為欄位從1開始
        cell.setBackground('#FF0000'); // 設定為紅色
    }

    static overAllError() {
        Logger.log("錯誤狀態: " + JSON.stringify(errorHandlerState));
        
        for (var i = 0; i < errorHandlerState.eventTitles.length; i++) {
            Logger.log(`錯誤 #${i + 1}:
            活動名稱: ${errorHandlerState.eventTitles[i]}
            錯誤訊息: ${errorHandlerState.errorMessages[i]}
            位置: 第 ${errorHandlerState.errorRow[i]} 列第 ${errorHandlerState.errorColumn[i]} 欄
            錯誤堆疊: ${errorHandlerState.errorStacks[i]}`);
        }

        var errorMessage = "";
        for (var i = 0; i < errorHandlerState.eventTitles.length; i++) {
            errorMessage += `${errorHandlerState.eventTitles[i]} 發生錯誤在（ ${errorHandlerState.errorRow[i]} , ${errorHandlerState.errorColumn[i]} ）`;
        }
        errorMessage += "如若更正後仍出現錯誤，請聯繫管理人。";
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMessage, "錯誤", 30);

        return {
            titles: errorHandlerState.eventTitles,
            messages: errorHandlerState.errorMessages,
            rows: errorHandlerState.errorRow,
            columns: errorHandlerState.errorColumn,
            fullMessage: errorMessage
        };
    }

    static clearErrors() {
        // 清除所有標記的儲存格背景顏色
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        for (let i = 0; i < errorHandlerState.errorRow.length; i++) {
            const cell = sheet.getRange(errorHandlerState.errorRow[i] + 1, 
                                      errorHandlerState.errorColumn[i].charCodeAt(0) - 64); // 將A-Z轉換為1-26
            cell.setBackground(null); // 清除背景顏色
        }

        errorHandlerState.eventTitles = [];
        errorHandlerState.errorMessages = [];
        errorHandlerState.errorRow = [];
        errorHandlerState.errorColumn = [];
        errorHandlerState.errorStacks = [];
    }
}