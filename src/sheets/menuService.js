class MenuService {
  static addCustomMenu(_calendar_id) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const recordService = new RecordService();

    const sheet = spreadsheet.getSheetByName("recordEventID");
    if(!sheet){
      const sheet = spreadsheet.insertSheet("recordEventID");
      sheet.hideSheet();
      Logger.log("recordEventID set!")
    }else{
      sheet.hideSheet();
      Logger.log("recordEventID exist!");
    }

    var ss = recordService.findSheetWithTimestamp();

    var ui = SpreadsheetApp.getUi();
    ui.createMenu('手打擴充功能')
      .addItem('更新行事曆', 'addEvent')
      .addToUi();
    
    const eventService = new EventService(_calendar_id);
    eventService.addEvent(); // 移除 ss 參數
  }
}

