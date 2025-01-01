class SetupUtils {
    static firstUsed() {
        Logger.log("複製下面的文字");
        Logger.log(
      `var _calendar_id = "更改為自己的 calendarID!!!";
      
      function addEvent() {
        _.main("addEvent", _calendar_id);  
      }
      
      function addCustomMenu() {
        _.main("addCustomMenu", _calendar_id);
      }
      
      function logSheetContent(){
        _.main("logSheetContent", _calendar_id);
      }
      `);
        Logger.log(
      `記得改 calendarID!!!
      記得改 calendarID!!!
      記得改 calendarID!!!`
      );
        Logger.log(
      `更改好 calendarID 後，\n
      1. 按下儲存 (在 + 號右邊3格)
      2. 選擇 addCustomMenu (在偵錯的右邊下拉式選單中)
      3. 點選執行        
      `);
      }
  }