function handleError(error) {
  Logger.log(`Error: ${error.message}`);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `發生錯誤: ${error.message}`, 
    "錯誤", 
    5
  );
}

function main(input, _calendar_id) {
  try {
    const calendarService = new CalendarService(_calendar_id);
    
    switch(input) {
      case "addEvent":
        addEvent(_calendar_id);
        break;
      case "addCustomMenu":
        addCustomMenu(_calendar_id);
        break;
      case "logSheetContent":
        logSheetContent();
        break;
      case "firstUsed":
        firstUsed();
        break;
      default:
        throw new Error("Function not found");
    }
  } catch(error) {
    handleError(error);
    throw error;
  }
}

// 第一次使用 給予使用者程式碼以及提醒事項
function firstUsed() {
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

// Function to handle matching events without time
function handleMatchingEventNoTime(calId, _title, _location, _description) {
  // Search the calendar for events with the same title
  const events = Calendar.Events.list(calId, {
    timeMin: "2000-01-01T00:00:00Z", // Start date for searching
    timeMax: "2100-01-01T00:00:00Z", // End date for searching
    q: _title, // Search query
  }).items;

  if (!events) {
    Logger.log("No events found matching the query.");
    return 0;
  }

  // Loop through found events to check for a full match
  for (let i = 0; i < events.length; i++) {
    const event = events[i];

    if (
      event.summary === _title &&
      event.location === _location &&
      event.description === _description
    ) {
      return event.id;
    }
  }
  return 0;
}

// Function to handle matching events with time
function eventIDEmpty(createTime, cal, _title, _location, _description, start_dateTime, end_dateTime, ss, i) {
  eventID = handleMatchingEventNoTime(cal, _title, _location, _description, start_dateTime, end_dateTime);
  // 如果eventID != 0
  if(eventID != 0){
    recordOrUpdateEventInfo(createTime, eventID);  // Column Q for Event ID
    return 1;
  }
  return 0;
}

// Function to record or update event info in "recordEventID" sheet
function recordOrUpdateEventInfo(createTime, eventID) {
  const recordService = new RecordService();
  return recordService.recordOrUpdateEventInfo(createTime, eventID);
}

// Function to get event ID by createTime
function getEventIDByCreateTime(createTime) {
  const recordService = new RecordService();
  return recordService.getEventIDByCreateTime(createTime);
}

// Function to add events to a Google Calendar
function addEvent(_calendar_id) {
  SpreadsheetApp.getActiveSpreadsheet().toast("行事曆更新開始", "提示", 1.5);
  const recordService = new RecordService();
  const ss = recordService.findSheetWithTimestamp();

  const data = ss.getRange("A1:Z" + ss.getLastRow()).getValues();
  const now = new Date();

  var columnH = ss.getRange(1, 8, ss.getMaxRows()); // 第 8 欄
  columnH.setNumberFormat('@'); // 設為純文字格式

  var columnI = ss.getRange(1, 9, ss.getMaxRows()); // 第 9 欄
  columnI.setNumberFormat('@'); // 設為純文字格式

  // Loop through the data starting from the second row
  for (let i = 1; i < data.length; i++) {
    const createTime = data[i][0]; // 時間戳記

    // 如果 createTime 為空，則往下找到不為空的 createTime
    if(createTime === '') {
        var blank = 0;
        while(data[i + blank][0] === '') {
          blank++;
        }
        var rangeToCut = ss.getRange(i + blank + 2, 1, 1, 17);
        var rowValues = rangeToCut.getValues();
        var rangeToPaste = ss.getRange(i + 2, 1, 1, 17);
        rangeToPaste.setValues(rowValues);
        rangeToCut.clearContent();
        i -= blank
        data = ss.getRange("A1:Z" + ss.getLastRow()).getValues();
    }

    const _title = data[i][2]; // 活動名稱
    const _location = data[i][4]; // 活動地點
    var discription_creater_displayName = "活動單位: " + data[i][1]; // 辦理活動單位
    var discription_createrEmailAndLineID = "名稱: " + data[i][12] +  // 聯絡人
                                            "\nLine ID: " + data[i][13] + // 聯絡人 Line ID
                                            "\ngmail: " + data[i][14]; // 聯絡人 email
    var discription_leasingEquipOrHuman = data[i][5]; // 借器材 or 人
    
    
    var start_dateTime, end_dateTime;
    // 如果是借器材
    if (discription_leasingEquipOrHuman === "器材") {
      start_dateTime = new Date(data[i][3]); // For all-day events
      // 結束時間設為開始時間的隔天
      end_dateTime = new Date(start_dateTime);
      end_dateTime.setDate(end_dateTime.getDate() + 1);

      // 詳細資訊包含器材名稱、借器材時間、還器材時間、借還器材地點
      var _description = discription_creater_displayName + '\n' +
                        discription_createrEmailAndLineID + '\n' +
                        "器材: " + data[i][6] + '\n' +
                        "借器材時間: " + data[i][7] + '\n' +
                        "還器材時間: " + data[i][8] + '\n' +
                        "借還器材地點: " + data[i][9];
    } else { // 如果是借人
      // 解析日期範圍
      var parsedRange = DateUtils.parseDateRangeWithYear(data[i][11], new Date(data[i][3])); // For ranged events
      start_dateTime = new Date(parsedRange.start);
      end_dateTime = new Date(parsedRange.end);
      // 詳細資訊包含借人員
      var _description = discription_creater_displayName + '\n' +
                        discription_createrEmailAndLineID + '\n' +
                        "租用人員: " + data[i][10];
    }

    // 如果開始時間小於現在時間，則跳過
    if (start_dateTime < now) {
      Logger.log("Skipped event: " + _title + " (past)");
      continue;
    }
    
    // --> 這裡開始新增
    const eventId = getEventIDByCreateTime(createTime);
    if (eventId) {
      try {
        const existingEvent = Calendar.Events.get(_calendar_id, eventId);

        if (
          existingEvent.summary === _title &&
          existingEvent.location === _location &&
          existingEvent.description === _description
        ) {
          Logger.log("Skipped event: " + _title + " (all same)");
          continue;
        }

        Calendar.Events.remove(_calendar_id, eventId);
        Logger.log("Deleted event: " + _title);
      } catch (e) {
        Logger.log("Error deleting event: " + e.message);
      }
    }

    const eventData = {
      summary: _title,
      location: _location,
      description: _description,
      start: {
        dateTime: start_dateTime.toISOString(),
        timeZone: "Asia/Taipei",
      },
      end: {
        dateTime: end_dateTime.toISOString(),
        timeZone: "Asia/Taipei",
      },
    };

    const newEvent = Calendar.Events.insert(eventData, _calendar_id);
    recordOrUpdateEventInfo(createTime, newEvent.id);
    Logger.log("Created new event: " + _title);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("行事曆已更新", "提示", 1.5);
}

function addCustomMenu(_calendar_id) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const recordService = new RecordService();

  const sheet = spreadsheet.getSheetByName("recordEventID");
  if(!sheet){
    const sheet = spreadsheet.insertSheet("recordEventID"); // Create a new sheet with a specific name
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
  addEvent(_calendar_id);
}

function logSheetContent() {
  const sheetName = "recordEventID";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  // Check if sheet exists
  if (!sheet) {
    Logger.log(`Sheet "${sheetName}" not found`);
    return;
  }

  // Get all data from the sheet
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(1, 1, lastRow, 2);
  const values = dataRange.getValues();
  
  // Log the header showing row numbers
  Logger.log("Sheet Content (recordEventID):");
  Logger.log("Row | Create Time | Event ID");
  Logger.log("-".repeat(30));
  
  // Log each row with row number
  values.forEach((row, index) => {
    Logger.log(`R${index + 1} | ${row[0]} | ${row[1]}`);
  });
  
  Logger.log("-".repeat(30));
  Logger.log(`Total Rows: ${lastRow}`);
}

function protectAndHideEventIdColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // Define the range and column to protect/hide
  var rangeToProtect = sheet.getRange("A:A");
  var columnIndex = 17; // Column Q
  
  // Check if already protected
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var isProtected = protections.some(p => p.getRange().getA1Notation() === rangeToProtect.getA1Notation());

  if (!isProtected) {
    var protection = rangeToProtect.protect();
    protection.setDescription("Event ID column is protected from editing.");
  
    var me = Session.getEffectiveUser();
    protection.addEditor(me); // Allow script owner to modify
    protection.removeEditors(protection.getEditors().filter(user => user.getEmail() !== me.getEmail()));
  
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  
    Logger.log("Protection applied to range: " + rangeToProtect.getA1Notation());
  } else {
    Logger.log("Range already protected: " + rangeToProtect.getA1Notation());
  }
}
