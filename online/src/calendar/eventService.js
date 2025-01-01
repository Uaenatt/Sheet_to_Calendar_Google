class EventService {
    constructor(calendarId) {
        this.calendarId = calendarId;
        this.recordService = new RecordService();
    }

    addEvent() {
        var updateCorrectly = true;
        SpreadsheetApp.getActiveSpreadsheet().toast("行事曆更新開始", "提示", 1.5);
        const ss = this.recordService.findSheetWithTimestamp();

        const data = ss.getRange("A1:Z" + ss.getLastRow()).getValues();
        const now = new Date();

        var columnH = ss.getRange(1, 8, ss.getMaxRows()); // 第 8 欄
        columnH.setNumberFormat('@'); // 設為純文字格式

        var columnI = ss.getRange(1, 9, ss.getMaxRows()); // 第 9 欄
        columnI.setNumberFormat('@'); // 設為純文字格式

        // Loop through the data starting from the second row
        for (let i = 1; i < data.length; i++) {
            Logger.log("Processing activity: " + data[i][2]); // 活動名稱
            try {
                const createTime = data[i][0]; // 時間戳記

                // 如果 createTime 為空，則往下找到不為空的 createTime
                if (createTime === '') {
                    var blank = 0;
                    while (data[i + blank][0] === '') {
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
                    if(parsedRange === null) {
                        ErrorHandler.handleError(new Error("Invalid date range"), data[i][2]);  // 傳遞活動名稱
                        continue;
                    }
                    
                    start_dateTime = new Date(parsedRange.start);
                    end_dateTime = new Date(parsedRange.end);
                    // 詳細資訊包含借人員
                    var _description = discription_creater_displayName + '\n' +
                        discription_createrEmailAndLineID + '\n' +
                        "租用人員: " + data[i][10];
                }

                
                Logger.log("now: " + now);
                Logger.log("start_dateTime: " + start_dateTime);
                Logger.log("end_dateTime: " + end_dateTime);
            
                // 如果開始時間小於現在時間加一天，則跳過
                var start_dateTime_for_comparing = new Date(start_dateTime);
                start_dateTime_for_comparing = start_dateTime_for_comparing.setDate(start_dateTime.getDate() + 1);
                if (start_dateTime_for_comparing < now) {
                    Logger.log("Skipped event: " + _title + " (past)");
                    continue;
                }


                // 使用 this.recordService 來調用方法
                const eventId = this.recordService.getEventIDByCreateTime(createTime);
                if (eventId) {
                    const existingEvent = Calendar.Events.get(this.calendarId, eventId);

                    if (
                        existingEvent.summary === _title &&
                        existingEvent.location === _location &&
                        existingEvent.description === _description
                    ) {
                        Logger.log("Skipped event: " + _title + " (all same)");
                        continue;
                    }

                    Calendar.Events.remove(this.calendarId, eventId);
                    Logger.log("Deleted event: " + _title);
                }
            
                // 創建新事件
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

                const newEvent = Calendar.Events.insert(eventData, this.calendarId);
                this.recordService.recordOrUpdateEventInfo(createTime, newEvent.id);
                Logger.log("Created new event: " + _title);

            } catch (e) {
                updateCorrectly = false;
                ErrorHandler.handleError(e, data[i][2]);  // 傳遞活動名稱
            }
            
        }
        if(updateCorrectly === true){ // 如果更新成功
            SpreadsheetApp.getActiveSpreadsheet().toast("行事曆已更新", "提示", 1.5);
        }else{
            ErrorHandler.overAllError();
        }
    }
}