/** 取得日曆 */
function getCalender(calenderID) {
  var getCal = CalendarApp.getCalendarById(calenderID);
  /** 若想要使用浮動取得Calender的方式，也可以使用以下方式 */
  // var calendarID = spreadsheet.getRange("在GoogleSheets資料表中的位置").getValue();
  // var eventCal = CalendarApp.getCalendarById(calendarID); 

  return getCal
}

function scheduleShifts() {
  
  eventCal = getCalender("輸入Google Calender ID")
  /** 取得所有資料的位置 */
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var signups = spreadsheet.getRange("輸入Sheets的資料範圍").getValues();
  /** 開始新增資料 */
  for (x=0; x<signups.length; x++) {
    var shift = signups[x];

    if (shift[1] != "") { // 若標題是空的，就不執行那行，節省計算資源

      var title = shift[0]; // 行程標題
      var theday = shift[1]; // 日期
      var startTime = shift[2]; // 開始時間
      var endTime = shift[3]; // 結束時間
      var location = shift[4]; // 地點
      var descripion = shift[5]; // 備註
      theday.setDate(theday.getDate() + 1) // 必須加一天

      /** 檢查title是不是有重複 */
      var doit = true;
      getTodayEvents = eventCal.getEventsForDay(theday);
      for (ch=0; ch<getTodayEvents.length; ch++){
        if (getTodayEvents[ch].getTitle() == title){
          doit = false; // 發現重複，改成flase
          Logger.log("有重複，不執行")
        }
      }

      /** 沒有重複，創建事件 */
      if (doit) {
        
        var start = new Date(
                              theday.getFullYear(), 
                              theday.getMonth(), 
                              theday.getDate(), 
                              startTime.getHours(), 
                              startTime.getMinutes()
                            )

        var end = new Date(
                            theday.getFullYear(), 
                            theday.getMonth(), 
                            theday.getDate(), 
                            endTime.getHours(), 
                            endTime.getMinutes()
                          )

        /** 為了避免跨日問題，可能開始日期會大於結束日期*/
        if (start.valueOf()  > end.valueOf()){
          start.setDate(start.getDate() - 1)
          Logger.log("跨日問題")
        }

        /** 建立行程 */
        event= eventCal.createEvent( title, start, end, {location: location});
    
        event.setColor("3") // 設定顏色為紫色
        event.setDescription(descripion) // 設定描述
      }
    }
  }
}