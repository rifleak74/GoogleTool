/** 取得日曆 */
function getCalender(calenderID) {
  var getCal = CalendarApp.getCalendarById(calenderID);
  /** 若想要使用浮動取得Calender的方式，也可以使用以下方式 */
  // var calendarID = spreadsheet.getRange("在GoogleSheets資料表中的位置").getValue();
  // var eventCal = CalendarApp.getCalendarById(calendarID); 

  return getCal
}

function main() {
  eventCal = getCalender("你的日曆Token")
  /** 取得所有行程 */
  var now = new Date();
  var longTimeAgo = new Date(2020, 7, 16, 0, 0)
  var allEvents = eventCal.getEvents(longTimeAgo, now) // 取得所有行程
  // var allEvents = eventCal.getEvents(longTimeAgo, now, {author: '作者@gmail.com'}) // 取得特定作者行程
  // var allEvents = eventCal.getEvents(longTimeAgo, now, {search: 'Ivan'}) // 取得特定者行程
  // Logger.log(allEvents[0].getTitle())

  /** 存檔 */
  var text = "行程標題,行程ID,創建者,開始時間,結束時間,地址,描述\n";
  for (var i in allEvents) {
    theTitle = allEvents[i].getTitle() // 取得標題
    theID = allEvents[i].getId() // 取得ID
    theCreators = allEvents[i].getCreators() // 取得創建者
    theStartTime = allEvents[i].getStartTime() // 取得開始時間
    theEndTime = allEvents[i].getEndTime() // 取得結束時間
    theLocation = allEvents[i].getLocation() // 取得地址
    theDescription = allEvents[i].getDescription() // 取得描述

    text = text + theTitle + "," + theID + "," + theCreators + "," + theStartTime + "," + theEndTime + "," + theLocation + "," + theDescription + "," + "\n";
    
  }
  var folder = DriveApp.getFoldersByName("備份檔").next(); //選擇儲存的資料夾
  var ff = folder.createFile("T2001D.csv",text,MimeType.PLAIN_TEXT);  

}
