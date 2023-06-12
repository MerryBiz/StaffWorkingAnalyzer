var FUNC_NAME = "collectStaffDailyAttendance";

// トリガーに指定。初期処理のためここでプロパティの値を初期化する。
function triggerCollectStaffDailyAttendance() {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty(FUNC_NAME + CNT_SUFFIX, "0");
  collectStaffDailyAttendance();
}

// 6分制限の対策として再実行時はこの関数から実行する。
function collectStaffDailyAttendance() {
  try {
    collect(FUNC_NAME);
  }
  catch (e) {
    sendSlack(e);
    throw e;
  }
}

function collect(funcName) {
  Logger.log("funcName is:"+funcName);
  console.time("CostCollector");
  var start_time = new Date();
  var properties = PropertiesService.getScriptProperties();
  var currentCntProp = properties.getProperty(funcName + CNT_SUFFIX);
  console.log("currentCntProp is :"+currentCntProp);
  var currentCnt = 0;
  if (currentCntProp) {
    currentCnt = parseInt(currentCntProp);
  }
  console.log("start count is :" + currentCnt);
  delete_specific_triggers(funcName);
  var cmSpreadSheet = SpreadsheetApp.openById(CM_SHEET_ID);
  var cmStockSheet = cmSpreadSheet.getSheetByName(STAFF_ATTENDANCE_STOCK_SHEET_NAME);
  if (!cmStockSheet) {
    console.warn("ストックシートが取得できません。実行関数：" + funcName);
    return;
  }
  var prevMonthTitle = getPrevMonthTitle();
  Logger.log(prevMonthTitle);
  if (currentCnt === 0) {
    removeCollectedRows(cmStockSheet, prevMonthTitle);
  }
  var targetFiles = this.getTargetFiles();
  var inputValues = new Array();
  for (var cnt = currentCnt; cnt < targetFiles.length; cnt++) {
    var file = targetFiles[cnt];
    var currentSpreadsheet = SpreadsheetApp.open(file);
    console.log("execute file : "+currentSpreadsheet.getName());
    var currentAttendanceSheet = currentSpreadsheet.getSheetByName(prevMonthTitle);
    if (!currentAttendanceSheet) {
      // notFindSheetCnt++;
      console.log("先月分の勤務シートが見つかりませんでした。処理をスキップします。:" + file.getName());
      continue;
    }
    var extractData = this.extractAttendanceSummary(currentAttendanceSheet);
    for (var i = 0; i < extractData.length; i++) {
      var inputData = extractData[i];
      var customerId = this.extractCustomerId(inputData);
      var fileName = currentSpreadsheet.getName()
      var staffId = extractStaffId(fileName);
      inputData.unshift(customerId);
      inputData.unshift(staffId);
      inputData.unshift(fileName);
      inputData.unshift(prevMonthTitle);
      inputValues.push(inputData);
    }
    if (needRestart(start_time, cnt, funcName)) {
      if (inputValues.length > 0) {
        console.log("append value. size is ["+inputValues.length+"] and sample data is ["+ inputValues[0][0] + " / "+inputValues[0][1] + " / "+inputValues[0][2]+"].");
        cmStockSheet.getRange(cmStockSheet.getLastRow() + 1, 1, inputValues.length, inputValues[0].length).setValues(inputValues);
      }
      console.log("Restart!! CurrentCnt is " + cnt);
      return;
    }
  }
  if (inputValues.length > 0) {
    cmStockSheet.getRange(cmStockSheet.getLastRow() + 1, 1, inputValues.length, inputValues[0].length).setValues(inputValues);
  }
  properties.setProperty(funcName + CNT_SUFFIX, "0");
  console.timeEnd("CostCollector");
}

function extractAttendanceSummary(curentSheet) {
  var dat = curentSheet.getRange(this.getTargetRangePosition(curentSheet)).getValues();
  var extractData = [];
  for (var cnt = 0; cnt < dat.length; cnt++) {
    if (dat[cnt][0]) {
      extractData.push(dat[cnt]);
    }
  }
  return extractData;
}

function getTargetFiles() {
  console.time("sortTime");
  var targetFolder = DriveApp.getFolderById(this.getTargetFolderId());
  Logger.log(targetFolder.getName());
  var files = targetFolder.searchFiles("title contains '勤務実績表'");
  //各スタッフのスプシ毎の処理
  var filesArray = [];
  //検証用のファイル制限
  // var verificationFileNameList = ["S0003_松尾 綾子様_勤務実績表", "S0006_皆見 佳子様_勤務実績表", "S0018_原田 雅美 様_勤務実績表", "S0021_近藤 昌代様_勤務実績表", "S0004_吉益 美江様_勤務実績表"];
  while (files.hasNext()) {
    var file = files.next();
    // for (var k = 0; k < verificationFileNameList.length; k++) {
    //   if (file.getName() === verificationFileNameList[k]) {
        filesArray.push(file);
      //   break;
      // }
    // }
  }
  // 非アクティブスタッフのファイル抽出
  var inactiveFolerId = this.getInactiveTargetFolderId();
  if (inactiveFolerId) {
    console.log("get Inactive file list");
    var inactiveTargetFolder = DriveApp.getFolderById(inactiveFolerId);
    Logger.log(inactiveTargetFolder.getName());
    var files = inactiveTargetFolder.searchFiles("title contains '勤務実績表'");
    while (files.hasNext()) {
      var file = files.next();
      filesArray.push(file);
    }
  }
  filesArray.sort(function (a, b) {
    if (a.getName() > b.getName()) {
      return 1;
    }
    else {
      return -1;
    }
  });
  // TODO Need to remove.
  // for (var i = 0; i < filesArray.length; i++) {
  //   Logger.log(filesArray[i].getName());
  // }
  console.timeEnd("sortTime");
  return filesArray;
}

function getTargetFolderId() {
  return getStaffAttendanceFolderId();
}
function getInactiveTargetFolderId() {
  return getInactiveStaffAttendanceFolderId();
}
function getTargetRangePosition(curentSheet) {
  var rowIndex = getLastRowIndex(curentSheet);
  var targetRange = "K7:T"+rowIndex;
  return targetRange;
}
function getLastRowIndex(curentSheet){
  var kRangeValues = curentSheet.getRange("K:K").getValues();
    for (var i = 0; i < kRangeValues.length; i++) {
        if (kRangeValues[i][0] === "合計金額(税込)") {
            return i;
        }
    }
    return 0;
}
function test(){
  var spredsheet = SpreadsheetApp.openById("14up1mArH4ER_r5tA6w1TwN0VGAoktrBVPZSYfc4WJes");
  var range = getTargetRangePosition(spredsheet.getSheetByName("2022年3月"));

}

function extractCustomerId(currentDataArray) {
  var customerIdColumn = currentDataArray[1];
  var customerIdCandidate = customerIdColumn.split(" ")[0];
  var regex = new RegExp(/^A[0-9]{5}$/);
  if (typeof (customerIdCandidate) == "string" && regex.test(customerIdCandidate)) {
    return customerIdCandidate;
  }
  return "";
}
function extractStaffId(fileName) {
  var staffIdCandidate = fileName.split("_")[0];
  var regex = new RegExp(/^S[0-9]{4}$/);
  if (typeof (staffIdCandidate) == "string" && regex.test(staffIdCandidate)) {
    return staffIdCandidate;
  }
  return "";
}
function getCostManagementStockSheetName() {
  return STAFF_ATTENDANCE_STOCK_SHEET_NAME;
}

