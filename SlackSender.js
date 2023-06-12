//定数
var SLACK_CHANNEL = "#6cp_admin_it_office";
var BOT_NAME = "スタッフ実績分析集計Bot";
// var MENTIONS = ["@hirokazu.nezu"];
/*
* @param <SalesMeetingMemo> salesMeetingMemo
*/
function sendSlack(errorText) {
    send(generateMessage(errorText));
}
/*
* @param {string} mesage
*/
function send(message) {
    var url = PropertiesService.getScriptProperties().getProperty('SLACK_URL');
    var data = { "channel": SLACK_CHANNEL, "username": BOT_NAME, "text": message };
    var payload = JSON.stringify(data);
    var options = {
        "method": "POST",
        "contentType": "application/json",
        "payload": payload
    };
    var response = UrlFetchApp.fetch(url, options);
}
function generateMessage(text) {
    var message = "";
    // for (var i = 0; i < MENTIONS.length; i++) {
    //     message += "<" + MENTIONS[i] + "> ";
    // }
    message += "\n";
    message += "【エラー発生報告】";
    message += "\n";
    message += "\n";
    message += "スタッフの実績分析集計ツールでエラーが発生しました。ログを確認してください。";
    message += "\n";
    message += "---------------------------------------";
    message += "\n";
    message += text;
    message += "\n";
    message += "---------------------------------------";
    return message;
}
