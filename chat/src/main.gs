// The MIT License (MIT)
//
// Copyright (c) 2017 SoftBank Corp.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.


// ----------------------------------------------------------------------------
// グローバル変数
/* globals CHATUTIL_load_creds */
/* globals CHATUTIL_load_config */
/* globals NLCUTIL_exec_check_clfs */
/* globals CHATUTIL_send_message */
/* globals CHATUTIL_store_reply */
/* globals LINE_REPLY_URL */

/**
 * 分類器数
 * @type {Integer}
 */
var NB_CLFS = 3; // eslint-disable-line no-unused-vars

/**
 * 設定シートフィールドインデックス
 * @type {Object}
 * @property {Integer} ws_name      シート名
 * @property {Integer} start_col    定義開始列
 * @property {Integer} start_row    定義開始行
 * @property {Integer} text_col     学習テキスト列
 * @property {Integer} intent1_col  インテント列1
 * @property {Integer} result1_col  分類結果列1
 * @property {Integer} resconf1_col 確信度列1
 * @property {Integer} restime1_col 分類日時列1
 * @property {Integer} intent2_col  インテント列2
 * @property {Integer} result2_col  分類結果列2
 * @property {Integer} resconf2_col 確信度列2
 * @property {Integer} restime2_col 分類日時列2
 * @property {Integer} intent3_col  インテント列3
 * @property {Integer} result3_col  分類結果列3
 * @property {Integer} resconf3_col 確信度列3
 * @property {Integer} restime3_col 分類日時列3
 * @property {Integer} log_ws       ログシート名
 * @property {Integer} conv_ws      応答設定シート名
 * @property {Integer} start_msg    開始メッセージ
 * @property {Integer} other_msg    該当なしメッセージ
 * @property {Integer} error_msg    例外メッセージ
 * @property {Integer} avatar_url   アバターアイコンURL
 * @property {Integer} giveup_msg   テキスト以外のメッセージ
 */
var CONF_INDEX = { // eslint-disable-line no-unused-vars
    ws_name: 0,
    start_col: 1,
    start_row: 2,
    text_col: 3,
    intent1_col: 4,
    result1_col: 5,
    resconf1_col: 6,
    restime1_col: 7,
    intent2_col: 8,
    result2_col: 9,
    resconf2_col: 10,
    restime2_col: 11,
    intent3_col: 12,
    result3_col: 13,
    resconf3_col: 14,
    restime3_col: 15,
    log_ws: 16,
    conv_ws: 17,
    start_msg: 18,
    other_msg: 19,
    error_msg: 20,
    avatar_url: 21,
    giveup_msg: 22,
};

/**
 * バインドされているスプレッドシートオブジェクト
 * @type {Spreadsheet}
 */
var SELF_SS = SpreadsheetApp.getActiveSpreadsheet();

/**
 * バインドされているスプレッドシートのID
 * @type {String}
 */
var SS_ID = SELF_SS.getId();

/**
 * 設定メタデータ
 * @type {ConfigSet}
 * @property {String} ss_id           スプレッドシートID
 * @property {String} ws_name         設定シート名
 * @property {Integer} st_start_row    設定シート定義開始行
 * @property {Integer} st_start_col    設定シート定義開始列
 * @property {Integer} notif_start_row 通知設定定義開始行
 * @property {Integer} notif_start_col 通知設定定義開始列
 * @property {Boolean} result_override 分類結果上書きオプション
 * @property {Integer} clfs_start_col  分類器表示開始列
 * @property {Integer} clfs_start_row  分類器表示開始行
 * @property {Integer} log_start_col   ログ開始列
 * @property {Integer} log_start_row   ログ開始行
 */
var CONFIG_SET = {
    ss_id: SS_ID,
    ws_name: '設定',
    st_start_row: 2,
    st_start_col: 2,
    conv_start_row: 2,
    conv_start_col: 1,
    result_override: false,
    clfs_start_col: 5,
    clfs_start_row: 3,
    log_start_col: 1,
    log_start_row: 2,
};
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * インクルード
 * htmlコンテンツをインクルードする
 * @param  {String} filename ファイル名
 * @return {Object}          HtmlService
 */
function include(filename) { // eslint-disable-line no-unused-vars
    return HtmlService.createHtmlOutputFromFile(filename)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .getContent();
}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * HTMLサービス
 * @return {Object}   HTML
 */
function doGet() { // eslint-disable-line no-unused-vars
    // (1) 画面表示時の処理

    var conf = CHATUTIL_load_config(CONFIG_SET);

    var web = HtmlService.createTemplateFromFile('index');
    web.data = {
        title: "Chat Bot Demo",
        bot_name: "名無しのボット君",
        user_name: "あなた",
        start_msg: conf.sheet_conf.start_msg,
        avatar_url: conf.sheet_conf.avatar_url,
    };

    return web.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * LINE応答処理
 * @param  {Event} e イベントパラメータ
 */
function doPost(e) { // eslint-disable-line no-unused-vars

    var conf = CHATUTIL_load_config(CONFIG_SET);

    var contents = JSON.parse(e.postData.contents);

    var event = contents.events[0];

    var reply_token = event.replyToken;
    if (typeof reply_token === 'undefined') {
        return;
    }

    var res_msg = "";
    if (event.type === "follow") {
        res_msg = conf.sheet_conf.start_msg;
        CHATUTIL_store_reply('#FOLLOW', res_msg);
    } else {
        if (event.message.type === "text") {
            var user_message = event.message.text;
            var res = CHATUTIL_send_message(user_message);
            res_msg = res.response[0];
        } else {
            res_msg = conf.sheet_conf.giveup_msg;
            CHATUTIL_store_reply('#UNDEFINED', res_msg);
        }
    }

    var url = LINE_REPLY_URL;

    var CREDS;
    try {
        CREDS = CHATUTIL_load_creds();
    } catch (err) {
        CHATUTIL_store_reply('#ERROR', err);
        return;
    }

    try {
        UrlFetchApp.fetch(url, {
            headers: {
                "Content-Type": 'application/json; charset=UTF-8',
                Authorization: 'Bearer ' + CREDS.channel_access_token,
            },
            method: 'post',
            payload: JSON.stringify({
                replyToken: reply_token,
                messages: [{
                    type: 'text',
                    text: res_msg,
                }],
            }),
        });
    } catch (err) {
        CHATUTIL_store_reply('#ERROR', err);
    }

}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * オープン時処理
 */
function onOpen() { // eslint-disable-line no-unused-vars

    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Watson')
        .addItem('学習', 'CHATUTIL_train_all')
        .addSubMenu(ui.createMenu('削除')
            .addItem('分類器1', 'NLCUTIL_del_clf1')
            .addItem('分類器2', 'NLCUTIL_del_clf2')
            .addItem('分類器3', 'NLCUTIL_del_clf3'))
        .addToUi();

    NLCUTIL_exec_check_clfs();
}
// ----------------------------------------------------

/**
 * テスト用メッセージ送信
 */
function test_send() { // eslint-disable-line no-unused-vars
    // ※ デバッグ機能のブレークポイントをここに設定して、処理を追います。

    CHATUTIL_send_message("電話番号を教えて");
}
