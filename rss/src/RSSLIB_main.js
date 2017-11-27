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
// グローバル
/* globals NLCUTIL_exec_check_clfs */
/* globals NLCUTIL_load_creds */

/**
 * 分類器数
 * @type {Integer}
 */
var NB_CLFS = 3;

/**
     * 設定シートフィールドインデックス
 * @type {Object}
 * @property {Integer} ws_name      シート名
 * @property {Integer} start_col    定義開始列
 * @property {Integer} start_row    定義開始行
 * @property {Integer} train_column 学習テキスト選択列
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
 * @property {Integer} notif_opt    通知オプション
 * @property {Integer} notif_ws     通知設定シート名
 */
var CONF_INDEX = {
    ws_name: 0,
    start_col: 1,
    start_row: 2,
    train_column: 3,
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
    notif_opt: 17,
    notif_ws: 18,
};

/**
 * バインドされているスプレッドシートオブジェクト
 * @type {Spreadsheet}}
 */
var SELF_SS = SpreadsheetApp.getActiveSpreadsheet();

/**
 * バインドされているスプレッドシートのID
 * @type {String}
 */
var SS_ID = SELF_SS.getId();

/**
 * クレデンシャル情報
 * @type {Creds}
 */
var CREDS = NLCUTIL_load_creds();

/**
 * 設定メタデータ
 * @type {ConfigSet} CONFIG_SET
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
    rss_start_row: 3,
    rss_start_col: 5,
    notif_start_row: 2,
    notif_start_col: 1,
    result_override: false,
    clfs_start_col: 9,
    clfs_start_row: 3,
    log_start_col: 1,
    log_start_row: 2,
};
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * オープン時処理
 */
function onOpen() {

    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Watson')
        .addItem('データ取得', 'RSSUTIL_crawl')
        .addItem('分類', 'RSSUTIL_classify_all')
        .addItem('学習', 'RSSUTIL_train_all')
        .addSubMenu(ui.createMenu('削除')
            .addItem('分類器1', 'NLCUTIL_del_clf1')
            .addItem('分類器2', 'NLCUTIL_del_clf2')
            .addItem('分類器3', 'NLCUTIL_del_clf3'))
        .addToUi();

    NLCUTIL_exec_check_clfs();
}
// ----------------------------------------------------------------------------
