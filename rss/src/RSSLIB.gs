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
/* globals SS_ID */
/* globals SELF_SS */
/* globals NB_CLFS */
/* globals CONFIG_SET */
/* globals CONF_INDEX */
/* globals CLFNAME_PREFIX */
/* globals CLF_SEP */
/* globals NOTIF_OPT */
/* globals NLCUTIL_norm_text */
/* globals NLCUTIL_load_creds */
/* globals NLCUTIL_open_dialog */
/* globals NLCUTIL_log_train */
/* globals NLCUTIL_clf_vers */
/* globals NLCUTIL_exec_check_clfs */
/* globals NLCUTIL_set_trigger */
/* globals NLCUTIL_log_classify */
/* globals NLCUTIL_check_notify */
/* globals NLCAPI_get_classifiers */
/* globals NLCAPI_post_classifiers */
/* globals NLCAPI_delete_classifier */
/* globals NLCAPI_post_classify */
/* globals NLCAPI_delete_classifier */
/* globals NLCAPI_delete_classifier */
/* globals NLCUTIL_select_clf */
/* globals NLCUTIL_load_notif_rules */
/* globals NLCUTIL_escape_formula */

/**
 * RSSデータフィールドインデックス
 * @type {Object}
 */
var RSS_FIELDS = {
    NAME: 0,
    TITLE: 1,
    URL: 2,
    CREATED: 3,
    SUMMARY: 4,
};

/**
 * 取得フィードフィールドインデックス
 * @type {Object}
 */
var FEED_FIELDS = {
    TITLE: 0,
    URL: 1,
    CREATED: 2,
    SUMMARY: 3,
};

/**
 * フィード取得用シート名
 * @type {String}
 */
var FEED_WS_NAME = "RSS_WORK";

/**
 * 取得フィード項目名
 * @type {String[]}
 */
var RSS_NAMES = ["title", "url", "created", "summary"];

/**
 * 学習対象選択オプション
 * @type {Object}
 */
var TRAIN_COLUMN = {
    TITLE: "タイトルのみ",
    SUMMARY: "サマリのみ",
    BOTH: "タイトル・サマリ両方",
};
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 設定情報を取得する
 * @param       {Object} config_set メタデータ
 * @return      {Object} 設定情報
 * @throws      {Error} 設定シートが不明です
 * @throws      {Error} 設定シートに問題があります
 */
function RSSUTIL_load_config(config_set) {

    var sheet = SELF_SS.getSheetByName(config_set.ws_name);
    if (sheet === null) {
        throw new Error('設定シートが不明です');
    }

    var nb_conf = Object.keys(CONF_INDEX)
        .length;

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (lastRow < (config_set.st_start_row + nb_conf) - 1) {
        throw new Error("設定シートに問題があります");
    }

    var conf_list = sheet.getRange(config_set.st_start_row, config_set.st_start_col, nb_conf, 1)
        .getValues();

    var sheet_conf = {};
    sheet_conf["ws_name"] = conf_list[CONF_INDEX.ws_name][0];
    sheet_conf["start_col"] = conf_list[CONF_INDEX.start_col][0];
    sheet_conf["start_row"] = conf_list[CONF_INDEX.start_row][0];
    sheet_conf["train_column"] = conf_list[CONF_INDEX.train_column][0];
    sheet_conf["intent_col"] = [
        conf_list[CONF_INDEX.intent1_col][0],
        conf_list[CONF_INDEX.intent2_col][0],
        conf_list[CONF_INDEX.intent3_col][0],
    ];
    sheet_conf["result_col"] = [
        conf_list[CONF_INDEX.result1_col][0],
        conf_list[CONF_INDEX.result2_col][0],
        conf_list[CONF_INDEX.result3_col][0],
    ];
    sheet_conf["resconf_col"] = [
        conf_list[CONF_INDEX.resconf1_col][0],
        conf_list[CONF_INDEX.resconf2_col][0],
        conf_list[CONF_INDEX.resconf3_col][0],
    ];
    sheet_conf["restime_col"] = [
        conf_list[CONF_INDEX.restime1_col][0],
        conf_list[CONF_INDEX.restime2_col][0],
        conf_list[CONF_INDEX.restime3_col][0],
    ];
    sheet_conf["log_ws"] = conf_list[CONF_INDEX.log_ws][0];

    var notif_conf = {};
    notif_conf["option"] = conf_list[CONF_INDEX.notif_opt][0];
    notif_conf["ws_name"] = conf_list[CONF_INDEX.notif_ws][0];

    if (lastRow < config_set.rss_start_row || lastCol < config_set.rss_start_col) {
        throw new Error("設定シートに問題があります");
    }

    var rss_list = sheet.getRange(config_set.rss_start_row, config_set.rss_start_col, (lastRow - config_set.rss_start_row) + 1, 2)
        .getValues();

    var rss_conf = [];
    for (var i = 0; i < rss_list.length; i += 1) {
        var rss_name = rss_list[i][0];
        var rss_url = rss_list[i][1];
        if (rss_name !== "" && rss_url !== "") {
            rss_conf.push({
                name: rss_name,
                url: rss_url,
            });
        }
    }

    return {
        sheet_conf: sheet_conf,
        rss_conf: rss_conf,
        notif_conf: notif_conf,
    };
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * RSSフィードを取得する
 * @param       {Object} feed_set 設定情報
 * @return      {Object} フィード
 */
function RSSUTIL_get_feeds(feed_set) {

    var sheet = SELF_SS.getSheetByName(feed_set.ws_name);

    var formulastring = "";
    formulastring = '=Importfeed( "' + feed_set.url + '", "items title", FALSE )';
    sheet.getRange("A1")
        .setFormula(formulastring);

    formulastring = '=Importfeed( "' + feed_set.url + '", "items url", FALSE )';
    sheet.getRange("B1")
        .setFormula(formulastring);

    formulastring = '=Importfeed( "' + feed_set.url + '", "items created", FALSE )';
    sheet.getRange("C1")
        .setFormula(formulastring);

    formulastring = '=Importfeed( "' + feed_set.url + '", "items summary", FALSE )';
    sheet.getRange("D1")
        .setFormula(formulastring);
    if (sheet.getRange("D1")
        .getValue() === "#N/A") {
        sheet.getRange("D1")
            .setFormula("");
    }

    var lastRow = sheet.getLastRow();

    var rss_data = sheet.getRange(1, 1, lastRow, 4)
        .getValues();

    var result = [];
    for (var cnt = 0; cnt < rss_data.length; cnt += 1) {

        var dateObj = new Date(rss_data[cnt][FEED_FIELDS.CREATED]);
        var localDate = Utilities.formatDate(dateObj, "JST", "yyyy_MMdd_HHmmss");
        result.push({
            name: feed_set.name,
            title: NLCUTIL_escape_formula(rss_data[cnt][FEED_FIELDS.TITLE]),
            url: rss_data[cnt][FEED_FIELDS.URL],
            created: localDate,
            summary: NLCUTIL_escape_formula(rss_data[cnt][FEED_FIELDS.SUMMARY].trim()),
        });
    }

    return result;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * ワークエリアを消去する
 * @param       {Object} feed_set 取得設定
 */
function RSSUTIL_clear_work(feed_set) {

    var sheet = SELF_SS.getSheetByName(feed_set.ws_name);
    if (sheet === null) {
        sheet = SELF_SS.insertSheet(feed_set.ws_name);
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < 1 && lastCol < 1) {
        return;
    }

    sheet.getRange(1, 1, lastRow, lastCol)
        .clear();
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * RSSフィードを追記する
 * @param       {Object} data_set 配置情報
 */
function RSSUTIL_update_data(data_set) {

    var sheet = SELF_SS.getSheetByName(data_set.ws_name);
    if (sheet === null) {
        sheet = SELF_SS.insertSheet(data_set.ws_name);
    }

    var lastRow = sheet.getLastRow();

    var nb_fields = Object.keys(RSS_FIELDS)
        .length;
    var entries = [];
    if (lastRow >= data_set.start_row) {
        entries = sheet.getRange(data_set.start_row, data_set.start_col, (lastRow - data_set.start_row) + 1, nb_fields)
            .getValues();
    } else {
        lastRow = data_set.start_row - 1;
    }

    var row_cnt = 0;
    for (var i = 0; i < data_set.feeds.length; i += 1) {

        var isMatch = 0;

        for (var j = 0; j < entries.length; j += 1) {

            if (entries[j][RSS_FIELDS.NAME] === data_set.feeds[i].name &&
                entries[j][RSS_FIELDS.CREATED] === data_set.feeds[i].created) {
                isMatch = 1;
                break;
            }
        }

        if (isMatch === 0) {
            var record = [
                data_set.feeds[i].name,
                data_set.feeds[i].title,
                data_set.feeds[i].url,
                data_set.feeds[i].created,
                data_set.feeds[i].summary,
            ];

            sheet.getRange(lastRow + row_cnt + 1, data_set.start_col, 1, nb_fields)
                .setValues([record]);
            row_cnt += 1;
        }
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * RSSフォードを取得してシートに追記する
 * @param       {Object} feed_set 取得設定
 * @param       {Object} rss_set  配置設定
 */
function RSSUTIL_load_rss(feed_set, rss_set) {

    RSSUTIL_clear_work(feed_set);

    var rss_feeds = RSSUTIL_get_feeds(feed_set);

    var data_set = {
        ss_id: SS_ID,
        ws_name: rss_set.ws_name,
        start_col: rss_set.start_col,
        start_row: rss_set.start_row,
        rss_name: feed_set.name,
        feeds: rss_feeds,
    };

    RSSUTIL_update_data(data_set);
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 全てのRSSフィードを取得してシートに追記する
 */
function RSSUTIL_crawl() { // eslint-disable-line no-unused-vars

    Logger.log("### RSSUTIL_crawl");

    var conf = RSSUTIL_load_config(CONFIG_SET);

    var SS_UI;
    try {
        SS_UI = SpreadsheetApp.getUi();
    } catch (e) {
        SS_UI = null;
    }

    if (SS_UI !== null) {
        var res = NLCUTIL_open_dialog("データ取得", "データ取得を開始します。よろしいですか？", SS_UI.ButtonSet.OK_CANCEL);
        if (res === SS_UI.Button.CANCEL) {
            NLCUTIL_open_dialog("データ取得", "データ取得を中止しました。", SS_UI.ButtonSet.OK);
            return;
        }
        var msg = "データ取得分類を開始しました。";
        NLCUTIL_open_dialog("データ取得", msg, SS_UI.ButtonSet.OK);
    }

    var rss_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_col: conf.sheet_conf.start_col,
        start_row: conf.sheet_conf.start_row,
    };

    for (var cnt = 0; cnt < conf.rss_conf.length; cnt += 1) {

        var feed_set = {
            ss_id: SS_ID,
            ws_name: FEED_WS_NAME,
            name: conf.rss_conf[cnt].name,
            url: conf.rss_conf[cnt].url,
            columns: RSS_NAMES,
        };

        RSSUTIL_load_rss(feed_set, rss_set);
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * RSSフィードのテキストを学習する
 * @param       {Object} train_set      学習設定
 * @param       {String} creds_username クレデンシャル
 * @param       {String} creds_password クレデンシャル
 * @return      {Object} 学習結果
 * @throws      {Error} データシートが不明です
 * @throws      {Error} 学習・分類対象が不正です
 */
function RSSUTIL_train(train_set, creds_username, creds_password) {

    Logger.log("### RSSUTIL_train");

    var clfs = NLCAPI_get_classifiers(creds_username, creds_password);
    if (clfs.status !== 200) {
        var result = {
            status: clfs.status,
            description: clfs.error,
        };
        return result;
    }

    var sheet = SELF_SS.getSheetByName(train_set.ws_name);
    if (sheet === null) {
        throw new Error("データシートが不明です");
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries;
    if (lastRow < train_set.start_row || lastCol < train_set.start_col) {
        entries = [];
    } else {
        entries = sheet.getRange(train_set.start_row, train_set.start_col, (lastRow - train_set.start_row) + 1, lastCol - train_set.start_col)
            .getValues();
    }

    var row_cnt = 0;
    var csvString = "";

    for (var cnt = 0; cnt < entries.length; cnt += 1) {

        var class_name = NLCUTIL_norm_text(String(entries[cnt][train_set.class_col - train_set.start_col]))
            .trim();
        if (class_name === "") continue;

        if (entries[cnt][RSS_FIELDS.NAME].length === 0) continue;
        if (entries[cnt][RSS_FIELDS.NAME] === 0) continue;

        var train_text = "";
        var title_text = NLCUTIL_norm_text(String(entries[cnt][RSS_FIELDS.TITLE]))
            .trim();
        var summary_text = NLCUTIL_norm_text(String(entries[cnt][RSS_FIELDS.SUMMARY]))
            .trim();

        if (train_set.train_column === TRAIN_COLUMN.TITLE) {
            train_text = title_text;
        } else if (train_set.train_column === TRAIN_COLUMN.SUMMARY) {
            train_text = summary_text;
        } else if (train_set.train_column === TRAIN_COLUMN.BOTH) {
            if (title_text === "") {
                train_text = summary_text;
            } else if (summary_text === "") {
                train_text = title_text;
            } else {
                train_text = title_text + " " + summary_text;
            }
        } else {
            throw new Error("学習・分類対象が不正です");
        }

        if (train_text.length === 0) continue;

        if (train_text.length > 1024) {
            train_text = train_text.substring(0, 1024);
        }

        csvString = csvString + '"' + train_text + '","' + class_name + '"' +
            "\r\n";
        row_cnt += 1;
    }

    if (row_cnt === 0) {
        return {
            status: 0,
            description: "学習データなし",
        };
    }

    var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, train_set.clf_name);

    var new_version = (clf_info.max_ver + 1);
    var clf_name = train_set.clf_name + CLF_SEP + new_version;

    var res = NLCAPI_post_classifiers(creds_username, creds_password, csvString, clf_name, 'ja');
    if (clf_info.count >= 2 && res.status === 200) {

        NLCAPI_delete_classifier(creds_username, creds_password, clf_info.clfs[clf_info.min_ver].classifier_id);

    }

    NLCUTIL_exec_check_clfs();
    NLCUTIL_set_trigger("NLCUTIL_exec_check_clfs", 1);

    var result2 = {
        status: res.status,
        nlc: res,
        rows: row_cnt,
        version: new_version,
    };
    return result2;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 特定分類器で学習して結果をログに出力する
 * @param       {Integer} clf_no 分類器番号
 */
function RSSUTIL_train_set(clf_no) {

    Logger.log("### RSSUTIL_train_set", clf_no);

    var CREDS = NLCUTIL_load_creds();

    var conf = RSSUTIL_load_config(CONFIG_SET);

    var train_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_row: conf.sheet_conf.start_row,
        start_col: conf.sheet_conf.start_col,
        end_row: -1,
        train_column: conf.sheet_conf.train_column,
        text_col: conf.sheet_conf.train_column,
        class_col: conf.sheet_conf.intent_col[clf_no - 1],
        clf_no: clf_no,
        clf_name: CLFNAME_PREFIX + clf_no,
    };

    var train_result = RSSUTIL_train(train_set, CREDS.username, CREDS.password);

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row,
    };

    NLCUTIL_log_train(log_set, train_set, train_result);
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * イベント実行用学習処理
 */
function RSSUTIL_train_all() { // eslint-disable-line no-unused-vars

    var CREDS = NLCUTIL_load_creds();

    var conf = RSSUTIL_load_config(CONFIG_SET);

    var SS_UI;
    try {
        SS_UI = SpreadsheetApp.getUi();
    } catch (e) {
        SS_UI = null;
    }

    if (SS_UI !== null) {
        var res = NLCUTIL_open_dialog("学習", "学習を開始します。よろしいですか？", SS_UI.ButtonSet.OK_CANCEL);
        if (res === SS_UI.Button.CANCEL) {
            NLCUTIL_open_dialog("学習", "学習を中止しました。", SS_UI.ButtonSet.OK);
            return;
        }

        var msg = "学習を開始しました。ログは「" + conf.sheet_conf.log_ws + "」シートをご参照ください。";
        msg += "\nステータスは「" + CONFIG_SET.ws_name + "」シートをご参照ください。";
        NLCUTIL_open_dialog("学習", msg, SS_UI.ButtonSet.OK);
    }

    for (var i = 1; i <= NB_CLFS; i += 1) {
        RSSUTIL_train_set(i);
    }

}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * RSSフィードのテキストを分類する
 * @param       {Object} test_set       分類設定
 * @param       {String} creds_username クレデンシャル
 * @param       {String} creds_password クレデンシャル
 * @param       {Boolean} override       上書きオプション
 * @return      {Object} 分類結果
 * @throws      {Error}  データシートが不明です
 * @throws      {Error} 学習・分類対象が不正です
 */
function RSSUTIL_classify(test_set, creds_username, creds_password, override) {

    Logger.log("### RSSUTIL_classify");

    var clf_id = NLCUTIL_select_clf(test_set.clf_name, creds_username, creds_password);
    if (clf_id === "") {
        return {
            status: 900,
            description: "classifier does not found",
        };
    }

    var sheet = SELF_SS.getSheetByName(test_set.ws_name);
    if (sheet === null) {
        throw new Error("データシートが不明です");
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries;
    if (lastRow < test_set.start_row || lastCol < test_set.start_col) {
        entries = [];
    } else {
        entries = sheet.getRange(test_set.start_row, test_set.start_col, (lastRow - test_set.start_row) + 1, (lastCol - test_set.start_col) + 1)
            .getValues();
    }

    var hasError = 0;
    var row_cnt = 0;
    var nlc_res;
    var err_res;
    for (var cnt = 0; cnt < entries.length; cnt += 1) {

        var rss_name = entries[cnt][RSS_FIELDS.NAME];
        if (rss_name.length === 0) continue;
        if (rss_name === 0) continue;

        var result_text;
        if (lastCol < test_set.result_col) {
            result_text = "";
        } else {
            result_text = entries[cnt][test_set.result_col - test_set.start_col];
        }
        if (result_text !== "" && override !== true) continue;

        var title_text = NLCUTIL_norm_text(String(entries[cnt][RSS_FIELDS.TITLE]))
            .trim();
        var summary_text = NLCUTIL_norm_text(String(entries[cnt][RSS_FIELDS.SUMMARY]))
            .trim();

        var test_text;
        if (test_set.train_column === TRAIN_COLUMN.TITLE) {
            test_text = title_text;
        } else if (test_set.train_column === TRAIN_COLUMN.SUMMARY) {
            test_text = summary_text;
        } else if (test_set.train_column === TRAIN_COLUMN.BOTH) {
            if (title_text === "") {
                test_text = summary_text;
            } else if (summary_text === "") {
                test_text = title_text;
            } else {
                test_text = title_text + " " + summary_text;
            }
        } else {
            throw new Error("学習・分類対象が不正です");
        }

        if (test_text.length === 0) continue;

        if (test_text.length > 1024) {
            test_text = test_text.substring(0, 1024);
        }

        nlc_res = NLCAPI_post_classify(creds_username, creds_password, clf_id, test_text);
        if (nlc_res.status !== 200) {
            hasError = 1;
            err_res = nlc_res;
        }

        var r = nlc_res.body.top_class;
        sheet.getRange(test_set.start_row + cnt, test_set.result_col, 1, 1)
            .setValue(r);

        var t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
        sheet.getRange(test_set.start_row + cnt, test_set.restime_col, 1, 1)
            .setValue(t);

        // TODO: confidence

        row_cnt += 1;

        if (test_set.notif_opt === NOTIF_OPT.ON) {

            var record = sheet.getRange(test_set.start_row + cnt, 1, 1, lastCol)
                .getValues();

            NLCUTIL_check_notify(test_set.notif_set, record[0]);
        }
    }

    if (row_cnt === 0) {
        return {
            status: 0,
            rows: 0,
        };
    }
    if (hasError === 1) {
        return {
            status: err_res.status,
            rows: row_cnt,
            nlc: err_res,
        };
    }
    return {
        status: nlc_res.status,
        rows: row_cnt,
        nlc: nlc_res,
    };
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 特定分類器で分類して結果をログに出力する
 * @param       {Integer} clf_no 分類器番号
 */
function RSSUTIL_classify_set(clf_no) {

    Logger.log("### RSSUTIL_classify_set " + clf_no);

    var CREDS = NLCUTIL_load_creds();

    var conf = RSSUTIL_load_config(CONFIG_SET);

    var notif_conf = {
        ss_id: SS_ID,
        ws_name: conf.notif_conf.ws_name,
        start_col: CONFIG_SET.notif_start_col,
        start_row: CONFIG_SET.notif_start_row,
    };

    var rules = NLCUTIL_load_notif_rules(notif_conf);

    var notif_set = {
        rules: rules,
        from: CONFIG_SET.notif_from,
        sender: CONFIG_SET.notif_sender,
        result_cols: conf.sheet_conf.result_col,
    };

    var test_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_col: conf.sheet_conf.start_col,
        start_row: conf.sheet_conf.start_row,
        end_row: -1,
        text_col: conf.sheet_conf.train_column,
        result_col: conf.sheet_conf.result_col[clf_no - 1],
        restime_col: conf.sheet_conf.restime_col[clf_no - 1],
        clf_name: CLFNAME_PREFIX + String(clf_no),
        notif_set: notif_set,
        notif_opt: conf.notif_conf.option,
    };

    var test_result = RSSUTIL_classify(test_set, CREDS.username, CREDS.password);

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row,
    };

    NLCUTIL_log_classify(log_set, test_set, test_result);
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * イベント実行用分類処理
 * @throws      {Error}  データシートが不明です
 * @throws      {Error} 学習・分類対象が不正です
 */
function RSSUTIL_classify_all() { // eslint-disable-line no-unused-vars

    Logger.log("### RSSUTIL_classify_all");

    var CREDS = NLCUTIL_load_creds();

    var conf = RSSUTIL_load_config(CONFIG_SET);

    var SS_UI;
    try {
        SS_UI = SpreadsheetApp.getUi();
    } catch (e) {
        SS_UI = null;
    }

    if (SS_UI !== null) {
        var res = NLCUTIL_open_dialog("分類", "分類を開始します。よろしいですか？", SS_UI.ButtonSet.OK_CANCEL);
        if (res === SS_UI.Button.CANCEL) {
            NLCUTIL_open_dialog("分類", "分類を中止しました。", SS_UI.ButtonSet.OK);
            return;
        }
        var msg = "分類を開始しました。ログは「" + conf.sheet_conf.log_ws + "」シートをご参照ください。";
        NLCUTIL_open_dialog("分類", msg, SS_UI.ButtonSet.OK);
    }

    var notif_conf = {
        ss_id: SS_ID,
        ws_name: conf.notif_conf.ws_name,
        start_col: CONFIG_SET.notif_start_col,
        start_row: CONFIG_SET.notif_start_row,
    };

    var rules = NLCUTIL_load_notif_rules(notif_conf);

    var notif_set = {
        rules: rules,
        from: CONFIG_SET.notif_from,
        sender: CONFIG_SET.notif_sender,
        result_cols: conf.sheet_conf.result_col,
    };

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row,
    };

    var test_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_col: conf.sheet_conf.start_col,
        start_row: conf.sheet_conf.start_row,
        end_row: -1,
        text_col: conf.sheet_conf.train_column,
        notif_set: notif_set,
        notif_opt: conf.notif_conf.option,
    };

    var clf_ids = [];
    for (var i = 0; i < NB_CLFS; i += 1) {

        var clf_name = CLFNAME_PREFIX + String(i + 1);

        test_set.clf_no = i + 1;
        test_set.clf_name = clf_name;
        test_set.result_col = conf.sheet_conf.result_col[i];
        test_set.restime_col = conf.sheet_conf.restime_col[i];

        var clf = NLCUTIL_select_clf(clf_name, CREDS.username, CREDS.password);
        if (clf.status === "Training") {
            NLCUTIL_log_classify(log_set, test_set, {
                status: 900,
                description: "トレーニング中",
                clf_id: clf.clf_id,
            });
        } else if (clf.status === "Nothing") {
            NLCUTIL_log_classify(log_set, test_set, {
                status: 900,
                description: "分類器なし",
                clf_id: "",
            });
        } else if (clf.status !== "Available") {
            NLCUTIL_log_classify(log_set, test_set, {
                status: 800,
                description: clf.status,
                clf_id: "",
            });
        }

        clf_ids.push({
            id: clf.clf_id,
            status: clf.status,
        });
    }

    var sheet = SELF_SS.getSheetByName(conf.sheet_conf.ws_name);
    if (sheet === null) {
        throw new Error("データシートが不明です");
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries;
    if (lastRow < conf.sheet_conf.start_row || lastCol < conf.sheet_conf.start_col) {
        entries = [];
    } else {
        entries = sheet.getRange(conf.sheet_conf.start_row, conf.sheet_conf.start_col, (lastRow - conf.sheet_conf.start_row) + 1, (lastCol - conf.sheet_conf.start_col) + 1)
            .getValues();
    }

    var hasError = 0;
    var nlc_res;
    var err_res;
    //var row_cnt = 0;
    var res_rows = [0, 0, 0];
    for (var cnt = 0; cnt < entries.length; cnt += 1) {

        var rss_name = entries[cnt][RSS_FIELDS.NAME];
        if (rss_name.length === 0) continue;
        if (rss_name === 0) continue;

        var title_text = NLCUTIL_norm_text(String(entries[cnt][RSS_FIELDS.TITLE]))
            .trim();
        var summary_text = NLCUTIL_norm_text(String(entries[cnt][RSS_FIELDS.SUMMARY]))
            .trim();

        var test_text;
        if (conf.sheet_conf.train_column === TRAIN_COLUMN.TITLE) {
            test_text = title_text;
        } else if (conf.sheet_conf.train_column === TRAIN_COLUMN.SUMMARY) {
            test_text = summary_text;
        } else if (conf.sheet_conf.train_column === TRAIN_COLUMN.BOTH) {
            if (title_text === "") {
                test_text = summary_text;
            } else if (summary_text === "") {
                test_text = title_text;
            } else {
                test_text = title_text + " " + summary_text;
            }
        } else {
            throw new Error("学習・分類対象が不正です");
        }

        if (test_text.length === 0) continue;

        if (test_text.length > 1024) {
            test_text = test_text.substring(0, 1024);
        }

        var updates = 0;
        var upd_flg = [0, 0, 0];
        for (var j = 0; j < NB_CLFS; j += 1) {

            if (clf_ids[j].status !== "Available") continue;

            var result_text;
            if (lastCol < conf.sheet_conf.result_col[j]) {
                result_text = "";
            } else {
                result_text = entries[cnt][conf.sheet_conf.result_col[j] - conf.sheet_conf.start_col];
            }

            if (result_text !== "" && CONFIG_SET.result_override !== true) continue;

            nlc_res = NLCAPI_post_classify(CREDS.username, CREDS.password, clf_ids[j].id, test_text);
            if (nlc_res.status !== 200) {
                hasError = 1;
                err_res = nlc_res;
            } else {
                var r = nlc_res.body.top_class;
                sheet.getRange(conf.sheet_conf.start_row + cnt, conf.sheet_conf.result_col[j], 1, 1)
                    .setValue(r);
                var c = nlc_res.body.classes[0].confidence;
                sheet.getRange(conf.sheet_conf.start_row + cnt, conf.sheet_conf.resconf_col[j], 1, 1)
                    .setValue(c);
                var t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
                sheet.getRange(conf.sheet_conf.start_row + cnt, conf.sheet_conf.restime_col[j], 1, 1)
                    .setValue(t);
                updates += 1;
                res_rows[j] += 1;
                upd_flg[j] = 1;
            }
        }

        if (test_set.notif_opt === NOTIF_OPT.ON) {
            if (updates > 0) {
                //row_cnt += 1;
                var record = sheet.getRange(conf.sheet_conf.start_row + cnt, 1, 1, lastCol)
                    .getValues();
                NLCUTIL_check_notify(notif_set, record[0], upd_flg);
            }
        }
    }

    for (var k = 0; k < NB_CLFS; k += 1) {

        if (clf_ids[k].status !== "Available") continue;

        var clfname = CLFNAME_PREFIX + String(k + 1);
        test_set.clf_no = k + 1;
        test_set.clf_name = clfname;
        test_set.result_col = conf.sheet_conf.result_col[k];
        test_set.restime_col = conf.sheet_conf.restime_col[k];

        var test_result;
        if (res_rows[k] === 0) {
            test_result = {
                status: 0,
                rows: 0,
            };
            NLCUTIL_log_classify(log_set, test_set, test_result);

        } else {

            if (hasError === 1) {
                test_result = {
                    status: err_res.status,
                    rows: res_rows[k],
                    nlc: err_res,
                };
                NLCUTIL_log_classify(log_set, test_set, test_result);

            } else {
                test_result = {
                    status: nlc_res.status,
                    rows: res_rows[k],
                    nlc: nlc_res,
                };
                NLCUTIL_log_classify(log_set, test_set, test_result);
            }
        }
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
// イベント用ラッパー
/**
 * 分類器1の学習
 */
function RSSUTIL_train_no1() { // eslint-disable-line no-unused-vars
    RSSUTIL_train_set(1);
}

/**
 * 分類器2の学習
 */
function RSSUTIL_train_no2() { // eslint-disable-line no-unused-vars
    RSSUTIL_train_set(2);
}

/**
 * 分類器3の学習
 */
function RSSUTIL_train_no3() { // eslint-disable-line no-unused-vars
    RSSUTIL_train_set(3);
}

/**
 * 分類器1の分類
 */
function RSSUTIL_classify_no1() { // eslint-disable-line no-unused-vars
    RSSUTIL_classify_set(1);
}

/**
 * 分類器2の分類
 */
function RSSUTIL_classify_no2() { // eslint-disable-line no-unused-vars
    RSSUTIL_classify_set(2);
}

/**
 * 分類器3の分類
 */
function RSSUTIL_classify_no3() { // eslint-disable-line no-unused-vars
    RSSUTIL_classify_set(3);
}
// ----------------------------------------------------------------------------
