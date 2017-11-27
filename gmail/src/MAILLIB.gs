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
/* globals CREDS */
/* globals SELF_SS */
/* globals SS_ID */
/* globals CONFIG_SET */
/* globals CONF_INDEX */
/* globals NB_CLFS */
/* globals CONF_INDEX */
/* globals CLFNAME_PREFIX */
/* globals CLF_SEP */
/* globals NLCUTIL_escape_formula */
/* globals NLCUTIL_open_dialog */
/* globals NLCAPI_get_classifiers */
/* globals NLCUTIL_norm_text */
/* globals NLCUTIL_clf_vers */
/* globals NLCUTIL_exec_check_clfs */
/* globals NLCUTIL_set_trigger */
/* globals NLCUTIL_log_train */
/* globals NLCUTIL_load_notif_rules */
/* globals NLCUTIL_select_clf */
/* globals NLCUTIL_log_classify */
/* globals NLCUTIL_check_notify */
/* globals NLCAPI_post_classifiers */
/* globals NLCAPI_delete_classifier */
/* globals NLCAPI_post_classify */

/**
 * メールデータフィールドインデックス
 * @type {Object}
 * @property {Integer} ID メールID
 * @property {Integer} DATE 受信日時
 * @property {Integer} SUBJECT 件名
 * @property {Integer} FROM 送信元メールアドレス
 * @property {Integer} TO 送信先メールアドレス
 * @property {Integer} CC CCメールアドレス
 * @property {Integer} BODY 本文
 * @property {Integer} NORM_BODY 本文(処理対象)
 */
var MAIL_FIELDS = {
    ID: 0,
    DATE: 1,
    SUBJECT: 2,
    FROM: 3,
    TO: 4,
    CC: 5,
    BODY: 6,
    NORM_BODY: 7,
};

/**
 * 学習対象選択オプション
 * @type {Object}
 * @property {String} SUBJECT 件名のみ
 * @property {String} BODY 本文
 * @property {String} BOTH 両方
 */
var TRAIN_COLUMN = {
    SUBJECT: "件名のみ",
    BODY: "本文のみ",
    BOTH: "件名・本文両方",
};

/**
 * GmailApp search 最大スレッド数
 * @type {Integer}
 */
var MAX_THREADS = 500;
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ConfigMeta 設定メタデータ
 * @property {String} ss_id スプレッドシートID
 * @property {String} ws_name 設定シート名
 * @property {Integer} st_start_row 定義開始行
 * @property {Integer} st_start_col 定義開始列
 */
/**
 * @typedef {Object} SheetConf データシート設定
 * @property {String}    ws_name      データシート名
 * @property {Integer}   start_row    定義開始行
 * @property {Integer}   start_col    定義開始列
 * @property {Integer[]} intent_col   インテント列1to3
 * @property {Integer[]} result_col   分類結果列1to3
 * @property {Integer[]} resconf_col  確信度列1to3
 * @property {Integer[]} restime_col  分類日時列1to3
 * @property {String}    log_ws       ログシート名
 * @property {String}    query        フィルタクエリ
 * @property {Integer}   search_limit 最大取得スレッド数
 * @property {Integer}   ago_days     過去分取得日数
 */
/**
 * @typedef {Object} NotifConf 通知設定
 * @property {String} notif_opt 通知オプション{On,Off}
 * @property {String} notif_ws 設定シート名
 */
/**
 * @typedef {Object} ExcConf 本文除外設定
 * @property {String[]} re_list 正規表現リスト
 */
/**
 * @typedef {Object} Config 設定情報
 * @property {SheetConf} sheet_conf データシート設定
 * @property {NotifConf} notif_conf 通知設定
 * @property {ExcConf}   exc_conf   本文除外設定
 */
/**
 * 設定情報の取得
 * @param       {ConfigMeta} config_set メタデータ
 * @return      {Config} 設定情報
 * @throws      {Error}  設定シートが不明です
 * @throws      {Error}  設定シートに問題があります
 */
function MAILUTIL_load_config(config_set) {

    var sheet = SELF_SS.getSheetByName(config_set.ws_name);
    if (sheet === null) {
        throw new Error('設定シートが不明です');
    }

    var nb_conf = Object.keys(CONF_INDEX)
        .length;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < (config_set.st_start_row + nb_conf) - 1 || lastCol < config_set.st_start_col) {
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
    sheet_conf["query"] = conf_list[CONF_INDEX.query][0];
    sheet_conf["search_limit"] = parseInt(conf_list[CONF_INDEX.search_limit][0], 10);
    if (sheet_conf.search_limit < 1) {
        sheet_conf.search_limit = 1;
    }
    if (sheet_conf.search_limit > MAX_THREADS) {
        sheet_conf.search_limit = MAX_THREADS;
    }

    sheet_conf["ago_days"] = conf_list[CONF_INDEX.ago_days][0];

    var notif_conf = {};
    notif_conf["option"] = conf_list[CONF_INDEX.notif_opt][0];
    notif_conf["ws_name"] = conf_list[CONF_INDEX.notif_ws][0];

    if (lastRow < config_set.exc_start_row || lastCol < config_set.exc_start_col) {
        throw new Error("設定シートに問題があります");
    }
    var exc_list = sheet.getRange(config_set.exc_start_row, config_set.exc_start_col, (lastRow - config_set.exc_start_row) + 1, 1)
        .getValues();

    var exc_conf = {};
    var re_list = [];
    for (var i = 0; i < exc_list.length; i += 1) {
        var exc_re = exc_list[i][0];
        if (exc_re !== "") {
            re_list.push(exc_re);
        }
    }
    exc_conf["re_list"] = re_list;

    return {
        sheet_conf: sheet_conf,
        notif_conf: notif_conf,
        exc_conf: exc_conf,
    };
}
// ----------------------------------------------------------------------------
/**
 * @typedef {Object} MailSet
 * @property {String} ss_id SS_ID
 * @property {String} ws_name シート名
 * @property {String} start_col 開始列
 * @property {String} start_row 開始行
 * @property {String} query クエリ
 * @property {String} search_limit 最大スレッド数
 * @property {String} exc_res 除外設定
 * @property {String} msgs メッセージ
 */
/**
 * 最新
 * @param       {MailSet} mail_set メール設定
 * @return      {Date} 最新日付
 * @throws      {Error} データシートが不明です
 */
function MAILUTIL_get_newest(mail_set) {

    var sheet = SELF_SS.getSheetByName(mail_set.ws_name);

    if (sheet === null) {
        throw new Error("データシートが不明です");
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < mail_set.start_row || lastCol < mail_set.start_col) {
        return new Date();
    }

    var entries = [];
    if (lastRow >= mail_set.start_row) {
        entries = sheet.getRange(mail_set.start_row, mail_set.start_col + MAIL_FIELDS.DATE, (lastRow - mail_set.start_row) + 1, 1)
            .getValues();

        var from_max = entries.map(function (flds) {
            if (flds[0] === '') {
                return 0;
            }
            return flds[0];
        });
        // 該当なしの場合は当日のみ
        var max_date = Math.max.apply(null, from_max);
        if (max_date === 0) {
            return new Date();
        }
        return new Date(max_date);
    }
    return new Date();
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @type {String[]} MailData
 * ID, date, subject, from, to, cc, body
 */
/**
 * メールを取得する
 * @param       {MailSet} mail_set 設定情報
 * @return      {MailData} メールデータ
 */
function MAILUTIL_get_messages(mail_set) {

    Logger.log("### MAILUTIL_get_messages");

    var result = [];
    var threads = GmailApp.search(mail_set.query, 0, mail_set.search_limit);

    for (var i = 0; i < threads.length; i += 1) {

        var thread = threads[i];
        var msgs = thread.getMessages();

        for (var j = 0; j < msgs.length; j += 1) {
            var msg = msgs[j];

            var res_msg = [
                msg.getId(), msg.getDate(), NLCUTIL_escape_formula(msg.getSubject()), msg.getFrom(),
                msg.getTo(), msg.getCc(), NLCUTIL_escape_formula(msg.getPlainBody()),
            ];

            if (res_msg[MAIL_FIELDS.DATE] >= mail_set.from_date) {
                result.push(res_msg);
            }
        }
    }

    return result;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 本文を正規表現で除外する
 * @param       {MailSet} mail_set 設定情報
 * @return      {MailSet} 編集結果
 */
function MAILUTIL_trim_exc(mail_set) {

    for (var i = 0; i < mail_set.msgs.length; i += 1) {

        var msg = mail_set.msgs[i];
        var buf = msg[MAIL_FIELDS.BODY];

        for (var j = 0; j < mail_set.exc_res.length; j += 1) {

            var regexp = new RegExp(mail_set.exc_res[j], 'gm');

            buf = buf.replace(regexp, "");

        }

        mail_set.msgs[i][MAIL_FIELDS.NORM_BODY] = buf.trim();

    }

    return mail_set;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * シートのメールデータを更新する
 * @param       {MailSet} mail_set メール情報
 */
function MAILUTIL_update_data(mail_set) {

    var sheet = SELF_SS.getSheetByName(mail_set.ws_name);

    if (sheet === null) {
        sheet = ss.insertSheet(mail_set.ws_name);
    }

    var lastRow = sheet.getLastRow();

    var nb_fields = Object.keys(MAIL_FIELDS)
        .length;

    var entries = [];
    if (lastRow >= mail_set.start_row) {

        entries = sheet.getRange(mail_set.start_row, mail_set.start_col, (lastRow - mail_set.start_row) + 1, nb_fields)
            .getValues();

    } else {
        lastRow = mail_set.start_row - 1;
    }

    var row_cnt = 0;
    for (var i = 0; i < mail_set.msgs.length; i += 1) {

        var isMatch = 0;
        for (var j = 0; j < entries.length; j += 1) {

            if (entries[j][MAIL_FIELDS.ID] === mail_set.msgs[i][MAIL_FIELDS.ID]) {
                isMatch = 1;
                break;
            }
        }

        if (isMatch === 0) {
            var record = [
                mail_set.msgs[i][MAIL_FIELDS.ID],
                mail_set.msgs[i][MAIL_FIELDS.DATE],
                mail_set.msgs[i][MAIL_FIELDS.SUBJECT],
                mail_set.msgs[i][MAIL_FIELDS.FROM],
                mail_set.msgs[i][MAIL_FIELDS.TO],
                mail_set.msgs[i][MAIL_FIELDS.CC],
                mail_set.msgs[i][MAIL_FIELDS.BODY],
                mail_set.msgs[i][MAIL_FIELDS.NORM_BODY],
            ];

            sheet.getRange(lastRow + row_cnt + 1, mail_set.start_col, 1, record.length)
                .setValues([record]);
            row_cnt += 1;
        }
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 取得メールをシートに配置する
 * @throws {Error} 過去分取得日数が不正です
 */
function MAILUTIL_load_messages() {

    var conf = MAILUTIL_load_config(CONFIG_SET);

    var SS_UI;
    try {
        SS_UI = SpreadsheetApp.getUi();
    } catch (e) {
        SS_UI = null;
    }

    if (SS_UI !== null) {
        var res = NLCUTIL_open_dialog("取得", "メールの取得を開始します。よろしいですか？", SS_UI.ButtonSet.OK_CANCEL);
        if (res === SS_UI.Button.CANCEL) {
            NLCUTIL_open_dialog("取得", "メールの取得を中止しました。", SS_UI.ButtonSet.OK);
            return;
        }
        var msg = "メールの取得を開始しました。";
        NLCUTIL_open_dialog("分類", msg, SS_UI.ButtonSet.OK);
    }

    var mail_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_col: conf.sheet_conf.start_col,
        start_row: conf.sheet_conf.start_row,
        query: conf.sheet_conf.query,
        search_limit: conf.sheet_conf.search_limit,
        exc_res: conf.exc_conf.re_list,
        msgs: [],
    };

    var from_date;
    if (conf.sheet_conf.ago_days === 0) {
        from_date = MAILUTIL_get_newest(mail_set);
        from_date.setHours(0, 0, 0, 0);
    } else if (conf.sheet_conf.ago_days > 0) {
        from_date = new Date();
        from_date.setHours(0, 0, 0, 0);
        from_date.setDate(from_date.getDate() - conf.sheet_conf.ago_days);
    } else {
        throw new Error("過去メール取得日数が不正です");
    }

    mail_set.from_date = from_date;

    var msgs = MAILUTIL_get_messages(mail_set);

    mail_set.msgs = msgs;

    mail_set = MAILUTIL_trim_exc(mail_set);

    MAILUTIL_update_data(mail_set);
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 学習処理
 * @param       {TrainSet} train_set      学習情報
 * @param       {String} creds_username クレデンシャル
 * @param       {String} creds_password クレデンシャル
 * @return      {TrainResult} 学習結果
 * @throws      {Error}  データシートが不明です
 * @throws      {Error}  学習・分類対象が不正です
 */
function MAILUTIL_train(train_set, creds_username, creds_password) {

    Logger.log("### MAILUTIL_train");

    var clfs = NLCAPI_get_classifiers(creds_username, creds_password);
    if (clfs.status !== 200) {
        return {
            status: clfs.code,
            description: clfs.status,
        };
    }

    var sheet = SELF_SS.getSheetByName(train_set.ws_name);
    if (sheet === null) {
        throw new Error('データシートが不明です');
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries;
    if (lastRow < train_set.start_row || lastCol < train_set.start_col || lastCol < train_set.class_col) {
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

        var subject_text = NLCUTIL_norm_text(String(entries[cnt][MAIL_FIELDS.SUBJECT]))
            .trim();
        var body_text = NLCUTIL_norm_text(String(entries[cnt][MAIL_FIELDS.NORM_BODY]))
            .trim();

        var train_text;
        if (train_set.train_column === TRAIN_COLUMN.SUBJECT) {
            train_text = subject_text;
        } else if (train_set.train_column === TRAIN_COLUMN.BODY) {
            train_text = body_text;
        } else if (train_set.train_column === TRAIN_COLUMN.BOTH) {
            if (subject_text === "") {
                train_text = body_text;
            } else if (body_text === "") {
                train_text = subject_text;
            } else {
                train_text = subject_text + " " + body_text;
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

    var result = {
        status: res.status,
        nlc: res,
        rows: row_cnt,
        version: new_version,
    };
    return result;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 対象分類器を学習する
 * @param       {Integer} clf_no 分類器番号
 */
function MAILUTIL_train_set(clf_no) {

    Logger.log("### MAILUTIL_train_set", clf_no);

    var conf = MAILUTIL_load_config(CONFIG_SET);

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

    var train_result = MAILUTIL_train(train_set, CREDS.username, CREDS.password);

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
 * 全分類器を学習
 */
function MAILUTIL_train_all() {

    var conf = MAILUTIL_load_config(CONFIG_SET);

    var SS_UI;
    try {
        SS_UI = SpreadsheetApp.getUi();
    } catch (e) {
        SS_UI = null;
    }

    if (SS_UI != null) {
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
        MAILUTIL_train_set(i);
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 全分類器で分類
 * @throws      {Error}  学習・分類対象が不正です
 * @throws      {Error}  データシートが不明です
 */
function MAILUTIL_classify_all() {

    Logger.log("### MAILUTIL_classify_all");

    var conf = MAILUTIL_load_config(CONFIG_SET);

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
    var res_rows = [0, 0, 0];
    for (var cnt = 0; cnt < entries.length; cnt += 1) {

        var subject_text = NLCUTIL_norm_text(String(entries[cnt][MAIL_FIELDS.SUBJECT]))
            .trim();
        var body_text = NLCUTIL_norm_text(String(entries[cnt][MAIL_FIELDS.NORM_BODY]))
            .trim();

        var test_text;
        if (conf.sheet_conf.train_column === TRAIN_COLUMN.SUBJECT) {
            test_text = subject_text;
        } else if (conf.sheet_conf.train_column === TRAIN_COLUMN.BODY) {
            test_text = body_text;
        } else if (conf.sheet_conf.train_column === TRAIN_COLUMN.BOTH) {
            if (subject_text === "") {
                test_text = body_text;
            } else if (body_text === "") {
                test_text = subject_text;
            } else {
                test_text = subject_text + " " + body_text;
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

            if (clf_ids[j].status !== "Available") {
                continue;
            }

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
                var t = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
                sheet.getRange(conf.sheet_conf.start_row + cnt, conf.sheet_conf.restime_col[j], 1, 1)
                    .setValue(t);
                var c = nlc_res.body.classes[0].confidence;
                sheet.getRange(conf.sheet_conf.start_row + cnt, conf.sheet_conf.resconf_col[j], 1, 1)
                    .setValue(c);

                updates += 1;
                res_rows[j] += 1;
                upd_flg[j] = 1;
            }
        }

        if (updates > 0) {
            var record = sheet.getRange(conf.sheet_conf.start_row + cnt, 1, 1, lastCol)
                .getValues();
            NLCUTIL_check_notify(notif_set, record[0], upd_flg);
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
