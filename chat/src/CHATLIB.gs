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
/* globals CONF_INDEX */
/* globals CONFIG_SET */
/* globals SS_ID */
/* globals SELF_SS */
/* globals NB_CLFS */
/* globals CLFNAME_PREFIX */
/* globals CLF_SEP */
/* globals NLCUTIL_load_creds */
/* globals NLCUTIL_log_classify */
/* globals NLCUTIL_log_train */
/* globals NLCUTIL_select_clf */
/* globals NLCUTIL_norm_text */
/* globals NLCUTIL_open_dialog */
/* globals NLCAPI_post_classify */
/* globals NLCAPI_get_classifiers */
/* globals NLCUTIL_clf_vers */
/* globals NLCAPI_post_classifiers */
/* globals NLCAPI_delete_classifier */
/* globals NLCUTIL_exec_check_clfs */
/* globals NLCUTIL_set_trigger */

/**
 * 応答設定フィールドインデックス
 * @type {Object}
 */
var CONV_INDEX = {
    result1: 0,
    resconf1: 1,
    result2: 2,
    resconf2: 3,
    result3: 4,
    resconf3: 5,
    message: 6,
};

/**
 * LINE応答メッセージ用URL
 * @type {String}
 */
var WORKCHAT_REPLY_URL = 'https://graph.facebook.com/v2.6/me/messages'; // eslint-disable-line no-unused-vars
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ChatCreds クレデンシャル情報
 * @property {String} username NLCのユーザー名
 * @property {String} password NLCのパスワード
 * @property {String} url      NLCのエンドポイント
 * @property {String} channel_access_token LINEチャネルアクセストークン
 */
/**
 * クレデンシャル情報の取得
 * <p>利用するNLCインスタンスのクレデンシャル情報をスクリプトプロパティから取得する</p>
 * <p>利用するLINEアカウントのアクセストークンをスクリプトプロパティから取得する</p>
 * @return {ChatCreds} クレデンシャル情報
 * @throws {Error}  NLCクレデンシャルが不明です
 * @throws {Error}  LINEクレデンシャルが不明です
 */
function CHATUTIL_load_creds() { // eslint-disable-line no-unused-vars

    var scriptProps = PropertiesService.getScriptProperties();

    var creds = {};

    creds['channel_access_token'] = scriptProps.getProperty('WORKPLACE_ACCESS_TOKEN');

    if (creds.channel_access_token === null) {
        throw new Error('Workplace アクセストークンが設定されていません');
    }

    return creds;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ChatConfig 設定情報
 * @property {SheetConf} sheet_conf データシート設定
 * @property {ConvConf}  conv_conf 応答設定
 */
/**
 * 設定情報の取得
 * @param       {ConfigMeta} config_set メタデータ
 * @return      {ChatConfig} 設定情報
 * @throws      {Error}  設定シートが不明です
 * @throws      {Error}  設定シートに問題があります
 */
function CHATUTIL_load_config(config_set) {

    var sheet = SELF_SS.getSheetByName(config_set.ws_name);
    if (sheet === null) {
        throw new Error("設定シートが不明です");
    }

    var nb_conf = Object.keys(CONF_INDEX)
        .length;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow < (config_set.st_start_row + nb_conf) - 1 || lastCol < config_set.st_start_col) {
        throw new Error("設定シートに問題があります");
    }

    var records = sheet.getRange(config_set.st_start_row, config_set.st_start_col, nb_conf, 1)
        .getValues();

    var i = 0;
    var sheet_conf = {
        ws_name: records[CONF_INDEX.ws_name][i],
        start_row: records[CONF_INDEX.start_row][i],
        start_col: records[CONF_INDEX.start_col][i],
        text_col: records[CONF_INDEX.text_col][i],
        intent_col: [
            records[CONF_INDEX.intent1_col][i],
            records[CONF_INDEX.intent2_col][i],
            records[CONF_INDEX.intent3_col][i],
        ],
        result_col: [
            records[CONF_INDEX.result1_col][i],
            records[CONF_INDEX.result2_col][i],
            records[CONF_INDEX.result3_col][i],
        ],
        confidence_col: [
            records[CONF_INDEX.resconf1_col][i],
            records[CONF_INDEX.resconf2_col][i],
            records[CONF_INDEX.resconf3_col][i],
        ],
        restime_col: [
            records[CONF_INDEX.restime1_col][i],
            records[CONF_INDEX.restime2_col][i],
            records[CONF_INDEX.restime3_col][i],
        ],
        log_ws: records[CONF_INDEX.log_ws][i],
        start_msg: records[CONF_INDEX.start_msg][i],
        other_msg: records[CONF_INDEX.other_msg][i],
        error_msg: records[CONF_INDEX.error_msg][i],
        avatar_url: records[CONF_INDEX.avatar_url][i],
        giveup_msg: records[CONF_INDEX.giveup_msg][i],
    };

    var conv_conf = {};
    conv_conf["ws_name"] = records[CONF_INDEX.conv_ws][0];

    return {
        sheet_conf: sheet_conf,
        conv_conf: conv_conf,
    };
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 応答設定の取得
 * @param       {ConfigMeta} config_set メタデータ
 * @return      {Config} 応答設定
 * @throws      {Error}  応答設定シートに問題があります
 */
function CHATUTIL_load_conv_rules(config_set) {

    var sheet = SELF_SS.getSheetByName(config_set.ws_name);
    if (sheet === null) {
        return [];
    }

    var nb_conf = Object.keys(CONV_INDEX)
        .length;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastCol < (config_set.start_col + nb_conf) - 1) {
        throw new Error("応答設定シートに問題があります");
    }

    if (lastRow < config_set.start_row) {
        return [];
    }

    var records = sheet.getRange(config_set.start_row, config_set.start_col, (lastRow - config_set.start_row) + 1, nb_conf)
        .getValues();

    var rules = [];
    for (var i = 0; i < records.length; i += 1) {

        rules.push({
            res_int: [
                String(records[i][CONV_INDEX.result1]),
                String(records[i][CONV_INDEX.result2]),
                String(records[i][CONV_INDEX.result3]),
            ],
            res_conf: [
                String(records[i][CONV_INDEX.resconf1]),
                String(records[i][CONV_INDEX.resconf2]),
                String(records[i][CONV_INDEX.resconf3]),
            ],
            message: records[i][CONV_INDEX.message],
        });
    }

    return rules;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 展開用辞書
 * @type {Object} EnvVars
 * @property {String} input 入力テキスト
 * @property {String} date 現在日付
 * @property {String} time 現在時刻
 */
/**
 * 展開用辞書
 * @type {Object} ExpandResult
 * @property {String} code
 * @property {String} text
 */
/**
 * 埋め込みタグ展開
 * @param       {String} p_temp 対象文字列
 * @param       {EnvVars} p_dict 辞書
 * @return      {Object} 展開結果
 */
function CHATUTIL_expand_tags(p_temp, p_dict) {

    var xbody = p_temp;
    var buf = "";
    Object.keys(p_dict)
        .forEach(function (key) {
            buf = xbody.replace(new RegExp('\\[\\[#' + key + '\\]\\]', 'g'), p_dict[key]);
            xbody = buf;
        });

    xbody.match(new RegExp('\\[\\[#.+\\]\\]', 'g'));

    return {
        code: 'OK',
        text: xbody,
    };
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * メッセージ選択
 * @param       {ConvSet} conv_set    応答設定
 * @param       {Object} res_classes 分類結果
 * @return      {String[]} 応答メッセージ
 */
function CHATUTIL_select_message(conv_set, res_classes) {

    Logger.log("### CHATUTIL_select_messsage");

    var messages = [];
    var match_cnt = 0;
    for (var i = 0; i < conv_set.rules.length; i += 1) {

        var chk_cnt = 0;
        for (var j = 0; j < NB_CLFS; j += 1) {

            if (conv_set.rules[i].res_int[j] === "") {
                chk_cnt += 1;
            } else {
                if (res_classes[j].class_name === conv_set.rules[i].res_int[j]) {
                    if (conv_set.rules[i].res_conf[j] === "") {
                        chk_cnt += 1;
                    } else {
                        if (res_classes[j].confidence >= conv_set.rules[i].res_conf[j]) {
                            chk_cnt += 1;
                        }
                    }
                }
            }
        }

        if (chk_cnt === NB_CLFS) {

            var env_vars = {
                input: conv_set.input_text,
                date: Utilities.formatDate(new Date(), "JST", "yyyy年MM月dd日"),
                time: Utilities.formatDate(new Date(), "JST", "HH時mm分ss秒"),
            };

            var res = CHATUTIL_expand_tags(conv_set.rules[i].message, env_vars);
            messages.push(res.text);
            match_cnt += 1;
            break;
        }
    }

    if (match_cnt === 0) {
        messages.push(conv_set.other_msg);
    }
    return messages;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 会話記録
 * @param       {ConvSet} conv_set    応答設定
 * @param       {Object} res_classes 分類結果
 */
function CHATUTIL_store_dialog(conv_set, res_classes) {

    var sheet = SELF_SS.getSheetByName(conv_set.ws_name);
    if (sheet === null) {
        sheet = SELF_SS.insertSheet(conv_set.ws_name);
    }

    var lastRow = sheet.getLastRow();
    lastRow += 1;
    if (lastRow < conv_set.start_row) {
        lastRow = conv_set.start_row;
    }

    var record = [
        conv_set.timestamp,
        conv_set.input_text,
        conv_set.messages,
        "", res_classes[0].class_name, res_classes[0].confidence, res_classes[0].timestamp,
        "", res_classes[1].class_name, res_classes[1].confidence, res_classes[1].timestamp,
        "", res_classes[2].class_name, res_classes[2].confidence, res_classes[2].timestamp,
    ];
    sheet.appendRow(record);

}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 応答記録
 * @param       {String} input   入力情報
 * @param       {String} res_msg 応答メッセージ
 */
function CHATUTIL_store_reply(input, res_msg) { // eslint-disable-line no-unused-vars

    var timestamp = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");

    var conf = CHATUTIL_load_config(CONFIG_SET);

    var sheet = SELF_SS.getSheetByName(conf.sheet_conf.ws_name);
    if (sheet === null) {
        sheet = SELF_SS.insertSheet(conf.sheet_conf.ws_name);
    }

    var lastRow = sheet.getLastRow();
    lastRow += 1;
    if (lastRow < conf.sheet_conf.start_row) {
        lastRow = conf.sheet_conf.start_row;
    }

    var record = [
        timestamp,
        input,
        res_msg,
        "", "", "", "",
        "", "", "", "",
        "", "", "", "",
    ];
    sheet.appendRow(record);

}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * メッセージ送信
 * @param       {String} input_text ユーザー入力
 * @return      {Object} 応答メッセージ
 */
function CHATUTIL_send_message(input_text) { // eslint-disable-line no-unused-vars

    Logger.log("CHATUTIL_send message");

    var timestamp = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");

    var conf = CHATUTIL_load_config(CONFIG_SET);

    var conv_conf = {
        ss_id: SS_ID,
        ws_name: conf.conv_conf.ws_name,
        start_col: CONFIG_SET.conv_start_col,
        start_row: CONFIG_SET.conv_start_row,
    };

    var rules = CHATUTIL_load_conv_rules(conv_conf);

    var conv_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_col: conf.sheet_conf.start_col,
        start_row: conf.sheet_conf.start_row,
        rules: rules,
        other_msg: conf.sheet_conf.other_msg,
        input_text: input_text,
    };

    var CREDS;
    var res_classes;
    var msgs;
    try {
        CREDS = NLCUTIL_load_creds();
    } catch (e) {
        msgs = [];
        msgs.push(conf.sheet_conf.error_msg);
        conv_set.input_text = input_text;
        conv_set.messages = msgs.join('\n');
        conv_set.timestamp = timestamp;

        res_classes = [
            {
                class_name: 'N/A',
                confidence: 0,
                timestamp: timestamp
            },
            {
                class_name: 'N/A',
                confidence: 0,
                timestamp: timestamp
            },
            {
                class_name: 'N/A',
                confidence: 0,
                timestamp: timestamp
            },
        ];
        CHATUTIL_store_dialog(conv_set, res_classes);
        return {
            response: msgs,
        };
    }

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
        text_col: conf.sheet_conf.text_col,
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

    // ３つの分類器にリクエストを投げる
    res_classes = [];
    var nlc_res;
    //var err_res;
    var has_error = 0;
    for (var j = 0; j < NB_CLFS; j += 1) {

        if (clf_ids[j].status !== "Available") {
            res_classes.push({
                class_name: "",
                confidence: "",
                timestamp: "",
            });
            if (clf_ids[j].status !== 'Nothing') {
                has_error = 1;
            }
            continue;
        }

        nlc_res = NLCAPI_post_classify(CREDS.username, CREDS.password, clf_ids[j].id, input_text);
        if (nlc_res.status !== 200) {
            //err_res = nlc_res;
            has_error = 2;
        } else {
            res_classes.push({
                class_name: nlc_res.body.top_class,
                confidence: nlc_res.body.classes[0].confidence,
                timestamp: Utilities.formatDate(new Date(nlc_res.from), "JST", "yyyy/MM/dd HH:mm:ss"),
            });
        }
    }

    msgs = [];
    if (has_error !== 0) {
        msgs.push(conf.sheet_conf.error_msg);
    } else {
        msgs = CHATUTIL_select_message(conv_set, res_classes);
    }

    conv_set.input_text = input_text;
    conv_set.messages = msgs.join('\n');
    conv_set.timestamp = timestamp;

    CHATUTIL_store_dialog(conv_set, res_classes);

    // 履歴をシートに保存
    var result = {
        response: msgs,
    };
    return result;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 分類器の学習
 * @param       {TrainSet} train_set 学習設定
 * @param       {LogSet} log_set   ログ設定
 * @throws      {Error}  データシートが不明です
 */
function CHATUTIL_train(train_set, log_set) {

    var train_result;

    var CREDS = NLCUTIL_load_creds();

    var clfs = NLCAPI_get_classifiers(CREDS.username, CREDS.password);
    if (clfs.status !== 200) {
        train_result = {
            status: clfs.status,
            description: clfs.error,
        };
        NLCUTIL_log_train();
        return;
    }

    var sheet = SELF_SS.getSheetByName(train_set.ws_name);
    if (sheet === null) {
        throw new Error("データシートが不明です");
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries;
    if (lastRow < train_set.start_row || lastCol < train_set.start_col || lastCol < train_set.class_col) {
        entries = [];
    } else {
        entries = sheet.getRange(train_set.start_row, 1, (lastRow - train_set.start_row) + 1, lastCol)
            .getValues();
    }

    var row_cnt = 0;
    var csvString = '';
    for (var i = 0; i < entries.length; i += 1) {

        var class_name = NLCUTIL_norm_text(String(entries[i][train_set.class_col - 1]));
        if (class_name.length === 0) continue;

        var train_text = String(entries[i][train_set.text_col - train_set.start_col]);
        train_text = NLCUTIL_norm_text(train_text);

        if (train_text.length === 0) continue;

        if (train_text.length > 1024) {
            train_text = train_text.substring(0, 1024);
        }

        csvString = csvString + '"' + train_text + '","' + class_name + '"' +
            "\r\n";
        row_cnt += 1;
    }

    if (row_cnt === 0) {
        train_result = {
            status: 0,
            description: "学習データなし",
        };
        NLCUTIL_log_train(log_set, train_set, train_result);
        return;
    }

    var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, train_set.clf_name);

    var new_version = (clf_info.max_ver + 1);
    var clf_name = train_set.clf_name + CLF_SEP + new_version;

    var nlc_res = NLCAPI_post_classifiers(CREDS.username, CREDS.password, csvString, clf_name, 'ja');

    if (clf_info.count >= 2 && nlc_res.status === 200) {

        NLCAPI_delete_classifier(CREDS.username, CREDS.password, clf_info.clfs[clf_info.min_ver].classifier_id);

    }

    NLCUTIL_exec_check_clfs();
    NLCUTIL_set_trigger('NLCUTIL_exec_check_clfs', 1);

    train_result = {
        status: nlc_res.status,
        nlc: nlc_res,
        rows: row_cnt,
        version: new_version,
    };
    NLCUTIL_log_train(log_set, train_set, train_result);
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * [CHATUTIL_train_set description]
 * @param       {Integer} clf_no 分類器番号
 */
function CHATUTIL_train_set(clf_no) {

    var conf = CHATUTIL_load_config(CONFIG_SET);

    var train_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_row: conf.sheet_conf.start_row,
        start_col: conf.sheet_conf.start_col,
        text_col: conf.sheet_conf.text_col,
        class_col: conf.sheet_conf.intent_col[clf_no - 1],
        clf_no: clf_no,
        clf_name: CLFNAME_PREFIX + String(clf_no),
    };

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row,
    };

    CHATUTIL_train(train_set, log_set);
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 全分類器を学習
 */
function CHATUTIL_train_all() { // eslint-disable-line no-unused-vars

    var SS_UI;
    try {
        SS_UI = SpreadsheetApp.getUi();
    } catch (e) {
        SS_UI = null;
    }

    var CREDS = NLCUTIL_load_creds()

    var conf = CHATUTIL_load_config(CONFIG_SET);

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
        CHATUTIL_train_set(i);
    }
}
// ----------------------------------------------------------------------------
