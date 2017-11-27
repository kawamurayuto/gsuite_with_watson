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
/* globals SS_ID */
/* globals SELF_SS */
/* globals CONFIG_SET */
/* globals NB_CLFS */
/* globals CONF_INDEX */


/**
 * 分類器のステータス
 * @type {Object}
 * @property {String} AVAILABLE 利用可能
 * @property {String} TRAINING トレーニング中
 * @property {String} NOTHING  利用不可
 */
var NLCAPI_CLF_STATUS = {
    AVAILABLE: 'Available',
    TRAINING: 'Training',
    NOTHING: 'Nothing',
};

/**
 * 分類器名のプリフィクス
 * @type {String}
 */
var CLFNAME_PREFIX = 'CLF';

/**
 * 分類器名のセパレータ
 * @type {String}
 */
var CLF_SEP = '#__#';

/**
 * 通知オプション
 * @type {Object}
 * @property {String} ON オン
 * @property {String} OFF オフ
 */
var NOTIF_OPT = {
    ON: 'On',
    OFF: 'Off',
};

/**
 * 通知ルールレコードのフィールドインデックス
 * @type {Object} NOTIF_INDEX
 * @property {Integer} result1 分類結果1
 * @property {Integer} result2 分類結果2
 * @property {Integer} result3 分類結果3
 * @property {Integer} from  送信元メールアドレス
 * @property {Integer} to 送信先メールアドレス
 * @property {Integer} cc CCメールアドレス
 * @property {Integer} bcc BCCメールアドレス
 * @property {Integer} subject 件名
 * @property {Integer} body 本文
 */
var NOTIF_INDEX = {
    result1: 0,
    result2: 1,
    result3: 2,
    from: 3,
    to: 4,
    cc: 5,
    bcc: 6,
    subject: 7,
    body: 8,
};
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} SheetSet シート基本情報
 * @property {String} ss_id スプレッドシートID
 * @property {String} ws_name シート名
 * @property {Integer} start_col 開始列
 * @property {Integer} start_row 開始行
 */
/**
 * @typedef {Object} Creds クレデンシャル情報
 * @property {String} username ユーザー名
 * @property {String} password パスワード
 * @property {String} url      エンドポイント
 */
/**
 * クレデンシャル情報の取得
 * <p>利用するNLCインスタンスのクレデンシャル情報をスクリプトプロパティから取得する</p>
 * @return {Creds} クレデンシャル情報
 * @throws {Error}  クレデンシャルが不明です
 */
function NLCUTIL_load_creds() {

    var scriptProps = PropertiesService.getScriptProperties();

    var creds = {};
    creds['url'] = scriptProps.getProperty('CREDS_URL');
    creds['username'] = scriptProps.getProperty('CREDS_USERNAME');
    creds['password'] = scriptProps.getProperty('CREDS_PASSWORD');

    if (creds.username === null || creds.password === null) {
        throw new Error('クレデンシャルが不明です');
    }

    return creds;
}
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
 * @property {String} ws_name データシート名
 * @property {Integer} start_row 定義開始行
 * @property {Integer} start_col 定義開始列
 * @property {Integer[]} intent_col インテント列1to3
 * @property {Integer[]} result_col 分類結果列1to3
 * @property {Integer[]} resconf_col 確信度列1to3
 * @property {Integer[]} restime_col 分類日時列1to3
 * @property {String} log_ws ログシート名
 */
/**
 * @typedef {Object} NotifConf 通知設定
 * @property {String} notif_opt 通知オプション{On,Off}
 * @property {String} notif_ws 設定シート名
 */
/**
 * @typedef {Object} Config 設定情報
 * @property {SheetConf} sheet_conf データシート設定
 * @property {NotifConf} notif_conf 通知設定
 */
/**
 * 設定情報の取得
 * <p>メタデータを元にユーザーの設定情報を取得する</p>
 * @param   {ConfigMeta}      config_set コンフィグメタデータ
 * @return {Config} コンフィグ
 * @throws {Error}  設定シートが不明です
 * @throws {Error}  設定シートに問題があります
 * @throws {Error}  学習・分類対象列が不正です'
 */
function NLCUTIL_load_config(config_set) {

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

    var records = sheet.getRange(config_set.st_start_row, config_set.st_start_col, nb_conf, 1)
        .getValues();

    var i = 0;
    var sheet_conf = {
        ws_name: records[CONF_INDEX.ws_name][i],
        start_row: records[CONF_INDEX.start_row][i],
        start_col: records[CONF_INDEX.start_col][i],
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
        resconf_col: [
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
    };

    if (CONF_INDEX.text_col) {
        sheet_conf.text_col = records[CONF_INDEX.text_col][i];
        if (typeof sheet_conf.text_col !== 'number') {
            throw new Error('学習・分類対象列が不正です');
        }
    }

    var notif_conf = {};
    if (CONF_INDEX.notif_opt) {
        notif_conf["option"] = records[CONF_INDEX.notif_opt][0];
    }
    if (CONF_INDEX.notif_ws) {
        notif_conf["ws_name"] = records[CONF_INDEX.notif_ws][0];
    }

    return {
        sheet_conf: sheet_conf,
        notif_conf: notif_conf,
    };
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ExpandResult 展開結果
 * @property {String} code 結果コード
 * @property {String} text 結果テキスト
 */
/**
 * 通知ルール用の埋め込みタグを展開する
 * <p>対象テキスト中に埋め込まれたタグをインデックスに該当する対象フィールドに置換する</p>
 * <p>埋め込みタグの形式 [[#インデックス]] ※インデックスは1以上の列として有効な整数</p>
 * @param {String} target 変換対象テキスト
 * @param {String[]} fields 展開対象フィールド
 * @return {ExpandResult} 展開結果
 */
function NLCUTIL_expand_tags(target, fields) {

    var xbody = target;
    var buf = '';

    for (var idx = 0; idx < fields.length; idx += 1) {
        buf = xbody.replace(new RegExp('\\[\\[#' + String(idx + 1) + '\\]\\]', 'g'), fields[idx]);
        xbody = buf;
    }

    xbody.match(new RegExp('\\[\\[.+\\]\\]', 'g'));

    return {
        code: 'OK',
        text: xbody,
    };
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} Mail メール設定
 * @property {String} from 送信元メールアドレス
 * @property {String} to 送信先メールアドレス
 * @property {String} cc ccメールアドレス
 * @property {String} bcc bccメールアドレス
 * @property {String} subject 件名
 * @property {String} body 本文
 */
/**
 * @typedef {Object} Rules 通知条件
 * @property {String[]} res_int 分類結果1to3
 * @property {Mail} mail メール設定
 */
/**
 * 通知条件の取得
 * <p>
 * </p>
 * @param       {SheetSet} config_set 設定情報
 * @return      {Rules[]} 通知条件
 * @throws      {Error}   通知設定シートに問題があります
 */
function NLCUTIL_load_notif_rules(config_set) {

    var sheet = SELF_SS.getSheetByName(config_set.ws_name);
    if (sheet === null) {
        return [];
    }

    var nb_conf = Object.keys(NOTIF_INDEX)
        .length;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastCol < (config_set.start_col + nb_conf) - 1) {
        return [];
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
                String(records[i][NOTIF_INDEX.result1]),
                String(records[i][NOTIF_INDEX.result2]),
                String(records[i][NOTIF_INDEX.result3]),
            ],
            mail: {
                from: records[i][NOTIF_INDEX.from],
                to: records[i][NOTIF_INDEX.to],
                cc: records[i][NOTIF_INDEX.cc],
                bcc: records[i][NOTIF_INDEX.bcc],
                subject: records[i][NOTIF_INDEX.subject],
                body: records[i][NOTIF_INDEX.body],
            },
        });
    }
    return rules;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 数式評価文字抑止
 * <p>セルの数式として評価される特殊文字をシングルクォートでエスケープする
 * @param       {String} text 対象テキスト
 * @return      {String} 編集結果
 */
function NLCUTIL_escape_formula(text) {

    var first = text.substring(0, 1);

    var result = text;
    if (first === '=' || first === '+') {
        result = "'" + result;
    }

    return result;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * NLC学習用にテキストをノーマライズする
 * <ul>
 * <li>シングルとダブルのクォートはダブルクォートでエスケープ</li>
 * <li>タブ、改行は削除</li>
 * <li>前後の空白をトリミング</li>
 * </ul>
 * @param       {String} target 変換対象テキスト
 * @return      {String} 変換結果
 */
function NLCUTIL_norm_text(target) {

    var res = target
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\"'")
        .replace(/"/g, '""')
        .replace(/(0x0D)/g, '')
        .replace(/(0x0A)/g, '')
        .replace(/\t/g, '')
        .replace(/\r?\n/g, '')
        .replace(/\r/g, '')
        .replace(/\n/g, '');

    res = res.trim();

    return res;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ClassifierInfoPayload 分類器情報
 * @property {String} classifier_id ID
 * @property {String} name 名称
 * @property {String} created 作成日
 */
/**
 * @typedef {Object} ClfVers 分類器バージョン一覧
 * @property {Integer} count 件数
 * @property {Integer} min_ver 最小バージョン
 * @property {Integer} max_ver 最大バージョン
 * @property {ClassifierInfoPayload[]} clfs 分類器情報
 */
/**
 * 分類器のバージョン一覧を取得
 * <p>分類器一覧から分類器名(ex. CLF1)に該当するバージョン一覧を生成する</p>
 * @param       {ClassifierInfoPayload[]} clf_list 分類器一覧
 * @param       {String} target_name 分類器名
 * @return      {ClfVers}  バージョン一覧
 */
function NLCUTIL_clf_vers(clf_list, target_name) {

    var clfs = [];
    var max_ver = 0;
    var min_ver = 99999999;
    var count = 0;

    for (var i = 0; i < clf_list.length; i += 1) {

        var base = clf_list[i].name.split(CLF_SEP);
        if (base[1] === null) continue;

        if (target_name === base[0]) {

            clfs[parseInt(base[1], 10)] = clf_list[i];
            count += 1;

            if (parseInt(base[1], 10) > max_ver) {
                max_ver = parseInt(base[1], 10);
            }
            if (parseInt(base[1], 10) < min_ver) {
                min_ver = parseInt(base[1], 10);
            }
        }
    }

    var result = {
        count: count,
        min_ver: min_ver,
        max_ver: max_ver,
        clfs: clfs,
    };
    return result;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} ClfInfo 分類器情報
 * @property {String} clf_id ID
 * @property {String} status ステータス
 */
/**
 * 利用可能な分類器の最新バージョンを取得する
 * <p>ステータスコードが200以外の場合、IDに空白、ステータスにNLCの実行ステータスをセットする</p>
 * <p>バージョン件数が０件の場合、IDに空白、ステータスに'Nothing'をセットする</p>
 * <p>各バージョンの状態を取得する</p>
 * <p>状態が'Available'でバージョンが最新の分類器情報を返す</p>
 * @param       {String} clf_name       分類器名 ex.CLF1
 * @param       {String} creds_username ユーザー名(クレデンシャル)
 * @param       {String} creds_password パスワード(クレデンシャル)
 * @return      {ClfInfo}  分類器情報
 */
function NLCUTIL_select_clf(clf_name, creds_username, creds_password) {

    var clfs = NLCAPI_get_classifiers(creds_username, creds_password);
    if (clfs.status !== 200) {
        return {
            clf_id: '',
            status: clfs.status,
        };
    }

    var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, clf_name);
    if (clf_info.count === 0) {
        return {
            clf_id: '',
            status: 'Nothing',
        };
    }

    var clf;
    var res;
    for (var i = clf_info.max_ver; i >= clf_info.min_ver; i -= 1) {

        clf = clf_info.clfs[i];

        res = NLCAPI_get_classifier(creds_username, creds_password, clf.classifier_id);
        if (res.body.status === 'Available') {
            return {
                clf_id: clf.classifier_id,
                status: res.body.status,
            };
        }
    }

    return {
        clf_id: clf.classifier_id,
        status: res.body.status,
    };
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * メールの送信
 * <p>GmailAppのsendEmailを実行する</p>
 * @param       {Mail} mail_set メール通知設定
 * @throws      {Error} メールの送信に失敗しました
 */
function NLCUTIL_send_mail(mail_set) {

    Logger.log('### NLCUTIL_send_mail');

    try {
        GmailApp.sendEmail(
            mail_set.to,
            mail_set.subject,
            mail_set.body, {
                from: mail_set.from,
                cc: mail_set.cc,
                bcc: mail_set.bcc,
            }
        );
    } catch (err) {
        Logger.log(err);
        throw new Error('メールの送信に失敗しました');
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * メール通知条件チェック
 * <p>通知条件にマッチした場合はメールを送信する</p>
 * <p>件名と本文の埋め込みタグを展開する</p>
 * <p>インテントがブランクの場合はワイルドカード扱いする</p>
 * @param       {NotifSet} notif_set 通知設定
 * @param       {String[]} record    通知対象データ
 * @param       {Integer[]} upd_flg   更新フラグ
 */
function NLCUTIL_check_notify(notif_set, record, upd_flg) {

    Logger.log('### NLCUTIL_check_notify');

    for (var i = 0; i < notif_set.rules.length; i += 1) {

        var chk_cnt = 0;
        var upd_chk = 0;

        for (var j = 0; j < NB_CLFS; j += 1) {

            if (notif_set.rules[i].res_int[j] === '') {
                chk_cnt += 1;
            } else {
                if (upd_flg[j] === 1) {
                    upd_chk = 1;
                }
                if (record[notif_set.result_cols[j] - 1] === notif_set.rules[i].res_int[j]) {
                    chk_cnt += 1;
                }
            }
        }

        if (chk_cnt === NB_CLFS && upd_chk === 1) {

            var res;
            res = NLCUTIL_expand_tags(notif_set.rules[i].mail.body, record);
            var body = res.text;

            res = NLCUTIL_expand_tags(notif_set.rules[i].mail.subject, record);
            var subject = res.text;

            var mail_set = {
                from: notif_set.rules[i].mail.from,
                to: notif_set.rules[i].mail.to,
                cc: notif_set.rules[i].mail.cc,
                bcc: notif_set.rules[i].mail.bcc,
                subject: subject,
                body: body,
            };

            NLCUTIL_send_mail(mail_set);
        }
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 分類結果ログ出力
 * @param       {Object} log_set     ログ出力設定
 * @param       {Object} test_set    テスト設定
 * @param       {Object} test_result テスト結果
 */
function NLCUTIL_log_classify(log_set, test_set, test_result) {

    Logger.log('### NLCUTIL_log_classify');

    var sheet = SELF_SS.getSheetByName(log_set.ws_name);
    if (sheet === null) {
        sheet = SELF_SS.insertSheet(log_set.ws_name);
    }

    var lastRow = sheet.getLastRow();
    lastRow += 1;
    if (lastRow < log_set.start_row) {
        lastRow = log_set.start_row;
    }

    var record = [];
    if (test_result.status === 200) {
        record = [
            '分類',
            Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
            '成功',
            test_set.clf_no,
            test_set.ws_name,
            test_result.rows,
            test_set.text_col,
            test_set.result_col,
            test_result.nlc.body.classifier_id,
            test_result.status,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record]);
    } else if (test_result.status === 0) {
        record = [
            '分類',
            Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
            '対象なし',
            test_set.clf_no,
            test_set.ws_name,
            test_result.rows,
            test_set.text_col,
            test_set.result_col,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record]);
    } else if (test_result.status === 900) {
        record = [
            '分類',
            Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
            test_result.description,
            test_set.clf_no,
            test_set.ws_name,
            'N/A',
            test_set.text_col,
            test_set.result_col,
            test_result.clf_id,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record]);
    } else {
        if (test_result['nlc']) {
            record = [
                '分類',
                Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
                test_result.description,
                test_set.clf_no,
                test_set.ws_name,
                'N/A',
                test_set.text_col,
                test_set.result_col,
            ];
        } else {
            record = [
                '分類',
                Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
                test_result.description,
                test_set.clf_no,
                test_set.ws_name,
                'N/A',
                test_set.text_col,
                test_set.result_col,
            ];
        }
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record])
            .setFontColor('red');
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 分類器削除ログ出力
 * @param       {Object} log_set    ログ設定
 * @param       {Object} del_set    削除設定
 * @param       {Object} del_result 削除結果
 */
function NLCUTIL_log_delete(log_set, del_set, del_result) {

    Logger.log('### NLCUTIL_log_delete');

    var sheet = SELF_SS.getSheetByName(log_set.ws_name);
    if (sheet === null) {
        sheet = SELF_SS.insertSheet(log_set.ws_name);
    }

    var lastRow = sheet.getLastRow();
    lastRow += 1;
    if (lastRow < log_set.start_row) {
        lastRow = log_set.start_row;
    }

    var record = [];
    if (del_result.status === 200) {
        record = ['削除',
            Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
            '成功',
            del_set.clf_no,
            '',
            '',
            '',
            '',
            del_set.clf_id,
            del_result.status,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record]);
    } else {
        record = ['削除',
            Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
            '失敗',
            del_set.clf_no,
            '',
            '',
            '',
            '',
            del_set.clf_id,
            del_result.status,
            del_result.nlc.body.status_description,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record])
            .setFontColor('red');

    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * ダイアログオープン
 * @param       {String} title   タイトル
 * @param       {String} msg     メッセージ
 * @param       {Object} buttons 配置ボタン
 * @return      {Object} 選択したボタン
 * @throws      {Error}  データシートが不明です
 */
function NLCUTIL_open_dialog(title, msg, buttons) {

    var SS_UI = SpreadsheetApp.getUi();
    var result = SS_UI.alert(title, msg, buttons);
    return result;
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * イベント実行用分類処理
 * @throws {Error} データシートが不明
 */
function NLCUTIL_classify_all() {

    Logger.log('### NLCUTIL_classify_all');

    var conf = NLCUTIL_load_config(CONFIG_SET);

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
        text_col: conf.sheet_conf.text_col,
        notif_set: notif_set,
    };

    // ------------------------------------------------------------------------
    // 分類器情報の取得
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

    // ------------------------------------------------------------------------
    // 対象データの読み込み
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

    // ------------------------------------------------------------------------
    // 分類
    var hasError = 0;
    var err_res;
    var nlc_res;
    var res_rows = [0, 0, 0];
    for (var cnt = 0; cnt < entries.length; cnt += 1) {

        var test_text = entries[cnt][conf.sheet_conf.text_col - conf.sheet_conf.start_col];
        test_text = NLCUTIL_norm_text(String(test_text))
            .trim();
        if (test_text.length === 0) continue;

        if (test_text.length > 1024) {
            test_text = test_text.substring(0, 1024);
        }

        var updates = 0;
        var upd_flg = [0, 0, 0];
        for (var j = 0; j < NB_CLFS; j += 1) {

            if (clf_ids[j].status !== 'Available') continue;

            var result_text;
            if (lastCol < conf.sheet_conf.result_col[j]) {
                result_text = '';
            } else {
                result_text = entries[cnt][conf.sheet_conf.result_col[j] - conf.sheet_conf.start_col];
            }

            if (result_text !== '' && CONFIG_SET.result_override !== true) continue;

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
    // 分類
    // ------------------------------------------------------------------------

    // ------------------------------------------------------------------------
    // ログ出力
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
/**
 * タイマー起動用分類器状態チェック
 */
function NLCUTIL_exec_check_clfs() {

    var clf_set = {
        ss_id: CONFIG_SET.ss_id,
        ws_name: CONFIG_SET.ws_name,
        start_col: CONFIG_SET.clfs_start_col,
        start_row: CONFIG_SET.clfs_start_row,
    };

    NLCUTIL_check_classifiers(clf_set, CREDS);
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * タイマーをセットする
 * @param       {String} func_name セットする関数名
 * @param       {Integer} int_min  実行間隔(分)
 */
function NLCUTIL_set_trigger(func_name, int_min) {

    var triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < triggers.length; i += 1) {
        if (triggers[i].getHandlerFunction() === func_name) {
            return;
        }
    }

    ScriptApp.newTrigger(func_name)
        .timeBased()
        .everyMinutes(int_min)
        .create();
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * タイマーを解除する
 * @param       {String} func_name 解除する関数名
 */
function NLCUTIL_del_trigger(func_name) {

    Logger.log('### NLCUTIL_del_trigger');

    var triggers = ScriptApp.getProjectTriggers();

    for (var i = 0; i < triggers.length; i += 1) {
        if (triggers[i].getHandlerFunction() === func_name) {
            ScriptApp.deleteTrigger(triggers[i]);
        }
    }
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 学習処理
 * @param       {Object} train_set      学習設定
 * @param       {String} creds_username クレデンシャル
 * @param       {String} creds_password クレデンシャル
 * @return      {Object} 学習結果
 * @throws      {Error} データシートが不明です
 */
function NLCUTIL_train(train_set, creds_username, creds_password) {

    var clfs = NLCAPI_get_classifiers(creds_username, creds_password);
    if (clfs.status !== 200) {
        return {
            status: clfs.status,
            description: clfs.error,
        };
    }

    var sheet = SELF_SS.getSheetByName(train_set.ws_name);
    if (sheet === null) {
        throw new Error('データシートが不明です');
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    var entries;
    if (lastRow < train_set.start_row || lastCol < train_set.class_col || lastCol < train_set.text_col) {
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

        var train_text = String(entries[i][train_set.text_col - 1]);
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
    NLCUTIL_set_trigger('NLCUTIL_exec_check_clfs', 1);

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
 * 特定分類器で学習して結果をログに出力する
 * @param       {Integer} clf_no 分類器番号
 */
function NLCUTIL_train_set(clf_no) {

    Logger.log('### NLCUTIL_train_set', clf_no);

    var conf = NLCUTIL_load_config(CONFIG_SET);

    var train_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.ws_name,
        start_row: conf.sheet_conf.start_row,
        start_col: conf.sheet_conf.start_col,
        end_row: -1,
        text_col: conf.sheet_conf.text_col,
        class_col: conf.sheet_conf.intent_col[clf_no - 1],
        clf_no: clf_no,
        clf_name: CLFNAME_PREFIX + clf_no,
    };

    var train_result = NLCUTIL_train(train_set, CREDS.username, CREDS.password);

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
function NLCUTIL_train_all() {

    var conf = NLCUTIL_load_config(CONFIG_SET);

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
        NLCUTIL_train_set(i);
    }

}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 分類器削除
 * @param       {Integer} clf_no 分類器番号
 * @throws      {Error}   サーバーエラーが発生しました
 */
function NLCUTIL_delete_classifier(clf_no) {

    var conf = NLCUTIL_load_config(CONFIG_SET);

    var SS_UI;
    try {
        SS_UI = SpreadsheetApp.getUi();
    } catch (e) {
        SS_UI = null;
    }

    var clfs = NLCAPI_get_classifiers(CREDS.username, CREDS.password);
    if (clfs.status !== 200) {
        throw new Error("サーバーエラーが発生しました");
    }

    var msg;
    var clf_name = CLFNAME_PREFIX + String(clf_no);
    var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, clf_name);
    if (clf_info.count === 0) {
        if (SS_UI != null) {
            msg = "分類器" + String(clf_no) + "は存在しません。";
            NLCUTIL_open_dialog("削除", msg, SS_UI.ButtonSet.OK);
        }
        return;
    }

    if (SS_UI != null) {
        msg = "分類器" + String(clf_no) + "を削除します。よろしいですか？";
        var res = NLCUTIL_open_dialog("削除", msg, SS_UI.ButtonSet.OK_CANCEL);
        if (res === SS_UI.Button.CANCEL) {
            NLCUTIL_open_dialog("削除", "削除を中止しました。", SS_UI.ButtonSet.OK);
            return;
        }
    }

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row,
    };

    for (var i = clf_info.min_ver; i <= clf_info.max_ver; i += 1) {
        var nlc_res = NLCAPI_delete_classifier(CREDS.username, CREDS.password, clf_info.clfs[i].classifier_id);
        var del_set = {
            clf_no: clf_no,
            clf_id: clf_info.clfs[i].classifier_id,
        };
        var del_result = {
            status: nlc_res.status,
            nlc: nlc_res,
        };

        NLCUTIL_log_delete(log_set, del_set, del_result);
    }

    NLCUTIL_exec_check_clfs();
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 分類器1削除
 */
function NLCUTIL_del_clf1() {
    NLCUTIL_delete_classifier(1);
}

/**
 * 分類器2削除
 */
function NLCUTIL_del_clf2() {
    NLCUTIL_delete_classifier(2);
}

/**
 * 分類器3削除
 */
function NLCUTIL_del_clf3() {
    NLCUTIL_delete_classifier(3);
}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * 学習結果ログ出力
 * @param       {Object} log_set   ログ出力設定
 * @param       {Object} train_set 学習設定
 * @param       {Object} res       学習結果
 */
function NLCUTIL_log_train(log_set, train_set, res) {

    var sheet = SELF_SS.getSheetByName(log_set.ws_name);
    if (sheet === null) {
        sheet = SELF_SS.insertSheet(log_set.ws_name);
    }

    var lastRow = sheet.getLastRow();
    lastRow += 1;
    if (lastRow < log_set.start_row) {
        lastRow = log_set.start_row;
    }

    var record = [];
    if (res.status === 200) {
        record = ['学習',
            Utilities.formatDate(new Date(res.nlc.from), 'JST', 'yyyy/MM/dd HH:mm:ss'),
            res.nlc.body.status,
            train_set.clf_no,
            train_set.ws_name,
            res.rows,
            train_set.text_col,
            train_set.class_col,
            res.nlc.body.classifier_id,
            res.status,
            res.nlc.body.status_description,
            res.nlc.body.created,
            res.version,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record]);
    } else if (res.status === 2000) {
        record = ['学習',
            Utilities.formatDate(new Date(res.nlc.from), 'JST', 'yyyy/MM/dd HH:mm:ss'),
            res.nlc.body.status,
            train_set.clf_no,
            '',
            '',
            '',
            '',
            res.nlc.body.classifier_id,
            res.nlc.status,
            res.nlc.body.status_description,
            res.nlc.body.created,
            res.version,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record]);
    } else if (res.status === 0) {
        record = ['学習',
            Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
            res.description,
            train_set.clf_no,
            train_set.ws_name,
            0,
            train_set.text_col,
            train_set.class_col,
        ];
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record]);
    } else {
        // エラー
        if (res['nlc']) {

            record = ['学習',
                Utilities.formatDate(new Date(res.nlc.from), 'JST', 'yyyy/MM/dd HH:mm:ss'),
                res.nlc.body.error,
                train_set.clf_no,
                train_set.ws_name,
                res.rows,
                train_set.text_col,
                train_set.class_col,
                '',
                res.status,
                res.nlc.body.description,
            ];
        } else {
            record = ['学習',
                Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'),
                res.description,
                train_set.clf_no,
                train_set.ws_name,
                'N/A',
                train_set.text_col,
                train_set.class_col,
                '',
                res.status,
            ];
        }
        sheet.getRange(lastRow, log_set.start_col, 1, record.length)
            .setValues([record])
            .setFontColor('red');
    }

}
// ----------------------------------------------------------------------------


// ----------------------------------------------------------------------------
/**
 * @typedef {Object} SheetSet クレデンシャル情報
 * @property {String}  ユーザー名
 */
/**
 * 分類器の状態確認
 * @param       {Object} clf_set メタデータ
 * @param       {Creds} creds   クレデンシャル
 * @throws      {Error} 設定シートが不明です
 */
function NLCUTIL_check_classifiers(clf_set, creds) {

    Logger.log('### NLCUTIL_check_classifiers');

    var sheet = SELF_SS.getSheetByName(clf_set.ws_name);
    if (sheet === null) {
        throw new Error("設定シートが不明です");
    }

    var conf = NLCUTIL_load_config(CONFIG_SET);

    var log_set = {
        ss_id: SS_ID,
        ws_name: conf.sheet_conf.log_ws,
        start_col: CONFIG_SET.log_start_col,
        start_row: CONFIG_SET.log_start_row,
    };

    var curr_stats = sheet.getRange(clf_set.start_row, clf_set.start_col, NB_CLFS, 2)
        .getValues();
    sheet.getRange(clf_set.start_row, clf_set.start_col, NB_CLFS, 2)
        .clear();

    var clfs = NLCAPI_get_classifiers(creds.username, creds.password);

    var all_status = 0;
    for (var cnt = 1; cnt <= NB_CLFS; cnt += 1) {

        if (clfs.status !== 200) {
            sheet.getRange(clf_set.start_row + (cnt - 1), clf_set.start_col, 1, 2)
                .setValues([
                    ['ERROR', clfs.error],
                ]);
        } else {

            var clf_name = CLFNAME_PREFIX + String(cnt);

            var clf_info = NLCUTIL_clf_vers(clfs.body.classifiers, clf_name);

            if (clf_info.count === 0) {
                sheet.getRange(clf_set.start_row + (cnt - 1), clf_set.start_col, 1, 2)
                    .setValues([
                        ['', ''],
                    ]);
                all_status += 1;
            } else {
                var clf = clf_info.clfs[clf_info.max_ver];

                var res = NLCAPI_get_classifier(creds.username, creds.password, clf.classifier_id);
                if (res.body.status === 'Available') {
                    all_status += 1;
                    if (curr_stats[cnt - 1][0] === clf.classifier_id &&
                        curr_stats[cnt - 1][1] === 'Training') {
                        var train_set = {
                            clf_no: cnt,
                        };
                        var train_result = {
                            nlc: res,
                            status: 2000,
                            version: clf_info.max_ver,
                        };
                        NLCUTIL_log_train(log_set, train_set, train_result);
                    }
                }

                sheet.getRange(clf_set.start_row + (cnt - 1), clf_set.start_col, 1, 2)
                    .setValues([
                        [clf.classifier_id, res.body.status],
                    ]);
            }
        }
    }

    // 全てAvailableでタイマー解除
    if (NB_CLFS === all_status) {
        NLCUTIL_del_trigger('NLCUTIL_exec_check_clfs');
    }
}
// ----------------------------------------------------------------------------


// Watson Natural Language Classifier API wrapper
// ----------------------------------------------------
// 定数
var URI_DOMAIN = "https://gateway.watsonplatform.net";
var URI_BASE = "/natural-language-classifier/api";
var URI_APIVERSION = "v1";
var EOL = '\r\n';
// DataCollection
// X-Watson-Learning-Opt-Out true
/**
 * @typedef {Object} APIResult 分類器情報
 * @property {Integer} status レスポンスコード
 * @property {Object} body レスポンスボディ
 * @property {Long} from 開始時刻
 * @property {Long} to 終了時刻
 * @property {Long} delta 応答時間
 */
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * Classifier一覧取得
 * <p>GET /v1/classifiers</p>
 * <p>概要: NLCサービスのClassifier一覧を取得する</p>
 * @param       {String} p_username NLCサービス資格情報のusername
 * @param       {String} p_password NLCサービス資格情報のpassword
 * @return      {APIResult}            APIの実行結果
 */
function NLCAPI_get_classifiers(p_username, p_password) {

    Logger.log("### NLCAPI_get_classifiers");

    // URIビルド
    var resource = "classifiers";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource;
    // リクエストオプション
    var options = {
        headers: {
            Authorization: " Basic " + Utilities.base64Encode(p_username + ":" + p_password),
        },
        method: "get",
        contentType: "application/json",
        muteHttpExceptions: true,
    };

    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    var result;
    result = {
        status: responseCode,
        body: JSON.parse(responseBody),
        from: fromTime.getTime(),
        to: toTime.getTime(),
        delta: delta,
    };
    // 戻り値を返す
    return result;
}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * Classifier生成
 * <p>POST /v1/classifiers</p>
 * <p>概要: 資格情報に該当するNLCサービスにClassifierを新規作成する</p>
 * @param       {String} p_username      NLCサービス資格情報のusername
 * @param       {String} p_password      NLCサービス資格情報のpassword
 * @param       {String} p_training_data トレーニングデータ(CSV)
 * @param       {String} p_classname     クラス名(training_metadata)
 * @param       {String} p_langcode      言語コード(オプション)
 * @return      {APIResult} APIの実行結果
 */
function NLCAPI_post_classifiers(p_username, p_password, p_training_data, p_classname, p_langcode) {

    // 日本語のデフォルト(ja)設定
    var l_langcode = p_langcode || "ja";

    // URIビルド
    var resource = "classifiers";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource;

    // メタデータ
    var training_metadata = {
        language: l_langcode,
        name: p_classname,
    };

    // バウンダリデータの生成
    var boundary = NLCAPI_createBoundary();

    // バイナリデータの生成
    var postbody = Utilities.newBlob('--' + boundary + EOL +
            'Content-Disposition: form-data; name="training_data"; filename="training.csv"' + EOL +
            'Content-Type: application/octet-stream' + EOL +
            'Content-Transfer-Encoding: binary' + EOL + EOL)
        .getBytes();
    postbody = postbody.concat(Utilities.newBlob(p_training_data)
        .getBytes());
    postbody = postbody.concat(Utilities.newBlob('--' + boundary + EOL +
            'Content-Disposition: form-data; name="training_metadata"' + EOL + EOL)
        .getBytes());
    postbody = postbody.concat(Utilities.newBlob(JSON.stringify(training_metadata))
        .getBytes());
    postbody = postbody.concat(Utilities.newBlob(EOL + '--' + boundary + '--' + EOL)
        .getBytes());

    // リクエストオプション
    var options = {
        headers: {
            Authorization: " Basic " + Utilities.base64Encode(p_username + ":" + p_password),
        },
        method: "post",
        contentType: 'multipart/form-data; boundary=' + boundary,
        payload: postbody,
        muteHttpExceptions: true,
    };

    var fromTime = new Date();

    var response;
    try {
        response = UrlFetchApp.fetch(uri, options);
    } catch (e) {
        Logger.log(e);
        return {
            status: 999,
            description: e,
        };
    }
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    var result;
    result = {
        status: responseCode,
        body: JSON.parse(responseBody),
        from: fromTime.getTime(),
        to: toTime.getTime(),
        delta: delta,
    };
    return result;
}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * クラス分類
 * <p>POST /v1/classifiers/{classifier_id}/classify</p>
 * <p>概要: 文章を対象の分類器で分類する</p>
 * @param       {String} p_username NLCサービス資格情報のusername
 * @param       {String} p_password NLCサービス資格情報のpassword
 * @param       {String} p_classid  分類器のクラスID
 * @param       {String} p_phrase   分類する文章
 * @return      {APIResult}            APIの実行結果
 */
function NLCAPI_post_classify(p_username, p_password, p_classid, p_phrase) {

    // URIビルド
    var resource = "classifiers";
    var verb = "classify";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource + "/" + p_classid + "/" + verb;

    // ポストデータ
    var postbody = JSON.stringify({
        text: p_phrase,
    });

    // リクエストオプション
    var options = {
        headers: {
            Authorization: " Basic " + Utilities.base64Encode(p_username + ":" + p_password),
        },
        method: "post",
        contentType: "application/json",
        payload: postbody,
        muteHttpExceptions: true,
    };

    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    var result;
    result = {
        status: responseCode,
        body: JSON.parse(responseBody),
        from: fromTime.getTime(),
        to: toTime.getTime(),
        delta: delta,
    };
    // 戻り値を返す
    return result;
}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * クラス分類
 * <p>GET /v1/classifiers/{classifier_id}/classify</p>
 * <p>概要: 文章を対象のClassifierで分類する</p>
 * @param       {String} p_username NLCサービス資格情報のusername
 * @param       {String} p_password NLCサービス資格情報のpassword
 * @param       {String} p_classid  Classifier_id
 * @param       {String} p_phrase   分類する文章
 * @return      {APIResult}            APIの実行結果
 */
function NLCAPI_get_classify(p_username, p_password, p_classid, p_phrase) {

    // URIビルド
    var resource = "classifiers";
    var verb = "classify";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource + "/" + p_classid + "/" + verb;
    uri = uri + "?text=" + p_phrase;

    // リクエストオプション
    var options = {
        headers: {
            Authorization: " Basic " + Utilities.base64Encode(p_username + ":" + p_password),
        },
        method: "get",
        contentType: "application/json",
        muteHttpExceptions: true,
    };

    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    var result;
    result = {
        status: responseCode,
        body: JSON.parse(responseBody),
        from: fromTime.getTime(),
        to: toTime.getTime(),
        delta: delta,
    };
    // 戻り値を返す
    return result;
}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * Classifier削除
 * <p>DELETE /v1/classifiers/{classifier_id}</p>
 * <p>概要: 対象のClassifierを削除する</p>
 * @param       {String} p_username NLCサービス資格情報のusername
 * @param       {String} p_password NLCサービス資格情報のpassword
 * @param       {String} p_classid  削除対象のClassifier_id
 * @return      {APIResult}            APIの実行結果
 */
function NLCAPI_delete_classifier(p_username, p_password, p_classid) {

    // URIビルド
    var resource = "classifiers";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource + "/" + p_classid;
    // リクエストオプション
    var options = {
        headers: {
            Authorization: " Basic " + Utilities.base64Encode(p_username + ":" + p_password),
        },
        method: "delete",
        contentType: "application/json",
        muteHttpExceptions: true,
    };

    // HTTPリクエスト
    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    var result;
    result = {
        status: responseCode,
        body: JSON.parse(responseBody),
        from: fromTime.getTime(),
        to: toTime.getTime(),
        delta: delta,
    };
    // 戻り値を返す
    return result;
}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * Classifier情報取得
 * <p>GET /v1/classifiers/{classifier_id}</p>
 * <p>概要: 対象のClassifierについてステータスなどの情報を取得する</p>
 * @param       {String} p_username NLCサービス資格情報のusername
 * @param       {String} p_password NLCサービス資格情報のpassword
 * @param       {String} p_classid  取得対象のClassifier_id
 * @return      {APIResult}            APIの実行結果
 */
function NLCAPI_get_classifier(p_username, p_password, p_classid) {

    // URIビルド
    var resource = "classifiers";
    var uri = URI_DOMAIN + URI_BASE + "/" + URI_APIVERSION + "/" + resource + "/" + p_classid;
    // リクエストオプション
    var options = {
        headers: {
            Authorization: " Basic " + Utilities.base64Encode(p_username + ":" + p_password),
        },
        method: "get",
        contentType: "application/json",
        muteHttpExceptions: true,
    };

    // HTTPリクエスト
    var fromTime = new Date();
    var response = UrlFetchApp.fetch(uri, options);
    var toTime = new Date();
    var delta = (toTime - fromTime);

    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();
    var result;
    result = {
        status: responseCode,
        body: JSON.parse(responseBody),
        from: fromTime.getTime(),
        to: toTime.getTime(),
        delta: delta,
    };
    return result;
}
// ----------------------------------------------------


// ----------------------------------------------------
/**
 * multipartデータ送信のバウンダリ生成
 * @return {String} バウンダリ
 */
function NLCAPI_createBoundary() {

    var multipartChars = "-_1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
    var length = 30 + Math.floor(Math.random() * 10);
    var boundary = "---------------------------";
    for (var i = 0; i < length; i += 1) {
        boundary += multipartChars.charAt(Math.floor(Math.random() * multipartChars.length));
    }
    return boundary;
}
// ----------------------------------------------------
