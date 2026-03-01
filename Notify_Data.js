// ■■■ Notify_Data.js : 通知設定のシート読み書き ■■■

var NOTIFY_SHEET_NAME = '通知設定';
var NOTIFY_SETTINGS_ROWS = 6;  // 設定行数 (行1〜6)
var NOTIFY_HEADER_ROW = 8;     // 宛先リストのヘッダー行

// 設定キー → 行番号マッピング (1始まり)
var NOTIFY_KEY_MAP = {
    ALERT_MONTHS: 1,
    VALIDITY_YEARS: 2,
    DAILY_REPORT_HOUR: 3,
    REMINDER_HOUR: 4,
    INSPECTION_HOUR: 5,
    UNRETURNED_HOUR: 6
};

// デフォルト値
var NOTIFY_DEFAULTS = {
    ALERT_MONTHS: 6,
    VALIDITY_YEARS: 3,
    DAILY_REPORT_HOUR: 18,
    REMINDER_HOUR: 12,
    INSPECTION_HOUR: 9,
    UNRETURNED_HOUR: 9
};

/**
 * 通知設定シートを取得（なければ作成）
 */
function getNotifySheet_() {
    var ss = getMainSpreadsheet();
    var sheet = ss.getSheetByName(NOTIFY_SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(NOTIFY_SHEET_NAME);

        // セクション1: システム設定 (行1〜6)
        var settings = [
            ['ALERT_MONTHS', NOTIFY_DEFAULTS.ALERT_MONTHS, 'アラート表示開始 (ヶ月前)'],
            ['VALIDITY_YEARS', NOTIFY_DEFAULTS.VALIDITY_YEARS, '耐圧検査有効期間 (年)'],
            ['DAILY_REPORT_HOUR', NOTIFY_DEFAULTS.DAILY_REPORT_HOUR, '日報 送信時刻 (0〜23)'],
            ['REMINDER_HOUR', NOTIFY_DEFAULTS.REMINDER_HOUR, 'リマインダー 送信時刻 (0〜23)'],
            ['INSPECTION_HOUR', NOTIFY_DEFAULTS.INSPECTION_HOUR, '月次耐圧検査通知 送信時刻 (0〜23)'],
            ['UNRETURNED_HOUR', NOTIFY_DEFAULTS.UNRETURNED_HOUR, '週次未回収通知 送信時刻 (0〜23)']
        ];
        sheet.getRange(1, 1, settings.length, 3).setValues(settings);

        // 区切り行
        sheet.getRange(7, 1).setValue('---');

        // セクション2: 宛先リストのヘッダー (行8)
        sheet.getRange(NOTIFY_HEADER_ROW, 1, 1, 4).setValues([
            ['種別(LINE/EMAIL)', 'グループID / メールアドレス', 'LINEトークン', '通知タイプ(カンマ区切り)']
        ]);
        sheet.getRange(NOTIFY_HEADER_ROW, 1, 1, 4).setFontWeight('bold');

        // 列幅の調整
        sheet.setColumnWidth(1, 130);
        sheet.setColumnWidth(2, 300);
        sheet.setColumnWidth(3, 400);
        sheet.setColumnWidth(4, 250);
    }

    return sheet;
}

/**
 * 通知設定を全て読み込む
 * @returns {{settings: Object, recipients: Array}}
 */
function getNotifyConfig() {
    try {
        var sheet = getNotifySheet_();
        var data = sheet.getDataRange().getValues();

        // セクション1: 設定値の読み込み
        var settings = {};
        Object.keys(NOTIFY_KEY_MAP).forEach(function (key) {
            var rowIdx = NOTIFY_KEY_MAP[key] - 1; // 0始まり
            var val = (data.length > rowIdx && data[rowIdx][1] !== '') ? Number(data[rowIdx][1]) : NOTIFY_DEFAULTS[key];
            settings[key] = isNaN(val) ? NOTIFY_DEFAULTS[key] : val;
        });

        // セクション2: 宛先リストの読み込み (行9以降: headerRow + 1)
        var recipients = [];
        for (var i = NOTIFY_HEADER_ROW; i < data.length; i++) { // NOTIFY_HEADER_ROW = 8, 0始まりで行8はindex8
            var row = data[i];
            var type = String(row[0] || '').trim().toUpperCase();
            var dest = String(row[1] || '').trim();
            if (!type || !dest) continue;

            var token = String(row[2] || '').trim();
            var rawTypes = String(row[3] || '').trim();
            var notifyTypes = rawTypes ? rawTypes.split(',').map(function (t) { return t.trim().toUpperCase(); }) : ['ALL'];

            recipients.push({
                type: type,        // 'LINE' or 'EMAIL'
                dest: dest,        // groupId or email address
                token: token,      // LINE only
                notifyTypes: notifyTypes  // ['DAILY', 'REMINDER', ...]
            });
        }

        return { success: true, settings: settings, recipients: recipients };
    } catch (e) {
        return { success: false, message: e.message, settings: NOTIFY_DEFAULTS, recipients: [] };
    }
}

/**
 * 通知設定を保存する
 * @param {Object} data - { settings: {}, recipients: [] }
 */
function saveNotifyConfig(data) {
    try {
        var sheet = getNotifySheet_();

        // セクション1: 設定値の書き込み
        if (data.settings) {
            Object.keys(NOTIFY_KEY_MAP).forEach(function (key) {
                if (data.settings[key] !== undefined) {
                    var rowNum = NOTIFY_KEY_MAP[key];
                    sheet.getRange(rowNum, 2).setValue(Number(data.settings[key]) || NOTIFY_DEFAULTS[key]);
                }
            });
        }

        // セクション2: 宛先リストの書き込み
        if (data.recipients) {
            // 既存データを消去 (行9以降)
            var lastRow = sheet.getLastRow();
            if (lastRow >= NOTIFY_HEADER_ROW + 1) {
                sheet.getRange(NOTIFY_HEADER_ROW + 1, 1, lastRow - NOTIFY_HEADER_ROW, 4).clearContent();
            }

            if (data.recipients.length > 0) {
                var writeRows = data.recipients
                    .filter(function (r) { return r.dest && r.type; })
                    .map(function (r) {
                        return [
                            r.type.toUpperCase(),
                            r.dest,
                            r.token || '',
                            Array.isArray(r.notifyTypes) ? r.notifyTypes.join(',') : (r.notifyTypes || 'ALL')
                        ];
                    });

                if (writeRows.length > 0) {
                    sheet.getRange(NOTIFY_HEADER_ROW + 1, 1, writeRows.length, 4).setValues(writeRows);
                }
            }
        }

        // トリガーの再登録（時刻が変わっている可能性があるため）
        try {
            setupNotifyTriggers_();
        } catch (te) {
            console.warn('トリガー再登録エラー (無視): ' + te.message);
        }

        return { success: true, message: '通知設定を保存しました。トリガーも更新されました。' };
    } catch (e) {
        return { success: false, message: '保存エラー: ' + e.message };
    }
}

/**
 * 日報用データを収集する
 */
function buildDailyReportData() {
    var ss = getMainSpreadsheet();
    var today = new Date();
    var tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
    var todayStr = Utilities.formatDate(today, tz, 'yyyy/MM/dd');

    var result = {
        date: Utilities.formatDate(today, tz, 'yyyy/MM/dd (E)'),
        lendings: [],    // 本日の貸出 [{dest, count, staff}]
        returns: [],     // 本日の返却 [{dest, count, staff}]
        unreturned: 0,   // 未回収タンク総数
        longUnreturned: 0, // 30日超未回収
        expiredCount: 0,
        warningCount: 0
    };

    // 1. ログシートから当日の貸出・返却を集計
    var year = today.getFullYear();
    var logSheetName = (typeof SHEET_NAMES !== 'undefined' ? SHEET_NAMES.LOG : '履歴ログ') + year;
    var logSheet = ss.getSheetByName(logSheetName) || ss.getSheetByName('履歴ログ');

    if (logSheet) {
        var logData = logSheet.getDataRange().getValues();
        // [0]UUID, [1]日時, [2]時刻, [3]タンクID, [4]操作, [5]場所, [6]備考, [7]担当者

        var lendMap = {};  // dest → {count, staffSet}
        var retMap = {};  // dest → {count, staffSet}

        for (var i = 1; i < logData.length; i++) {
            var row = logData[i];
            if (!isValidDate(row[1])) continue;
            var dStr = Utilities.formatDate(new Date(row[1]), tz, 'yyyy/MM/dd');
            if (dStr !== todayStr) continue;

            var action = String(row[4] || '');
            var dest = String(row[5] || '自社');
            var staff = String(row[7] || '不明');

            if (action === '貸出') {
                if (!lendMap[dest]) lendMap[dest] = { count: 0, staffSet: {} };
                lendMap[dest].count++;
                lendMap[dest].staffSet[staff] = true;
            } else if (action === '返却' || action === '返却(未充填)' || action === '未使用返却') {
                if (!retMap[dest]) retMap[dest] = { count: 0, staffSet: {} };
                retMap[dest].count++;
                retMap[dest].staffSet[staff] = true;
            }
        }

        Object.keys(lendMap).forEach(function (d) {
            result.lendings.push({ dest: d, count: lendMap[d].count, staff: Object.keys(lendMap[d].staffSet).join('・') });
        });
        Object.keys(retMap).forEach(function (d) {
            result.returns.push({ dest: d, count: retMap[d].count, staff: Object.keys(retMap[d].staffSet).join('・') });
        });
    }

    // 2. ステータスシートから未回収・期限情報を取得
    var statusName = (typeof SHEET_NAMES !== 'undefined' ? SHEET_NAMES.STATUS : 'タンクステータス');
    var statusSheet = ss.getSheetByName(statusName);

    if (statusSheet) {
        var statusData = statusSheet.getDataRange().getValues();
        var warnDate = new Date();
        warnDate.setMonth(warnDate.getMonth() + ((typeof NOTIFY_CONFIG !== 'undefined' ? NOTIFY_CONFIG.ALERT_MONTHS : 6)));
        var thirtyDaysAgo = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);

        for (var i = 1; i < statusData.length; i++) {
            var row = statusData[i];
            var status = String(row[1] || '');
            var updateDate = row[7];
            var limitDate = row[4];

            if (status === '貸出中' || status === '未返却') {
                result.unreturned++;
                if (isValidDate(updateDate) && new Date(updateDate) < thirtyDaysAgo) {
                    result.longUnreturned++;
                }
            }

            if (isValidDate(limitDate) && status !== '耐圧検査' && status !== '廃棄') {
                var limit = new Date(limitDate);
                if (limit < today) {
                    result.expiredCount++;
                } else if (limit < warnDate) {
                    result.warningCount++;
                }
            }
        }
    }

    return result;
}

/**
 * 月次耐圧検査アラート用データを収集する
 */
function buildInspectionAlertData() {
    var ss = getMainSpreadsheet();
    var today = new Date();
    var tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
    var config = getNotifyConfig();
    var alertMonths = (config.settings && config.settings.ALERT_MONTHS) ? config.settings.ALERT_MONTHS : 6;
    var warnDate = new Date();
    warnDate.setMonth(warnDate.getMonth() + alertMonths);

    var expired = [];
    var warning = [];

    var statusName = (typeof SHEET_NAMES !== 'undefined' ? SHEET_NAMES.STATUS : 'タンクステータス');
    var sheet = ss.getSheetByName(statusName);
    if (!sheet) return { expired: [], warning: [], alertMonths: alertMonths };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        var tankId = String(data[i][0] || '');
        var limitDate = data[i][4];
        var status = String(data[i][1] || '');
        if (status === '廃棄' || status === '耐圧検査') continue;

        if (isValidDate(limitDate)) {
            var limit = new Date(limitDate);
            var limitStr = Utilities.formatDate(limit, tz, 'yyyy/MM');
            if (limit < today) {
                expired.push({ id: formatDisplayId(tankId), limit: limitStr });
            } else if (limit < warnDate) {
                warning.push({ id: formatDisplayId(tankId), limit: limitStr });
            }
        }
    }

    return { expired: expired, warning: warning, alertMonths: alertMonths };
}

/**
 * 週次未回収タンクデータを収集する
 */
function buildUnreturnedData() {
    var ss = getMainSpreadsheet();
    var today = new Date();
    var tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';

    var thresholds = [7, 14, 30, 60]; // 日数のしきい値
    var buckets = {};
    thresholds.forEach(function (d) { buckets[d] = []; });
    var over60 = [];

    var statusName = (typeof SHEET_NAMES !== 'undefined' ? SHEET_NAMES.STATUS : 'タンクステータス');
    var sheet = ss.getSheetByName(statusName);
    if (!sheet) return { buckets: buckets, over60: over60 };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
        var status = String(data[i][1] || '');
        if (status !== '貸出中' && status !== '未返却') continue;

        var tankId = formatDisplayId(String(data[i][0] || ''));
        var dest = String(data[i][2] || '不明');
        var updateDate = data[i][7];

        if (!isValidDate(updateDate)) continue;
        var diffDays = Math.floor((today.getTime() - new Date(updateDate).getTime()) / (1000 * 60 * 60 * 24));

        if (diffDays > 60) {
            over60.push({ id: tankId, dest: dest, days: diffDays });
        } else {
            for (var t = thresholds.length - 1; t >= 0; t--) {
                if (diffDays >= thresholds[t]) {
                    buckets[thresholds[t]].push({ id: tankId, dest: dest, days: diffDays });
                    break;
                }
            }
        }
    }

    return { buckets: buckets, over60: over60, total: data.length - 1 };
}
