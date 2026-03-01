// ■■■ Notify_Triggers.js : 通知トリガー関数 ■■■

// -----------------------------------------------------------
// ① 日報 (毎日 設定時刻)
// -----------------------------------------------------------
function runDailyReport() {
    var config = getNotifyConfig();
    var data = buildDailyReportData();
    var message = formatDailyReportMessage_(data);

    sendToRecipients_(config.recipients, 'DAILY', {
        lineMessage: message,
        emailSubject: '【タンク管理】日報 ' + data.date,
        emailBody: message.replace(/\n/g, '<br>')
    });

    console.log('日報送信完了: ' + data.date);
}

// -----------------------------------------------------------
// ② リマインダー (毎日 設定時刻)
// -----------------------------------------------------------
function runReminder() {
    var ss = getMainSpreadsheet();
    var today = new Date();
    var tz = ss.getSpreadsheetTimeZone() || 'Asia/Tokyo';
    var todayStr = Utilities.formatDate(today, tz, 'yyyy/MM/dd');

    // 本日の操作件数を確認
    var year = today.getFullYear();
    var logSheetName = (typeof SHEET_NAMES !== 'undefined' ? SHEET_NAMES.LOG : '履歴ログ') + year;
    var logSheet = ss.getSheetByName(logSheetName) || ss.getSheetByName('履歴ログ');

    var todayCount = 0;
    if (logSheet) {
        var logData = logSheet.getDataRange().getValues();
        for (var i = 1; i < logData.length; i++) {
            if (!isValidDate(logData[i][1])) continue;
            if (Utilities.formatDate(new Date(logData[i][1]), tz, 'yyyy/MM/dd') === todayStr) {
                todayCount++;
            }
        }
    }

    if (todayCount > 0) {
        console.log('本日は ' + todayCount + ' 件の操作あり。リマインダーをスキップします。');
        return; // 操作があればリマインダー不要
    }

    var config = getNotifyConfig();
    var dateStr = Utilities.formatDate(today, tz, 'M月d日 (E)');
    var message = [
        '⚠️ 【リマインダー】',
        dateStr + ' の操作が現時点でまだ記録されていません。',
        '',
        '現場操作（貸出・返却・充填等）があれば、必ずアプリから入力をお願いします。',
        '',
        '問題がなければこのメッセージは無視してください。'
    ].join('\n');

    sendToRecipients_(config.recipients, 'REMINDER', {
        lineMessage: message,
        emailSubject: '【タンク管理】本日の作業入力リマインダー',
        emailBody: message.replace(/\n/g, '<br>')
    });

    console.log('リマインダー送信完了');
}

// -----------------------------------------------------------
// ③ 月次耐圧検査アラート (毎月1日 設定時刻)
// -----------------------------------------------------------
function runMonthlyInspection() {
    var config = getNotifyConfig();
    var data = buildInspectionAlertData();

    if (data.expired.length === 0 && data.warning.length === 0) {
        console.log('耐圧検査アラート: 対象なし。送信スキップ。');
        return;
    }

    var lines = [
        '🔧 【月次耐圧検査アラート】',
        Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月') + ' 分',
        ''
    ];

    if (data.expired.length > 0) {
        lines.push('❌ 期限切れ (' + data.expired.length + '本):');
        data.expired.slice(0, 10).forEach(function (t) {
            lines.push('  ' + t.id + ' (期限: ' + t.limit + ')');
        });
        if (data.expired.length > 10) lines.push('  ... 他' + (data.expired.length - 10) + '本');
        lines.push('');
    }

    if (data.warning.length > 0) {
        lines.push('⚠️ ' + data.alertMonths + 'ヶ月以内に期限 (' + data.warning.length + '本):');
        data.warning.slice(0, 10).forEach(function (t) {
            lines.push('  ' + t.id + ' (期限: ' + t.limit + ')');
        });
        if (data.warning.length > 10) lines.push('  ... 他' + (data.warning.length - 10) + '本');
    }

    lines.push('');
    lines.push('耐圧検査の手配をお願いします。');

    var message = lines.join('\n');

    sendToRecipients_(config.recipients, 'INSPECTION', {
        lineMessage: message,
        emailSubject: '【タンク管理】月次耐圧検査アラート',
        emailBody: message.replace(/\n/g, '<br>')
    });

    console.log('耐圧検査アラート送信完了: 期限切れ' + data.expired.length + '本、警告' + data.warning.length + '本');
}

// -----------------------------------------------------------
// ④ 週次未回収アラート (毎週月曜 設定時刻)
// -----------------------------------------------------------
function runWeeklyUnreturnedAlert() {
    var config = getNotifyConfig();
    var data = buildUnreturnedData();

    var totalLong = (data.over60 ? data.over60.length : 0) +
        (data.buckets && data.buckets[30] ? data.buckets[30].length : 0);

    if (totalLong === 0) {
        console.log('週次未回収アラート: 長期未回収なし。送信スキップ。');
        return;
    }

    var lines = [
        '📦 【週次 未回収タンクアラート】',
        Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd') + ' 時点',
        ''
    ];

    if (data.over60 && data.over60.length > 0) {
        lines.push('🚨 60日超 (' + data.over60.length + '本):');
        data.over60.slice(0, 8).forEach(function (t) {
            lines.push('  ' + t.id + ' → ' + t.dest + ' (' + t.days + '日)');
        });
        if (data.over60.length > 8) lines.push('  ... 他' + (data.over60.length - 8) + '本');
        lines.push('');
    }

    if (data.buckets && data.buckets[30] && data.buckets[30].length > 0) {
        lines.push('⚠️ 30〜60日 (' + data.buckets[30].length + '本):');
        data.buckets[30].slice(0, 8).forEach(function (t) {
            lines.push('  ' + t.id + ' → ' + t.dest + ' (' + t.days + '日)');
        });
        if (data.buckets[30].length > 8) lines.push('  ... 他' + (data.buckets[30].length - 8) + '本');
        lines.push('');
    }

    lines.push('回収・確認をお願いします。');

    var message = lines.join('\n');

    sendToRecipients_(config.recipients, 'WEEKLY', {
        lineMessage: message,
        emailSubject: '【タンク管理】週次未回収タンクアラート',
        emailBody: message.replace(/\n/g, '<br>')
    });

    console.log('週次未回収アラート送信完了');
}

// -----------------------------------------------------------
// 送信共通関数
// -----------------------------------------------------------

/**
 * 指定の通知タイプに該当する宛先に送信する
 */
function sendToRecipients_(recipients, notifyType, payload) {
    if (!recipients || recipients.length === 0) return;

    recipients.forEach(function (r) {
        var shouldSend = r.notifyTypes.indexOf('ALL') !== -1 || r.notifyTypes.indexOf(notifyType) !== -1;
        if (!shouldSend) return;

        if (r.type === 'LINE') {
            if (r.token && r.dest) {
                sendLineMessage_(r.dest, r.token, payload.lineMessage);
            }
        } else if (r.type === 'EMAIL') {
            if (r.dest) {
                sendEmailNotification_([r.dest], payload.emailSubject, payload.emailBody);
            }
        }
    });
}

/**
 * LINE Messaging APIでメッセージを送信する
 * @param {string} groupId - LINE グループID
 * @param {string} token - チャネルアクセストークン
 * @param {string} message - 送信テキスト
 */
function sendLineMessage_(groupId, token, message) {
    try {
        var url = 'https://api.line.me/v2/bot/message/push';
        var payload = JSON.stringify({
            to: groupId,
            messages: [{ type: 'text', text: message }]
        });
        var options = {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + token },
            payload: payload,
            muteHttpExceptions: true
        };
        var res = UrlFetchApp.fetch(url, options);
        var code = res.getResponseCode();
        if (code !== 200) {
            console.error('LINE送信エラー: HTTP ' + code + ' / ' + res.getContentText());
        }
    } catch (e) {
        console.error('LINE送信例外: ' + e.message);
    }
}

/**
 * メール通知を送信する
 */
function sendEmailNotification_(emails, subject, htmlBody) {
    if (!emails || emails.length === 0) return;
    emails.forEach(function (email) {
        if (!email) return;
        try {
            GmailApp.sendEmail(email, subject, '', { htmlBody: htmlBody });
        } catch (e) {
            console.error('メール送信エラー (' + email + '): ' + e.message);
        }
    });
}

// -----------------------------------------------------------
// メッセージフォーマット
// -----------------------------------------------------------

function formatDailyReportMessage_(data) {
    var lines = [
        '📋 【タンク管理 日報】' + data.date,
        ''
    ];

    // 貸出
    if (data.lendings.length > 0) {
        lines.push('【本日の貸出】');
        data.lendings.forEach(function (l) {
            lines.push('  ・' + l.dest + ' × ' + l.count + '本 / 担当: ' + l.staff);
        });
    } else {
        lines.push('【本日の貸出】なし');
    }
    lines.push('');

    // 返却
    if (data.returns.length > 0) {
        lines.push('【本日の返却】');
        data.returns.forEach(function (r) {
            lines.push('  ・' + r.dest + ' × ' + r.count + '本 / 担当: ' + r.staff);
        });
    } else {
        lines.push('【本日の返却】なし');
    }
    lines.push('');

    // 未回収
    lines.push('【未回収タンク】現在 ' + data.unreturned + '本');
    if (data.longUnreturned > 0) {
        lines.push('  うち30日超: ' + data.longUnreturned + '本 ⚠️');
    }
    lines.push('');

    // 耐圧アラート
    if (data.expiredCount > 0 || data.warningCount > 0) {
        lines.push('【耐圧検査アラート】');
        if (data.expiredCount > 0) lines.push('  期限切れ: ' + data.expiredCount + '本 ❌');
        if (data.warningCount > 0) lines.push('  期限間近: ' + data.warningCount + '本 ⚠️');
        lines.push('');
    }

    lines.push('━━━━━━━━━━');
    lines.push('タンク管理システム 自動送信');

    return lines.join('\n');
}

// -----------------------------------------------------------
// トリガー管理
// -----------------------------------------------------------

/**
 * 全通知トリガーをセットアップ（管理画面から呼び出し）
 * 既存のNotify系トリガーを削除してから再登録する
 */
function setupNotifyTriggers_() {
    var config = getNotifyConfig();
    var s = config.settings || NOTIFY_DEFAULTS;

    // 既存のNotify系トリガーを全削除
    deleteNotifyTriggers_();

    // 日報トリガー (毎日)
    ScriptApp.newTrigger('runDailyReport')
        .timeBased()
        .everyDays(1)
        .atHour(s.DAILY_REPORT_HOUR)
        .create();

    // リマインダートリガー (毎日)
    ScriptApp.newTrigger('runReminder')
        .timeBased()
        .everyDays(1)
        .atHour(s.REMINDER_HOUR)
        .create();

    // 月次耐圧検査トリガー (毎月1日)
    ScriptApp.newTrigger('runMonthlyInspection')
        .timeBased()
        .onMonthDay(1)
        .atHour(s.INSPECTION_HOUR)
        .create();

    // 週次未回収トリガー (毎週月曜)
    ScriptApp.newTrigger('runWeeklyUnreturnedAlert')
        .timeBased()
        .everyWeeks(1)
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(s.UNRETURNED_HOUR)
        .create();

    console.log('通知トリガーを登録しました。');
}

/**
 * Notify系トリガーを全削除
 */
function deleteNotifyTriggers_() {
    var notifyFunctions = ['runDailyReport', 'runReminder', 'runMonthlyInspection', 'runWeeklyUnreturnedAlert'];
    ScriptApp.getProjectTriggers().forEach(function (trigger) {
        if (notifyFunctions.indexOf(trigger.getHandlerFunction()) !== -1) {
            ScriptApp.deleteTrigger(trigger);
        }
    });
}

/**
 * フロントエンドから呼ばれる: トリガーを登録
 */
function setupTriggers() {
    try {
        setupNotifyTriggers_();
        var triggers = ScriptApp.getProjectTriggers().filter(function (t) {
            return ['runDailyReport', 'runReminder', 'runMonthlyInspection', 'runWeeklyUnreturnedAlert'].indexOf(t.getHandlerFunction()) !== -1;
        });
        return { success: true, message: '✅ ' + triggers.length + '件のトリガーを登録しました。', count: triggers.length };
    } catch (e) {
        return { success: false, message: 'トリガー登録エラー: ' + e.message };
    }
}

/**
 * フロントエンドから呼ばれる: トリガーを削除
 */
function deleteTriggers() {
    try {
        deleteNotifyTriggers_();
        return { success: true, message: '通知トリガーを全て削除しました。' };
    } catch (e) {
        return { success: false, message: 'トリガー削除エラー: ' + e.message };
    }
}

/**
 * フロントエンドから呼ばれる: 現在のトリガー状況を取得
 */
function getTriggerStatus() {
    try {
        var notifyFunctions = { runDailyReport: false, runReminder: false, runMonthlyInspection: false, runWeeklyUnreturnedAlert: false };
        ScriptApp.getProjectTriggers().forEach(function (t) {
            if (notifyFunctions.hasOwnProperty(t.getHandlerFunction())) {
                notifyFunctions[t.getHandlerFunction()] = true;
            }
        });
        return { success: true, triggers: notifyFunctions };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

/**
 * フロントエンドから呼ばれる: 今すぐ日報を手動送信（テスト用）
 */
function sendDailyReportNow() {
    try {
        runDailyReport();
        return { success: true, message: '日報を送信しました。' };
    } catch (e) {
        return { success: false, message: '送信エラー: ' + e.message };
    }
}
