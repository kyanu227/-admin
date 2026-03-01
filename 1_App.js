// ■■■ Admin.gs : ダッシュボード・管理機能 ■■■

function doGet(e) {
  var userEmail = getSafeUserEmail();
  var userInfo = getUserInfo(userEmail, ""); // No passcode in admin URL usually

  // Optional: Check if the user is truly an admin before rendering
  if (!checkAdminRole(userInfo)) {
    return HtmlService.createHtmlOutput("<h1>アクセス権限がありません。管理者アカウントでログインしてください。</h1>");
  }

  return createAdminPage(userInfo, userEmail);
}

function createAdminPage(userInfo, userEmail) {
  var template = HtmlService.createTemplateFromFile('index'); // We renamed admin.html to index.html

  template.staffName = userInfo.name;
  template.userRole = userInfo.role;
  template.scriptUrl = ScriptApp.getService().getUrl();
  template.loginMode = getLoginMode();
  template.userEmail = userEmail || "";

  try { template.menuNames = MENU_NAMES; } catch (e) { template.menuNames = {}; }

  return template.evaluate()
    .setTitle("タンク管理システム - 管理画面")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function checkAdminRole(userInfo) {
  return (userInfo.role.indexOf('管理者') !== -1 ||
    userInfo.role.indexOf('準管理者') !== -1 ||
    userInfo.role.toLowerCase().indexOf('admin') !== -1);
}

// ダッシュボード集計用のシート名・列インデックス設定
const ADMIN_CONFIG = {
  SHEET_PRICE: 'M_設定_単価',      // 既存のFeature系と同じ単価マスタ
  SHEET_TANK: 'タンクステータス',    // ボンベの現在の状況

  // 各シートの列番号 (0始まり: A列=0, B列=1...)
  COL_LOG_DATE: 1,    // ログの日付列 (B列)
  COL_LOG_ACTION: 4,  // ログのアクション列 (E列)
  COL_LOG_STAFF: 7,   // ログの担当者列 (H列)

  COL_PRICE_NAME: 0,  // 単価マスタの品名 (A列)
  COL_PRICE_VAL: 1,   // 単価マスタの価格 (B列)

  COL_TANK_ID: 0,     // タンクID列 (A列)
  COL_TANK_STATUS: 1, // ステータス列 (B列)
  COL_TANK_LIMIT: 4   // 耐圧検査期限の列 (E列)
};

/**
 * ダッシュボード用データ一括取得
 */
function getAdminDashboardData() {
  var today = new Date();
  var yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);

  // 1. 各部から詳細データを取得
  var salesStats = getDetailedSalesStats();
  var staffStats = getDetailedStaffStats();
  var expiryInfo = checkTankExpiry_();
  var orderCount = 0; // いったん固定

  // 2. 取得したデータをダッシュボードトップ用の形式にマッピング
  var salesRatio = salesStats.momRatio || 0;

  // 今日の売上を算出する処理 (ActionBreakdown等から取るか、別途当日だけ再取得。
  // 簡単のために、getDetailedSalesStats では当月全体のトータルを currentMonthTotal として返しているので、
  // 今日の売上は、元々の calcDailySales_ ではなく、昨日/今日を計算するか、または月間トータルをそのまま「今月の売上」として表示するように変更するのもありです。
  // ユーザーの要件に合わせて、ダッシュボードトップには「本日の売上」を残すため、当日分だけ簡易計算します)
  var priceMap = getPriceMap_();
  var salesToday = calcDailySales_(today, priceMap);
  var salesYesterday = calcDailySales_(yesterday, priceMap);
  if (salesYesterday > 0) {
    salesRatio = Math.round(((salesToday - salesYesterday) / salesYesterday) * 100);
  } else if (salesToday > 0) {
    salesRatio = 100;
  }

  var trendData = getSalesTrend_(7, priceMap);

  return {
    sales: salesToday,
    salesRatio: salesRatio,
    alertCount: expiryInfo.expiredCount,
    warningCount: expiryInfo.warningCount,
    orderCount: orderCount,
    activeStaff: staffStats.activeStaffCount,

    // チャート用データ
    chartLabels: trendData.labels,
    chartData: trendData.data,

    // お知らせリスト (期限切れ情報を追加)
    notifications: expiryInfo.messages
  };
}

// -------------------------------------------------------
// 内部計算用関数 (このファイル内でのみ使用)
// -------------------------------------------------------

/**
 * 単価マスタをアクション名→価格の連想配列で取得 { '充填': 5000, ... }
 * ※本スプレッドシートの 'M_設定_単価' シートを参照
 */
function getPriceMap_() {
  var ss = getMainSpreadsheet();
  var sheet = ss.getSheetByName(ADMIN_CONFIG.SHEET_PRICE);
  var map = {};
  if (!sheet) return map;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][ADMIN_CONFIG.COL_PRICE_NAME]).trim();
    var price = Number(data[i][ADMIN_CONFIG.COL_PRICE_VAL]) || 0;
    if (name) map[name] = price;
  }
  return map;
}

/**
 * フロントエンドからの詳細売上データ取得用エンドポイント
 */
function getAdminSalesData() {
  try {
    return getDetailedSalesStats(); // Admin_Feature_Sales.js の関数を呼び出す
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * フロントエンドからの詳細スタッフデータ取得用エンドポイント
 */
function getAdminStaffData() {
  try {
    return getDetailedStaffStats(); // Admin_Feature_Staff.js の関数を呼び出す
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * 耐圧検査期限のチェック (期限切れ件数・警告件数・メッセージ一覧を返す)
 */
function checkTankExpiry_() {
  var ss = getMainSpreadsheet();
  var sheet = ss.getSheetByName(ADMIN_CONFIG.SHEET_TANK);
  var expiredCount = 0;
  var warningCount = 0;
  var messages = [];

  if (!sheet) return { expiredCount: 0, warningCount: 0, messages: [] };

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var warnDate = new Date();
  warnDate.setDate(today.getDate() + 30); // 30日以内を警告ライン

  for (var i = 1; i < data.length; i++) {
    var tankId = data[i][ADMIN_CONFIG.COL_TANK_ID];
    var limitDate = data[i][ADMIN_CONFIG.COL_TANK_LIMIT];

    if (isValidDate(limitDate)) {
      var limit = new Date(limitDate);
      if (limit < today) {
        expiredCount++;
        messages.push({ type: 'Warn', msg: '期限切れ: ' + tankId + ' (' + Utilities.formatDate(limit, "JST", "yyyy/MM") + ')' });
      } else if (limit < warnDate) {
        warningCount++;
      }
    }
  }
  return { expiredCount: expiredCount, warningCount: warningCount, messages: messages };
}

// isValidDate は 2_Utils.gs に定義済み
// countActiveStaff_ は getDetailedStaffStats (Admin_Feature_Staff.gs) の activeStaffCount で代替済みのため削除