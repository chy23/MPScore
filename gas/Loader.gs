/**
 * MPScore - GitHub 自動熱更新引導核心 (比照 QRcode 專案架構)
 * 只要在 Apps Script 貼上這段程式碼，未來 GitHub (chy23/MPScore) 有任何演算法或功能更新，
 * 您的 Google 試算表都會全自動即時抓取最新版，免去重複手動貼代碼的麻煩！
 * 
 * 部署指引：
 * 1. 貼上後儲存。
 * 2. 第一次使用請先點上方「執行」> 選擇「setup」並執行，完成 Google 帳號授權。
 * 3. 點選右上角「部署」>「新部署」，選「網頁應用程式」，【誰可以存取】務必選【所有人 (Anyone)】！
 */
var GITHUB_CORE_URL = "https://raw.githubusercontent.com/chy23/MPScore/main/gas/Code.gs";
var _cachedEngine = null;

function setup() {
  getEngine();
  SpreadsheetApp.getActiveSpreadsheet();
}

function getEngine() {
  if (!_cachedEngine) {
    var t = new Date().getTime();
    var response = UrlFetchApp.fetch(GITHUB_CORE_URL + "?t=" + t);
    var code = response.getContentText();
    (1, eval)(code);
    if (typeof MPScoreEngine !== 'undefined') {
      _cachedEngine = MPScoreEngine;
    } else {
      throw new Error("無法載入 MPScore 核心模組，請檢查 GitHub 網路連線。");
    }
  }
  return _cachedEngine;
}

// 靜態建立選單（不使用 UrlFetchApp，確保試算表每次開啟時都能 100% 穩定出現選單！）
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ MPScore 工具')
    .addItem('🛠️ 一鍵自動建立標準分頁範本 (全新空白表專用)', 'setupStandardSheets')
    .addSeparator()
    .addItem('➕ 從【學生名冊】建立新考試分頁 (自動帶入全班姓名)', 'createNewExamSheetFromRoster')
    .addItem('🔄 從最新考試分頁同步名冊至【學生名冊】', 'syncStudentsFromExamSheet')
    .addItem('🔑 一鍵為全班產生防偽查詢代碼 (Token)', 'generateTokensForStudents')
    .addSeparator()
    .addItem('📐 一鍵自動美化所有分頁與最佳欄寬 (立即修復文字擠壓)', 'autoFormatAndResizeAllSheets')
    .addItem('📊 預覽全班成績統計分析與各科平均', 'previewClassStatistics')
    .addSeparator()
    .addItem('🔄 測試與 GitHub 核心連線狀態', 'testGitHubConnection')
    .addToUi();
}

// 選單動作分發 (由 GitHub 核心即時執行最新邏輯)
function setupStandardSheets() { getEngine().setupStandardSheets(); }
function createNewExamSheetFromRoster() { getEngine().createNewExamSheetFromRoster(); }
function syncStudentsFromExamSheet() { getEngine().syncStudentsFromExamSheet(); }
function generateTokensForStudents() { getEngine().generateTokensForStudents(); }
function autoFormatAndResizeAllSheets() { getEngine().autoFormatAndResizeAllSheets(); }
function previewClassStatistics() { getEngine().previewClassStatistics(); }

// Web API 入口點 (附完整錯誤捕捉保護，絕不回傳 HTML 當機頁面)
function doGet(e) {
  try {
    return getEngine().handleDoGet(e);
  } catch (err) {
    var out = {
      success: false,
      message: "【Loader 熱更新執行異常】: " + err.toString() + "\n\n💡 快速解決：\n1. 請在 Apps Script 上方將函式選為「setup」，並點擊「執行」完成帳號授權；\n2. 或者直接複製網站【步驟 2】的【模式 B：完整單機版程式碼】(零外部依賴，100% 穩定秒連！)"
    };
    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  return doGet(e);
}

function testGitHubConnection() {
  try {
    var engine = getEngine();
    SpreadsheetApp.getUi().alert("🎉 連線成功！已成功自 GitHub (chy23/MPScore) 載入最新版 MPScore 核心模組。");
  } catch (err) {
    SpreadsheetApp.getUi().alert("⚠️ 連線失敗: " + err.toString());
  }
}
