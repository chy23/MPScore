/**
 * MPScore - GitHub 自動熱更新引導核心 (比照 QRcode 專案架構)
 * 只要貼上這段程式碼，未來 GitHub (chy23/MPScore) 有任何演算法或功能更新，
 * 您的 Google 試算表都會全自動即時抓取最新版，免去重新複製貼上的麻煩！
 * 
 * 第一次使用請先點「部署 > 新部署 (所有人)」，並完成授權。
 */
var GITHUB_URL = "https://raw.githubusercontent.com/chy23/MPScore/main/gas/Code.gs";

function setup() {
  UrlFetchApp.fetch(GITHUB_URL);
  SpreadsheetApp.getActiveSpreadsheet();
}

function loadCore() {
  var t = new Date().getTime();
  var res = UrlFetchApp.fetch(GITHUB_URL + "?t=" + t);
  (1, eval)(res.getContentText());
}

function onOpen(e) { loadCore(); onOpen(e); }
function doGet(e) { loadCore(); return doGet(e); }
function doPost(e) { loadCore(); return doPost(e); }
function setupStandardSheets() { loadCore(); setupStandardSheets(); }
function syncStudentsFromExamSheet() { loadCore(); syncStudentsFromExamSheet(); }
function generateTokensForStudents() { loadCore(); generateTokensForStudents(); }
function previewClassStatistics() { loadCore(); previewClassStatistics(); }
