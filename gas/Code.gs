/**
 * MPScore - 智慧學生成績與學習診斷系統 (Google Apps Script 後端)
 * 專為不擅長資訊的老師設計，極致簡化、高資安防護
 */

// 1. 設定試算表自訂選單（讓老師在試算表上一鍵操作）
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ MPScore 工具')
    .addItem('🔑 一鍵為全班產生防偽代碼 (Token)', 'generateTokensForStudents')
    .addItem('📊 自動計算班級平均並填入', 'autoCalculateAverages')
    .addToUi();
}

// 2. Web API 入口 (支援 GET 查詢與跨網域 CORS)
function doGet(e) {
  var output = {};
  var params = e ? e.parameter : {};
  var action = params.action || (params.token ? 'query' : 'ping');

  try {
    if (action === 'ping') {
      // 測試連線狀態
      output = checkSystemStatus();
    } else if (action === 'query') {
      // 家長查詢個別成績 (安全過濾)
      output = getStudentScoreByToken(params.token);
    } else if (action === 'getStudentsForPrint') {
      // 教師端列印 QR Code 用 (需驗證管理密碼或預設金鑰)
      output = getStudentListForTeacher(params.adminKey);
    } else {
      output = { success: false, message: '未知的操作請求' };
    }
  } catch (err) {
    output = { success: false, message: '系統處理錯誤: ' + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

// 支援 POST 請求
function doPost(e) {
  return doGet(e);
}

/**
 * 檢查系統連線狀態
 */
function checkSystemStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('學生名冊');
  var scoreSheet = ss.getSheetByName('成績明細');

  if (!studentSheet || !scoreSheet) {
    return {
      success: false,
      message: '工作表名稱不完整！必須包含【學生名冊】與【成績明細】兩個分頁。'
    };
  }

  var studentCount = Math.max(0, studentSheet.getLastRow() - 1);
  var scoreCount = Math.max(0, scoreSheet.getLastRow() - 1);

  return {
    success: true,
    message: '連線正常！系統已就緒。',
    sheetTitle: ss.getName(),
    studentCount: studentCount,
    scoreCount: scoreCount
  };
}

/**
 * 家長安全查詢：依 Token 讀取單一生歷年成績 (資安核心：絕不洩漏他人成績)
 */
function getStudentScoreByToken(token) {
  if (!token) {
    return { success: false, message: '請提供專屬查詢代碼！' };
  }
  token = token.toString().trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('學生名冊');
  var scoreSheet = ss.getSheetByName('成績明細');

  if (!studentSheet || !scoreSheet) {
    return { success: false, message: '試算表分頁設定不正確，請洽任課老師。' };
  }

  // 1. 尋找學生名冊 (A:學號, B:姓名, C:班級, D:座號, E:專屬代碼, F:啟用狀態)
  var studentData = studentSheet.getDataRange().getValues();
  var student = null;

  for (var i = 1; i < studentData.length; i++) {
    var row = studentData[i];
    var rowToken = row[4] ? row[4].toString().trim() : '';
    var isEnabled = row[5];

    if (rowToken === token) {
      if (isEnabled === false) {
        return { success: false, message: '此學生帳號目前已被設為停用。' };
      }
      student = {
        studentId: row[0].toString().trim(),
        name: row[1],
        className: row[2],
        seatNumber: row[3]
      };
      break;
    }
  }

  if (!student) {
    return { success: false, message: '驗證失敗：查無此代碼，請確認是否輸入正確。' };
  }

  // 2. 抓取該學生「歷年所有成績明細」
  var scoreData = scoreSheet.getDataRange().getValues();
  var records = [];

  // [0:學號, 1:學年度, 2:學期, 3:考試名稱, 4:科目, 5:個人得分, 6:班級平均, 7:PR值, 8:解讀評語]
  for (var j = 1; j < scoreData.length; j++) {
    var sRow = scoreData[j];
    var sId = sRow[0].toString().trim();

    if (sId === student.studentId) {
      records.push({
        year: sRow[1],
        semester: sRow[2],
        examName: sRow[3],
        examKey: sRow[1] + '-' + sRow[2] + ' ' + sRow[3],
        subject: sRow[4],
        score: Number(sRow[5]) || 0,
        average: Number(sRow[6]) || 0,
        pr: Number(sRow[7]) || 0,
        diagnosis: sRow[8] || '學習態度良好，請持續保持！'
      });
    }
  }

  return {
    success: true,
    student: student,
    records: records
  };
}

/**
 * 教師端取得全班學生清單 (供自動產生 QR Code 列印小卡)
 */
function getStudentListForTeacher(adminKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var studentSheet = ss.getSheetByName('學生名冊');
  if (!studentSheet) return { success: false, message: '找不到【學生名冊】工作表' };

  var data = studentSheet.getDataRange().getValues();
  var students = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] && row[4]) {
      students.push({
        studentId: row[0].toString(),
        name: row[1],
        className: row[2],
        seatNumber: row[3],
        token: row[4].toString().trim()
      });
    }
  }

  return {
    success: true,
    students: students
  };
}

/**
 * 試算表內一鍵生成隨機防偽 Token (6 碼隨機英文數字，避免混淆字元)
 */
function generateTokensForStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('學生名冊');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('找不到【學生名冊】工作表，請先建立！');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz';
  var updated = 0;

  for (var i = 1; i < data.length; i++) {
    var existingToken = data[i][4];
    if (!existingToken || existingToken.toString().trim() === '') {
      var token = '';
      for (var c = 0; c < 6; c++) {
        token += chars.charAt(Math.floor(Math.random() * chars.length));
      }
      sheet.getRange(i + 1, 5).setValue(token);
      sheet.getRange(i + 1, 6).setValue(true);
      updated++;
    }
  }

  SpreadsheetApp.getUi().alert('✨ 成功！已為 ' + updated + ' 位尚未有代碼的學生生成專屬代碼。');
}

/**
 * 自動為成績明細表計算班級平均 (可選輔助工具)
 */
function autoCalculateAverages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoreSheet = ss.getSheetByName('成績明細');
  if (!scoreSheet) return;

  var data = scoreSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  // 依據 (學年度 + 學期 + 考試名稱 + 科目) 計算平均
  var sumMap = {};
  var countMap = {};

  for (var i = 1; i < data.length; i++) {
    var key = data[i][1] + '|' + data[i][2] + '|' + data[i][3] + '|' + data[i][4];
    var score = Number(data[i][5]);
    if (!isNaN(score)) {
      sumMap[key] = (sumMap[key] || 0) + score;
      countMap[key] = (countMap[key] || 0) + 1;
    }
  }

  // 填回班級平均
  for (var j = 1; j < data.length; j++) {
    var k = data[j][1] + '|' + data[j][2] + '|' + data[j][3] + '|' + data[j][4];
    if (countMap[k]) {
      var avg = Math.round((sumMap[k] / countMap[k]) * 10) / 10;
      scoreSheet.getRange(j + 1, 7).setValue(avg);
    }
  }

  SpreadsheetApp.getUi().alert('✨ 班級平均已自動計算並更新完畢！');
}
