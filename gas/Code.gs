/**
 * MPScore - 智慧學生成績與學習診斷系統 (Google Apps Script 後端)
 * 配合自訂欄位順序：
 * 【學生名冊】A: 班級, B: 座號, C: 姓名, D: 專屬代碼, E: 啟用狀態
 * 【成績明細】動態欄位辨識（支援 座號、學年度、學期、考試名稱、科目、個人得分、班級平均、PR值、解讀評語）
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ MPScore 工具')
    .addItem('🔑 一鍵為全班產生防偽代碼 (Token)', 'generateTokensForStudents')
    .addItem('📊 自動計算班級平均與 PR 值', 'autoCalculateStats')
    .addToUi();
}

function doGet(e) {
  var params = e ? e.parameter : {};
  var action = params.action || (params.token ? 'query' : 'ping');
  var output = {};

  try {
    if (action === 'ping') {
      output = checkSystemStatus();
    } else if (action === 'query') {
      output = getStudentScoreByToken(params.token);
    } else if (action === 'getStudentsForPrint') {
      output = getStudentListForTeacher();
    } else if (action === 'appendExamScores') {
      var body = params.payload || (e && e.postData && e.postData.contents ? e.postData.contents : '');
      output = appendExamScores(body);
    } else {
      output = { success: false, message: '未知的操作請求' };
    }
  } catch (err) {
    output = { success: false, message: '系統錯誤: ' + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}

/**
 * 從網頁端直接新增一整次考試成績到【成績明細】表的最下方 (不覆蓋舊考試)
 */
function appendExamScores(payloadStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scSheet = ss.getSheetByName('成績明細');
  if (!scSheet) return { success: false, message: '找不到【成績明細】分頁' };

  if (!payloadStr) return { success: false, message: '請提供要新增的成績資料' };
  var payload = typeof payloadStr === 'string' ? JSON.parse(payloadStr) : payloadStr;
  var rows = payload.rows;
  if (!rows || rows.length === 0) return { success: false, message: '成績列表為空' };

  var newRows = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    newRows.push([
      r.seatNumber,
      r.year,
      r.semester,
      r.examName,
      r.subject,
      Number(r.score) || 0,
      r.average ? Number(r.average) : '',
      r.pr ? Number(r.pr) : '',
      r.diagnosis || ''
    ]);
  }

  var startRow = scSheet.getLastRow() + 1;
  scSheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);

  return {
    success: true,
    message: '成功在【成績明細】最下方新增 ' + newRows.length + ' 筆考試成績！',
    count: newRows.length
  };
}

function checkSystemStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  var scSheet = ss.getSheetByName('成績明細');

  if (!sSheet || !scSheet) {
    return { success: false, message: '工作表分頁需有名為【學生名冊】與【成績明細】' };
  }

  return {
    success: true,
    message: '連線正常！系統已就緒',
    sheetTitle: ss.getName(),
    studentCount: Math.max(0, sSheet.getLastRow() - 1),
    scoreCount: Math.max(0, scSheet.getLastRow() - 1)
  };
}

/**
 * 依據專屬 Token 查詢該名學生歷年成績
 */
function getStudentScoreByToken(token) {
  if (!token) return { success: false, message: '請提供專屬查詢代碼' };
  token = token.toString().trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  var scSheet = ss.getSheetByName('成績明細');

  if (!sSheet || !scSheet) {
    return { success: false, message: '試算表分頁設定不正確' };
  }

  // 1. 學生名冊欄位：A:班級, B:座號, C:姓名, D:專屬代碼, E:啟用狀態
  var sData = sSheet.getDataRange().getValues();
  var student = null;

  for (var i = 1; i < sData.length; i++) {
    var row = sData[i];
    var rowToken = row[3] ? row[3].toString().trim() : ''; // D 欄: 專屬代碼
    var isEnabled = row[4]; // E 欄: 啟用狀態

    if (rowToken === token) {
      if (isEnabled === false) {
        return { success: false, message: '此代碼帳號已被停用。' };
      }
      student = {
        className: row[0].toString().trim(),
        seatNumber: row[1].toString().trim(),
        name: row[2].toString().trim(),
        studentId: (row[0] + '-' + (row[1] < 10 && row[1].length === 1 ? '0' + row[1] : row[1]))
      };
      break;
    }
  }

  if (!student) {
    return { success: false, message: '驗證失敗：查無此代碼，請確認是否輸入正確。' };
  }

  // 2. 抓取成績明細
  var scData = scSheet.getDataRange().getValues();
  if (scData.length <= 1) {
    return { success: true, student: student, records: [] };
  }

  // 智慧欄位索引判定
  var headers = scData[0].map(function(h) { return h.toString().trim(); });
  var colIndex = {
    seat: headers.indexOf('座號'),
    className: headers.indexOf('班級'),
    year: headers.indexOf('學年度'),
    semester: headers.indexOf('學期'),
    examName: headers.indexOf('考試名稱'),
    subject: headers.indexOf('科目'),
    score: headers.indexOf('個人得分'),
    average: headers.indexOf('班級平均'),
    pr: headers.indexOf('PR值'),
    diagnosis: headers.indexOf('解讀評語')
  };

  // 若沒找到座號標題，預設嘗試從前幾欄推斷
  if (colIndex.seat === -1 && colIndex.className === -1) {
    colIndex.seat = 0; // 假定第 1 欄為座號
  }

  // 收集全班各科目成績，以便在 PR 欄空白時自動即時計算 PR
  var examSubjectPool = {};

  for (var k = 1; k < scData.length; k++) {
    var rowK = scData[k];
    var eKey = (colIndex.year !== -1 ? rowK[colIndex.year] : '') + '_' +
               (colIndex.semester !== -1 ? rowK[colIndex.semester] : '') + '_' +
               (colIndex.examName !== -1 ? rowK[colIndex.examName] : '') + '_' +
               (colIndex.subject !== -1 ? rowK[colIndex.subject] : '');
    var sc = Number(colIndex.score !== -1 ? rowK[colIndex.score] : 0);

    if (!examSubjectPool[eKey]) examSubjectPool[eKey] = [];
    if (!isNaN(sc)) examSubjectPool[eKey].push(sc);
  }

  // 撈取該學生的成績
  var records = [];
  var targetSeat = student.seatNumber;
  var targetClass = student.className;

  for (var j = 1; j < scData.length; j++) {
    var r = scData[j];
    
    // 比對學生：若有班級與座號就同時比對，否則比對座號
    var rowSeat = colIndex.seat !== -1 ? r[colIndex.seat].toString().trim() : '';
    var rowClass = colIndex.className !== -1 ? r[colIndex.className].toString().trim() : '';
    
    var isMatch = false;
    if (colIndex.className !== -1 && colIndex.seat !== -1) {
      isMatch = (rowClass === targetClass && rowSeat === targetSeat);
    } else if (colIndex.seat !== -1) {
      isMatch = (rowSeat === targetSeat);
    }

    if (isMatch) {
      var scoreVal = Number(colIndex.score !== -1 ? r[colIndex.score] : 0) || 0;
      var avgVal = Number(colIndex.average !== -1 ? r[colIndex.average] : 0) || 0;
      var prVal = Number(colIndex.pr !== -1 ? r[colIndex.pr] : 0) || 0;

      // 如果試算表 PR 值未填，自動透過同科全體得分計算 PR
      var eKeyMatch = (colIndex.year !== -1 ? r[colIndex.year] : '') + '_' +
                      (colIndex.semester !== -1 ? r[colIndex.semester] : '') + '_' +
                      (colIndex.examName !== -1 ? r[colIndex.examName] : '') + '_' +
                      (colIndex.subject !== -1 ? r[colIndex.subject] : '');
      
      if ((!prVal || prVal <= 0) && examSubjectPool[eKeyMatch] && examSubjectPool[eKeyMatch].length > 1) {
        var pool = examSubjectPool[eKeyMatch];
        var belowCount = 0;
        var equalCount = 0;
        for (var p = 0; p < pool.length; p++) {
          if (pool[p] < scoreVal) belowCount++;
          else if (pool[p] === scoreVal) equalCount++;
        }
        // 常模百分等級公式: (低於人數 + 0.5 * 同分人數) / 總人數 * 100
        prVal = Math.round(((belowCount + 0.5 * equalCount) / pool.length) * 100);
      }

      records.push({
        year: colIndex.year !== -1 ? r[colIndex.year] : '',
        semester: colIndex.semester !== -1 ? r[colIndex.semester] : '',
        examName: colIndex.examName !== -1 ? r[colIndex.examName] : '',
        examKey: (colIndex.year !== -1 ? r[colIndex.year] : '') + '-' +
                 (colIndex.semester !== -1 ? r[colIndex.semester] : '') + ' ' +
                 (colIndex.examName !== -1 ? r[colIndex.examName] : ''),
        subject: colIndex.subject !== -1 ? r[colIndex.subject] : '',
        score: scoreVal,
        average: avgVal,
        pr: prVal,
        diagnosis: (colIndex.diagnosis !== -1 ? r[colIndex.diagnosis] : '') || ''
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
 * 取得全班學生清單 (供列印 QR Code 小卡)
 */
function getStudentListForTeacher() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('學生名冊');
  if (!sheet) return { success: false, message: '找不到【學生名冊】工作表' };

  var data = sheet.getDataRange().getValues();
  var students = [];

  // A: 班級, B: 座號, C: 姓名, D: 專屬代碼, E: 啟用狀態
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[2] && row[3]) { // 有姓名與代碼
      students.push({
        className: row[0].toString().trim(),
        seatNumber: row[1].toString().trim(),
        name: row[2].toString().trim(),
        token: row[3].toString().trim(),
        studentId: row[0] + '-' + (row[1] < 10 && row[1].length === 1 ? '0' + row[1] : row[1])
      });
    }
  }

  return {
    success: true,
    students: students
  };
}

/**
 * 自動為全班生成 6 碼隨機 Token (填入 D 欄，並在 E 欄設為 TRUE)
 */
function generateTokensForStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('學生名冊');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('找不到【學生名冊】工作表！');
    return;
  }

  var data = sheet.getDataRange().getValues();
  var chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz';
  var updated = 0;

  for (var i = 1; i < data.length; i++) {
    var existingToken = data[i][3]; // D 欄是索引 3
    if (!existingToken || existingToken.toString().trim() === '') {
      var token = '';
      for (var c = 0; c < 6; c++) {
        token += chars.charAt(Math.floor(Math.random() * chars.length));
      }
      sheet.getRange(i + 1, 4).setValue(token); // 填入 D 欄
      sheet.getRange(i + 1, 5).setValue(true);  // 填入 E 欄 (TRUE)
      updated++;
    }
  }

  SpreadsheetApp.getUi().alert('✨ 成功！已為 ' + updated + ' 位學生生成專屬查詢代碼。');
}

/**
 * 自動計算班級平均與 PR 值並回填試算表 (輔助工具)
 */
function autoCalculateStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scSheet = ss.getSheetByName('成績明細');
  if (!scSheet) return;

  var data = scSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var avgCol = headers.indexOf('班級平均');
  var prCol = headers.indexOf('PR值');
  var scoreCol = headers.indexOf('個人得分');
  var subjCol = headers.indexOf('科目');
  var examCol = headers.indexOf('考試名稱');

  if (scoreCol === -1 || subjCol === -1) {
    SpreadsheetApp.getUi().alert('找不到【個人得分】或【科目】欄位！');
    return;
  }

  // 分群計算
  var groupMap = {};
  for (var i = 1; i < data.length; i++) {
    var key = data[i][examCol] + '_' + data[i][subjCol];
    var sc = Number(data[i][scoreCol]);
    if (!groupMap[key]) groupMap[key] = [];
    if (!isNaN(sc)) groupMap[key].push(sc);
  }

  // 回填平均與 PR
  for (var j = 1; j < data.length; j++) {
    var k = data[j][examCol] + '_' + data[j][subjCol];
    var list = groupMap[k];
    if (list && list.length > 0) {
      // 平均
      var sum = list.reduce(function(a, b) { return a + b; }, 0);
      var avg = Math.round((sum / list.length) * 10) / 10;
      if (avgCol !== -1) scSheet.getRange(j + 1, avgCol + 1).setValue(avg);

      // PR
      var myScore = Number(data[j][scoreCol]);
      var below = 0, equal = 0;
      for (var p = 0; p < list.length; p++) {
        if (list[p] < myScore) below++;
        else if (list[p] === myScore) equal++;
      }
      var pr = Math.round(((below + 0.5 * equal) / list.length) * 100);
      if (prCol !== -1) scSheet.getRange(j + 1, prCol + 1).setValue(pr);
    }
  }

  SpreadsheetApp.getUi().alert('✨ 班級平均與 PR 值已全數計算並更新完成！');
}
