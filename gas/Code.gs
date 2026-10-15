/**
 * MPScore - 智慧學生成績與學習診斷系統 (Google Apps Script 後端)
 * 網站作者 / 系統規劃：楊家驊老師 (Chiahua Yang)
 *
 * 【核心架構 - 每次考試獨立分頁】：
 * 1. 【學生名冊】分頁：班級 | 座號 | 姓名 | 專屬代碼 | 啟用狀態
 * 2. 考試分頁（例如：114-1期中考、114-1期末考... 每次考試一個新分頁）：
 *    - 第 1 列：考試名稱標題（例如：114學年度第2學期期中考）
 *    - 第 2 列：座號 | 姓名 | 國語 | 數學 | 社會 | 英文 | 自然 (僅需上傳原始成績)
 *    - 第 3 列起：各學生原始分數
 *    ★ 老師只需上傳原始成績！所有「各科班級平均、個人平均、班級排名、PR值、等第區間、學習診斷評語」
 *      全由本系統自動運算，無需老師手動計算！
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ MPScore 工具')
    .addItem('🛠️ 一鍵自動建立標準分頁範本 (全新空白表專用)', 'setupStandardSheets')
    .addSeparator()
    .addItem('➕ 從【學生名冊】建立新考試分頁 (自動帶入全班姓名)', 'createNewExamSheetFromRoster')
    .addItem('🔄 從最新考試分頁同步名冊至【學生名冊】', 'syncStudentsFromExamSheet')
    .addItem('🔑 一鍵為全班產生防偽查詢代碼 (Token)', 'generateTokensForStudents')
    .addSeparator()
    .addItem('📊 預覽全班成績統計分析與各科平均', 'previewClassStatistics')
    .addToUi();
}

/**
 * 🛠️ 一鍵自動建立標準分頁範本
 * 建立【學生名冊】及兩個考試分頁【114-1期中考】、【114-1期末考】，供立即體驗多次考試歷程比對
 */
function setupStandardSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 1. 建立或格式化【學生名冊】
  var sSheet = ss.getSheetByName('學生名冊') || ss.getSheets()[0];
  sSheet.setName('學生名冊');
  var sHeaders = [['班級', '座號', '姓名', '專屬代碼', '啟用狀態']];
  sSheet.getRange(1, 1, 1, 5).setValues(sHeaders)
    .setBackground('#312e81') // 深靛藍
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  var sampleRoster = [
    ['701', '01', '陳小明', 'DEMO01', true],
    ['701', '02', '陳小華', 'DEMO02', true],
    ['701', '03', '陳大美', 'DEMO03', true],
    ['701', '04', '陳志豪', 'DEMO04', true],
    ['701', '06', '黃小琪', 'DEMO06', true],
    ['701', '07', '李大偉', 'DEMO07', true],
    ['701', '08', '林子涵', 'DEMO08', true],
    ['701', '09', '凌雅婷', 'DEMO09', true],
    ['701', '10', '劉建銘', 'DEMO10', true],
    ['701', '11', '王冠宇', 'DEMO11', true],
    ['701', '12', '吳宗憲', 'DEMO12', true],
    ['701', '13', '嚴立婷', 'DEMO13', true],
    ['701', '14', '余承恩', 'DEMO14', true],
    ['701', '15', '嚴佳慧', 'DEMO15', true],
    ['701', '16', '羅品安', 'DEMO16', true],
    ['701', '17', '羅廷宇', 'DEMO17', true],
    ['701', '18', '蕭雅云', 'DEMO18', true],
    ['701', '19', '張凱傑', 'DEMO19', true],
    ['701', '20', '茅文傑', 'DEMO20', true],
    ['701', '21', '洪宇成', 'DEMO21', true],
    ['701', '22', '林郁潔', 'DEMO22', true],
    ['701', '23', '周奕翔', 'DEMO23', true],
    ['701', '24', '吳秉翰', 'DEMO24', true],
    ['701', '25', '李勝吉', 'DEMO25', true]
  ];
  if (sSheet.getLastRow() <= 1) {
    sSheet.getRange(2, 1, sampleRoster.length, 5).setValues(sampleRoster);
    sSheet.autoResizeColumns(1, 5);
  }

  // 2. 建立第 1 次考試分頁【114-1期中考】(依楊老師真實成績單格式)
    var exam1Data = [
    ['01', '陳小明', 88, 70, 77, 100, 57],
    ['02', '陳小華', 86, 74, 81, 79, 69],
    ['03', '陳大美', 84, 90, 81, 86, 90],
    ['04', '陳志豪', 62, 80, 55, 55, 43],
    ['06', '黃小琪', 77, 76, 70, 92, 52],
    ['07', '李大偉', 92, 96, 96, 79, 89],
    ['08', '林子涵', 96, 96, 96, 98, 97],
    ['09', '凌雅婷', 72, 70, 80, 70, 58],
    ['10', '劉建銘', 69, 82, 79, 86, 79],
    ['11', '王冠宇', 47, 36, 45, 43, 30],
    ['12', '吳宗憲', 77, 76, 98, 56, 68],
    ['13', '嚴立婷', 89, 100, 94, 90, 85],
    ['14', '余承恩', 73, 64, 58, 64, 44],
    ['15', '嚴佳慧', 88, 90, 91, 97, 91],
    ['16', '羅品安', 97, 92, 94, 96, 80],
    ['17', '羅廷宇', 52, 56, 34, 41, 44],
    ['18', '蕭雅云', 90, 90, 95, 89, 94],
    ['19', '張凱傑', 76, 68, 65, 77, 54],
    ['20', '茅文傑', 88, 84, 92, 95, 83],
    ['21', '洪宇成', 69, 62, 66, 80, 57],
    ['22', '林郁潔', 87, 88, 92, 89, 89],
    ['23', '周奕翔', 100, 80, 92, 98, 89],
    ['24', '吳秉翰', 80, 70, 78, 31, 63],
    ['25', '李勝吉', 54, 36, 32, 35, 41]
  ];
  var exam1Sheet = ss.getSheetByName('114-1期中考') || ss.insertSheet('114-1期中考');
  formatExamSheet(exam1Sheet, '114學年度第1學期期中考', exam1Data);

  // 3. 建立第 2 次考試分頁【114-1期末考】(供跨次考試進退步對比)
  var exam2Data = exam1Data.map(function(st) {
    return [
      st[0],
      st[1],
      Math.min(100, st[2] + (st[0] === '01' ? 4 : 2)),
      Math.min(100, st[3] + (st[0] === '01' ? 8 : 1)),
      Math.min(100, st[4] + (st[0] === '01' ? 5 : 2)),
      Math.min(100, st[5] + (st[0] === '01' ? -2 : 3)),
      Math.min(100, st[6] + (st[0] === '01' ? 11 : 4))
    ];
  });
  var exam2Sheet = ss.getSheetByName('114-1期末考') || ss.insertSheet('114-1期末考');
  formatExamSheet(exam2Sheet, '114學年度第1學期期末考', exam2Data);

  ui.alert(
    '🎉 建置完成！\n\n' +
    '已為您建立：\n' +
    '1. 【學生名冊】(含專屬代碼與啟用狀態)\n' +
    '2. 【114-1期中考】(只需原始成績：座號、姓名、國語、數學、社會、英文、自然)\n' +
    '3. 【114-1期末考】(供跨次對比示範)\n\n' +
    '💡 小提醒：未來每次考試，只需點選「+」新增一個分頁（例如：114-2期中考），貼上學生的各科原始成績即可，網站會自動計算班級平均、排名與PR值！'
  );
}

/**
 * 輔助格式化考試分頁
 */
function formatExamSheet(sheet, examTitle, rowsData) {
  var headers = ['座號', '姓名', '國語', '數學', '社會', '英文', '自然'];
  
  // 第 1 列：考試名稱標題
  sheet.getRange(1, 1, 1, headers.length).merge()
    .setValue(examTitle)
    .setBackground('#1e293b') // 深灰藍
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');

  // 第 2 列：欄位標題
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setBackground('#334155')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 第 3 列起：學生原始成績
  if (rowsData && rowsData.length > 0) {
    sheet.getRange(3, 1, rowsData.length, headers.length).setValues(rowsData)
      .setHorizontalAlignment('center');
  }

  sheet.autoResizeColumns(1, headers.length);
}

/**
 * Web API 入口點 (doGet / doPost)
 */
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
 * 系統連線檢測 (ping)
 * 偵測試算表內【學生名冊】與所有考試分頁
 */
function checkSystemStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { success: false, message: '試算表內找不到名為【學生名冊】的分頁，請先點選【MPScore 工具 > 🛠️ 一鍵自動建立標準分頁範本】' };
  }

  var allSheets = ss.getSheets();
  var examTabs = [];
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    if (name !== '學生名冊' && !name.startsWith('_')) {
      examTabs.push(name);
    }
  }

  var studentCount = Math.max(0, sSheet.getLastRow() - 1);

  return {
    success: true,
    message: '連線正常！系統已就緒',
    sheetTitle: ss.getName(),
    studentCount: studentCount,
    examCount: examTabs.length,
    examTabs: examTabs,
    autoCalculationEnabled: true
  };
}

/**
 * 依據專屬 Token 查詢該名學生歷次考試成績與自動運算成果
 */
function getStudentScoreByToken(token) {
  if (!token) return { success: false, message: '請提供專屬查詢代碼' };
  token = token.toString().trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { success: false, message: '試算表缺少【學生名冊】分頁' };
  }

  // 1. 在【學生名冊】中尋找此代碼對應的學生
  var sData = sSheet.getDataRange().getValues();
  var student = null;

  for (var i = 1; i < sData.length; i++) {
    var row = sData[i];
    var rowCls = row[0] ? row[0].toString().trim() : '';
    var rowSeat = normalizeSeatNumber(row[1]);
    var rowName = row[2] ? row[2].toString().trim() : '';
    var rowToken = row[3] ? row[3].toString().trim() : ''; // D 欄: 專屬代碼
    var isEnabled = row[4]; // E 欄: 啟用狀態

    // 支援代碼 (不分大小寫)、學生姓名或座號精確比對
    var tokenUpper = token.toUpperCase();
    var isMatch = (rowToken && rowToken.toUpperCase() === tokenUpper) || 
                  (rowName && rowName === token) || 
                  (token.length <= 2 && rowSeat === normalizeSeatNumber(token));

    if (isMatch) {
      if (isEnabled === false) {
        return { success: false, message: '此學生帳號已被停用。' };
      }
      student = {
        className: rowCls,
        seatNumber: rowSeat,
        name: rowName,
        token: rowToken,
        studentId: (rowCls || '') + '-' + rowSeat
      };
      break;
    }
  }

  if (!student) {
    return { success: false, message: '驗證失敗：查無代碼或學生姓名「' + token + '」，請確認是否輸入正確。' };
  }

  // 2. 遍歷試算表中所有的考試分頁
  var allSheets = ss.getSheets();
  var exams = [];
  var flatRecords = [];

  for (var s = 0; s < allSheets.length; s++) {
    var curSheet = allSheets[s];
    var sheetName = curSheet.getName();
    if (sheetName === '學生名冊' || sheetName.startsWith('_')) {
      continue;
    }

    var examAnalysis = analyzeExamSheet(curSheet, student);
    if (examAnalysis && examAnalysis.studentFound) {
      exams.push(examAnalysis.examData);

      // 同時轉成平鋪的 records 陣列，完美向下相容既有圖表與報表
      var curExam = examAnalysis.examData;
      for (var subIdx = 0; subIdx < curExam.subjects.length; subIdx++) {
        var subItem = curExam.subjects[subIdx];
        flatRecords.push({
          examKey: curExam.examTitle,
          examName: curExam.examTitle,
          sheetName: curExam.sheetName,
          subject: subItem.name,
          score: subItem.score,
          average: subItem.classAverage,
          pr: subItem.pr,
          zone: subItem.zone,
          zoneStatus: subItem.zoneStatus,
          diagnosis: subItem.advice
        });
      }
    }
  }

  return {
    success: true,
    student: student,
    exams: exams,
    records: flatRecords
  };
}

/**
 * 深入解析單一考試分頁：全自動運算該次考試的全班平均、PR值、排名與個人診斷
 */
function analyzeExamSheet(sheet, student) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var sheetName = sheet.getName();
  var examTitle = sheetName;

  // 1. 判定標題與欄位標題列
  var headerRowIdx = 1; // 預設第 2 列 (索引 1)
  if (data.length > 0) {
    var row0Str = data[0].join(' ');
    // 如果第 1 列就包含「座號」或「姓名」，代表沒有第 1 列合併大標題
    if (row0Str.indexOf('座號') !== -1 || row0Str.indexOf('姓名') !== -1) {
      headerRowIdx = 0;
    } else {
      // 第 1 列為考試名稱大標題
      var titleCandidate = data[0][0] ? data[0][0].toString().trim() : '';
      if (titleCandidate) examTitle = titleCandidate;
      headerRowIdx = 1;
    }
  }

  var headerRow = data[headerRowIdx].map(function(h) { return h.toString().trim(); });
  var seatCol = -1;
  var nameCol = -1;

  for (var c = 0; c < headerRow.length; c++) {
    if (headerRow[c].indexOf('座號') !== -1) seatCol = c;
    else if (headerRow[c].indexOf('姓名') !== -1) nameCol = c;
  }

  if (seatCol === -1) seatCol = 0;
  if (nameCol === -1 && headerRow.length > 1) nameCol = 1;

  // 2. 辨識科目欄位（從姓名右側開始，到最後一個非平均/排名的有效欄位）
  var subjectCols = [];
  var startCol = Math.max(seatCol, nameCol) + 1;

  for (var sc = startCol; sc < headerRow.length; sc++) {
    var hName = headerRow[sc];
    if (!hName) continue;
    // 排除「平均」、「排名」或統計欄位（因為由網站自動精算）
    if (hName === '平均' || hName === '總分' || hName === '排名' || hName === '名次') {
      continue;
    }
    subjectCols.push({ index: sc, name: hName });
  }

  if (subjectCols.length === 0) return null;

  // 3. 收集所有學生的原始分數（自動排除底部的「平均」、「標準差」或空列）
  var studentsScores = [];
  var myStudentRow = null;

  for (var r = headerRowIdx + 1; r < data.length; r++) {
    var row = data[r];
    var seatVal = row[seatCol] ? row[seatCol].toString().trim() : '';
    var nameVal = nameCol !== -1 && row[nameCol] ? row[nameCol].toString().trim() : '';

    // 遇到平均、標準差或空列，代表學生資料已結束
    if (!seatVal || seatVal.indexOf('平均') !== -1 || seatVal.indexOf('標準差') !== -1 || seatVal.indexOf('分') !== -1) {
      break;
    }
    if (nameVal && (nameVal.indexOf('平均') !== -1 || nameVal.indexOf('標準差') !== -1)) {
      break;
    }

    var sObj = {
      seat: normalizeSeatNumber(seatVal),
      name: nameVal,
      scores: {},
      total: 0,
      count: 0
    };

    for (var sIdx = 0; sIdx < subjectCols.length; sIdx++) {
      var colInfo = subjectCols[sIdx];
      var rawScore = row[colInfo.index];
      var numScore = Number(rawScore);
      if (!isNaN(numScore) && rawScore !== '') {
        sObj.scores[colInfo.name] = numScore;
        sObj.total += numScore;
        sObj.count++;
      } else {
        sObj.scores[colInfo.name] = 0;
      }
    }

    sObj.average = sObj.count > 0 ? Math.round((sObj.total / sObj.count) * 100) / 100 : 0;
    studentsScores.push(sObj);

    // 比對目標學生
    if (sObj.seat === normalizeSeatNumber(student.seatNumber) || (student.name && sObj.name === student.name)) {
      myStudentRow = sObj;
    }
  }

  if (!myStudentRow || studentsScores.length === 0) {
    return { studentFound: false };
  }

  var totalStudents = studentsScores.length;

  // 4. 計算全班排名 (依學生個人平均由高至低)
  var sortedByAvg = studentsScores.slice().sort(function(a, b) {
    return b.average - a.average;
  });

  var myRank = 1;
  for (var rankIdx = 0; rankIdx < sortedByAvg.length; rankIdx++) {
    if (sortedByAvg[rankIdx].seat === myStudentRow.seat) {
      myRank = rankIdx + 1;
      break;
    }
  }

  // 5. 計算各科目的「班級平均」與「PR值」
  var subjectsResult = [];
  var classTotalAvgSum = 0;

  for (var j = 0; j < subjectCols.length; j++) {
    var subName = subjectCols[j].name;
    var allScoresForSub = [];

    for (var k = 0; k < studentsScores.length; k++) {
      allScoresForSub.push(studentsScores[k].scores[subName] || 0);
    }

    // 計算該科班級平均
    var subSum = allScoresForSub.reduce(function(a, b) { return a + b; }, 0);
    var subClassAvg = Math.round((subSum / totalStudents) * 100) / 100;
    classTotalAvgSum += subClassAvg;

    // 計算該名學生的 PR 值: (低於人數 + 0.5 * 同分人數) / 總人數 * 100
    var mySubScore = myStudentRow.scores[subName] || 0;
    var below = 0;
    var equal = 0;

    for (var p = 0; p < allScoresForSub.length; p++) {
      if (allScoresForSub[p] < mySubScore) below++;
      else if (allScoresForSub[p] === mySubScore) equal++;
    }

    var prVal = Math.round(((below + 0.5 * equal) / totalStudents) * 100);

    // 學習落點區間判定 (依楊家驊老師 GradeInterpr 標準)
    var zone = '穩健';
    var zoneStatus = '穩健發揮區間';
    var zoneType = 'leading';

    if (prVal >= 75) {
      zone = '精熟';
      zoneStatus = '精熟表現區間';
      zoneType = 'leading';
    } else if (prVal >= 50) {
      zone = '穩健';
      zoneStatus = '穩健發揮區間';
      zoneType = 'leading';
    } else if (prVal >= 25) {
      zone = '努力';
      zoneStatus = '努力成長區間';
      zoneType = 'improving';
    } else {
      zone = '開發';
      zoneStatus = '開發潛力區間';
      zoneType = 'improving';
    }

    // 產生專家建議
    var adviceText = getExpertSubjectAdvice(subName, zoneType, mySubScore);

    subjectsResult.push({
      name: subName,
      score: mySubScore,
      classAverage: subClassAvg,
      diff: Math.round((mySubScore - subClassAvg) * 100) / 100,
      pr: prVal,
      zone: zone,
      zoneStatus: zoneStatus,
      advice: adviceText
    });
  }

  var classOverallAvg = Math.round((classTotalAvgSum / subjectCols.length) * 100) / 100;

  return {
    studentFound: true,
    examData: {
      examTitle: examTitle,
      sheetName: sheetName,
      studentAverage: myStudentRow.average,
      studentRank: myRank,
      totalStudents: totalStudents,
      classOverallAverage: classOverallAvg,
      subjects: subjectsResult
    }
  };
}

/**
 * 標準化座號（例如：1 轉為 "01"）
 */
function normalizeSeatNumber(val) {
  if (val === null || val === undefined) return '';
  var s = val.toString().trim();
  var num = parseInt(s, 10);
  if (!isNaN(num) && num < 10 && s.length === 1) {
    return '0' + num;
  }
  return s;
}

/**
 * 取得全班學生名單及所有考試統計摘要 (供列印 QR Code 小卡與自選呈現統計)
 */
function getStudentListForTeacher() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('學生名冊');
  if (!sheet) return { success: false, message: '找不到【學生名冊】工作表' };

  var data = sheet.getDataRange().getValues();
  var students = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[2] && row[3]) { // 有姓名與代碼
      var seat = normalizeSeatNumber(row[1]);
      var cls = row[0] ? row[0].toString().trim() : '';
      students.push({
        className: cls,
        seatNumber: seat,
        name: row[2].toString().trim(),
        token: row[3].toString().trim(),
        studentId: (cls || '') + '-' + seat,
        examStats: {}
      });
    }
  }

  // 遍歷所有考試分頁，單次掃描即時運算全班名次與各科 PR
  var allSheets = ss.getSheets();
  var examList = [];

  for (var s = 0; s < allSheets.length; s++) {
    var curSheet = allSheets[s];
    var sName = curSheet.getName();
    if (sName === '學生名冊' || sName.startsWith('_')) continue;

    var examStats = getExamAllStudentsStats(curSheet);
    if (examStats) {
      examList.push(examStats.examTitle);
      for (var stIdx = 0; stIdx < students.length; stIdx++) {
        var stSeat = students[stIdx].seatNumber;
        var stName = students[stIdx].name;
        if (examStats.resultMap[stSeat]) {
          students[stIdx].examStats[examStats.examTitle] = examStats.resultMap[stSeat];
        } else if (stName && examStats.resultMap[stName]) {
          students[stIdx].examStats[examStats.examTitle] = examStats.resultMap[stName];
        }
      }
    }
  }

  return {
    success: true,
    students: students,
    examList: examList
  };
}

/**
 * 高速解析單一考試分頁內所有學生的統計摘要
 */
function getExamAllStudentsStats(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  var sheetName = sheet.getName();
  var examTitle = sheetName;
  var headerRowIdx = 1;
  if (data.length > 0 && (data[0].join(' ').indexOf('座號') !== -1 || data[0].join(' ').indexOf('姓名') !== -1)) {
    headerRowIdx = 0;
  } else if (data.length > 0 && data[0][0]) {
    examTitle = data[0][0].toString().trim();
  }

  var headerRow = data[headerRowIdx].map(function(h) { return h.toString().trim(); });
  var seatCol = 0, nameCol = 1;
  for (var c = 0; c < headerRow.length; c++) {
    if (headerRow[c].indexOf('座號') !== -1) seatCol = c;
    else if (headerRow[c].indexOf('姓名') !== -1) nameCol = c;
  }

  var subjectCols = [];
  var startCol = Math.max(seatCol, nameCol) + 1;
  for (var sc = startCol; sc < headerRow.length; sc++) {
    var hName = headerRow[sc];
    if (hName && hName !== '平均' && hName !== '總分' && hName !== '排名' && hName !== '名次') {
      subjectCols.push({ index: sc, name: hName });
    }
  }
  if (subjectCols.length === 0) return null;

  var studentsList = [];
  for (var r = headerRowIdx + 1; r < data.length; r++) {
    var row = data[r];
    var seatVal = row[seatCol] ? row[seatCol].toString().trim() : '';
    var nameVal = nameCol !== -1 && row[nameCol] ? row[nameCol].toString().trim() : '';
    if (!seatVal || seatVal.indexOf('平均') !== -1 || seatVal.indexOf('標準差') !== -1 || seatVal.indexOf('分') !== -1) break;
    if (nameVal && nameVal.indexOf('平均') !== -1) break;

    var sObj = {
      seat: normalizeSeatNumber(seatVal),
      name: nameVal,
      scores: {},
      total: 0,
      count: 0
    };
    for (var sIdx = 0; sIdx < subjectCols.length; sIdx++) {
      var colInfo = subjectCols[sIdx];
      var rawScore = row[colInfo.index];
      var numScore = Number(rawScore);
      if (!isNaN(numScore) && rawScore !== '') {
        sObj.scores[colInfo.name] = numScore;
        sObj.total += numScore;
        sObj.count++;
      } else {
        sObj.scores[colInfo.name] = 0;
      }
    }
    sObj.average = sObj.count > 0 ? Math.round((sObj.total / sObj.count) * 100) / 100 : 0;
    studentsList.push(sObj);
  }

  var total = studentsList.length;
  if (total === 0) return null;

  // 排序計算名次
  var sorted = studentsList.slice().sort(function(a, b) { return b.average - a.average; });
  var rankMap = {};
  for (var k = 0; k < sorted.length; k++) {
    rankMap[sorted[k].seat] = k + 1;
  }

  // 計算各科班均與各生 PR
  var subAvgs = {};
  var subScoresPool = {};
  var classTotalAvgSum = 0;
  for (var j = 0; j < subjectCols.length; j++) {
    var subName = subjectCols[j].name;
    var pList = [];
    for (var m = 0; m < total; m++) {
      pList.push(studentsList[m].scores[subName] || 0);
    }
    subScoresPool[subName] = pList;
    var sum = pList.reduce(function(a, b) { return a + b; }, 0);
    var avg = Math.round((sum / total) * 100) / 100;
    subAvgs[subName] = avg;
    classTotalAvgSum += avg;
  }
  var classOverallAvg = Math.round((classTotalAvgSum / subjectCols.length) * 100) / 100;

  var resultMap = {};
  for (var n = 0; n < total; n++) {
    var st = studentsList[n];
    var subjectsResult = [];
    for (var p = 0; p < subjectCols.length; p++) {
      var sName = subjectCols[p].name;
      var myScore = st.scores[sName] || 0;
      var pool = subScoresPool[sName];
      var below = 0, equal = 0;
      for (var q = 0; q < pool.length; q++) {
        if (pool[q] < myScore) below++;
        else if (pool[q] === myScore) equal++;
      }
      var pr = Math.round(((below + 0.5 * equal) / total) * 100);
      subjectsResult.push({
        name: sName,
        score: myScore,
        classAverage: subAvgs[sName],
        pr: pr
      });
    }

    var statObj = {
      examTitle: examTitle,
      sheetName: sheetName,
      studentAverage: st.average,
      studentRank: rankMap[st.seat],
      totalStudents: total,
      classOverallAverage: classOverallAvg,
      subjects: subjectsResult
    };
    resultMap[st.seat] = statObj;
    if (st.name) {
      resultMap[st.name] = statObj;
    }
  }

  return {
    examTitle: examTitle,
    sheetName: sheetName,
    resultMap: resultMap
  };
}

/**
 * 🔄 從最新考試分頁自動同步座號與姓名至【學生名冊】並生成代碼
 * 老師只要貼上考試成績，點此按鈕立即完成全班防偽名冊建置！
 */
function syncStudentsFromExamSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    sSheet = ss.insertSheet('學生名冊', 0);
    sSheet.getRange(1, 1, 1, 5).setValues([['班級', '座號', '姓名', '專屬代碼', '啟用狀態']])
      .setBackground('#312e81').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  // 尋找第一個非學生名冊的考試分頁
  var sheets = ss.getSheets();
  var examSheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== '學生名冊' && !sheets[i].getName().startsWith('_')) {
      examSheet = sheets[i];
      break;
    }
  }

  if (!examSheet) {
    ui.alert('請先建立至少一個考試分頁（例如：114-1期中考）並貼上座號與姓名！');
    return;
  }

  var data = examSheet.getDataRange().getValues();
  var headerRowIdx = 1;
  if (data.length > 0 && (data[0].join(' ').indexOf('座號') !== -1 || data[0].join(' ').indexOf('姓名') !== -1)) {
    headerRowIdx = 0;
  }

  var headerRow = data[headerRowIdx].map(function(h) { return h.toString().trim(); });
  var seatCol = 0;
  var nameCol = 1;
  for (var c = 0; c < headerRow.length; c++) {
    if (headerRow[c].indexOf('座號') !== -1) seatCol = c;
    else if (headerRow[c].indexOf('姓名') !== -1) nameCol = c;
  }

  // 讀取既有代碼避免覆蓋
  var existingTokens = {};
  var sData = sSheet.getDataRange().getValues();
  for (var s = 1; s < sData.length; s++) {
    var exSeat = normalizeSeatNumber(sData[s][1]);
    var exTok = sData[s][3];
    if (exSeat && exTok) existingTokens[exSeat] = exTok;
  }

  var chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz';
  var newRoster = [];

  for (var r = headerRowIdx + 1; r < data.length; r++) {
    var seatVal = data[r][seatCol] ? data[r][seatCol].toString().trim() : '';
    var nameVal = data[r][nameCol] ? data[r][nameCol].toString().trim() : '';

    if (!seatVal || seatVal.indexOf('平均') !== -1 || seatVal.indexOf('標準差') !== -1) break;
    if (nameVal && nameVal.indexOf('平均') !== -1) break;

    var normSeat = normalizeSeatNumber(seatVal);
    var token = existingTokens[normSeat];
    if (!token) {
      token = '';
      for (var t = 0; t < 6; t++) token += chars.charAt(Math.floor(Math.random() * chars.length));
    }

    newRoster.push(['701', normSeat, nameVal, token, true]);
  }

  if (newRoster.length === 0) {
    ui.alert('在考試分頁【' + examSheet.getName() + '】中未偵測到學生座號與姓名！');
    return;
  }

  sSheet.getRange(2, 1, sSheet.getLastRow() > 1 ? sSheet.getLastRow() - 1 : 1, 5).clearContent();
  sSheet.getRange(2, 1, newRoster.length, 5).setValues(newRoster);
  sSheet.autoResizeColumns(1, 5);

  ui.alert('🎉 成功！已從【' + examSheet.getName() + '】同步 ' + newRoster.length + ' 位學生座號姓名，並建立專屬防偽代碼！');
}

/**
 * 為【學生名冊】中尚未有代碼的學生生成 6 碼隨機 Token
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
    var existingToken = data[i][3];
    if (!existingToken || existingToken.toString().trim() === '') {
      var token = '';
      for (var c = 0; c < 6; c++) {
        token += chars.charAt(Math.floor(Math.random() * chars.length));
      }
      sheet.getRange(i + 1, 4).setValue(token);
      sheet.getRange(i + 1, 5).setValue(true);
      updated++;
    }
  }

  SpreadsheetApp.getUi().alert('✨ 成功！已為 ' + updated + ' 位學生生成專屬查詢代碼。');
}

/**
 * 📊 預覽全班成績統計分析與各科平均
 */
function previewClassStatistics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheets = ss.getSheets();
  var report = '📊 MPScore 即時成績精算分析：\n\n';

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var sName = sh.getName();
    if (sName === '學生名冊' || sName.startsWith('_')) continue;

    var mockStudent = { seatNumber: '01', name: '' };
    var result = analyzeExamSheet(sh, mockStudent);
    if (result && result.examData) {
      var d = result.examData;
      report += '【' + d.examTitle + '】(學生 ' + d.totalStudents + ' 人，總平均 ' + d.classOverallAverage + ' 分)\n';
      for (var s = 0; s < d.subjects.length; s++) {
        report += '  • ' + d.subjects[s].name + ' 班平均: ' + d.subjects[s].classAverage + ' 分\n';
      }
      report += '\n';
    }
  }

  ui.alert(report);
}

/**
 * 來自楊家驊老師 GradeInterpr 的專家學習診斷建議生成器
 */
function getExpertSubjectAdvice(subject, zoneType, score) {
  var adviceDb = {
    '國語': {
      leading: [
        "你對文字的敏感度很高，可以多接觸不同背景的文學作品，練習體會文字背後的弦外之音。",
        "嘗試將閱讀後的感想與生活連結，這能幫助你將零散的思緒整理成更有系統的觀點。",
        "在閱讀長篇文章後，練習提取幾個關鍵概念，這對掌握複雜文本非常有幫助。"
      ],
      improving: [
        "建議透過『大聲朗讀』來感受課文的語氣與節奏，這對於建立語感很有幫助。",
        "可以試著將課文中的優美詞彙應用在平日的簡短紀錄中，這會讓這些文字真正變成你的。",
        "先從自己感興趣的故事主題開始讀起，累積對文字的親近感，閱讀會變得越來越輕鬆。"
      ]
    },
    '國文': {
      leading: [
        "你對文字的敏感度很高，可以多接觸不同背景的文學作品，練習體會文字背後的弦外之音。",
        "嘗試將閱讀後的感想與生活連結，這能幫助你將零散的思緒整理成更有系統的觀點。"
      ],
      improving: [
        "建議透過『大聲朗讀』來感受課文的語氣與節奏，這對於建立語感很有幫助。",
        "可以試著將課文中的優美詞彙應用在平日的簡短紀錄中，這會讓這些文字真正變成你的。"
      ]
    },
    '數學': {
      leading: [
        "你展現了優秀的思維，可以試著探索公式背後的由來，理解原理會讓應用更靈活。",
        "練習看看『一題多解』，從不同的路徑切入問題，能讓你的思維死角越來越少。",
        "試著把學過的邏輯教給別人看，當你能講得通時，就代表你已經內化了這個概念。"
      ],
      improving: [
        "學習數學像是堆積木，先把課本裡的基礎定義弄清楚，基礎穩了，解題就會有成就感。",
        "不需要急著寫很難的題目，先把基本例題反覆練習幾次，確保運算的流程是流暢的。",
        "遇到卡住的題目時，可以換個心情再回頭看，或者試著畫出圖表來幫助大腦思考。"
      ]
    },
    '社會': {
      leading: [
        "你很擅長整理資訊，可以練習將不同時間的事件串聯起來，看見歷史或地理背後的因果。",
        "關注一下最近的新聞時事，並思考課本內容在現實中是如何運作的，這能讓知識變得很立體。",
        "試著從不同的社會角色來看同一件事，這能幫助你培養更全面、更具深度的思辨力。"
      ],
      improving: [
        "社會科其實是許多精彩故事的集合。想像自己是當時的人，這能幫你理解為何會有那樣的發展。",
        "對於複雜的資訊，可以試著用畫圖或拉時間軸的方式呈現，讓文字敘述變成圖像記憶。",
        "可以從你喜歡的地圖、圖表或特定人物故事開始讀起，興趣是理解學科最好的動力。"
      ]
    },
    '英文': {
      leading: [
        "你的語感表現優異，可以開始接觸課外的英文讀物或短片，挑戰更自然的語言環境。",
        "練習聽不同背景的英文資源，這能提升你在真實生活情境中的聽力適應能力。",
        "試著用英文寫下簡短的計畫或心情，讓語言成為你生活中自然的一部分。"
      ],
      improving: [
        "英文需要時間來讓耳朵習慣。每天花幾分鐘聽聽朗讀，即便沒聽懂全部，也是在累積感度。",
        "每天掌握幾個對你有意義、用得到的單字，積少成多會產生非常驚人的進步。",
        "不要怕犯錯，試著勇敢讀出聲來。語言是溝通的橋樑，只要敢開口就是跨出了一大步。"
      ]
    },
    '英語': {
      leading: [
        "你的語感表現優異，可以開始接觸課外的英文讀物或短片，挑戰更自然的語言環境。",
        "練習聽不同背景的英文資源，這能提升你在真實生活情境中的聽力適應能力。"
      ],
      improving: [
        "英文需要時間來讓耳朵習慣。每天花幾分鐘聽聽朗讀，即便沒聽懂全部，也是在累積感度。",
        "每天掌握幾個對你有意義、用得到的單字，積少成多會產生非常驚人的進步。"
      ]
    },
    '自然': {
      leading: [
        "你擁有敏銳的觀察力。對於生活中的現象，試著用科學的視角去推論它發生的原因。",
        "可以閱讀一些科普短文或新聞，探索前瞻的科學發展，這會大大擴展你的科學視野。",
        "針對實驗結果，試著提出『如果改變某個條件會怎樣？』的假設，這是科學探究的靈魂。"
      ],
      improving: [
        "科學其實就在你的生活週遭。觀察植物生長或廚房裡的物理現象，你會發現原理很親切。",
        "試著將課本上的圖表與生活現象對照，找回對世界運作的好奇心，學習會更有動力。",
        "先從記錄你觀察到的細節開始。當你開始好奇『這是在做什麼？』時，你就跨出了第一步。"
      ]
    }
  };

  var subData = adviceDb[subject] || adviceDb['自然'];
  var list = subData[zoneType] || subData['improving'];
  var seed = (subject.charCodeAt(0) + Math.floor(score)) % list.length;
  return list[seed] || list[0];
}

/**
 * ➕ 從【學生名冊】一鍵建立全新考試分頁（自動帶入全班座號與姓名）
 * 解決老師每次考試都要手動重打學生座號與姓名的困擾！
 */
function createNewExamSheetFromRoster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    ui.alert('找不到【學生名冊】分頁，請先點選【🛠️ 一鍵自動建立標準分頁範本】！');
    return;
  }

  var promptResult = ui.prompt(
    '➕ 建立新考試分頁',
    '請輸入考試分頁名稱（例如：114-1第二次段考、114-2期中考、隨堂測驗）：',
    ui.ButtonSet.OK_CANCEL
  );

  if (promptResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var examTitle = promptResult.getResponseText().trim();
  if (!examTitle) {
    ui.alert('考試名稱不可為空白！');
    return;
  }

  if (ss.getSheetByName(examTitle)) {
    ui.alert('分頁【' + examTitle + '】已經存在！請使用不同名稱。');
    return;
  }

  // 從【學生名冊】讀取所有學生的座號與姓名
  var sData = sSheet.getDataRange().getValues();
  var rosterRows = [];

  for (var i = 1; i < sData.length; i++) {
    var seat = normalizeSeatNumber(sData[i][1]);
    var name = sData[i][2] ? sData[i][2].toString().trim() : '';
    if (seat || name) {
      // 座號, 姓名, 國語, 數學, 社會, 英文, 自然 (分數留空讓老師直接登錄)
      rosterRows.push([seat, name, '', '', '', '', '']);
    }
  }

  if (rosterRows.length === 0) {
    ui.alert('【學生名冊】中沒有偵測到學生資料，請先在【學生名冊】中填入學生！');
    return;
  }

  // 插入新分頁並自動排版
  var newSheet = ss.insertSheet(examTitle);
  formatExamSheet(newSheet, examTitle, rosterRows);

  ui.alert(
    '🎉 建立完成！\n\n' +
    '已成功建立新考試分頁【' + examTitle + '】！\n' +
    '已為您自動帶入全班 ' + rosterRows.length + ' 位學生的座號與姓名。\n\n' +
    '💡 老師您只需在各科欄位填入原始分數，MPScore 網站就會全自動算好班平均與排名！'
  );
}

/**
 * 接收網站發送之成績資料，直接建立並寫入新考試分頁
 */
function appendExamScores(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var payload = typeof body === 'string' ? JSON.parse(body) : body;
  var examTitle = (payload.examTitle || '新考試成績').trim();
  var rows = payload.rows || [];

  if (rows.length === 0) {
    return { success: false, message: '沒有成績資料可儲存' };
  }

  var sheet = ss.getSheetByName(examTitle);
  if (!sheet) {
    sheet = ss.insertSheet(examTitle);
  } else {
    sheet.clear();
  }

  var formattedRows = rows.map(function(r) {
    return [
      normalizeSeatNumber(r.seat),
      r.name || '',
      r.scores && r.scores['國語'] !== undefined ? r.scores['國語'] : (r['國語'] || ''),
      r.scores && r.scores['數學'] !== undefined ? r.scores['數學'] : (r['數學'] || ''),
      r.scores && r.scores['社會'] !== undefined ? r.scores['社會'] : (r['社會'] || ''),
      r.scores && r.scores['英文'] !== undefined ? r.scores['英文'] : (r['英文'] || ''),
      r.scores && r.scores['自然'] !== undefined ? r.scores['自然'] : (r['自然'] || '')
    ];
  });

  formatExamSheet(sheet, examTitle, formattedRows);

  return {
    success: true,
    message: '成功建立並上傳【' + examTitle + '】成績分頁，共 ' + formattedRows.length + ' 筆學生資料！',
    sheetName: examTitle,
    studentCount: formattedRows.length
  };
}

// ========================================================
// 核心模組導出 (供 Loader.gs 自動熱更新引導器調用)
// ========================================================
var MPScoreEngine = {
  setupStandardSheets: setupStandardSheets,
  createNewExamSheetFromRoster: createNewExamSheetFromRoster,
  syncStudentsFromExamSheet: syncStudentsFromExamSheet,
  generateTokensForStudents: generateTokensForStudents,
  previewClassStatistics: previewClassStatistics,
  handleDoGet: doGet,
  handleDoPost: doPost
};

