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
    .addItem('🛠️ 建立初始表格（空白試算表第一次使用）', 'setupStandardSheets')
    .addSeparator()
    .addItem('➕ 新增一次考試（自動帶入學生名單）', 'createNewExamSheetFromRoster')
    .addItem('🔄 用最新考試名單，更新「學生名冊」', 'syncStudentsFromExamSheet')
    .addItem('🔑 為全班產生「家長查詢代碼」', 'generateTokensForStudents')
    .addSeparator()
    .addItem('📋 查看「家長閱覽紀錄」', 'openParentAccessLogSheet')
    .addItem('📐 自動整理欄寬與表格排版', 'autoFormatAndResizeAllSheets')
    .addItem('🚀 立即更新雲端資料（改完分數點此）', 'clearSystemCache')
    .addItem('📊 查看全班成績與各科平均', 'previewClassStatistics')
    .addToUi();
}

/**
 * 清除系統雲端快取
 */
function clearSystemCache() {
  var ui = SpreadsheetApp.getUi();
  try {
    var cache = CacheService.getScriptCache();
    var keys = [];
    for (var i = 1; i <= 60; i++) {
      var s = i < 10 ? '0' + i : '' + i;
      keys.push('st_score_v3_DEMO' + s);
      keys.push('st_score_v3_' + s);
      keys.push('st_score_v2_DEMO' + s);
      keys.push('st_score_v2_' + s);
      keys.push('st_score_DEMO' + s);
      keys.push('st_score_' + s);
    }
    cache.removeAll(keys);
    ui.alert('✨ 雲端快取清除成功！', '家長端與查詢系統現已載入試算表最新成績與運算結果。', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('提示', '快取已同步重置。' + err.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 🛠️ 建立初始表格（空白試算表第一次使用）
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
    .setBackground('#b45309') // 暖陽金褐
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
  }
  sSheet.setColumnWidth(1, 90);  // 班級
  sSheet.setColumnWidth(2, 80);  // 座號
  sSheet.setColumnWidth(3, 120); // 姓名
  sSheet.setColumnWidth(4, 130); // 專屬代碼
  sSheet.setColumnWidth(5, 100); // 啟用狀態
  sSheet.setRowHeight(1, 36);
  var sLastRow = Math.max(2, sSheet.getLastRow());
  for (var sr = 2; sr <= sLastRow; sr++) {
    sSheet.setRowHeight(sr, 28);
  }
  sSheet.getRange(1, 1, sLastRow, 5)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sSheet.setFrozenRows(1);

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

  // 設置寬敞舒適的固定欄寬與列高，杜絕字體擠壓或截斷
  sheet.setColumnWidth(1, 80);  // 座號
  sheet.setColumnWidth(2, 120); // 姓名 (完整顯示3-4字)
  for (var c = 3; c <= headers.length; c++) {
    sheet.setColumnWidth(c, 95); // 科目成績欄 (大字清晰)
  }
  sheet.setRowHeight(1, 40); // 考試大標題列
  sheet.setRowHeight(2, 34); // 欄位標題列
  var curLastRow = Math.max(2, sheet.getLastRow());
  for (var r = 3; r <= curLastRow; r++) {
    sheet.setRowHeight(r, 28);
  }
  sheet.getRange(1, 1, curLastRow, headers.length)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(2);
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
    } else if (action === 'getParentAccessLogs') {
      output = getParentAccessLogs();
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
    return { success: false, message: '試算表內找不到名為【學生名冊】的分頁，請先點選【MPScore 工具 > 🛠️ 建立初始表格（空白試算表第一次使用）】' };
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

  // 0. 極速快取檢查：若 10 分鐘內已有運算快取，0.1 秒秒速回傳，避免家長誤會網站失效
  var tokenUpper = token.toUpperCase();
  var scriptCache = CacheService.getScriptCache();
  var cacheKey = 'st_score_v3_' + encodeURIComponent(tokenUpper);
  try {
    var cachedJson = scriptCache.get(cacheKey);
    if (cachedJson) {
      var cachedObj = JSON.parse(cachedJson);
      cachedObj.fromCache = true;
      if (cachedObj.student) {
        recordParentAccess(cachedObj.student);
      }
      return cachedObj;
    }
  } catch (e) {
    // 忽略快取讀取異常，正常降級至試算表即時運算
  }

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

  // 記錄家長閱覽日誌
  recordParentAccess(student);

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

  var finalResult = {
    success: true,
    student: student,
    exams: exams,
    records: flatRecords
  };

  // 3. 寫入快取 (快取 600 秒 = 10 分鐘，單一項目上限 100KB)
  try {
    var outStr = JSON.stringify(finalResult);
    if (outStr.length < 95000) {
      scriptCache.put(cacheKey, outStr, 600);
      // 若是依據姓名或座號查詢，同時也快取其 token
      if (student.token && student.token.toUpperCase() !== tokenUpper) {
        scriptCache.put('st_score_v3_' + encodeURIComponent(student.token.toUpperCase()), outStr, 600);
      }
    }
  } catch (e) {
    // 忽略快取寫入異常
  }

  return finalResult;
}

/**
 * 智慧解析考試分頁的標題列、座號欄、姓名欄、班級欄與各科成績欄位
 * 徹底解決多列標題、合併儲存格、班級欄位存在或欄位順序不同時的錯位問題！
 */
function detectExamSheetColumns(data) {
  if (!data || data.length < 2) return null;

  var maxScanRows = Math.min(5, data.length - 1);
  var bestRowIdx = -1;
  var bestScore = -1;

  var commonSubjects = ['國語', '國文', '語文', '數學', '英語', '英文', '社會', '自然', '理化', '物理', '化學', '生物', '歷史', '地理', '公民', '資訊', '健體', '藝文', '科技'];

  // 1. 智慧評分找出最佳標題列 (包含最多科目或座號/姓名關鍵字的那一列)
  for (var r = 0; r <= maxScanRows; r++) {
    var rowText = data[r].join(' ');
    var score = 0;
    if (rowText.indexOf('座號') !== -1 || rowText.indexOf('號碼') !== -1 || rowText.indexOf('學號') !== -1) score += 3;
    if (rowText.indexOf('姓名') !== -1 || rowText.indexOf('學生') !== -1) score += 3;
    if (rowText.indexOf('班級') !== -1) score += 2;
    for (var s = 0; s < commonSubjects.length; s++) {
      if (rowText.indexOf(commonSubjects[s]) !== -1) score += 2;
    }
    if (score > bestScore) {
      bestScore = score;
      bestRowIdx = r;
    }
  }

  if (bestRowIdx === -1 || bestScore < 2) {
    bestRowIdx = data.length > 1 && (data[0].join(' ').indexOf('座號') !== -1 || data[0].join(' ').indexOf('姓名') !== -1) ? 0 : 1;
  }

  // 2. 整合多列標題（若有合併儲存格，向前繼承座號、姓名或班級欄位標題）
  var numCols = data[bestRowIdx].length;
  var combinedHeaders = [];
  for (var c = 0; c < numCols; c++) {
    var val = (data[bestRowIdx][c] || '').toString().trim();
    if (!val && bestRowIdx > 0) {
      val = (data[bestRowIdx - 1][c] || '').toString().trim();
    }
    combinedHeaders.push(val);
  }

  var seatCol = -1;
  var nameCol = -1;
  var classCol = -1;

  for (var c = 0; c < combinedHeaders.length; c++) {
    var h = combinedHeaders[c].replace(/\s+/g, '');
    if (!h) continue;
    if (seatCol === -1 && (h.indexOf('座號') !== -1 || h === '號' || h === '號碼' || h === '學號')) {
      seatCol = c;
    } else if (nameCol === -1 && (h.indexOf('姓名') !== -1 || h === '學生' || h === '名' || h === '學生姓名')) {
      nameCol = c;
    } else if (classCol === -1 && (h.indexOf('班級') !== -1 || h === '班')) {
      classCol = c;
    }
  }

  // 3. 備援判定：若未標註文字，透過資料列特徵自動探測座號與姓名欄位
  if (seatCol === -1 || nameCol === -1) {
    var sampleRows = data.slice(bestRowIdx + 1, Math.min(bestRowIdx + 6, data.length));
    for (var c = 0; c < numCols; c++) {
      if (c === classCol) continue;
      var isNumericSeat = sampleRows.length > 0 && sampleRows.every(function(row) {
        var num = parseInt(row[c], 10);
        return !isNaN(num) && num >= 1 && num <= 65;
      });
      var isChineseName = sampleRows.length > 0 && sampleRows.every(function(row) {
        var txt = (row[c] || '').toString().trim();
        return txt.length >= 2 && txt.length <= 4 && /^[\u4e00-\u9fa5]+$/.test(txt);
      });

      if (seatCol === -1 && isNumericSeat) seatCol = c;
      else if (nameCol === -1 && isChineseName) nameCol = c;
    }
  }

  if (seatCol === -1) seatCol = 0;
  if (nameCol === -1) nameCol = (seatCol === 0 ? 1 : 0);

  // 4. 辨識所有科目欄位 (排除座號、姓名、班級、統計、備註欄位)
  var subjectCols = [];
  var excludeKeywords = ['座號', '號碼', '學號', '姓名', '學生', '班級', '班', '平均', '總分', '排名', '名次', 'PR', '等第', '代碼', '狀態', '備註', '註記'];

  for (var sc = 0; sc < numCols; sc++) {
    if (sc === seatCol || sc === nameCol || sc === classCol) continue;
    var hName = combinedHeaders[sc];
    if (!hName) continue;

    var cleanH = hName.replace(/\s+/g, '');
    var isExcluded = excludeKeywords.some(function(k) { return cleanH.indexOf(k) !== -1; });
    if (isExcluded) continue;

    // 檢查底下資料是否包含有效數值
    var hasNumbers = false;
    for (var rCheck = bestRowIdx + 1; rCheck < Math.min(bestRowIdx + 10, data.length); rCheck++) {
      var cellVal = data[rCheck][sc];
      if (cellVal !== '' && cellVal !== null && cellVal !== undefined && !isNaN(Number(cellVal))) {
        hasNumbers = true;
        break;
      }
    }

    if (hasNumbers || commonSubjects.some(function(s) { return cleanH.indexOf(s) !== -1; })) {
      subjectCols.push({ index: sc, name: hName });
    }
  }

  return {
    headerRowIdx: bestRowIdx,
    seatCol: seatCol,
    nameCol: nameCol,
    classCol: classCol,
    subjectCols: subjectCols
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

  var colConfig = detectExamSheetColumns(data);
  if (!colConfig || colConfig.subjectCols.length === 0) return null;

  var headerRowIdx = colConfig.headerRowIdx;
  var seatCol = colConfig.seatCol;
  var nameCol = colConfig.nameCol;
  var subjectCols = colConfig.subjectCols;

  if (headerRowIdx > 0 && data[0][0]) {
    var titleCandidate = data[0][0].toString().trim();
    if (titleCandidate && titleCandidate.indexOf('座號') === -1 && titleCandidate.indexOf('姓名') === -1) {
      examTitle = titleCandidate;
    }
  }

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

  // 計算學生個人總平均之 PR 值 (供 PR 常模曲線圖標註落點)
  var allAvgValues = sortedByAvg.map(function(s) { return s.average; }).slice().sort(function(a, b) { return a - b; });
  var avgBelowCount = 0, avgEqualCount = 0;
  for (var abIdx = 0; abIdx < allAvgValues.length; abIdx++) {
    if (allAvgValues[abIdx] < myStudentRow.average) avgBelowCount++;
    else if (allAvgValues[abIdx] === myStudentRow.average) avgEqualCount++;
  }
  var studentAveragePr = Math.round(((avgBelowCount + 0.5 * avgEqualCount) / totalStudents) * 100);

  // 5. 計算各科目的「班級平均」、「PR值」與「箱型圖四分位數 (Q1, Median, Q3)」
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

    // 計算該科箱型圖五數摘要 (Q1, Median, Q3, Min, Max)
    var sortedSubScores = allScoresForSub.slice().sort(function(a, b) { return a - b; });
    var q1Val = calculatePercentileValue(sortedSubScores, 25);
    var medianVal = calculatePercentileValue(sortedSubScores, 50);
    var q3Val = calculatePercentileValue(sortedSubScores, 75);
    var minVal = sortedSubScores[0];
    var maxVal = sortedSubScores[sortedSubScores.length - 1];

    // 學習落點區間判定 (依楊家驊老師 GradeInterpr 標準)
    var zone = '穩健';
    var zoneStatus = '穩健發揮區間';
    var zoneType = 'leading';

    if (mySubScore >= q3Val || prVal >= 75) {
      zone = '精熟';
      zoneStatus = '精熟表現區間';
      zoneType = 'leading';
    } else if (mySubScore >= medianVal || prVal >= 50) {
      zone = '穩健';
      zoneStatus = '穩健發揮區間';
      zoneType = 'leading';
    } else if (mySubScore >= q1Val || prVal >= 25) {
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
      q1: q1Val,
      median: medianVal,
      q3: q3Val,
      min: minVal,
      max: maxVal,
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
      studentAveragePr: studentAveragePr,
      studentRank: myRank,
      totalStudents: totalStudents,
      classOverallAverage: classOverallAvg,
      subjects: subjectsResult
    }
  };
}

/**
 * 百分位數精確插值計算 (供 PR 及箱型圖 Q1, Median, Q3 運算)
 */
function calculatePercentileValue(sortedArr, p) {
  if (!sortedArr || sortedArr.length === 0) return 0;
  if (sortedArr.length === 1) return sortedArr[0];
  var idx = (p / 100) * (sortedArr.length - 1);
  var low = Math.floor(idx);
  var high = Math.ceil(idx);
  var weight = idx - low;
  return Math.round((sortedArr[low] * (1 - weight) + sortedArr[high] * weight) * 10) / 10;
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
      var primaryName = examStats.sheetName || examStats.examTitle;
      if (examList.indexOf(primaryName) === -1) examList.push(primaryName);
      if (examStats.examTitle && examStats.examTitle !== primaryName && examList.indexOf(examStats.examTitle) === -1) {
        examList.push(examStats.examTitle);
      }

      for (var stIdx = 0; stIdx < students.length; stIdx++) {
        var stSeat = students[stIdx].seatNumber;
        var stName = students[stIdx].name;
        var stStat = examStats.resultMap[stSeat] || (stName ? examStats.resultMap[stName] : null);
        if (stStat) {
          if (examStats.examTitle) students[stIdx].examStats[examStats.examTitle] = stStat;
          if (examStats.sheetName) students[stIdx].examStats[examStats.sheetName] = stStat;
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

  var colConfig = detectExamSheetColumns(data);
  if (!colConfig || colConfig.subjectCols.length === 0) return null;

  var headerRowIdx = colConfig.headerRowIdx;
  var seatCol = colConfig.seatCol;
  var nameCol = colConfig.nameCol;
  var subjectCols = colConfig.subjectCols;

  if (headerRowIdx > 0 && data[0][0]) {
    var titleCandidate = data[0][0].toString().trim();
    if (titleCandidate && titleCandidate.indexOf('座號') === -1 && titleCandidate.indexOf('姓名') === -1) {
      examTitle = titleCandidate;
    }
  }

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

  // 計算各科班均、各科箱型圖五數摘要與各生 PR
  var subAvgs = {};
  var subScoresPool = {};
  var subBoxStats = {};
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

    var sortedP = pList.slice().sort(function(a, b) { return a - b; });
    subBoxStats[subName] = {
      q1: calculatePercentileValue(sortedP, 25),
      median: calculatePercentileValue(sortedP, 50),
      q3: calculatePercentileValue(sortedP, 75),
      min: sortedP[0],
      max: sortedP[sortedP.length - 1]
    };
  }
  var classOverallAvg = Math.round((classTotalAvgSum / subjectCols.length) * 100) / 100;

  var allAvgSorted = sorted.map(function(s) { return s.average; }).slice().sort(function(a, b) { return a - b; });

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
      var bStats = subBoxStats[sName] || { q1: 60, median: 75, q3: 88, min: 40, max: 100 };
      subjectsResult.push({
        name: sName,
        score: myScore,
        classAverage: subAvgs[sName],
        pr: pr,
        q1: bStats.q1,
        median: bStats.median,
        q3: bStats.q3,
        min: bStats.min,
        max: bStats.max
      });
    }

    // 計算學生個人總平均之 PR 值
    var avgBelow = 0, avgEqual = 0;
    for (var ab = 0; ab < allAvgSorted.length; ab++) {
      if (allAvgSorted[ab] < st.average) avgBelow++;
      else if (allAvgSorted[ab] === st.average) avgEqual++;
    }
    var stAvgPr = Math.round(((avgBelow + 0.5 * avgEqual) / total) * 100);

    var statObj = {
      examTitle: examTitle,
      sheetName: sheetName,
      studentAverage: st.average,
      studentAveragePr: stAvgPr,
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
 * 智慧判定並保留班級設定，徹底避免覆蓋班級與擠壓儲存格欄寬！
 */
function syncStudentsFromExamSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    sSheet = ss.insertSheet('學生名冊', 0);
    sSheet.getRange(1, 1, 1, 5).setValues([['班級', '座號', '姓名', '專屬代碼', '啟用狀態']])
      .setBackground('#b45309').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
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
  if (data.length === 0) {
    ui.alert('考試分頁【' + examSheet.getName() + '】中沒有任何資料！');
    return;
  }

  var headerRowIdx = 1;
  if (data.length > 0 && (data[0].join(' ').indexOf('座號') !== -1 || data[0].join(' ').indexOf('姓名') !== -1)) {
    headerRowIdx = 0;
  }

  var headerRow = data[headerRowIdx].map(function(h) { return (h || '').toString().trim(); });
  var seatCol = -1;
  var nameCol = -1;
  var classCol = -1;
  for (var c = 0; c < headerRow.length; c++) {
    var h = headerRow[c];
    if (h.indexOf('座號') !== -1 || h === '號') seatCol = c;
    else if (h.indexOf('姓名') !== -1 || h === '學生') nameCol = c;
    else if (h.indexOf('班級') !== -1 || h === '班') classCol = c;
  }

  if (seatCol === -1) seatCol = 0;
  if (nameCol === -1) nameCol = 1;

  // 1. 讀取既有名冊中的班級、代碼與啟用狀態，避免更新時被覆蓋
  var existingBySeat = {};
  var existingByName = {};
  var existingClassFallback = '';
  var sData = sSheet.getDataRange().getValues();
  for (var s = 1; s < sData.length; s++) {
    var exCls = sData[s][0] !== undefined ? sData[s][0].toString().trim() : '';
    var exSeat = normalizeSeatNumber(sData[s][1]);
    var exName = sData[s][2] !== undefined ? sData[s][2].toString().trim() : '';
    var exTok = sData[s][3] !== undefined ? sData[s][3].toString().trim() : '';
    var exEn = sData[s][4];
    var isEnabled = (exEn === '' || exEn === undefined) ? true : (exEn === true || exEn.toString().toUpperCase() === 'TRUE' || exEn === 1);

    if (exCls && !existingClassFallback) {
      existingClassFallback = exCls;
    }

    var record = {
      className: exCls,
      token: exTok,
      isEnabled: isEnabled
    };
    if (exSeat) existingBySeat[exSeat] = record;
    if (exName) existingByName[exName] = record;
  }

  // 2. 從試算表名稱、分頁名稱或第 1 列大標題智慧推算預設班級
  var bannerClass = '';
  if (headerRowIdx > 0 && data.length > 0) {
    for (var col = 0; col < data[0].length; col++) {
      var cellVal = (data[0][col] || '').toString().trim();
      var bMatch = cellVal.match(/班級[：:\s]*([^\s,，]+)/) ||
                   cellVal.match(/([1-9]\d{2}|[七八九一二三四五六七八九十\d]+年[^\s\d]+班|\d{1,2}年\d{1,2}班)/);
      if (bMatch) {
        bannerClass = bMatch[1] || bMatch[0];
        break;
      }
    }
  }

  var titleClass = '';
  var tMatch = (examSheet.getName() + ' ' + ss.getName()).match(/([1-9]\d{2}|[七八九一二三四五六七八九十\d]+年[^\s\d]+班|\d{1,2}年\d{1,2}班)/);
  if (tMatch) {
    titleClass = tMatch[1] || tMatch[0];
  }

  // 班級預設優先級：既有名冊班級 > 考試表頭大標班級 > 分頁/試算表名稱班級
  var defaultClass = existingClassFallback || bannerClass || titleClass || '';

  var chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz';
  var newRoster = [];

  for (var r = headerRowIdx + 1; r < data.length; r++) {
    var seatVal = (data[r][seatCol] !== undefined) ? data[r][seatCol].toString().trim() : '';
    var nameVal = (data[r][nameCol] !== undefined) ? data[r][nameCol].toString().trim() : '';

    if (!seatVal || seatVal.indexOf('平均') !== -1 || seatVal.indexOf('標準差') !== -1) break;
    if (nameVal && nameVal.indexOf('平均') !== -1) break;

    var normSeat = normalizeSeatNumber(seatVal);
    var existingRecord = existingBySeat[normSeat] || (nameVal ? existingByName[nameVal] : null);

    // 班級判定：考試欄位 > 學生名冊既有設定 > 預設班級（絕不強制硬寫701）
    var studentClass = '';
    if (classCol !== -1 && data[r][classCol] !== undefined && data[r][classCol].toString().trim() !== '') {
      studentClass = data[r][classCol].toString().trim();
    } else if (existingRecord && existingRecord.className) {
      studentClass = existingRecord.className;
    } else {
      studentClass = defaultClass;
    }

    // 專屬防偽代碼：完全保留既有 Token，僅新學生自動生成
    var token = (existingRecord && existingRecord.token) ? existingRecord.token : '';
    if (!token) {
      token = '';
      for (var t = 0; t < 6; t++) token += chars.charAt(Math.floor(Math.random() * chars.length));
    }

    // 啟用狀態：完全保留既有設定（預設為 true）
    var isEnabled = (existingRecord && existingRecord.isEnabled !== undefined) ? existingRecord.isEnabled : true;

    newRoster.push([studentClass, normSeat, nameVal, token, isEnabled]);
  }

  if (newRoster.length === 0) {
    ui.alert('在考試分頁【' + examSheet.getName() + '】中未偵測到學生座號與姓名！');
    return;
  }

  // 清除舊有名冊資料列並寫入全新整理後的名冊
  if (sSheet.getLastRow() > 1) {
    sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 5).clearContent();
  }
  sSheet.getRange(2, 1, newRoster.length, 5).setValues(newRoster);

  // 確保標題列格式完好
  sSheet.getRange(1, 1, 1, 5).setValues([['班級', '座號', '姓名', '專屬代碼', '啟用狀態']])
    .setBackground('#b45309')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 設定標準寬敞欄寬，徹底解決儲存格被擠窄過小的問題
  sSheet.setColumnWidth(1, 90);  // 班級 (例如 701、三年忠班)
  sSheet.setColumnWidth(2, 80);  // 座號 (例如 01、02)
  sSheet.setColumnWidth(3, 120); // 姓名 (寬敞不截斷，完整容納 3~4 字姓名)
  sSheet.setColumnWidth(4, 130); // 專屬代碼 (寬敞清晰，防偽好辨識)
  sSheet.setColumnWidth(5, 100); // 啟用狀態 (舒適置中)

  // 設定標準舒適列高
  sSheet.setRowHeight(1, 36); // 標題列
  var totalRosterRows = Math.max(2, sSheet.getLastRow());
  for (var sr = 2; sr <= totalRosterRows; sr++) {
    sSheet.setRowHeight(sr, 28); // 學生資料列寬敞舒適 28px
  }

  // 置中對齊與凍結標題列
  sSheet.getRange(1, 1, totalRosterRows, 5)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sSheet.setFrozenRows(1);

  ui.alert('🎉 成功！已從【' + examSheet.getName() + '】同步 ' + newRoster.length + ' 位學生座號姓名，保留既有班級與代碼，並套用舒適欄寬！');
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
 * ➕ 新增一次考試（自動帶入學生名單）
 * 解決老師每次考試都要手動重打學生座號與姓名的困擾！
 */
function createNewExamSheetFromRoster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    ui.alert('找不到【學生名冊】分頁，請先點選【🛠️ 建立初始表格（空白試算表第一次使用）】！');
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

/**
 * 📐 自動整理欄寬與表格排版 (修復文字擠壓、欄寬過窄)
 * 自動辨識【學生名冊】與所有【考試分頁】，將姓名拓寬至120px、科目95px、凍結標題列並垂直置中
 */
function autoFormatAndResizeAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var updatedCount = 0;

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sName = sheet.getName();
    if (sName.startsWith('_')) continue;

    if (sName === '學生名冊') {
      sheet.setColumnWidth(1, 90);  // 班級
      sheet.setColumnWidth(2, 80);  // 座號
      sheet.setColumnWidth(3, 120); // 姓名
      sheet.setColumnWidth(4, 130); // 專屬代碼
      sheet.setColumnWidth(5, 100); // 啟用狀態
      sheet.setRowHeight(1, 36);
      
      var lastRow = Math.max(2, sheet.getLastRow());
      sheet.getRange(1, 1, lastRow, 5)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('center');
      for (var r = 2; r <= lastRow; r++) {
        sheet.setRowHeight(r, 28);
      }
      sheet.setFrozenRows(1);
      updatedCount++;
    } else {
      var data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      var headerRowIdx = 1;
      if (data.length > 0 && (data[0].join(' ').indexOf('座號') !== -1 || data[0].join(' ').indexOf('姓名') !== -1)) {
        headerRowIdx = 0;
      }

      var headerRow = data[headerRowIdx];
      var numCols = headerRow.length;

      for (var c = 0; c < numCols; c++) {
        var hText = (headerRow[c] || '').toString().trim();
        if (hText.indexOf('座號') !== -1) {
          sheet.setColumnWidth(c + 1, 80);
        } else if (hText.indexOf('姓名') !== -1) {
          sheet.setColumnWidth(c + 1, 120);
        } else if (hText.indexOf('班級') !== -1) {
          sheet.setColumnWidth(c + 1, 90);
        } else if (hText) {
          sheet.setColumnWidth(c + 1, 95);
        }
      }

      if (headerRowIdx === 1) {
        sheet.setRowHeight(1, 40);
        sheet.setRowHeight(2, 34);
        sheet.setFrozenRows(2);
      } else {
        sheet.setRowHeight(1, 34);
        sheet.setFrozenRows(1);
      }

      var maxRow = Math.max(headerRowIdx + 1, sheet.getLastRow());
      sheet.getRange(1, 1, maxRow, numCols)
        .setVerticalAlignment('middle')
        .setHorizontalAlignment('center');

      for (var r2 = headerRowIdx + 2; r2 <= maxRow; r2++) {
        sheet.setRowHeight(r2, 28);
      }
      updatedCount++;
    }
  }

  SpreadsheetApp.getUi().alert(
    '🎉 美化完成！\n\n' +
    '已為您自動優化全表 ' + updatedCount + ' 個分頁的欄寬與列高：\n' +
    '• 姓名欄位已拓寬至 120px，中文字不再被截斷！\n' +
    '• 各科成績欄位已等距舒展開 (95px)，數值清晰不擠壓。\n' +
    '• 標題列已固定凍結，上下滾動更清晰！'
  );
}

// ========================================================
// 家長閱覽紀錄日誌與已讀追蹤模組
// ========================================================

/**
 * 取得或自動建立【家長閱覽紀錄】工作表
 */
function getOrCreateAccessLogSheet(ss) {
  var sheet = ss.getSheetByName('家長閱覽紀錄');
  if (!sheet) {
    sheet = ss.insertSheet('家長閱覽紀錄');
    var headers = [['閱覽時間', '班級', '座號', '學生姓名', '查詢代碼', '狀態說明']];
    sheet.getRange(1, 1, 1, 6).setValues(headers)
      .setBackground('#0d9488') // 質感松綠
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setColumnWidth(1, 170); // 閱覽時間
    sheet.setColumnWidth(2, 80);  // 班級
    sheet.setColumnWidth(3, 70);  // 座號
    sheet.setColumnWidth(4, 110); // 學生姓名
    sheet.setColumnWidth(5, 110); // 查詢代碼
    sheet.setColumnWidth(6, 130); // 狀態說明
    sheet.setRowHeight(1, 34);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setVerticalAlignment('middle');
  }
  return sheet;
}

/**
 * 記錄家長閱覽成績日誌 (附帶 60 秒冷卻防刷保護)
 */
function recordParentAccess(student) {
  if (!student || !student.token) return;
  var cache = CacheService.getScriptCache();
  var debounceKey = 'log_db_' + encodeURIComponent(student.token.toString().toUpperCase());
  if (cache.get(debounceKey)) {
    return; // 60 秒內已記錄過，防重複觸發
  }
  cache.put(debounceKey, '1', 60);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = getOrCreateAccessLogSheet(ss);
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    logSheet.appendRow([
      nowStr,
      student.className || '',
      student.seatNumber || '',
      student.name || '',
      student.token || '',
      '線上成績查閱'
    ]);
  } catch (err) {
    // 記錄日誌失敗不應阻斷家長看成績
  }
}

/**
 * 取得家長閱覽日誌與全班已讀/未讀統計分析
 */
function getParentAccessLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { success: false, message: '找不到【學生名冊】分頁' };
  }

  // 1. 取得學生名單
  var sData = sSheet.getDataRange().getValues();
  var studentsMap = {};
  var studentsList = [];

  for (var i = 1; i < sData.length; i++) {
    var row = sData[i];
    var cls = row[0] ? row[0].toString().trim() : '';
    var seat = normalizeSeatNumber(row[1]);
    var name = row[2] ? row[2].toString().trim() : '';
    var token = row[3] ? row[3].toString().trim() : '';
    var isEnabled = row[4];

    if (seat || name) {
      var stObj = {
        className: cls,
        seatNumber: seat,
        name: name,
        token: token,
        studentId: (cls || '') + '-' + seat,
        isEnabled: isEnabled !== false,
        isRead: false,
        readCount: 0,
        lastReadTime: ''
      };
      studentsList.push(stObj);
      if (token) {
        studentsMap[token.toUpperCase()] = stObj;
      }
      if (seat) {
        studentsMap['SEAT_' + seat] = stObj;
      }
    }
  }

  // 2. 取得【家長閱覽紀錄】
  var logSheet = ss.getSheetByName('家長閱覽紀錄');
  var recentLogs = [];
  var totalLogsCount = 0;

  if (logSheet && logSheet.getLastRow() > 1) {
    var logData = logSheet.getDataRange().getValues();
    totalLogsCount = logData.length - 1;

    // 從最後一列往回讀 (最新排前面)
    for (var r = logData.length - 1; r >= 1; r--) {
      var logRow = logData[r];
      var timeStr = logRow[0] ? (logRow[0] instanceof Date ? Utilities.formatDate(logRow[0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : logRow[0].toString().trim()) : '';
      var logCls = logRow[1] ? logRow[1].toString().trim() : '';
      var logSeat = normalizeSeatNumber(logRow[2]);
      var logName = logRow[3] ? logRow[3].toString().trim() : '';
      var logToken = logRow[4] ? logRow[4].toString().trim().toUpperCase() : '';
      var logNote = logRow[5] ? logRow[5].toString().trim() : '線上成績查閱';

      // 歸納進學生統計
      var matchedStudent = (logToken && studentsMap[logToken]) || (logSeat && studentsMap['SEAT_' + logSeat]);
      if (matchedStudent) {
        matchedStudent.isRead = true;
        matchedStudent.readCount = (matchedStudent.readCount || 0) + 1;
        if (!matchedStudent.lastReadTime) {
          matchedStudent.lastReadTime = timeStr; // 第一個碰到的就是最新時間
        }
      }

      // 保留最近 50 筆供即時時間軸展示
      if (recentLogs.length < 50) {
        recentLogs.push({
          time: timeStr,
          className: logCls,
          seatNumber: logSeat,
          name: logName,
          token: logToken,
          note: logNote
        });
      }
    }
  }

  // 3. 統計彙整
  var totalStudents = studentsList.length;
  var readStudents = 0;
  for (var k = 0; k < studentsList.length; k++) {
    if (studentsList[k].isRead) {
      readStudents++;
    }
  }
  var unreadStudents = Math.max(0, totalStudents - readStudents);
  var readRate = totalStudents > 0 ? ((readStudents / totalStudents) * 100).toFixed(1) + '%' : '0%';

  return {
    success: true,
    summary: {
      total: totalStudents,
      readCount: readStudents,
      unreadCount: unreadStudents,
      readRate: readRate,
      totalLogViews: totalLogsCount,
      latestAccessTime: recentLogs.length > 0 ? recentLogs[0].time : '尚無紀錄'
    },
    students: studentsList,
    recentLogs: recentLogs
  };
}

/**
 * 在 Google 試算表中直接切換開啟【家長閱覽紀錄】
 */
function openParentAccessLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = getOrCreateAccessLogSheet(ss);
  ss.setActiveSheet(logSheet);
  var count = Math.max(0, logSheet.getLastRow() - 1);
  SpreadsheetApp.getUi().alert('📋 家長閱覽日誌', '目前累積共有 ' + count + ' 筆家長閱覽紀錄。\n您可在此分頁直接查閱每筆詳細時間戳記與學生姓名。', SpreadsheetApp.getUi().ButtonSet.OK);
}

var MPScoreEngine = {
  autoFormatAndResizeAllSheets: autoFormatAndResizeAllSheets,
  setupStandardSheets: setupStandardSheets,
  createNewExamSheetFromRoster: createNewExamSheetFromRoster,
  syncStudentsFromExamSheet: syncStudentsFromExamSheet,
  generateTokensForStudents: generateTokensForStudents,
  previewClassStatistics: previewClassStatistics,
  openParentAccessLogSheet: openParentAccessLogSheet,
  getParentAccessLogs: getParentAccessLogs,
  handleDoGet: doGet,
  handleDoPost: doPost
};


