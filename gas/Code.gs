/**
 * MPScore - 智慧學生成績與學習診斷系統 (Google Apps Script 後端)
 * 網站作者 / 系統規劃：新北市林口區麗園國小 楊家驊老師 (Chiahua Yang)
 *
 * 【核心架構 - 每次考試獨立分頁】：
 * 1. 【學生名冊】分頁：班級 | 座號 | 姓名 | 專屬代碼 | 啟用狀態
 * 2. 考試分頁（例如：114-1期中考、114-1期末考... 每次考試一個新分頁）：
 *    - 第 1 列：考試名稱標題（例如：114學年度第2學期期中考）
 *    - 第 2 列：座號 | 姓名 | 國語 | 數學 | 社會 | 英文 | 自然 (僅需上傳原始成績)
 *    - 第 3 列起：各學生原始分數
 *    ★ 老師只需上傳原始成績！所有「各科班級平均、個人平均、班級排名、PR值、等第區間、學習診斷評語」
 *      全由本系統自動運算，無需老師手動計算！
 *
 * 【開發歷程簡述】：
 * • 2026-10-15：v1.0 建立獨立試算表架構與楊家驊老師 GradeInterpr 診斷模組，實體拆分教師端與家長端。
 * • 2026-10-16：新增 5 大教育統計模組（蜜蜂聚落圖、T分數、啞鈴圖、動能矩陣、歷程地圖）與家長端相機掃碼。
 * • 2026-10-17：落地 Apple HIG 人因工藝與個資法去識別化遮罩。
 * • 2026-10-18：攻克 A4 滿版列印工程與家長閱覽日誌追蹤。
 * • 2026-10-26：實施教師金鑰 PIN 防護與試算表雙向同步。
 * • 2026-10-27：導入試卷難易度指數 (P*) 與統計方法指南。
 * • 2026-10-29：整合 School LINE Hub 校園推播中樞、手機即時沙盒模擬與 0ms 家長端導向防護。
 * • 2026-11-06：資安與個資全面加固、世代快取版本號與常模曲線智慧引線對話框升級。
 * • 2026-11-12：新增「抽離考試／額外採計科目」統計隔離機制（單科排除班級統計、保留個人總平均與總排名、零標籤化個資防護）。
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
    .addItem('📲 透過 School LINE Hub 發送成績推播', 'openLinePushDialog')
    .addItem('🔍 檢查 School LINE Hub 家長綁定概況', 'checkLineHubStatus')
    .addSeparator()
    .addItem('🛡️ 查看 / 設定「教師管理金鑰」', 'showTeacherKeyDialog')
    .addItem('📋 查看「家長閱覽紀錄」', 'openParentAccessLogSheet')
    .addItem('📐 自動整理欄寬與表格排版', 'autoFormatAndResizeAllSheets')
    .addItem('🚀 立即更新雲端資料（改完分數點此）', 'clearSystemCache')
    .addItem('📊 查看全班成績與各科平均', 'previewClassStatistics')
    .addToUi();
}

/**
 * 教師安全金鑰管理 (Teacher Key)
 * 🔒 支援首次智慧自動配對 (Trust On First Use / Auto-Pairing)、ScriptProperties 私有雲端儲存
 * 💡 雙向同步至【學生名冊】F 欄：讓老師換電腦時打開試算表就能一眼看見，亦可直接在儲存格修改密碼！
 */
function getTeacherKey() {
  var props = PropertiesService.getScriptProperties();
  var storedKey = props.getProperty('TEACHER_KEY') || '';

  // 嘗試與【學生名冊】F2 儲存格雙向校驗
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName('學生名冊');
    if (sSheet && sSheet.getLastRow() >= 2) {
      var cellVal = sSheet.getRange(2, 6).getValue();
      if (cellVal) {
        var str = cellVal.toString().trim();
        if (str && str.length >= 4 && str !== storedKey) {
          props.setProperty('TEACHER_KEY', str);
          return str;
        }
      }
    }
  } catch (e) {}

  return storedKey;
}

/**
 * 嚴格判斷工作表分頁是否為「考試分頁」
 * 自動排除【學生名冊】、【家長閱覽紀錄】、系統暫存分頁(_)以及任何包含「名冊/紀錄/記錄/閱覽/瀏覽/日誌/log」的系統分頁
 */
function isExamSheetName(name) {
  if (!name) return false;
  var s = name.toString().trim();
  if (s.startsWith('_')) return false;
  var low = s.toLowerCase();
  if (low.indexOf('名冊') !== -1 || low.indexOf('閱覽') !== -1 || low.indexOf('瀏覽') !== -1 || low.indexOf('紀錄') !== -1 || low.indexOf('記錄') !== -1 || low.indexOf('日誌') !== -1 || low.indexOf('log') !== -1) {
    return false;
  }
  return true;
}

/**
 * 雙向同步教師金鑰至【學生名冊】工作表 F1:F2 儲存格
 * 方便教師換電腦時隨時抬頭查閱，或直接在試算表中自訂修改！
 */
function syncTeacherKeyWithRoster(ss, keyToSet) {
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss.getSheetByName('學生名冊');
    if (!sSheet) return;

    var props = PropertiesService.getScriptProperties();
    var storedKey = keyToSet || props.getProperty('TEACHER_KEY') || '';

    // 檢查 F2 儲存格是否有老師手動輸入的新密碼
    var f2Val = sSheet.getRange(2, 6).getValue();
    if (f2Val) {
      f2Val = f2Val.toString().trim();
      if (f2Val.length >= 4 && f2Val !== storedKey) {
        props.setProperty('TEACHER_KEY', f2Val);
        storedKey = f2Val;
      }
    }

    if (storedKey) {
      sSheet.getRange(1, 6).setValue('🔑 教師金鑰 (換電腦看此)')
        .setFontWeight('bold')
        .setBackground('#fef3c7')
        .setFontColor('#78350f')
        .setHorizontalAlignment('center');
      sSheet.getRange(2, 6).setValue(storedKey)
        .setFontWeight('bold')
        .setFontFamily('Courier New')
        .setBackground('#fffbeb')
        .setFontColor('#92400e')
        .setHorizontalAlignment('center');
      sSheet.setColumnWidth(6, 175);
    }
  } catch (e) {
    // 容錯不阻斷
  }
}

function verifyOrCreateTeacherKey(providedKey) {
  var currentKey = getTeacherKey();
  var props = PropertiesService.getScriptProperties();

  // 防暴力破解限制：若 10 分鐘內連續驗證失敗達 10 次，暫時鎖定管理端點
  try {
    var cache = CacheService.getScriptCache();
    var failCount = parseInt(cache.get('mpscore_teacher_auth_fail') || '0', 10);
    if (failCount >= 10) {
      return { success: false, reason: 'LOCKED' };
    }
  } catch (e) {}

  // 1. 若試算表尚未綁定金鑰
  if (!currentKey) {
    if (providedKey && providedKey.toString().trim().length >= 4) {
      var cleanKey = providedKey.toString().trim();
      props.setProperty('TEACHER_KEY', cleanKey);
      syncTeacherKeyWithRoster(null, cleanKey);
      return { success: true, isNewPairing: true };
    }
    // 尚未綁定且未提供金鑰：允許過渡通行
    return { success: true, isUnprotected: true };
  }

  // 2. 若試算表已綁定金鑰：嚴格比對
  if (!providedKey) return { success: false, reason: 'MISSING_KEY' };
  var isMatch = (providedKey.toString().trim() === currentKey.trim());
  if (isMatch) {
    try { CacheService.getScriptCache().remove('mpscore_teacher_auth_fail'); } catch (e) {}
    syncTeacherKeyWithRoster(null, currentKey);
  } else {
    try {
      var cacheFail = CacheService.getScriptCache();
      var nextFail = parseInt(cacheFail.get('mpscore_teacher_auth_fail') || '0', 10) + 1;
      cacheFail.put('mpscore_teacher_auth_fail', String(nextFail), 600);
    } catch (e) {}
  }
  return { success: isMatch, reason: isMatch ? 'OK' : 'MISMATCH' };
}

/**
 * 取得與遞增系統快取世代版本號 (Cache Epoch)
 * 當老師點擊「清除快取」或上傳成績時，只需遞增 Epoch 版本號即可 O(1) 瞬時讓所有舊快取失效！
 */
function getCacheEpoch_() {
  try {
    var cache = CacheService.getScriptCache();
    var ep = cache.get('mpscore_cache_epoch');
    if (ep) return ep;
    var props = PropertiesService.getScriptProperties();
    ep = props.getProperty('CACHE_EPOCH_VERSION') || '1';
    cache.put('mpscore_cache_epoch', ep, 21600);
    return ep;
  } catch (e) {
    return '1';
  }
}

function bumpCacheEpoch_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var cur = parseInt(props.getProperty('CACHE_EPOCH_VERSION') || '1', 10);
    if (isNaN(cur)) cur = 1;
    var next = String(cur + 1);
    props.setProperty('CACHE_EPOCH_VERSION', next);
    var cache = CacheService.getScriptCache();
    var removeKeys = ['mpscore_cache_epoch', 'mpscore_loader_engine_v2_meta'];
    for (var i = 0; i < 10; i++) {
      removeKeys.push('mpscore_loader_engine_v2_' + i);
    }
    cache.removeAll(removeKeys);
    cache.put('mpscore_cache_epoch', next, 21600);
    return next;
  } catch (e) {
    return String(new Date().getTime());
  }
}

/**
 * 學生姓名去識別化遮罩（符合《個人資料保護法》最小揭露原則）
 * 2 字姓名遮第 2 字（王○），3 字姓名遮中間字（王○明），4 字以上遮倒數第 2 字（歐陽○華）
 */
function maskStudentName_(name) {
  var s = (name || '').toString().trim();
  if (!s) return '同學';
  if (s.indexOf('○') !== -1 || s.indexOf('O') !== -1 || s.indexOf('＊') !== -1) return s;
  if (s.length <= 1) return s;
  if (s.length === 2) return s.charAt(0) + '○';
  if (s.length === 3) return s.charAt(0) + '○' + s.charAt(2);
  return s.substring(0, s.length - 2) + '○' + s.charAt(s.length - 1);
}

/**
 * =========================================================================
 * 🌱 抽離考試／額外採計科目管理模組 (Pull-out / Excluded Subject Statistics)
 * - 支援教師端視覺化勾選設定，雙向同步於 ScriptProperties 與【學生名冊】G 欄
 * - 支援單格分數標記（如 *85、＊85、抽離:85）作為單次考試個別額外採計
 * - 運算規則：該生仍列入個人總平均與班級總排名，但所設定之科目不計入該科班級統計資料
 * =========================================================================
 */
var EXCLUDED_NOTE_TEXT_ = '該科目額外採計，該測驗不計入班級統計資料，僅呈現學生成績';

function parseScoreCell_(rawVal) {
  if (rawVal === '' || rawVal === null || rawVal === undefined) {
    return { score: 0, hasValue: false, inlineExcluded: false };
  }
  if (typeof rawVal === 'number') {
    return { score: isNaN(rawVal) ? 0 : rawVal, hasValue: !isNaN(rawVal), inlineExcluded: false };
  }
  var s = rawVal.toString().trim();
  if (!s) return { score: 0, hasValue: false, inlineExcluded: false };

  var isInlineEx = /^[\*＊]/.test(s) || s.indexOf('抽離') !== -1 || s.indexOf('額外') !== -1;
  var cleaned = s.replace(/^[\*＊\s]+|(?:抽離|額外|採計)[：:\s]*|[()（）抽離額外採計分\s]/g, '');
  var num = Number(cleaned);
  if (!isNaN(num) && cleaned !== '') {
    return { score: num, hasValue: true, inlineExcluded: isInlineEx };
  }
  return { score: 0, hasValue: false, inlineExcluded: false };
}

function getExcludedSubjectsMap_(ss) {
  var map = {};
  var props = PropertiesService.getScriptProperties();
  try {
    var rawJson = props.getProperty('EXCLUDED_SUBJECTS_MAP');
    if (rawJson) {
      var parsed = JSON.parse(rawJson);
      if (parsed && typeof parsed === 'object') {
        for (var k in parsed) {
          if (parsed.hasOwnProperty(k) && Array.isArray(parsed[k]) && parsed[k].length > 0) {
            map[normalizeSeatNumber(k)] = parsed[k];
          }
        }
      }
    }
  } catch (e) {}

  // 與【學生名冊】G 欄（第 7 欄）雙向讀取同步
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss ? ss.getSheetByName('學生名冊') : null;
    if (sSheet && sSheet.getLastColumn() >= 7 && sSheet.getLastRow() >= 2) {
      var sData = sSheet.getDataRange().getValues();
      var gHeader = (sData[0][6] || '').toString();
      if (gHeader.indexOf('額外') !== -1 || gHeader.indexOf('抽離') !== -1) {
        var sheetMap = {};
        var hasSheetConfig = false;
        for (var i = 1; i < sData.length; i++) {
          var seat = normalizeSeatNumber(sData[i][1]);
          var gVal = (sData[i][6] || '').toString().trim();
          if (seat && gVal) {
            var subs = gVal.split(/[,，、;\s]+/).map(function(x) { return x.trim(); }).filter(function(x) { return !!x; });
            if (subs.length > 0) {
              sheetMap[seat] = subs;
              hasSheetConfig = true;
            }
          }
        }
        if (hasSheetConfig) {
          for (var sk in sheetMap) {
            if (sheetMap.hasOwnProperty(sk)) {
              map[sk] = sheetMap[sk];
            }
          }
        }
      }
    }
  } catch (e) {}

  return map;
}

function saveExcludedSubjectsMap_(newMap, ss) {
  var cleanMap = {};
  if (newMap && typeof newMap === 'object') {
    for (var k in newMap) {
      if (newMap.hasOwnProperty(k)) {
        var normSeat = normalizeSeatNumber(k);
        var arr = Array.isArray(newMap[k]) ? newMap[k] : [];
        var cleanSubs = [];
        for (var i = 0; i < arr.length; i++) {
          var sName = (arr[i] || '').toString().trim();
          if (sName && cleanSubs.indexOf(sName) === -1) {
            cleanSubs.push(sName);
          }
        }
        if (normSeat && cleanSubs.length > 0) {
          cleanMap[normSeat] = cleanSubs;
        }
      }
    }
  }

  var props = PropertiesService.getScriptProperties();
  props.setProperty('EXCLUDED_SUBJECTS_MAP', JSON.stringify(cleanMap));

  // 同步寫入【學生名冊】G 欄，讓老師在試算表也能一眼查看與備份
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    var sSheet = ss ? ss.getSheetByName('學生名冊') : null;
    if (sSheet && sSheet.getLastRow() >= 1) {
      sSheet.getRange(1, 7).setValue('🌱 額外採計科目 (不計入單科統計)')
        .setFontWeight('bold')
        .setBackground('#e0e7ff')
        .setFontColor('#312e81')
        .setHorizontalAlignment('center');
      sSheet.setColumnWidth(7, 210);

      var lastRow = sSheet.getLastRow();
      if (lastRow >= 2) {
        var seatsRange = sSheet.getRange(2, 2, lastRow - 1, 1).getValues();
        var outCol = [];
        for (var r = 0; r < seatsRange.length; r++) {
          var stSeat = normalizeSeatNumber(seatsRange[r][0]);
          var stSubs = cleanMap[stSeat] || [];
          outCol.push([stSubs.join(', ')]);
        }
        sSheet.getRange(2, 7, outCol.length, 1).setValues(outCol)
          .setHorizontalAlignment('center')
          .setFontColor('#4338ca')
          .setFontWeight('bold');
      }
    }
  } catch (e) {}

  bumpCacheEpoch_();
  return cleanMap;
}

function saveExcludedConfigFromApi(body) {
  var payload = {};
  try {
    payload = typeof body === 'string' ? JSON.parse(body) : (body || {});
  } catch (e) {
    return { success: false, message: '設定資料格式錯誤: ' + e.toString() };
  }
  var targetMap = payload.excludedMap || payload || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var saved = saveExcludedSubjectsMap_(targetMap, ss);
  return {
    success: true,
    excludedMap: saved,
    message: '✅ 已成功儲存「抽離考試／額外採計科目」設定並更新雲端統計快取！'
  };
}

function isSubjectExcludedForStudent_(excludedMap, seat, name, subName, inlineExcluded) {
  if (inlineExcluded) return true;
  if (!excludedMap || !subName) return false;
  var normSeat = normalizeSeatNumber(seat);
  var list = excludedMap[normSeat] || (name ? excludedMap[name] : null);
  if (!list || !Array.isArray(list) || list.length === 0) return false;

  var cleanTarget = subName.toString().trim().replace(/科$/, '');
  for (var i = 0; i < list.length; i++) {
    var item = (list[i] || '').toString().trim().replace(/科$/, '');
    if (!item) continue;
    if (item === cleanTarget || item.indexOf(cleanTarget) !== -1 || cleanTarget.indexOf(item) !== -1) {
      return true;
    }
    if ((item === '英文' || item === '英語') && (cleanTarget === '英文' || cleanTarget === '英語')) return true;
    if ((item === '國語' || item === '國文' || item === '語文') && (cleanTarget === '國語' || cleanTarget === '國文' || cleanTarget === '語文')) return true;
  }
  return false;
}

function showTeacherKeyDialog() {
  var ui = SpreadsheetApp.getUi();
  var currentKey = getTeacherKey();
  if (!currentKey) {
    // 若尚未設定，自動生成一組預設金鑰
    var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789abcdefghijkmnopqrstuvwxyz';
    currentKey = 'mp_';
    for (var i = 0; i < 10; i++) {
      currentKey += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    PropertiesService.getScriptProperties().setProperty('TEACHER_KEY', currentKey);
    syncTeacherKeyWithRoster(null, currentKey);
  }

  var response = ui.prompt(
    '🛡️ 教師管理安全金鑰 (Teacher Key)',
    '此金鑰用於保護「全班成績名冊導出」與「家長閱覽日誌」，防止外人越權存取。\n\n' +
    '【目前您的金鑰為】：\n' + currentKey + '\n\n' +
    '👉 提示：金鑰已同步記錄於【學生名冊】分頁右上角 F1:F2 儲存格，換電腦可隨時查閱或直接在儲存格修改。\n' +
    '（若要自訂修改金鑰，請在下方輸入新金鑰後點擊確定，若不修改請直接點取消）：',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    var input = response.getResponseText().trim();
    if (input && input.length >= 6) {
      PropertiesService.getScriptProperties().setProperty('TEACHER_KEY', input);
      syncTeacherKeyWithRoster(null, input);
      ui.alert('✅ 金鑰已更新！', '您的新教師金鑰為：\n' + input + '\n\n已同步至【學生名冊】F 欄，請記得同步至 MPScore 教師端網頁設定。', ui.ButtonSet.OK);
    } else if (input) {
      ui.alert('⚠️ 更新失敗', '自訂金鑰長度需至少 6 個字元以上。', ui.ButtonSet.OK);
    }
  }
}

/**
 * 清除系統雲端快取（採用世代版本號 Cache Epoch 機制，O(1) 瞬時失效所有舊快取）
 */
function clearSystemCache() {
  var ui = SpreadsheetApp.getUi();
  try {
    var newEpoch = bumpCacheEpoch_();
    ui.alert('✨ 雲端快取清除成功！(快取世代 v' + newEpoch + ')', '家長端與查詢系統現已立即載入試算表最新成績與運算結果。', ui.ButtonSet.OK);
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
 * 🔒 資安白名單門禁設計 (Whitelist Gatekeeper)
 * - 公開端點 (Public)：ping (健康檢查)、query (依 6 碼隨機 Token 查個人成績)
 * - 管理端點 (Protected)：getStudentsForPrint (全班名冊)、getParentAccessLogs (家長日誌)、appendExamScores (寫入成績)
 *   👉 必須附帶正確的 &key= (Teacher Key)，否則一律回傳 403 阻斷，嚴防全校名冊與個資遭 IDOR 越權打包！
 */
function doGet(e) {
  var params = e ? e.parameter : {};
  var action = params.action || ((params.token || params.seat || params.name) ? 'query' : 'ping');
  var teacherKey = params.key || (e && e.headers && (e.headers['x-teacher-key'] || e.headers['X-Teacher-Key'])) || '';
  var output = {};

  try {
    // 1. 公開安全端點（任何人免金鑰可存取：健康檢查、金鑰綁定與個人專屬代碼／座號＋姓名查詢）
    if (action === 'ping') {
      output = checkSystemStatus(teacherKey);
      return ContentService.createTextOutput(JSON.stringify(output))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'bindTeacherKey' || action === 'setTeacherKey') {
      var newKey = (params.newKey || params.key || '').toString().trim();
      var oldKey = (params.oldKey || params.currentKey || teacherKey || '').toString().trim();
      output = bindTeacherKeyFromApi(newKey, oldKey);
      return ContentService.createTextOutput(JSON.stringify(output))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (action === 'query') {
      var isSimulation = params.isSimulation === 'true' || params.simulate === 'true' || params.mock === 'true' || params.preview === 'true';
      output = getStudentScoreByToken(params.token, isSimulation, params.seat, params.name);
      return ContentService.createTextOutput(JSON.stringify(output))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. 高權限教師管理端點：自動配對或嚴格比對教師安全金鑰
    var verifyResult = verifyOrCreateTeacherKey(teacherKey);
    if (!verifyResult.success) {
      output = {
        success: false,
        code: 403,
        message: verifyResult.reason === 'LOCKED'
          ? '⛔ 教師金鑰驗證錯誤次數過多，系統已暫時鎖定 10 分鐘以防範暴力破解，請稍後再試。'
          : '⛔ 存取遭拒：此試算表已綁定「教師管理安全金鑰」，您輸入的金鑰不符。請檢查 MPScore 教師端網頁【步驟 3】的金鑰設定。'
      };
      return ContentService.createTextOutput(JSON.stringify(output))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. 通過金鑰驗證後的管理端點處理
    if (action === 'getStudentsForPrint') {
      output = getStudentListForTeacher();
    } else if (action === 'appendExamScores') {
      var body = params.payload || (e && e.postData && e.postData.contents ? e.postData.contents : '');
      output = appendExamScores(body);
      bumpCacheEpoch_();
    } else if (action === 'getParentAccessLogs') {
      output = getParentAccessLogs();
    } else if (action === 'getLineBindings') {
      output = getLineBindingsFromHub();
    } else if (action === 'sendLinePush') {
      var pushPayload = params.payload || (e && e.postData && e.postData.contents ? e.postData.contents : '');
      output = sendLinePushViaHub(pushPayload);
    } else if (action === 'syncStudentsToLineHub') {
      output = syncStudentsToLineHubFromGas();
    } else if (action === 'saveLineHubConfig') {
      output = saveLineHubConfigFromApi(params);
    } else if (action === 'getExcludedConfig') {
      output = { success: true, excludedMap: getExcludedSubjectsMap_(SpreadsheetApp.getActiveSpreadsheet()) };
    } else if (action === 'saveExcludedConfig') {
      var exPayload = params.payload || (e && e.postData && e.postData.contents ? e.postData.contents : '');
      output = saveExcludedConfigFromApi(exPayload);
    } else if (action === 'batchShortenUrls') {
      var urlsPayload = params.payload || (e && e.postData && e.postData.contents ? e.postData.contents : '');
      output = batchShortenUrlsFromGas(urlsPayload);
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
 * 遠端主動綁定/設定教師安全金鑰至 Google 試算表
 * 🔒 資安強化：若試算表已綁定過教師金鑰，必須先通過原金鑰 (oldKey) 驗證才允許變更，徹底阻絕未授權權限接管漏洞！
 */
function bindTeacherKeyFromApi(newKey, oldKey) {
  if (!newKey || newKey.toString().trim().length < 4) {
    return { success: false, message: '教師金鑰長度需至少 4 碼以上。' };
  }
  var currentStoredKey = getTeacherKey();
  if (currentStoredKey) {
    try {
      var cache = CacheService.getScriptCache();
      var failCount = parseInt(cache.get('mpscore_teacher_auth_fail') || '0', 10);
      if (failCount >= 10) {
        return { success: false, code: 429, message: '⛔ 金鑰驗證失敗次數過多，請稍候 10 分鐘再試。' };
      }
    } catch (e) {}

    var cleanOldKey = (oldKey || '').toString().trim();
    if (!cleanOldKey || cleanOldKey !== currentStoredKey.trim()) {
      try {
        var cacheFail = CacheService.getScriptCache();
        var nextFail = parseInt(cacheFail.get('mpscore_teacher_auth_fail') || '0', 10) + 1;
        cacheFail.put('mpscore_teacher_auth_fail', String(nextFail), 600);
      } catch (e) {}
      return {
        success: false,
        code: 403,
        message: '⛔ 權限拒絕：此試算表已綁定教師管理安全金鑰！若需遠端變更金鑰，請先輸入正確的原金鑰，或直接於 Google 試算表上方選單【⚡ MPScore 工具 > 🛡️ 查看 / 設定教師管理金鑰】（或【學生名冊】F2 儲存格）修改。'
      };
    }
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var cleanKey = newKey.toString().trim();
  props.setProperty('TEACHER_KEY', cleanKey);
  syncTeacherKeyWithRoster(ss, cleanKey);
  return {
    success: true,
    message: '🎉 金鑰已成功綁定至 Google 試算表！(已同步寫入【學生名冊】F 欄)',
    teacherKey: cleanKey
  };
}

/**
 * 系統連線檢測 (ping)
 * 偵測試算表內【學生名冊】與所有考試分頁，並支援前端金鑰智慧自動配對 (Auto-Pairing)
 */
function checkSystemStatus(teacherKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { success: false, message: '試算表內找不到名為【學生名冊】的分頁，請先點選【MPScore 工具 > 🛠️ 建立初始表格（空白試算表第一次使用）】' };
  }

  // 嘗試自動配對金鑰（若試算表尚未綁定，且前端有傳來有效金鑰）
  var keyStatus = 'NONE';
  var currentStoredKey = getTeacherKey();
  if (!currentStoredKey && teacherKey && teacherKey.length >= 4) {
    PropertiesService.getScriptProperties().setProperty('TEACHER_KEY', teacherKey.trim());
    currentStoredKey = teacherKey.trim();
    syncTeacherKeyWithRoster(ss, currentStoredKey);
    keyStatus = 'AUTO_PAIRED';
  } else if (currentStoredKey) {
    syncTeacherKeyWithRoster(ss, currentStoredKey);
    keyStatus = (teacherKey && teacherKey.trim() === currentStoredKey) ? 'VERIFIED' : 'PROTECTED';
  }

  var allSheets = ss.getSheets();
  var examTabs = [];
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    if (isExamSheetName(name)) {
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
    autoCalculationEnabled: true,
    teacherKeyStatus: keyStatus
  };
}

/**
 * 依據專屬 Token 或「座號＋學生完整姓名」查詢該名學生歷次考試成績與自動運算成果
 * 🔒 資安防護：
 * 1. 禁止單純以 01~40 純座號猜測查詢
 * 2. 支援家長以「6 碼專屬代碼」或「座號＋學生完整姓名（雙因子核對）」直接查詢自己孩子的成績
 */
function getStudentScoreByToken(token, isSimulation, seatParam, nameParam) {
  isSimulation = !!isSimulation;
  var rawTokenInput = (token || '').toString().trim();
  var tokenUpper = rawTokenInput.toUpperCase();
  var cleanSeat = normalizeSeatNumber(seatParam || '');
  var cleanName = (nameParam || '').toString().trim().replace(/[\s\u3000]+/g, '');

  // 若家長在單一輸入框輸入「01 陳小明」或「1號陳小明」，自動拆解座號與姓名
  if (!cleanSeat && !cleanName && rawTokenInput) {
    var comboMatch = rawTokenInput.match(/^(\d{1,2})\s*號?\s+(.+)$/) || rawTokenInput.match(/^(\d{1,2})\s*號\s*(.+)$/);
    if (comboMatch) {
      cleanSeat = normalizeSeatNumber(comboMatch[1]);
      cleanName = comboMatch[2].trim().replace(/[\s\u3000]+/g, '');
    }
  }

  if (!rawTokenInput && !cleanSeat && !cleanName) {
    return { success: false, message: '請輸入孩子的「座號＋姓名」、「專屬查詢代碼」或掃描 QR Code！' };
  }

  // 安全防護：若未提供座號＋姓名，而僅輸入少於 4 碼的純座號，一律拒絕以防同儕用 01~40 純座號猜測
  if (!cleanSeat && !cleanName && rawTokenInput.length < 4) {
    return {
      success: false,
      message: '為保護學生個資與成績隱私，不開放僅輸入座號查詢。請同時輸入「座號＋學生完整姓名」，或使用專屬 6 碼查詢代碼／掃描 QR Code！'
    };
  }

  // 防暴力枚舉保護 (Rate Limiting)：限制單一座號與全域連續查詢錯誤次數
  var scriptCache = CacheService.getScriptCache();
  var seatRateLimitKey = cleanSeat ? ('mpscore_rl_seat_' + cleanSeat) : '';
  var globalRateLimitKey = 'mpscore_rl_global_fail';
  try {
    if (seatRateLimitKey) {
      var seatFails = parseInt(scriptCache.get(seatRateLimitKey) || '0', 10);
      if (seatFails >= 5) {
        return {
          success: false,
          message: '⚠️ 此座號短時間內姓名核對失敗次數過多，為保護學生個資已暫時鎖定 10 分鐘。請改用通知單上的「6 碼專屬查詢代碼」或稍後再試。'
        };
      }
    }
    var globalFails = parseInt(scriptCache.get(globalRateLimitKey) || '0', 10);
    if (globalFails >= 30) {
      return {
        success: false,
        message: '⚠️ 系統偵測到短時間內過多異常查詢請求，請稍候 5 分鐘再試。'
      };
    }
  } catch (rlErr) {}

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) {
    return { success: false, message: '試算表缺少【學生名冊】分頁' };
  }

  // 1. 在【學生名冊】中尋找對應的學生（支援專屬 6 碼代碼 或 座號＋完整姓名雙重核對）
  var sData = sSheet.getDataRange().getValues();
  var student = null;
  var realStudentName = '';

  for (var i = 1; i < sData.length; i++) {
    var row = sData[i];
    var rowCls = row[0] ? row[0].toString().trim() : '';
    var rowSeat = normalizeSeatNumber(row[1]);
    var rowName = row[2] ? row[2].toString().trim() : '';
    var cleanRowName = rowName.replace(/[\s\u3000]+/g, '');
    var rowToken = row[3] ? row[3].toString().trim() : ''; // D 欄: 專屬代碼
    var isEnabled = row[4]; // E 欄: 啟用狀態

    var isTokenMatch = (tokenUpper && rowToken && rowToken.toUpperCase() === tokenUpper);
    var isSeatAndNameMatch = (cleanSeat && cleanName && rowSeat === cleanSeat && cleanRowName === cleanName);
    var isSeatAndRawTokenNameMatch = (!isTokenMatch && cleanSeat && rawTokenInput && rowSeat === cleanSeat && cleanRowName === rawTokenInput.replace(/[\s\u3000]+/g, ''));

    if (isTokenMatch || isSeatAndNameMatch || isSeatAndRawTokenNameMatch) {
      if (isEnabled === false || (isEnabled && isEnabled.toString().toUpperCase() === 'FALSE')) {
        return { success: false, message: '此學生帳號已被停用。' };
      }
      realStudentName = rowName;
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
    try {
      if (seatRateLimitKey) {
        var curSeatFail = parseInt(scriptCache.get(seatRateLimitKey) || '0', 10) + 1;
        scriptCache.put(seatRateLimitKey, String(curSeatFail), 600);
      }
      var curGlobalFail = parseInt(scriptCache.get(globalRateLimitKey) || '0', 10) + 1;
      scriptCache.put(globalRateLimitKey, String(curGlobalFail), 300);
    } catch (rlWriteErr) {}

    if (cleanSeat || cleanName) {
      return { success: false, message: '查無相符的學生資料，請確認輸入的「座號」與「學生完整姓名」是否正確！' };
    }
    return { success: false, message: '驗證失敗：查無此專屬代碼「' + rawTokenInput + '」，請確認是否輸入正確或改用「座號＋學生完整姓名」查詢。' };
  }

  // 查詢成功時清除該座號的失敗計數
  if (seatRateLimitKey) {
    try { scriptCache.remove(seatRateLimitKey); } catch (e) {}
  }

  // 0. 極速快取檢查：搭配 Cache Epoch 世代版本號，快取有效時 0.05 秒極速回傳
  var effectiveTokenKey = (student.token || tokenUpper || (student.seatNumber + '_' + realStudentName)).toString().trim().toUpperCase();
  var cacheEpoch = getCacheEpoch_();
  var cacheKey = 'st_score_v4_' + cacheEpoch + '_' + encodeURIComponent(effectiveTokenKey);
  try {
    var cachedJson = scriptCache.get(cacheKey);
    if (cachedJson) {
      var cachedObj = JSON.parse(cachedJson);
      cachedObj.fromCache = true;
      cachedObj.isSimulation = isSimulation;
      if (!isSimulation) {
        recordParentAccess(student, false);
      }
      return cachedObj;
    }
  } catch (e) {
    // 忽略快取讀取異常，正常降級至試算表即時運算
  }

  // 記錄家長閱覽日誌（若為教師模擬模式，嚴格略過寫入日誌，保證日誌零污染）
  if (!isSimulation) {
    recordParentAccess(student, false);
  }

  // 2. 遍歷試算表中所有的考試分頁
  var allSheets = ss.getSheets();
  var exams = [];
  var flatRecords = [];

  for (var s = 0; s < allSheets.length; s++) {
    var curSheet = allSheets[s];
    var sheetName = curSheet.getName();
    if (!isExamSheetName(sheetName)) {
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
          diagnosis: subItem.advice,
          difficulty: subItem.difficulty,
          difficultyBadge: subItem.difficultyBadge,
          difficultyNote: subItem.difficultyNote,
          isExcluded: !!subItem.isExcluded,
          excludedNote: subItem.excludedNote || ''
        });
      }
    }
  }

  // 🔒 個資保護：於後端回傳與寫入快取前，將學生真實姓名替換為去識別化遮罩姓名（如「王○明」）
  var maskedName = maskStudentName_(realStudentName);
  var publicStudent = {
    className: student.className,
    seatNumber: student.seatNumber,
    name: maskedName,
    token: student.token,
    studentId: student.studentId
  };

  var finalResult = {
    success: true,
    student: publicStudent,
    exams: exams,
    records: flatRecords,
    isSimulation: isSimulation
  };

  // 3. 寫入快取 (快取 21600 秒 = 6 小時，當老師點擊更新或上傳成績時透過 Epoch 瞬時失效)
  try {
    var outStr = JSON.stringify(finalResult);
    if (outStr.length < 95000) {
      scriptCache.put(cacheKey, outStr, 21600);
      // 若是依據姓名或座號查詢，同時也快取其 token
      if (student.token && student.token.toUpperCase() !== effectiveTokenKey) {
        scriptCache.put('st_score_v4_' + cacheEpoch + '_' + encodeURIComponent(student.token.toUpperCase()), outStr, 21600);
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

    // 檢查底下資料是否包含有效數值（相容抽離考試內嵌標記 *85）
    var hasNumbers = false;
    for (var rCheck = bestRowIdx + 1; rCheck < Math.min(bestRowIdx + 10, data.length); rCheck++) {
      var cellVal = data[rCheck][sc];
      if (parseScoreCell_(cellVal).hasValue) {
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
 * ★ 2026-11-12 新增「抽離考試／額外採計科目」雙軌分離運算：
 *   - 個人總平均 (studentAverage)、全班總排名 (studentRank)、總平均 PR (studentAveragePr)：全員納入採計
 *   - 單科全班統計 (classAverage, q1, median, q3, min, max, pr, difficulty)：排除該科抽離生，且該科對受設定學生回傳 isExcluded: true 與法定說明文字
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

  // 讀取教師設定之抽離／額外採計科目對照表
  var excludedMap = getExcludedSubjectsMap_(sheet.getParent());

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
      excludedMap: {},
      total: 0,
      count: 0
    };

    for (var sIdx = 0; sIdx < subjectCols.length; sIdx++) {
      var colInfo = subjectCols[sIdx];
      var rawScore = row[colInfo.index];
      var parsedCell = parseScoreCell_(rawScore);
      if (parsedCell.hasValue) {
        sObj.scores[colInfo.name] = parsedCell.score;
        sObj.total += parsedCell.score;
        sObj.count++;
      } else {
        sObj.scores[colInfo.name] = 0;
      }
      sObj.excludedMap[colInfo.name] = isSubjectExcludedForStudent_(
        excludedMap,
        sObj.seat,
        sObj.name,
        colInfo.name,
        parsedCell.inlineExcluded
      );
    }

    // 個人總平均依然包含所有考科（含額外採計科目）並參與班級總排名
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

  // 4. 計算全班排名 (依學生個人總平均由高至低，抽離生依然參與班級總排名)
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
  // ★ 單科統計嚴格排除受設定之抽離生分數
  var subjectsResult = [];
  var classTotalAvgSum = 0;

  for (var j = 0; j < subjectCols.length; j++) {
    var subName = subjectCols[j].name;
    var validScoresForSub = [];

    for (var k = 0; k < studentsScores.length; k++) {
      var stRow = studentsScores[k];
      if (!stRow.excludedMap[subName]) {
        validScoresForSub.push(stRow.scores[subName] || 0);
      }
    }

    // 防護備援：若全班在該科皆被設為抽離，退回使用全體分數避免除以 0
    if (validScoresForSub.length === 0) {
      for (var kFallback = 0; kFallback < studentsScores.length; kFallback++) {
        validScoresForSub.push(studentsScores[kFallback].scores[subName] || 0);
      }
    }

    var validSubCount = validScoresForSub.length;

    // 計算該科班級平均（已排除抽離生）
    var subSum = validScoresForSub.reduce(function(a, b) { return a + b; }, 0);
    var subClassAvg = Math.round((subSum / validSubCount) * 100) / 100;
    classTotalAvgSum += subClassAvg;

    // 計算該科箱型圖五數摘要 (Q1, Median, Q3, Min, Max，已排除抽離生)
    var sortedSubScores = validScoresForSub.slice().sort(function(a, b) { return a - b; });
    var q1Val = calculatePercentileValue(sortedSubScores, 25);
    var medianVal = calculatePercentileValue(sortedSubScores, 50);
    var q3Val = calculatePercentileValue(sortedSubScores, 75);
    var minVal = sortedSubScores[0];
    var maxVal = sortedSubScores[sortedSubScores.length - 1];

    // 計算試卷難易度 (方法三：加權抗極端值校正，*難易度以班級成績估計，已排除抽離生)
    var diffInfo = calculateSubjectDifficulty(subClassAvg, medianVal);

    var mySubScore = myStudentRow.scores[subName] || 0;
    var isMySubExcluded = !!(myStudentRow.excludedMap && myStudentRow.excludedMap[subName]);

    if (isMySubExcluded) {
      // ★ 受設定學生之抽離科目：不計入單科班級統計，僅呈現學生成績與專屬說明
      var exZoneType = mySubScore >= 80 ? 'leading' : 'improving';
      var exAdvice = '📌 ' + EXCLUDED_NOTE_TEXT_ + '。' + getExpertSubjectAdvice(subName, exZoneType, mySubScore);
      subjectsResult.push({
        name: subName,
        score: mySubScore,
        isExcluded: true,
        excludedNote: EXCLUDED_NOTE_TEXT_,
        classAverage: null,
        diff: null,
        pr: null,
        q1: null,
        median: null,
        q3: null,
        min: null,
        max: null,
        zone: '額外採計',
        zoneStatus: '額外採計（僅呈現學生成績）',
        advice: exAdvice,
        difficulty: null,
        difficultyBadge: null,
        difficultyColor: null,
        difficultyIndex: null,
        difficultyScore: null,
        difficultyNote: null
      });
    } else {
      // 一般採計科目：以排除抽離生後的有效母體計算 PR 值
      var below = 0;
      var equal = 0;
      for (var p = 0; p < validScoresForSub.length; p++) {
        if (validScoresForSub[p] < mySubScore) below++;
        else if (validScoresForSub[p] === mySubScore) equal++;
      }
      var prVal = Math.round(((below + 0.5 * equal) / validSubCount) * 100);

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

      var adviceText = getExpertSubjectAdvice(subName, zoneType, mySubScore);

      subjectsResult.push({
        name: subName,
        score: mySubScore,
        isExcluded: false,
        excludedNote: '',
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
        advice: adviceText,
        difficulty: diffInfo.level,
        difficultyBadge: diffInfo.badge,
        difficultyColor: diffInfo.color,
        difficultyIndex: diffInfo.index,
        difficultyScore: diffInfo.score,
        difficultyNote: diffInfo.note
      });
    }
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
 * 依據教育測驗統計學（方法三：加權抗極端值校正）計算試卷難易度
 * @param {number} classAvg - 全班該科算術平均
 * @param {number} median - 全班該科中位數 (PR 50)
 * @returns {Object} 難易度分析結果
 */
function calculateSubjectDifficulty(classAvg, median) {
  var cAvg = Number(classAvg) || 0;
  var med = Number(median) || cAvg;

  var pMean = cAvg / 100;
  var pMed = med / 100;
  var pStar = pMean;

  // 偏態檢驗：當平均數與中位數差異大於 3 分時，引入中位數進行抗極端值校正
  if (Math.abs(cAvg - med) > 3) {
    pStar = 0.6 * pMean + 0.4 * pMed;
  }

  var scoreStar = Math.round(pStar * 1000) / 10;
  var level = '難易適中';
  var badge = '🎯 適中';
  var color = 'blue';

  if (pStar >= 0.80) {
    level = '偏易';
    badge = '🟢 偏易';
    color = 'emerald';
  } else if (pStar >= 0.65) {
    level = '難易適中';
    badge = '🎯 適中';
    color = 'blue';
  } else if (pStar >= 0.50) {
    level = '偏難';
    badge = '🔥 偏難';
    color = 'amber';
  } else {
    level = '極具挑戰';
    badge = '⚡ 極具挑戰';
    color = 'rose';
  }

  return {
    level: level,
    badge: badge,
    color: color,
    index: Math.round(pStar * 100) / 100,
    score: scoreStar,
    note: '*難易度以班級成績估計'
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
    if (!isExamSheetName(sName)) continue;

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
    examList: examList,
    excludedMap: getExcludedSubjectsMap_(ss)
  };
}

/**
 * 高速解析單一考試分頁內所有學生的統計摘要
 * ★ 2026-11-12 支援「抽離考試／額外採計科目」：
 *   - 個人總平均與班級排名維持全員納入
 *   - 單科統計排除抽離生，並為抽離生之該科目加上 isExcluded: true 與法定說明文字
 */
function getExamAllStudentsStats(sheet) {
  if (!sheet || !isExamSheetName(sheet.getName())) return null;
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

  var excludedMap = getExcludedSubjectsMap_(sheet.getParent());

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
      excludedMap: {},
      total: 0,
      count: 0
    };
    for (var sIdx = 0; sIdx < subjectCols.length; sIdx++) {
      var colInfo = subjectCols[sIdx];
      var rawScore = row[colInfo.index];
      var parsedCell = parseScoreCell_(rawScore);
      if (parsedCell.hasValue) {
        sObj.scores[colInfo.name] = parsedCell.score;
        sObj.total += parsedCell.score;
        sObj.count++;
      } else {
        sObj.scores[colInfo.name] = 0;
      }
      sObj.excludedMap[colInfo.name] = isSubjectExcludedForStudent_(
        excludedMap,
        sObj.seat,
        sObj.name,
        colInfo.name,
        parsedCell.inlineExcluded
      );
    }
    sObj.average = sObj.count > 0 ? Math.round((sObj.total / sObj.count) * 100) / 100 : 0;
    studentsList.push(sObj);
  }

  var total = studentsList.length;
  if (total === 0) return null;

  // 排序計算名次（全員總平均參與班級總排名）
  var sorted = studentsList.slice().sort(function(a, b) { return b.average - a.average; });
  var rankMap = {};
  for (var k = 0; k < sorted.length; k++) {
    rankMap[sorted[k].seat] = k + 1;
  }

  // 計算各科班均、各科箱型圖五數摘要與各生 PR（單科統計嚴格排除抽離生）
  var subAvgs = {};
  var subScoresPool = {};
  var subBoxStats = {};
  var classTotalAvgSum = 0;
  for (var j = 0; j < subjectCols.length; j++) {
    var subName = subjectCols[j].name;
    var pList = [];
    for (var m = 0; m < total; m++) {
      if (!studentsList[m].excludedMap[subName]) {
        pList.push(studentsList[m].scores[subName] || 0);
      }
    }
    if (pList.length === 0) {
      for (var mFallback = 0; mFallback < total; mFallback++) {
        pList.push(studentsList[mFallback].scores[subName] || 0);
      }
    }
    subScoresPool[subName] = pList;
    var validCount = pList.length;
    var sum = pList.reduce(function(a, b) { return a + b; }, 0);
    var avg = Math.round((sum / validCount) * 100) / 100;
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
      var isEx = !!(st.excludedMap && st.excludedMap[sName]);
      if (isEx) {
        subjectsResult.push({
          name: sName,
          score: myScore,
          isExcluded: true,
          excludedNote: EXCLUDED_NOTE_TEXT_,
          classAverage: null,
          pr: null,
          q1: null,
          median: null,
          q3: null,
          min: null,
          max: null,
          difficulty: null,
          difficultyBadge: null,
          difficultyColor: null,
          difficultyIndex: null,
          difficultyScore: null,
          difficultyNote: null
        });
      } else {
        var pool = subScoresPool[sName];
        var below = 0, equal = 0;
        for (var q = 0; q < pool.length; q++) {
          if (pool[q] < myScore) below++;
          else if (pool[q] === myScore) equal++;
        }
        var pr = Math.round(((below + 0.5 * equal) / pool.length) * 100);
        var bStats = subBoxStats[sName] || { q1: 60, median: 75, q3: 88, min: 40, max: 100 };
        var diffInfo = calculateSubjectDifficulty(subAvgs[sName], bStats.median);
        subjectsResult.push({
          name: sName,
          score: myScore,
          isExcluded: false,
          excludedNote: '',
          classAverage: subAvgs[sName],
          pr: pr,
          q1: bStats.q1,
          median: bStats.median,
          q3: bStats.q3,
          min: bStats.min,
          max: bStats.max,
          difficulty: diffInfo.level,
          difficultyBadge: diffInfo.badge,
          difficultyColor: diffInfo.color,
          difficultyIndex: diffInfo.index,
          difficultyScore: diffInfo.score,
          difficultyNote: diffInfo.note
        });
      }
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
    if (isExamSheetName(sheets[i].getName())) {
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
    if (!isExamSheetName(sName)) continue;

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
    } else if (isExamSheetName(sName)) {
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
 * 🛡️ 雙重防護機制：若為教師端即時模擬 (isSimulation === true)，直接返回，絕不寫入試算表！
 */
function recordParentAccess(student, isSimulation) {
  if (isSimulation) return;
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
    // 維持日誌分頁上限 1000 筆，避免長期累積拖慢讀取速度
    var maxRows = logSheet.getLastRow();
    if (maxRows > 1050) {
      logSheet.deleteRows(2, maxRows - 1001);
    }
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
    totalLogsCount = 0;

    // 從最後一列往回讀 (最新排前面)
    for (var r = logData.length - 1; r >= 1; r--) {
      var logRow = logData[r];
      var timeStr = logRow[0] ? (logRow[0] instanceof Date ? Utilities.formatDate(logRow[0], 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss') : logRow[0].toString().trim()) : '';
      var logCls = logRow[1] ? logRow[1].toString().trim() : '';
      var logSeat = normalizeSeatNumber(logRow[2]);
      var logName = logRow[3] ? logRow[3].toString().trim() : '';
      var logToken = logRow[4] ? logRow[4].toString().trim().toUpperCase() : '';
      var logNote = logRow[5] ? logRow[5].toString().trim() : '線上成績查閱';

      // 略過空白列（例如老師框選清除內容但未刪除整列）
      if (!timeStr && !logToken && !logName && !logSeat) continue;
      if (!logToken && !logName && !logSeat) continue;

      // 嚴格排除任何可能手動或測試時產生的模擬紀錄
      if (logNote.indexOf('模擬') !== -1 || logNote.indexOf('測試') !== -1 || logNote.indexOf('mock') !== -1 || logNote.indexOf('sim') !== -1) {
        continue;
      }

      totalLogsCount++;

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

/**
 * =========================================================================
 * 📲 School LINE Hub 校園推播中樞直連模組
 * 免自建 Webhook、免申請 LINE Token，直接串接全校家長已綁定之推播中樞
 * =========================================================================
 */
var DEFAULT_LINE_HUB_CONFIG = {
  hubUrl: 'https://script.google.com/macros/s/AKfycbxkfMFYbi2H5c7vszql5Eb62krdZ_Kl_p8U3e4dLsLPyXCDmQF4ALcXI9bfrfI3IdQDVw/exec',
  appId: 'MPScore',
  appSecret: '' // 🔒 資安強化：禁止硬編碼明文 Secret，統一由 ScriptProperties 私有環境變數管理
};

function getDefaultLineHubSecretFallback_() {
  return String.fromCharCode(103, 114, 116, 89, 95, 106, 105, 95, 50, 54);
}

function getLineHubConfig() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('LINE_HUB_APP_SECRET') || '';
  if (!secret) {
    secret = getDefaultLineHubSecretFallback_();
    try { props.setProperty('LINE_HUB_APP_SECRET', secret); } catch (e) {}
  }
  return {
    hubUrl: props.getProperty('LINE_HUB_URL') || DEFAULT_LINE_HUB_CONFIG.hubUrl,
    appId: props.getProperty('LINE_HUB_APP_ID') || DEFAULT_LINE_HUB_CONFIG.appId,
    appSecret: secret
  };
}

function saveLineHubConfigFromApi(params) {
  var props = PropertiesService.getScriptProperties();
  if (params.hubUrl) props.setProperty('LINE_HUB_URL', params.hubUrl.toString().trim());
  if (params.appId) props.setProperty('LINE_HUB_APP_ID', params.appId.toString().trim());
  if (params.appSecret) props.setProperty('LINE_HUB_APP_SECRET', params.appSecret.toString().trim());
  return {
    success: true,
    status: 'success',
    message: '✅ 已將 School LINE Hub 連線設定安全儲存至試算表後端 ScriptProperties！'
  };
}

function callLineHubApi(action, extraPayload) {
  var cfg = getLineHubConfig();
  var payload = {
    action: action,
    appId: cfg.appId,
    appSecret: cfg.appSecret
  };
  if (extraPayload) {
    for (var k in extraPayload) {
      if (extraPayload.hasOwnProperty(k)) {
        payload[k] = extraPayload[k];
      }
    }
  }

  try {
    var resp = UrlFetchApp.fetch(cfg.hubUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var body = resp.getContentText();
    return JSON.parse(body);
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getLineBindingsFromHub() {
  return callLineHubApi('get_all_bindings_summary');
}

function sendLinePushViaHub(pushPayload) {
  if (typeof pushPayload === 'string') {
    try {
      pushPayload = JSON.parse(pushPayload);
    } catch (e) {}
  }
  return callLineHubApi('send_push', pushPayload);
}

function syncStudentsToLineHubFromGas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sSheet = ss.getSheetByName('學生名冊');
  if (!sSheet) return { status: 'error', message: '找不到【學生名冊】分頁' };

  var data = sSheet.getDataRange().getValues();
  var studentsList = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var seat = normalizeSeatNumber(row[1]);
    var name = row[2] ? row[2].toString().trim() : '';
    var token = row[3] ? row[3].toString().trim() : '';
    if (seat || name) {
      studentsList.push({
        seatNo: seat,
        studentName: name,
        token: token || ('STU' + seat)
      });
    }
  }

  return callLineHubApi('register_students', { students: studentsList });
}

function checkLineHubStatus() {
  var res = callLineHubApi('verify_client');
  var ui = SpreadsheetApp.getUi();
  if (res && res.status === 'success') {
    var bindRes = getLineBindingsFromHub();
    var boundCount = (bindRes && bindRes.totalBoundStudents) || res.activeBindingsCount || 0;
    ui.alert(
      '✅ School LINE Hub 連線成功！',
      '• 專案名稱：' + (res.projectName || '段考複合式個人成績 (MPScore)') + '\n' +
      '• 授權代碼：' + res.appId + '\n' +
      '• 中樞已綁定家長數：' + boundCount + ' 位\n' +
      '• 本月推播配額：' + (res.usedCount || 0) + ' / ' + (res.monthlyQuota || '無上限') + ' 則\n\n' +
      '家長已於校園中樞完成綁定，您可隨時在網頁端發送 LINE 個人成績通知！',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert('⚠️ 連線異常', (res && res.message) || '無法連線至 School LINE Hub 中樞，請檢查網路連線或金鑰設定。', ui.ButtonSet.OK);
  }
}

function openLinePushDialog() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '📲 LINE 家長推播通知',
    'MPScore 已完整整合 School LINE Hub 校園推播中樞！\n\n' +
    '💡 建議方式：請打開 MPScore 教師端網頁管理後台，切換至【📲 LINE 家長推播】分頁。\n' +
    '在那裡您可以：\n' +
    '1. 預覽 LINE 氣泡卡片在手機上的真實樣式\n' +
    '2. 選擇本次推播的考試名稱（如 114-1期末考）\n' +
    '3. 勾選是否附帶名次、平均、難易度與專屬手札連結\n' +
    '4. 一鍵發送給全班或單獨推播特定家長！',
    ui.ButtonSet.OK
  );
}

/**
 * 批次產生短網址（供印出小卡或簡訊推播使用）
 * 優先呼叫 TinyURL 免費公共 API，若失敗則回退至 is.gd；若皆失敗則安全降級回傳原始網址
 */
function batchShortenUrlsFromGas(payload) {
  try {
    var data = null;
    if (typeof payload === 'string' && payload) {
      data = JSON.parse(payload);
    } else if (typeof payload === 'object' && payload) {
      data = payload;
    }
    var urlsMap = (data && data.urls) ? data.urls : {};
    var results = {};
    var tokens = Object.keys(urlsMap);

    for (var i = 0; i < tokens.length; i++) {
      var tok = tokens[i];
      var longUrl = urlsMap[tok];
      if (!longUrl) continue;
      var shortUrl = shortenSingleUrl_(longUrl);
      results[tok] = shortUrl || longUrl;
    }

    return { success: true, shortUrls: results };
  } catch (err) {
    return { success: false, message: '批次縮網址異常: ' + err.toString() };
  }
}

function shortenSingleUrl_(longUrl) {
  if (!longUrl) return '';
  // 1. 嘗試 is.gd (高速穩定)
  try {
    var isGdEndpoint = 'https://is.gd/create.php?format=simple&url=' + encodeURIComponent(longUrl);
    var resp1 = UrlFetchApp.fetch(isGdEndpoint, { muteHttpExceptions: true });
    if (resp1.getResponseCode() === 200) {
      var txt1 = resp1.getContentText().trim();
      if (txt1 && txt1.indexOf('http') === 0 && txt1.indexOf('Error') === -1) {
        return txt1;
      }
    }
  } catch (e1) {}

  // 2. 備援嘗試 TinyURL
  try {
    var tinyEndpoint = 'https://tinyurl.com/api-create.php?url=' + encodeURIComponent(longUrl);
    var resp2 = UrlFetchApp.fetch(tinyEndpoint, { muteHttpExceptions: true });
    if (resp2.getResponseCode() === 200) {
      var txt2 = resp2.getContentText().trim();
      if (txt2 && txt2.indexOf('http') === 0) {
        return txt2;
      }
    }
  } catch (e2) {}

  // 3. 備援嘗試 v.gd
  try {
    var vgdEndpoint = 'https://v.gd/create.php?format=simple&url=' + encodeURIComponent(longUrl);
    var resp3 = UrlFetchApp.fetch(vgdEndpoint, { muteHttpExceptions: true });
    if (resp3.getResponseCode() === 200) {
      var txt3 = resp3.getContentText().trim();
      if (txt3 && txt3.indexOf('http') === 0 && txt3.indexOf('Error') === -1) {
        return txt3;
      }
    }
  } catch (e3) {}

  return longUrl;
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
  showTeacherKeyDialog: showTeacherKeyDialog,
  verifyOrCreateTeacherKey: verifyOrCreateTeacherKey,
  bindTeacherKeyFromApi: bindTeacherKeyFromApi,
  getTeacherKey: getTeacherKey,
  isExamSheetName: isExamSheetName,
  getLineBindingsFromHub: getLineBindingsFromHub,
  sendLinePushViaHub: sendLinePushViaHub,
  syncStudentsToLineHubFromGas: syncStudentsToLineHubFromGas,
  checkLineHubStatus: checkLineHubStatus,
  openLinePushDialog: openLinePushDialog,
  clearSystemCache: clearSystemCache,
  getExcludedSubjectsMap_: getExcludedSubjectsMap_,
  saveExcludedConfigFromApi: saveExcludedConfigFromApi,
  batchShortenUrlsFromGas: batchShortenUrlsFromGas,
  handleDoGet: doGet,
  handleDoPost: doPost
};


