/**
 * けやっきーのせかいりょこう迷路 — 授業用 API
 *
 * セットアップ:
 * 1. 新しい Google スプレッドシートを作成
 * 2. 拡張機能 → Apps Script を開き、このファイルの内容を貼り付け
 * 3. 下記 SPREADSHEET_ID を自分のシートIDに書き換え（またはこのスクリプトを
 *    スプレッドシート紐付けで作れば getActiveSpreadsheet で動く）
 * 4. 初回は ensureSheets_() が config / progress シートを自動作成
 * 5. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 実行ユーザー: 自分
 *    - アクセスできるユーザー: 全員（匿名ユーザーを含む）
 *      ※「組織内のみ」だと GitHub Pages から CORS/401 エラーになります
 * 6. 発行された URL をリポジトリの config.js の gasUrl に貼る
 */

var SPREADSHEET_ID = ''; // 空ならスクリプトに紐付いたスプレッドシートを使う
var TEACHER_SESSION_HOURS = 8;

function doGet(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var action = p.action || '';
    if (!action) {
      return json_({ ok: true, service: 'labyrinth-classroom', version: 1 });
    }
    var body = bodyFromParams_(p);
    if (action === 'studentLogin') return json_(studentLogin_(body));
    if (action === 'saveProgress') return json_(saveProgress_(body));
    return json_({ ok: false, error: 'use_post_for_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function bodyFromParams_(p) {
  return {
    action: p.action || '',
    grade: p.grade !== undefined && p.grade !== '' ? Number(p.grade) : undefined,
    className: p.className !== undefined && p.className !== '' ? Number(p.className) : undefined,
    studentNo: p.studentNo !== undefined && p.studentNo !== '' ? Number(p.studentNo) : undefined,
    cleared: p.cleared || '',
    perfect: p.perfect || '',
    masterCleared: p.masterCleared || '',
    masterPerfect: p.masterPerfect || '',
    medals: p.medals || '',
    password: p.password || '',
    token: p.token || ''
  };
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action || '';
    if (action === 'studentLogin') return json_(studentLogin_(body));
    if (action === 'saveProgress') return json_(saveProgress_(body));
    if (action === 'teacherLogin') return json_(teacherLogin_(body));
    if (action === 'getClassProgress') return json_(getClassProgress_(body));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheets_() {
  var ss = getSpreadsheet_();
  var progress = ss.getSheetByName('progress');
  if (!progress) {
    progress = ss.insertSheet('progress');
    progress.appendRow([
      'key', 'grade', 'className', 'studentNo',
      'cleared', 'perfect', 'masterCleared', 'masterPerfect',
      'medals', 'updatedAt'
    ]);
  }

  var config = ss.getSheetByName('config');
  if (!config) {
    config = ss.insertSheet('config');
    config.appendRow(['key', 'value']);
    config.appendRow(['teacherPassword', 'sensei']);
    config.appendRow(['allowedGrades', '1,2,3,4,5,6']);
    config.appendRow(['maxStudentNo', '40']);
    config.appendRow(['maxClass', '6']);
  }
  return { ss: ss, progress: progress, config: config };
}

function getConfigMap_() {
  var sheets = ensureSheets_();
  var values = sheets.config.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var k = String(values[i][0] || '').trim();
    if (k) map[k] = String(values[i][1] == null ? '' : values[i][1]).trim();
  }
  return map;
}

function makeStudentKey_(grade, className, studentNo) {
  return String(grade) + '-' + String(className) + '-' + String(studentNo);
}

function parseList_(s) {
  if (!s) return [];
  return String(s).split(',').map(function (x) {
    return String(x).trim();
  }).filter(Boolean);
}

function mergeCsv_(a, b) {
  var set = {};
  parseList_(a).forEach(function (x) { set[x] = true; });
  parseList_(b).forEach(function (x) { set[x] = true; });
  return Object.keys(set).sort(function (x, y) {
    var nx = Number(x);
    var ny = Number(y);
    if (!isNaN(nx) && !isNaN(ny)) return nx - ny;
    return String(x).localeCompare(String(y));
  }).join(',');
}

function findProgressRow_(sheet, key) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === key) return i + 1; // 1-based row
  }
  return -1;
}

function rowToProgress_(row) {
  return {
    key: String(row[0] || ''),
    grade: Number(row[1]),
    className: Number(row[2]),
    studentNo: Number(row[3]),
    cleared: String(row[4] || ''),
    perfect: String(row[5] || ''),
    masterCleared: String(row[6] || ''),
    masterPerfect: String(row[7] || ''),
    medals: String(row[8] || ''),
    updatedAt: String(row[9] || '')
  };
}

function validateStudent_(grade, className, studentNo) {
  var cfg = getConfigMap_();
  var allowed = parseList_(cfg.allowedGrades || '1,2,3,4,5,6');
  var maxNo = Number(cfg.maxStudentNo || 40);
  var maxClass = Number(cfg.maxClass || 6);
  grade = Number(grade);
  className = Number(className);
  studentNo = Number(studentNo);
  if (!allowed.some(function (g) { return Number(g) === grade; })) {
    return { ok: false, error: 'invalid_grade' };
  }
  if (!(className >= 1 && className <= maxClass)) {
    return { ok: false, error: 'invalid_class' };
  }
  if (!(studentNo >= 1 && studentNo <= maxNo) || studentNo !== Math.floor(studentNo)) {
    return { ok: false, error: 'invalid_student_no' };
  }
  return { ok: true, grade: grade, className: className, studentNo: studentNo };
}

function studentLogin_(body) {
  var v = validateStudent_(body.grade, body.className, body.studentNo);
  if (!v.ok) return v;
  var sheets = ensureSheets_();
  var key = makeStudentKey_(v.grade, v.className, v.studentNo);
  var rowIndex = findProgressRow_(sheets.progress, key);
  if (rowIndex < 0) {
    var now = new Date().toISOString();
    sheets.progress.appendRow([
      key, v.grade, v.className, v.studentNo,
      '', '', '', '', '', now
    ]);
    return {
      ok: true,
      created: true,
      progress: {
        key: key,
        grade: v.grade,
        className: v.className,
        studentNo: v.studentNo,
        cleared: '',
        perfect: '',
        masterCleared: '',
        masterPerfect: '',
        medals: '',
        updatedAt: now
      }
    };
  }
  var row = sheets.progress.getRange(rowIndex, 1, rowIndex, 10).getValues()[0];
  return { ok: true, created: false, progress: rowToProgress_(row) };
}

function saveProgress_(body) {
  var v = validateStudent_(body.grade, body.className, body.studentNo);
  if (!v.ok) return v;
  var sheets = ensureSheets_();
  var key = makeStudentKey_(v.grade, v.className, v.studentNo);
  var rowIndex = findProgressRow_(sheets.progress, key);
  var now = new Date().toISOString();

  var incoming = {
    cleared: body.cleared || '',
    perfect: body.perfect || '',
    masterCleared: body.masterCleared || '',
    masterPerfect: body.masterPerfect || '',
    medals: body.medals || ''
  };

  if (rowIndex < 0) {
    sheets.progress.appendRow([
      key, v.grade, v.className, v.studentNo,
      incoming.cleared, incoming.perfect,
      incoming.masterCleared, incoming.masterPerfect,
      incoming.medals, now
    ]);
    return { ok: true, progress: Object.assign({ key: key, grade: v.grade, className: v.className, studentNo: v.studentNo, updatedAt: now }, incoming) };
  }

  var current = rowToProgress_(sheets.progress.getRange(rowIndex, 1, rowIndex, 10).getValues()[0]);
  var merged = {
    cleared: mergeCsv_(current.cleared, incoming.cleared),
    perfect: mergeCsv_(current.perfect, incoming.perfect),
    masterCleared: mergeCsv_(current.masterCleared, incoming.masterCleared),
    masterPerfect: mergeCsv_(current.masterPerfect, incoming.masterPerfect),
    medals: mergeCsv_(current.medals, incoming.medals)
  };
  sheets.progress.getRange(rowIndex, 5, rowIndex, 10).setValues([[
    merged.cleared, merged.perfect, merged.masterCleared, merged.masterPerfect, merged.medals, now
  ]]);
  return {
    ok: true,
    progress: {
      key: key,
      grade: v.grade,
      className: v.className,
      studentNo: v.studentNo,
      cleared: merged.cleared,
      perfect: merged.perfect,
      masterCleared: merged.masterCleared,
      masterPerfect: merged.masterPerfect,
      medals: merged.medals,
      updatedAt: now
    }
  };
}

function makeTeacherToken_() {
  return Utilities.getUuid().replace(/-/g, '');
}

function teacherLogin_(body) {
  var cfg = getConfigMap_();
  var password = String(body.password || '');
  if (!password || password !== (cfg.teacherPassword || 'sensei')) {
    return { ok: false, error: 'bad_password' };
  }
  var token = makeTeacherToken_();
  CacheService.getScriptCache().put('teacher_' + token, '1', TEACHER_SESSION_HOURS * 3600);
  return { ok: true, token: token, expiresInHours: TEACHER_SESSION_HOURS };
}

function requireTeacher_(token) {
  if (!token) return false;
  return CacheService.getScriptCache().get('teacher_' + token) === '1';
}

function getClassProgress_(body) {
  if (!requireTeacher_(body.token)) {
    return { ok: false, error: 'unauthorized' };
  }
  var grade = Number(body.grade);
  var className = Number(body.className);
  if (!grade || !className) return { ok: false, error: 'invalid_class_filter' };

  var sheets = ensureSheets_();
  var values = sheets.progress.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < values.length; i++) {
    var g = Number(values[i][1]);
    var c = Number(values[i][2]);
    if (g === grade && c === className) {
      students.push(rowToProgress_(values[i]));
    }
  }
  students.sort(function (a, b) { return a.studentNo - b.studentNo; });
  return { ok: true, grade: grade, className: className, students: students };
}
