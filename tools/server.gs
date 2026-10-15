/* =====================================================================
   MÁY CHỦ NỘP BÀI & CÔNG BỐ ĐIỂM — EMS SIMULATOR (Google Apps Script)
   ---------------------------------------------------------------------
   Cài đặt (làm 1 lần — xem README.md):
     1. Tạo Google Sheet mới → Tiện ích mở rộng → Apps Script.
     2. Dán toàn bộ file này vào Code.gs, lưu, chạy hàm setup() 1 lần.
     3. Triển khai → Tùy chọn triển khai mới → Ứng dụng web:
          Thực thi với tư cách: Tôi | Người có quyền truy cập: Bất kỳ ai
     4. Sao chép URL /exec, dán vào hằng SUBMIT_API_URL trong index.html.
   Điều khiển trực tiếp trên trang tính "CauHinh" (không cần triển khai lại).
   ===================================================================== */

var SHEET_SUBMISSIONS = 'BaiNop';
var SHEET_SCORES = 'BangDiem';
var SHEET_ROSTER = 'DanhSach';
var SHEET_CONFIG = 'CauHinh';

var SUBMISSION_HEADERS = ['Thời gian nộp', 'MSSV', 'Họ và tên', 'Lớp', 'Lần nộp', 'P-f (/30)', 'Q-V (/30)', 'MVA-OPF (/30)', 'Nhận xét (/10)', 'Tổng (/100)', 'Điểm hệ 10', 'Số bản ghi', 'Nhận xét của SV', 'Chi tiết chấm (JSON)', 'Bản ghi (JSON)'];
var SCORE_HEADERS = ['MSSV', 'Họ và tên', 'Lớp', 'Số lần nộp', 'P-f (/30)', 'Q-V (/30)', 'MVA-OPF (/30)', 'Nhận xét (/10)', 'Điểm tự động (hệ 10)', 'Điểm GV điều chỉnh', 'Ghi chú GV', 'Cập nhật lúc'];
var COL = { mssv: 0, name: 1, cls: 2, count: 3, pf: 4, qv: 5, mva: 6, nx: 7, auto: 8, override: 9, gvNote: 10, updated: 11 };

var DEFAULT_CONFIG = [
  ['MO_NOP_BAI', 'TRUE', 'TRUE = nhận bài; FALSE = khóa nộp bài'],
  ['HAN_NOP', '', 'Hạn nộp, ví dụ 2026-10-30 23:59 (để trống = không giới hạn)'],
  ['SO_LAN_NOP_TOI_DA', 3, 'Số lần nộp tối đa của mỗi sinh viên'],
  ['CACH_TINH_DIEM', 'cao_nhat', 'cao_nhat = lấy lần nộp điểm cao nhất; lan_cuoi = lấy lần nộp cuối'],
  ['CONG_BO_DIEM', 'TRUE', 'TRUE = bảng điểm hiển thị công khai trên web ngay khi SV nộp'],
  ['KIEM_TRA_DANH_SACH', 'FALSE', 'TRUE = chỉ nhận MSSV có trong trang tính DanhSach'],
  ['AN_MSSV', 'TRUE', 'TRUE = che bớt chữ số MSSV trên bảng điểm công khai']
];

/* ---------------- Khởi tạo ---------------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_SUBMISSIONS, SUBMISSION_HEADERS, [2]);
  ensureSheet_(ss, SHEET_SCORES, SCORE_HEADERS, [1]);
  ensureSheet_(ss, SHEET_ROSTER, ['MSSV', 'Họ và tên', 'Lớp'], [1]);
  var cfg = ensureSheet_(ss, SHEET_CONFIG, ['Tham số', 'Giá trị', 'Giải thích'], []);
  if (cfg.getLastRow() < 2) cfg.getRange(2, 1, DEFAULT_CONFIG.length, 3).setValues(DEFAULT_CONFIG);
  cfg.autoResizeColumns(1, 3);
  return 'Đã khởi tạo xong các trang tính.';
}

function ensureSheet_(ss, name, headers, textCols) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#dbeafe');
    sh.setFrozenRows(1);
  }
  // Cột MSSV để dạng văn bản, tránh mất số 0 đầu
  textCols.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  return sh;
}

function readConfig_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  var cfg = {};
  DEFAULT_CONFIG.forEach(function (r) { cfg[r[0]] = r[1]; });
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0]) cfg[String(r[0]).trim()] = r[1];
    });
  }
  var bool = function (v) { return String(v).trim().toUpperCase() === 'TRUE'; };
  var deadline = null;
  if (cfg.HAN_NOP instanceof Date) deadline = cfg.HAN_NOP;
  else if (String(cfg.HAN_NOP).trim()) {
    var d = new Date(String(cfg.HAN_NOP).trim().replace(' ', 'T'));
    if (!isNaN(d.getTime())) deadline = d;
  }
  return {
    open: bool(cfg.MO_NOP_BAI),
    deadline: deadline,
    maxSubmissions: Math.max(1, Number(cfg.SO_LAN_NOP_TOI_DA) || 3),
    scoreMode: String(cfg.CACH_TINH_DIEM).trim() === 'lan_cuoi' ? 'lan_cuoi' : 'cao_nhat',
    publish: bool(cfg.CONG_BO_DIEM),
    checkRoster: bool(cfg.KIEM_TRA_DANH_SACH),
    maskId: bool(cfg.AN_MSSV)
  };
}

/* ---------------- Tiện ích ---------------- */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Chặn chèn công thức vào trang tính (=, +, -, @ ở đầu chuỗi)
function cell_(s, maxLen) {
  s = String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, maxLen || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function normId_(s) { return String(s == null ? '' : s).trim().toUpperCase(); }

function maskId_(id) {
  id = String(id);
  if (id.length <= 4) return id;
  return id.slice(0, 2) + new Array(id.length - 3).join('•') + id.slice(-2);
}

function fmtTime_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

/* ---------------- API đọc (GET) ---------------- */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'scores';
  var cfg = readConfig_();
  if (action === 'info') {
    var used = 0;
    var mssv = normId_(e.parameter.mssv);
    if (mssv) used = countSubmissions_(mssv);
    return json_({
      ok: true, open: cfg.open && !(cfg.deadline && new Date() > cfg.deadline),
      deadline: cfg.deadline ? fmtTime_(cfg.deadline) : '', maxSubmissions: cfg.maxSubmissions,
      used: used, scoreMode: cfg.scoreMode, publish: cfg.publish, version: EMS_GRADER.VERSION
    });
  }
  if (!cfg.publish) return json_({ ok: true, published: false, rows: [] });
  return json_({ ok: true, published: true, updated: fmtTime_(new Date()), scoreMode: cfg.scoreMode, rows: publicScores_(cfg) });
}

function publicScores_(cfg) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCORES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, SCORE_HEADERS.length).getValues()
    .filter(function (r) { return r[COL.mssv] !== ''; })
    .map(function (r) {
      var ov = r[COL.override];
      var hasOv = ov !== '' && ov !== null && isFinite(Number(ov));
      return {
        mssv: cfg.maskId ? maskId_(r[COL.mssv]) : String(r[COL.mssv]),
        name: String(r[COL.name]), cls: String(r[COL.cls]), count: Number(r[COL.count]) || 0,
        pf: Number(r[COL.pf]) || 0, qv: Number(r[COL.qv]) || 0, mva: Number(r[COL.mva]) || 0, nx: Number(r[COL.nx]) || 0,
        score: hasOv ? Number(ov) : Number(r[COL.auto]) || 0, adjusted: hasOv,
        updated: r[COL.updated] instanceof Date ? fmtTime_(r[COL.updated]) : String(r[COL.updated])
      };
    });
}

function countSubmissions_(mssv) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUBMISSIONS);
  if (!sh || sh.getLastRow() < 2) return 0;
  return sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues()
    .filter(function (r) { return normId_(r[0]) === mssv; }).length;
}

/* ---------------- API nộp bài (POST) ---------------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'Máy chủ đang bận, vui lòng thử lại sau vài giây.' });
  }
  try {
    return json_(handleSubmit_(JSON.parse(e.postData.contents)));
  } catch (err) {
    return json_({ ok: false, error: 'Dữ liệu nộp bài không hợp lệ: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

function handleSubmit_(body) {
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCORES)) setup();
  var cfg = readConfig_();
  var now = new Date();
  if (!cfg.open) return { ok: false, error: 'Giảng viên đã khóa nộp bài.' };
  if (cfg.deadline && now > cfg.deadline) return { ok: false, error: 'Đã quá hạn nộp bài (' + fmtTime_(cfg.deadline) + ').' };

  var mssv = normId_(body.mssv);
  var name = cell_(body.name, 80);
  var cls = cell_(body.className, 40);
  if (!/^[A-Z0-9][A-Z0-9._-]{3,19}$/.test(mssv)) return { ok: false, error: 'MSSV không hợp lệ (4–20 ký tự chữ/số).' };
  if (name.length < 3) return { ok: false, error: 'Vui lòng nhập đầy đủ họ và tên.' };

  if (cfg.checkRoster) {
    var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
    var found = roster && roster.getLastRow() > 1 && roster.getRange(2, 1, roster.getLastRow() - 1, 1).getValues()
      .some(function (r) { return normId_(r[0]) === mssv; });
    if (!found) return { ok: false, error: 'MSSV ' + mssv + ' không có trong danh sách lớp. Liên hệ giảng viên.' };
  }

  var used = countSubmissions_(mssv);
  if (used >= cfg.maxSubmissions) return { ok: false, error: 'Bạn đã dùng hết ' + cfg.maxSubmissions + ' lần nộp.' };

  var records = (Array.isArray(body.records) ? body.records : []).slice(0, 150).map(function (r) {
    return { module: String(r.module), scenario: Number(r.scenario), controls: r.controls || {}, hintUsed: !!r.hintUsed, note: cell_(r.note, 300) };
  });
  if (records.length === 0) return { ok: false, error: 'Bài nộp chưa có bản ghi kịch bản nào.' };
  var conclusion = cell_(body.conclusion, 5000);

  // Máy chủ tự chấm lại từ thông số thô — không tin điểm gửi từ trình duyệt
  var g = EMS_GRADER.gradeSubmission(records, conclusion);
  var attempt = used + 1;
  var detail = JSON.stringify(g);
  var recJson = JSON.stringify(records);

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUBMISSIONS).appendRow([
    now, mssv, name, cls, attempt, g.modules.pf.total, g.modules.qv.total, g.modules.mva.total, g.conclusion,
    g.total100, g.score10, records.length, conclusion, detail.slice(0, 49000), recJson.slice(0, 49000)
  ]);

  var official = upsertScore_(mssv, name, cls, attempt, g, cfg, now);
  return {
    ok: true, attempt: attempt, maxSubmissions: cfg.maxSubmissions, grading: g,
    officialScore: official, scoreMode: cfg.scoreMode, published: cfg.publish, submittedAt: fmtTime_(now)
  };
}

// Cập nhật 1 dòng/sinh viên trong BangDiem, giữ nguyên cột "Điểm GV điều chỉnh" và "Ghi chú GV"
function upsertScore_(mssv, name, cls, attempt, g, cfg, now) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCORES);
  var n = Math.max(0, sh.getLastRow() - 1);
  var rows = n ? sh.getRange(2, 1, n, SCORE_HEADERS.length).getValues() : [];
  var idx = -1;
  for (var i = 0; i < rows.length; i++) if (normId_(rows[i][COL.mssv]) === mssv) { idx = i; break; }

  var prev = idx >= 0 ? rows[idx] : null;
  var takeNew = !prev || cfg.scoreMode === 'lan_cuoi' || g.score10 >= (Number(prev[COL.auto]) || 0);
  var row = takeNew
    ? [mssv, name, cls, attempt, g.modules.pf.total, g.modules.qv.total, g.modules.mva.total, g.conclusion, g.score10]
    : [mssv, name, cls, attempt].concat(prev.slice(COL.pf, COL.auto + 1));
  row.push(prev ? prev[COL.override] : '', prev ? prev[COL.gvNote] : '', now);

  if (idx >= 0) sh.getRange(idx + 2, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return row[COL.auto];
}
