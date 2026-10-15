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

/* =====================================================================
   EMS_GRADER — MÔ HÌNH TÍNH TOÁN & THANG CHẤM ĐIỂM DÙNG CHUNG
   Khối mã này được chép NGUYÊN VĂN vào index.html (trình duyệt) và
   apps-script/Code.gs (máy chủ) bởi tools/build.py. Chỉ sửa ở file
   tools/grader.js rồi chạy lại build.py để hai nơi luôn đồng bộ.
   Máy chủ tự tính lại điểm từ thông số điều khiển thô, nên sinh viên
   không thể sửa điểm ở phía trình duyệt.
   ===================================================================== */
var EMS_GRADER = (function () {
  var VERSION = '2026.10';

  var SCENARIOS = {
    pf: {
      1: { name: 'Kịch bản 1: Cao điểm tối 19:00', load: 4200, resMax: 300, coalMax: 2200, coalInit: 1800, hydroInit: 1600, resInit: 100 },
      2: { name: 'Kịch bản 2: Giờ trưa 12:00 (Đỉnh ĐMT)', load: 2800, resMax: 2000, coalMax: 2200, coalInit: 1400, hydroInit: 400, resInit: 100 },
      3: { name: 'Kịch bản 3: Sự cố N-1 Nhảy tổ máy 600MW', load: 3800, resMax: 800, coalMax: 1400, coalInit: 1400, hydroInit: 1100, resInit: 100 }
    },
    qv: {
      1: { name: 'Kịch bản 1: Cao điểm chiều (Tải cảm lớn)', P_L: 600, Q_L: 380, Q_cap: 50, R12: 0.012, X12: 0.085, R34: 0.015, X34: 0.065 },
      2: { name: 'Kịch bản 2: Thấp điểm đêm (Hiệu ứng Ferranti)', P_L: 120, Q_L: 30, Q_cap: 280, R12: 0.012, X12: 0.085, R34: 0.015, X34: 0.065 },
      3: { name: 'Kịch bản 3: Sự cố N-1 Đường dây 500 kV', P_L: 520, Q_L: 320, Q_cap: 25, R12: 0.024, X12: 0.170, R34: 0.015, X34: 0.065 }
    },
    mva: {
      1: { name: 'Kịch bản 1: Cao điểm truyền tải (Bẫy Merit Order)', pLoad: 1000, qLoad: 360, s13Max: 500, s23Max: 600, s12Max: 450, pg1Init: 660, pg2Init: 300, qInit: 0 },
      2: { name: 'Kịch bản 2: Sự cố N-1 Nhảy 1 mạch ĐZ 1-3', pLoad: 950, qLoad: 340, s13Max: 350, s23Max: 600, s12Max: 450, pg1Init: 550, pg2Init: 300, qInit: 0 },
      3: { name: 'Kịch bản 3: Bùng nổ NLTT Nút 1 (Nghẽn mạch)', pLoad: 1100, qLoad: 380, s13Max: 480, s23Max: 550, s12Max: 400, pg1Init: 720, pg2Init: 320, qInit: 0 }
    }
  };

  var MODULE_LABEL = { pf: 'P-f', qv: 'Q-V', mva: 'MVA-OPF' };
  var OPF_HINT_PENALTY = 3;

  // Miền giá trị hợp lệ của từng thanh trượt: [min, max, bước]
  function limitsFor(mod, sc) {
    if (mod === 'pf') return { coal: [600, sc.coalMax, 20], hydro: [0, 2500, 20], res: [0, 100, 5] };
    if (mod === 'qv') return { v1: [0.95, 1.05, 0.01], qComp: [-150, 300, 10], tap: [0.92, 1.08, 0.01] };
    return { pg1: [0, 800, 10], pg2: [0, 700, 10], qComp: [0, 300, 10] };
  }

  /* ---------------- Mô hình tính toán (giữ nguyên công thức gốc) ---------------- */
  function computePF(sc, c) {
    var resMW = Math.round(sc.resMax * (c.res / 100));
    var totalGen = c.coal + c.hydro + resMW;
    var deltaP = totalGen - sc.load;
    var freq = 50.0 + (deltaP / 1000.0);
    var reserveMW = 2500 - c.hydro;
    var totalCost = (c.coal * 1.25) + (c.hydro * 1.85) + (resMW * 0.20);
    var unitCost = totalGen > 0 ? Math.round((totalCost * 1000) / totalGen) : 0;
    return { resMW: resMW, totalGen: totalGen, deltaP: deltaP, freq: freq, reserveMW: reserveMW, totalCost: totalCost, unitCost: unitCost };
  }

  function computeQV(sc, c) {
    var qNet4 = sc.Q_L - c.qComp;
    var qFlow12 = qNet4 - sc.Q_cap + 30;
    var v1 = c.v1;
    var v2 = v1 - ((sc.P_L * sc.R12 + qFlow12 * sc.X12) / 1000.0);
    var v3 = (v2 * c.tap) - ((qNet4 * 0.025) / 1000.0);
    var v4 = v3 - ((sc.P_L * sc.R34 + qNet4 * sc.X34) / (1000.0 * Math.max(0.8, v3)));
    var pLoss = (((sc.P_L * sc.P_L + qFlow12 * qFlow12) * sc.R12) + ((sc.P_L * sc.P_L + qNet4 * qNet4) * sc.R34)) / (1000.0 * v2 * v2);
    var cosPhi = sc.P_L / Math.sqrt(sc.P_L * sc.P_L + qNet4 * qNet4);
    return { qNet4: qNet4, qFlow12: qFlow12, v1: v1, v2: v2, v3: v3, v4: v4, pLoss: pLoss, cosPhi: cosPhi };
  }

  function computeMVA(sc, c) {
    var pg1 = c.pg1, pg2 = c.pg2;
    var pg3 = Math.max(0, sc.pLoad - pg1 - pg2);
    var overGen = Math.max(0, (pg1 + pg2) - sc.pLoad);
    var qNet = sc.qLoad - c.qComp;
    var p13 = (2.0 / 3.0) * pg1 + (1.0 / 3.0) * pg2;
    var p23 = (1.0 / 3.0) * pg1 + (2.0 / 3.0) * pg2;
    var p12 = Math.abs((1.0 / 3.0) * pg1 - (1.0 / 3.0) * pg2);
    var s13 = Math.sqrt(p13 * p13 + Math.pow(0.50 * qNet, 2));
    var s23 = Math.sqrt(p23 * p23 + Math.pow(0.35 * qNet, 2));
    var s12 = Math.sqrt(p12 * p12 + Math.pow(0.15 * qNet, 2));
    var pct13 = (s13 / sc.s13Max) * 100.0;
    var pct23 = (s23 / sc.s23Max) * 100.0;
    var pct12 = (s12 / sc.s12Max) * 100.0;
    var totalCost = (pg1 * 1.10) + (pg2 * 1.60) + (pg3 * 2.60);
    var idealG1 = Math.min(800, sc.pLoad);
    var idealG2 = Math.max(0, sc.pLoad - idealG1);
    var unconstrainedCost = (idealG1 * 1.10) + (idealG2 * 1.60);
    var congestionCost = Math.max(0, totalCost - unconstrainedCost);
    return {
      pg3: pg3, overGen: overGen, qNet: qNet, s13: s13, s23: s23, s12: s12,
      pct13: pct13, pct23: pct23, pct12: pct12, maxPct: Math.max(pct13, pct23, pct12),
      totalCost: totalCost, congestionCost: congestionCost
    };
  }

  /* ---------------- Kiểm tra & chuẩn hóa thông số ---------------- */
  // Trả về bộ thông số đã làm tròn theo bước thanh trượt, hoặc null nếu không hợp lệ.
  function sanitize(mod, scenarioId, controls) {
    var sc = SCENARIOS[mod] && SCENARIOS[mod][scenarioId];
    if (!sc || !controls || typeof controls !== 'object') return null;
    var lim = limitsFor(mod, sc);
    var out = {};
    for (var key in lim) {
      var v = Number(controls[key]);
      var lo = lim[key][0], hi = lim[key][1], step = lim[key][2];
      if (!isFinite(v) || v < lo - step / 2 || v > hi + step / 2) return null;
      var snapped = lo + Math.round((v - lo) / step) * step;
      snapped = Math.min(hi, Math.max(lo, snapped));
      out[key] = Math.round(snapped * 1000) / 1000;
    }
    return out;
  }

  /* ---------------- Lời giải chuẩn (tìm vét cạn theo bước thanh trượt) ---------------- */
  var benchCache = {};

  function benchmark(mod, id) {
    var key = mod + id;
    if (benchCache[key] !== undefined) return benchCache[key];
    var sc = SCENARIOS[mod][id];
    var best = Infinity;
    if (mod === 'pf') {
      // Chi phí thấp nhất khi cân bằng tuyệt đối (Δf = 0) và dự trữ quay ≥ 300 MW
      for (var coal = 600; coal <= sc.coalMax; coal += 20) {
        for (var res = 0; res <= 100; res += 5) {
          var hydro = sc.load - coal - Math.round(sc.resMax * res / 100);
          if (hydro < 0 || hydro > 2200 || hydro % 20 !== 0) continue;
          var r = computePF(sc, { coal: coal, hydro: hydro, res: res });
          if (r.totalCost < best) best = r.totalCost;
        }
      }
    } else if (mod === 'qv') {
      // Tổn thất nhỏ nhất khi cả 4 nút đạt chuẩn và |V4 − 1| ≤ 0,02 p.u.
      for (var i = 0; i <= 10; i++) {
        for (var q = -150; q <= 300; q += 10) {
          for (var t = 0; t <= 16; t++) {
            var rq = computeQV(sc, { v1: 0.95 + i * 0.01, qComp: q, tap: 0.92 + t * 0.01 });
            var ok = [rq.v1, rq.v2, rq.v3, rq.v4].every(function (v) { return v >= 0.95 && v <= 1.05; });
            if (ok && Math.abs(rq.v4 - 1) <= 0.02 && rq.pLoss < best) best = rq.pLoss;
          }
        }
      }
    } else {
      // Chi phí thấp nhất khi mọi đường dây ≤ 100 % và không thừa nguồn
      for (var g1 = 0; g1 <= 800; g1 += 10) {
        for (var g2 = 0; g2 <= 700 && g1 + g2 <= sc.pLoad; g2 += 10) {
          for (var qc = 0; qc <= 300; qc += 10) {
            var rm = computeMVA(sc, { pg1: g1, pg2: g2, qComp: qc });
            if (rm.maxPct <= 100 && rm.totalCost < best) { best = rm.totalCost; break; }
          }
        }
      }
    }
    benchCache[key] = best;
    return best;
  }

  function tier(value, steps) {
    // steps: [[ngưỡng, điểm], ...] theo thứ tự ngưỡng tăng dần
    for (var i = 0; i < steps.length; i++) if (value <= steps[i][0]) return steps[i][1];
    return 0;
  }

  /* ---------------- Chấm 1 bản ghi (thang 10 điểm / kịch bản) ---------------- */
  function gradeRecord(mod, scenarioId, rawControls, hintUsed) {
    var c = sanitize(mod, scenarioId, rawControls);
    if (!c) return { score: 0, max: 10, parts: [], note: 'Thông số điều khiển không hợp lệ' };
    var sc = SCENARIOS[mod][scenarioId];
    var parts = [];
    var note = '';

    if (mod === 'pf') {
      var r = computePF(sc, c);
      var df = Math.abs(r.freq - 50);
      var pFreq = tier(df - 1e-9, [[0.02, 5], [0.10, 4], [0.20, 3], [0.50, 1]]);
      var pRes = r.reserveMW >= 300 ? 2 : 0;
      var pEco = 0;
      if (df <= 0.20 + 1e-9) pEco = tier(r.totalCost / benchmark('pf', scenarioId), [[1.005, 3], [1.03, 2], [1.08, 1]]);
      parts = [
        { label: 'Tần số |Δf| = ' + df.toFixed(2) + ' Hz', pts: pFreq, max: 5 },
        { label: 'Dự trữ quay ' + r.reserveMW + ' MW (≥ 300)', pts: pRes, max: 2 },
        { label: 'Chi phí ' + r.totalCost.toFixed(0) + ' / tối ưu ' + benchmark('pf', scenarioId).toFixed(0) + ' Tr.đ/h', pts: pEco, max: 3 }
      ];
      if (df > 0.20) note = 'Tần số ngoài dải ±0,2 Hz nên không xét điểm kinh tế';
    } else if (mod === 'qv') {
      var q = computeQV(sc, c);
      var vs = [q.v1, q.v2, q.v3, q.v4];
      var inBand = vs.filter(function (v) { return v >= 0.95 && v <= 1.05; }).length;
      var dv4 = Math.abs(q.v4 - 1);
      var pV4 = tier(dv4, [[0.02, 3], [0.035, 2], [0.05, 1]]);
      var pLoss = 0;
      if (inBand === 4) pLoss = tier(q.pLoss / benchmark('qv', scenarioId), [[1.03, 3], [1.10, 2], [1.25, 1]]);
      parts = [
        { label: inBand + '/4 nút trong dải 0,95–1,05 p.u.', pts: inBand, max: 4 },
        { label: 'V4 = ' + q.v4.toFixed(3) + ' p.u. (lệch ' + dv4.toFixed(3) + ')', pts: pV4, max: 3 },
        { label: 'Tổn thất ' + q.pLoss.toFixed(2) + ' / tối ưu ' + benchmark('qv', scenarioId).toFixed(2) + ' MW', pts: pLoss, max: 3 }
      ];
      if (inBand < 4) note = 'Còn nút vi phạm điện áp nên không xét điểm tổn thất';
    } else {
      var m = computeMVA(sc, c);
      if (m.overGen > 0) {
        parts = [{ label: 'Thừa nguồn ' + m.overGen + ' MW (PG1 + PG2 > phụ tải)', pts: 0, max: 10 }];
        note = 'Phương án không khả thi';
      } else {
        var pTh = tier(m.maxPct, [[100, 5], [105, 2], [110, 1]]);
        var pCost = 0;
        if (m.maxPct <= 100) pCost = tier(m.totalCost / benchmark('mva', scenarioId), [[1.005, 5], [1.02, 4], [1.05, 3], [1.10, 2], [1.20, 1]]);
        parts = [
          { label: 'Mang tải lớn nhất ' + m.maxPct.toFixed(1) + ' % (≤ 100)', pts: pTh, max: 5 },
          { label: 'Chi phí ' + m.totalCost.toFixed(0) + ' / tối ưu ' + benchmark('mva', scenarioId).toFixed(0) + ' Tr.đ/h', pts: pCost, max: 5 }
        ];
        if (m.maxPct > 100) note = 'Còn đường dây quá tải nên không xét điểm kinh tế';
        if (hintUsed) parts.push({ label: 'Dùng nút "Gợi ý Điểm Tối ưu OPF"', pts: -OPF_HINT_PENALTY, max: 0 });
      }
    }

    var score = 0;
    parts.forEach(function (p) { score += p.pts; });
    return { score: Math.max(0, Math.min(10, score)), max: 10, parts: parts, note: note };
  }

  /* ---------------- Chấm phần nhận xét / kết luận (10 điểm) ---------------- */
  function gradeConclusion(text) {
    var n = String(text || '').replace(/\s+/g, ' ').trim().length;
    return tier(-n, [[-400, 10], [-250, 8], [-120, 5], [-40, 2]]);
  }

  /* ---------------- Chấm toàn bộ bài nộp (thang 100 → hệ 10) ---------------- */
  // records: [{ module: 'pf'|'qv'|'mva', scenario: 1..3, controls: {...}, hintUsed: bool }]
  function gradeSubmission(records, conclusion) {
    var result = { version: VERSION, modules: {}, total100: 0, score10: 0 };
    ['pf', 'qv', 'mva'].forEach(function (mod) {
      var mres = { label: MODULE_LABEL[mod], total: 0, max: 30, scenarios: {} };
      [1, 2, 3].forEach(function (id) {
        var best = { score: 0, max: 10, parts: [], note: 'Chưa ghi nhận kịch bản này', recordIndex: -1 };
        (records || []).forEach(function (rec, idx) {
          if (!rec || rec.module !== mod || Number(rec.scenario) !== id) return;
          var g = gradeRecord(mod, id, rec.controls, !!rec.hintUsed);
          if (best.recordIndex < 0 || g.score > best.score) { g.recordIndex = idx; best = g; }
        });
        mres.scenarios[id] = best;
        mres.total += best.score;
      });
      result.modules[mod] = mres;
      result.total100 += mres.total;
    });
    result.conclusion = gradeConclusion(conclusion);
    result.total100 += result.conclusion;
    result.score10 = Math.round(result.total100) / 10;
    return result;
  }

  return {
    VERSION: VERSION,
    SCENARIOS: SCENARIOS,
    MODULE_LABEL: MODULE_LABEL,
    OPF_HINT_PENALTY: OPF_HINT_PENALTY,
    computePF: computePF,
    computeQV: computeQV,
    computeMVA: computeMVA,
    sanitize: sanitize,
    benchmark: benchmark,
    gradeRecord: gradeRecord,
    gradeConclusion: gradeConclusion,
    gradeSubmission: gradeSubmission
  };
})();
