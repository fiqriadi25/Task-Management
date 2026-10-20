// ==================== KONFIGURASI ====================
var CONFIG = {
  SPREADSHEET_ID: "1NVmo_Zqbp_5N8BFrBmtD2-tIb7rViQ9VsKDz8LaHU0c",
  SHEET_TASKS: "TRACKER",
  SHEET_USERS: "USERS",
  SHEET_COMMENTS: "COMMENTS",
  SHEET_KENDALA: "FOLLOW_UP_KENDALA",
  SHEET_TEMPLATE: "TEMPLATE_TASK",
  CACHE_TTL: 300 // seconds
};

// ==================== CACHE HELPER ====================
var _cache = {};

function _getCached(key) {
  var cached = _cache[key];
  if (cached && (Date.now() - cached.ts) < CONFIG.CACHE_TTL * 1000) {
    return cached.data;
  }
  return null;
}


// ==================== PERFORMANCE OPTIMIZATION v1 (2026-08-11) ====================
// Satu pembacaan TRACKER untuk tabel + statistik + jumlah data + assignee.
var PM_CACHE_TTL_SECONDS = 300;
var PM_USERS_CACHE_KEY = "pm_users_v1";
var PM_CONFIG_CACHE_KEY = "pm_config_v1";

function _pmCacheGet(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    console.warn("_pmCacheGet: " + e.message);
    return null;
  }
}

function _pmCachePut(key, value) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), PM_CACHE_TTL_SECONDS);
  } catch (e) {
    console.warn("_pmCachePut: " + e.message);
  }
}

function _pmReadUsedRange(sheetName) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { sheet: sheet, headers: [], rows: [], col: {} };
  }
  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0] || [];
  return {
    sheet: sheet,
    headers: headers,
    rows: data.slice(1),
    col: buildColMap(headers)
  };
}

// Override cache USERS agar berlaku lintas eksekusi Apps Script.
function getUsers() {
  var cached = _pmCacheGet(PM_USERS_CACHE_KEY);
  if (cached) return cached;

  try {
    var s = _pmReadUsedRange(CONFIG.SHEET_USERS);
    if (!s.rows.length) return [];

    var required = ["USER_ID", "NAMA_LENGKAP", "ROLE", "EMAIL"];
    for (var r = 0; r < required.length; r++) {
      if (s.col[required[r]] === undefined) {
        throw new Error("Kolom wajib tidak ditemukan di USERS: " + required[r]);
      }
    }

    var users = s.rows.map(function(row) {
      return {
        USER_ID: row[s.col.USER_ID],
        NAMA_LENGKAP: row[s.col.NAMA_LENGKAP] || "",
        ROLE: row[s.col.ROLE] || "user",
        EMAIL: row[s.col.EMAIL] || ""
      };
    }).filter(function(user) {
      return user.USER_ID !== "" && user.USER_ID !== null;
    });

    _pmCachePut(PM_USERS_CACHE_KEY, users);
    return users;
  } catch (e) {
    console.error("getUsers optimized error: " + e.message);
    return [];
  }
}

// Override cache CONFIG agar dropdown tidak selalu membaca sheet.
function getAppConfig() {
  var cached = _pmCacheGet(PM_CONFIG_CACHE_KEY);
  if (cached) return cached;

  var empty = { KATEGORI: [], PRIORITAS: [], STATUS: [], PENYEBAB_KENDALA: [] };
  try {
    var sheet = getSpreadsheet().getSheetByName("CONFIG");
    if (!sheet) return empty;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return empty;

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var config = { KATEGORI: [], PRIORITAS: [], STATUS: [], PENYEBAB_KENDALA: [] };

    for (var i = 1; i < data.length; i++) {
      var type = data[i][0];
      var value = data[i][1];
      if (!type || value === "" || value === null) continue;
      var item = {
        value: value,
        label: data[i][2] || value,
        order: data[i][3] || 999
      };
      if (type === "KATEGORI") config.KATEGORI.push(item);
      else if (type === "PRIORITAS") config.PRIORITAS.push(item);
      else if (type === "STATUS") config.STATUS.push(item);
      else if (type === "PENYEBAB_KENDALA") config.PENYEBAB_KENDALA.push(item);
    }

    Object.keys(config).forEach(function(key) {
      config[key].sort(function(a, b) { return a.order - b.order; });
    });

    _pmCachePut(PM_CONFIG_CACHE_KEY, config);
    return config;
  } catch (e) {
    console.error("getAppConfig optimized error: " + e.message);
    return empty;
  }
}

function _pmResolveCurrentUser(users) {
  try {
    var email = Session.getEffectiveUser().getEmail();
    if (!email) return _guestUser();
    var target = String(email).toLowerCase();
    for (var i = 0; i < users.length; i++) {
      if (String(users[i].EMAIL || "").toLowerCase() === target) return users[i];
    }
    return {
      USER_ID: null,
      NAMA_LENGKAP: email.split("@")[0],
      ROLE: "user",
      EMAIL: email
    };
  } catch (e) {
    return _guestUser();
  }
}

function _pmGetKendalaSnapshot() {
  var snapshot = {
    map: {},
    stats: { total: 0, byStatus: {}, byPenyebab: {} }
  };

  try {
    var sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_KENDALA);
    if (!sheet) return snapshot;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return snapshot;

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = data[0];
    var taskCol = headers.indexOf("TASK_ID");
    var statusCol = headers.indexOf("STATUS");
    var penyebabCol = headers.indexOf("PENYEBAB");

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var taskId = taskCol !== -1 ? String(row[taskCol] || "").trim() : "";
      if (taskId) snapshot.map[taskId] = true;

      snapshot.stats.total++;
      var status = statusCol !== -1 ? (row[statusCol] || "Open") : "Open";
      snapshot.stats.byStatus[status] = (snapshot.stats.byStatus[status] || 0) + 1;

      if (penyebabCol !== -1 && row[penyebabCol]) {
        var penyebab = row[penyebabCol];
        snapshot.stats.byPenyebab[penyebab] =
          (snapshot.stats.byPenyebab[penyebab] || 0) + 1;
      }
    }
  } catch (e) {
    console.error("_pmGetKendalaSnapshot error: " + e.message);
  }
  return snapshot;
}

function _pmBuildDashboardBundle(filters, page, pageSize, users, kendalaSnapshot) {
  var startedAt = Date.now();
  filters = filters || {};
  page = Math.max(1, Number(page) || 1);
  pageSize = Math.max(1, Math.min(200, Number(pageSize) || 25));

  var source = _pmReadUsedRange(CONFIG.SHEET_TASKS);
  if (!source.rows.length) {
    return {
      tasks: [],
      totalCount: 0,
      totalPages: 0,
      currentPage: page,
      pageSize: pageSize,
      stats: { total: 0, byStatus: {}, byPriority: {}, byKategori: {}, overdue: 0 },
      totalJumlahData: 0,
      assigneeStats: [],
      durationMs: Date.now() - startedAt
    };
  }

  if (source.col.TASK_ID === undefined) {
    throw new Error("Kolom TASK_ID tidak ditemukan di sheet " + CONFIG.SHEET_TASKS);
  }

  var userMap = {};
  (users || []).forEach(function(user) {
    userMap[user.USER_ID] = user.NAMA_LENGKAP;
  });

  var filtered = [];
  var assigneeMap = {};
  var stats = { total: 0, byStatus: {}, byPriority: {}, byKategori: {}, overdue: 0 };
  var totalJumlahData = 0;
  var statsAssignee = filters.assignee ? String(filters.assignee) : "";
  var today = new Date();
  today.setHours(0, 0, 0, 0);

  for (var i = 0; i < source.rows.length; i++) {
    var row = source.rows[i];
    if (!row[source.col.TASK_ID] && !row[source.col.JUDUL]) continue;
    if (source.col.IS_DELETED !== undefined && row[source.col.IS_DELETED] === true) continue;

    var task = rowToTask(row, source.col, userMap);
    task.HAS_KENDALA = !!kendalaSnapshot.map[String(task.TASK_ID).trim()];

    var assignee = task.ASSIGN_TO || "Unassigned";
    assigneeMap[assignee] = (assigneeMap[assignee] || 0) + 1;

    if (!statsAssignee || String(task.ASSIGN_TO) === statsAssignee) {
      stats.total++;
      if (task.STATUS) stats.byStatus[task.STATUS] = (stats.byStatus[task.STATUS] || 0) + 1;
      if (task.PRIORITY) stats.byPriority[task.PRIORITY] = (stats.byPriority[task.PRIORITY] || 0) + 1;
      if (task.KATEGORI) stats.byKategori[task.KATEGORI] = (stats.byKategori[task.KATEGORI] || 0) + 1;

      var jumlah = Number(task.JUMLAH_DATA);
      if (!isNaN(jumlah)) totalJumlahData += jumlah;

      if (task.DUE_DATE && task.STATUS !== "Done" && task.STATUS !== "Closed") {
        if (new Date(task.DUE_DATE) < today) stats.overdue++;
      }
    }

    if (taskMatchesFilter(task, filters)) filtered.push(task);
  }

  filtered.sort(function(a, b) {
    return new Date(b.CREATED_AT) - new Date(a.CREATED_AT);
  });

  var totalCount = filtered.length;
  var totalPages = Math.ceil(totalCount / pageSize);
  var start = (page - 1) * pageSize;

  var assigneeStats = Object.keys(assigneeMap).map(function(name) {
    return { assignee_name: name, count: assigneeMap[name] };
  }).sort(function(a, b) {
    return b.count - a.count;
  });

  return {
    tasks: filtered.slice(start, start + pageSize),
    totalCount: totalCount,
    totalPages: totalPages,
    currentPage: page,
    pageSize: pageSize,
    stats: stats,
    totalJumlahData: totalJumlahData,
    assigneeStats: assigneeStats,
    durationMs: Date.now() - startedAt
  };
}

function getDashboardBundle(filters, page, pageSize) {
  var users = getUsers();
  var kendalaSnapshot = _pmGetKendalaSnapshot();
  return _pmBuildDashboardBundle(
    filters,
    page,
    pageSize,
    users,
    kendalaSnapshot
  );
}

function getDashboardBootstrap(filters, page, pageSize) {
  var users = getUsers();
  var kendalaSnapshot = _pmGetKendalaSnapshot();
  return {
    config: getAppConfig(),
    currentUser: _pmResolveCurrentUser(users),
    users: users,
    kendalaStats: kendalaSnapshot.stats,
    dashboard: _pmBuildDashboardBundle(
      filters,
      page,
      pageSize,
      users,
      kendalaSnapshot
    )
  };
}


function _setCached(key, data) {
  _cache[key] = { data: data, ts: Date.now() };
}

function _invalidateCache(key) {
  delete _cache[key];
}

// ==================== SPREADSHEET HELPER ====================
var _ss = null;
var _timezoneCache = null;

function getSpreadsheet() {
  if (!_ss) {
    _ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return _ss;
}

function getSheet(name) {
  var sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error("Sheet '" + name + "' tidak ditemukan.");
  return sheet;
}

function _getSpreadsheetTimeZone() {
  if (_timezoneCache === null) {
    _timezoneCache = getSpreadsheet().getSpreadsheetTimeZone();
  }
  return _timezoneCache;
}

function _formatDate(date, format) {
  if (!date) return "";
  if (!(date instanceof Date)) return String(date);
  var tz = _getSpreadsheetTimeZone();
  return Utilities.formatDate(date, tz, format || "yyyy-MM-dd HH:mm:ss");
}

function buildColMap(headers) {
  var col = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] !== "") col[headers[i]] = i;
  }
  return col;
}

function readSheet(sheetName) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  if (data.length < 1) return { sheet: sheet, headers: [], rows: [], col: {} };
  var headers = data[0];
  var col = buildColMap(headers);
  var rows = data.slice(1);
  return { sheet: sheet, headers: headers, rows: rows, col: col };
}

function generateNextId(sheet, col, colIndex) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  var ids = sheet
    .getRange(2, colIndex + 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(function(v) {
      var num = Number(v);
      if (!isNaN(num)) return num;
      var match = String(v).match(/\d+/);
      return match ? parseInt(match[0], 10) : 0;
    })
    .filter(function(v) { return v > 0; });
  return ids.length ? Math.max.apply(null, ids) + 1 : 1;
}

function _normalizeSearchTerm(term) {
  return String(term || "").toLowerCase().trim().replace(/^#/, "");
}

// ==================== TANGGAL: HELPER ANTI-BUG TIMEZONE ====================
/**
 * Parse tanggal dengan aman dari pergeseran timezone DAN dari ambiguitas format.
 * Tanggal yang diinput/di-upload user TIDAK BOLEH berubah akibat timezone maupun
 * akibat native Date() menebak format (mis. "11/07/2026" ditebak sebagai MM/DD/YYYY
 * padahal maksud user adalah DD/MM/YYYY).
 *
 * Urutan pengecekan:
 * 1. Date object            -> ambil komponen UTC (bukan local/project timezone)
 * 2. Angka (serial Excel)   -> hasil baca file .xlsx di frontend kadang berupa serial number
 * 3. String "yyyy-MM-dd..." -> format ISO, paling aman & eksplisit
 * 4. String "dd/mm/yyyy" atau "dd-mm-yyyy" -> format umum pada file upload Indonesia
 * 5. Fallback native Date() -> hanya jika semua format eksplisit di atas tidak cocok
 *
 * Semua hasil dinormalisasi ke jam 12:00 siang WAKTU ASIA/JAKARTA (WIB, UTC+7)
 * pada tanggal yang dimaksud (lihat _jakartaAnchor), supaya tanggal yang tersimpan
 * & ditampilkan di Sheets selalu jatuh di hari yang benar, terlepas dari timezone
 * proyek Apps Script atau timezone spreadsheet.
 */
// Asia/Jakarta (WIB) = UTC+7, tanpa DST, jadi offset-nya selalu tetap.
var JAKARTA_UTC_OFFSET_HOURS = 7;

/**
 * Membuat Date yang "jangkar"-nya adalah jam 12:00 siang waktu Asia/Jakarta
 * pada tanggal (y, m, d) yang dimaksud. Dengan begini, berapa pun timezone
 * proyek Apps Script / timezone spreadsheet, tanggal yang tampil di Sheets
 * akan selalu jatuh di hari yang benar (tidak pernah mundur/maju sehari).
 * m: 0-indexed (Januari = 0), d: 1-indexed.
 */
function _jakartaAnchor(y, m, d) {
  return new Date(Date.UTC(y, m, d, 12 - JAKARTA_UTC_OFFSET_HOURS, 0, 0));
}

function _parseDueDateSafe(dateStr) {
  if (dateStr === null || dateStr === undefined || dateStr === "") return null;

  // 1) Date object -> PENTING: pakai komponen UTC, BUKAN local getter.
  //    Local getter (getFullYear/getMonth/getDate) mengikuti timezone proyek
  //    Apps Script yang belum tentu Asia/Jakarta, dan itu penyebab tanggal
  //    bisa mundur satu hari. Nilai tanggal-saja (date-only) dari JSON/xlsx
  //    secara standar direpresentasikan sebagai UTC-midnight, jadi komponen
  //    UTC adalah yang paling bisa dipercaya sebagai "tanggal yang dimaksud".
  if (dateStr instanceof Date) {
    if (isNaN(dateStr.getTime())) return null;
    return _jakartaAnchor(dateStr.getUTCFullYear(), dateStr.getUTCMonth(), dateStr.getUTCDate());
  }

  // 2) Serial number ala Excel (epoch 1899-12-30, termasuk bug leap-year bawaan Excel)
  if (typeof dateStr === "number") {
    if (isNaN(dateStr) || dateStr <= 0) return null;
    var excelEpochMs = Date.UTC(1899, 11, 30);
    var ms = excelEpochMs + Math.round(dateStr) * 24 * 60 * 60 * 1000;
    var dExcel = new Date(ms);
    if (isNaN(dExcel.getTime())) return null;
    return _jakartaAnchor(dExcel.getUTCFullYear(), dExcel.getUTCMonth(), dExcel.getUTCDate());
  }

  var s = String(dateStr).trim();
  if (!s) return null;

  // 3) Format ISO: yyyy-MM-dd (boleh diikuti waktu, diabaikan)
  var mIso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (mIso) {
    var y1 = parseInt(mIso[1], 10), mo1 = parseInt(mIso[2], 10), d1 = parseInt(mIso[3], 10);
    var dtIso = _jakartaAnchor(y1, mo1 - 1, d1);
    return isNaN(dtIso.getTime()) ? null : dtIso;
  }

  // 4) Format dd/mm/yyyy atau dd-mm-yyyy (format umum di file template bulk upload)
  var mDmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (mDmy) {
    var d2 = parseInt(mDmy[1], 10), mo2 = parseInt(mDmy[2], 10), y2 = parseInt(mDmy[3], 10);
    if (mo2 >= 1 && mo2 <= 12 && d2 >= 1 && d2 <= 31) {
      var dtDmy = _jakartaAnchor(y2, mo2 - 1, d2);
      return isNaN(dtDmy.getTime()) ? null : dtDmy;
    }
  }

  // 5) Fallback terakhir: parse native. String tanggal-saja tanpa zona waktu
  //    diperlakukan JS sebagai UTC-midnight, jadi tetap pakai komponen UTC.
  var parsed = new Date(s);
  if (isNaN(parsed.getTime())) return null;
  return _jakartaAnchor(parsed.getUTCFullYear(), parsed.getUTCMonth(), parsed.getUTCDate());
}

function _shiftDate(baseDate, offsetSeconds) {
  if (!baseDate) baseDate = new Date();
  return new Date(baseDate.getTime() + offsetSeconds * 1000);
}

// ==================== WEB APP ====================
function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Project Management")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== USERS ====================
function getUsers() {
  var cached = _getCached("users");
  if (cached) return cached;

  try {
    var s = readSheet(CONFIG.SHEET_USERS);
    if (!s.rows.length) return [];

    var required = ["USER_ID", "NAMA_LENGKAP", "ROLE", "EMAIL"];
    for (var r = 0; r < required.length; r++) {
      if (s.col[required[r]] === undefined) {
        throw new Error("Kolom wajib tidak ditemukan di USERS: " + required[r]);
      }
    }

    var users = s.rows.map(function(row) {
      return {
        USER_ID: row[s.col.USER_ID],
        NAMA_LENGKAP: row[s.col.NAMA_LENGKAP] || "",
        ROLE: row[s.col.ROLE] || "user",
        EMAIL: row[s.col.EMAIL] || ""
      };
    }).filter(function(u) { return u.USER_ID !== "" && u.USER_ID !== null; });

    _setCached("users", users);
    return users;
  } catch (e) {
    console.error("getUsers error: " + e.message);
    return [];
  }
}

function buildUserMap() {
  var map = {};
  getUsers().forEach(function(u) { map[u.USER_ID] = u.NAMA_LENGKAP; });
  return map;
}

function getCurrentUser() {
  try {
    var email = Session.getEffectiveUser().getEmail();
    if (!email) return _guestUser();
    var users = getUsers();
    var found = users.find(function(u) {
      return u.EMAIL.toLowerCase() === email.toLowerCase();
    });
    if (found) return found;
    return {
      USER_ID: null,
      NAMA_LENGKAP: email.split("@")[0],
      ROLE: "user",
      EMAIL: email
    };
  } catch (e) {
    console.error("getCurrentUser error: " + e.message);
    return _guestUser();
  }
}

function _guestUser() {
  return { USER_ID: null, NAMA_LENGKAP: "Guest", ROLE: "viewer", EMAIL: "" };
}

// ==================== VALIDASI TERPUSAT ====================
/**
 * Membangun "master data" (users + config) sekali pakai untuk divalidasi
 * terhadap setiap baris task (single maupun bulk). Dibangun sekali per
 * request supaya ringan (tidak query sheet berulang per baris).
 */
function _buildValidationContext() {
  var users = getUsers();
  var userIdSet = {};
  var userNameSet = {};
  users.forEach(function(u) {
    if (u.USER_ID !== "" && u.USER_ID !== null && u.USER_ID !== undefined) {
      userIdSet[String(u.USER_ID).trim().toLowerCase()] = true;
    }
    if (u.NAMA_LENGKAP) {
      userNameSet[String(u.NAMA_LENGKAP).trim().toLowerCase()] = true;
    }
  });

  var cfg = getAppConfig();
  var kategoriSet = {};
  (cfg.KATEGORI || []).forEach(function(k) {
    kategoriSet[String(k.value).trim().toLowerCase()] = true;
  });
  var priorityValueSet = {};
  (cfg.PRIORITAS || []).forEach(function(p) {
    priorityValueSet[String(p.value).trim().toLowerCase()] = true;
  });

  return {
    userIdSet: userIdSet,
    userNameSet: userNameSet,
    kategoriSet: kategoriSet,
    priorityValueSet: priorityValueSet
  };
}

/**
 * Validasi field wajib dasar (JUDUL, DESKRIPSI, JUMLAH_DATA, STATUS, PRIORITY, DUE_DATE).
 * Dipakai untuk addTask, updateTask, dan bulkAddTasks.
 */
function validateTaskData(taskData) {
  var errors = [];

  if (!taskData.JUDUL || String(taskData.JUDUL).trim() === "") {
    errors.push("JUDUL tidak boleh kosong.");
  }
  if (!taskData.DESKRIPSI || String(taskData.DESKRIPSI).trim() === "") {
    errors.push("DESKRIPSI tidak boleh kosong.");
  }
  if (taskData.JUMLAH_DATA === undefined || taskData.JUMLAH_DATA === null || String(taskData.JUMLAH_DATA).trim() === "") {
    errors.push("JUMLAH_DATA tidak boleh kosong.");
  } else {
    var jd = Number(taskData.JUMLAH_DATA);
    if (isNaN(jd) || jd < 0) {
      errors.push("JUMLAH_DATA harus berupa angka positif.");
    }
  }
  if (!taskData.STATUS || String(taskData.STATUS).trim() === "") {
    errors.push("STATUS tidak boleh kosong.");
  }
  if (!taskData.PRIORITY || String(taskData.PRIORITY).trim() === "") {
    errors.push("PRIORITY tidak boleh kosong.");
  }
  if (taskData.DUE_DATE && !_parseDueDateSafe(taskData.DUE_DATE)) {
    errors.push("DUE_DATE format tidak valid.");
  }

  return errors;
}

/**
 * Validasi field terhadap master data: KATEGORI & PRIORITY harus terdaftar
 * di sheet CONFIG, ASSIGN_TO (jika diisi) harus terdaftar di sheet USERS.
 */
function _validateAgainstMasterData(taskData, ctx) {
  var errors = [];

  if (taskData.KATEGORI && String(taskData.KATEGORI).trim() !== "") {
    var kat = String(taskData.KATEGORI).trim().toLowerCase();
    if (!ctx.kategoriSet[kat]) {
      errors.push("KATEGORI '" + taskData.KATEGORI + "' tidak terdaftar pada CONFIG.");
    }
  }

  if (taskData.PRIORITY && String(taskData.PRIORITY).trim() !== "") {
    var pri = String(taskData.PRIORITY).trim().toLowerCase();
    if (!ctx.priorityValueSet[pri]) {
      errors.push("PRIORITY '" + taskData.PRIORITY + "' tidak terdaftar pada CONFIG.");
    }
  }

  if (taskData.ASSIGN_TO && String(taskData.ASSIGN_TO).trim() !== "") {
    var assignee = String(taskData.ASSIGN_TO).trim().toLowerCase();
    if (!ctx.userNameSet[assignee]) {
      errors.push("ASSIGN_TO '" + taskData.ASSIGN_TO + "' tidak terdaftar di sheet USERS.");
    }
  }

  return errors;
}

// ==================== TASKS ====================
var TASK_FIELDS = [
  "TASK_ID", "TIMESTAMP", "JUDUL", "KATEGORI", "DESKRIPSI",
  "JUMLAH_DATA",
  "ASSIGN_TO", "ASSIGN_USER_ID", "STATUS", "PRIORITY",
  "DUE_DATE", "ATTACHMENT_LINK", "LAST_UPDATE"
];

function rowToTask(row, col, userMap) {
  var dueDate = row[col.DUE_DATE];
  if (dueDate instanceof Date) {
    dueDate = _formatDate(dueDate, "yyyy-MM-dd");
  }
  var createdAt = (col.TIMESTAMP !== undefined) ? row[col.TIMESTAMP] : "";
  if (createdAt instanceof Date) {
    createdAt = _formatDate(createdAt, "yyyy-MM-dd HH:mm:ss");
  }
  var lastUpdate = (col.LAST_UPDATE !== undefined) ? row[col.LAST_UPDATE] : "";
  if (lastUpdate instanceof Date) {
    lastUpdate = _formatDate(lastUpdate, "yyyy-MM-dd HH:mm:ss");
  }

  return {
    TASK_ID:         row[col.TASK_ID] || "",
    CREATED_AT:      createdAt || null,
    JUDUL:           row[col.JUDUL] || "",
    KATEGORI:        row[col.KATEGORI] || "",
    DESKRIPSI:       row[col.DESKRIPSI] || "",
    JUMLAH_DATA:     (col.JUMLAH_DATA !== undefined && row[col.JUMLAH_DATA] !== "") ? Number(row[col.JUMLAH_DATA]) : 1,
    ASSIGN_TO:       row[col.ASSIGN_TO] || "",
    ASSIGN_USER_ID:  row[col.ASSIGN_USER_ID] || "",
    ASSIGNEE_NAME:   userMap[row[col.ASSIGN_USER_ID]] || "",
    STATUS:          row[col.STATUS] || "",
    PRIORITY:        row[col.PRIORITY] || "",
    DUE_DATE:        dueDate || null,
    ATTACHMENT_LINK: row[col.ATTACHMENT_LINK] || "",
    LAST_UPDATE:     lastUpdate || null
  };
}

function taskMatchesFilter(task, filters) {
  if (!filters) return true;
  if (filters.status && task.STATUS !== filters.status) return false;
  if (filters.priority && task.PRIORITY !== filters.priority) return false;
  if (filters.kategori && task.KATEGORI !== filters.kategori) return false;
  if (filters.assignee && String(task.ASSIGN_TO) !== String(filters.assignee)) return false;
  if (filters.creator && String(task.ASSIGN_USER_ID) !== String(filters.creator)) return false;

  if (filters.search) {
    var s = _normalizeSearchTerm(filters.search);
    var inId        = String(task.TASK_ID || "").toLowerCase().indexOf(s) !== -1;
    var inJudul     = (task.JUDUL || "").toLowerCase().indexOf(s) !== -1;
    var inDeskripsi = (task.DESKRIPSI || "").toLowerCase().indexOf(s) !== -1;
    var inKategori  = (task.KATEGORI || "").toLowerCase().indexOf(s) !== -1;
    if (!inId && !inJudul && !inDeskripsi && !inKategori) return false;
  }
  return true;
}

function getTasks(filters, page, pageSize) {
  filters = filters || {};
  try {
    var s = readSheet(CONFIG.SHEET_TASKS);
    if (!s.rows.length) {
      if (pageSize && pageSize > 0 && page && page > 0) {
        return { tasks: [], totalCount: 0, totalPages: 0, currentPage: page, pageSize: pageSize };
      }
      return [];
    }

    if (s.col.TASK_ID === undefined) throw new Error("Kolom TASK_ID tidak ditemukan di sheet " + CONFIG.SHEET_TASKS);

    var userMap = buildUserMap();
    var kendalaMap = {};
    try { kendalaMap = _buildKendalaMap(); } catch(e) { console.warn("Gagal build kendala map: " + e.message); }

    var tasks = [];
    for (var i = 0; i < s.rows.length; i++) {
      var row = s.rows[i];
      if (!row[s.col.TASK_ID] && !row[s.col.JUDUL]) continue;
      if (s.col.IS_DELETED !== undefined && row[s.col.IS_DELETED] === true) continue;
      var task = rowToTask(row, s.col, userMap);
      task.HAS_KENDALA = !!kendalaMap[String(task.TASK_ID).trim()];
      if (taskMatchesFilter(task, filters)) {
        tasks.push(task);
      }
    }

    tasks.sort(function(a, b) {
      return new Date(b.CREATED_AT) - new Date(a.CREATED_AT);
    });

    if (pageSize && pageSize > 0 && page && page > 0) {
      var totalCount = tasks.length;
      var totalPages = Math.ceil(totalCount / pageSize);
      var start = (page - 1) * pageSize;
      var end = start + pageSize;
      return {
        tasks: tasks.slice(start, end),
        totalCount: totalCount,
        totalPages: totalPages,
        currentPage: page,
        pageSize: pageSize
      };
    }

    return tasks;
  } catch (e) {
    console.error("getTasks error: " + e.message);
    return { error: e.message };
  }
}

function getTaskById(taskId) {
  try {
    var s = readSheet(CONFIG.SHEET_TASKS);
    var userMap = buildUserMap();
    var targetId = String(taskId).trim();
    for (var i = 0; i < s.rows.length; i++) {
      var row = s.rows[i];
      if (String(row[s.col.TASK_ID]).trim() === targetId) {
        if (s.col.IS_DELETED !== undefined && row[s.col.IS_DELETED] === true) return null;
        var task = rowToTask(row, s.col, userMap);
        task.HAS_KENDALA = !!_buildKendalaMap()[targetId];
        return task;
      }
    }
    return null;
  } catch (e) {
    console.error("getTaskById error: " + e.message);
    return null;
  }
}

function addTask(taskData) {
  try {
    var ctx = _buildValidationContext();
    var errors = validateTaskData(taskData).concat(_validateAgainstMasterData(taskData, ctx));
    if (errors.length) return { success: false, message: errors.join(" ") };

    var s = readSheet(CONFIG.SHEET_TASKS);
    var sheet = s.sheet;
    var headers = s.headers;

    var now = new Date();
    var dueDate = _parseDueDateSafe(taskData.DUE_DATE);
    if (!dueDate) {
      dueDate = new Date();
      dueDate.setDate(dueDate.getDate() + 3);
    }

    var currentUserObj = getCurrentUser();
    if (!currentUserObj.USER_ID) {
      throw new Error("USER_ID tidak ditemukan. Pastikan user sudah terdaftar di sheet USERS.");
    }

    var assignUserId = currentUserObj.USER_ID;
    var createdByEmail = currentUserObj.EMAIL || _safeGetEmail();
    var newId = generateTaskId(sheet, currentUserObj, undefined, 0, now);

    var jumlahData = Number(taskData.JUMLAH_DATA);

    var newRow = headers.map(function(header) {
      switch (header) {
        case "TASK_ID":         return newId;
        case "TIMESTAMP":       return now;
        case "JUDUL":           return String(taskData.JUDUL || "").trim();
        case "KATEGORI":        return taskData.KATEGORI || "";
        case "DESKRIPSI":       return String(taskData.DESKRIPSI || "").trim();
        case "JUMLAH_DATA":     return jumlahData;
        case "ASSIGN_TO":       return taskData.ASSIGN_TO || "";
        case "ASSIGN_USER_ID":  return assignUserId;
        case "STATUS":          return taskData.STATUS || "";
        case "PRIORITY":        return taskData.PRIORITY || "";
        case "DUE_DATE":        return dueDate;
        case "ATTACHMENT_LINK": return taskData.ATTACHMENT_LINK || "";
        case "CREATED_BY":      return createdByEmail;
        case "LAST_UPDATE":     return now;
        case "LAST_STATUS":     return taskData.STATUS || "";
        case "IS_DELETED":      return false;
        default:                return "";
      }
    });

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    return { success: true, taskId: newId };
  } catch (e) {
    console.error("addTask error: " + e.message);
    return { success: false, message: e.message };
  }
}

function updateTask(taskId, taskData) {
  try {
    var ctx = _buildValidationContext();
    var errors = validateTaskData(taskData).concat(_validateAgainstMasterData(taskData, ctx));
    if (errors.length) return { success: false, message: errors.join(" ") };

    var s = readSheet(CONFIG.SHEET_TASKS);
    var sheet = s.sheet;
    var col = s.col;
    var taskIdStr = String(taskId).trim();
    var rowIndex = -1;

    for (var i = 0; i < s.rows.length; i++) {
      if (String(s.rows[i][col.TASK_ID]).trim() === taskIdStr) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "Task ID " + taskIdStr + " tidak ditemukan." };
    }

    var sheetRowNumber = rowIndex + 2;
    var jumlahData = Number(taskData.JUMLAH_DATA);

    var updates = {
      JUDUL:           String(taskData.JUDUL || "").trim(),
      KATEGORI:        taskData.KATEGORI || "",
      DESKRIPSI:       String(taskData.DESKRIPSI || "").trim(),
      JUMLAH_DATA:     jumlahData,
      ASSIGN_TO:       taskData.ASSIGN_TO || "",
      STATUS:          taskData.STATUS || "",
      PRIORITY:        taskData.PRIORITY || "",
      DUE_DATE:        _parseDueDateSafe(taskData.DUE_DATE) || "",
      ATTACHMENT_LINK: taskData.ATTACHMENT_LINK || "",
      LAST_UPDATE:     new Date(),
      LAST_STATUS:     taskData.STATUS || ""
    };

    for (var key in updates) {
      var idx = col[key];
      if (idx !== undefined) {
        sheet.getRange(sheetRowNumber, idx + 1).setValue(updates[key]);
      }
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    console.error("updateTask error: " + e.message);
    return { success: false, message: e.message };
  }
}

function deleteTask(taskId) {
  try {
    var s = readSheet(CONFIG.SHEET_TASKS);
    var sheet = s.sheet;
    var col = s.col;
    var taskIdStr = String(taskId).trim();
    var rowIndex = -1;

    for (var i = 0; i < s.rows.length; i++) {
      if (String(s.rows[i][col.TASK_ID]).trim() === taskIdStr) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "Task ID " + taskIdStr + " tidak ditemukan." };

    var sheetRowNumber = rowIndex + 2;
    if (col.IS_DELETED !== undefined) {
      sheet.getRange(sheetRowNumber, col.IS_DELETED + 1).setValue(true);
      if (col.LAST_UPDATE !== undefined) {
        sheet.getRange(sheetRowNumber, col.LAST_UPDATE + 1).setValue(new Date());
      }
    } else {
      sheet.deleteRow(sheetRowNumber);
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    console.error("deleteTask error: " + e.message);
    return { success: false, message: e.message };
  }
}

function bulkUpdateStatus(taskIds, newStatus) {
  try {
    if (!taskIds || !taskIds.length) return { success: false, message: "Tidak ada task yang dipilih." };
    if (!newStatus) return { success: false, message: "Status tidak boleh kosong." };

    var s = readSheet(CONFIG.SHEET_TASKS);
    var sheet = s.sheet;
    var col = s.col;
    var now = new Date();
    var updated = 0;

    taskIds.forEach(function(taskId) {
      var taskIdStr = String(taskId).trim();
      var rowIndex = -1;
      for (var i = 0; i < s.rows.length; i++) {
        if (String(s.rows[i][col.TASK_ID]).trim() === taskIdStr) {
          rowIndex = i; break;
        }
      }
      if (rowIndex === -1) return;
      var sheetRowNumber = rowIndex + 2;
      if (col.STATUS !== undefined)      sheet.getRange(sheetRowNumber, col.STATUS + 1).setValue(newStatus);
      if (col.LAST_STATUS !== undefined) sheet.getRange(sheetRowNumber, col.LAST_STATUS + 1).setValue(newStatus);
      if (col.LAST_UPDATE !== undefined) sheet.getRange(sheetRowNumber, col.LAST_UPDATE + 1).setValue(now);
      updated++;
    });

    SpreadsheetApp.flush();
    return { success: true, updated: updated };
  } catch (e) {
    console.error("bulkUpdateStatus error: " + e.message);
    return { success: false, message: e.message };
  }
}

// ==================== BULK UPLOAD TASK ====================
var BULK_TEMPLATE_HEADERS = [
  "JUDUL", "KATEGORI", "DESKRIPSI", "JUMLAH_DATA",
  "ASSIGN_TO", "STATUS", "PRIORITY", "DUE_DATE", "ATTACHMENT_LINK"
];

function downloadTaskTemplate() {
  var tempFile = null;
  try {
    var tempSs = SpreadsheetApp.create("TEMP_Template_Bulk_Task_" + new Date().getTime());
    // Samakan timezone spreadsheet sementara ini dengan Asia/Jakarta supaya
    // konsisten dengan bagaimana tanggal nantinya diproses & ditampilkan.
    try { tempSs.setSpreadsheetTimeZone("Asia/Jakarta"); } catch (eTz) { console.warn("Gagal set timezone: " + eTz.message); }

    var sheet = tempSs.getSheets()[0];
    sheet.setName(CONFIG.SHEET_TEMPLATE);

    sheet.getRange(1, 1, 1, BULK_TEMPLATE_HEADERS.length)
      .setValues([BULK_TEMPLATE_HEADERS])
      .setFontWeight("bold")
      .setBackground("#d9e1f2");

    var dueDateColIdx = BULK_TEMPLATE_HEADERS.indexOf("DUE_DATE") + 1;

    // PENTING: paksa kolom DUE_DATE berformat TEXT (bukan Date/Number) SEBELUM
    // diisi nilai apa pun. Ini mencegah Google Sheets otomatis menebak-nebak
    // dan mengonversi string tanggal ("11/07/2026") menjadi objek Date sesuai
    // locale-nya sendiri (yang bisa salah menafsirkan DD/MM jadi MM/DD).
    // Dengan kolom bertipe Text, apa pun yang diketik/ditempel user akan
    // tersimpan APA ADANYA, dan diparse manual & eksplisit di Code.gs.
    sheet.getRange(2, dueDateColIdx, 499, 1).setNumberFormat("@");

    sheet.getRange(2, 1, 1, BULK_TEMPLATE_HEADERS.length).setValues([[
      "Contoh: Update data toko cabang A", "Daily Task", "Deskripsi singkat task", 1,
      "FIQRI ADIPERDANA", "To Do", "Medium", "11/07/2026", ""
    ]]);

    // Catatan panduan format tanggal langsung di header kolom DUE_DATE
    sheet.getRange(1, dueDateColIdx).setNote(
      "Format tanggal yang didukung:\n" +
      "- dd/mm/yyyy (contoh: 11/07/2026 = 11 Juli 2026)\n" +
      "- yyyy-MM-dd (contoh: 2026-07-11)\n\n" +
      "Kolom ini sengaja berformat TEXT agar tanggal yang Anda ketik " +
      "tidak diubah otomatis oleh Sheets/Excel."
    );

    try {
      var cfg = getAppConfig();
      _applyDropdownValidation(sheet, cfg.KATEGORI, BULK_TEMPLATE_HEADERS.indexOf("KATEGORI") + 1);
      _applyDropdownValidation(sheet, cfg.STATUS, BULK_TEMPLATE_HEADERS.indexOf("STATUS") + 1);
      _applyDropdownValidation(sheet, cfg.PRIORITAS, BULK_TEMPLATE_HEADERS.indexOf("PRIORITY") + 1);
    } catch (e) { console.warn("Gagal menerapkan dropdown template: " + e.message); }

    sheet.autoResizeColumns(1, BULK_TEMPLATE_HEADERS.length);
    sheet.setColumnWidth(BULK_TEMPLATE_HEADERS.indexOf("DESKRIPSI") + 1, 250);
    SpreadsheetApp.flush();

    var tempId = tempSs.getId();
    tempFile = DriveApp.getFileById(tempId);

    var exportMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    var exportUrl = "https://www.googleapis.com/drive/v3/files/" + tempId + "/export?mimeType=" + encodeURIComponent(exportMime);
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("Gagal mengekspor template (HTTP " + response.getResponseCode() + ")");
    }

    var base64 = Utilities.base64Encode(response.getBlob().getBytes());
    return { success: true, base64: base64, filename: "Template_Bulk_Task.xlsx" };
  } catch (e) {
    console.error("downloadTaskTemplate error: " + e.message);
    return { success: false, message: e.message };
  } finally {
    if (tempFile) {
      try { tempFile.setTrashed(true); } catch (e2) { console.warn("Gagal hapus temp template: " + e2.message); }
    }
  }
}

function _applyDropdownValidation(sheet, options, colIndex) {
  if (!options || !options.length || !colIndex) return;
  var values = options.map(function(o) { return o.value; });
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true).build();
  sheet.getRange(2, colIndex, 500, 1).setDataValidation(rule);
}

/**
 * Bulk upload task dari file (JUDUL, KATEGORI, DESKRIPSI, JUMLAH_DATA, ASSIGN_TO,
 * STATUS, PRIORITY, DUE_DATE, ATTACHMENT_LINK per baris).
 *
 * Validasi per baris (baris gagal TIDAK menghentikan baris lain, hanya dicatat di errors[]):
 * - JUDUL, DESKRIPSI, JUMLAH_DATA wajib diisi (JUMLAH_DATA harus angka >= 0)
 * - STATUS, PRIORITY wajib diisi
 * - KATEGORI (jika diisi) harus terdaftar di CONFIG
 * - PRIORITY harus terdaftar di CONFIG
 * - ASSIGN_TO (jika diisi) harus terdaftar di sheet USERS
 * - DUE_DATE mengikuti tanggal apa adanya dari file (anti pergeseran timezone/format)
 *
 * Validasi sebelum proses (menghentikan seluruh batch):
 * - User yang melakukan upload (USER_ID) harus terdaftar di sheet USERS
 */
function bulkAddTasks(tasksData) {
  try {
    if (!tasksData || !tasksData.length) {
      return { success: false, message: "Tidak ada data untuk diupload." };
    }
    if (tasksData.length > 1000) {
      return { success: false, message: "Maksimal 1000 baris per upload. Silakan pecah file menjadi beberapa batch." };
    }

    var currentUserObj = getCurrentUser();
    if (!currentUserObj.USER_ID) {
      return { success: false, message: "USER_ID Anda belum terdaftar di sheet USERS. Hubungi admin sebelum melakukan bulk upload." };
    }

    var s = readSheet(CONFIG.SHEET_TASKS);
    var sheet = s.sheet;
    var headers = s.headers;
    if (s.col.TASK_ID === undefined) {
      throw new Error("Kolom TASK_ID tidak ditemukan di sheet " + CONFIG.SHEET_TASKS);
    }

    var ctx = _buildValidationContext();
    var now = new Date();
    var createdByEmail = currentUserObj.EMAIL || _safeGetEmail();
    var baseLastRow = sheet.getLastRow();
    var rowsToAppend = [];
    var errors = [];

    tasksData.forEach(function(taskData, idx) {
      var rowNum = idx + 1;
      var cleanTask = {
        JUDUL:           taskData.JUDUL,
        KATEGORI:        taskData.KATEGORI,
        DESKRIPSI:       taskData.DESKRIPSI,
        JUMLAH_DATA:     taskData.JUMLAH_DATA,
        ASSIGN_TO:       taskData.ASSIGN_TO,
        STATUS:          taskData.STATUS,
        PRIORITY:        taskData.PRIORITY,
        DUE_DATE:        taskData.DUE_DATE,
        ATTACHMENT_LINK: taskData.ATTACHMENT_LINK
      };

      var rowErrors = validateTaskData(cleanTask).concat(_validateAgainstMasterData(cleanTask, ctx));
      if (rowErrors.length) {
        errors.push({ row: rowNum, judul: cleanTask.JUDUL || "(kosong)", message: rowErrors.join(" ") });
        return;
      }

      var jumlahData = Number(cleanTask.JUMLAH_DATA);
      // Tanggal deadline mengikuti PERSIS apa yang ada di file upload (tidak digeser timezone)
      var dueDate = _parseDueDateSafe(cleanTask.DUE_DATE) || "";

      // Timestamp & TASK_ID digeser +1 detik per baris agar tetap unik/berurutan
      var shiftedTime = _shiftDate(now, idx);
      var taskId = generateTaskId(sheet, currentUserObj, baseLastRow, idx, now);

      var newRow = headers.map(function(header) {
        switch (header) {
          case "TASK_ID":         return taskId;
          case "TIMESTAMP":       return shiftedTime;
          case "JUDUL":           return String(cleanTask.JUDUL).trim();
          case "KATEGORI":        return cleanTask.KATEGORI || "";
          case "DESKRIPSI":       return String(cleanTask.DESKRIPSI).trim();
          case "JUMLAH_DATA":     return jumlahData;
          case "ASSIGN_TO":       return cleanTask.ASSIGN_TO || "";
          case "ASSIGN_USER_ID":  return currentUserObj.USER_ID;
          case "STATUS":          return cleanTask.STATUS || "";
          case "PRIORITY":        return cleanTask.PRIORITY || "";
          case "DUE_DATE":        return dueDate;
          case "ATTACHMENT_LINK": return cleanTask.ATTACHMENT_LINK || "";
          case "CREATED_BY":      return createdByEmail;
          case "LAST_UPDATE":     return shiftedTime;
          case "LAST_STATUS":     return cleanTask.STATUS || "";
          case "IS_DELETED":      return false;
          default:                return "";
        }
      });

      rowsToAppend.push(newRow);
    });

    if (rowsToAppend.length) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
      SpreadsheetApp.flush();
    }

    return {
      success: true,
      added: rowsToAppend.length,
      failed: errors.length,
      errors: errors
    };
  } catch (e) {
    console.error("bulkAddTasks error: " + e.message);
    return { success: false, message: e.message };
  }
}

function getTaskStats() {
  try {
    var tasks = getTasks();
    if (tasks && tasks.error) return { error: tasks.error };

    var stats = { total: tasks.length, byStatus: {}, byPriority: {}, byKategori: {}, overdue: 0 };
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    tasks.forEach(function(task) {
      stats.byStatus[task.STATUS]   = (stats.byStatus[task.STATUS] || 0) + 1;
      if (task.PRIORITY) stats.byPriority[task.PRIORITY] = (stats.byPriority[task.PRIORITY] || 0) + 1;
      if (task.KATEGORI) stats.byKategori[task.KATEGORI] = (stats.byKategori[task.KATEGORI] || 0) + 1;
      if (task.DUE_DATE && task.STATUS !== "Done" && task.STATUS !== "Closed") {
        if (new Date(task.DUE_DATE) < today) stats.overdue++;
      }
    });

    return stats;
  } catch (e) {
    console.error("getTaskStats error: " + e.message);
    return { error: e.message };
  }
}

// ==================== ASSIGNEE STATS ====================
function getTasksPerAssignee() {
  try {
    var tasks = getTasks();
    if (tasks.error) return [];
    var assigneeMap = {};
    tasks.forEach(function(task) {
      var assignee = task.ASSIGNEE_NAME || task.ASSIGN_TO || 'Unassigned';
      assigneeMap[assignee] = (assigneeMap[assignee] || 0) + 1;
    });
    var result = [];
    for (var name in assigneeMap) {
      result.push({ assignee_name: name, count: assigneeMap[name] });
    }
    result.sort(function(a, b) { return b.count - a.count; });
    return result;
  } catch (e) {
    console.error('getTasksPerAssignee error: ' + e.message);
    return [];
  }
}

// ==================== COMMENTS ====================
function getComments(taskId) {
  try {
    if (!taskId) return [];
    var s = readSheet(CONFIG.SHEET_COMMENTS);
    if (!s.rows.length) return [];

    var userMap = buildUserMap();
    var taskIdStr = String(taskId).trim();

    var comments = s.rows
      .filter(function(row) {
        return String(row[s.col.TASK_ID]).trim() === taskIdStr;
      })
      .map(function(row) {
        var ts = row[s.col.TIMESTAMP];
        return {
          COMMENT_ID:        row[s.col.COMMENT_ID] || "",
          TASK_ID:           row[s.col.TASK_ID] || "",
          USER_ID:           row[s.col.USER_ID] || "",
          USER_NAME:         userMap[row[s.col.USER_ID]] || "Unknown",
          COMMENT:           row[s.col.COMMENT] || "",
          TIMESTAMP:         ts instanceof Date ? _formatDate(ts, "yyyy-MM-dd HH:mm:ss") : (ts || ""),
          PARENT_COMMENT_ID: row[s.col.PARENT_COMMENT_ID] || null
        };
      });

    comments.sort(function(a, b) { return new Date(a.TIMESTAMP) - new Date(b.TIMESTAMP); });
    return comments;
  } catch (e) {
    console.error("getComments error: " + e.message);
    return [];
  }
}

function addComment(commentData) {
  try {
    if (!commentData.TASK_ID) return { success: false, message: "TASK_ID wajib diisi." };
    if (!commentData.COMMENT || String(commentData.COMMENT).trim() === "") {
      return { success: false, message: "Komentar tidak boleh kosong." };
    }

    var taskIdStr = String(commentData.TASK_ID).trim();
    var task = getTaskById(taskIdStr);
    if (!task) return { success: false, message: "Task ID " + taskIdStr + " tidak ditemukan." };

    var s = readSheet(CONFIG.SHEET_COMMENTS);
    var sheet = s.sheet;
    var col = s.col;
    var headers = s.headers;

    if (col.COMMENT_ID === undefined) {
      return { success: false, message: "Sheet COMMENTS tidak memiliki kolom COMMENT_ID." };
    }

    var newId = generateNextId(sheet, col, col.COMMENT_ID);

    var newRow = headers.map(function(header) {
      switch (header) {
        case "COMMENT_ID":        return newId;
        case "TASK_ID":           return taskIdStr;
        case "USER_ID":           return commentData.USER_ID || "";
        case "COMMENT":           return String(commentData.COMMENT).trim();
        case "TIMESTAMP":         return new Date();
        case "PARENT_COMMENT_ID": return commentData.PARENT_COMMENT_ID || "";
        default:                  return "";
      }
    });

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    return { success: true, commentId: newId };
  } catch (e) {
    console.error("addComment error: " + e.message);
    return { success: false, message: e.message };
  }
}

// ==================== FOLLOW UP KENDALA ====================
function _buildKendalaMap() {
  try {
    var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_KENDALA);
    if (!sh) return {};
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return {};
    var headers = data[0];
    var taskCol = headers.indexOf("TASK_ID");
    if (taskCol === -1) return {};
    var map = {};
    for (var i = 1; i < data.length; i++) {
      var tid = String(data[i][taskCol]).trim();
      if (tid) map[tid] = true;
    }
    return map;
  } catch(e) {
    console.error("_buildKendalaMap error: " + e.message);
    return {};
  }
}

function getKendala(taskId) {
  try {
    var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_KENDALA);
    if (!sh) return [];
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = data[0];
    var taskIdStr = String(taskId).trim();

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = row[idx]; });

      if (String(obj.TASK_ID).trim() !== taskIdStr) continue;

      if (obj.CREATED_AT instanceof Date) {
        obj.CREATED_AT = _formatDate(obj.CREATED_AT, "yyyy-MM-dd HH:mm:ss");
      }
      if (obj.LAST_UPDATE instanceof Date) {
        obj.LAST_UPDATE = _formatDate(obj.LAST_UPDATE, "yyyy-MM-dd HH:mm:ss");
      }
      result.push(obj);
    }

    result.sort(function(a, b) { return new Date(b.CREATED_AT) - new Date(a.CREATED_AT); });
    return result;
  } catch(e) {
    console.error("getKendala error: " + e.message);
    return [];
  }
}

function addKendala(data) {
  try {
    if (!data.TASK_ID) return { success: false, message: "TASK_ID wajib diisi." };
    if (!data.KENDALA || String(data.KENDALA).trim() === "") {
      return { success: false, message: "Deskripsi kendala tidak boleh kosong." };
    }

    var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_KENDALA);
    if (!sh) return { success: false, message: "Sheet " + CONFIG.SHEET_KENDALA + " tidak ditemukan." };

    var now = new Date();
    var user = _safeGetEmail();
    var newId = Utilities.getUuid();

    var lastRow = sh.getLastRow();
    if (lastRow < 1) {
      sh.appendRow(["ID", "TASK_ID", "KENDALA", "PENYEBAB", "STATUS", "CREATED_AT", "CREATED_BY", "LAST_UPDATE"]);
    }

    var headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var hMap = buildColMap(headerRow);

    var newRow = headerRow.map(function(h) {
      switch (h) {
        case "ID":          return newId;
        case "TASK_ID":     return String(data.TASK_ID).trim();
        case "KENDALA":     return String(data.KENDALA).trim();
        case "PENYEBAB":    return data.PENYEBAB || "";
        case "STATUS":      return data.STATUS || "Open";
        case "CREATED_AT":  return now;
        case "CREATED_BY":  return user;
        case "LAST_UPDATE": return now;
        default:            return "";
      }
    });

    sh.appendRow(newRow);
    SpreadsheetApp.flush();
    return { success: true, id: newId };
  } catch(e) {
    console.error("addKendala error: " + e.message);
    return { success: false, message: e.message };
  }
}

function updateKendalaStatus(id, status) {
  try {
    if (!id) return { success: false, message: "ID wajib diisi." };
    if (!status) return { success: false, message: "Status wajib diisi." };

    var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_KENDALA);
    if (!sh) return { success: false, message: "Sheet " + CONFIG.SHEET_KENDALA + " tidak ditemukan." };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: "Data kendala kosong." };

    var headers = data[0];
    var idCol     = headers.indexOf("ID");
    var statusCol = headers.indexOf("STATUS");
    var updateCol = headers.indexOf("LAST_UPDATE");

    if (idCol === -1) return { success: false, message: "Kolom ID tidak ditemukan di sheet " + CONFIG.SHEET_KENDALA + "." };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(id).trim()) {
        if (statusCol !== -1) sh.getRange(i + 1, statusCol + 1).setValue(status);
        if (updateCol !== -1) sh.getRange(i + 1, updateCol + 1).setValue(new Date());
        SpreadsheetApp.flush();
        return { success: true };
      }
    }

    return { success: false, message: "ID kendala tidak ditemukan." };
  } catch(e) {
    console.error("updateKendalaStatus error: " + e.message);
    return { success: false, message: e.message };
  }
}

function deleteKendala(id) {
  try {
    if (!id) return { success: false, message: "ID wajib diisi." };

    var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_KENDALA);
    if (!sh) return { success: false, message: "Sheet " + CONFIG.SHEET_KENDALA + " tidak ditemukan." };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: "Data kendala kosong." };

    var headers = data[0];
    var idCol = headers.indexOf("ID");
    if (idCol === -1) return { success: false, message: "Kolom ID tidak ditemukan." };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]).trim() === String(id).trim()) {
        sh.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }

    return { success: false, message: "ID kendala tidak ditemukan." };
  } catch(e) {
    console.error("deleteKendala error: " + e.message);
    return { success: false, message: e.message };
  }
}

function getKendalaStats() {
  try {
    var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_KENDALA);
    if (!sh) return { total: 0, byStatus: {}, byPenyebab: {} };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { total: 0, byStatus: {}, byPenyebab: {} };

    var headers = data[0];
    var statusCol  = headers.indexOf("STATUS");
    var penyebabCol = headers.indexOf("PENYEBAB");

    var stats = { total: 0, byStatus: {}, byPenyebab: {} };

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      stats.total++;
      var st = statusCol !== -1 ? (row[statusCol] || "Open") : "Open";
      stats.byStatus[st] = (stats.byStatus[st] || 0) + 1;
      if (penyebabCol !== -1 && row[penyebabCol]) {
        var py = row[penyebabCol];
        stats.byPenyebab[py] = (stats.byPenyebab[py] || 0) + 1;
      }
    }

    return stats;
  } catch(e) {
    console.error("getKendalaStats error: " + e.message);
    return { total: 0, byStatus: {}, byPenyebab: {} };
  }
}

function getAllKendala(filters, page, pageSize) {
  filters = filters || {};
  try {
    var sh = getSpreadsheet().getSheetByName(CONFIG.SHEET_KENDALA);
    if (!sh) return { kendala: [], totalCount: 0, totalPages: 0 };

    var data = sh.getDataRange().getValues();
    if (data.length < 2) {
      if (pageSize && page) return { kendala: [], totalCount: 0, totalPages: 0, currentPage: page };
      return [];
    }

    var headers = data[0];
    var userMap = buildUserMap();

    var taskTitleMap = {};
    try {
      var ts = readSheet(CONFIG.SHEET_TASKS);
      ts.rows.forEach(function(row) {
        if (row[ts.col.TASK_ID]) {
          taskTitleMap[String(row[ts.col.TASK_ID]).trim()] = row[ts.col.JUDUL] || "";
        }
      });
    } catch(e) {}

    var list = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = row[idx]; });

      if (obj.CREATED_AT instanceof Date) {
        obj.CREATED_AT = _formatDate(obj.CREATED_AT, "yyyy-MM-dd HH:mm:ss");
      }
      if (obj.LAST_UPDATE instanceof Date) {
        obj.LAST_UPDATE = _formatDate(obj.LAST_UPDATE, "yyyy-MM-dd HH:mm:ss");
      }

      obj.TASK_JUDUL = taskTitleMap[String(obj.TASK_ID).trim()] || "";

      if (filters.status && obj.STATUS !== filters.status) continue;
      if (filters.taskId && String(obj.TASK_ID).trim() !== String(filters.taskId).trim()) continue;
      if (filters.penyebab && obj.PENYEBAB !== filters.penyebab) continue;
      if (filters.search) {
        var s = _normalizeSearchTerm(filters.search);
        var inTaskId    = String(obj.TASK_ID || "").toLowerCase().indexOf(s) !== -1;
        var inKendala   = (obj.KENDALA || "").toLowerCase().indexOf(s) !== -1;
        var inTaskJudul = (obj.TASK_JUDUL || "").toLowerCase().indexOf(s) !== -1;
        var inPenyebab  = (obj.PENYEBAB || "").toLowerCase().indexOf(s) !== -1;
        if (!inTaskId && !inKendala && !inTaskJudul && !inPenyebab) continue;
      }

      list.push(obj);
    }

    list.sort(function(a, b) { return new Date(b.CREATED_AT) - new Date(a.CREATED_AT); });

    if (pageSize && pageSize > 0 && page && page > 0) {
      var totalCount = list.length;
      var totalPages = Math.ceil(totalCount / pageSize);
      var start = (page - 1) * pageSize;
      return {
        kendala: list.slice(start, start + pageSize),
        totalCount: totalCount,
        totalPages: totalPages,
        currentPage: page,
        pageSize: pageSize
      };
    }

    return list;
  } catch(e) {
    console.error("getAllKendala error: " + e.message);
    return { error: e.message };
  }
}

// ==================== CONFIG ====================
function getAppConfig() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('CONFIG');
  if (!sheet) return { KATEGORI: [], PRIORITAS: [], STATUS: [], PENYEBAB_KENDALA: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { KATEGORI: [], PRIORITAS: [], STATUS: [], PENYEBAB_KENDALA: [] };

  var headers = data[0];
  var config = { KATEGORI: [], PRIORITAS: [], STATUS: [], PENYEBAB_KENDALA: [] };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var type  = row[0];
    var value = row[1];
    var label = row[2] || value;
    var order = row[3] || 999;

    if (type === 'KATEGORI')         config.KATEGORI.push({ value: value, label: label, order: order });
    else if (type === 'PRIORITAS')   config.PRIORITAS.push({ value: value, label: label, order: order });
    else if (type === 'STATUS')      config.STATUS.push({ value: value, label: label, order: order });
    else if (type === 'PENYEBAB_KENDALA') config.PENYEBAB_KENDALA.push({ value: value, label: label, order: order });
  }

  config.KATEGORI.sort(function(a,b) { return a.order - b.order; });
  config.PRIORITAS.sort(function(a,b) { return a.order - b.order; });
  config.STATUS.sort(function(a,b) { return a.order - b.order; });
  config.PENYEBAB_KENDALA.sort(function(a,b) { return a.order - b.order; });

  return config;
}

// ==================== FILTERED STATS ====================
function getFilteredStats(filters) {
  try {
    var sheet = getSheet(CONFIG.SHEET_TASKS);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var statusCol   = headers.indexOf('STATUS');
    var priorityCol = headers.indexOf('PRIORITY');
    var dueDateCol  = headers.indexOf('DUE_DATE');
    var assignToCol = headers.indexOf('ASSIGN_TO');
    var deletedCol  = headers.indexOf('IS_DELETED');

    var tasks = data.slice(1);
    if (deletedCol !== -1) {
      tasks = tasks.filter(function(row) { return row[deletedCol] !== true; });
    }
    if (filters && filters.assignee && filters.assignee !== '') {
      var assigneeId = String(filters.assignee).trim();
      tasks = tasks.filter(function(row) {
        return String(row[assignToCol]).trim() === assigneeId;
      });
    }

    var total = tasks.length;
    var byStatus   = { 'To Do': 0, 'In Progress': 0, 'Done': 0, 'Closed': 0 };
    var byPriority = { 'Urgent': 0, 'High': 0, 'Medium': 0, 'Low': 0 };
    var overdue = 0;
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    tasks.forEach(function(row) {
      var status   = row[statusCol];
      var priority = row[priorityCol];
      var dueDate  = row[dueDateCol] ? new Date(row[dueDateCol]) : null;
      if (status)   { byStatus[status]   = (byStatus[status]   || 0) + 1; }
      if (priority) { byPriority[priority] = (byPriority[priority] || 0) + 1; }
      if (dueDate && dueDate < today && status !== 'Done' && status !== 'Closed') overdue++;
    });

    return { total: total, byStatus: byStatus, byPriority: byPriority, overdue: overdue };
  } catch(e) {
    console.error("getFilteredStats error: " + e.message);
    return { total: 0, byStatus: {}, byPriority: {}, overdue: 0 };
  }
}

function getTotalJumlahData(assigneeId) {
  try {
    var sheet = getSheet(CONFIG.SHEET_TASKS);
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return 0;

    var headers     = data[0];
    var idxJumlah   = headers.indexOf('JUMLAH_DATA');
    var idxDeleted  = headers.indexOf('IS_DELETED');
    var idxAssignTo = headers.indexOf('ASSIGN_TO');

    if (idxJumlah === -1) { console.warn("Kolom JUMLAH_DATA tidak ditemukan"); return 0; }

    var total = 0;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (idxDeleted !== -1 && row[idxDeleted] === true) continue;
      if (assigneeId && assigneeId !== '') {
        if (idxAssignTo === -1) continue;
        if (String(row[idxAssignTo]).trim() !== String(assigneeId).trim()) continue;
      }
      var val = Number(row[idxJumlah]);
      if (!isNaN(val)) total += val;
    }
    return total;
  } catch(e) {
    console.error("getTotalJumlahData error: " + e.message);
    return 0;
  }
}

// ==================== DRIVE UPLOAD ====================
var DRIVE_CONFIG = {
  FOLDER_ID: "1bOfIXr7WkicRZu2CeJtB6j_u8fE_dxcO"
};

function uploadFileToDrive(fileData) {
  try {
    if (!fileData || !fileData.data || !fileData.fileName) {
      throw new Error("Data file tidak lengkap.");
    }
    var folder;
    try {
      folder = DriveApp.getFolderById(DRIVE_CONFIG.FOLDER_ID);
    } catch (e) {
      console.warn("Folder tidak ditemukan, menggunakan root folder.");
      folder = DriveApp.getRootFolder();
    }
    var base64Data = fileData.data;
    if (base64Data.indexOf(',') !== -1) base64Data = base64Data.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), fileData.mimeType, fileData.fileName);
    var file = folder.createFile(blob);
    return { success: true, url: file.getUrl() };
  } catch (error) {
    console.error("Error uploadFileToDrive: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

function deleteDriveFile(fileUrl) {
  try {
    var match = fileUrl.match(/[-\w]{25,}/);
    if (!match) throw new Error("Tidak dapat mengekstrak ID file dari URL.");
    var file = DriveApp.getFileById(match[0]);
    file.setTrashed(true);
    return { success: true };
  } catch (error) {
    console.error("Error deleteDriveFile: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ==================== PRIVATE HELPERS ====================
function _findRowIndex(rows, colIndex, targetId) {
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][colIndex]) === String(targetId)) return i;
  }
  return -1;
}

function _safeGetEmail() {
  try {
    return Session.getActiveUser().getEmail() || "system";
  } catch (e) {
    return "system";
  }
}

function generateTaskId(sheet, userObj, baseLastRow, offset, baseTimestamp) {
  var lastRow = (baseLastRow !== undefined) ? baseLastRow : sheet.getLastRow();
  var extra = offset || 0;
  var nextNo = lastRow < 2 ? 1 + extra : lastRow + extra;

  var base = (baseTimestamp instanceof Date) ? baseTimestamp : new Date();
  var rowDate = _shiftDate(base, extra);
  var timestamp = Utilities.formatDate(rowDate, _getSpreadsheetTimeZone(), "yyyyMMddHHmmss");

  var nama = (userObj && userObj.NAMA_LENGKAP) ? userObj.NAMA_LENGKAP.trim() : "";
  var inisial = nama.substring(0, 3).toUpperCase();
  if (inisial === "") inisial = "USR";
  return [nextNo, timestamp, inisial].join("-");
}

// ==================== DEBUG / UTILITIES ====================
function checkSheetStructure() {
  var result = {};
  [CONFIG.SHEET_TASKS, CONFIG.SHEET_USERS, CONFIG.SHEET_COMMENTS, CONFIG.SHEET_KENDALA].forEach(function(name) {
    try {
      var sheet = getSpreadsheet().getSheetByName(name);
      if (!sheet) { result[name] = { error: "Sheet tidak ditemukan" }; return; }
      var headers = sheet.getDataRange().getValues()[0];
      result[name] = { headers: headers, totalDataRows: sheet.getLastRow() - 1 };
    } catch (e) {
      result[name] = { error: e.message };
    }
  });
  return result;
}

function runSmokeTest() {
  Logger.log("=== SMOKE TEST ===");
  Logger.log("Sheet structure: " + JSON.stringify(checkSheetStructure()));
  Logger.log("Users: " + JSON.stringify(getUsers()));
  Logger.log("Current user: " + JSON.stringify(getCurrentUser()));
  Logger.log("Tasks (no filter): " + JSON.stringify(getTasks()));
  Logger.log("Task stats: " + JSON.stringify(getTaskStats()));
  Logger.log("Kendala stats: " + JSON.stringify(getKendalaStats()));
  Logger.log("=== DONE ===");
}

function runBulkUploadSmokeTest() {
  Logger.log("=== BULK UPLOAD SMOKE TEST ===");
  var dummy = [
    { JUDUL: "Test Bulk 1", KATEGORI: "Daily Task", DESKRIPSI: "desc 1", JUMLAH_DATA: 5, ASSIGN_TO: "", STATUS: "To Do", PRIORITY: "Low", DUE_DATE: "01/08/2026", ATTACHMENT_LINK: "" },
    { JUDUL: "Test Bulk 2 (cek ID +1 detik)", KATEGORI: "Daily Task", DESKRIPSI: "desc 2", JUMLAH_DATA: 3, ASSIGN_TO: "", STATUS: "To Do", PRIORITY: "Low", DUE_DATE: "01/08/2026", ATTACHMENT_LINK: "" },
    { JUDUL: "", KATEGORI: "Daily Task", DESKRIPSI: "desc invalid (judul kosong)", JUMLAH_DATA: 3, ASSIGN_TO: "", STATUS: "To Do", PRIORITY: "Low", DUE_DATE: "01/08/2026", ATTACHMENT_LINK: "" },
    { JUDUL: "Test kategori invalid", KATEGORI: "Kategori Ngasal", DESKRIPSI: "desc", JUMLAH_DATA: 1, ASSIGN_TO: "", STATUS: "To Do", PRIORITY: "Low", DUE_DATE: "01/08/2026", ATTACHMENT_LINK: "" },
    { JUDUL: "Test assign_to invalid", KATEGORI: "Daily Task", DESKRIPSI: "desc", JUMLAH_DATA: 1, ASSIGN_TO: "Nama Tidak Ada", STATUS: "To Do", PRIORITY: "Low", DUE_DATE: "01/08/2026", ATTACHMENT_LINK: "" }
  ];
  Logger.log("bulkAddTasks result: " + JSON.stringify(bulkAddTasks(dummy)));
  Logger.log("=== DONE ===");
}


// Final overrides: placed after legacy declarations so V8 resolves optimized versions.
function _pmGetUsersFinal() {
  var cached = _pmCacheGet(PM_USERS_CACHE_KEY);
  if (cached) return cached;
  try {
    var source = _pmReadUsedRange(CONFIG.SHEET_USERS);
    var required = ["USER_ID", "NAMA_LENGKAP", "ROLE", "EMAIL"];
    for (var r = 0; r < required.length; r++) {
      if (source.col[required[r]] === undefined) {
        throw new Error("Kolom wajib tidak ditemukan di USERS: " + required[r]);
      }
    }
    var result = source.rows.map(function(row) {
      return {
        USER_ID: row[source.col.USER_ID],
        NAMA_LENGKAP: row[source.col.NAMA_LENGKAP] || "",
        ROLE: row[source.col.ROLE] || "user",
        EMAIL: row[source.col.EMAIL] || ""
      };
    }).filter(function(user) {
      return user.USER_ID !== "" && user.USER_ID !== null;
    });
    _pmCachePut(PM_USERS_CACHE_KEY, result);
    return result;
  } catch (e) {
    console.error("_pmGetUsersFinal: " + e.message);
    return [];
  }
}

function getUsers() {
  return _pmGetUsersFinal();
}

function _pmGetAppConfigFinal() {
  var cached = _pmCacheGet(PM_CONFIG_CACHE_KEY);
  if (cached) return cached;

  var empty = { KATEGORI: [], PRIORITAS: [], STATUS: [], PENYEBAB_KENDALA: [] };
  try {
    var sheet = getSpreadsheet().getSheetByName("CONFIG");
    if (!sheet) return empty;
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return empty;

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var config = { KATEGORI: [], PRIORITAS: [], STATUS: [], PENYEBAB_KENDALA: [] };

    for (var i = 1; i < data.length; i++) {
      var type = data[i][0];
      var value = data[i][1];
      if (!type || value === "" || value === null) continue;
      var item = {
        value: value,
        label: data[i][2] || value,
        order: data[i][3] || 999
      };
      if (type === "KATEGORI") config.KATEGORI.push(item);
      else if (type === "PRIORITAS") config.PRIORITAS.push(item);
      else if (type === "STATUS") config.STATUS.push(item);
      else if (type === "PENYEBAB_KENDALA") config.PENYEBAB_KENDALA.push(item);
    }

    Object.keys(config).forEach(function(key) {
      config[key].sort(function(a, b) { return a.order - b.order; });
    });
    _pmCachePut(PM_CONFIG_CACHE_KEY, config);
    return config;
  } catch (e) {
    console.error("_pmGetAppConfigFinal: " + e.message);
    return empty;
  }
}

function getAppConfig() {
  return _pmGetAppConfigFinal();
}
