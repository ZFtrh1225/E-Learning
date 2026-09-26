/*************************************************************
 * EduFSR — E-Learning Platform  |  BACKEND API (Code.gs)
 * ------------------------------------------------------------
 * Versi 3.0 — Maximized for Independent Learning
 * Fokus: Siswa mandiri + Admin full control user
 *
 * PERUBAHAN UTAMA v3:
 *  - Login siswa HANYA jika sudah didaftarkan admin (no auto-register)
 *  - CRUD Siswa lengkap oleh admin
 *  - Scaffolding / Mode Belajar Bertahap
 *  - Streak harian + Badge
 *  - Antrian soal salah (retry wrong)
 *  - Analisis kelemahan per kategori
 *  - Tingkat kesulitan soal
 *  - Flashcard
 *  - Tujuan mingguan
 *  - Statistik admin lebih kaya
 *  - Focus mode support + better meta
 *
 * CARA PAKAI:
 *  1. Buat Google Spreadsheet baru (kosong).
 *  2. Extensions > Apps Script → timpa Code.gs dengan file ini.
 *  3. Jalankan initSheets() sekali (Run) → izinkan akses Spreadsheet & Drive.
 *  4. Jika error "Spreadsheet tidak ditemukan":
 *       setSpreadsheetId("ID_SPREADSHEET_ANDA")
 *  5. Deploy > New deployment > Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *  6. Salin URL Web App (…/exec) → tempel di frontend (API_URL).
 *  7. Jalankan migrateLegacyPasswordsFromEditor() dan
 *     resetAdminPasswordFromEditor() dari editor; simpan password baru.
 *  8. Untuk akun yang sudah memakai PBKDF2 lama, jalankan
 *     migratePasswordVerifiersFromEditor() dari editor hingga tersisa 0
 *     agar login pertama setelah pembaruan juga cepat.
 *
 * Setiap perubahan kode → Manage deployments > Edit > New version.
 *************************************************************/

const SHEET_USERS     = 'Users';
const SHEET_KATEGORI  = 'Kategori';
const SHEET_MATERI    = 'Materi';
const SHEET_SOAL      = 'Soal';
const SHEET_HASIL     = 'Hasil';
const SHEET_BELAJAR   = 'Belajar';
const SHEET_WRONG     = 'WrongQueue';
const SHEET_FLASHCARD = 'Flashcard';
const SHEET_BADGE     = 'Badge';
const MATERI_FOLDER_NAME = 'EduFSR - File Materi';
const UPLOAD_TEMP_FOLDER_NAME = 'EduFSR - Upload Temp';
const DEFAULT_KKM = 70;
const MAX_UPLOAD_BYTES = 35 * 1024 * 1024; // 35MB
const PASSWORD_ITERATIONS_ = 120000;
const ADMIN_RESET_REQUIRED_ = '!ADMIN_RESET_REQUIRED!';

/* ============================================================
 * CORS + JSON RESPONSE HELPERS
 * ==========================================================*/
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
 * ENTRY POINTS — API ROUTER
 * ==========================================================*/
function doGet(e) {
  try {
    // GET is a health check only. Never process protected actions through URL parameters.
    return jsonResponse_({ success: true, message: 'EduFSR API ready' });
  } catch (err) {
    return jsonResponse_({ success: false, message: err.message || String(err) });
  }
}

function doPost(e) {
  // A Spreadsheet object is reused only during this single API execution.
  WEB_REQUEST_ACTIVE_ = true;
  REQUEST_SS_ = null;
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        return jsonResponse_({ success: false, message: 'Body harus JSON valid.' });
      }
    }
    const action = payload.action || '';
    if (!action) {
      return jsonResponse_({ success: false, message: 'Parameter "action" wajib diisi.' });
    }
    if (action === 'login') return routeAction_(action, payload);
    if (action === 'loginChallenge') return routeAction_(action, payload);
    const session = requireSession_(payload.authToken);
    authorizeAction_(action, payload, session);
    if (action === 'logout') {
      const found = findRowByValue_(SHEET_USERS, 'Username', session.username);
      const epochCol = found && found.headers.indexOf('AuthEpoch');
      if (found && epochCol >= 0) getSS().getSheetByName(SHEET_USERS)
        .getRange(found.rowNumber, epochCol + 1).setValue(Number(found.values[epochCol] || 0) + 1);
      return jsonResponse_({ success: true });
    }
    return routeAction_(action, payload);
  } catch (err) {
    return jsonResponse_({ success: false, message: err.message || String(err) });
  } finally {
    REQUEST_SS_ = null;
    WEB_REQUEST_ACTIVE_ = false;
  }
}

function routeAction_(action, data) {
  switch (String(action)) {
    // Auth
    case 'login':
      return jsonResponse_(login(data.username, data.password, data.passwordProof));
    case 'loginChallenge':
      return jsonResponse_(getLoginChallenge_(data.username));

    // Student management (Admin only)
    case 'getStudentList':
      return jsonResponse_(getStudentList());
    case 'addStudent':
      return jsonResponse_(addStudent(data.data || data));
    case 'updateStudent':
      return jsonResponse_(updateStudent(data.data || data));
    case 'deleteStudent':
      return jsonResponse_(deleteStudent(data.username || (data.data && data.data.username)));
    case 'resetStudentPassword':
      return jsonResponse_(resetStudentPassword(data.data || data));

    // Materi
    case 'getKategoriList':
      return jsonResponse_(getKategoriList());
    case 'getKategoriFull':
      return jsonResponse_(getKategoriFull());
    case 'saveKategori':
      return jsonResponse_(saveKategori(data.data || data));
    case 'deleteKategori':
      return jsonResponse_(deleteKategori(data.id || data.nama));
    case 'getMateriByKategori':
      return jsonResponse_(getMateriByKategori(data.kategori));
    case 'getMateriList':
      return jsonResponse_(getMateriList());
    case 'getMateriById':
      return jsonResponse_(getMateriById(data.id));
    case 'getMateriForLatihan':
      return jsonResponse_(getMateriForLatihan());
    case 'getKategoriForLatihan':
      return jsonResponse_(getKategoriForLatihan());
    case 'saveMateri':
      return jsonResponse_(saveMateri(data.data || data));
    case 'initChunkUpload':
      return jsonResponse_(initChunkUpload(data.data || data));
    case 'uploadChunk':
      return jsonResponse_(uploadChunk(data.data || data));
    case 'finalizeChunkUpload':
      return jsonResponse_(finalizeChunkUpload(data.data || data));
    case 'abortChunkUpload':
      return jsonResponse_(abortChunkUpload(data.sessionId || (data.data && data.data.sessionId)));
    case 'deleteMateri':
      return jsonResponse_(deleteMateri(data.id));
    case 'duplicateMateri':
      return jsonResponse_(duplicateMateri(data.id));

    // Soal
    case 'getSoalByKategoriAdmin':
      return jsonResponse_(getSoalByKategoriAdmin(data.kategori));
    case 'getSoalByKategoriStudent':
      return jsonResponse_(getSoalByKategoriStudent(data.kategori));
    case 'saveSoal':
      return jsonResponse_(saveSoal(data.data || data));
    case 'importSoalBulk':
      return jsonResponse_(importSoalBulk(data.data || data));
    case 'deleteSoal':
      return jsonResponse_(deleteSoal(data.id));
    case 'toggleSoalStatus':
      return jsonResponse_(toggleSoalStatus(data.id));
    case 'duplicateSoal':
      return jsonResponse_(duplicateSoal(data.id));
    case 'getSoalByMateriAdmin':
      return jsonResponse_(getSoalByMateriAdmin(data.materiId || data.id));
    case 'getSoalByMateriStudent':
      return jsonResponse_(getSoalByMateriStudent(data.materiId || data.id));
    case 'setLearningStage':
      return jsonResponse_({ success: false, message: 'Tahap belajar berubah setelah aktivitas belajar.' });

    // Hasil / Nilai
    case 'submitJawaban':
      return jsonResponse_(submitJawaban(data.payload || data));
    case 'gradeEssay':
      return jsonResponse_(gradeEssay(data.data || data));
    case 'getHasilById':
      return jsonResponse_(getHasilById(data.id));
    case 'getHasilByStudent':
      return jsonResponse_(getHasilByStudent(data.username));
    case 'getAllHasil':
      return jsonResponse_(getAllHasil());
    case 'exportHasil':
      return jsonResponse_(exportHasil(data.username || ''));

    // Progress & Belajar
    case 'getProgressByStudent':
      return jsonResponse_(getProgressByStudent(data.username));
    case 'getStudentDashboard':
      return jsonResponse_(getStudentDashboard(data.username));
    case 'getBelajarMeta':
      return jsonResponse_(getBelajarMeta(data.username));
    case 'markMateriDibaca':
      return jsonResponse_(markMateriDibaca(data.data || data));
    case 'toggleBookmark':
      return jsonResponse_(toggleBookmark(data.data || data));
    case 'saveCatatanMateri':
      return jsonResponse_(saveCatatanMateri(data.data || data));
    case 'getCekPemahaman':
      return jsonResponse_(getCekPemahaman(data.materiId || data.id));
    case 'startQuizAttempt':
      return jsonResponse_(startQuizAttempt(data.data || data));
    case 'getLearningPath':
      return jsonResponse_(getLearningPath(data.materiId || data.id, data.username));
    case 'getMateriReader':
      return jsonResponse_(getMateriReader(data.materiId || data.id, data.username));

    // Latihan list
    case 'getLatihanList':
      return jsonResponse_(getLatihanList());

    // Settings
    case 'getSettings':
      return jsonResponse_(getSettings());
    case 'saveSettings':
      return jsonResponse_(saveSettings(data.data || data));

    // Streak, Badge, Goals
    case 'getStudentProfile':
      return jsonResponse_(getStudentProfile(data.username));
    case 'recordActivity':
      return jsonResponse_(recordActivity(data.data || data));
    case 'getBadges':
      return jsonResponse_(getBadges(data.username));
    case 'setWeeklyGoal':
      return jsonResponse_(setWeeklyGoal(data.data || data));

    // Wrong queue (retry)
    case 'getWrongQueue':
      return jsonResponse_(getWrongQueue(data.username));
    case 'addToWrongQueue':
      return jsonResponse_({ success: false, message: 'Tindakan internal.' });
    case 'clearWrongItem':
      return jsonResponse_(clearWrongItem(data.data || data));
    case 'getRetryQuiz':
      return jsonResponse_(getRetryQuiz(data.username));

    // Weakness analysis
    case 'getWeaknessAnalysis':
      return jsonResponse_(getWeaknessAnalysis(data.username));

    // Flashcard
    case 'getFlashcards':
      return jsonResponse_(getFlashcards(data.username, data.materiId || ''));
    case 'getFlashcardOverview':
      return jsonResponse_(getFlashcardOverview(data.username));
    case 'saveFlashcard':
      return jsonResponse_(saveFlashcard(data.data || data));
    case 'deleteFlashcard':
      return jsonResponse_(deleteFlashcard(data.id, data._actor));
    case 'generateFlashcardsFromMateri':
      return jsonResponse_(generateFlashcardsFromMateri(data.data || data));
    case 'reviewFlashcard':
      return jsonResponse_(reviewFlashcard(data.data || data));

    // Admin
    case 'getAdminStats':
      return jsonResponse_(getAdminStats());
    case 'getRichStats':
      return jsonResponse_(getRichStats());

    // Setup
    case 'initSheets':
      return jsonResponse_({ success: false, message: 'Jalankan setup dari editor Apps Script.' });
    case 'setSpreadsheetId':
      return jsonResponse_({ success: false, message: 'Jalankan setup dari editor Apps Script.' });

    default:
      return jsonResponse_({ success: false, message: 'Action tidak dikenal: ' + action });
  }
}

/* ============================================================
 * SPREADSHEET ACCESS
 * ==========================================================*/
let REQUEST_SS_ = null;
let WEB_REQUEST_ACTIVE_ = false;
function getSS() {
  if (WEB_REQUEST_ACTIVE_ && REQUEST_SS_) return REQUEST_SS_;
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }

  if (!ss) {
    const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (id) {
      try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
    }
  }

  if (!ss) {
    throw new Error(
      'Spreadsheet tidak ditemukan. Jalankan fungsi setSpreadsheetId("ID_SPREADSHEET_ANDA") ' +
      'sekali dari editor Apps Script (ID ada di URL spreadsheet, antara /d/ dan /edit), ' +
      'lalu buat deployment versi baru.'
    );
  }
  if (WEB_REQUEST_ACTIVE_) REQUEST_SS_ = ss;
  return ss;
}

function setSpreadsheetId(id) {
  if (!id) throw new Error('ID Spreadsheet wajib diisi. Jalankan dari editor Apps Script.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  initSheets();
  return 'Tersimpan. Spreadsheet ID: ' + id;
}

function getMateriFolder_() {
  const folders = DriveApp.getFoldersByName(MATERI_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(MATERI_FOLDER_NAME);
}

/* ============================================================
 * SETUP / INIT
 * ==========================================================*/
function initSheets() {
  signingKey_();
  const ss = getSS();

  // Users
  let sh = ss.getSheetByName(SHEET_USERS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_USERS);
    sh.appendRow(['Username', 'Password', 'Role', 'NamaLengkap', 'TanggalDaftar', 'LastActive', 'Streak', 'LongestStreak', 'WeeklyGoal', 'WeeklyProgress', 'WeeklyWeek', 'AuthEpoch']);
    sh.appendRow(['admin', ADMIN_RESET_REQUIRED_, 'admin', 'Administrator', new Date(), '', 0, 0, 0, 0]);
    sh.getRange(1, 1, 1, 10).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ensureColumns_(SHEET_USERS, [
    { name: 'LastActive', value: '' },
    { name: 'Streak', value: 0 },
    { name: 'LongestStreak', value: 0 },
    { name: 'WeeklyGoal', value: 0 },
    { name: 'WeeklyProgress', value: 0 },
    { name: 'WeeklyWeek', value: '' },
    { name: 'AuthEpoch', value: 0 }
  ]);

  // Kategori
  sh = ss.getSheetByName(SHEET_KATEGORI);
  if (!sh) {
    sh = ss.insertSheet(SHEET_KATEGORI);
    sh.appendRow(['ID', 'Nama', 'Deskripsi', 'Tanggal']);
    sh.getRange(1, 1, 1, 4).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  // Migrasi kategori orphan
  try {
    const existingCats = {};
    sheetToObjects_(SHEET_KATEGORI).forEach(c => {
      const n = String(c.Nama || '').trim();
      if (n) existingCats[n.toLowerCase()] = true;
    });
    const matSh = ss.getSheetByName(SHEET_MATERI);
    if (matSh) {
      const matRows = sheetToObjects_(SHEET_MATERI);
      const toAdd = {};
      matRows.forEach(m => {
        const n = String(m.Kategori || '').trim();
        if (n && !existingCats[n.toLowerCase()]) toAdd[n] = true;
      });
      Object.keys(toAdd).forEach(nama => {
        ss.getSheetByName(SHEET_KATEGORI).appendRow([generateId_(), nama, '', new Date()]);
        existingCats[nama.toLowerCase()] = true;
      });
    }
  } catch (e) {}

  // Materi
  sh = ss.getSheetByName(SHEET_MATERI);
  if (!sh) {
    sh = ss.insertSheet(SHEET_MATERI);
    sh.appendRow(['ID', 'Kategori', 'Judul', 'FileID', 'FileName', 'FileType', 'FileURL', 'Tanggal', 'KKM', 'Ringkasan', 'Urutan']);
    sh.getRange(1, 1, 1, 11).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ensureColumns_(SHEET_MATERI, [
    { name: 'KKM', value: DEFAULT_KKM },
    { name: 'Ringkasan', value: '' },
    { name: 'Urutan', value: 0 }
  ]);

  // Soal
  sh = ss.getSheetByName(SHEET_SOAL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SOAL);
    sh.appendRow(['ID', 'Kategori', 'Tipe', 'Pertanyaan', 'PilihanA', 'PilihanB', 'PilihanC', 'PilihanD', 'JawabanBenar', 'Status', 'MateriID', 'Kesulitan', 'Tags', 'Pembahasan', 'GambarFileID', 'GambarURL']);
    sh.getRange(1, 1, 1, 16).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ensureColumns_(SHEET_SOAL, [
    { name: 'Status', value: 'Aktif' },
    { name: 'MateriID', value: '' },
    { name: 'Kesulitan', value: 'Sedang' },
    { name: 'Tags', value: '' },
    { name: 'Pembahasan', value: '' },
    { name: 'GambarFileID', value: '' },
    { name: 'GambarURL', value: '' }
  ]);

  // Hasil
  sh = ss.getSheetByName(SHEET_HASIL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_HASIL);
    sh.appendRow(['ID', 'Username', 'MateriID', 'JudulMateri', 'Skor', 'BenarPG', 'TotalPG', 'JawabanEssay', 'Tanggal', 'Lulus', 'EssaySkor', 'SkorAkhir', 'CatatanAdmin', 'EssayStatus', 'KKM', 'TipeLatihan', 'AttemptID', 'BobotPG', 'BobotEssay']);
    sh.getRange(1, 1, 1, 14).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ensureColumns_(SHEET_HASIL, [
    { name: 'Lulus', value: '' },
    { name: 'EssaySkor', value: '' },
    { name: 'SkorAkhir', value: '' },
    { name: 'CatatanAdmin', value: '' },
    { name: 'EssayStatus', value: '' },
    { name: 'KKM', value: DEFAULT_KKM },
    { name: 'TipeLatihan', value: 'latihan' },
    { name: 'AttemptID', value: '' },
    { name: 'BobotPG', value: '' },
    { name: 'BobotEssay', value: '' }
  ]);

  // Belajar
  sh = ss.getSheetByName(SHEET_BELAJAR);
  if (!sh) {
    sh = ss.insertSheet(SHEET_BELAJAR);
    sh.appendRow(['ID', 'Username', 'MateriID', 'Dibaca', 'TanggalBaca', 'Bookmark', 'Catatan', 'UpdatedAt', 'Stage']);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ensureColumns_(SHEET_BELAJAR, [{ name: 'Stage', value: 'baca' }]);

  // Wrong Queue (soal yang salah untuk diulang)
  sh = ss.getSheetByName(SHEET_WRONG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_WRONG);
    sh.appendRow(['ID', 'Username', 'SoalID', 'Kategori', 'MateriID', 'Pertanyaan', 'JawabanBenar', 'JawabanUser', 'AddedAt', 'RetryCount']);
    sh.getRange(1, 1, 1, 10).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  // Flashcard
  sh = ss.getSheetByName(SHEET_FLASHCARD);
  if (!sh) {
    sh = ss.insertSheet(SHEET_FLASHCARD);
    sh.appendRow(['ID', 'Username', 'MateriID', 'Kategori', 'Front', 'Back', 'CreatedAt', 'ReviewStatus', 'ReviewCount', 'LastReviewedAt']);
    sh.getRange(1, 1, 1, 10).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ensureColumns_(SHEET_FLASHCARD, [
    { name: 'ReviewStatus', value: 'baru' },
    { name: 'ReviewCount', value: 0 },
    { name: 'LastReviewedAt', value: '' }
  ]);

  // Badge log
  sh = ss.getSheetByName(SHEET_BADGE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_BADGE);
    sh.appendRow(['ID', 'Username', 'BadgeCode', 'BadgeName', 'EarnedAt']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  return 'OK — v3.0 initialized';
}

function ensureColumns_(sheetName, columns) {
  const sh = getSS().getSheetByName(sheetName);
  if (!sh) return;
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  columns.forEach(col => {
    if (headers.indexOf(col.name) !== -1) return;
    const newColIdx = sh.getLastColumn() + 1;
    sh.getRange(1, newColIdx).setValue(col.name).setFontWeight('bold');
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const values = [];
      for (let i = 0; i < lastRow - 1; i++) values.push([col.value]);
      sh.getRange(2, newColIdx, lastRow - 1, 1).setValues(values);
    }
  });
}

/* ============================================================
 * HELPERS
 * ==========================================================*/
function sheetToObjects_(sheetName) {
  const sh = getSS().getSheetByName(sheetName);
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map((row, idx) => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      obj._row = idx + 2;
      return obj;
    });
}

function generateId_() {
  return Utilities.getUuid().split('-')[0] + Date.now().toString(36);
}

function toIso_(d) {
  if (!d) return '';
  try {
    if (Object.prototype.toString.call(d) === '[object Date]') return d.toISOString();
    return String(d);
  } catch (e) { return ''; }
}

function findRowByValue_(sheetName, colName, value) {
  const sh = getSS().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const colIdx = headers.indexOf(colName);
  for (let i = 1; i < data.length; i++) {
    if (idEq_(data[i][colIdx], value)) {
      return { rowNumber: i + 1, headers: headers, values: data[i] };
    }
  }
  return null;
}

function idEq_(a, b) {
  return String(a == null ? '' : a).trim() === String(b == null ? '' : b).trim();
}

// SpreadsheetApp interprets strings starting with = as formulas. Keep notes,
// flashcards and question text as literal data, including leading whitespace.
function sheetLiteral_(value) {
  const text = String(value == null ? '' : value);
  return /^\s*=/.test(text) ? "'" + text : text;
}

function todayStr_() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function weekStartStr_() {
  const d = new Date();
  d.setDate(d.getDate() - (d.getDay() + 6) % 7);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

/* ============================================================
 * AUTH — Hanya user yang sudah didaftarkan
 * ==========================================================*/
function passwordDigest_(password, salt, iterations) {
  // PBKDF2-HMAC-SHA256, 32-byte output. The salt is unique for each account.
  const key = Utilities.newBlob(String(password)).getBytes();
  const saltBytes = Utilities.newBlob(String(salt)).getBytes();
  const block = saltBytes.concat([0, 0, 0, 1]);
  let u = Utilities.computeHmacSha256Signature(block, key);
  const derived = u.map(n => n & 255);
  for (let i = 1; i < iterations; i++) {
    u = Utilities.computeHmacSha256Signature(u, key);
    for (let j = 0; j < derived.length; j++) derived[j] ^= (u[j] & 255);
  }
  return Utilities.base64Encode(derived.map(n => n > 127 ? n - 256 : n));
}

function hashPassword_(password) {
  const salt = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  return ['pbkdf2_sha256_peppered', PASSWORD_ITERATIONS_, salt,
    passwordVerifier_(passwordDigest_(password, salt, PASSWORD_ITERATIONS_))].join('$');
}

// The stored verifier needs the server signing key as well as the password-derived
// value. A copy of the Users sheet alone must not be usable as a login proof.
function passwordVerifier_(derivedBase64) {
  return Utilities.base64Encode(Utilities.computeHmacSha256Signature(
    'edufsr-password-v2:' + derivedBase64, signingKey_()));
}

function passwordParts_(stored) {
  const parts = String(stored || '').split('$');
  if (parts.length !== 4 ||
    (parts[0] !== 'pbkdf2_sha256' && parts[0] !== 'pbkdf2_sha256_peppered') ||
    !/^\d+$/.test(parts[1]) || Number(parts[1]) < 10000 || Number(parts[1]) > 600000 ||
    !/^[a-f0-9]{64}$/.test(parts[2]) || !/^[A-Za-z0-9+/]{43}=$/.test(parts[3])) return null;
  return parts;
}

function sameSecret_(left, right) {
  left = String(left || ''); right = String(right || '');
  if (left.length !== right.length) return false;
  let diff = 0;
  for (let i = 0; i < left.length; i++) diff |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return diff === 0;
}

/** Editor-only: convert existing password verifiers without knowing passwords.
 * Repeat when the result reports remaining rows. Existing sessions are revoked. */
function migratePasswordVerifiersFromEditor(maxUsers) {
  const sh = getSS().getSheetByName(SHEET_USERS);
  if (!sh) throw new Error('Jalankan initSheets() dari editor terlebih dahulu.');
  const rows = sh.getDataRange().getValues();
  const col = rows[0].indexOf('Password');
  if (col < 0) throw new Error('Kolom Password tidak ditemukan.');
  const limit = Math.min(100, Math.max(1, Number(maxUsers) || 50));
  let converted = 0, remaining = 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    for (let i = 1; i < rows.length; i++) {
      const stored = String(rows[i][col] || '');
      const parts = passwordParts_(stored);
      if (!parts || parts[0] !== 'pbkdf2_sha256') continue;
      if (converted >= limit) { remaining++; continue; }
      const cell = sh.getRange(i + 1, col + 1);
      if (String(cell.getValue()) !== stored) continue;
      cell.setValue(['pbkdf2_sha256_peppered', parts[1], parts[2], passwordVerifier_(parts[3])].join('$'));
      converted++;
    }
  } finally { lock.releaseLock(); }
  const message = 'Verifikator diperbarui: ' + converted + '; tersisa: ' + remaining +
    '. Sesi lama akan meminta login ulang.';
  Logger.log(message);
  return message;
}

function signingKey_() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('AUTH_SIGNING_KEY');
  if (key) return key;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    key = props.getProperty('AUTH_SIGNING_KEY');
    if (!key) {
      key = (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
      props.setProperty('AUTH_SIGNING_KEY', key);
    }
    return key;
  } finally { lock.releaseLock(); }
}

function signedToken_(payload) {
  const body = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/, '');
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(body, signingKey_())
  ).replace(/=+$/, '');
  return body + '.' + sig;
}

function credentialTag_(passwordRecord) {
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(String(passwordRecord), signingKey_())
  ).replace(/=+$/, '');
}

function readSignedToken_(token, kind) {
  const error = kind === 'session' ? 'Sesi habis. Silakan masuk kembali.'
    : 'Sesi latihan tidak berlaku. Mulai latihan kembali.';
  const input = String(token || '');
  if (input.length > 50000 || !/^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(input)) throw new Error(error);
  const parts = input.split('.');
  const expected = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(parts[0], signingKey_())
  ).replace(/=+$/, '');
  if (expected.length !== parts[1].length) throw new Error(error);
  let diff = 0;
  for (let i = 0; i < expected.length; i++) diff |= expected.charCodeAt(i) ^ parts[1].charCodeAt(i);
  if (diff) throw new Error(error);
  let payload;
  try { payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); }
  catch (e) { throw new Error(error); }
  if (!payload || payload.kind !== kind || !Number.isFinite(payload.exp) || Date.now() >= payload.exp) throw new Error(error);
  return payload;
}

function passwordMatches_(input, stored) {
  if (stored === ADMIN_RESET_REQUIRED_) return false;
  const parts = passwordParts_(stored);
  if (!parts && !String(stored).startsWith('pbkdf2_')) {
    // Legacy passwords remain usable until the editor migration or next login.
    return String(input) === String(stored);
  }
  if (!parts) return false;
  const derived = passwordDigest_(input, parts[2], Number(parts[1]));
  return sameSecret_(parts[0] === 'pbkdf2_sha256_peppered' ? passwordVerifier_(derived) : derived, parts[3]);
}

function passwordProofMatches_(proof, stored) {
  const parts = passwordParts_(stored);
  return !!parts && parts[0] === 'pbkdf2_sha256_peppered' &&
    /^[A-Za-z0-9+/]{43}=$/.test(String(proof || '')) &&
    sameSecret_(passwordVerifier_(String(proof)), parts[3]);
}

function getLoginChallenge_(username) {
  username = String(username || '').trim().toLowerCase().slice(0, 180);
  const user = username && sheetToObjects_(SHEET_USERS).find(u =>
    String(u.Username).toLowerCase() === username);
  const parts = user && passwordParts_(user.Password);
  // A dummy salt keeps the response shape the same for unknown accounts.
  return { salt: parts ? parts[2] : (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''),
    iterations: parts ? Number(parts[1]) : PASSWORD_ITERATIONS_ };
}

/** Run once from the Apps Script editor; never expose this function in doPost. */
function migrateLegacyPasswordsFromEditor(maxUsers) {
  const sh = getSS().getSheetByName(SHEET_USERS);
  if (!sh) throw new Error('Jalankan initSheets() dari editor terlebih dahulu.');
  const rows = sh.getDataRange().getValues();
  const userCol = rows[0].indexOf('Username');
  const passCol = rows[0].indexOf('Password');
  const roleCol = rows[0].indexOf('Role');
  if (userCol < 0 || passCol < 0 || roleCol < 0) throw new Error('Kolom Users tidak lengkap.');
  const limit = Math.min(10, Math.max(1, Number(maxUsers) || 5));
  let migrated = 0, remaining = 0;
  for (let i = 1; i < rows.length; i++) {
    const old = String(rows[i][passCol] || '');
    if (!old || old.startsWith('pbkdf2_sha256$') ||
      old.startsWith('pbkdf2_sha256_peppered$') || old === ADMIN_RESET_REQUIRED_) continue;
    if (migrated >= limit) { remaining++; continue; }
    const isDefaultAdmin = String(rows[i][roleCol]) === 'admin' &&
      String(rows[i][userCol]).toLowerCase() === 'admin' && old === 'admin';
    sh.getRange(i + 1, passCol + 1).setValue(isDefaultAdmin ? ADMIN_RESET_REQUIRED_ : hashPassword_(old));
    const epochCol = rows[0].indexOf('AuthEpoch');
    if (epochCol >= 0) sh.getRange(i + 1, epochCol + 1).setValue(Number(rows[i][epochCol] || 0) + 1);
    migrated++;
  }
  return 'Dimigrasikan: ' + migrated + '; tersisa: ' + remaining + '. Ulangi sampai tersisa 0, lalu jalankan resetAdminPasswordFromEditor().';
}

/** Generates a strong new admin password, shown only in the editor's execution result. */
function resetAdminPasswordFromEditor() {
  const found = findRowByValue_(SHEET_USERS, 'Username', 'admin');
  if (!found || String(found.values[found.headers.indexOf('Role')]) !== 'admin') {
    throw new Error('Akun admin belum dibuat. Jalankan initSheets() dari editor.');
  }
  const password = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  getSS().getSheetByName(SHEET_USERS)
    .getRange(found.rowNumber, found.headers.indexOf('Password') + 1)
    .setValue(hashPassword_(password));
  const epochCol = found.headers.indexOf('AuthEpoch');
  if (epochCol >= 0) getSS().getSheetByName(SHEET_USERS)
    .getRange(found.rowNumber, epochCol + 1).setValue(Number(found.values[epochCol] || 0) + 1);
  Logger.log('Password admin baru (simpan secara aman): ' + password);
  return 'Password admin baru tercatat di Execution log.';
}

function login(username, password, passwordProof) {
  username = (username || '').toString().trim();
  password = (password || '').toString().trim();

  if (!username || (!password && !passwordProof)) {
    return { success: false, message: 'Username dan password wajib diisi.' };
  }
  const cache = CacheService.getScriptCache();
  const attemptKey = 'login_fail_' + username.toLowerCase().slice(0, 180);
  if (Number(cache.get(attemptKey) || 0) >= 8) {
    return { success: false, message: 'Terlalu banyak percobaan. Coba lagi beberapa menit lagi.' };
  }
  const failed = () => cache.put(attemptKey, String(Number(cache.get(attemptKey) || 0) + 1), 900);

  const users = sheetToObjects_(SHEET_USERS);
  const user = users.find(u => String(u.Username).toLowerCase() === username.toLowerCase());

  if (!user) {
    failed();
    return { success: false, message: 'Username atau password tidak sesuai.' };
  }

  if (String(user.Password) === ADMIN_RESET_REQUIRED_ ||
    (String(user.Role) === 'admin' && String(user.Username).toLowerCase() === 'admin' && String(user.Password) === 'admin')) {
    return { success: false, message: 'Akun admin perlu password baru. Jalankan resetAdminPasswordFromEditor() dari editor Apps Script.' };
  }
  const storedPassword = String(user.Password);
  const storedParts = passwordParts_(storedPassword);
  if (passwordProof && (!storedParts || storedParts[0] === 'pbkdf2_sha256')) {
    return { success: false, message: 'Akun memerlukan login lama sekali sebelum mode cepat aktif.' };
  }
  const validPassword = passwordProof
    ? passwordProofMatches_(passwordProof, storedPassword)
    : passwordMatches_(password, storedPassword);
  if (!validPassword) {
    failed();
    return { success: false, message: 'Username atau password tidak sesuai.' };
  }
  cache.remove(attemptKey);
  if (user.AuthEpoch === undefined) {
    return { success: false, message: 'Pembaruan data belum selesai. Jalankan initSheets() dari editor Apps Script.' };
  }

  if (!String(user.Password).startsWith('pbkdf2_sha256_peppered$')) {
    const upgraded = hashPassword_(password);
    const found = findRowByValue_(SHEET_USERS, 'Username', user.Username);
    const passCol = found && found.headers.indexOf('Password');
    if (!found || passCol < 0 || String(found.values[passCol]) !== String(user.Password)) {
      throw new Error('Data akun berubah. Silakan coba masuk kembali.');
    }
    getSS().getSheetByName(SHEET_USERS).getRange(found.rowNumber, passCol + 1).setValue(upgraded);
    user.Password = upgraded;
  }

  // Update last active + streak
  try {
    updateStreakOnLogin_(user.Username);
  } catch (e) {}

  const token = signedToken_({ kind: 'session', username: user.Username,
    role: user.Role || 'student', epoch: Number(user.AuthEpoch || 0),
    credentialTag: credentialTag_(user.Password), exp: Date.now() + 4 * 3600000 });
  return {
    success: true,
    token: token,
    role: user.Role || 'student',
    username: user.Username,
    fullname: user.NamaLengkap || user.Username,
    streak: Number(user.Streak || 0),
    longestStreak: Number(user.LongestStreak || 0)
  };
}

function updateStreakOnLogin_(username) {
  const found = findRowByValue_(SHEET_USERS, 'Username', username);
  if (!found) return;
  const sh = getSS().getSheetByName(SHEET_USERS);
  const headers = found.headers;
  const idx = (name) => headers.indexOf(name);

  const lastActive = found.values[idx('LastActive')];
  let streak = Number(found.values[idx('Streak')] || 0);
  let longest = Number(found.values[idx('LongestStreak')] || 0);
  const today = todayStr_();

  let lastDateStr = '';
  if (lastActive) {
    try {
      const d = new Date(lastActive);
      lastDateStr = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    } catch (e) {}
  }

  if (lastDateStr === today) {
    // already counted today
  } else {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const yStr = yesterday.getFullYear() + '-' + String(yesterday.getMonth() + 1).padStart(2, '0') + '-' + String(yesterday.getDate()).padStart(2, '0');
    if (lastDateStr === yStr) {
      streak += 1;
    } else {
      streak = 1;
    }
    if (streak > longest) longest = streak;
  }

  const setCol = (name, val) => {
    const c = idx(name) + 1;
    if (c > 0) sh.getRange(found.rowNumber, c).setValue(val);
  };
  setCol('LastActive', new Date());
  setCol('Streak', streak);
  setCol('LongestStreak', longest);

  // Award streak badges
  if (streak >= 3) awardBadge_(username, 'streak3', 'Konsisten 3 Hari');
  if (streak >= 7) awardBadge_(username, 'streak7', 'Konsisten 7 Hari');
  if (streak >= 14) awardBadge_(username, 'streak14', 'Master Konsistensi');
}

/* ============================================================
 * STUDENT MANAGEMENT (Admin only)
 * ==========================================================*/
function addStudent(data) {
  const username = String(data.username || '').trim();
  const password = String(data.password || '').trim();
  const fullname = String(data.fullname || data.namaLengkap || username).trim();

  if (!username || !password) throw new Error('Username dan password wajib diisi.');
  if (username.toLowerCase() === 'admin') throw new Error('Username "admin" tidak boleh digunakan untuk siswa.');

  const users = sheetToObjects_(SHEET_USERS);
  if (users.find(u => String(u.Username).toLowerCase() === username.toLowerCase())) {
    throw new Error('Username sudah digunakan.');
  }

  const sh = getSS().getSheetByName(SHEET_USERS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  const set = (name, val) => { const i = headers.indexOf(name); if (i >= 0) row[i] = val; };
  set('Username', username);
  set('Password', hashPassword_(password));
  set('Role', 'student');
  set('NamaLengkap', fullname);
  set('TanggalDaftar', new Date());
  set('Streak', 0);
  set('LongestStreak', 0);
  set('WeeklyGoal', 0);
  set('WeeklyProgress', 0);
  set('WeeklyWeek', weekStartStr_());
  set('AuthEpoch', 0);
  sh.appendRow(row);

  return { success: true, username: username, fullname: fullname };
}

function updateStudent(data) {
  const username = String(data.username || '').trim();
  if (!username) throw new Error('Username wajib.');

  const found = findRowByValue_(SHEET_USERS, 'Username', username);
  if (!found) throw new Error('Siswa tidak ditemukan.');
  if (String(found.values[found.headers.indexOf('Role')]) === 'admin') {
    throw new Error('Tidak bisa mengubah akun admin lewat sini.');
  }

  const sh = getSS().getSheetByName(SHEET_USERS);
  const headers = found.headers;
  const setCol = (name, val) => {
    const c = headers.indexOf(name) + 1;
    if (c > 0) sh.getRange(found.rowNumber, c).setValue(val);
  };

  if (data.fullname || data.namaLengkap) setCol('NamaLengkap', String(data.fullname || data.namaLengkap).trim());
  if (data.password) {
    setCol('Password', hashPassword_(String(data.password).trim()));
    setCol('AuthEpoch', Number(found.values[headers.indexOf('AuthEpoch')] || 0) + 1);
  }
  if (data.weeklyGoal !== undefined) setCol('WeeklyGoal', Number(data.weeklyGoal) || 0);

  return { success: true, username: username };
}

function deleteStudent(username) {
  username = String(username || '').trim();
  if (!username) throw new Error('Username wajib.');
  if (username.toLowerCase() === 'admin') throw new Error('Tidak bisa menghapus admin.');

  const found = findRowByValue_(SHEET_USERS, 'Username', username);
  if (!found) throw new Error('Siswa tidak ditemukan.');
  if (String(found.values[found.headers.indexOf('Role')]) === 'admin') {
    throw new Error('Tidak bisa menghapus akun admin.');
  }

  getSS().getSheetByName(SHEET_USERS).deleteRow(found.rowNumber);
  return { success: true };
}

function resetStudentPassword(data) {
  const username = String(data.username || '').trim();
  const password = String(data.password || '').trim();
  if (!username || !password) throw new Error('Username dan password baru wajib.');
  return updateStudent({ username: username, password: password });
}

/* ============================================================
 * MATERI (existing + duplicate)
 * ==========================================================*/
function getKategoriList() {
  const set = {};
  sheetToObjects_(SHEET_KATEGORI).forEach(c => {
    const n = String(c.Nama || '').trim();
    if (n) set[n] = true;
  });
  sheetToObjects_(SHEET_MATERI).forEach(m => {
    const n = String(m.Kategori || '').trim();
    if (n) set[n] = true;
  });
  return Object.keys(set).sort();
}

function getKategoriFull() {
  const cats = sheetToObjects_(SHEET_KATEGORI);
  const materi = sheetToObjects_(SHEET_MATERI);
  const countMap = {};
  materi.forEach(m => {
    const n = String(m.Kategori || '').trim();
    if (n) countMap[n] = (countMap[n] || 0) + 1;
  });

  const result = [];
  const seen = {};
  cats.forEach(c => {
    const nama = String(c.Nama || '').trim();
    if (!nama) return;
    seen[nama.toLowerCase()] = true;
    result.push({
      id: String(c.ID || ''),
      nama: nama,
      deskripsi: String(c.Deskripsi || ''),
      tanggal: toIso_(c.Tanggal),
      materiCount: countMap[nama] || 0
    });
  });
  Object.keys(countMap).forEach(nama => {
    if (seen[nama.toLowerCase()]) return;
    result.push({ id: '', nama: nama, deskripsi: '', tanggal: '', materiCount: countMap[nama] || 0 });
  });
  return result.sort((a, b) => (a.nama > b.nama ? 1 : -1));
}

function saveKategori(data) {
  const sheet = getSS().getSheetByName(SHEET_KATEGORI);
  const nama = String(data.nama || '').trim();
  if (!nama) throw new Error('Nama kategori wajib diisi.');
  const deskripsi = String(data.deskripsi || '').trim();

  const all = sheetToObjects_(SHEET_KATEGORI);
  const dup = all.find(c =>
    String(c.Nama || '').trim().toLowerCase() === nama.toLowerCase() &&
    !idEq_(c.ID, data.id)
  );
  if (dup) throw new Error('Kategori dengan nama tersebut sudah ada.');

  if (data.id) {
    const found = findRowByValue_(SHEET_KATEGORI, 'ID', data.id);
    if (found) {
      const oldNama = String(found.values[1] || '').trim();
      sheet.getRange(found.rowNumber, 2, 1, 2).setValues([[nama, deskripsi]]);
      if (oldNama && oldNama !== nama) renameKategoriCascade_(oldNama, nama);
      return { success: true, id: data.id, nama: nama };
    }
  }

  if (data.oldNama) {
    const oldNama = String(data.oldNama).trim();
    const found = all.find(c => String(c.Nama || '').trim().toLowerCase() === oldNama.toLowerCase());
    if (found) {
      const row = findRowByValue_(SHEET_KATEGORI, 'ID', found.ID);
      if (row) {
        sheet.getRange(row.rowNumber, 2, 1, 2).setValues([[nama, deskripsi]]);
        if (oldNama !== nama) renameKategoriCascade_(oldNama, nama);
        return { success: true, id: found.ID, nama: nama };
      }
    }
    const id = generateId_();
    sheet.appendRow([id, nama, deskripsi, new Date()]);
    if (oldNama !== nama) renameKategoriCascade_(oldNama, nama);
    return { success: true, id: id, nama: nama };
  }

  const id = generateId_();
  sheet.appendRow([id, nama, deskripsi, new Date()]);
  return { success: true, id: id, nama: nama };
}

function renameKategoriCascade_(oldNama, newNama) {
  const matSh = getSS().getSheetByName(SHEET_MATERI);
  if (matSh) {
    const data = matSh.getDataRange().getValues();
    const headers = data[0];
    const col = headers.indexOf('Kategori');
    if (col >= 0) {
      for (let i = 1; i < data.length; i++) {
        if (idEq_(data[i][col], oldNama)) matSh.getRange(i + 1, col + 1).setValue(newNama);
      }
    }
  }
  const soalSh = getSS().getSheetByName(SHEET_SOAL);
  if (soalSh) {
    const data = soalSh.getDataRange().getValues();
    const headers = data[0];
    const col = headers.indexOf('Kategori');
    if (col >= 0) {
      for (let i = 1; i < data.length; i++) {
        if (idEq_(data[i][col], oldNama)) soalSh.getRange(i + 1, col + 1).setValue(newNama);
      }
    }
  }
}

function deleteKategori(idOrNama) {
  const key = String(idOrNama || '').trim();
  if (!key) throw new Error('ID atau nama kategori wajib.');

  let nama = key;
  const found = findRowByValue_(SHEET_KATEGORI, 'ID', key);
  if (found) {
    nama = String(found.values[1] || '').trim();
    getSS().getSheetByName(SHEET_KATEGORI).deleteRow(found.rowNumber);
  } else {
    const byName = sheetToObjects_(SHEET_KATEGORI).find(c =>
      String(c.Nama || '').trim().toLowerCase() === key.toLowerCase()
    );
    if (byName) {
      nama = String(byName.Nama || '').trim();
      const row = findRowByValue_(SHEET_KATEGORI, 'ID', byName.ID);
      if (row) getSS().getSheetByName(SHEET_KATEGORI).deleteRow(row.rowNumber);
    }
  }

  const remaining = sheetToObjects_(SHEET_MATERI).filter(m => idEq_(m.Kategori, nama));
  if (remaining.length) {
    throw new Error('Kategori masih berisi ' + remaining.length + ' materi. Pindahkan atau hapus materi terlebih dahulu.');
  }
  return { success: true };
}

function getMateriByKategori(kategori) {
  return sheetToObjects_(SHEET_MATERI)
    .filter(m => idEq_(m.Kategori, kategori))
    .map(buildMateriObj_)
    .sort((a, b) => {
      if (a.urutan !== b.urutan) return a.urutan - b.urutan;
      return (a.judul > b.judul ? 1 : -1);
    });
}

function buildMateriObj_(m) {
  const urutan = (m.Urutan === '' || m.Urutan === undefined || m.Urutan === null) ? 0 : Number(m.Urutan);
  return {
    id: String(m.ID || ''), kategori: String(m.Kategori || ''), judul: String(m.Judul || ''),
    fileId: String(m.FileID || ''), fileName: String(m.FileName || ''), fileType: String(m.FileType || ''),
    fileUrl: String(m.FileURL || ''), tanggal: toIso_(m.Tanggal),
    kkm: (m.KKM === '' || m.KKM === undefined || m.KKM === null) ? DEFAULT_KKM : Number(m.KKM),
    ringkasan: String(m.Ringkasan || ''),
    urutan: isNaN(urutan) ? 0 : urutan,
    previewUrl: m.FileID ? ('https://drive.google.com/file/d/' + m.FileID + '/preview') : '',
    downloadUrl: m.FileID ? ('https://drive.google.com/uc?export=download&id=' + m.FileID) : ''
  };
}

function getMateriList() {
  return sheetToObjects_(SHEET_MATERI)
    .map(buildMateriObj_)
    .sort((a, b) => {
      if (a.kategori !== b.kategori) return a.kategori > b.kategori ? 1 : -1;
      if (a.urutan !== b.urutan) return a.urutan - b.urutan;
      return (a.judul > b.judul ? 1 : -1);
    });
}

function getMateriById(id) {
  const m = sheetToObjects_(SHEET_MATERI).find(x => idEq_(x.ID, id));
  return m ? buildMateriObj_(m) : null;
}

function getMateriForLatihan() {
  return sheetToObjects_(SHEET_MATERI).map(m => buildMateriObj_(m)).sort((a, b) => (a.kategori > b.kategori ? 1 : -1));
}

function getKategoriForLatihan() {
  const materiRows = sheetToObjects_(SHEET_MATERI);
  const soalRows = sheetToObjects_(SHEET_SOAL);
  const kategoriMap = {};
  materiRows.forEach(m => {
    const cat = String(m.Kategori || '').trim();
    if (cat && !kategoriMap[cat]) kategoriMap[cat] = { kategori: cat, soalCount: 0, materiCount: 0 };
    if (cat) kategoriMap[cat].materiCount++;
  });
  soalRows.forEach(s => {
    const cat = String(s.Kategori || '').trim();
    const mid = String(s.MateriID || '').trim();
    if (!cat) return;
    if (!kategoriMap[cat]) kategoriMap[cat] = { kategori: cat, soalCount: 0, materiCount: 0 };
    if (!mid && (s.Status || 'Aktif') !== 'Terkunci') kategoriMap[cat].soalCount++;
  });
  return Object.values(kategoriMap).sort((a, b) => (a.kategori > b.kategori ? 1 : -1));
}

function getLatihanList() {
  const materiRows = sheetToObjects_(SHEET_MATERI);
  const soalRows = sheetToObjects_(SHEET_SOAL);
  const settings = getSettings();

  const catSoal = {};
  const matSoal = {};
  soalRows.forEach(s => {
    if ((s.Status || 'Aktif') === 'Terkunci') return;
    const cat = String(s.Kategori || '').trim();
    const mid = String(s.MateriID || '').trim();
    if (mid) matSoal[mid] = (matSoal[mid] || 0) + 1;
    else if (cat) catSoal[cat] = (catSoal[cat] || 0) + 1;
  });

  const items = [];
  const cats = {};
  materiRows.forEach(m => {
    const cat = String(m.Kategori || '').trim();
    if (cat) cats[cat] = (cats[cat] || 0) + 1;
  });
  Object.keys(cats).sort().forEach(cat => {
    items.push({
      type: 'kategori', id: cat, label: cat, kategori: cat, materiId: '',
      materiCount: cats[cat], soalCount: catSoal[cat] || 0, judul: 'Latihan Kategori: ' + cat
    });
  });
  Object.keys(catSoal).forEach(cat => {
    if (!cats[cat]) {
      items.push({
        type: 'kategori', id: cat, label: cat, kategori: cat, materiId: '',
        materiCount: 0, soalCount: catSoal[cat] || 0, judul: 'Latihan Kategori: ' + cat
      });
    }
  });
  materiRows.forEach(m => {
    const mid = String(m.ID || '');
    const count = matSoal[mid] || 0;
    if (count <= 0) return;
    items.push({
      type: 'materi', id: mid, label: String(m.Judul || ''), kategori: String(m.Kategori || ''),
      materiId: mid, materiCount: 1, soalCount: count, judul: String(m.Judul || '')
    });
  });

  items.sort((a, b) => {
    if (a.kategori !== b.kategori) return a.kategori > b.kategori ? 1 : -1;
    if (a.type !== b.type) return a.type === 'kategori' ? -1 : 1;
    return (a.judul || '') > (b.judul || '') ? 1 : -1;
  });

  return { items: items, quizMinutes: settings.quizMinutes };
}

function saveMateri(data) {
  const sh = getSS().getSheetByName(SHEET_MATERI);
  const kkm = (data.kkm === undefined || data.kkm === null || data.kkm === '' || isNaN(Number(data.kkm)))
    ? DEFAULT_KKM : Math.max(0, Math.min(100, Number(data.kkm)));

  let uploaded = null;
  if (data.fileId) {
    uploaded = {
      fileId: String(data.fileId),
      fileName: data.fileName || 'file',
      fileType: detectFileType_(data.fileName || '', data.mimeType || ''),
      fileUrl: data.fileUrl || ('https://drive.google.com/file/d/' + data.fileId + '/view'),
      sharingOk: data.sharingOk !== false
    };
  } else if (data.base64) {
    uploaded = uploadFileToDrive_(data.fileName, data.mimeType, data.base64);
  }

  if (data.id) {
    const found = findRowByValue_(SHEET_MATERI, 'ID', data.id);
    if (found) {
      let fileId = found.values[3], fileName = found.values[4], fileType = found.values[5], fileUrl = found.values[6];
      let sharingOk = true;
      if (uploaded) {
        if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
        fileId = uploaded.fileId; fileName = uploaded.fileName; fileType = uploaded.fileType; fileUrl = uploaded.fileUrl;
        sharingOk = uploaded.sharingOk;
      }
      sh.getRange(found.rowNumber, 2, 1, 6).setValues([[data.kategori, data.judul, fileId, fileName, fileType, fileUrl]]);
      sh.getRange(found.rowNumber, 9).setValue(kkm);
      ensureColumns_(SHEET_MATERI, [{ name: 'Ringkasan', value: '' }, { name: 'Urutan', value: 0 }]);
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const ringCol = headers.indexOf('Ringkasan') + 1;
      const urutCol = headers.indexOf('Urutan') + 1;
      if (ringCol > 0) sh.getRange(found.rowNumber, ringCol).setValue(String(data.ringkasan || ''));
      if (urutCol > 0) {
        let u = Number(data.urutan);
        if (isNaN(u)) u = 0;
        sh.getRange(found.rowNumber, urutCol).setValue(u);
      }
      return { success: true, id: data.id, sharingOk: sharingOk };
    }
  }

  if (!uploaded) throw new Error('File materi wajib diunggah.');
  const id = generateId_();
  ensureColumns_(SHEET_MATERI, [{ name: 'Ringkasan', value: '' }, { name: 'Urutan', value: 0 }]);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  const set = (name, val) => { const i = headers.indexOf(name); if (i >= 0) row[i] = val; };
  let u = Number(data.urutan); if (isNaN(u)) u = 0;
  set('ID', id);
  set('Kategori', data.kategori);
  set('Judul', data.judul);
  set('FileID', uploaded.fileId);
  set('FileName', uploaded.fileName);
  set('FileType', uploaded.fileType);
  set('FileURL', uploaded.fileUrl);
  set('Tanggal', new Date());
  set('KKM', kkm);
  set('Ringkasan', String(data.ringkasan || ''));
  set('Urutan', u);
  sh.appendRow(row);
  return { success: true, id: id, sharingOk: uploaded.sharingOk };
}

function duplicateMateri(id) {
  const m = sheetToObjects_(SHEET_MATERI).find(x => idEq_(x.ID, id));
  if (!m) throw new Error('Materi tidak ditemukan.');
  const sh = getSS().getSheetByName(SHEET_MATERI);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  const set = (name, val) => { const i = headers.indexOf(name); if (i >= 0) row[i] = val; };
  const newId = generateId_();
  set('ID', newId);
  set('Kategori', m.Kategori);
  set('Judul', String(m.Judul || '') + ' (Salinan)');
  set('FileID', m.FileID);
  set('FileName', m.FileName);
  set('FileType', m.FileType);
  set('FileURL', m.FileURL);
  set('Tanggal', new Date());
  set('KKM', m.KKM || DEFAULT_KKM);
  set('Ringkasan', m.Ringkasan || '');
  set('Urutan', m.Urutan || 0);
  sh.appendRow(row);
  return { success: true, id: newId };
}

function uploadFileToDrive_(fileName, mimeType, base64) {
  if (!base64) throw new Error('Data file kosong.');
  const approxBytes = Math.floor((String(base64).length * 3) / 4);
  if (approxBytes > MAX_UPLOAD_BYTES) throw new Error('File terlalu besar (maks 35MB). Gunakan unggahan bertahap (chunk).');
  const folder = getMateriFolder_();
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName || 'file');
  const file = folder.createFile(blob);
  const sharingOk = applyFileSharing_(file);
  return {
    fileId: file.getId(),
    fileName: fileName,
    fileType: detectFileType_(fileName, mimeType),
    fileUrl: file.getUrl(),
    sharingOk: sharingOk
  };
}

function getUploadTempFolder_() {
  const folders = DriveApp.getFoldersByName(UPLOAD_TEMP_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(UPLOAD_TEMP_FOLDER_NAME);
}

function applyFileSharing_(file) {
  let sharingOk = true;
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    sharingOk = false;
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      sharingOk = true;
    } catch (e2) {}
  }
  return sharingOk;
}

function initChunkUpload(data) {
  const fileName = String(data.fileName || 'file').trim();
  const mimeType = String(data.mimeType || 'application/octet-stream');
  const totalChunks = Number(data.totalChunks || 0);
  const totalBytes = Number(data.totalBytes || 0);
  if (!fileName || totalChunks < 1) throw new Error('Parameter upload tidak valid.');
  if (totalBytes > MAX_UPLOAD_BYTES) throw new Error('File terlalu besar. Maksimal 35MB.');
  if (totalChunks > 120) throw new Error('Terlalu banyak chunk.');

  const sessionId = generateId_();
  const meta = { sessionId, fileName, mimeType, totalChunks, totalBytes, received: 0, created: Date.now() };
  CacheService.getScriptCache().put('up_' + sessionId, JSON.stringify(meta), 21600);
  return { success: true, sessionId: sessionId, maxBytes: MAX_UPLOAD_BYTES };
}

function uploadChunk(data) {
  const sessionId = String(data.sessionId || '');
  const index = Number(data.index);
  const base64 = data.base64;
  if (!sessionId || isNaN(index) || index < 0 || !base64) throw new Error('Chunk tidak valid.');
  const cache = CacheService.getScriptCache();
  const raw = cache.get('up_' + sessionId);
  if (!raw) throw new Error('Sesi upload kedaluwarsa. Silakan unggah ulang.');
  const meta = JSON.parse(raw);

  const folder = getUploadTempFolder_();
  const partName = sessionId + '__' + index;
  const existing = folder.getFilesByName(partName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, 'application/octet-stream', partName);
  folder.createFile(blob);

  meta.received = Number(meta.received || 0) + 1;
  cache.put('up_' + sessionId, JSON.stringify(meta), 21600);
  return { success: true, index: index, received: meta.received, totalChunks: meta.totalChunks };
}

function finalizeChunkUpload(data) {
  const sessionId = String(data.sessionId || '');
  if (!sessionId) throw new Error('sessionId wajib.');
  const cache = CacheService.getScriptCache();
  const raw = cache.get('up_' + sessionId);
  if (!raw) throw new Error('Sesi upload kedaluwarsa.');
  const meta = JSON.parse(raw);
  const total = Number(meta.totalChunks);
  const folder = getUploadTempFolder_();

  const parts = [];
  let totalLen = 0;
  for (let i = 0; i < total; i++) {
    const partName = sessionId + '__' + i;
    const files = folder.getFilesByName(partName);
    if (!files.hasNext()) throw new Error('Chunk #' + (i + 1) + ' hilang. Unggah ulang.');
    const f = files.next();
    const b = f.getBlob().getBytes();
    parts.push(b);
    totalLen += b.length;
    f.setTrashed(true);
  }
  if (totalLen > MAX_UPLOAD_BYTES) throw new Error('File hasil gabungan melebihi 35MB.');

  const merged = [];
  for (let p = 0; p < parts.length; p++) {
    const arr = parts[p];
    for (let j = 0; j < arr.length; j++) merged.push(arr[j]);
  }

  const materiFolder = getMateriFolder_();
  const finalBlob = Utilities.newBlob(merged, meta.mimeType || 'application/octet-stream', meta.fileName);
  const file = materiFolder.createFile(finalBlob);
  const sharingOk = applyFileSharing_(file);
  cache.remove('up_' + sessionId);

  return {
    success: true,
    fileId: file.getId(),
    fileName: meta.fileName,
    fileType: detectFileType_(meta.fileName, meta.mimeType),
    fileUrl: file.getUrl(),
    sharingOk: sharingOk,
    size: totalLen
  };
}

function abortChunkUpload(sessionId) {
  sessionId = String(sessionId || '');
  if (!sessionId) return { success: true };
  try {
    CacheService.getScriptCache().remove('up_' + sessionId);
    const folder = getUploadTempFolder_();
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (String(f.getName()).indexOf(sessionId + '__') === 0) {
        try { f.setTrashed(true); } catch (e) {}
      }
    }
  } catch (e) {}
  return { success: true };
}

function detectFileType_(fileName, mimeType) {
  const ext = (fileName.split('.').pop() || '').toLowerCase();
  if (['doc', 'docx'].includes(ext)) return 'Word';
  if (['xls', 'xlsx', 'csv'].includes(ext)) return 'Excel';
  if (['ppt', 'pptx'].includes(ext)) return 'PowerPoint';
  if (ext === 'pdf') return 'PDF';
  return ext ? ext.toUpperCase() : 'File';
}

function deleteMateri(id) {
  const found = findRowByValue_(SHEET_MATERI, 'ID', id);
  if (found) {
    const fileId = found.values[3];
    if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
    getSS().getSheetByName(SHEET_MATERI).deleteRow(found.rowNumber);
  }
  return { success: true };
}

/* ============================================================
 * SOAL
 * ==========================================================*/
function mapSoalAdmin_(s) {
  return {
    id: s.ID,
    kategori: s.Kategori,
    materiId: String(s.MateriID || ''),
    tipe: s.Tipe,
    pertanyaan: s.Pertanyaan,
    a: s.PilihanA, b: s.PilihanB, c: s.PilihanC, d: s.PilihanD,
    jawabanBenar: s.JawabanBenar,
    status: s.Status || 'Aktif',
    kesulitan: s.Kesulitan || 'Sedang',
    tags: String(s.Tags || ''),
    pembahasan: String(s.Pembahasan || ''),
    gambarUrl: String(s.GambarURL || '')
  };
}

function mapSoalStudent_(s) {
  return {
    id: s.ID,
    kategori: s.Kategori,
    materiId: String(s.MateriID || ''),
    tipe: s.Tipe,
    pertanyaan: s.Pertanyaan,
    a: s.PilihanA, b: s.PilihanB, c: s.PilihanC, d: s.PilihanD,
    status: s.Status || 'Aktif',
    kesulitan: s.Kesulitan || 'Sedang',
    gambarUrl: String(s.GambarURL || '')
  };
}

function getSoalByKategoriAdmin(kategori) {
  return sheetToObjects_(SHEET_SOAL)
    .filter(s => idEq_(s.Kategori, kategori) && !String(s.MateriID || '').trim())
    .map(mapSoalAdmin_);
}

function getSoalByKategoriStudent(kategori) {
  return getSoalByKategoriAdmin(kategori).map(s => ({
    id: s.id, kategori: s.kategori, materiId: s.materiId, tipe: s.tipe,
    pertanyaan: s.pertanyaan, a: s.a, b: s.b, c: s.c, d: s.d, status: s.status,
    kesulitan: s.kesulitan, gambarUrl: s.gambarUrl
  }));
}

function getSoalByMateriAdmin(materiId) {
  return sheetToObjects_(SHEET_SOAL)
    .filter(s => idEq_(s.MateriID, materiId))
    .map(mapSoalAdmin_);
}

function getSoalByMateriStudent(materiId) {
  return getSoalByMateriAdmin(materiId).map(s => ({
    id: s.id, kategori: s.kategori, materiId: s.materiId, tipe: s.tipe,
    pertanyaan: s.pertanyaan, a: s.a, b: s.b, c: s.c, d: s.d, status: s.status,
    kesulitan: s.kesulitan, gambarUrl: s.gambarUrl
  }));
}

function uploadSoalImage_(data) {
  const base64 = String(data.gambarBase64 || '');
  const mime = String(data.gambarMimeType || '');
  if (!/^image\/(jpeg|png|webp|gif)$/.test(mime) || !/^[A-Za-z0-9+/=]+$/.test(base64)) {
    throw new Error('Format gambar soal tidak valid.');
  }
  if (Math.floor(base64.length * 3 / 4) > 2 * 1024 * 1024) {
    throw new Error('Gambar soal maksimal 2 MB setelah kompresi.');
  }
  const name = String(data.gambarFileName || 'gambar-soal.jpg').split(/[\\/]/).pop().slice(0, 120);
  const uploaded = uploadFileToDrive_(name, mime, base64);
  if (!uploaded.sharingOk) {
    try { DriveApp.getFileById(uploaded.fileId).setTrashed(true); } catch (e) {}
    throw new Error('Gambar tidak dapat dibagikan ke siswa. Periksa pengaturan Drive.');
  }
  return { id: uploaded.fileId, url: 'https://drive.google.com/thumbnail?id=' + uploaded.fileId + '&sz=w1600' };
}

function saveSoal(data) {
  const sh = getSS().getSheetByName(SHEET_SOAL);
  const status = data.status === 'Terkunci' ? 'Terkunci' : 'Aktif';
  const materiId = String(data.materiId || data.materiID || '').trim();
  let kategori = String(data.kategori || '').trim();
  const kesulitan = ['Mudah', 'Sedang', 'Sulit'].includes(data.kesulitan) ? data.kesulitan : 'Sedang';
  const tags = String(data.tags || '').trim();
  const pembahasan = String(data.pembahasan || '').trim();

  if (materiId) {
    const m = sheetToObjects_(SHEET_MATERI).find(x => idEq_(x.ID, materiId));
    if (m) kategori = String(m.Kategori || kategori).trim();
  }
  if (!kategori) throw new Error('Kategori wajib diisi (atau pilih materi).');

  ensureColumns_(SHEET_SOAL, [
    { name: 'MateriID', value: '' },
    { name: 'Kesulitan', value: 'Sedang' },
    { name: 'Tags', value: '' },
    { name: 'Pembahasan', value: '' },
    { name: 'GambarFileID', value: '' },
    { name: 'GambarURL', value: '' }
  ]);
  const image = data.gambarBase64 ? uploadSoalImage_(data)
    : (data.gambarExistingFileId ? { id: String(data.gambarExistingFileId), url: String(data.gambarExistingUrl || '') } : null);

  if (data.id) {
    const found = findRowByValue_(SHEET_SOAL, 'ID', data.id);
    if (found) {
      const oldImageId = String(found.values[found.headers.indexOf('GambarFileID')] || '');
      sh.getRange(found.rowNumber, 2, 1, 8).setValues([[
        kategori, data.tipe, sheetLiteral_(data.pertanyaan),
        sheetLiteral_(data.a || ''), sheetLiteral_(data.b || ''),
        sheetLiteral_(data.c || ''), sheetLiteral_(data.d || ''),
        data.jawabanBenar || ''
      ]]);
      const headers = found.headers;
      const setCol = (name, val) => {
        const c = headers.indexOf(name) + 1;
        if (c > 0) sh.getRange(found.rowNumber, c).setValue(val);
      };
      setCol('Status', status);
      setCol('MateriID', materiId);
      setCol('Kesulitan', kesulitan);
      setCol('Tags', sheetLiteral_(tags));
      setCol('Pembahasan', sheetLiteral_(pembahasan));
      if (image || data.removeGambar) {
        setCol('GambarFileID', image ? image.id : '');
        setCol('GambarURL', image ? image.url : '');
        if (oldImageId && oldImageId !== (image && image.id)) {
          try { DriveApp.getFileById(oldImageId).setTrashed(true); } catch (e) {}
        }
      }
      return { success: true, id: data.id };
    }
  }

  const id = generateId_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  const set = (name, val) => { const i = headers.indexOf(name); if (i >= 0) row[i] = val; };
  set('ID', id);
  set('Kategori', kategori);
  set('Tipe', data.tipe);
  set('Pertanyaan', sheetLiteral_(data.pertanyaan));
  set('PilihanA', sheetLiteral_(data.a || ''));
  set('PilihanB', sheetLiteral_(data.b || ''));
  set('PilihanC', sheetLiteral_(data.c || ''));
  set('PilihanD', sheetLiteral_(data.d || ''));
  set('JawabanBenar', data.jawabanBenar || '');
  set('Status', status);
  set('MateriID', materiId);
  set('Kesulitan', kesulitan);
  set('Tags', sheetLiteral_(tags));
  set('Pembahasan', sheetLiteral_(pembahasan));
  set('GambarFileID', image ? image.id : '');
  set('GambarURL', image ? image.url : '');
  sh.appendRow(row);
  return { success: true, id: id };
}

/**
 * Bulk import soal dari array (hasil parse CSV/XLSX di frontend).
 * data.soals = [{ kategori, tipe, pertanyaan, a,b,c,d, jawabanBenar, kesulitan, materiId, tags, pembahasan, status }, ...]
 * data.defaultKategori / data.defaultMateriId opsional sebagai fallback.
 */
function importSoalBulk(data) {
  const soals = data.soals || data.items || [];
  if (!Array.isArray(soals) || !soals.length) {
    throw new Error('Daftar soal kosong. Pastikan file berisi minimal 1 baris data.');
  }
  if (soals.length > 200) {
    throw new Error('Maksimal 200 soal per unggahan. Pecah file menjadi beberapa bagian.');
  }

  const defaultKategori = String(data.defaultKategori || data.kategori || '').trim();
  const defaultMateriId = String(data.defaultMateriId || data.materiId || '').trim();
  const sh = getSS().getSheetByName(SHEET_SOAL);
  if (!sh) throw new Error('Sheet Soal tidak ditemukan. Jalankan initSheets() dari editor.');
  ensureColumns_(SHEET_SOAL, [
    { name: 'MateriID', value: '' }, { name: 'Kesulitan', value: 'Sedang' },
    { name: 'Tags', value: '' }, { name: 'Pembahasan', value: '' },
    { name: 'GambarFileID', value: '' }, { name: 'GambarURL', value: '' }
  ]);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const materiById = new Map(sheetToObjects_(SHEET_MATERI).map(m => [String(m.ID), m]));
  const prepared = [];
  const errors = [];

  soals.forEach((row, idx) => {
    try {
      const tipeRaw = String(row.tipe || row.Tipe || 'PG').trim();
      const tipe = /^essay$/i.test(tipeRaw) ? 'Essay' : 'PG';
      const pertanyaan = String(row.pertanyaan || row.Pertanyaan || '').trim();
      if (!pertanyaan) throw new Error('Pertanyaan kosong');

      let kategori = String(row.kategori || row.Kategori || defaultKategori || '').trim();
      let materiId = String(row.materiId || row.MateriID || row.materiID || defaultMateriId || '').trim();

      if (materiId) {
        const m = materiById.get(materiId);
        if (m) kategori = String(m.Kategori || kategori).trim();
      }
      if (!kategori) throw new Error('Kategori wajib');

      const jawabanBenar = String(row.jawabanBenar || row.JawabanBenar || '').trim().toUpperCase();
      if (tipe === 'PG' && !['A', 'B', 'C', 'D'].includes(jawabanBenar)) {
        throw new Error('JawabanBenar PG harus A/B/C/D');
      }

      const difficulty = row.kesulitan || row.Kesulitan || 'Sedang';
      const record = {
        ID: generateId_(), Kategori: kategori, Tipe: tipe,
        Pertanyaan: sheetLiteral_(pertanyaan),
        PilihanA: sheetLiteral_(row.a || row.PilihanA || row.pilihanA || ''),
        PilihanB: sheetLiteral_(row.b || row.PilihanB || row.pilihanB || ''),
        PilihanC: sheetLiteral_(row.c || row.PilihanC || row.pilihanC || ''),
        PilihanD: sheetLiteral_(row.d || row.PilihanD || row.pilihanD || ''),
        JawabanBenar: tipe === 'PG' ? jawabanBenar : '',
        Status: (row.status || row.Status) === 'Terkunci' ? 'Terkunci' : 'Aktif',
        MateriID: materiId,
        Kesulitan: ['Mudah', 'Sedang', 'Sulit'].includes(difficulty) ? difficulty : 'Sedang',
        Tags: sheetLiteral_(String(row.tags || row.Tags || '').trim()),
        Pembahasan: sheetLiteral_(String(row.pembahasan || row.Pembahasan || '').trim()),
        GambarFileID: '', GambarURL: ''
      };
      prepared.push(headers.map(h => Object.prototype.hasOwnProperty.call(record, h) ? record[h] : ''));
    } catch (e) {
      errors.push({ row: idx + 2, message: e.message || String(e) }); // +2 = header + 1-index
    }
  });

  if (prepared.length) {
    // One contiguous write instead of up to 200 appendRow and schema lookups.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      sh.getRange(sh.getLastRow() + 1, 1, prepared.length, headers.length).setValues(prepared);
    } finally {
      lock.releaseLock();
    }
  }

  return {
    success: true,
    imported: prepared.length,
    failed: errors.length,
    total: soals.length,
    errors: errors.slice(0, 20)
  };
}

function deleteSoal(id) {
  const found = findRowByValue_(SHEET_SOAL, 'ID', id);
  if (found) {
    const fileId = String(found.values[found.headers.indexOf('GambarFileID')] || '');
    getSS().getSheetByName(SHEET_SOAL).deleteRow(found.rowNumber);
    if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }
  }
  return { success: true };
}

function toggleSoalStatus(id) {
  const found = findRowByValue_(SHEET_SOAL, 'ID', id);
  if (!found) throw new Error('Soal tidak ditemukan.');
  const sh = getSS().getSheetByName(SHEET_SOAL);
  const headers = found.headers;
  const statusColIdx = headers.indexOf('Status') + 1;
  const currentStatus = found.values[statusColIdx - 1] || 'Aktif';
  const newStatus = currentStatus === 'Terkunci' ? 'Aktif' : 'Terkunci';
  sh.getRange(found.rowNumber, statusColIdx).setValue(newStatus);
  return { success: true, status: newStatus };
}

function duplicateSoal(id) {
  const s = sheetToObjects_(SHEET_SOAL).find(x => idEq_(x.ID, id));
  if (!s) throw new Error('Soal tidak ditemukan.');
  let image = null;
  if (s.GambarFileID) {
    const copy = DriveApp.getFileById(String(s.GambarFileID)).makeCopy(getMateriFolder_());
    if (!applyFileSharing_(copy)) { copy.setTrashed(true); throw new Error('Gambar salinan tidak dapat dibagikan.'); }
    image = { id: copy.getId(), url: 'https://drive.google.com/thumbnail?id=' + copy.getId() + '&sz=w1600' };
  }
  return saveSoal({
    kategori: s.Kategori,
    materiId: s.MateriID,
    tipe: s.Tipe,
    pertanyaan: String(s.Pertanyaan || '') + ' (Salinan)',
    a: s.PilihanA, b: s.PilihanB, c: s.PilihanC, d: s.PilihanD,
    jawabanBenar: s.JawabanBenar,
    status: s.Status || 'Aktif',
    kesulitan: s.Kesulitan || 'Sedang',
    tags: s.Tags || '',
    pembahasan: s.Pembahasan || '',
    gambarExistingFileId: image && image.id,
    gambarExistingUrl: image && image.url
  });
}

/* ============================================================
 * HASIL / NILAI + Wrong Queue
 * ==========================================================*/
function submitJawaban(payload) {
  const attemptId = String(payload.attemptId || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const attempt = readSignedToken_(attemptId, 'quiz');
    if (attempt.username !== String(payload.username)) throw new Error('Sesi latihan milik akun lain.');
    const sh = getSS().getSheetByName(SHEET_HASIL);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (headers.indexOf('AttemptID') === -1) throw new Error('Data hasil belum diperbarui. Jalankan initSheets() dari editor.');
    if (!attempt.nonce) throw new Error('Sesi latihan tidak valid.');
    const previous = findRowByValue_(SHEET_HASIL, 'AttemptID', attempt.nonce);
    if (previous) {
      const stored = {};
      previous.headers.forEach((h, i) => { stored[h] = previous.values[i]; });
      const summary = mapHasil_(stored);
      return { success: true, alreadySubmitted: true, skor: summary.skor,
        benar: Number(summary.benarPG || 0), totalPG: Number(summary.totalPG || 0),
        essayCount: summary.essayCount, essayStatus: summary.essayStatus,
        lulus: summary.lulus, skorAkhir: summary.skorAkhir,
        kkm: stored.KKM === '' || stored.KKM == null ? DEFAULT_KKM : Number(stored.KKM),
        details: [] };
    }
    const ids = attempt.questionIds || [];
    const idSet = new Set(ids);
    const allSoal = sheetToObjects_(SHEET_SOAL);
    const byQuestion = new Map(allSoal.map(s => [String(s.ID), s]));
    const soalList = ids.map(id => byQuestion.get(String(id)));
    if (!soalList.length || soalList.some(s => !s || (s.Status || 'Aktif') === 'Terkunci')) {
      throw new Error('Daftar soal berubah. Mulai latihan kembali.');
    }
    if (attempt.questionVersion && questionVersion_(soalList) !== attempt.questionVersion) {
      throw new Error('Soal diperbarui saat latihan berlangsung. Mulai latihan kembali agar penilaian adil.');
    }
    const submitted = Array.isArray(payload.answers) ? payload.answers : [];
    if (submitted.length > ids.length) throw new Error('Daftar jawaban tidak sesuai latihan.');
    const answers = new Map();
    submitted.forEach(ans => {
      const id = String(ans.soalId || '');
      if (!idSet.has(id) || answers.has(id)) throw new Error('Daftar jawaban tidak sesuai latihan.');
      answers.set(id, String(ans.jawaban || '').trim().slice(0, 10000));
    });

    const materi = attempt.materiId ? getMateriById(attempt.materiId) : null;
    const kkm = materi && Number.isFinite(Number(materi.kkm)) ? Number(materi.kkm) : DEFAULT_KKM;
    let benar = 0, totalPG = 0;
    const essayJawaban = [], details = [], wrongItems = [], correctIds = [];
    soalList.forEach(soal => {
      const userAns = answers.get(String(soal.ID)) || '';
      if (soal.Tipe === 'PG') {
        totalPG++;
        const selected = userAns.toUpperCase();
        const correct = String(soal.JawabanBenar || '').trim().toUpperCase();
        const isCorrect = selected === correct && selected !== '';
        if (isCorrect) { benar++; correctIds.push(String(soal.ID)); }
        else wrongItems.push({
          username: payload.username, soalId: String(soal.ID), kategori: String(soal.Kategori || ''),
          materiId: String(soal.MateriID || ''), pertanyaan: String(soal.Pertanyaan || ''),
          jawabanBenar: correct, jawabanUser: selected
        });
        details.push({
          soalId: String(soal.ID), tipe: 'PG', pertanyaan: String(soal.Pertanyaan || ''),
          a: String(soal.PilihanA || ''), b: String(soal.PilihanB || ''),
          c: String(soal.PilihanC || ''), d: String(soal.PilihanD || ''),
          jawabanUser: selected, jawabanBenar: correct, benar: isCorrect,
          pembahasan: String(soal.Pembahasan || '')
        });
      } else {
        essayJawaban.push({
          soalId: String(soal.ID), pertanyaan: String(soal.Pertanyaan || ''),
          jawaban: userAns, skor: null, maxSkor: 100
        });
        details.push({
          soalId: String(soal.ID), tipe: 'Essay', pertanyaan: String(soal.Pertanyaan || ''),
          jawabanUser: userAns, benar: null
        });
      }
    });
    const skorPG = totalPG ? Math.round(benar / totalPG * 100) : 0;
    const essayCount = essayJawaban.length;
    const essayStatus = essayCount ? 'Menunggu' : '';
    const lulus = essayCount ? 'Menunggu Penilaian'
      : (totalPG ? (skorPG >= kkm ? 'Lulus' : 'Tidak Lulus') : '');
    const hasilMateriId = attempt.type === 'retry' ? 'RETRY'
      : (attempt.materiId || 'CAT-' + attempt.kategori);
    const hasilJudul = attempt.type === 'retry' ? 'Ulangi Soal Salah'
      : (materi ? materi.judul : 'Kategori: ' + attempt.kategori);
    const row = new Array(headers.length).fill('');
    const set = (name, val) => { const i = headers.indexOf(name); if (i >= 0) row[i] = val; };
    set('ID', generateId_());
    set('Username', payload.username);
    set('MateriID', hasilMateriId);
    set('JudulMateri', hasilJudul);
    set('Skor', skorPG);
    set('BenarPG', benar);
    set('TotalPG', totalPG);
    set('JawabanEssay', JSON.stringify(essayJawaban));
    set('Tanggal', new Date());
    set('Lulus', lulus);
    set('SkorAkhir', essayCount ? '' : skorPG);
    set('EssayStatus', essayStatus);
    set('KKM', kkm);
    set('TipeLatihan', attempt.type === 'materi' || attempt.type === 'kategori' ? 'latihan' : attempt.type);
    set('AttemptID', attempt.nonce);
    sh.appendRow(row);

    if (attempt.type === 'retry' && correctIds.length) {
      const solved = new Set(correctIds);
      const wrongRows = sheetToObjects_(SHEET_WRONG).filter(w =>
        String(w.Username) === payload.username && solved.has(String(w.SoalID))
      ).sort((a, b) => b._row - a._row);
      const wrongSh = getSS().getSheetByName(SHEET_WRONG);
      wrongRows.forEach(w => wrongSh.deleteRow(w._row));
    }
    wrongItems.forEach(w => { try { addToWrongQueue(w); } catch (e) {} });
    if (materi) {
      const stage = attempt.type === 'cek'
        ? (!essayCount && skorPG >= kkm ? 'latihan' : 'cek')
        : (!essayCount && skorPG >= Math.max(85, kkm) ? 'kuasai' : 'latihan');
      upsertBelajar_(payload.username, attempt.materiId, { Stage: stage });
    }
    try {
      incrementWeeklyProgress_(payload.username);
      if (!essayCount && skorPG >= 80) awardBadge_(payload.username, 'score80', 'Skor Tinggi (≥80)');
      if (!essayCount && skorPG === 100 && totalPG >= 3) awardBadge_(payload.username, 'perfect', 'Sempurna!');
    } catch (e) {}
    return {
      success: true, skor: skorPG, benar: benar, totalPG: totalPG,
      essayCount: essayCount, lulus: lulus, kkm: kkm,
      skorAkhir: essayCount ? null : skorPG,
      essayStatus: essayStatus, details: details, wrongCount: wrongItems.length
    };
  } finally {
    lock.releaseLock();
  }
}

function mapHasil_(h) {
  let essay = [];
  try { essay = JSON.parse(h.JawabanEssay || '[]'); } catch (e) { essay = []; }
  const essayCount = Array.isArray(essay) ? essay.length : 0;
  const essayStatus = String(h.EssayStatus || (essayCount ? 'Menunggu' : '') || '');
  const skorPG = Number(h.Skor || 0);
  const essaySkor = (h.EssaySkor === '' || h.EssaySkor === null || h.EssaySkor === undefined) ? null : Number(h.EssaySkor);
  const pending = essayCount > 0 && essayStatus !== 'Dinilai';
  const skorAkhir = (h.SkorAkhir === '' || h.SkorAkhir === null || h.SkorAkhir === undefined)
    ? (pending ? null : skorPG)
    : Number(h.SkorAkhir);

  return {
    id: h.ID,
    username: h.Username,
    materiId: h.MateriID,
    judulMateri: h.JudulMateri,
    skor: skorPG,
    benarPG: h.BenarPG,
    totalPG: h.TotalPG,
    lulus: pending ? 'Menunggu Penilaian' : (h.Lulus || ''),
    jawabanEssay: h.JawabanEssay,
    essayCount: essayCount,
    essayStatus: essayStatus,
    essaySkor: essaySkor,
    skorAkhir: pending ? null : skorAkhir,
    bobotPG: h.BobotPG === '' || h.BobotPG == null ? null : Number(h.BobotPG),
    bobotEssay: h.BobotEssay === '' || h.BobotEssay == null ? null : Number(h.BobotEssay),
    catatanAdmin: String(h.CatatanAdmin || ''),
    tanggal: toIso_(h.Tanggal),
    _sortTs: h.Tanggal ? new Date(h.Tanggal).getTime() : 0
  };
}

function finalScore_(h) {
  // A PG score is only a component while any essay is awaiting grading.
  if (String(h.EssayStatus) === 'Menunggu') return null;
  let essays = [];
  try { essays = JSON.parse(h.JawabanEssay || '[]'); } catch (e) {}
  if (Array.isArray(essays) && essays.length && String(h.EssayStatus) !== 'Dinilai') return null;
  if (h.SkorAkhir !== '' && h.SkorAkhir !== null && h.SkorAkhir !== undefined) {
    return Number(h.SkorAkhir) || 0;
  }
  if (Array.isArray(essays) && essays.length) return null;
  return Number(h.Skor) || 0;
}

function gradingWeights_(totalPG, savedPG, savedEssay, suppliedPG, suppliedEssay) {
  const providedPG = suppliedPG !== undefined && suppliedPG !== null;
  const providedEssay = suppliedEssay !== undefined && suppliedEssay !== null;
  if (providedPG !== providedEssay) throw new Error('Kedua bobot harus diisi bersama.');
  const hasSaved = savedPG !== '' && savedPG != null && savedEssay !== '' && savedEssay != null;
  const pg = Number(providedPG ? suppliedPG : (hasSaved ? savedPG : (totalPG ? 60 : 0)));
  const essay = Number(providedEssay ? suppliedEssay : (hasSaved ? savedEssay : (totalPG ? 40 : 100)));
  if ((providedPG && (suppliedPG === '' || suppliedEssay === '')) ||
      !Number.isFinite(pg) || !Number.isFinite(essay) || pg < 0 || essay <= 0 ||
      pg > 100 || essay > 100 || Math.abs(pg + essay - 100) > 1e-8 ||
      (totalPG > 0 && pg <= 0) || (totalPG === 0 && (pg !== 0 || essay !== 100))) {
    throw new Error('Bobot PG dan esai harus berjumlah 100%; tiap bagian yang ada harus berbobot positif.');
  }
  return { pg: pg, essay: essay };
}

function getHasilById(id) {
  const h = sheetToObjects_(SHEET_HASIL).find(x => idEq_(x.ID, id));
  if (!h) return null;
  const mapped = mapHasil_(h);
  delete mapped._sortTs;
  return mapped;
}

function gradeEssay(data) {
  const found = findRowByValue_(SHEET_HASIL, 'ID', data.id);
  if (!found) throw new Error('Hasil latihan tidak ditemukan.');

  ensureColumns_(SHEET_HASIL, [
    { name: 'EssaySkor', value: '' },
    { name: 'SkorAkhir', value: '' },
    { name: 'CatatanAdmin', value: '' },
    { name: 'EssayStatus', value: '' },
    { name: 'BobotPG', value: '' },
    { name: 'BobotEssay', value: '' }
  ]);

  const sh = getSS().getSheetByName(SHEET_HASIL);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const values = found.values;
  const idx = (name) => headers.indexOf(name);

  let essay = [];
  try { essay = JSON.parse(values[idx('JawabanEssay')] || '[]'); } catch (e) { essay = []; }
  if (!Array.isArray(essay) || !essay.length) throw new Error('Tidak ada jawaban essay pada hasil ini.');

  const grades = data.essays || data.grades || [];
  grades.forEach(g => {
    let target = null;
    if (g.soalId) target = essay.find(e => idEq_(e.soalId, g.soalId));
    if (!target && g.index !== undefined && g.index !== null) target = essay[Number(g.index)];
    if (!target) return;
    const s = Number(g.skor);
    if (g.skor === '' || g.skor === null || g.skor === undefined ||
        !Number.isFinite(s) || s < 0 || s > 100) {
      throw new Error('Nilai tiap esai harus berupa angka 0–100.');
    }
    target.skor = s;
    target.maxSkor = 100;
    if (g.catatan) target.catatan = String(g.catatan);
  });

  const pending = essay.filter(e => e.skor === null || e.skor === undefined || e.skor === '');
  if (pending.length) throw new Error('Semua soal essay harus diberi nilai (0–100).');

  const essaySkor = Math.round(essay.reduce((sum, e) => sum + Number(e.skor || 0), 0) / essay.length);
  const skorPG = Number(values[idx('Skor')] || 0);
  const totalPG = Number(values[idx('TotalPG')] || 0);

  const weights = gradingWeights_(totalPG,
    idx('BobotPG') >= 0 ? values[idx('BobotPG')] : '',
    idx('BobotEssay') >= 0 ? values[idx('BobotEssay')] : '',
    data.bobotPG, data.bobotEssay);
  const skorAkhir = Math.round((skorPG * weights.pg + essaySkor * weights.essay) / 100);

  const rawKkm = idx('KKM') >= 0 ? values[idx('KKM')] : '';
  const savedKkm = rawKkm === '' || rawKkm === null ? NaN : Number(rawKkm);
  const resultMateriId = String(values[idx('MateriID')] || '');
  const material = resultMateriId && !resultMateriId.startsWith('CAT-') && resultMateriId !== 'RETRY'
    ? getMateriById(resultMateriId) : null;
  const kkm = Number.isFinite(savedKkm) && savedKkm >= 0 ? savedKkm
    : (material && Number.isFinite(Number(material.kkm)) ? Number(material.kkm) : DEFAULT_KKM);
  const lulus = skorAkhir >= kkm ? 'Lulus' : 'Tidak Lulus';
  const catatan = sheetLiteral_(data.catatanAdmin || data.catatan || '');

  const setCol = (name, val) => {
    const c = idx(name) + 1;
    if (c > 0) sh.getRange(found.rowNumber, c).setValue(val);
  };
  setCol('JawabanEssay', JSON.stringify(essay));
  setCol('EssaySkor', essaySkor);
  setCol('SkorAkhir', skorAkhir);
  setCol('BobotPG', weights.pg);
  setCol('BobotEssay', weights.essay);
  setCol('CatatanAdmin', catatan);
  setCol('EssayStatus', 'Dinilai');
  setCol('Lulus', lulus);

  if (material) {
    try {
      const kind = idx('TipeLatihan') >= 0 ? String(values[idx('TipeLatihan')] || 'latihan') : 'latihan';
      const stage = kind === 'cek' ? (skorAkhir >= kkm ? 'latihan' : 'cek')
        : (skorAkhir >= Math.max(85, kkm) ? 'kuasai' : 'latihan');
      upsertBelajar_(String(values[idx('Username')]), resultMateriId, { Stage: stage });
    } catch (e) { /* Final score remains authoritative if optional progress update fails. */ }
  }

  try {
    const username = String(values[idx('Username')]);
    if (skorAkhir >= 80) awardBadge_(username, 'score80', 'Skor Tinggi (≥80)');
    if (skorAkhir === 100 && totalPG + essay.length >= 3) awardBadge_(username, 'perfect', 'Sempurna!');
  } catch (e) { /* Badges are optional; the stored final grade remains authoritative. */ }

  return {
    success: true,
    id: data.id,
    essaySkor: essaySkor,
    skorPG: skorPG,
    skorAkhir: skorAkhir,
    bobotPG: weights.pg, bobotEssay: weights.essay,
    lulus: lulus,
    essayStatus: 'Dinilai'
  };
}

function getHasilByStudent(username) {
  return sheetToObjects_(SHEET_HASIL)
    .filter(h => String(h.Username) === String(username))
    .map(mapHasil_)
    .sort((a, b) => b._sortTs - a._sortTs)
    .map(({ _sortTs, ...rest }) => rest);
}

function getAllHasil() {
  return sheetToObjects_(SHEET_HASIL)
    .map(mapHasil_)
    .sort((a, b) => b._sortTs - a._sortTs)
    .map(({ _sortTs, ...rest }) => rest);
}

function exportHasil(username) {
  let list = username
    ? getHasilByStudent(username)
    : getAllHasil();
  return {
    success: true,
    count: list.length,
    data: list,
    exportedAt: new Date().toISOString()
  };
}

/* ============================================================
 * BELAJAR + Learning Path (Scaffolding)
 * ==========================================================*/
function ensureBelajarSheet_() {
  const sh = getSS().getSheetByName(SHEET_BELAJAR);
  if (!sh) throw new Error('Data belajar belum disiapkan. Jalankan initSheets() dari editor.');
  return sh;
}

function findBelajarRow_(username, materiId) {
  const rows = sheetToObjects_(SHEET_BELAJAR);
  return rows.find(b => String(b.Username) === String(username) && idEq_(b.MateriID, materiId)) || null;
}

function upsertBelajar_(username, materiId, fields) {
  const sh = ensureBelajarSheet_();
  const existing = findBelajarRow_(username, materiId);
  if (existing) {
    const found = findRowByValue_(SHEET_BELAJAR, 'ID', existing.ID);
    if (!found) throw new Error('Baris belajar tidak ditemukan.');
    const headers = found.headers;
    const setCol = (name, val) => {
      const c = headers.indexOf(name) + 1;
      if (c > 0) sh.getRange(found.rowNumber, c).setValue(val);
    };
    Object.keys(fields).forEach(k => setCol(k, fields[k]));
    setCol('UpdatedAt', new Date());
    return existing.ID;
  }
  const id = generateId_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = new Array(headers.length).fill('');
  const set = (name, val) => { const i = headers.indexOf(name); if (i >= 0) row[i] = val; };
  set('ID', id);
  set('Username', username);
  set('MateriID', materiId);
  set('Dibaca', fields.Dibaca !== undefined ? fields.Dibaca : false);
  set('TanggalBaca', fields.TanggalBaca || '');
  set('Bookmark', fields.Bookmark !== undefined ? fields.Bookmark : false);
  set('Catatan', fields.Catatan !== undefined ? fields.Catatan : '');
  set('Stage', fields.Stage || 'baca');
  set('UpdatedAt', new Date());
  sh.appendRow(row);
  return id;
}

function getBelajarMeta(username) {
  username = String(username || '');
  if (!username) return [];
  ensureBelajarSheet_();
  return sheetToObjects_(SHEET_BELAJAR)
    .filter(b => String(b.Username) === username)
    .map(b => ({
      id: b.ID,
      materiId: String(b.MateriID || ''),
      dibaca: String(b.Dibaca) === 'true' || b.Dibaca === true || b.Dibaca === 'TRUE' || b.Dibaca === 1,
      tanggalBaca: toIso_(b.TanggalBaca),
      bookmark: String(b.Bookmark) === 'true' || b.Bookmark === true || b.Bookmark === 'TRUE' || b.Bookmark === 1,
      catatan: String(b.Catatan || ''),
      stage: String(b.Stage || 'baca'),
      updatedAt: toIso_(b.UpdatedAt)
    }));
}

function markMateriDibaca(data) {
  const username = String(data.username || '').trim();
  const materiId = String(data.materiId || data.id || '').trim();
  if (!username || !materiId) throw new Error('username dan materiId wajib.');
  const dibaca = data.dibaca === false || data.dibaca === 'false' ? false : true;
  upsertBelajar_(username, materiId, {
    Dibaca: dibaca,
    TanggalBaca: dibaca ? new Date() : '',
    Stage: dibaca ? 'cek' : 'baca'
  });
  if (dibaca) {
    try { incrementWeeklyProgress_(username); } catch (e) {}
  }
  return { success: true, materiId: materiId, dibaca: dibaca };
}

function toggleBookmark(data) {
  const username = String(data.username || '').trim();
  const materiId = String(data.materiId || data.id || '').trim();
  if (!username || !materiId) throw new Error('username dan materiId wajib.');
  const existing = findBelajarRow_(username, materiId);
  let next = true;
  if (existing) {
    const cur = String(existing.Bookmark) === 'true' || existing.Bookmark === true || existing.Bookmark === 'TRUE' || existing.Bookmark === 1;
    next = !cur;
  }
  upsertBelajar_(username, materiId, { Bookmark: next });
  return { success: true, materiId: materiId, bookmark: next };
}

function saveCatatanMateri(data) {
  const username = String(data.username || '').trim();
  const materiId = String(data.materiId || data.id || '').trim();
  const catatan = sheetLiteral_(data.catatan || '');
  if (!username || !materiId) throw new Error('username dan materiId wajib.');
  upsertBelajar_(username, materiId, { Catatan: catatan });
  return { success: true, materiId: materiId };
}

/** Learning Path / Scaffolding stages: baca → cek → latihan → kuasai */
function getLearningPath(materiId, username) {
  materiId = String(materiId || '').trim();
  username = String(username || '').trim();
  if (!materiId) throw new Error('materiId wajib.');

  const m = getMateriById(materiId);
  if (!m) throw new Error('Materi tidak ditemukan.');

  const meta = findBelajarRow_(username, materiId);
  return buildLearningPath_(m, meta, sheetToObjects_(SHEET_SOAL), sheetToObjects_(SHEET_HASIL), username);
}

/** One authenticated request and one read per sheet needed by the reader. */
function getMateriReader(materiId, username) {
  materiId = String(materiId || '').trim();
  username = String(username || '').trim();
  if (!materiId || !username) throw new Error('Materi dan akun wajib diisi.');
  const row = sheetToObjects_(SHEET_MATERI).find(m => idEq_(m.ID, materiId));
  if (!row) throw new Error('Materi tidak ditemukan.');
  const materi = buildMateriObj_(row);
  const meta = sheetToObjects_(SHEET_BELAJAR).find(b =>
    String(b.Username) === username && idEq_(b.MateriID, materiId)) || null;
  const path = buildLearningPath_(materi, meta,
    sheetToObjects_(SHEET_SOAL), sheetToObjects_(SHEET_HASIL), username);
  const yes = value => value === true || value === 1 || String(value).toLowerCase() === 'true';
  return {
    materi: materi,
    path: path,
    flashcards: getFlashcards(username, materiId),
    belajar: {
      materiId: materiId,
      dibaca: meta ? yes(meta.Dibaca) : false,
      bookmark: meta ? yes(meta.Bookmark) : false,
      catatan: meta ? String(meta.Catatan || '') : '',
      stage: meta ? String(meta.Stage || 'baca') : 'baca'
    }
  };
}

function buildLearningPath_(m, meta, questions, hasilRows, username) {
  const materiId = String(m.id);
  const stage = meta ? String(meta.Stage || 'baca') : 'baca';
  const dibaca = meta ? (String(meta.Dibaca) === 'true' || meta.Dibaca === true) : false;

  // Both flags can be calculated from one sheet read without serializing questions.
  const active = s => (s.Status || 'Aktif') !== 'Terkunci';
  const hasSoalMateri = questions.some(s => active(s) && idEq_(s.MateriID, materiId));
  const hasSoalKategori = questions.some(s => active(s) && idEq_(s.Kategori, m.kategori) && !String(s.MateriID || '').trim());

  // Cek apakah sudah pernah latihan materi ini
  const attempts = hasilRows.filter(h =>
    String(h.Username) === username && idEq_(h.MateriID, materiId) &&
    (!h.TipeLatihan || String(h.TipeLatihan) === 'latihan')
  );
  const hasil = attempts.filter(h => finalScore_(h) !== null);
  const pernahLatihan = attempts.length > 0;
  const skorTerbaik = hasil.length ? Math.max(...hasil.map(finalScore_)) : null;
  const threshold = Number.isFinite(Number(m.kkm)) ? Number(m.kkm) : DEFAULT_KKM;
  const lulus = skorTerbaik !== null && skorTerbaik >= threshold;

  const stages = [
    { id: 'baca', label: 'Baca Materi', done: dibaca, unlocked: true },
    { id: 'cek', label: 'Cek Pemahaman', done: stage === 'latihan' || stage === 'kuasai' || pernahLatihan, unlocked: dibaca },
    { id: 'latihan', label: 'Latihan Penuh', done: lulus, unlocked: dibaca },
    { id: 'kuasai', label: 'Kuasai', done: skorTerbaik !== null && skorTerbaik >= Math.max(85, threshold), unlocked: pernahLatihan }
  ];

  return {
    materiId: materiId,
    judul: m.judul,
    kategori: m.kategori,
    currentStage: stage,
    stages: stages,
    hasSoalMateri: hasSoalMateri,
    hasSoalKategori: hasSoalKategori,
    skorTerbaik: skorTerbaik,
    dibaca: dibaca
  };
}

function setLearningStage(data) {
  const username = String(data.username || '').trim();
  const materiId = String(data.materiId || data.id || '').trim();
  const stage = String(data.stage || '').trim().toLowerCase();
  const allowed = ['baca', 'cek', 'latihan', 'kuasai'];
  if (!username || !materiId) throw new Error('username dan materiId wajib.');
  if (allowed.indexOf(stage) === -1) throw new Error('Stage tidak valid.');
  upsertBelajar_(username, materiId, { Stage: stage });
  return { success: true, materiId: materiId, stage: stage };
}

function getCekPemahaman(materiId) {
  materiId = String(materiId || '').trim();
  if (!materiId) throw new Error('materiId wajib.');
  const m = sheetToObjects_(SHEET_MATERI).find(x => idEq_(x.ID, materiId));
  if (!m) throw new Error('Materi tidak ditemukan.');
  const kategori = String(m.Kategori || '');

  let soal = getSoalByMateriStudent(materiId).filter(s => s.status !== 'Terkunci');
  if (soal.length < 3) {
    const extra = getSoalByKategoriStudent(kategori).filter(s => s.status !== 'Terkunci');
    const ids = {};
    soal.forEach(s => ids[s.id] = true);
    extra.forEach(s => { if (!ids[s.id]) soal.push(s); });
  }
  for (let i = soal.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = soal[i]; soal[i] = soal[j]; soal[j] = t;
  }
  const picked = soal.slice(0, Math.min(5, Math.max(soal.length, 0)));
  return {
    materiId: materiId,
    kategori: kategori,
    judul: String(m.Judul || ''),
    soal: picked,
    total: picked.length
  };
}

function getProgressByStudent(username) {
  const currentIds = new Set(sheetToObjects_(SHEET_MATERI).map(m => String(m.ID)));
  const totalMateri = currentIds.size;
  if (totalMateri === 0) return { total: 0, selesai: 0, persen: 0, dibaca: 0, bookmark: 0 };
  const belajar = sheetToObjects_(SHEET_BELAJAR).filter(b => String(b.Username) === String(username));
  const dibacaSet = {};
  let bookmarkCount = 0;
  belajar.forEach(b => {
    if (currentIds.has(String(b.MateriID)) && (String(b.Dibaca) === 'true' || b.Dibaca === true || b.Dibaca === 'TRUE' || b.Dibaca === 1)) {
      dibacaSet[String(b.MateriID)] = true;
    }
    if (currentIds.has(String(b.MateriID)) && (String(b.Bookmark) === 'true' || b.Bookmark === true || b.Bookmark === 'TRUE' || b.Bookmark === 1)) {
      bookmarkCount++;
    }
  });
  const hasil = sheetToObjects_(SHEET_HASIL).filter(h => String(h.Username) === String(username));
  hasil.forEach(h => {
    const mid = String(h.MateriID || '');
    if (currentIds.has(mid) && String(h.Lulus) === 'Lulus' &&
      (!h.TipeLatihan || String(h.TipeLatihan) === 'latihan')) dibacaSet[mid] = true;
  });
  const selesai = Object.keys(dibacaSet).length;
  return {
    total: totalMateri,
    selesai: selesai,
    dibaca: selesai,
    bookmark: bookmarkCount,
    persen: Math.round((selesai / totalMateri) * 100)
  };
}

/* ============================================================
 * STREAK, BADGE, GOALS
 * ==========================================================*/
function getStudentProfile(username) {
  return getStudentDashboard(username).profile;
}

function recordActivity(data) {
  const username = String(data.username || '').trim();
  if (!username) throw new Error('username wajib.');
  updateStreakOnLogin_(username);
  return { success: true };
}

function awardBadge_(username, code, name) {
  const existing = sheetToObjects_(SHEET_BADGE).find(b =>
    String(b.Username) === username && String(b.BadgeCode) === code
  );
  if (existing) return;
  const sh = getSS().getSheetByName(SHEET_BADGE);
  sh.appendRow([generateId_(), username, code, name, new Date()]);
}

function getBadges(username) {
  return sheetToObjects_(SHEET_BADGE)
    .filter(b => String(b.Username) === String(username))
    .map(b => ({
      code: b.BadgeCode,
      name: b.BadgeName,
      earnedAt: toIso_(b.EarnedAt)
    }));
}

function setWeeklyGoal(data) {
  const username = String(data.username || '').trim();
  const goal = Number(data.goal || 0);
  if (!username) throw new Error('username wajib.');
  const found = findRowByValue_(SHEET_USERS, 'Username', username);
  if (!found) throw new Error('User tidak ditemukan.');
  const sh = getSS().getSheetByName(SHEET_USERS);
  const col = found.headers.indexOf('WeeklyGoal') + 1;
  if (col > 0) sh.getRange(found.rowNumber, col).setValue(Math.max(0, goal));
  return { success: true, weeklyGoal: goal };
}

function incrementWeeklyProgress_(username) {
  const found = findRowByValue_(SHEET_USERS, 'Username', username);
  if (!found) return;
  const sh = getSS().getSheetByName(SHEET_USERS);
  const headers = found.headers;
  const idx = headers.indexOf('WeeklyProgress');
  if (idx < 0) return;
  const weekIdx = headers.indexOf('WeeklyWeek');
  const thisWeek = weekStartStr_();
  let prog = (weekIdx >= 0 && String(found.values[weekIdx] || '') === thisWeek)
    ? Number(found.values[idx] || 0) + 1 : 1;
  sh.getRange(found.rowNumber, idx + 1).setValue(prog);
  if (weekIdx >= 0) sh.getRange(found.rowNumber, weekIdx + 1).setValue(thisWeek);

  const goalIdx = headers.indexOf('WeeklyGoal');
  const goal = goalIdx >= 0 ? Number(found.values[goalIdx] || 0) : 0;
  if (goal > 0 && prog >= goal) {
    awardBadge_(username, 'weekly_goal', 'Tujuan Mingguan Tercapai');
  }
}

/* ============================================================
 * WRONG QUEUE (Retry soal yang salah)
 * ==========================================================*/
function addToWrongQueue(data) {
  const username = String(data.username || '').trim();
  const soalId = String(data.soalId || '').trim();
  if (!username || !soalId) return { success: false };

  // Cek sudah ada?
  const existing = sheetToObjects_(SHEET_WRONG).find(w =>
    String(w.Username) === username && idEq_(w.SoalID, soalId)
  );
  if (existing) {
    // increment retry count
    const found = findRowByValue_(SHEET_WRONG, 'ID', existing.ID);
    if (found) {
      const sh = getSS().getSheetByName(SHEET_WRONG);
      const col = found.headers.indexOf('RetryCount') + 1;
      if (col > 0) {
        const cur = Number(found.values[col - 1] || 0) + 1;
        sh.getRange(found.rowNumber, col).setValue(cur);
      }
    }
    return { success: true, updated: true };
  }

  const sh = getSS().getSheetByName(SHEET_WRONG);
  sh.appendRow([
    generateId_(),
    username,
    soalId,
    data.kategori || '',
    data.materiId || '',
    data.pertanyaan || '',
    data.jawabanBenar || '',
    data.jawabanUser || '',
    new Date(),
    0
  ]);
  return { success: true };
}

function getWrongQueue(username) {
  return sheetToObjects_(SHEET_WRONG)
    .filter(w => String(w.Username) === String(username))
    .map(w => ({
      id: w.ID,
      soalId: w.SoalID,
      kategori: w.Kategori,
      materiId: w.MateriID,
      pertanyaan: w.Pertanyaan,
      jawabanBenar: w.JawabanBenar,
      jawabanUser: w.JawabanUser,
      addedAt: toIso_(w.AddedAt),
      retryCount: Number(w.RetryCount || 0)
    }));
}

function clearWrongItem(data) {
  const id = String(data.id || '').trim();
  const username = String(data.username || '').trim();
  if (!id) throw new Error('id wajib.');
  const found = findRowByValue_(SHEET_WRONG, 'ID', id);
  if (found) {
    if (username && String(found.values[found.headers.indexOf('Username')]) !== username) {
      throw new Error('Tidak berwenang.');
    }
    getSS().getSheetByName(SHEET_WRONG).deleteRow(found.rowNumber);
  }
  return { success: true };
}

function getRetryQuiz(username) {
  const queue = getWrongQueue(username);
  if (!queue.length) return { soal: [], total: 0 };

  const soalIds = queue.map(w => w.soalId);
  const allSoal = sheetToObjects_(SHEET_SOAL);
  const soal = [];
  soalIds.forEach(sid => {
    const s = allSoal.find(x => idEq_(x.ID, sid));
    if (s && (s.Status || 'Aktif') !== 'Terkunci') {
      soal.push(mapSoalStudent_(s));
    }
  });
  return { soal: soal, total: soal.length, fromQueue: true };
}

/* ============================================================
 * WEAKNESS ANALYSIS
 * ==========================================================*/
function getWeaknessAnalysis(username) {
  username = String(username || '').trim();
  const hasil = sheetToObjects_(SHEET_HASIL).filter(h =>
    String(h.Username) === username && String(h.TipeLatihan || 'latihan') !== 'retry' &&
    finalScore_(h) !== null
  );
  if (!hasil.length) return { categories: [], message: 'Belum ada data latihan.' };

  const catStats = {};
  const materiById = new Map(sheetToObjects_(SHEET_MATERI).map(m => [String(m.ID), m]));
  hasil.forEach(h => {
    let cat = '';
    const mid = String(h.MateriID || '');
    if (mid.indexOf('CAT-') === 0) {
      cat = mid.substring(4);
    } else {
      const m = materiById.get(mid);
      cat = m ? String(m.Kategori || '') : 'Lainnya';
    }
    if (!cat) cat = 'Lainnya';
    if (!catStats[cat]) catStats[cat] = { total: 0, sumSkor: 0, count: 0 };
    const skor = finalScore_(h);
    catStats[cat].sumSkor += skor;
    catStats[cat].count += 1;
  });

  const categories = Object.keys(catStats).map(cat => {
    const s = catStats[cat];
    const avg = s.count ? Math.round(s.sumSkor / s.count) : 0;
    return {
      kategori: cat,
      avgSkor: avg,
      attempts: s.count,
      level: avg >= 80 ? 'kuat' : (avg >= 60 ? 'sedang' : 'lemah')
    };
  }).sort((a, b) => a.avgSkor - b.avgSkor);

  return { categories: categories };
}

/* ============================================================
 * FLASHCARD
 * ==========================================================*/
/** Daftar ringan untuk ponsel: tidak mengirim teks lengkap setiap kartu. */
function getFlashcardOverview(username) {
  username = String(username || '').trim();
  const materi = getMateriList().map(m => ({ id: m.id, judul: m.judul, kategori: m.kategori }));
  const counts = {};
  sheetToObjects_(SHEET_FLASHCARD).forEach(f => {
    if (String(f.Username) !== username) return;
    const id = String(f.MateriID || '');
    if (!id) return;
    if (!counts[id]) counts[id] = { total: 0, mastered: 0 };
    counts[id].total++;
    if (String(f.ReviewStatus) === 'paham') counts[id].mastered++;
  });
  return { materi: materi, counts: counts };
}

function getFlashcards(username, materiId) {
  username = String(username || '').trim();
  let list = sheetToObjects_(SHEET_FLASHCARD).filter(f => String(f.Username) === username);
  if (materiId) list = list.filter(f => idEq_(f.MateriID, materiId));
  return list.map(f => ({
    id: f.ID,
    materiId: f.MateriID,
    kategori: f.Kategori,
    front: f.Front,
    back: f.Back,
    createdAt: toIso_(f.CreatedAt),
    reviewStatus: ['paham', 'ulang'].indexOf(String(f.ReviewStatus)) !== -1 ? String(f.ReviewStatus) : 'baru',
    reviewCount: Number(f.ReviewCount) || 0,
    lastReviewedAt: f.LastReviewedAt ? toIso_(f.LastReviewedAt) : ''
  }));
}

function saveFlashcard(data) {
  const username = String(data.username || '').trim();
  const front = sheetLiteral_(String(data.front || '').trim());
  const back = sheetLiteral_(String(data.back || '').trim());
  const materiId = String(data.materiId || '').trim();
  const kategori = String(data.kategori || '').trim();
  if (!username || !front || !back) throw new Error('username, front, dan back wajib.');

  // Sheet initialization is an editor-only setup operation.
  const sh = getSS().getSheetByName(SHEET_FLASHCARD);
  if (!sh) throw new Error('Sheet Flashcard tidak ditemukan. Jalankan initSheets().');

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = (name) => headers.indexOf(name) + 1;

  if (data.id) {
    const found = findRowByValue_(SHEET_FLASHCARD, 'ID', data.id);
    if (found) {
      if (String(found.values[found.headers.indexOf('Username')]) !== username) throw new Error('Kartu ini milik akun lain.');
      const setCol = (name, val) => {
        const c = col(name);
        if (c > 0) sh.getRange(found.rowNumber, c).setValue(val);
      };
      setCol('Front', front);
      setCol('Back', back);
      if (materiId) setCol('MateriID', materiId);
      if (kategori) setCol('Kategori', kategori);
      return { success: true, id: data.id };
    }
    throw new Error('Kartu tidak ditemukan. Muat ulang daftar kartu.');
  }

  const id = generateId_();
  // Urutan kolom: ID, Username, MateriID, Kategori, Front, Back, CreatedAt
  const row = new Array(headers.length).fill('');
  const set = (name, val) => {
    const i = headers.indexOf(name);
    if (i >= 0) row[i] = val;
  };
  set('ID', id);
  set('Username', username);
  set('MateriID', materiId);
  set('Kategori', kategori);
  set('Front', front);
  set('Back', back);
  set('CreatedAt', new Date());
  sh.appendRow(row);
  return { success: true, id: id, materiId: materiId };
}

function deleteFlashcard(id, username) {
  const found = findRowByValue_(SHEET_FLASHCARD, 'ID', id);
  if (found) {
    if (String(found.values[found.headers.indexOf('Username')]) !== String(username)) throw new Error('Kartu ini milik akun lain.');
    getSS().getSheetByName(SHEET_FLASHCARD).deleteRow(found.rowNumber);
  }
  return { success: true };
}

function reviewFlashcard(data) {
  const username = String(data.username || '').trim();
  const id = String(data.id || '').trim();
  const status = String(data.status || '').trim();
  if (!username || !id || ['paham', 'ulang'].indexOf(status) === -1) {
    throw new Error('Kartu dan hasil latihan tidak valid.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const found = findRowByValue_(SHEET_FLASHCARD, 'ID', id);
    if (!found) throw new Error('Kartu tidak ditemukan.');
    if (String(found.values[found.headers.indexOf('Username')]) !== username) {
      throw new Error('Kartu ini milik akun lain.');
    }
    // Tambahkan kolom pada sheet lama saat penggunaan pertama, tanpa mengubah kartu yang ada.
    ensureColumns_(SHEET_FLASHCARD, [
      { name: 'ReviewStatus', value: 'baru' },
      { name: 'ReviewCount', value: 0 },
      { name: 'LastReviewedAt', value: '' }
    ]);
    const sh = getSS().getSheetByName(SHEET_FLASHCARD);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const col = name => headers.indexOf(name) + 1;
    const count = Number(sh.getRange(found.rowNumber, col('ReviewCount')).getValue()) || 0;
    const now = new Date();
    sh.getRange(found.rowNumber, col('ReviewStatus')).setValue(status);
    sh.getRange(found.rowNumber, col('ReviewCount')).setValue(count + 1);
    sh.getRange(found.rowNumber, col('LastReviewedAt')).setValue(now);
    return { success: true, id: id, reviewStatus: status, reviewCount: count + 1,
      lastReviewedAt: toIso_(now) };
  } finally {
    lock.releaseLock();
  }
}

function generateFlashcardsFromMateri(data) {
  const username = String(data.username || '').trim();
  const materiId = String(data.materiId || '').trim();
  if (!username || !materiId) throw new Error('username dan materiId wajib.');

  const m = getMateriById(materiId);
  if (!m) throw new Error('Materi tidak ditemukan.');
  const ringkasan = String(m.ringkasan || '').trim();
  if (!ringkasan) {
    return { success: true, created: 0, skipped: 0, message: 'Materi belum punya ringkasan.' };
  }

  const lines = ringkasan.split(/\n+/).map(l => l.trim()).filter(l => l.length > 10);
  // Dua klik bersamaan tidak boleh membuat kartu yang sama dua kali.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const key = (front, back) => String(front).trim().replace(/\s+/g, ' ').toLowerCase() +
      '\u0000' + String(back).trim().replace(/\s+/g, ' ').toLowerCase();
    const existing = new Set(getFlashcards(username, materiId).map(f => key(f.front, f.back)));
    let created = 0;
    let skipped = 0;
    lines.slice(0, 8).forEach((line, i) => {
      const parts = line.split(/[:–—-]/);
      const front = (parts.length >= 2 ? parts[0].trim() : 'Poin ' + (i + 1) + ' — ' + m.judul).substring(0, 120);
      const back = (parts.length >= 2 ? parts.slice(1).join(':').trim() : line).substring(0, 300);
      if (!front || !back) return;
      const signature = key(front, back);
      if (existing.has(signature)) {
        skipped++;
        return;
      }
      saveFlashcard({ username: username, materiId: materiId, kategori: m.kategori,
        front: front, back: back });
      existing.add(signature);
      created++;
    });
    return { success: true, created: created, skipped: skipped,
      message: created ? '' : (skipped ? 'Semua kartu dari ringkasan sudah ada.' : 'Tidak ada poin yang dapat dijadikan kartu.') };
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * SETTINGS
 * ==========================================================*/
function getSettings() {
  const props = PropertiesService.getScriptProperties();
  let quizMinutes = Number(props.getProperty('QUIZ_MINUTES') || '20');
  if (isNaN(quizMinutes) || quizMinutes < 1) quizMinutes = 20;
  if (quizMinutes > 180) quizMinutes = 180;
  return { quizMinutes: quizMinutes };
}

function saveSettings(data) {
  const props = PropertiesService.getScriptProperties();
  let quizMinutes = Number(data.quizMinutes);
  if (isNaN(quizMinutes) || quizMinutes < 1) quizMinutes = 20;
  if (quizMinutes > 180) quizMinutes = 180;
  props.setProperty('QUIZ_MINUTES', String(quizMinutes));
  return { success: true, quizMinutes: quizMinutes };
}

/* ============================================================
 * ADMIN STATS
 * ==========================================================*/
function getStudentList() {
  const users = sheetToObjects_(SHEET_USERS).filter(u => u.Role === 'student');
  const allHasil = sheetToObjects_(SHEET_HASIL);
  const materiIds = new Set(sheetToObjects_(SHEET_MATERI).map(m => String(m.ID)));
  const totalMateri = materiIds.size;

  return users.map(u => {
    const hasilUser = allHasil.filter(h => String(h.Username) === String(u.Username));
    const graded = hasilUser.filter(h =>
      String(h.TipeLatihan || 'latihan') === 'latihan' && finalScore_(h) !== null);
    const rataRata = graded.length
      ? Math.round(graded.reduce((sum, h) => sum + finalScore_(h), 0) / graded.length)
      : 0;
    const materiSelesai = {};
    hasilUser.forEach(h => {
      const mid = String(h.MateriID || '');
      if (materiIds.has(mid) && String(h.Lulus) === 'Lulus' &&
        (!h.TipeLatihan || String(h.TipeLatihan) === 'latihan')) materiSelesai[mid] = true;
    });
    return {
      username: u.Username,
      fullname: u.NamaLengkap,
      tanggalDaftar: toIso_(u.TanggalDaftar),
      lastActive: toIso_(u.LastActive),
      streak: Number(u.Streak || 0),
      _sortTs: u.TanggalDaftar ? new Date(u.TanggalDaftar).getTime() : 0,
      totalLatihan: hasilUser.length,
      rataRata: rataRata,
      progress: totalMateri > 0 ? Math.round((Object.keys(materiSelesai).length / totalMateri) * 100) : 0
    };
  }).sort((a, b) => b._sortTs - a._sortTs).map(({ _sortTs, ...rest }) => rest);
}

function getAdminStats() {
  const settings = getSettings();
  const count = name => Math.max(0, getSS().getSheetByName(name).getLastRow() - 1);
  return {
    totalMateri: count(SHEET_MATERI),
    totalSoal: count(SHEET_SOAL),
    totalStudent: sheetToObjects_(SHEET_USERS).filter(u => u.Role === 'student').length,
    totalLatihan: count(SHEET_HASIL),
    quizMinutes: settings.quizMinutes
  };
}

function getRichStats() {
  const base = getAdminStats();
  const hasil = sheetToObjects_(SHEET_HASIL);
  const soal = sheetToObjects_(SHEET_SOAL);

  // Rata-rata skor global
  let sum = 0, cnt = 0;
  hasil.forEach(h => {
    if (String(h.TipeLatihan || 'latihan') !== 'latihan') return;
    const score = finalScore_(h);
    if (score !== null) { sum += score; cnt++; }
  });
  base.avgSkor = cnt ? Math.round(sum / cnt) : 0;

  // Soal terkunci
  base.soalTerkunci = soal.filter(s => (s.Status || 'Aktif') === 'Terkunci').length;

  // Essay menunggu
  base.essayMenunggu = hasil.filter(h => String(h.EssayStatus) === 'Menunggu').length;

  return base;
}

/** Start an assessed attempt. The question set stays on the server so a
 * shortened or edited browser submission cannot change the denominator. */
function questionVersion_(questions) {
  // The keyed digest prevents exposing answer keys through the readable token.
  const snapshot = questions.map(s => [s.ID, s.Tipe, s.Pertanyaan,
    s.PilihanA, s.PilihanB, s.PilihanC, s.PilihanD,
    s.JawabanBenar, s.GambarURL, s.Status || 'Aktif']);
  return Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(JSON.stringify(snapshot), signingKey_())
  ).replace(/=+$/, '');
}

function startQuizAttempt(data) {
  const username = String(data.username || '').trim();
  const type = String(data.type || 'kategori');
  const materiId = String(data.materiId || '').trim();
  const kategori = String(data.kategori || '').trim();
  const all = sheetToObjects_(SHEET_SOAL).filter(s => (s.Status || 'Aktif') !== 'Terkunci');
  let list = [];
  let material = null;
  if (type === 'materi' || type === 'cek') {
    material = getMateriById(materiId);
    if (!material) throw new Error('Materi tidak ditemukan.');
    list = all.filter(s => idEq_(s.MateriID, materiId));
    if (type === 'cek') {
      if (list.length < 3) {
        const extras = all.filter(s => idEq_(s.Kategori, material.kategori) && !String(s.MateriID || '').trim());
        list = list.concat(extras);
      }
      list = list.sort(() => Math.random() - .5).slice(0, 5);
    }
  } else if (type === 'kategori') {
    list = all.filter(s => idEq_(s.Kategori, kategori) && !String(s.MateriID || '').trim());
  } else if (type === 'retry') {
    const ids = new Set(getWrongQueue(username).map(w => String(w.soalId)));
    list = all.filter(s => ids.has(String(s.ID)));
  } else {
    throw new Error('Jenis latihan tidak valid.');
  }
  if (!list.length) return { soal: [], attemptId: '', quizMinutes: getSettings().quizMinutes };
  const attemptId = signedToken_({
    kind: 'quiz', username: username, nonce: Utilities.getUuid().replace(/-/g, ''),
    type: type, materiId: material ? materiId : '',
    kategori: material ? material.kategori : kategori,
    questionIds: list.map(s => String(s.ID)), questionVersion: questionVersion_(list),
    exp: Date.now() + 3600000
  });
  return {
    soal: list.map(mapSoalStudent_), attemptId: attemptId,
    quizMinutes: getSettings().quizMinutes,
    kategori: material ? material.kategori : kategori,
    judul: material ? material.judul : (type === 'retry' ? 'Ulangi Soal Salah' : kategori)
  };
}

/* Validate every request on the server. The role and username sent by a
 * browser are never accepted as proof of identity or authorization. */
const ADMIN_ACTIONS_ = [
  'getStudentList','addStudent','updateStudent','deleteStudent','resetStudentPassword',
  'saveKategori','deleteKategori','saveMateri','initChunkUpload','uploadChunk',
  'finalizeChunkUpload','abortChunkUpload','deleteMateri','duplicateMateri',
  'getSoalByKategoriAdmin','getSoalByMateriAdmin','saveSoal','importSoalBulk',
  'deleteSoal','toggleSoalStatus','duplicateSoal','gradeEssay','getHasilById',
  'getAllHasil','exportHasil','saveSettings','getAdminStats','getRichStats'
];
const SHARED_ACTIONS_ = [
  'getKategoriList','getKategoriFull','getMateriByKategori','getMateriList',
  'getMateriById','getMateriForLatihan','getKategoriForLatihan','getLatihanList',
  'getSoalByKategoriStudent','getSoalByMateriStudent','getSettings'
];
const OWN_ACTIONS_ = [
  'getHasilByStudent','getProgressByStudent','getBelajarMeta','markMateriDibaca',
  'toggleBookmark','saveCatatanMateri','getLearningPath','getMateriReader','getStudentDashboard',
  'getStudentProfile','recordActivity','getBadges','setWeeklyGoal','getWrongQueue',
  'clearWrongItem','getRetryQuiz','getWeaknessAnalysis','getFlashcards','getFlashcardOverview',
  'saveFlashcard','deleteFlashcard','generateFlashcardsFromMateri','reviewFlashcard',
  'getCekPemahaman','startQuizAttempt','submitJawaban'
];

function requireSession_(token) {
  const session = readSignedToken_(token, 'session');
  // Password reset, logout, role change, or account removal revokes old tokens.
  const user = sheetToObjects_(SHEET_USERS).find(u =>
    String(u.Username) === session.username && String(u.Role || 'student') === session.role
  );
  if (!user || Number(user.AuthEpoch || 0) !== session.epoch ||
    credentialTag_(user.Password) !== session.credentialTag) {
    throw new Error('Sesi tidak berlaku. Silakan masuk kembali.');
  }
  return session;
}

function authorizeAction_(action, data, session) {
  if (action === 'logout') return;
  if (ADMIN_ACTIONS_.indexOf(action) !== -1) {
    if (session.role !== 'admin') throw new Error('Akses admin diperlukan.');
    return;
  }
  if (SHARED_ACTIONS_.indexOf(action) !== -1) return;
  if (OWN_ACTIONS_.indexOf(action) === -1) throw new Error('Tindakan tidak tersedia.');
  if (session.role !== 'student') throw new Error('Akses siswa diperlukan.');
  // ID-only handlers verify item ownership before changing data.
  if (action === 'deleteFlashcard' || action === 'getCekPemahaman') {
    data._actor = session.username;
    return;
  }
  const container = action === 'submitJawaban' ? (data.payload || data) : (data.data || data);
  const given = container.username || data.username;
  if (given && String(given) !== session.username) throw new Error('Akses akun lain ditolak.');
  container.username = session.username;
  data.username = session.username;
  data._actor = session.username;
}

/** One read per relevant sheet and one response for the student home screen. */
function getStudentDashboard(username) {
  username = String(username || '').trim();
  const user = sheetToObjects_(SHEET_USERS).find(u => String(u.Username) === username);
  if (!user) throw new Error('Akun tidak ditemukan.');
  const materi = sheetToObjects_(SHEET_MATERI);
  const materiMap = new Map(materi.map(m => [String(m.ID), m]));
  const belajar = sheetToObjects_(SHEET_BELAJAR).filter(b => String(b.Username) === username);
  const hasil = sheetToObjects_(SHEET_HASIL).filter(h => String(h.Username) === username);
  const badgeRows = sheetToObjects_(SHEET_BADGE).filter(b => String(b.Username) === username);
  const wrongRows = sheetToObjects_(SHEET_WRONG).filter(w => String(w.Username) === username);
  const completed = new Set();
  const saved = [];
  let bookmarkCount = 0;
  belajar.forEach(b => {
    const mid = String(b.MateriID || '');
    if (!materiMap.has(mid)) return;
    const dibaca = b.Dibaca === true || String(b.Dibaca).toLowerCase() === 'true' || b.Dibaca === 1;
    const bookmark = b.Bookmark === true || String(b.Bookmark).toLowerCase() === 'true' || b.Bookmark === 1;
    if (dibaca) completed.add(mid);
    if (bookmark) bookmarkCount++;
    saved.push({ materiId: mid, dibaca: dibaca, bookmark: bookmark,
      catatan: String(b.Catatan || ''), stage: String(b.Stage || 'baca') });
  });
  const categories = {};
  let scoreSum = 0, scoreCount = 0;
  hasil.forEach(h => {
    const mid = String(h.MateriID || '');
    const type = String(h.TipeLatihan || 'latihan');
    if (materiMap.has(mid) && String(h.Lulus) === 'Lulus' && type === 'latihan') completed.add(mid);
    const score = finalScore_(h);
    const hasFinal = score !== null;
    if (hasFinal && type === 'latihan') { scoreSum += score; scoreCount++; }
    if (type === 'retry' || !hasFinal) return;
    const cat = mid.indexOf('CAT-') === 0 ? mid.substring(4)
      : (materiMap.get(mid) ? String(materiMap.get(mid).Kategori || '') : 'Lainnya');
    if (!categories[cat]) categories[cat] = { sum: 0, count: 0 };
    categories[cat].sum += score;
    categories[cat].count++;
  });
  const weakness = Object.keys(categories).map(k => {
    const avg = Math.round(categories[k].sum / categories[k].count);
    return { kategori: k, avgSkor: avg, attempts: categories[k].count,
      level: avg >= 80 ? 'kuat' : (avg >= 60 ? 'sedang' : 'lemah') };
  }).sort((a, b) => a.avgSkor - b.avgSkor);
  const total = materi.length;
  const progress = { total: total, selesai: completed.size, dibaca: completed.size,
    bookmark: bookmarkCount, persen: total ? Math.round(completed.size / total * 100) : 0 };
  const profile = {
    username: username, fullname: String(user.NamaLengkap || username),
    streak: Number(user.Streak || 0), longestStreak: Number(user.LongestStreak || 0),
    weeklyGoal: Number(user.WeeklyGoal || 0),
    weeklyProgress: String(user.WeeklyWeek || '') === weekStartStr_() ? Number(user.WeeklyProgress || 0) : 0,
    badges: badgeRows.map(b => ({ code: b.BadgeCode, name: b.BadgeName, earnedAt: toIso_(b.EarnedAt) })),
    wrongCount: wrongRows.length, weakness: { categories: weakness }, progress: progress
  };
  const latest = materi.slice().sort((a, b) => {
    const aTime = a.Tanggal ? new Date(a.Tanggal).getTime() || 0 : 0;
    const bTime = b.Tanggal ? new Date(b.Tanggal).getTime() || 0 : 0;
    return bTime - aTime;
  }).slice(0, 3).map(buildMateriObj_);
  const ordered = materi.slice().sort((a, b) => {
    if (String(a.Kategori) !== String(b.Kategori)) return String(a.Kategori).localeCompare(String(b.Kategori));
    return (Number(a.Urutan) || 0) - (Number(b.Urutan) || 0);
  });
  const nextRow = ordered.find(m => !completed.has(String(m.ID))) || ordered[0];
  return { profile: profile, progress: progress,
    totalLatihan: hasil.length, rataRata: scoreCount ? Math.round(scoreSum / scoreCount) : 0,
    materi: latest, nextMateri: nextRow ? buildMateriObj_(nextRow) : null, belajar: saved };
}
