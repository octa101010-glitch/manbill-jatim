/**
 * ==============================================================================
 * APIK — APLIKASI PENILAIAN & INSENTIF KINERJA MANBILL
 * (Upload, login multi-role, penilaian kinerja 100% & reward per periode)
 * File: Code.gs
 * 5 Jenjang Jabatan Sesuai Matriks Terbaru:
 * 1. SPV MANBILL AREA (Kinerja MO 10% + SLA Kontrak 30% + Penyelesaian BA Tagihan 20% = 60%, Penugasan Lain 10%, Absensi 10%, Atasan 20% = 100%)
 * 2. SPV MANBILL ULP  (Kinerja MO 10% + SLA Kontrak 30% + Penyelesaian BA Tagihan 20% = 60%, Penugasan Lain 10%, Absensi 10%, Atasan 20% = 100%)
 * 3. KORDINATOR       (Kinerja MO 40% + SLA Kontrak 15% + Penugasan PLN 5% = 60%, SIQ 10%, Absensi 10%, Atasan 20% = 100%)
 * 4. PETUGAS MANBILL  (Kinerja MO 40% + SLA Kontrak 15% + Penugasan PLN 5% = 60%, SIQ 10%, Absensi 10%, Atasan 20% = 100%)
 * 5. ADMIN MANBILL    (Pembentukkan RPP 40% + SLA Kontrak 15% + Penugasan PLN 5% = 60%, SIQ 10%, Absensi 10%, Atasan 20% = 100%)
 *
 * PERBAIKAN VERSI INI:
 * - Konfigurasi reward disimpan di sheet tersembunyi "_Config" (berpotongan/chunk),
 *   karena ScriptProperties dibatasi ±9 KB per nilai sehingga nilai kinerja individu
 *   dulu gagal tersimpan tanpa pesan error.
 * - Login & ubah password diverifikasi di server, password disimpan dalam bentuk hash,
 *   dan daftar password TIDAK lagi dikirim ke browser.
 * - Akun MUL punya password sendiri (kunci "MUL <AREA>"), tidak lagi berbagi dengan akun Area.
 * - ID Kordinator / SPV ULP / SPV Area / Admin dibuat stabil (berdasarkan nama unit),
 *   bukan nomor urut, sehingga nilai kinerja tidak "pindah orang" saat data berubah.
 *   Nilai lama ber-ID nomor urut otomatis dimigrasikan.
 * - Data sheet yang berisi tanggal disanitasi sebelum dikirim ke browser
 *   (objek Date membuat google.script.run gagal mengembalikan data).
 * - Parsing kolom TOTAL: "5%" dibaca 5%, pecahan Excel (0,95) dibaca 95%.
 * - LockService untuk mencegah dua penyimpanan bersamaan saling menimpa.
 * - Pagu Area = (Nilai Reward - Nilai Denda) x 70%.
 * - SLA Kontrak ULP menjadi pengali bobot reward seluruh penerima di ULP tsb.
 * - PERIODE KINERJA (BLTH): data, nilai & reward dihitung per bulan. Data lama = April 2026.
 * - KUNCI PERIODE: UP dapat mengunci periode (verifikasi password UP di server);
 *   periode terkunci tidak bisa diubah (server menolak semua penyimpanan).
 * - v2.0: SESI LOGIN di server (token), user SDM (read only + unduh Excel),
 *   ALUR PENGAJUAN & APPROVAL REWARD (Area → MUL → UP 3 → SDM) dengan kunci per Area per bulan,
 *   snapshot reward yang disetujui (sheet Reward_Approved), log (sheet Log_Pengajuan),
 *   dan NIP untuk Kordinator / SPV / Admin.
 * ==============================================================================
 */

const SHEET_DATA_NAME = "Data_Manbill";
const SHEET_DASHBOARD_NAME = "Dashboard";
const SHEET_REWARD_NAME = "Data_Reward";
const SHEET_CONFIG_NAME = "_Config";
const CONFIG_CHUNK_SIZE = 45000;
const CONFIG_CHUNK_PREFIX = "J:";
const DEFAULT_USER_PASSWORD = "pln@123";
// Akun "admin": kewenangan sama dengan UP 3 Jatim, password terpisah (default pln@123)
const ADMIN_ACCOUNT_KEY = "ADMIN";
const UP_LEVEL_ACCOUNT_KEYS = ["UP 3 JATIM", ADMIN_ACCOUNT_KEY];
const DEFAULT_AREA_PAGU = 50000000;
const PASSWORD_HASH_PREFIX = "sha256:";
const NETTO_PAGU_PCT = 0.7; // Pagu Area = (Nilai Reward - Nilai Denda) x 70%
const CONFIG_MAP_KEYS = ["areaRewards", "areaRewardGross", "areaDenda", "areaSlaScores", "ulpSlaScores", "customScores", "adminNames", "spvUlpNames", "spvAreaNames", "unitNips"];

function doGet(e) {
  // API untuk Portal AREK Manbill: ?api=portal (ringkasan & pengumuman) / ?api=guide (Buku Panduan)
  const api = e && e.parameter ? String(e.parameter.api || "") : "";
  if (api) return portalApiResponse_(api, e.parameter || {});
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("APIK - Aplikasi Penilaian & Insentif Kinerja Manbill | PLN")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================== UTILITAS UMUM ============================== */

function withScriptLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function findHeaderIdx(headers, names) {
  return (headers || []).findIndex(h => names.indexOf(String(h === null || h === undefined ? "" : h).trim().toUpperCase()) !== -1);
}

function parseScoreValue(rawT) {
  if (rawT === null || rawT === undefined) return 100;
  const str = String(rawT).trim();
  if (str === "" || str === "-%" || str === "-") return 100;
  const isPctString = str.indexOf("%") !== -1;
  const num = parseFloat(str.replace("%", "").trim().replace(",", "."));
  if (isNaN(num)) return 100;
  return (!isPctString && num <= 10) ? num * 100 : num;
}

function makeKoordId(up3, ulp, koord) { return "KOORD|" + up3 + "|" + ulp + "|" + String(koord || "").trim().toUpperCase(); }
function makeSpvUlpId(up3, ulp) { return "SPV-ULP|" + up3 + "|" + ulp; }
function makeSpvAreaId(up3) { return "SPV-AREA|" + up3; }
function makeAdminId(up3) { return "ADMIN|" + up3; }

// Ambil nilai kustom berdasarkan ID stabil; jika masih tersimpan dengan ID lama (nomor urut), pindahkan.
function takeCustomScore(cfg, newId, legacyId) {
  const cs = cfg.customScores;
  if (cs[newId] !== undefined) return cs[newId];
  if (legacyId && cs[legacyId] !== undefined) {
    cs[newId] = cs[legacyId];
    delete cs[legacyId];
    return cs[newId];
  }
  return undefined;
}

function sanitizeValuesForClient(values) {
  const tz = Session.getScriptTimeZone() || "Asia/Jakarta";
  return values.map(row => row.map(v => {
    if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
    if (v === null || v === undefined) return "";
    return v;
  }));
}

/* ============================== AKUN & PASSWORD ============================== */

function normalizeAccountKey(raw) {
  const clean = String(raw || "").trim().toUpperCase().replace(/\s+/g, " ");
  if (clean === "UP 3 JATIM" || clean === "UP3 JATIM" || clean === "UP JATIM" || clean === "UID JATIM") {
    return "UP 3 JATIM";
  }
  return clean;
}

function getUserPasswordsMap() {
  try {
    const saved = PropertiesService.getScriptProperties().getProperty("USER_PASSWORDS_MAP");
    if (saved) return JSON.parse(saved);
  } catch (e) {}
  return {};
}

function saveUserPasswordsMap(map) {
  PropertiesService.getScriptProperties().setProperty("USER_PASSWORDS_MAP", JSON.stringify(map || {}));
}

function getPasswordSalt() {
  const props = PropertiesService.getScriptProperties();
  let salt = props.getProperty("PASSWORD_SALT");
  if (!salt) {
    salt = Utilities.getUuid();
    props.setProperty("PASSWORD_SALT", salt);
  }
  return salt;
}

function hashPassword(plain) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, getPasswordSalt() + "|" + String(plain), Utilities.Charset.UTF_8);
  return PASSWORD_HASH_PREFIX + bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, "0")).join("");
}

function isPasswordMatch(stored, plain) {
  if (!stored) return plain === DEFAULT_USER_PASSWORD;
  if (String(stored).indexOf(PASSWORD_HASH_PREFIX) === 0) return hashPassword(plain) === stored;
  return plain === stored; // format lama (teks biasa)
}

/* ---------- Reset password oleh akun admin ---------- */

function assertAdminPassword(adminPass) {
  const passMap = getUserPasswordsMap();
  if (!isPasswordMatch(passMap[ADMIN_ACCOUNT_KEY], String(adminPass || "").trim())) {
    throw new Error("Password akun admin tidak sesuai. Fitur ini hanya untuk akun admin.");
  }
  return passMap;
}

/** Daftar akun yang passwordnya sudah diganti (bukan default). Wajib password admin. */
function adminGetPasswordStatus(adminPass) {
  try {
    const passMap = assertAdminPassword(adminPass);
    return { success: true, changedKeys: Object.keys(passMap).filter(k => k !== ADMIN_ACCOUNT_KEY) };
  } catch (err) {
    return { success: false, message: err.message || String(err) };
  }
}

/** Reset password satu akun ke default (pln@123). Wajib password admin. */
function adminResetUserPassword(adminPass, targetUsername) {
  try {
    assertAdminPassword(adminPass);
    const key = normalizeAccountKey(targetUsername);
    if (!key) return { success: false, message: "Akun yang akan di-reset belum dipilih." };
    if (key === ADMIN_ACCOUNT_KEY) return { success: false, message: "Password akun admin tidak dapat di-reset dari menu ini. Gunakan Ubah Password." };
    return withScriptLock(() => {
      const map = getUserPasswordsMap();
      const wasChanged = Object.prototype.hasOwnProperty.call(map, key);
      delete map[key];
      saveUserPasswordsMap(map);
      return {
        success: true,
        wasChanged: wasChanged,
        changedKeys: Object.keys(map).filter(k => k !== ADMIN_ACCOUNT_KEY),
        message: wasChanged
          ? "Password akun " + key + " berhasil di-reset ke default (" + DEFAULT_USER_PASSWORD + ")."
          : "Akun " + key + " memang masih memakai password default (" + DEFAULT_USER_PASSWORD + ")."
      };
    });
  } catch (err) {
    return { success: false, message: err.message || String(err) };
  }
}

/* ============================== SESI LOGIN (v2.0) ==============================
 * Setelah login berhasil, server memberi kode sesi (token) yang disimpan di CacheService.
 * Setiap penyimpanan / pengajuan / approval wajib membawa token ini, sehingga server tahu
 * siapa penggunanya (level & wilayah) tanpa bergantung pada tampilan di browser.
 * ============================================================================== */

const SDM_ACCOUNT_KEY = "SDM";
const SESSION_TTL_SEC = 21600; // 6 jam (batas maksimum CacheService), diperpanjang setiap kali dipakai
const SESSION_PREFIX = "SES:";
const SESSION_EXPIRED_TAG = "SESI_BERAKHIR";
const AREAS_WITHOUT_MUL = ["BANYUWANGI", "SITUBONDO", "JEMBER", "KEDIRI"];

function createSession(key) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
  CacheService.getScriptCache().put(SESSION_PREFIX + token, JSON.stringify({ key: key, at: Date.now() }), SESSION_TTL_SEC);
  return token;
}

function readSession(token) {
  const t = String(token || "").trim();
  if (!t) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get(SESSION_PREFIX + t);
  if (!raw) return null;
  try {
    cache.put(SESSION_PREFIX + t, raw, SESSION_TTL_SEC);
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function logoutSession(token) {
  try { CacheService.getScriptCache().remove(SESSION_PREFIX + String(token || "")); } catch (e) {}
  return { success: true };
}

// Tentukan level & wilayah dari kunci akun (urutan sama dengan halaman login)
function resolveAccountRole(key, dir) {
  if (key === "UP 3 JATIM") return { role: "UP", label: "UP 3 Jatim" };
  if (key === ADMIN_ACCOUNT_KEY) return { role: "UP", isAdmin: true, label: "Admin" };
  if (key === SDM_ACCOUNT_KEY) return { role: "SDM", label: "SDM" };
  if (/^MUL /.test(key)) {
    const area = key.substring(4).trim();
    if (!area || AREAS_WITHOUT_MUL.indexOf(area) !== -1) return null;
    return { role: "MUL", up3: area, label: "MUL Area " + area };
  }
  if (dir.areas.indexOf(key) !== -1) return { role: "AREA", up3: key, label: "Area " + key };
  const ulp = dir.ulps.find(x => x[1] === key);
  if (ulp) return { role: "ULP", up3: ulp[0], ulp: ulp[1], label: "SPV ULP " + ulp[1] };
  const koords = dir.koords.filter(x => String(x[2]).trim().toUpperCase() === key);
  if (koords.length) {
    const areas = Array.from(new Set(koords.map(x => x[0])));
    return { role: "KOORD", up3: koords[0][0], ulp: koords[0][1], koord: key, areas: areas, label: "Kordinator " + koords[0][2] };
  }
  return null;
}

function requireSession(token, all) {
  const s = readSession(token);
  if (!s || !s.key) {
    throw new Error(SESSION_EXPIRED_TAG + ": Sesi login Anda sudah berakhir. Silakan keluar, lalu login kembali.");
  }
  const data = all || readAllDataRows(SpreadsheetApp.getActiveSpreadsheet());
  const r = resolveAccountRole(s.key, buildUnitDirectory(data));
  if (!r) throw new Error("Akun " + s.key + " tidak dikenali di data unit.");
  return Object.assign({ key: s.key }, r);
}

// null = seluruh Jawa Timur
function sessionAreas(sess) {
  if (sess.role === "UP" || sess.role === "SDM") return null;
  if (sess.role === "KOORD") return sess.areas || [sess.up3];
  return [sess.up3];
}

function assertCanMutate(sess) {
  if (sess.role === "SDM") throw new Error("Akun SDM hanya dapat melihat dan mengunduh data (read only).");
}

function errText(err) {
  return String(err && err.message ? err.message : err);
}

function verifyLogin(username, password) {
  try {
    const key = normalizeAccountKey(username);
    if (!key) return { success: false, message: "Username akun tidak valid." };
    const pass = String(password || "").trim();
    const passMap = getUserPasswordsMap();
    const stored = passMap[key];
    const ok = isPasswordMatch(stored, pass);

    // Upgrade otomatis password lama (teks biasa) menjadi hash
    if (ok && stored && String(stored).indexOf(PASSWORD_HASH_PREFIX) !== 0) {
      try {
        withScriptLock(() => {
          const fresh = getUserPasswordsMap();
          fresh[key] = hashPassword(pass);
          saveUserPasswordsMap(fresh);
        });
      } catch (e) {}
    }
    if (!ok) return { success: false, message: "Password yang Anda masukkan tidak sesuai." };
    return { success: true, message: "Login berhasil.", token: createSession(key) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function changeUserPassword(username, oldPass, newPass) {
  try {
    const key = normalizeAccountKey(username);
    if (!key) return { success: false, message: "Username akun tidak valid." };
    const cleanNew = String(newPass || "").trim();
    if (cleanNew.length < 4) return { success: false, message: "Password baru minimal 4 karakter." };

    return withScriptLock(() => {
      const passMap = getUserPasswordsMap();
      if (!isPasswordMatch(passMap[key], String(oldPass || "").trim())) {
        return { success: false, message: "Password lama tidak sesuai." };
      }
      passMap[key] = hashPassword(cleanNew);
      saveUserPasswordsMap(passMap);
      return { success: true, message: "Password berhasil diperbarui!" };
    });
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/* ============================== SHEET HELPER ============================== */

function ensureSheetDimensions(sheet, minRows, minCols) {
  const targetRows = Math.max(Number(minRows) || 1, 1);
  const targetCols = Math.max(Number(minCols) || 1, 1);
  const curRows = sheet.getMaxRows();
  if (curRows < targetRows) sheet.insertRowsAfter(curRows, targetRows - curRows);
  const curCols = sheet.getMaxColumns();
  if (curCols < targetCols) sheet.insertColumnsAfter(curCols, targetCols - curCols);
}

function resetSheetSafely(sheet, minRows, minCols) {
  ensureSheetDimensions(sheet, minRows, minCols);
  try { sheet.getDataRange().breakApart(); } catch (e) {}
  sheet.clear();
}

function normalizeTableMatrix(headers, rows) {
  const safeHeaders = Array.isArray(headers) ? Array.from(headers, h => String(h === null || h === undefined ? "" : h).trim()) : [];
  let maxCols = Math.max(safeHeaders.length, 1);
  const validRows = (rows || []).filter(r => Array.isArray(r) && r.length > 0 && r.some(c => c !== null && c !== undefined && String(c).trim() !== ""));
  validRows.forEach(r => { if (r.length > maxCols) maxCols = r.length; });
  while (safeHeaders.length < maxCols) safeHeaders.push("KOLOM_" + (safeHeaders.length + 1));

  const normalizedRows = validRows.map(r => {
    const rowCopy = [];
    for (let i = 0; i < maxCols; i++) {
      const val = r[i];
      rowCopy.push(val === undefined || val === null ? "" : val);
    }
    return rowCopy;
  });
  return { headers: safeHeaders, rows: normalizedRows, colCount: maxCols };
}

/* ============================== PENYIMPANAN KONFIGURASI ============================== */

function readConfigStore() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      const sh = ss.getSheetByName(SHEET_CONFIG_NAME);
      if (sh && sh.getLastRow() > 0) {
        const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
        const json = vals.map(r => {
          const s = String(r[0] || "");
          return s.indexOf(CONFIG_CHUNK_PREFIX) === 0 ? s.substring(CONFIG_CHUNK_PREFIX.length) : s;
        }).join("");
        if (json) return JSON.parse(json);
      }
    }
  } catch (e) {}
  // Fallback: format lama di ScriptProperties
  try {
    const saved = PropertiesService.getScriptProperties().getProperty("REWARD_CONFIG");
    if (saved) return JSON.parse(saved);
  } catch (e) {}
  return null;
}

function writeConfigStore(obj) {
  const json = JSON.stringify(obj || {});
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    PropertiesService.getScriptProperties().setProperty("REWARD_CONFIG", json);
    return;
  }
  let sh = ss.getSheetByName(SHEET_CONFIG_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_CONFIG_NAME);
    try { sh.hideSheet(); } catch (e) {}
  }
  try { CacheService.getScriptCache().remove("PORTAL_DATA_V1"); } catch (e) {}  // ringkasan portal ikut segar
  const chunks = [];
  for (let i = 0; i < json.length; i += CONFIG_CHUNK_SIZE) {
    chunks.push([CONFIG_CHUNK_PREFIX + json.substring(i, i + CONFIG_CHUNK_SIZE)]);
  }
  if (chunks.length === 0) chunks.push([CONFIG_CHUNK_PREFIX + "{}"]);
  ensureSheetDimensions(sh, chunks.length + 1, 1);
  sh.getRange(1, 1, sh.getMaxRows(), 1).clearContent();
  const rng = sh.getRange(1, 1, chunks.length, 1);
  rng.setNumberFormat("@");
  rng.setValues(chunks);
}

/* ============================== PERIODE (BLTH) ==============================
 * BLTH = TahunBulan "YYYYMM" (contoh "202606" = Juni 2026).
 * - Data_Manbill punya kolom BLTH. Data lama tanpa kolom BLTH = periode April 2026.
 * - Pengaturan yang berbeda tiap bulan (pagu, denda, SLA, nilai kinerja, koefisien)
 *   disimpan per periode di store.periods[BLTH]. Nama SPV/Admin berlaku global.
 * - Periode yang dikunci (store.lockedPeriods) tidak bisa diubah lagi.
 * ========================================================================== */

const LEGACY_BLTH = "202604";
// Kolom pada Data_Manbill untuk realisasi absensi (%) hasil upload file Absensi
const ABSENSI_HEADER = "ABSENSI (%)";

// Nilai sel absensi → persen. Kosong / tidak valid → null (memakai nilai manual / standar 100%)
function parseAbsensiCell(v) {
  if (v === null || v === undefined) return null;
  const str = String(v).trim();
  if (str === "" || str === "-") return null;
  const isPct = str.indexOf("%") !== -1;
  const num = parseFloat(str.replace("%", "").trim().replace(",", "."));
  if (isNaN(num) || num < 0) return null;
  return Number(((!isPct && num <= 10) ? num * 100 : num).toFixed(2));
}
const BLTH_HEADER = "BLTH";
const NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
// Kunci pengaturan yang berbeda tiap periode
const PERIOD_MAP_KEYS = ["areaRewards", "areaRewardGross", "areaDenda", "areaSlaScores", "ulpSlaScores", "customScores"];
const PERIOD_SCALAR_KEYS = ["koefAdmin", "koefPetugas", "koefKoordinator", "koefTLULP", "koefTLArea", "useScoreInWeight", "onlyApproved"];
// Kunci pengaturan yang berlaku untuk semua periode
const GLOBAL_MAP_KEYS = ["adminNames", "spvUlpNames", "spvAreaNames", "unitNips"]; // unitNips: NIP Kordinator / SPV / Admin per ID
const DEFAULT_PERIOD_SCALARS = { koefAdmin: 0.3, koefPetugas: 0.4, koefKoordinator: 0.6, koefTLULP: 0.8, koefTLArea: 1, useScoreInWeight: true, onlyApproved: false };

function isValidBlth(blth) {
  const s = String(blth || "");
  if (!/^\d{6}$/.test(s)) return false;
  const m = Number(s.substring(4));
  return m >= 1 && m <= 12;
}

function blthLabel(blth) {
  const s = String(blth || "");
  if (!isValidBlth(s)) return s;
  return NAMA_BULAN[Number(s.substring(4)) - 1] + " " + s.substring(0, 4);
}

// Nilai BLTH di sheet bisa terbaca sebagai angka; samakan menjadi teks "YYYYMM"
function cellToBlth(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone() || "Asia/Jakarta", "yyyyMM");
  const s = String(v).trim();
  if (/^\d{6}(\.0+)?$/.test(s)) return s.substring(0, 6);
  return s;
}

/* ------------------------- Store konfigurasi ------------------------- */

function emptyPeriodConfig(scalars) {
  const p = {};
  PERIOD_MAP_KEYS.forEach(k => { p[k] = {}; });
  Object.assign(p, DEFAULT_PERIOD_SCALARS, scalars || {});
  return p;
}

// Bentuk store: { defaultAreaReward, defaultAdminName, adminNames, spvUlpNames, spvAreaNames, periods: {BLTH: {...}}, lockedPeriods: {BLTH: {...}} }
function getConfigStore() {
  const raw = readConfigStore();
  const store = {
    defaultAreaReward: DEFAULT_AREA_PAGU,
    defaultAdminName: "ADMIN MANBILL",
    adminNames: {}, spvUlpNames: {}, spvAreaNames: {}, unitNips: {},
    periods: {}, lockedPeriods: {}, submissions: {}
  };
  if (!raw || typeof raw !== "object") return store;

  if (raw.defaultAreaReward !== undefined) store.defaultAreaReward = raw.defaultAreaReward;
  if (raw.defaultAdminName) store.defaultAdminName = raw.defaultAdminName;
  GLOBAL_MAP_KEYS.forEach(k => { if (raw[k] && typeof raw[k] === "object") store[k] = Object.assign({}, raw[k]); });
  if (raw.lockedPeriods && typeof raw.lockedPeriods === "object") store.lockedPeriods = Object.assign({}, raw.lockedPeriods);
  if (raw.submissions && typeof raw.submissions === "object") store.submissions = Object.assign({}, raw.submissions);

  if (raw.periods && typeof raw.periods === "object") {
    Object.keys(raw.periods).forEach(b => {
      const src = raw.periods[b] || {};
      const p = emptyPeriodConfig();
      PERIOD_MAP_KEYS.forEach(k => { if (src[k] && typeof src[k] === "object") p[k] = Object.assign({}, src[k]); });
      PERIOD_SCALAR_KEYS.forEach(k => { if (src[k] !== undefined) p[k] = src[k]; });
      store.periods[b] = p;
    });
  } else {
    // MIGRASI: konfigurasi lama (tanpa periode) menjadi milik periode April 2026
    const p = emptyPeriodConfig();
    PERIOD_MAP_KEYS.forEach(k => { if (raw[k] && typeof raw[k] === "object") p[k] = Object.assign({}, raw[k]); });
    PERIOD_SCALAR_KEYS.forEach(k => { if (raw[k] !== undefined) p[k] = raw[k]; });
    store.periods[LEGACY_BLTH] = p;
  }
  // Pagu lama (sebelum ada Nilai Reward & Denda) dipakai sebagai Nilai Reward kotor
  Object.keys(store.periods).forEach(b => {
    const p = store.periods[b];
    Object.keys(p.areaRewards || {}).forEach(a => {
      if (p.areaRewardGross[a] === undefined) p.areaRewardGross[a] = Math.max(Number(p.areaRewards[a]) || 0, 0);
    });
  });
  return store;
}

// Periode baru mewarisi koefisien & opsi dari periode terakhir sebelumnya
function ensurePeriodConfig(store, blth) {
  if (store.periods[blth]) return store.periods[blth];
  const prev = Object.keys(store.periods).filter(b => b < blth).sort().pop()
    || Object.keys(store.periods).sort().pop();
  const scalars = {};
  if (prev) PERIOD_SCALAR_KEYS.forEach(k => { scalars[k] = store.periods[prev][k]; });
  store.periods[blth] = emptyPeriodConfig(scalars);
  return store.periods[blth];
}

// Konfigurasi "datar" untuk satu periode (bentuk yang dipakai perhitungan & browser)
function flattenConfigForPeriod(store, blth) {
  const p = store.periods[blth] || ensurePeriodConfig(JSON.parse(JSON.stringify(store)), blth);
  const flat = {
    defaultAreaReward: store.defaultAreaReward,
    defaultAdminName: store.defaultAdminName
  };
  GLOBAL_MAP_KEYS.forEach(k => { flat[k] = Object.assign({}, store[k]); });
  PERIOD_MAP_KEYS.forEach(k => { flat[k] = Object.assign({}, p[k]); });
  PERIOD_SCALAR_KEYS.forEach(k => { flat[k] = p[k]; });
  return flat;
}

// Gabungkan konfigurasi datar dari browser ke store untuk satu periode
function mergeFlatIntoStore(store, flat, blth) {
  if (!flat || typeof flat !== "object") return;
  const p = ensurePeriodConfig(store, blth);
  PERIOD_MAP_KEYS.forEach(k => {
    if (flat[k] && typeof flat[k] === "object") p[k] = Object.assign({}, p[k], flat[k]);
  });
  PERIOD_SCALAR_KEYS.forEach(k => { if (flat[k] !== undefined) p[k] = flat[k]; });
  GLOBAL_MAP_KEYS.forEach(k => {
    if (flat[k] && typeof flat[k] === "object") store[k] = Object.assign({}, store[k], flat[k]);
  });
  if (flat.defaultAdminName) store.defaultAdminName = flat.defaultAdminName;
}

function saveConfigStore(store) {
  try {
    writeConfigStore(store);
  } catch (e) {
    throw new Error("Gagal menyimpan konfigurasi reward: " + e);
  }
}

function assertPeriodUnlocked(store, blth) {
  if (store.lockedPeriods && store.lockedPeriods[blth]) {
    throw new Error("Periode " + blthLabel(blth) + " sudah DIKUNCI. Data dan nilai tidak dapat diubah.");
  }
}

/* ------------------------- Data sheet per periode ------------------------- */

// Baca seluruh Data_Manbill → { headers (tanpa BLTH), rows: [{blth, cells}] }
function readAllDataRows(ss) {
  const sheet = ss.getSheetByName(SHEET_DATA_NAME);
  if (!sheet || sheet.getLastRow() < 1) return { headers: [], rows: [] };
  const values = sanitizeValuesForClient(sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).getValues());
  const rawHeaders = values[0].map(h => String(h).trim());
  const blthIdx = findHeaderIdx(rawHeaders, [BLTH_HEADER]);
  const headers = rawHeaders.filter((h, i) => i !== blthIdx);
  const rows = values.slice(1)
    .filter(r => r.some(c => String(c).trim() !== ""))
    .map(r => ({
      blth: blthIdx === -1 ? LEGACY_BLTH : (cellToBlth(r[blthIdx]) || LEGACY_BLTH),
      cells: r.filter((c, i) => i !== blthIdx)
    }));
  return { headers, rows };
}

// Gabungkan dua daftar header (tanpa membedakan huruf besar/kecil)
function unionHeaders(a, b) {
  const out = a.slice();
  b.forEach(h => { if (findHeaderIdx(out, [String(h).trim().toUpperCase()]) === -1) out.push(String(h).trim()); });
  return out.filter(h => h !== "" && h.toUpperCase() !== BLTH_HEADER);
}

function remapRow(cells, fromHeaders, toHeaders) {
  return toHeaders.map(h => {
    const i = findHeaderIdx(fromHeaders, [String(h).trim().toUpperCase()]);
    return i === -1 ? "" : (cells[i] === undefined || cells[i] === null ? "" : cells[i]);
  });
}

function writeAllDataRows(ss, headers, rows) {
  const sheet = ss.getSheetByName(SHEET_DATA_NAME) || ss.insertSheet(SHEET_DATA_NAME);
  const fullHeaders = headers.concat([BLTH_HEADER]);
  const allData = [fullHeaders].concat(rows.map(r => r.cells.concat([r.blth])));
  resetSheetSafely(sheet, allData.length + 5, fullHeaders.length + 2);
  // Kolom BLTH sebagai teks agar "202606" tidak berubah format
  sheet.getRange(1, fullHeaders.length, allData.length, 1).setNumberFormat("@");
  sheet.getRange(1, 1, allData.length, fullHeaders.length).setValues(allData);
  sheet.getRange(1, 1, 1, fullHeaders.length).setBackground("#005c8a").setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);
}

// Daftar unit (Area, ULP, Kordinator) dari SEMUA periode — dipakai untuk daftar login
function buildUnitDirectory(all) {
  const iA = findHeaderIdx(all.headers, ["UP3"]), iU = findHeaderIdx(all.headers, ["ULP"]), iK = findHeaderIdx(all.headers, ["KOORDINATOR", "KORDINATOR"]);
  const areas = {}, ulps = {}, koords = {};
  all.rows.forEach(r => {
    const a = iA === -1 ? "" : String(r.cells[iA] || "").trim().toUpperCase();
    const u = iU === -1 ? "" : String(r.cells[iU] || "").trim().toUpperCase();
    const k = iK === -1 ? "" : String(r.cells[iK] || "").trim();
    if (!a) return;
    areas[a] = true;
    if (u) ulps[a + "||" + u] = [a, u];
    if (u && k) koords[a + "||" + u + "||" + k.toUpperCase()] = [a, u, k];
  });
  return { areas: Object.keys(areas).sort(), ulps: Object.values(ulps), koords: Object.values(koords) };
}

function getPeriodRows(all, blth) {
  return all.rows.filter(r => r.blth === blth).map(r => r.cells);
}

function listDataPeriods(all) {
  return Array.from(new Set(all.rows.map(r => r.blth))).filter(isValidBlth).sort();
}

function recalcAndBuildSheets(ss, store, all, blth) {
  const flat = flattenConfigForPeriod(store, blth);
  const stats = calculateRekapAndReward(all.headers, getPeriodRows(all, blth), flat);
  buildRewardSheet(ss, stats, blth);
  buildSpreadsheetDashboard(ss, stats, blth);
  return stats;
}

/* ============================== ENDPOINT PENYIMPANAN ============================== */

/**
 * Simpan data petugas satu periode.
 * payload = { headers, rows }   → baris periode tsb (tanpa kolom BLTH)
 * rewardConfig                  → konfigurasi datar periode tsb (boleh null untuk periode baru)
 * blth                          → "YYYYMM"
 * scopeUp3 (opsional)           → bila diisi, hanya baris Area tsb yang diganti; Area lain di periode yg sama tetap
 */
function saveUploadedData(payload, rewardConfig, blth, scopeUp3, token) {
  try {
    const target = String(blth || LEGACY_BLTH);
    if (!isValidBlth(target)) throw new Error("Periode (BLTH) tidak valid: " + target);
    if (!payload || !Array.isArray(payload.rows)) throw new Error("Data kosong.");
    const scope = String(scopeUp3 || "").trim().toUpperCase();
    if (!scope && payload.rows.length === 0) throw new Error("Data kosong.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet tidak ditemukan.");

    return withScriptLock(() => {
      const all = readAllDataRows(ss);
      const sess = requireSession(token, all);
      assertCanMutate(sess);
      const sessAreas = sessionAreas(sess);
      if (sessAreas && (!scope || sessAreas.indexOf(scope) === -1)) throw new Error("Akses ditolak: Anda hanya dapat menyimpan data Area di wilayah Anda.");
      const store = getConfigStore();
      assertPeriodUnlocked(store, target);
      const frozen = frozenAreasForPeriod(store, target);
      if (scope && frozen[scope]) throw new Error(frozenMessage(scope, target, frozen[scope]));
      ensurePeriodConfig(store, target);
      assertConfigAllowed(store, all, target, rewardConfig, sess);
      mergeFlatIntoStore(store, rewardConfig, target);
      saveConfigStore(store);

      const normalized = normalizeTableMatrix(payload.headers, payload.rows);
      const inHeaders = normalized.headers.filter(h => h.toUpperCase() !== BLTH_HEADER);
      const inRows = normalized.rows.map(r => remapRow(r, normalized.headers, inHeaders));

      const headers = unionHeaders(all.headers, inHeaders);
      const up3OfOld = cells => { const i = findHeaderIdx(all.headers, ["UP3"]); return i === -1 ? "" : String(cells[i] || "").trim().toUpperCase(); };
      const up3OfNew = cells => { const i = findHeaderIdx(inHeaders, ["UP3"]); return i === -1 ? "" : String(cells[i] || "").trim().toUpperCase(); };

      // Pertahankan: periode lain, (jika scope Area) Area lain di periode yang sama,
      // dan Area yang sedang dalam proses pengajuan / sudah disetujui (dibekukan)
      const frozenSkipped = {};
      const kept = all.rows
        .filter(r => r.blth !== target || (scope && up3OfOld(r.cells) !== scope) || (!scope && frozen[up3OfOld(r.cells)]))
        .map(r => ({ blth: r.blth, cells: remapRow(r.cells, all.headers, headers) }));
      const incoming = inRows
        .filter(cells => !scope || up3OfNew(cells) === scope)
        .filter(cells => { const a = up3OfNew(cells); if (frozen[a]) { frozenSkipped[a] = true; return false; } return true; })
        .map(cells => ({ blth: target, cells: remapRow(cells, inHeaders, headers) }));

      // Realisasi absensi yang sudah di-upload untuk periode ini tetap dipertahankan (berdasarkan NIP)
      const hNip = findHeaderIdx(headers, ["NIP"]), hAbs = findHeaderIdx(headers, [ABSENSI_HEADER]);
      if (hNip !== -1 && hAbs !== -1) {
        const oldAbs = {};
        const oNip = findHeaderIdx(all.headers, ["NIP"]), oAbs = findHeaderIdx(all.headers, [ABSENSI_HEADER]);
        if (oNip !== -1 && oAbs !== -1) {
          all.rows.forEach(r => {
            if (r.blth !== target) return;
            const v = r.cells[oAbs];
            if (v !== "" && v !== null && v !== undefined) oldAbs[String(r.cells[oNip]).trim().toUpperCase()] = v;
          });
        }
        incoming.forEach(r => {
          const cur = r.cells[hAbs];
          const key = String(r.cells[hNip] || "").trim().toUpperCase();
          if ((cur === "" || cur === null || cur === undefined) && oldAbs[key] !== undefined) r.cells[hAbs] = oldAbs[key];
        });
      }

      const merged = { headers, rows: kept.concat(incoming) };
      writeAllDataRows(ss, headers, merged.rows);
      recalcAndBuildSheets(ss, store, merged, target);

      const skipped = Object.keys(frozenSkipped);
      return {
        success: true, blth: target, periods: listDataPeriods(merged), frozenSkipped: skipped,
        message: "Tersimpan di Spreadsheet (" + blthLabel(target) + ")!" + (skipped.length
          ? " Data Area " + skipped.join(", ") + " TIDAK diganti karena sedang dalam proses pengajuan / sudah disetujui." : "")
      };
    });
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Simpan realisasi absensi (%) ke data petugas periode `blth`.
 * items   : [{ nip, pct }]  (pct dalam persen, contoh 119 = 119%)
 * scopeUp3: bila diisi (user Area / MUL), hanya petugas di Area tsb yang diperbarui.
 * Data kinerja periode tsb harus sudah ada. NIP yang tidak ditemukan dilaporkan kembali.
 */
function saveAbsensiData(items, blth, scopeArg, token) {
  try {
    const target = String(blth || "");
    if (!isValidBlth(target)) throw new Error("Periode (BLTH) tidak valid: " + target);
    if (!Array.isArray(items) || !items.length) throw new Error("Data absensi kosong.");
    // scopeArg: "" / null (seluruh Jatim), "NAMA AREA" (format lama), atau { up3, ulp } (Area / Supervisor ULP)
    const sc = (scopeArg && typeof scopeArg === "object")
      ? { up3: String(scopeArg.up3 || "").trim().toUpperCase(), ulp: String(scopeArg.ulp || "").trim().toUpperCase() }
      : { up3: String(scopeArg || "").trim().toUpperCase(), ulp: "" };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet tidak ditemukan.");

    return withScriptLock(() => {
      const all = readAllDataRows(ss);
      const sess = requireSession(token, all);
      assertCanMutate(sess);
      // Wilayah upload ditentukan oleh akun yang login, bukan oleh browser
      if (sess.role === "KOORD") throw new Error("Akses ditolak: Kordinator tidak dapat meng-upload absensi.");
      if (sess.role === "ULP") { sc.up3 = sess.up3; sc.ulp = sess.ulp; }
      else if (sess.role === "MUL" || sess.role === "AREA") { sc.up3 = sess.up3; }
      const scopeLabel = sc.ulp ? " di ULP " + sc.ulp : (sc.up3 ? " di Area " + sc.up3 : "");
      const store = getConfigStore();
      assertPeriodUnlocked(store, target);
      const frozen = frozenAreasForPeriod(store, target);
      const periodRows = all.rows.filter(r => r.blth === target);
      if (!periodRows.length) {
        throw new Error("Data kinerja periode " + blthLabel(target) + " belum ada. Upload file Kinerja periode tersebut terlebih dahulu.");
      }
      const headers = unionHeaders(all.headers, [ABSENSI_HEADER]);
      const rows = all.rows.map(r => ({ blth: r.blth, cells: remapRow(r.cells, all.headers, headers) }));
      const iNip = findHeaderIdx(headers, ["NIP"]), iUp3 = findHeaderIdx(headers, ["UP3"]), iUlp = findHeaderIdx(headers, ["ULP"]), iAbs = findHeaderIdx(headers, [ABSENSI_HEADER]);
      if (iNip === -1) throw new Error("Kolom NIP tidak ditemukan pada data petugas.");

      const absByNip = {};
      items.forEach(it => {
        const nip = String(it && it.nip || "").trim().toUpperCase();
        const pct = parseAbsensiCell(it && it.pct);
        if (nip && pct !== null) absByNip[nip] = pct;
      });

      const matched = {}, frozenSkip = {};
      let outOfScope = 0;
      rows.forEach(r => {
        if (r.blth !== target) return;
        const nip = String(r.cells[iNip] || "").trim().toUpperCase();
        if (absByNip[nip] === undefined) return;
        const U = i => i === -1 ? "" : String(r.cells[i] || "").trim().toUpperCase();
        if ((sc.up3 && U(iUp3) !== sc.up3) || (sc.ulp && U(iUlp) !== sc.ulp)) { outOfScope++; return; }
        if (frozen[U(iUp3)]) { frozenSkip[U(iUp3)] = (frozenSkip[U(iUp3)] || 0) + 1; matched[nip] = "frozen"; return; }
        r.cells[iAbs] = absByNip[nip];
        matched[nip] = true;
      });
      const notFound = Object.keys(absByNip).filter(n => !matched[n]);
      const updated = Object.keys(matched).filter(n => matched[n] === true).length;
      const frozenAreas = Object.keys(frozenSkip);
      if (!updated) {
        if (frozenAreas.length) throw new Error(frozenMessage(frozenAreas[0], target, frozen[frozenAreas[0]]));
        throw new Error("Tidak ada NIP pada file absensi yang cocok dengan data petugas periode " + blthLabel(target) +
          scopeLabel + ".");
      }

      writeAllDataRows(ss, headers, rows);
      recalcAndBuildSheets(ss, store, { headers: headers, rows: rows }, target);
      return {
        success: true, blth: target, updated: updated, notFound: notFound.slice(0, 50), notFoundCount: notFound.length,
        outOfScope: outOfScope, periods: listDataPeriods({ headers: headers, rows: rows }), frozenSkipped: frozenAreas,
        message: "Realisasi absensi " + updated + " petugas periode " + blthLabel(target) + " tersimpan." +
          (frozenAreas.length ? " Petugas di Area " + frozenAreas.join(", ") + " dilewati karena sedang dalam proses pengajuan / sudah disetujui." : "")
      };
    });
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveRewardCalculation(rewardConfig, blth, token) {
  try {
    const target = String(blth || LEGACY_BLTH);
    if (!isValidBlth(target)) throw new Error("Periode (BLTH) tidak valid: " + target);
    return withScriptLock(() => {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const all = ss ? readAllDataRows(ss) : { headers: [], rows: [] };
      const sess = requireSession(token, all);
      assertCanMutate(sess);
      const store = getConfigStore();
      assertPeriodUnlocked(store, target);
      assertConfigAllowed(store, all, target, rewardConfig, sess);
      mergeFlatIntoStore(store, rewardConfig, target);
      saveConfigStore(store);

      if (!ss) return { success: true };
      if (getPeriodRows(all, target).length === 0) return { success: true };
      recalcAndBuildSheets(ss, store, all, target);
      return { success: true, message: "Perhitungan " + blthLabel(target) + " diperbarui!" };
    });
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/* ============================== KUNCI PERIODE ============================== */

/**
 * Kunci / buka kunci periode. Hanya user UP 3 Jatim: password akun UP diverifikasi di server.
 */
function setPeriodLock(blth, lock, upPassword) {
  try {
    const target = String(blth || "");
    if (!isValidBlth(target)) return { success: false, message: "Periode tidak valid." };
    const passMap = getUserPasswordsMap();
    const plain = String(upPassword || "").trim();
    // Diterima password akun UP 3 Jatim atau akun admin (kewenangan setara UP)
    const byKey = UP_LEVEL_ACCOUNT_KEYS.find(k => isPasswordMatch(passMap[k], plain));
    if (!byKey) {
      return { success: false, message: "Password tidak sesuai. Masukkan password akun UP 3 Jatim atau admin yang sedang dipakai." };
    }
    return withScriptLock(() => {
      const store = getConfigStore();
      if (lock) {
        ensurePeriodConfig(store, target);
        store.lockedPeriods[target] = { at: new Date().toISOString(), by: byKey };
      } else {
        delete store.lockedPeriods[target];
      }
      saveConfigStore(store);
      // Pastikan sheet Data_Reward berisi hasil final periode yang dikunci
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss && lock) {
        const all = readAllDataRows(ss);
        if (getPeriodRows(all, target).length > 0) recalcAndBuildSheets(ss, store, all, target);
      }
      return {
        success: true,
        lockedPeriods: store.lockedPeriods,
        message: "Periode " + blthLabel(target) + (lock ? " berhasil DIKUNCI." : " berhasil DIBUKA kembali.")
      };
    });
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Rekap reward per bulan untuk semua periode yang ada datanya.
 * scopeUp3 kosong → semua Area (untuk UP); diisi → hanya Area tsb (untuk MUL / SPV Area).
 * Hasil per periode per Area: Nilai Reward, Denda, Netto, Netto 70% (pagu),
 * total reward SPV Area / SPV ULP / Kordinator / Petugas beserta jumlah penerimanya.
 */
/* ============================== HISTORI PETUGAS ============================== */

// scope dari browser: { all: true } atau { up3, ulp, koord } (huruf besar). Membatasi data sesuai wilayah user.
function normalizeScope(scope) {
  const s = scope && typeof scope === "object" ? scope : {};
  const U = v => String(v || "").trim().toUpperCase();
  if (s.all) return { all: true };
  return { up3: U(s.up3), ulp: U(s.ulp), koord: U(s.koord) };
}

function rowInScope(cells, idx, scope) {
  if (scope.all) return true;
  const U = i => i === -1 ? "" : String(cells[i] || "").trim().toUpperCase();
  if (scope.up3 && U(idx.up3) !== scope.up3) return false;
  if (scope.ulp && U(idx.ulp) !== scope.ulp) return false;
  if (scope.koord && U(idx.koord) !== scope.koord) return false;
  return true;
}

function dataColumnIdx(headers) {
  return {
    nip: findHeaderIdx(headers, ["NIP"]),
    nama: findHeaderIdx(headers, ["PETUGAS"]),
    koord: findHeaderIdx(headers, ["KOORDINATOR", "KORDINATOR"]),
    ulp: findHeaderIdx(headers, ["ULP"]),
    up3: findHeaderIdx(headers, ["UP3"])
  };
}

const HISTORY_JABATAN_RANK = { "SPV MANBILL AREA": 1, "SPV MANBILL ULP": 2, "KORDINATOR": 3, "PETUGAS MANBILL": 4 };

// Penerima (dari hasil perhitungan) yang boleh dilihat user sesuai wilayahnya
function recipientInScope(rec, sc) {
  if (sc.all) return true;
  const U = v => String(v || "").trim().toUpperCase();
  if (sc.up3 && U(rec.up3) !== sc.up3) return false;
  if (rec.jabatan === "SPV MANBILL AREA" || rec.jabatan === "ADMIN MANBILL") return !sc.ulp;
  if (sc.ulp && U(rec.ulp) !== sc.ulp) return false;
  if (rec.jabatan === "SPV MANBILL ULP") return !sc.koord;
  const koord = rec.jabatan === "KORDINATOR" ? rec.nama : rec.koordinator;
  if (sc.koord && U(koord) !== sc.koord) return false;
  return true;
}

/**
 * Cari Petugas, Kordinator, SPV ULP, dan SPV Area di SEMUA periode.
 * - Petugas   : NIP / Nama / ULP
 * - Kordinator: Nama Kordinator / ULP
 * - SPV ULP   : Nama SPV / ULP
 * - SPV Area  : Nama SPV / Area
 * jabatanFilter (opsional): "PETUGAS MANBILL" | "KORDINATOR" | "SPV MANBILL ULP" | "SPV MANBILL AREA".
 * Hasil unik per ID penerima, maksimal 50.
 */
function searchPetugas(query, scope, jabatanFilter) {
  try {
    const q = String(query || "").trim().toLowerCase();
    if (q.length < 2) return { success: false, message: "Ketik minimal 2 karakter NIP, Nama, atau ULP." };
    const jf = String(jabatanFilter || "").trim().toUpperCase();
    const sc = normalizeScope(scope);
    const store = getConfigStore();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: true, results: [], total: 0 };
    const all = readAllDataRows(ss);
    const idx = dataColumnIdx(all.headers);

    const found = {};
    const add = (id, info, blth) => {
      const cur = found[id];
      if (!cur) found[id] = Object.assign({ id: id, lastBlth: blth, periods: {} }, info);
      else if (blth > cur.lastBlth) Object.assign(cur, info, { lastBlth: blth });
      found[id].periods[blth] = true;
    };
    const has = (...vals) => vals.some(v => String(v || "").toLowerCase().indexOf(q) !== -1);

    all.rows.forEach(r => {
      if (!rowInScope(r.cells, idx, sc)) return;
      const get = i => i === -1 ? "" : String(r.cells[i] || "").trim();
      const nip = get(idx.nip), nama = get(idx.nama);
      const ulp = get(idx.ulp).toUpperCase() || "TIDAK TERDEFINISI";
      const up3 = get(idx.up3).toUpperCase() || "TIDAK TERDEFINISI";
      const koord = get(idx.koord) || "TANPA KOORDINATOR";

      if ((!jf || jf === "PETUGAS MANBILL") && nip && has(nip, nama, ulp)) {
        add(nip, { jabatan: "PETUGAS MANBILL", nip: nip, nama: nama, koordinator: koord, ulp: ulp, up3: up3 }, r.blth);
      }
      if ((!jf || jf === "KORDINATOR") && has(koord, ulp)) {
        add(makeKoordId(up3, ulp, koord), { jabatan: "KORDINATOR", nip: "", nama: koord, koordinator: koord, ulp: ulp, up3: up3 }, r.blth);
      }
      if ((!jf || jf === "SPV MANBILL ULP") && !sc.koord) {
        const nm = String(store.spvUlpNames[up3 + "||" + ulp] || "").trim() || ("SPV MANBILL ULP " + ulp);
        if (has(nm, ulp)) add(makeSpvUlpId(up3, ulp), { jabatan: "SPV MANBILL ULP", nip: "", nama: nm, koordinator: "-", ulp: ulp, up3: up3 }, r.blth);
      }
      if ((!jf || jf === "SPV MANBILL AREA") && !sc.ulp) {
        const nm = String(store.spvAreaNames[up3] || "").trim() || ("SPV MANBILL AREA " + up3);
        if (has(nm, up3)) add(makeSpvAreaId(up3), { jabatan: "SPV MANBILL AREA", nip: "", nama: nm, koordinator: "-", ulp: "AREA " + up3, up3: up3 }, r.blth);
      }
    });

    const list = Object.values(found).map(x => ({
      id: x.id, jabatan: x.jabatan, nip: x.nip, nama: x.nama, koordinator: x.koordinator, ulp: x.ulp, up3: x.up3,
      lastBlth: x.lastBlth, lastLabel: blthLabel(x.lastBlth), periodCount: Object.keys(x.periods).length
    }));
    list.sort((a, b) => {
      const ea = (a.nip && a.nip.toLowerCase() === q) ? 0 : 1, eb = (b.nip && b.nip.toLowerCase() === q) ? 0 : 1;
      return ea - eb || HISTORY_JABATAN_RANK[a.jabatan] - HISTORY_JABATAN_RANK[b.jabatan] || a.nama.localeCompare(b.nama);
    });
    return { success: true, total: list.length, results: list.slice(0, 50) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Histori nilai kinerja & reward satu penerima per periode.
 * id = NIP (Petugas) atau ID unit: "KOORD|UP3|ULP|NAMA", "SPV-ULP|UP3|ULP", "SPV-AREA|UP3".
 * Untuk Kordinator / SPV disertakan info tim: jumlah petugas & yang Kinerja MO >= 100%.
 */
function getPetugasHistory(id, scope) {
  try {
    const target = String(id || "").trim();
    if (!target) return { success: false, message: "ID penerima tidak valid." };
    const sc = normalizeScope(scope);
    const store = getConfigStore();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: "Spreadsheet tidak ditemukan." };
    const all = readAllDataRows(ss);

    const history = listDataPeriods(all).map(b => {
      const stats = calculateRekapAndReward(all.headers, getPeriodRows(all, b), flattenConfigForPeriod(store, b));
      const r = (stats.recipients || []).find(x => String(x.nip) === target && x.jabatan !== "ADMIN MANBILL");
      if (!r || !recipientInScope(r, sc)) return null;
      const bd = r.breakdown || {};
      const area = (stats.rekapUP3 || []).find(u => u.up3 === r.up3) || {};

      // Tim untuk Kordinator / SPV
      let team = null;
      if (r.jabatan !== "PETUGAS MANBILL") {
        const members = (stats.recipients || []).filter(x => x.jabatan === "PETUGAS MANBILL" && x.up3 === r.up3 &&
          (r.jabatan === "SPV MANBILL AREA" || x.ulp === r.ulp) &&
          (r.jabatan !== "KORDINATOR" || String(x.koordinator || "").toUpperCase() === String(r.nama || "").toUpperCase()));
        team = {
          count: members.length,
          moOk: members.filter(x => Number((x.breakdown || {}).rawU1) >= 100).length,
          approved: members.filter(x => x.status === "approved").length
        };
      }
      return {
        blth: b, label: blthLabel(b), locked: Boolean(store.lockedPeriods[b]),
        jabatan: r.jabatan, nama: r.nama, koordinator: r.koordinator, ulp: r.ulp, up3: r.up3, status: r.status,
        rawMO: bd.rawU1, nilaiMO: bd.u1, rawSLA: bd.rawU2, nilaiSLA: bd.u2, nilaiK3: bd.u3,
        subUtama: bd.subUtama, tambahan: bd.tambahan, absensi: bd.absensi, subAtasan: bd.subAtasan,
        score: r.score, koef: r.koefJabatan, bobot: r.bobot,
        indeks: Math.round(r.nilaiPerBobot || 0), reward: Math.round(r.reward || 0),
        slaArea: area.slaArea, paguArea: Math.round(area.paguArea || 0), team: team
      };
    }).filter(Boolean);

    if (!history.length) return { success: false, message: "Data penerima tidak ditemukan di wilayah Anda." };
    const last = history[history.length - 1];
    return {
      success: true,
      profile: {
        id: target, jabatan: last.jabatan, nip: last.jabatan === "PETUGAS MANBILL" ? target : "",
        nama: last.nama, koordinator: last.koordinator, ulp: last.ulp, up3: last.up3
      },
      history: history
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getMonthlyRecap(scopeUp3) {
  try {
    const scope = String(scopeUp3 || "").trim().toUpperCase();
    const store = getConfigStore();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: true, periods: [] };
    const all = readAllDataRows(ss);
    const countKey = { "SPV MANBILL AREA": "nSpvArea", "SPV MANBILL ULP": "nSpvUlp", "KORDINATOR": "nKoord", "PETUGAS MANBILL": "nPetugas" };

    const periods = listDataPeriods(all).map(b => {
      const flat = flattenConfigForPeriod(store, b);
      const stats = calculateRekapAndReward(all.headers, getPeriodRows(all, b), flat);
      const areas = {};
      (stats.rekapUP3 || []).forEach(u => {
        if (scope && u.up3 !== scope) return;
        areas[u.up3] = {
          gross: Math.round(u.rewardKotor || 0), denda: Math.round(u.dendaArea || 0),
          netto: Math.round(u.rewardNetto || 0), pagu: Math.round(u.paguArea || 0),
          spvArea: Math.round(u.rewardTLArea || 0), spvUlp: Math.round(u.rewardTLULP || 0),
          koord: Math.round(u.rewardKoordinator || 0), petugas: Math.round(u.rewardPetugas || 0),
          nSpvArea: 0, nSpvUlp: 0, nKoord: 0, nPetugas: 0
        };
      });
      (stats.recipients || []).forEach(r => {
        const a = areas[r.up3], k = countKey[r.jabatan];
        if (a && k && r.bobot > 0) a[k]++;
      });
      return { blth: b, label: blthLabel(b), locked: Boolean(store.lockedPeriods[b]), areas: areas };
    });
    return { success: true, periods: periods };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// Ambil pengaturan (datar) suatu periode, dipakai untuk "Salin pengaturan dari bulan sebelumnya"
function getPeriodSettings(blth) {
  try {
    const target = String(blth || "");
    if (!isValidBlth(target)) return { success: false, message: "Periode tidak valid." };
    const store = getConfigStore();
    if (!store.periods[target]) return { success: false, message: "Pengaturan periode " + blthLabel(target) + " belum ada." };
    return { success: true, blth: target, rewardConfig: flattenConfigForPeriod(store, target) };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/* ============================== BUKU PANDUAN (v2.2) ==============================
 * File PDF Buku Panduan disimpan di sheet tersembunyi "_Panduan" (base64, dipecah per sel),
 * sehingga TIDAK memerlukan izin Google Drive dan selalu dapat dibaca siapa pun yang membuka
 * aplikasi (termasuk dari halaman login). Hanya akun admin yang dapat mengunggah.
 * (Versi 2.1 menyimpan di Google Drive; sejak 2.2 cukup unggah ulang sekali melalui aplikasi.)
 * ================================================================================ */
const SHEET_GUIDE_NAME = "_Panduan";
const GUIDE_FILE_PROP = "GUIDE_FILE_ID";   // sisa v2.1 (Drive), dihapus saat unggah ulang
const GUIDE_CHUNK = 45000;                 // batas aman karakter per sel (maks 50.000)
const GUIDE_MAX_BYTES = 10 * 1024 * 1024;

// Baca dari sheet: { meta, base64 } atau null
function readGuideFromSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss ? ss.getSheetByName(SHEET_GUIDE_NAME) : null;
  if (!sh || sh.getLastRow() < 2) return null;
  const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0] || ""));
  let meta = {};
  try { meta = JSON.parse(vals[0] || "{}"); } catch (e) { meta = {}; }
  const base64 = vals.slice(1).join("");
  if (!base64) return null;
  return { meta: meta, base64: base64 };
}

function writeGuideToSheet_(meta, base64) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Spreadsheet tidak ditemukan.");
  let sh = ss.getSheetByName(SHEET_GUIDE_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_GUIDE_NAME);
    try { sh.hideSheet(); } catch (e) {}
  }
  const rows = [[JSON.stringify(meta)]];
  for (let i = 0; i < base64.length; i += GUIDE_CHUNK) rows.push([base64.substring(i, i + GUIDE_CHUNK)]);
  ensureSheetDimensions(sh, rows.length + 1, 1);
  sh.getRange(1, 1, sh.getMaxRows(), 1).clearContent();
  const rng = sh.getRange(1, 1, rows.length, 1);
  rng.setNumberFormat("@");
  rng.setValues(rows);
}

function loadGuide_() {
  return readGuideFromSheet_();
}

/** Info Buku Panduan (tanpa login). */
function getGuideInfo() {
  try {
    const g = loadGuide_();
    if (!g) return { success: true, available: false };
    const m = g.meta || {};
    return { success: true, available: true, name: m.name || "Buku_Panduan_APIK.pdf", size: m.size || Math.floor(g.base64.length * 3 / 4),
      updatedAt: m.at || "", updatedBy: m.by || "", version: m.version || "" };
  } catch (err) {
    return { success: false, message: "Gagal membaca Buku Panduan: " + errText(err) };
  }
}

/** Isi file Buku Panduan (base64) untuk diunduh browser (tanpa login). */
function downloadGuide() {
  try {
    const g = loadGuide_();
    if (!g) return { success: false, notUploaded: true, message: "Buku Panduan belum diunggah. Minta admin APIK mengunggahnya melalui tombol Buku Panduan setelah login." };
    return { success: true, name: (g.meta && g.meta.name) || "Buku_Panduan_APIK.pdf", mimeType: "application/pdf", base64: g.base64 };
  } catch (err) {
    return { success: false, message: "Gagal membaca Buku Panduan: " + errText(err) };
  }
}

/** Unggah / ganti Buku Panduan (khusus akun admin). */
function uploadGuide(token, fileName, base64, versionLabel) {
  try {
    const sess = requireSession(token);
    if (!sess.isAdmin) throw new Error("Akses ditolak: hanya akun admin yang dapat mengunggah Buku Panduan.");
    const b64 = String(base64 || "").replace(/\s+/g, "");
    const bytes = Utilities.base64Decode(b64);
    if (!bytes || bytes.length < 5) throw new Error("File kosong.");
    if (bytes.length > GUIDE_MAX_BYTES) throw new Error("Ukuran file melebihi 10 MB.");
    const head = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
    if (head !== "%PDF") throw new Error("File harus berformat PDF.");
    let name = String(fileName || "Buku_Panduan_APIK.pdf").replace(/[\\/:*?"<>|]+/g, "_").trim();
    if (!/\.pdf$/i.test(name)) name += ".pdf";
    return withScriptLock(() => {
      const meta = { name: name, size: bytes.length, at: new Date().toISOString(), by: sess.label, version: String(versionLabel || "").trim().substring(0, 60) };
      writeGuideToSheet_(meta, b64);
      try { PropertiesService.getScriptProperties().deleteProperty(GUIDE_FILE_PROP); } catch (e) {}
      return { success: true, message: "Buku Panduan berhasil diunggah: " + name, info: getGuideInfo() };
    });
  } catch (err) {
    return { success: false, message: errText(err) };
  }
}

/* ============================== API PORTAL AREK MANBILL (v2.4) ==============================
 * Dipanggil halaman portal (situs statis di domain sendiri) melalui JSONP:
 *   <URL /exec>?api=portal&callback=fn  → ringkasan agregat APIK + pengumuman + daftar aplikasi
 *   <URL /exec>?api=guide&callback=fn   → file Buku Panduan (base64)
 * Hanya angka gabungan (tanpa nama / nilai per orang). Hasil di-cache 5 menit.
 * Pengumuman & daftar aplikasi portal dikelola di sheet "Portal_Pengumuman" dan "Portal_Aplikasi"
 * (dibuat otomatis berisi contoh saat portal pertama kali dibuka).
 * ======================================================================================= */
const PORTAL_SHOW_REWARD_TOTAL = true;   // false = portal tidak menampilkan nominal Rupiah
const PORTAL_CACHE_KEY = "PORTAL_DATA_V1";
const PORTAL_CACHE_SEC = 300;
const SHEET_PORTAL_NEWS = "Portal_Pengumuman";
const SHEET_PORTAL_APPS = "Portal_Aplikasi";
const PORTAL_NEWS_HEADERS = ["Tanggal", "Judul", "Isi", "Penting (Ya/Tidak)", "Tampil (Ya/Tidak)", "Tautan (opsional)"];
const PORTAL_APPS_HEADERS = ["Urutan", "Nama Aplikasi", "Keterangan", "Kategori", "Status (Aktif / Segera Hadir / Pemeliharaan)", "URL", "Ikon (Font Awesome)", "Warna (hex)", "Buka di Tab Baru (Ya/Tidak)", "Label Baru (Ya/Tidak)"];
const PORTAL_DEFAULT_APPS = [
  [1, "APIK Manbill", "Penilaian & Insentif Kinerja", "Kinerja & Reward", "Aktif", "apik/", "fa-ranking-star", "#0b5a7a", "Tidak", "Tidak"],
  [2, "Buku Panduan APIK", "Unduh panduan penggunaan (PDF)", "Dokumen", "Aktif", "#panduan", "fa-book-open", "#b45309", "Tidak", "Ya"],
  [3, "Absensi Rapat", "Kehadiran, Notulen & Eviden Rapat", "Operasional", "Segera Hadir", "", "fa-clipboard-user", "#0f766e", "Ya", "Tidak"],
  [4, "Monitoring SLA", "SLA Kontrak Manbill per Area", "Laporan", "Segera Hadir", "", "fa-gauge-high", "#1d4ed8", "Ya", "Tidak"],
  [5, "Laporan Bulanan", "Rekap Kinerja & Reward", "Laporan", "Segera Hadir", "", "fa-chart-column", "#7c3aed", "Ya", "Tidak"],
  [6, "Arsip Dokumen", "Dokumen & Berkas Manbill", "Dokumen", "Segera Hadir", "", "fa-box-archive", "#7c4a1e", "Ya", "Tidak"]
];

function portalApiResponse_(api, p) {
  let data;
  try {
    if (api === "portal") data = getPortalData_(String(p.refresh || "") === "1");
    else if (api === "guide") data = downloadGuide();
    else data = { success: false, message: "API tidak dikenal." };
  } catch (err) {
    data = { success: false, message: errText(err) };
  }
  const json = JSON.stringify(data);
  const cb = String(p.callback || "");
  if (/^[A-Za-z_$][\w$.]{0,60}$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + json + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function portalYes_(v) { return /^(ya|y|yes|true|1)$/i.test(String(v || "").trim()); }

function ensurePortalSheet_(ss, name, headers, sampleRows) {
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  const rows = [headers].concat(sampleRows || []);
  ensureSheetDimensions(sh, rows.length + 5, headers.length);
  sh.getRange(1, 1, rows.length, headers.length).setValues(rows);
  sh.getRange(1, 1, 1, headers.length).setBackground("#003b5c").setFontColor("#ffffff").setFontWeight("bold");
  sh.setFrozenRows(1);
  return sh;
}

function readPortalNews_(ss) {
  const today = Utilities.formatDate(new Date(), APP_TZ, "yyyy-MM-dd");
  const sh = ensurePortalSheet_(ss, SHEET_PORTAL_NEWS, PORTAL_NEWS_HEADERS, [
    [today, "Portal AREK Manbill resmi digunakan", "AREK Manbill (Aplikasi Rekap & Evaluasi Kinerja Manbill): seluruh aplikasi pendukung Manbill Optimization UP 3 Jatim kini dapat diakses dari satu halaman.", "Tidak", "Ya", ""],
    [today, "Buku Panduan APIK tersedia", "Unduh Buku Panduan APIK terbaru melalui kartu Buku Panduan APIK.", "Tidak", "Ya", ""]
  ]);
  if (sh.getLastRow() < 2) return [];
  return sanitizeValuesForClient(sh.getRange(2, 1, sh.getLastRow() - 1, PORTAL_NEWS_HEADERS.length).getValues())
    .filter(r => String(r[1] || "").trim() && portalYes_(r[4] === "" ? "Ya" : r[4]))
    .map(r => {
      const d = r[0] instanceof Date ? Utilities.formatDate(r[0], APP_TZ, "yyyy-MM-dd") : String(r[0] || "").trim();
      return { date: d, title: String(r[1]).trim(), body: String(r[2] || "").trim(), important: portalYes_(r[3]), link: String(r[5] || "").trim() };
    })
    .sort((a, b) => (b.important - a.important) || b.date.localeCompare(a.date))
    .slice(0, 8);
}

function readPortalApps_(ss) {
  const sh = ensurePortalSheet_(ss, SHEET_PORTAL_APPS, PORTAL_APPS_HEADERS, PORTAL_DEFAULT_APPS);
  if (sh.getLastRow() < 2) return null;
  const apps = sanitizeValuesForClient(sh.getRange(2, 1, sh.getLastRow() - 1, PORTAL_APPS_HEADERS.length).getValues())
    .filter(r => String(r[1] || "").trim())
    .map((r, i) => ({
      order: Number(r[0]) || (i + 1), name: String(r[1]).trim(), desc: String(r[2] || "").trim(), category: String(r[3] || "Lainnya").trim(),
      status: String(r[4] || "Aktif").trim(), url: String(r[5] || "").trim(), icon: String(r[6] || "fa-cube").trim(),
      color: String(r[7] || "#0b5a7a").trim(), newTab: portalYes_(r[8]), isNew: portalYes_(r[9])
    }))
    .sort((a, b) => a.order - b.order);
  return apps.length ? apps : null;
}

function getPortalData_(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const hit = cache.get(PORTAL_CACHE_KEY);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const store = getConfigStore();
  const nowM = monthKeyNow();
  const subs = Object.keys(store.submissions || {}).map(k => store.submissions[k]).filter(Boolean);
  const monthOf = iso => iso ? Utilities.formatDate(new Date(iso), APP_TZ, "yyyyMM") : "";
  const approvedNow = subs.filter(s => FINAL_STATUSES.indexOf(s.status) !== -1 && s.batch === nowM);
  const processedNow = subs.filter(s => s.status === "DIPROSES" && monthOf(s.sdmAt) === nowM);
  const sumReward = list => list.reduce((t, s) => t + (Number(s.totals && s.totals.reward) || 0), 0);

  let latest = "", areaCount = 0;
  if (ss) {
    const all = readAllDataRows(ss);
    const periods = listDataPeriods(all);
    latest = periods.length ? periods[periods.length - 1] : "";
    if (latest) {
      const idx = dataColumnIdx(all.headers);
      areaCount = new Set(getPeriodRows(all, latest).map(c => idx.up3 === -1 ? "" : String(c[idx.up3] || "").trim().toUpperCase()).filter(Boolean)).size;
    }
  }
  let lastActivity = "";
  subs.forEach(s => (s.history || []).forEach(h => { if (h && h.at && h.at > lastActivity) lastActivity = h.at; }));

  let news = [], apps = null;
  if (ss) {
    try { news = readPortalNews_(ss); } catch (e) {}
    try { apps = readPortalApps_(ss); } catch (e) {}
  }
  const data = {
    success: true,
    generatedAt: new Date().toISOString(),
    month: nowM, monthLabel: blthLabel(nowM),
    showAmount: PORTAL_SHOW_REWARD_TOTAL,
    stats: {
      latestPeriod: latest, latestPeriodLabel: latest ? blthLabel(latest) : "-", areaCount: areaCount,
      waitingMul: subs.filter(s => s.status === "MENUNGGU_MUL").length,
      waitingUp: subs.filter(s => s.status === "MENUNGGU_UP").length,
      returned: subs.filter(s => s.status === "DIKEMBALIKAN").length,
      readyForSdm: subs.filter(s => s.status === "DISETUJUI").length,
      approvedThisMonth: approvedNow.length,
      approvedThisMonthReward: PORTAL_SHOW_REWARD_TOTAL ? Math.round(sumReward(approvedNow)) : null,
      processedThisMonth: processedNow.length,
      processedThisMonthReward: PORTAL_SHOW_REWARD_TOTAL ? Math.round(sumReward(processedNow)) : null,
      lastActivity: lastActivity
    },
    news: news,
    apps: apps,
    guide: (() => { const g = getGuideInfo(); return g && g.available ? { name: g.name, updatedAt: g.updatedAt, version: g.version } : null; })()
  };
  try { cache.put(PORTAL_CACHE_KEY, JSON.stringify(data), PORTAL_CACHE_SEC); } catch (e) {}
  return data;
}

/* ============================== PENGAJUAN & APPROVAL REWARD (v2.0) ==============================
 * Alur per Area per BLTH:
 *   (Belum Diajukan) --Area ajukan--> MENUNGGU_MUL --MUL setuju--> MENUNGGU_UP --UP setuju--> DISETUJUI --SDM--> DIPROSES
 *   Area tanpa MUL (Banyuwangi, Situbondo, Jember, Kediri): Area ajukan --> MENUNGGU_UP.
 *   MUL / UP dapat MENGEMBALIKAN (wajib catatan) → DIKEMBALIKAN (data terbuka lagi untuk revisi).
 *   Sejak diajukan, data Area pada periode tsb DIBEKUKAN (tidak dapat diubah siapa pun).
 *   Saat UP menyetujui: nominal reward per penerima disimpan (snapshot) ke sheet "Reward_Approved"
 *   dan Bulan Proses SDM = bulan tanggal approval UP. Bila seluruh Area di periode tsb sudah disetujui,
 *   periode otomatis terkunci penuh.
 * ============================================================================================== */

const SUBMISSION_STATUS = {
  BELUM: "Belum Diajukan",
  MENUNGGU_MUL: "Menunggu Approval MUL",
  MENUNGGU_UP: "Menunggu Approval UP 3",
  DISETUJUI: "Disetujui UP 3 (Siap Diproses SDM)",
  DIPROSES: "Sudah Diproses SDM",
  DIKEMBALIKAN: "Dikembalikan / Perlu Revisi"
};
const FROZEN_STATUSES = ["MENUNGGU_MUL", "MENUNGGU_UP", "DISETUJUI", "DIPROSES"];
const FINAL_STATUSES = ["DISETUJUI", "DIPROSES"];
const SHEET_APPROVED_NAME = "Reward_Approved";
const SHEET_LOG_NAME = "Log_Pengajuan";
const APP_TZ = "Asia/Jakarta";
const APPROVED_HEADERS = ["Bulan Proses", "Periode Bulan Proses", "BLTH", "Periode", "Area / UP3", "ULP", "NIP", "ID Penerima",
  "Nama Penerima", "Jabatan", "Kordinator", "Total Nilai (%)", "Bobot", "Indeks / Bobot Area (Rp)", "Reward (Rp)", "Disetujui UP 3", "Kunci"];
const APPROVED_KEY_COL = APPROVED_HEADERS.length - 1;
const JABATAN_ORDER = { "SPV MANBILL AREA": 1, "SPV MANBILL ULP": 2, "KORDINATOR": 3, "PETUGAS MANBILL": 4, "ADMIN MANBILL": 5 };

function submissionKey(blth, area) { return blth + "|" + area; }
function splitSubmissionKey(k) {
  const s = String(k || ""), i = s.indexOf("|");
  return i === -1 ? { blth: s, area: "" } : { blth: s.substring(0, i), area: s.substring(i + 1) };
}
function areaHasMul(area) { return AREAS_WITHOUT_MUL.indexOf(String(area || "").toUpperCase()) === -1; }
function monthKeyNow() { return Utilities.formatDate(new Date(), APP_TZ, "yyyyMM"); }
function fmtDateTime(iso) { return iso ? Utilities.formatDate(new Date(iso), APP_TZ, "dd-MM-yyyy HH:mm") : ""; }

function frozenAreasForPeriod(store, blth) {
  const m = {};
  Object.keys(store.submissions || {}).forEach(k => {
    const s = store.submissions[k];
    if (s && s.blth === blth && FROZEN_STATUSES.indexOf(s.status) !== -1) m[s.area] = s.status;
  });
  return m;
}

function frozenMessage(area, blth, status) {
  if (FINAL_STATUSES.indexOf(status) !== -1) {
    return "Reward Area " + area + " periode " + blthLabel(blth) + " sudah disetujui UP 3 dan DIKUNCI (" + (SUBMISSION_STATUS[status] || status) +
      "). Data tidak dapat diubah" + (status === "DISETUJUI" ? "; bila perlu revisi, UP 3 dapat membatalkan persetujuan." : ".");
  }
  return "Reward Area " + area + " periode " + blthLabel(blth) + " sedang dalam proses pengajuan (" + (SUBMISSION_STATUS[status] || status) +
    "), sehingga datanya tidak dapat diubah. Bila perlu revisi, minta MUL / UP 3 mengembalikan pengajuan.";
}

/**
 * Periksa konfigurasi yang akan disimpan: wilayah user, status beku, dan kewenangan.
 * Kunci di luar wilayah / tidak berubah dibuang diam-diam; perubahan pada Area yang dibekukan ditolak.
 */
function assertConfigAllowed(store, all, blth, flat, sess) {
  if (!flat || typeof flat !== "object") return;
  const frozen = frozenAreasForPeriod(store, blth);
  const areas = sessionAreas(sess);
  const role = sess.role;
  const idx = dataColumnIdx(all.headers);
  const nipArea = {};
  if (idx.nip !== -1) {
    all.rows.forEach(r => {
      if (r.blth !== blth) return;
      nipArea[String(r.cells[idx.nip] || "").trim().toUpperCase()] = idx.up3 === -1 ? "" : String(r.cells[idx.up3] || "").trim().toUpperCase();
    });
  }
  const areaOf = (m, k) => {
    const key = String(k);
    if (m === "ulpSlaScores" || m === "spvUlpNames") return key.split("||")[0].toUpperCase();
    if (m === "customScores" || m === "unitNips") {
      if (/^(KOORD|SPV-ULP|SPV-AREA|ADMIN)\|/.test(key)) return key.split("|")[1];
      return nipArea[key.toUpperCase()] || "";
    }
    return key.toUpperCase();
  };
  const period = store.periods[blth] || {};
  const same = (m, k, v) => {
    const src = PERIOD_MAP_KEYS.indexOf(m) !== -1 ? (period[m] || {}) : (store[m] || {});
    return JSON.stringify(src[k]) === JSON.stringify(v);
  };
  const onlyCustom = role === "ULP" || role === "KOORD";

  CONFIG_MAP_KEYS.forEach(m => {
    const part = flat[m];
    if (!part || typeof part !== "object") return;
    Object.keys(part).forEach(k => {
      const a = areaOf(m, k);
      const unchanged = same(m, k, part[k]);
      if ((onlyCustom && m !== "customScores") || (areas && a && areas.indexOf(a) === -1)) {
        if (!unchanged && onlyCustom && m !== "customScores") throw new Error("Akses ditolak: level Anda tidak dapat mengubah pengaturan ini.");
        if (!unchanged && areas && a && areas.indexOf(a) === -1) throw new Error("Akses ditolak: data Area " + a + " di luar wilayah Anda.");
        delete part[k];
        return;
      }
      if (PERIOD_MAP_KEYS.indexOf(m) !== -1 && a && frozen[a]) {
        if (!unchanged) throw new Error(frozenMessage(a, blth, frozen[a]));
        delete part[k];
      }
    });
  });

  const changedScalars = PERIOD_SCALAR_KEYS.filter(k => flat[k] !== undefined && period[k] !== flat[k]);
  PERIOD_SCALAR_KEYS.forEach(k => { if (flat[k] !== undefined && changedScalars.indexOf(k) === -1) delete flat[k]; });
  if (changedScalars.length) {
    if (role !== "UP") throw new Error("Akses ditolak: koefisien jabatan & opsi perhitungan hanya dapat diubah oleh UP 3 Jatim.");
    const fa = Object.keys(frozen);
    if (fa.length) {
      throw new Error("Koefisien periode " + blthLabel(blth) + " tidak dapat diubah karena Area " + fa.join(", ") +
        " sudah mengajukan / disetujui. Kembalikan pengajuan Area tersebut terlebih dahulu.");
    }
  }
  if (flat.defaultAdminName !== undefined && role !== "UP" && flat.defaultAdminName !== store.defaultAdminName) {
    delete flat.defaultAdminName;
  }
}

/* ---------- Daftar periksa sebelum mengajukan ---------- */
function areaLiveTotals(stats, area) {
  const u = (stats.rekapUP3 || []).find(x => x.up3 === area) || {};
  const recs = (stats.recipients || []).filter(r => r.up3 === area);
  return {
    gross: Math.round(u.rewardKotor || 0), denda: Math.round(u.dendaArea || 0), netto: Math.round(u.rewardNetto || 0),
    pagu: Math.round(u.paguArea || 0), slaArea: u.slaArea === undefined ? 100 : u.slaArea,
    reward: Math.round(u.totalReward || 0), n: recs.filter(r => r.bobot > 0).length, nPetugas: recs.filter(r => r.jabatan === "PETUGAS MANBILL").length
  };
}

function computeSubmissionCheck(store, all, blth, area, stats) {
  const errors = [], warnings = [];
  const label = blthLabel(blth);
  const idx = dataColumnIdx(all.headers);
  const rows = getPeriodRows(all, blth);
  const areaRows = rows.filter(c => idx.up3 !== -1 && String(c[idx.up3] || "").trim().toUpperCase() === area);
  const p = store.periods[blth] || {};
  const sub = getSubmission(store, blth, area);

  if (sub && FROZEN_STATUSES.indexOf(sub.status) !== -1) errors.push("Sudah diajukan (" + SUBMISSION_STATUS[sub.status] + ").");
  if (!areaRows.length) errors.push("Data kinerja Area " + area + " periode " + label + " belum di-upload.");
  const gross = p.areaRewardGross ? Number(p.areaRewardGross[area]) : NaN;
  if (!(gross > 0)) errors.push("Nilai Reward (Rp) Area belum diisi.");

  if (areaRows.length) {
    if (!p.areaSlaScores || p.areaSlaScores[area] === undefined) warnings.push("SLA Kontrak Area belum pernah diisi (dipakai 100%).");
    const ulps = Array.from(new Set(areaRows.map(c => idx.ulp === -1 ? "" : String(c[idx.ulp] || "").trim().toUpperCase()).filter(Boolean)));
    const noSla = ulps.filter(u => !p.ulpSlaScores || p.ulpSlaScores[area + "||" + u] === undefined);
    if (noSla.length) warnings.push(noSla.length + " dari " + ulps.length + " ULP belum diisi SLA Kontrak ULP (dipakai 100%): " + noSla.slice(0, 6).join(", ") + (noSla.length > 6 ? ", ..." : "") + ".");
    const iAbs = findHeaderIdx(all.headers, [ABSENSI_HEADER]);
    const noAbs = areaRows.filter(c => iAbs === -1 || parseAbsensiCell(c[iAbs]) === null).length;
    if (noAbs) warnings.push(noAbs + " dari " + areaRows.length + " petugas belum ada realisasi Absensi dari upload (dipakai nilai manual / standar 100%).");
    const cs = p.customScores || {};
    const noScore = areaRows.filter(c => idx.nip === -1 || cs[String(c[idx.nip] || "").trim()] === undefined).length;
    if (noScore) warnings.push(noScore + " petugas belum dinilai melalui Edit Nilai (dipakai nilai standar).");
  }
  const st = stats || calculateRekapAndReward(all.headers, rows, flattenConfigForPeriod(store, blth));
  const totals = areaLiveTotals(st, area);
  if (areaRows.length && Number(totals.slaArea) < 100) warnings.push("SLA Kontrak Area " + totals.slaArea + "% (< 100%) sehingga pagu reward Area = Rp 0.");
  if (areaRows.length && totals.reward <= 0 && gross > 0) warnings.push("Total reward yang terbagi Rp 0.");
  return { blth: blth, label: label, area: area, ok: errors.length === 0, errors: errors, warnings: warnings, totals: totals };
}

/* ---------- Penyimpanan status, riwayat, log & snapshot ---------- */
function getSubmission(store, blth, area) {
  return (store.submissions || {})[submissionKey(blth, area)] || null;
}

function pushSubmissionHistory(sub, entry) {
  sub.history = (Array.isArray(sub.history) ? sub.history : []).concat([entry]).slice(-40);
}

function appendSubmissionLog(ss, entries) {
  if (!ss || !entries.length) return;
  const H = ["Waktu", "BLTH", "Periode", "Area / UP3", "Aksi", "Status Baru", "Oleh", "Catatan"];
  const sh = ss.getSheetByName(SHEET_LOG_NAME) || ss.insertSheet(SHEET_LOG_NAME);
  let last = sh.getLastRow();
  if (last === 0) {
    ensureSheetDimensions(sh, 2, H.length);
    sh.getRange(1, 1, 1, H.length).setValues([H]).setBackground("#003b5c").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
    last = 1;
  }
  const rows = entries.map(e => [fmtDateTime(e.at), e.blth, blthLabel(e.blth), e.area, e.actionLabel, SUBMISSION_STATUS[e.status] || e.status, e.byLabel, e.note || ""]);
  ensureSheetDimensions(sh, last + rows.length + 1, H.length);
  sh.getRange(last + 1, 2, rows.length, 1).setNumberFormat("@");
  sh.getRange(last + 1, 1, rows.length, H.length).setValues(rows);
}

function readApprovedRows(ss) {
  const sh = ss ? ss.getSheetByName(SHEET_APPROVED_NAME) : null;
  if (!sh || sh.getLastRow() < 2) return [];
  return sanitizeValuesForClient(sh.getRange(2, 1, sh.getLastRow() - 1, APPROVED_HEADERS.length).getValues())
    .filter(r => r.some(c => String(c).trim() !== ""))
    .map(r => { r[0] = cellToBlth(r[0]); r[2] = cellToBlth(r[2]); return r; });
}

function writeApprovedRows(ss, removeKeys, addRows) {
  const sh = ss.getSheetByName(SHEET_APPROVED_NAME) || ss.insertSheet(SHEET_APPROVED_NAME);
  const remove = {};
  (removeKeys || []).forEach(k => { remove[k] = true; });
  const rows = readApprovedRows(ss).filter(r => !remove[String(r[APPROVED_KEY_COL])]).concat(addRows || []);
  rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])) || String(a[2]).localeCompare(String(b[2])) ||
    String(a[4]).localeCompare(String(b[4])) || (JABATAN_ORDER[a[9]] || 9) - (JABATAN_ORDER[b[9]] || 9) || (Number(b[14]) || 0) - (Number(a[14]) || 0));
  const values = [APPROVED_HEADERS].concat(rows);
  resetSheetSafely(sh, values.length + 5, APPROVED_HEADERS.length + 1);
  sh.getRange(1, 1, values.length, 1).setNumberFormat("@");
  sh.getRange(1, 3, values.length, 1).setNumberFormat("@");
  sh.getRange(1, 7, values.length, 1).setNumberFormat("@");
  sh.getRange(1, 1, values.length, APPROVED_HEADERS.length).setValues(values);
  sh.getRange(1, 1, 1, APPROVED_HEADERS.length).setBackground("#003b5c").setFontColor("#ffffff").setFontWeight("bold");
  sh.setFrozenRows(1);
  if (rows.length) sh.getRange(2, 14, rows.length, 2).setNumberFormat('"Rp"#,##0');
}

function buildSnapshotRows(stats, blth, area, batch, upAtIso) {
  const key = submissionKey(blth, area);
  return (stats.recipients || [])
    .filter(r => r.up3 === area)
    .sort((a, b) => (JABATAN_ORDER[a.jabatan] || 9) - (JABATAN_ORDER[b.jabatan] || 9) || b.reward - a.reward)
    .map(r => [batch, blthLabel(batch), blth, blthLabel(blth), r.up3, r.ulp, r.nipPegawai || "", r.nip, r.nama, r.jabatan,
      r.jabatan === "PETUGAS MANBILL" ? r.koordinator : "-", r.score, r.bobot, r.nilaiPerBobot, Math.round(r.reward), fmtDateTime(upAtIso), key]);
}

// Kunci penuh otomatis bila seluruh Area yang punya data di periode tsb sudah disetujui UP
function autoLockCheck(store, all, blth) {
  const idx = dataColumnIdx(all.headers);
  const areas = Array.from(new Set(getPeriodRows(all, blth).map(c => idx.up3 === -1 ? "" : String(c[idx.up3] || "").trim().toUpperCase()).filter(Boolean)));
  const allFinal = areas.length > 0 && areas.every(a => { const s = getSubmission(store, blth, a); return s && FINAL_STATUSES.indexOf(s.status) !== -1; });
  const cur = store.lockedPeriods[blth];
  if (allFinal && !cur) store.lockedPeriods[blth] = { at: new Date().toISOString(), by: "OTOMATIS", auto: true };
  if (!allFinal && cur && cur.auto) delete store.lockedPeriods[blth];
}

/* ---------- Endpoint ---------- */

function getSubmissionPrecheck(token, blths) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const all = readAllDataRows(ss);
    const sess = requireSession(token, all);
    if (sess.role !== "AREA") throw new Error("Pengajuan reward hanya dapat dilakukan oleh user Area.");
    const store = getConfigStore();
    const list = Array.from(new Set((blths || []).map(String))).filter(isValidBlth).sort();
    return { success: true, checks: list.map(b => computeSubmissionCheck(store, all, b, sess.up3)) };
  } catch (err) {
    return { success: false, message: errText(err) };
  }
}

function submitRewardRequest(token, blths, note) {
  try {
    const cleanNote = String(note || "").trim().substring(0, 500);
    const list = Array.from(new Set((blths || []).map(String))).filter(isValidBlth).sort();
    if (!list.length) throw new Error("Pilih minimal satu bulan yang akan diajukan.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return withScriptLock(() => {
      const all = readAllDataRows(ss);
      const sess = requireSession(token, all);
      if (sess.role !== "AREA") throw new Error("Pengajuan reward hanya dapat dilakukan oleh user Area.");
      const store = getConfigStore();
      const area = sess.up3, now = new Date().toISOString();
      const results = [], log = [];
      list.forEach(b => {
        const chk = computeSubmissionCheck(store, all, b, area);
        if (!chk.ok) { results.push({ key: submissionKey(b, area), blth: b, ok: false, message: blthLabel(b) + ": " + chk.errors.join(" ") }); return; }
        const key = submissionKey(b, area);
        const prev = store.submissions[key] || { blth: b, area: area, history: [] };
        const next = areaHasMul(area) ? "MENUNGGU_MUL" : "MENUNGGU_UP";
        const sub = {
          blth: b, area: area, status: next, round: (prev.round || 0) + 1,
          submittedAt: now, submittedBy: sess.label, submitNote: cleanNote,
          mulAt: "", mulBy: "", mulNote: "", upAt: "", upBy: "", upNote: "", sdmAt: "", sdmBy: "", sdmNote: "",
          returnedAt: "", returnedBy: "", returnNote: "",
          lastReturn: prev.returnedAt ? { at: prev.returnedAt, by: prev.returnedBy, note: prev.returnNote } : (prev.lastReturn || null),
          batch: "", totals: null, history: prev.history || []
        };
        const entry = { at: now, blth: b, area: area, action: "AJUKAN", actionLabel: sub.round > 1 ? "Diajukan ulang" : "Diajukan", status: next, by: sess.key, byLabel: sess.label, note: cleanNote };
        pushSubmissionHistory(sub, entry);
        store.submissions[key] = sub;
        log.push(entry);
        results.push({ key: key, blth: b, ok: true, message: blthLabel(b) + ": " + SUBMISSION_STATUS[next] });
      });
      if (log.length) {
        saveConfigStore(store);
        appendSubmissionLog(ss, log);
      }
      const okN = log.length;
      return {
        success: okN > 0, results: results, okCount: okN, failCount: results.length - okN,
        message: okN ? (okN + " bulan berhasil diajukan" + (results.length > okN ? ", " + (results.length - okN) + " bulan gagal." : ".")) : "Pengajuan gagal: " + results.map(r => r.message).join(" | ")
      };
    });
  } catch (err) {
    return { success: false, message: errText(err) };
  }
}

/**
 * Aksi atas pengajuan: action = "SETUJU" | "KEMBALIKAN" | "PROSES" | "BATAL_PROSES".
 * - SETUJU     : MUL (MENUNGGU_MUL → MENUNGGU_UP), UP 3 (MENUNGGU_UP → DISETUJUI + snapshot + Bulan Proses)
 * - KEMBALIKAN : MUL (MENUNGGU_MUL), UP 3 (MENUNGGU_MUL / MENUNGGU_UP / DISETUJUI = batalkan persetujuan),
 *                Area (tarik pengajuan yang belum disetujui siapa pun). Catatan wajib.
 * - PROSES / BATAL_PROSES : SDM (DISETUJUI ↔ DIPROSES)
 */
function processRewardRequests(token, keys, action, note) {
  try {
    const cleanNote = String(note || "").trim().substring(0, 500);
    const list = Array.from(new Set((keys || []).map(String))).filter(Boolean);
    if (!list.length) throw new Error("Pilih minimal satu pengajuan.");
    if (["SETUJU", "KEMBALIKAN", "PROSES", "BATAL_PROSES"].indexOf(action) === -1) throw new Error("Aksi tidak dikenal.");
    if (action === "KEMBALIKAN" && !cleanNote) throw new Error("Catatan / alasan pengembalian wajib diisi.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return withScriptLock(() => {
      const all = readAllDataRows(ss);
      const sess = requireSession(token, all);
      const store = getConfigStore();
      const areas = sessionAreas(sess);
      const now = new Date().toISOString();
      const statsCache = {};
      const getStats = b => statsCache[b] || (statsCache[b] = calculateRekapAndReward(all.headers, getPeriodRows(all, b), flattenConfigForPeriod(store, b)));
      const results = [], log = [], snapRemove = [], snapAdd = [], touched = {};

      list.forEach(key => {
        const sub = store.submissions[key];
        const kk = splitSubmissionKey(key);
        const name = kk.area + " " + blthLabel(kk.blth);
        const fail = m => results.push({ key: key, ok: false, message: name + ": " + m });
        if (!sub) return fail("belum diajukan.");
        if (areas && areas.indexOf(sub.area) === -1) return fail("di luar wilayah Anda.");
        const role = sess.role, st = sub.status;
        let next = "", actLabel = "";

        if (action === "SETUJU") {
          if (role === "MUL" && st === "MENUNGGU_MUL") {
            next = "MENUNGGU_UP"; actLabel = "Disetujui MUL";
            sub.mulAt = now; sub.mulBy = sess.label; sub.mulNote = cleanNote;
          } else if (role === "UP" && st === "MENUNGGU_UP") {
            next = "DISETUJUI"; actLabel = "Disetujui UP 3";
            const stats = getStats(sub.blth);
            sub.upAt = now; sub.upBy = sess.label; sub.upNote = cleanNote;
            sub.batch = monthKeyNow();
            sub.totals = areaLiveTotals(stats, sub.area);
            snapRemove.push(key);
            Array.prototype.push.apply(snapAdd, buildSnapshotRows(stats, sub.blth, sub.area, sub.batch, now));
          } else {
            return fail(role === "MUL" ? "hanya pengajuan berstatus Menunggu Approval MUL yang dapat disetujui MUL."
              : role === "UP" ? "hanya pengajuan berstatus Menunggu Approval UP 3 yang dapat disetujui UP 3."
              : "level Anda tidak dapat menyetujui pengajuan.");
          }
        } else if (action === "KEMBALIKAN") {
          const areaWithdraw = role === "AREA" && (st === "MENUNGGU_MUL" || (st === "MENUNGGU_UP" && !areaHasMul(sub.area)));
          const ok = (role === "MUL" && st === "MENUNGGU_MUL") || (role === "UP" && ["MENUNGGU_MUL", "MENUNGGU_UP", "DISETUJUI"].indexOf(st) !== -1) || areaWithdraw;
          if (!ok) {
            return fail(st === "DIPROSES" ? "sudah diproses SDM, tidak dapat dikembalikan."
              : role === "AREA" ? "pengajuan yang sudah disetujui tidak dapat ditarik. Hubungi MUL / UP 3."
              : "status " + (SUBMISSION_STATUS[st] || st) + " tidak dapat dikembalikan oleh level Anda.");
          }
          next = "DIKEMBALIKAN";
          actLabel = areaWithdraw ? "Ditarik oleh Area" : (st === "DISETUJUI" ? "Persetujuan dibatalkan UP 3" : "Dikembalikan");
          sub.returnedAt = now; sub.returnedBy = sess.label; sub.returnNote = cleanNote; sub.returnedFrom = st;
          if (st === "DISETUJUI") { snapRemove.push(key); sub.batch = ""; sub.totals = null; }
        } else if (action === "PROSES") {
          if (role !== "SDM") return fail("hanya user SDM yang dapat menandai sudah diproses.");
          if (st !== "DISETUJUI") return fail("hanya pengajuan yang sudah disetujui UP 3 yang dapat diproses.");
          next = "DIPROSES"; actLabel = "Diproses SDM";
          sub.sdmAt = now; sub.sdmBy = sess.label; sub.sdmNote = cleanNote;
        } else if (action === "BATAL_PROSES") {
          if (role !== "SDM") return fail("hanya user SDM yang dapat membatalkan tanda proses.");
          if (st !== "DIPROSES") return fail("belum ditandai diproses.");
          next = "DISETUJUI"; actLabel = "Tanda proses SDM dibatalkan";
          sub.sdmAt = ""; sub.sdmBy = ""; sub.sdmNote = "";
        }

        sub.status = next;
        const entry = { at: now, blth: sub.blth, area: sub.area, action: action, actionLabel: actLabel, status: next, by: sess.key, byLabel: sess.label, note: cleanNote };
        pushSubmissionHistory(sub, entry);
        log.push(entry);
        touched[sub.blth] = true;
        results.push({ key: key, ok: true, message: name + ": " + SUBMISSION_STATUS[next] });
      });

      if (log.length) {
        Object.keys(touched).forEach(b => autoLockCheck(store, all, b));
        saveConfigStore(store);
        if (snapRemove.length || snapAdd.length) writeApprovedRows(ss, snapRemove, snapAdd);
        appendSubmissionLog(ss, log);
      }
      const okN = log.length;
      return {
        success: okN > 0, results: results, okCount: okN, failCount: results.length - okN, lockedPeriods: store.lockedPeriods,
        message: okN ? (okN + " pengajuan berhasil diperbarui" + (results.length > okN ? ", " + (results.length - okN) + " gagal." : "."))
          : results.map(r => r.message).join(" | ")
      };
    });
  } catch (err) {
    return { success: false, message: errText(err) };
  }
}

function buildMonitorItem(store, blth, area, sub, live, hasData) {
  const s = sub || {};
  const status = s.status || "BELUM";
  const final = FINAL_STATUSES.indexOf(status) !== -1 && s.totals;
  return {
    key: submissionKey(blth, area), blth: blth, label: blthLabel(blth), area: area, hasMul: areaHasMul(area), hasData: hasData,
    status: status, round: s.round || 0,
    submittedAt: s.submittedAt || "", submittedBy: s.submittedBy || "", submitNote: s.submitNote || "",
    mulAt: s.mulAt || "", mulBy: s.mulBy || "", mulNote: s.mulNote || "",
    upAt: s.upAt || "", upBy: s.upBy || "", upNote: s.upNote || "",
    sdmAt: s.sdmAt || "", sdmBy: s.sdmBy || "", sdmNote: s.sdmNote || "",
    returnedAt: s.returnedAt || "", returnedBy: s.returnedBy || "", returnNote: s.returnNote || "", returnedFrom: s.returnedFrom || "",
    lastReturn: s.lastReturn || null,
    batch: s.batch || "", batchLabel: s.batch ? blthLabel(s.batch) : "",
    totals: final ? s.totals : live, snapshot: Boolean(final),
    locked: Boolean(store.lockedPeriods[blth]), history: s.history || []
  };
}

/** Data dashboard monitoring pengajuan: seluruh Area × periode di wilayah user. */
function getSubmissionMonitor(token) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const all = ss ? readAllDataRows(ss) : { headers: [], rows: [] };
    const sess = requireSession(token, all);
    const store = getConfigStore();
    const areas = sessionAreas(sess);
    const items = [], seen = {};
    listDataPeriods(all).forEach(b => {
      const stats = calculateRekapAndReward(all.headers, getPeriodRows(all, b), flattenConfigForPeriod(store, b));
      (stats.rekapUP3 || []).forEach(u => {
        if (areas && areas.indexOf(u.up3) === -1) return;
        const key = submissionKey(b, u.up3);
        seen[key] = true;
        items.push(buildMonitorItem(store, b, u.up3, store.submissions[key], areaLiveTotals(stats, u.up3), true));
      });
    });
    Object.keys(store.submissions).forEach(key => {
      const s = store.submissions[key];
      if (seen[key] || !s || (areas && areas.indexOf(s.area) === -1)) return;
      items.push(buildMonitorItem(store, s.blth, s.area, s, s.totals || { gross: 0, denda: 0, netto: 0, pagu: 0, reward: 0, n: 0 }, false));
    });
    items.sort((a, b) => b.blth.localeCompare(a.blth) || a.area.localeCompare(b.area));
    const batches = Array.from(new Set(items.filter(i => i.batch).map(i => i.batch))).sort().reverse();
    return {
      success: true, items: items, batches: batches, nowMonth: monthKeyNow(), statusLabels: SUBMISSION_STATUS,
      user: { role: sess.role, label: sess.label, up3: sess.up3 || "", isAdmin: Boolean(sess.isAdmin) }
    };
  } catch (err) {
    return { success: false, message: errText(err) };
  }
}

/**
 * Data Excel gabungan untuk SDM / UP 3: reward yang sudah disetujui UP 3 (snapshot).
 * opts = { batch: "YYYYMM" (Bulan Proses) | "", keys: [...] (opsional), onlyUnprocessed: bool }
 */
function getSdmExportData(token, opts) {
  try {
    const o = opts || {};
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const all = readAllDataRows(ss);
    const sess = requireSession(token, all);
    if (sess.role !== "SDM" && sess.role !== "UP") throw new Error("Unduhan gabungan hanya untuk user SDM dan UP 3 Jatim.");
    const store = getConfigStore();
    const keySet = {};
    (o.keys || []).forEach(k => { keySet[String(k)] = true; });
    const useKeys = Object.keys(keySet).length > 0;
    const items = Object.keys(store.submissions).map(k => store.submissions[k]).filter(s => s && FINAL_STATUSES.indexOf(s.status) !== -1)
      .filter(s => useKeys ? keySet[submissionKey(s.blth, s.area)] : (!o.batch || s.batch === String(o.batch)))
      .filter(s => !o.onlyUnprocessed || s.status === "DISETUJUI")
      .map(s => buildMonitorItem(store, s.blth, s.area, s, s.totals, true))
      .sort((a, b) => a.blth.localeCompare(b.blth) || a.area.localeCompare(b.area));
    const wanted = {};
    items.forEach(i => { wanted[i.key] = true; });
    const rows = readApprovedRows(ss).filter(r => wanted[String(r[APPROVED_KEY_COL])]);
    return { success: true, items: items, headers: APPROVED_HEADERS.slice(0, APPROVED_KEY_COL), rows: rows.map(r => r.slice(0, APPROVED_KEY_COL)) };
  } catch (err) {
    return { success: false, message: errText(err) };
  }
}

/* ============================== PERHITUNGAN NILAI ============================== */

// Kinerja MO & Kinerja SLA Kontrak memakai aturan AMBANG: pencapaian >= 100% → bobot penuh; < 100% → 0.
// (Komponen 1 Admin Manbill = Pembentukkan RPP, tetap proporsional.)
function calcThresholdScore(rawInputPct, weightPct) {
  if (weightPct <= 0) return 0;
  return Number(Number(rawInputPct || 0).toFixed(2)) >= 100 ? weightPct : 0;
}

function calcComponentWeightedScore(rawInputPct, weightPct) {
  if (weightPct <= 0) return 0;
  const cleanVal = Math.max(Number(rawInputPct) || 0, 0);
  const effectivePct = cleanVal >= 100 ? 100 : cleanVal;
  return Number(((effectivePct / 100) * weightPct).toFixed(2));
}

function getAtasanCategoryLabel(rawVal) {
  const val = Number(rawVal) || 0;
  if (val >= 80) return "Baik (80-100)";
  if (val >= 60) return "Sedang (60-80)";
  return "Kurang (<60)";
}

function buildScoreBreakdown(jabatan, excelTotalScoreRaw, customOverride, areaSlaVal, uploadedAbsensi) {
  const isSpv = (jabatan === "SPV MANBILL AREA" || jabatan === "SPV MANBILL ULP");
  const maxU1 = isSpv ? 10 : 40;
  const maxU2 = isSpv ? 30 : 15;
  const maxU3 = isSpv ? 20 : 5;
  const maxTambahan = 10;
  const maxAbsensi = 10;
  const maxAtitude = 5, maxResponsif = 5, maxInisiatif = 5, maxKepedulian = 5;

  const defaultRawU1 = (excelTotalScoreRaw !== undefined && excelTotalScoreRaw !== null && !isNaN(Number(excelTotalScoreRaw)))
    ? Number(Number(excelTotalScoreRaw).toFixed(2))
    : 100;
  const hasAreaSla = areaSlaVal !== undefined && areaSlaVal !== null && !isNaN(Number(areaSlaVal));

  let rawU1 = defaultRawU1;
  let rawU2 = hasAreaSla ? Number(areaSlaVal) : 100;
  let rawU3 = 100;
  let rawTambahan = 100;
  let rawAbsensi = 100;
  let rawAtitude = 100;
  let rawResponsif = 100;
  let rawInisiatif = 100;
  let rawKepedulian = 100;

  if (customOverride && typeof customOverride === "object") {
    if (customOverride.isRawPct) {
      rawU1 = Number(customOverride.rawU1 ?? defaultRawU1);
      rawU2 = hasAreaSla ? Number(areaSlaVal) : Number(customOverride.rawU2 ?? 100);
      rawU3 = Number(customOverride.rawU3 ?? 100);
      rawTambahan = Number(customOverride.rawTambahan ?? 100);
      rawAbsensi = Number(customOverride.rawAbsensi ?? 100);
      rawAtitude = Number(customOverride.rawAtitude ?? 100);
      rawResponsif = Number(customOverride.rawResponsif ?? 100);
      rawInisiatif = Number(customOverride.rawInisiatif ?? 100);
      rawKepedulian = Number(customOverride.rawKepedulian ?? 100);
    } else {
      rawU1 = maxU1 > 0 ? Number(((Number(customOverride.u1 ?? maxU1) / maxU1) * 100).toFixed(2)) : 100;
      rawU2 = hasAreaSla ? Number(areaSlaVal) : (maxU2 > 0 ? Number(((Number(customOverride.u2 ?? maxU2) / maxU2) * 100).toFixed(2)) : 100);
      rawU3 = maxU3 > 0 ? Number(((Number(customOverride.u3 ?? maxU3) / maxU3) * 100).toFixed(2)) : 100;
      rawTambahan = maxTambahan > 0 ? Number(((Number(customOverride.tambahan ?? maxTambahan) / maxTambahan) * 100).toFixed(2)) : 100;
      rawAbsensi = maxAbsensi > 0 ? Number(((Number(customOverride.absensi ?? maxAbsensi) / maxAbsensi) * 100).toFixed(2)) : 100;
      rawAtitude = Number(((Number(customOverride.atitude ?? 5) / 5) * 100).toFixed(2));
      rawResponsif = Number(((Number(customOverride.responsif ?? 5) / 5) * 100).toFixed(2));
      rawInisiatif = Number(((Number(customOverride.inisiatif ?? 5) / 5) * 100).toFixed(2));
      rawKepedulian = Number(((Number(customOverride.kepedulian ?? 5) / 5) * 100).toFixed(2));
    }
  }

  // Realisasi absensi dari upload file Absensi selalu dipakai (menimpa input manual)
  const absensiUpload = uploadedAbsensi !== null && uploadedAbsensi !== undefined && !isNaN(Number(uploadedAbsensi));
  if (absensiUpload) rawAbsensi = Number(uploadedAbsensi);

  [rawU1, rawU2, rawU3, rawTambahan, rawAbsensi, rawAtitude, rawResponsif, rawInisiatif, rawKepedulian] =
    [rawU1, rawU2, rawU3, rawTambahan, rawAbsensi, rawAtitude, rawResponsif, rawInisiatif, rawKepedulian].map(v => isNaN(v) ? 0 : v);

  const u1 = jabatan === "ADMIN MANBILL" ? calcComponentWeightedScore(rawU1, maxU1) : calcThresholdScore(rawU1, maxU1);
  const u2 = calcThresholdScore(rawU2, maxU2);
  const u3 = calcComponentWeightedScore(rawU3, maxU3);
  const subUtama = Number((u1 + u2 + u3).toFixed(2));

  const tambahan = calcComponentWeightedScore(rawTambahan, maxTambahan);
  const absensi = calcComponentWeightedScore(rawAbsensi, maxAbsensi);

  const atitude = calcComponentWeightedScore(rawAtitude, maxAtitude);
  const responsif = calcComponentWeightedScore(rawResponsif, maxResponsif);
  const inisiatif = calcComponentWeightedScore(rawInisiatif, maxInisiatif);
  const kepedulian = calcComponentWeightedScore(rawKepedulian, maxKepedulian);
  const subAtasan = Number((atitude + responsif + inisiatif + kepedulian).toFixed(2));

  const totalNilai = Number((subUtama + tambahan + absensi + subAtasan).toFixed(2));

  return {
    excelScore: defaultRawU1,
    rawU1: Number(rawU1.toFixed(2)),
    rawU2: Number(rawU2.toFixed(2)),
    rawU3: Number(rawU3.toFixed(2)),
    rawTambahan: Number(rawTambahan.toFixed(2)),
    rawAbsensi: Number(rawAbsensi.toFixed(2)),
    rawAtitude: Number(rawAtitude.toFixed(2)),
    rawResponsif: Number(rawResponsif.toFixed(2)),
    rawInisiatif: Number(rawInisiatif.toFixed(2)),
    rawKepedulian: Number(rawKepedulian.toFixed(2)),
    absensiUpload: absensiUpload,
    catAtitude: getAtasanCategoryLabel(rawAtitude),
    catResponsif: getAtasanCategoryLabel(rawResponsif),
    catInisiatif: getAtasanCategoryLabel(rawInisiatif),
    catKepedulian: getAtasanCategoryLabel(rawKepedulian),
    u1, u2, u3, subUtama,
    tambahan, absensi,
    atitude, responsif, inisiatif, kepedulian, subAtasan,
    totalNilai
  };
}

// Pengali reward per ULP: SLA Kontrak ULP (maks 100%) dikalikan ke bobot Petugas, Kordinator & SPV ULP
// di ULP tsb. SPV Area & Admin tidak terpengaruh. Pagu Area tetap terbagi habis antar-ULP.
function ulpSlaToFactor(sla) {
  const v = Number(sla);
  if (isNaN(v)) return 1;
  return Math.min(Math.max(v, 0), 100) / 100;
}

// Nilai Reward (kotor) - Nilai Denda = Netto; Pagu Area = Netto x 70%.
// Config lama (sebelum ada Denda): nilai pagu lama di areaRewards dipakai sebagai Nilai Reward kotor.
function calcAreaPaguParts(cfg, up3, fallbackGross) {
  const rawGross = cfg.areaRewardGross[up3] !== undefined ? Number(cfg.areaRewardGross[up3])
    : (cfg.areaRewards[up3] !== undefined ? Number(cfg.areaRewards[up3]) : fallbackGross);
  const rawDenda = cfg.areaDenda[up3] !== undefined ? Number(cfg.areaDenda[up3]) : 0;
  const gross = Math.max(isNaN(rawGross) ? 0 : rawGross, 0);
  const denda = Math.max(isNaN(rawDenda) ? 0 : rawDenda, 0);
  const netto = Math.max(gross - denda, 0);
  return { gross, denda, netto, pagu: Math.round(netto * NETTO_PAGU_PCT) };
}

function calculateRekapAndReward(headers, rows, cfg) {
  const nipIdx = findHeaderIdx(headers, ["NIP"]);
  const petugasIdx = findHeaderIdx(headers, ["PETUGAS"]);
  const koordIdx = findHeaderIdx(headers, ["KOORDINATOR", "KORDINATOR"]);
  const ulpIdx = findHeaderIdx(headers, ["ULP"]);
  const up3Idx = findHeaderIdx(headers, ["UP3"]);
  const statusIdx = findHeaderIdx(headers, ["STATUS"]);
  const totalIdx = findHeaderIdx(headers, ["TOTAL"]);
  const absIdx = findHeaderIdx(headers, [ABSENSI_HEADER]);

  CONFIG_MAP_KEYS.forEach(k => { if (!cfg[k] || typeof cfg[k] !== "object") cfg[k] = {}; });
  const fallbackAreaPagu = Number(cfg.defaultAreaReward) >= 0 ? Number(cfg.defaultAreaReward) : DEFAULT_AREA_PAGU;
  const jAdmin = Number(cfg.koefAdmin) >= 0 ? Number(cfg.koefAdmin) : 0.3;
  const jPetugas = Number(cfg.koefPetugas) >= 0 ? Number(cfg.koefPetugas) : 0.4;
  const jKoord = Number(cfg.koefKoordinator) >= 0 ? Number(cfg.koefKoordinator) : 0.6;
  const jTLULP = Number(cfg.koefTLULP) >= 0 ? Number(cfg.koefTLULP) : 0.8;
  const jTLArea = Number(cfg.koefTLArea) >= 0 ? Number(cfg.koefTLArea) : 1;
  const useScoreInWeight = cfg.useScoreInWeight !== false;
  const onlyApproved = cfg.onlyApproved === true;

  const validRows = (rows || []).filter(r => Array.isArray(r) && r.some(c => c !== null && c !== undefined && String(c).trim() !== ""));
  let totalData = validRows.length, approved = 0, notApproved = 0;
  const up3Map = {}, ulpMap = {}, koordTeams = {}, ulpTeams = {}, up3Teams = {}, recipients = [];

  validRows.forEach((r, idx) => {
    const nip = (nipIdx !== -1 && r[nipIdx] !== "" && r[nipIdx] != null) ? String(r[nipIdx]).trim() : "PTG-" + (idx + 1);
    const nama = (petugasIdx !== -1 && r[petugasIdx]) ? String(r[petugasIdx]).trim() : "PETUGAS " + (idx + 1);
    const koord = (koordIdx !== -1 && r[koordIdx]) ? String(r[koordIdx]).trim() : "TANPA KOORDINATOR";
    const ulp = (ulpIdx !== -1 && r[ulpIdx]) ? String(r[ulpIdx]).trim().toUpperCase() : "TIDAK TERDEFINISI";
    const up3 = (up3Idx !== -1 && r[up3Idx]) ? String(r[up3Idx]).trim().toUpperCase() : "TIDAK TERDEFINISI";
    const isApp = String(statusIdx !== -1 ? r[statusIdx] : "").trim().toLowerCase() === "approved";
    const rawScore = parseScoreValue(totalIdx !== -1 ? r[totalIdx] : null);

    if (isApp) approved++; else notApproved++;

    if (!up3Map[up3]) {
      const parts = calcAreaPaguParts(cfg, up3, fallbackAreaPagu);
      const cleanPagu = parts.pagu;
      const slaVal = cfg.areaSlaScores[up3] !== undefined ? Number(cfg.areaSlaScores[up3]) : 100;
      const cleanSla = (!isNaN(slaVal) && slaVal >= 0) ? slaVal : 100;
      cfg.areaRewardGross[up3] = parts.gross;
      cfg.areaDenda[up3] = parts.denda;
      cfg.areaRewards[up3] = cleanPagu;
      cfg.areaSlaScores[up3] = cleanSla;
      up3Map[up3] = {
        up3, approved: 0, notApproved: 0, total: 0,
        rewardKotor: parts.gross, dendaArea: parts.denda, rewardNetto: parts.netto,
        paguArea: cleanPagu, slaArea: cleanSla, totalBobot: 0, nilaiPerBobot: 0,
        rewardAdmin: 0, rewardPetugas: 0, rewardKoordinator: 0, rewardTLULP: 0, rewardTLArea: 0,
        totalReward: 0
      };
    }
    up3Map[up3].total++;
    if (isApp) up3Map[up3].approved++; else up3Map[up3].notApproved++;

    const ulpKey = up3 + "||" + ulp;
    if (!ulpMap[ulpKey]) {
      const ulpSlaRaw = cfg.ulpSlaScores[ulpKey] !== undefined ? Number(cfg.ulpSlaScores[ulpKey]) : 100;
      const cleanUlpSla = (!isNaN(ulpSlaRaw) && ulpSlaRaw >= 0) ? ulpSlaRaw : 100;
      cfg.ulpSlaScores[ulpKey] = cleanUlpSla;
      ulpMap[ulpKey] = {
        ulp, up3, slaUlp: cleanUlpSla, approved: 0, notApproved: 0, total: 0,
        rewardAdmin: 0, rewardPetugas: 0, rewardKoordinator: 0, rewardTLULP: 0,
        totalReward: 0
      };
    }
    ulpMap[ulpKey].total++;
    if (isApp) ulpMap[ulpKey].approved++; else ulpMap[ulpKey].notApproved++;

    const jabatan = "PETUGAS MANBILL";
    const ulpSlaVal = cfg.ulpSlaScores[ulpKey];
    const breakdown = buildScoreBreakdown(jabatan, rawScore, cfg.customScores[nip], ulpSlaVal, absIdx !== -1 ? parseAbsensiCell(r[absIdx]) : null);
    const finalScore = breakdown.totalNilai;
    const scoreFactor = useScoreInWeight ? (finalScore / 100) : 1.0;
    const slaFactor = ulpSlaToFactor(ulpSlaVal);
    const bobot = (!onlyApproved || isApp) ? (jPetugas * scoreFactor * slaFactor) : 0;

    recipients.push({
      no: recipients.length + 1, nip, nipPegawai: /^PTG-\d+$/.test(nip) ? "" : nip, nama, koordinator: koord, jabatan, ulp, up3,
      excelScore: Number(rawScore.toFixed(2)),
      breakdown,
      score: finalScore, status: isApp ? "approved" : "not approved",
      koefJabatan: jPetugas, bobot: Number(bobot.toFixed(4)), nilaiPerBobot: 0, reward: 0
    });

    const kKey = up3 + "||" + ulp + "||" + koord.toUpperCase();
    if (!koordTeams[kKey]) koordTeams[kKey] = { koord, ulp, up3, excelScores: [], anyApproved: false };
    koordTeams[kKey].excelScores.push(rawScore);
    if (isApp) koordTeams[kKey].anyApproved = true;

    if (!ulpTeams[ulpKey]) ulpTeams[ulpKey] = { ulp, up3, excelScores: [], anyApproved: false };
    ulpTeams[ulpKey].excelScores.push(rawScore);
    if (isApp) ulpTeams[ulpKey].anyApproved = true;

    if (!up3Teams[up3]) up3Teams[up3] = { up3, excelScores: [], anyApproved: false };
    up3Teams[up3].excelScores.push(rawScore);
    if (isApp) up3Teams[up3].anyApproved = true;
  });

  const avg = arr => arr.length > 0 ? arr.reduce((a, b) => a + b, 0) / arr.length : 100;

  Object.values(koordTeams).forEach((kt, idx) => {
    const id = makeKoordId(kt.up3, kt.ulp, kt.koord);
    const jabatan = "KORDINATOR";
    const avgExcelScore = avg(kt.excelScores);
    const ulpSlaVal = cfg.ulpSlaScores[kt.up3 + "||" + kt.ulp];
    const breakdown = buildScoreBreakdown(jabatan, avgExcelScore, takeCustomScore(cfg, id, "KOORD-" + (idx + 1)), ulpSlaVal);
    const finalScore = breakdown.totalNilai;
    const scoreFactor = useScoreInWeight ? (finalScore / 100) : 1.0;
    const slaFactor = ulpSlaToFactor(ulpSlaVal);
    const bobot = (!onlyApproved || kt.anyApproved) ? (jKoord * scoreFactor * slaFactor) : 0;
    recipients.push({
      no: recipients.length + 1, nip: id, nipPegawai: String(cfg.unitNips[id] || ""), nama: kt.koord, koordinator: kt.koord, jabatan, ulp: kt.ulp, up3: kt.up3,
      excelScore: Number(avgExcelScore.toFixed(2)),
      breakdown,
      score: finalScore, status: kt.anyApproved ? "approved" : "not approved",
      koefJabatan: jKoord, bobot: Number(bobot.toFixed(4)), nilaiPerBobot: 0, reward: 0
    });
  });

  Object.values(ulpTeams).forEach((ut, idx) => {
    const id = makeSpvUlpId(ut.up3, ut.ulp);
    const jabatan = "SPV MANBILL ULP";
    const avgExcelScore = avg(ut.excelScores);
    const ulpKey = ut.up3 + "||" + ut.ulp;
    const ulpSlaVal = cfg.ulpSlaScores[ulpKey];
    const customSpvUlpName = (cfg.spvUlpNames[ulpKey] && String(cfg.spvUlpNames[ulpKey]).trim())
      ? String(cfg.spvUlpNames[ulpKey]).trim()
      : ("SPV MANBILL ULP " + ut.ulp);
    const breakdown = buildScoreBreakdown(jabatan, avgExcelScore, takeCustomScore(cfg, id, "SPV-ULP-" + (idx + 1)), ulpSlaVal);
    const finalScore = breakdown.totalNilai;
    const scoreFactor = useScoreInWeight ? (finalScore / 100) : 1.0;
    const slaFactor = ulpSlaToFactor(ulpSlaVal);
    const bobot = (!onlyApproved || ut.anyApproved) ? (jTLULP * scoreFactor * slaFactor) : 0;
    recipients.push({
      no: recipients.length + 1, nip: id, nipPegawai: String(cfg.unitNips[id] || ""), nama: customSpvUlpName, koordinator: "-", jabatan, ulp: ut.ulp, up3: ut.up3,
      excelScore: Number(avgExcelScore.toFixed(2)),
      breakdown,
      score: finalScore, status: ut.anyApproved ? "approved" : "not approved",
      koefJabatan: jTLULP, bobot: Number(bobot.toFixed(4)), nilaiPerBobot: 0, reward: 0
    });
  });

  Object.values(up3Teams).forEach((at, idx) => {
    const id = makeSpvAreaId(at.up3);
    const jabatan = "SPV MANBILL AREA";
    const avgExcelScore = avg(at.excelScores);
    const areaSlaVal = cfg.areaSlaScores[at.up3];
    const customSpvAreaName = (cfg.spvAreaNames[at.up3] && String(cfg.spvAreaNames[at.up3]).trim())
      ? String(cfg.spvAreaNames[at.up3]).trim()
      : ("SPV MANBILL AREA " + at.up3);
    const breakdown = buildScoreBreakdown(jabatan, avgExcelScore, takeCustomScore(cfg, id, "SPV-AREA-" + (idx + 1)), areaSlaVal);
    const finalScore = breakdown.totalNilai;
    const scoreFactor = useScoreInWeight ? (finalScore / 100) : 1.0;
    const bobot = (!onlyApproved || at.anyApproved) ? (jTLArea * scoreFactor) : 0;
    recipients.push({
      no: recipients.length + 1, nip: id, nipPegawai: String(cfg.unitNips[id] || ""), nama: customSpvAreaName, koordinator: "-", jabatan, ulp: "AREA " + at.up3, up3: at.up3,
      excelScore: Number(avgExcelScore.toFixed(2)),
      breakdown,
      score: finalScore, status: at.anyApproved ? "approved" : "not approved",
      koefJabatan: jTLArea, bobot: Number(bobot.toFixed(4)), nilaiPerBobot: 0, reward: 0
    });

    const adminId = makeAdminId(at.up3);
    const jabatanAdmin = "ADMIN MANBILL";
    const customAdminName = (cfg.adminNames[at.up3] && String(cfg.adminNames[at.up3]).trim())
      ? String(cfg.adminNames[at.up3]).trim()
      : (cfg.defaultAdminName && String(cfg.defaultAdminName).trim() !== "" && String(cfg.defaultAdminName).trim() !== "ADMIN MANBILL"
          ? String(cfg.defaultAdminName).trim() + " (" + at.up3 + ")"
          : "ADMIN MANBILL " + at.up3);
    const breakdownAdmin = buildScoreBreakdown(jabatanAdmin, 100, takeCustomScore(cfg, adminId, "ADMIN-" + (idx + 1)), areaSlaVal);
    const finalScoreAdmin = breakdownAdmin.totalNilai;
    const scoreFactorAdmin = useScoreInWeight ? (finalScoreAdmin / 100) : 1.0;
    const bobotAdmin = (!onlyApproved || at.anyApproved) ? (jAdmin * scoreFactorAdmin) : 0;
    recipients.push({
      no: recipients.length + 1, nip: adminId, nipPegawai: String(cfg.unitNips[adminId] || ""), nama: customAdminName, koordinator: "-", jabatan: jabatanAdmin, ulp: "AREA " + at.up3, up3: at.up3,
      excelScore: 100,
      breakdown: breakdownAdmin,
      score: finalScoreAdmin, status: at.anyApproved ? "approved" : "not approved",
      koefJabatan: jAdmin, bobot: Number(bobotAdmin.toFixed(4)), nilaiPerBobot: 0, reward: 0
    });
  });

  recipients.forEach(item => { if (up3Map[item.up3]) up3Map[item.up3].totalBobot += item.bobot; });

  let totalRewardPoolJatim = 0;
  Object.values(up3Map).forEach(u => {
    u.totalBobot = Number(u.totalBobot.toFixed(4));
    const isAreaSlaEligible = Number(u.slaArea ?? 100) >= 100;
    const effectiveAreaPagu = isAreaSlaEligible ? u.paguArea : 0;
    u.nilaiPerBobot = (isAreaSlaEligible && u.totalBobot > 0) ? (effectiveAreaPagu / u.totalBobot) : 0;
    totalRewardPoolJatim += effectiveAreaPagu;
  });

  const totalBobotJatim = recipients.reduce((s, i) => s + i.bobot, 0);
  let totalDistributed = 0, rewardAdminTotal = 0, rewardPetugasTotal = 0, rewardKoordTotal = 0, rewardTLULPTotal = 0, rewardTLAreaTotal = 0;
  let countAdminReward = 0, countPetugasReward = 0, countKoordReward = 0, countTLULPReward = 0, countTLAreaReward = 0;

  const matrix = {
    admin: { count: 0, koefJ: jAdmin, totalBobot: 0, rewardPerPerson: 0, totalReward: 0 },
    petugas: { count: 0, koefJ: jPetugas, totalBobot: 0, rewardPerPerson: 0, totalReward: 0 },
    koordinator: { count: 0, koefJ: jKoord, totalBobot: 0, rewardPerPerson: 0, totalReward: 0 },
    tlUlp: { count: 0, koefJ: jTLULP, totalBobot: 0, rewardPerPerson: 0, totalReward: 0 },
    tlArea: { count: 0, koefJ: jTLArea, totalBobot: 0, rewardPerPerson: 0, totalReward: 0 }
  };

  const addToMatrix = (cell, item, rw) => {
    if (item.bobot > 0) {
      cell.count++;
      cell.totalBobot += item.bobot;
      cell.totalReward += rw;
    }
  };

  recipients.forEach(item => {
    const uObj = up3Map[item.up3], ulpObj = ulpMap[item.up3 + "||" + item.ulp];
    const areaRate = uObj ? uObj.nilaiPerBobot : 0;
    const rw = item.bobot * areaRate;
    item.nilaiPerBobot = Math.round(areaRate);
    item.reward = rw;
    totalDistributed += rw;

    if (item.jabatan === "PETUGAS MANBILL") {
      rewardPetugasTotal += rw;
      if (item.bobot > 0) countPetugasReward++;
      addToMatrix(matrix.petugas, item, rw);
      if (uObj) uObj.rewardPetugas += rw;
      if (ulpObj) ulpObj.rewardPetugas += rw;
    } else if (item.jabatan === "KORDINATOR") {
      rewardKoordTotal += rw;
      if (item.bobot > 0) countKoordReward++;
      addToMatrix(matrix.koordinator, item, rw);
      if (uObj) uObj.rewardKoordinator += rw;
      if (ulpObj) ulpObj.rewardKoordinator += rw;
    } else if (item.jabatan === "SPV MANBILL ULP") {
      rewardTLULPTotal += rw;
      if (item.bobot > 0) countTLULPReward++;
      addToMatrix(matrix.tlUlp, item, rw);
      if (uObj) uObj.rewardTLULP += rw;
      if (ulpObj) ulpObj.rewardTLULP += rw;
    } else if (item.jabatan === "SPV MANBILL AREA") {
      rewardTLAreaTotal += rw;
      if (item.bobot > 0) countTLAreaReward++;
      addToMatrix(matrix.tlArea, item, rw);
      if (uObj) uObj.rewardTLArea += rw;
    } else if (item.jabatan === "ADMIN MANBILL") {
      rewardAdminTotal += rw;
      if (item.bobot > 0) countAdminReward++;
      addToMatrix(matrix.admin, item, rw);
      if (uObj) uObj.rewardAdmin += rw;
    }

    if (uObj) uObj.totalReward += rw;
    if (ulpObj) ulpObj.totalReward += rw;
  });

  ["admin", "petugas", "koordinator", "tlUlp", "tlArea"].forEach(jKey => {
    const cell = matrix[jKey];
    cell.totalBobot = Number(cell.totalBobot.toFixed(2));
    cell.rewardPerPerson = cell.count > 0 ? (cell.totalReward / cell.count) : 0;
  });

  const rekapUP3 = Object.values(up3Map).sort((a, b) => b.total - a.total);
  rekapUP3.forEach(u => {
    u.totalBobot = Number(u.totalBobot.toFixed(2));
    u.nilaiPerBobot = Math.round(u.nilaiPerBobot);
    u.pctApproved = u.total > 0 ? ((u.approved / u.total) * 100).toFixed(1) : "0.0";
    u.pctReward = totalDistributed > 0 ? ((u.totalReward / totalDistributed) * 100).toFixed(1) : "0.0";
  });

  const rekapULP = Object.values(ulpMap).sort((a, b) => b.total - a.total);
  rekapULP.forEach(u => {
    const parentArea = up3Map[u.up3];
    u.paguArea = parentArea ? parentArea.paguArea : 0;
    u.pctApproved = u.total > 0 ? ((u.approved / u.total) * 100).toFixed(1) : "0.0";
    u.pctFromArea = (parentArea && parentArea.totalReward > 0) ? ((u.totalReward / parentArea.totalReward) * 100).toFixed(1) : "0.0";
  });

  const avgNilaiPerBobotJatim = totalBobotJatim > 0 ? (totalDistributed / totalBobotJatim) : 0;

  return {
    total: totalData, approved, notApproved,
    pctApproved: totalData > 0 ? ((approved / totalData) * 100).toFixed(1) : "0.0",
    pctNotApproved: totalData > 0 ? ((notApproved / totalData) * 100).toFixed(1) : "0.0",
    rekapUP3, rekapULP, rewardConfig: cfg,
    rewardSummary: {
      totalPool: Math.round(totalRewardPoolJatim),
      totalDistributed: Math.round(totalDistributed),
      pctFromGlobal: "100.0",
      totalBobot: Number(totalBobotJatim.toFixed(2)),
      nilaiPerBobot: Math.round(avgNilaiPerBobotJatim),
      rewardAdmin: Math.round(rewardAdminTotal),
      rewardPetugas: Math.round(rewardPetugasTotal),
      rewardKoordinator: Math.round(rewardKoordTotal),
      rewardTLULP: Math.round(rewardTLULPTotal),
      rewardTLArea: Math.round(rewardTLAreaTotal),
      countAdmin: countAdminReward,
      countPetugas: countPetugasReward,
      countKoordinator: countKoordReward,
      countTLULP: countTLULPReward,
      countTLArea: countTLAreaReward,
      matrix
    },
    recipients
  };
}

/* ============================== OUTPUT KE SHEET ============================== */

// Data_Reward menyimpan semua periode (kolom BLTH); hanya baris periode yang dihitung yang diganti
function buildRewardSheet(ss, stats, blth) {
  const rSheet = ss.getSheetByName(SHEET_REWARD_NAME) || ss.insertSheet(SHEET_REWARD_NAME);
  const headers = [
    "BLTH", "Periode", "No", "NIP / ID", "Nama Penerima", "Jabatan", "ULP", "Area / UP3",
    "Total Excel / MO (%)", "Kinerja Utama (%)", "Kinerja Tambahan (%)", "Absensi (%)",
    "Penilaian Atasan (%)", "Total Nilai (%)", "Status", "Bobot Jabatan", "Total Nilai x Bobot",
    "Indeks / Bobot Area (Rp)", "Nominal Reward (Rp)"
  ];

  // Baris periode lain yang sudah ada dipertahankan
  let keptRows = [];
  if (rSheet.getLastRow() > 1) {
    const old = rSheet.getRange(1, 1, rSheet.getLastRow(), rSheet.getLastColumn()).getValues();
    const oldHasBlth = String(old[0][0]).trim().toUpperCase() === "BLTH";
    keptRows = old.slice(1)
      .filter(r => r.some(c => String(c).trim() !== ""))
      .map(r => oldHasBlth ? r.slice(0, headers.length) : [LEGACY_BLTH, blthLabel(LEGACY_BLTH)].concat(r).slice(0, headers.length))
      .filter(r => cellToBlth(r[0]) !== blth)
      .map(r => { r[0] = cellToBlth(r[0]); while (r.length < headers.length) r.push(""); return r; });
  }

  const label = blthLabel(blth);
  const newRows = (stats.recipients || []).map(item => {
    const b = item.breakdown || {};
    return [
      blth, label, item.no, item.nip, item.nama, item.jabatan, item.ulp, item.up3,
      item.excelScore ?? 100, b.subUtama ?? 0, b.tambahan ?? 0, b.absensi ?? 0,
      b.subAtasan ?? 0, item.score, item.status, item.koefJabatan, item.bobot,
      item.nilaiPerBobot, Math.round(item.reward)
    ];
  });

  const rows = keptRows.concat(newRows).sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  const allValues = [headers].concat(rows);
  resetSheetSafely(rSheet, allValues.length + 5, headers.length + 2);
  rSheet.getRange(1, 1, allValues.length, 1).setNumberFormat("@");
  rSheet.getRange(1, 1, allValues.length, headers.length).setValues(allValues);
  rSheet.getRange(1, 1, 1, headers.length).setBackground("#003b5c").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  rSheet.setFrozenRows(1);

  if (rows.length > 0) {
    rSheet.getRange(2, 18, rows.length, 2).setNumberFormat('"Rp"#,##0');
  }
}

// Sheet Dashboard menampilkan rekap periode yang terakhir dihitung
function buildSpreadsheetDashboard(ss, stats, blth) {
  const dash = ss.getSheetByName(SHEET_DASHBOARD_NAME) || ss.insertSheet(SHEET_DASHBOARD_NAME, 0);
  const hUP3 = ["No", "Area / UP3", "Approved", "Not App", "Total Petugas", "Nilai Reward (Rp)", "Nilai Denda (Rp)", "Reward Netto (Rp)", "Pagu Netto 70% (Rp)", "Indeks/Bobot", "Rw SPV Area", "Rw SPV ULP", "Rw Kordinator", "Rw Petugas", "Rw Admin", "Total Reward (Rp)"];
  const maxRowsNeeded = Math.max((stats.rekapUP3 || []).length, (stats.rekapULP || []).length) + 18;
  resetSheetSafely(dash, maxRowsNeeded, 24);

  const rw = stats.rewardSummary;
  const fmtRp = val => "Rp " + Math.round(val || 0).toLocaleString("id-ID");

  dash.getRange(1, 1, 1, hUP3.length).merge().setValue("APIK — REKAPITULASI PENILAIAN & INSENTIF KINERJA MANBILL — PERIODE " + blthLabel(blth).toUpperCase())
    .setFontSize(13).setFontWeight("bold").setBackground("#003b5c").setFontColor("#ffffff").setHorizontalAlignment("center");
  dash.getRange(2, 1, 1, hUP3.length).merge().setValue("Periode: " + blthLabel(blth) + " (BLTH " + blth + ") | Total Akumulasi Pagu Jatim: " + fmtRp(rw.totalPool) + " | Total Bobot: " + rw.totalBobot)
    .setFontSize(10).setBackground("#e0f2fe").setHorizontalAlignment("center");

  dash.getRange(4, 1, 1, hUP3.length).setValues([hUP3]).setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");

  const rowsUP3 = (stats.rekapUP3 || []).map((u, i) => [
    i + 1, u.up3, u.approved, u.notApproved, u.total,
    Math.round(u.rewardKotor || 0), Math.round(u.dendaArea || 0), Math.round(u.rewardNetto || 0),
    Math.round(u.paguArea), Math.round(u.nilaiPerBobot),
    Math.round(u.rewardTLArea), Math.round(u.rewardTLULP), Math.round(u.rewardKoordinator),
    Math.round(u.rewardPetugas), Math.round(u.rewardAdmin || 0), Math.round(u.totalReward)
  ]);

  if (rowsUP3.length > 0) {
    dash.getRange(5, 1, rowsUP3.length, hUP3.length).setValues(rowsUP3);
    dash.getRange(5, 6, rowsUP3.length, hUP3.length - 5).setNumberFormat('"Rp"#,##0');
  }
}

/* ============================== AMBIL DATA AWAL ============================== */

/**
 * Ambil data satu periode. blth kosong → periode terbaru yang ada datanya.
 * Hanya baris periode tsb yang dikirim (tanpa kolom BLTH) agar payload kecil & login cepat.
 */
function getExistingData(blth) {
  try {
    const store = getConfigStore();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const all = ss ? readAllDataRows(ss) : { headers: [], rows: [] };
    const periods = listDataPeriods(all);
    let selected = isValidBlth(blth) ? String(blth) : (periods.length ? periods[periods.length - 1] : LEGACY_BLTH);
    const rows = getPeriodRows(all, selected);
    return {
      success: rows.length > 0,
      blth: selected,
      periods: periods,
      unitDirectory: buildUnitDirectory(all),
      lockedPeriods: store.lockedPeriods,
      submissionStatuses: Object.keys(store.submissions).reduce((m, k) => { const s = store.submissions[k]; if (s && s.status) m[k] = s.status; return m; }, {}),
      rawHeaders: all.headers,
      rawRows: rows,
      rewardConfig: flattenConfigForPeriod(store, selected)
    };
  } catch (err) {
    return { success: false, message: err.toString(), blth: LEGACY_BLTH, periods: [], lockedPeriods: {} };
  }
}
