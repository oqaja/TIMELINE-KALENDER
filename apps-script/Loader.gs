/**
 * LOADER — TIMELINE KALENDER + INVOICE (Shoe Police)
 * Tempel SEKALI di Apps Script (gantiin semua isi Code.gs), lalu file ini nggak perlu diubah lagi.
 *
 * Kode aslinya diambil otomatis dari GitHub: oqaja/timeline-kalender -> apps-script/Code.gs
 * Update di GitHub kepakai sendiri dalam ±5–10 menit. Nggak perlu copas / Deploy ulang.
 * Kalau GitHub lagi nggak bisa diakses, dipakai salinan terakhir yang tersimpan.
 *
 * Fungsi yang bisa di-Run manual dari editor:
 *   izinSemua      -> SEKALI di awal, buat ngasih izin (Sheet, Drive, internet)
 *   updateSekarang -> paksa ambil kode terbaru dari GitHub saat itu juga
 */
const LOADER_SRC = 'https://raw.githubusercontent.com/oqaja/timeline-kalender/main/apps-script/Code.gs';
const LOADER_TTL = 300; // detik: seberapa sering ngecek GitHub
const LOADER_KEY = 'LOADER_SRC_';

function doGet(e) { return app_('doGet')(e); }
function doPost(e) { return app_('doPost')(e); }

function izinSemua() {
  // nyentuh semua layanan yang dipakai kode utama, biar Google minta izinnya sekalian
  SpreadsheetApp.getActive().getName();
  DriveApp.getRootFolder().getName();
  UrlFetchApp.fetch(LOADER_SRC, { muteHttpExceptions: true });
  CacheService.getScriptCache().get('x');
  LockService.getScriptLock();
  PropertiesService.getScriptProperties().getKeys();
  updateSekarang();
  const izin = app_('izinDrive');
  if (izin) izin();
}

function updateSekarang() {
  CacheService.getScriptCache().remove('app_src');
  const src = fetchSrc_();
  if (!src) throw new Error('Gagal ambil kode dari GitHub');
  Logger.log('Kode terbaru kepasang (' + src.length + ' karakter).');
}

// ---------------------------------------------------------------------------
function fetchSrc_() {
  try {
    const r = UrlFetchApp.fetch(LOADER_SRC + '?t=' + Date.now(), { muteHttpExceptions: true });
    const s = r.getResponseCode() === 200 ? r.getContentText() : '';
    if (s.indexOf('TIMELINE KALENDER') >= 0 && s.indexOf('function doPost') >= 0) {
      CacheService.getScriptCache().put('app_src', s, LOADER_TTL);
      saveBackup_(s);
      return s;
    }
  } catch (err) {}
  return '';
}
function saveBackup_(s) {
  const p = PropertiesService.getScriptProperties(), size = 8000, n = Math.ceil(s.length / size), o = {};
  for (let i = 0; i < n; i++) o[LOADER_KEY + i] = s.slice(i * size, (i + 1) * size);
  o[LOADER_KEY + 'N'] = String(n);
  p.setProperties(o, false); // false = properti lain (EDIT_PIN dll) tetap aman
}
function readBackup_() {
  const all = PropertiesService.getScriptProperties().getProperties(), n = Number(all[LOADER_KEY + 'N'] || 0);
  let s = '';
  for (let i = 0; i < n; i++) s += all[LOADER_KEY + i] || '';
  return s;
}
let app__ = null;
function app_(name) {
  if (!app__) {
    let src = CacheService.getScriptCache().get('app_src') || fetchSrc_() || readBackup_();
    if (!src) throw new Error('loader: kode belum pernah ke-download dari GitHub. Run updateSekarang.');
    // eval di dalam closure ini bisa ngakses semua fungsi/const dari Code.gs
    app__ = new Function(src + '\n;return function (n) { try { return eval(n); } catch (e) { return null; } };')();
  }
  const f = app__(name);
  if (typeof f !== 'function' && name !== 'izinDrive') throw new Error('loader: fungsi ' + name + ' nggak ada di Code.gs');
  return f;
}
