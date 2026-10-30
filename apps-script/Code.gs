/**
 * TIMELINE KALENDER — penghubung Google Sheet <-> kalender (GitHub Pages)
 * Pasang di: Google Sheet > Ekstensi > Apps Script, lalu Deploy sebagai Web App.
 *
 * Kolom dibaca berdasarkan NAMA HEADER di baris 1 (urutan kolom bebas).
 * Tab KONTEN : ID | TANGGAL | JUDUL KONTEN | JENIS KONTEN | BRAND TAP IN | AKUN | SEGMEN | ... | DIUBAH
 * Tab BRAND  : BRAND | WARNA
 */
const TZ = 'Asia/Jakarta';
const JENIS = { carousel: 'Carousel', short: 'Vid.Pendek', long: 'Vid.Panjang', story: 'Story' };

// key di kalender  ->  nama header di sheet
const FIELDS = {
  id: 'ID', date: 'TANGGAL', title: 'JUDUL KONTEN', type: 'JENIS KONTEN', brand: 'BRAND TAP IN',
  akun: 'AKUN', segmen: 'SEGMEN',
  jamYT: 'JAM UP YT', jamIG: 'JAM UP IG', jamTT: 'JAM UP TT', jamFB: 'JAM UP FB',
  producer: 'PRODUCER', writing: 'WRITING', shooting: 'SHOOTING', designer: 'DESIGNER', editor: 'EDITOR',
  talent1: 'TALENT 1', talent2: 'TALENT 2', production: 'PRODUCTION',
  stYT: 'STATUS YT', stIG: 'STATUS IG', stTT: 'STATUS TT', stFB: 'STATUS FB',
  idYT: 'POST ID YT', idIG: 'POST ID IG', idTT: 'POST ID TT', idFB: 'POST ID FB',
  catatan: 'CATATAN', eventId: 'EVENT ID', updated: 'DIUBAH'
};

function sh_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
// pastikan semua header ada (yang belum ada ditambah di kanan), kembalikan {key: kolom (1-based)}
function cols_() {
  const ks = sh_('KONTEN');
  let lastCol = Math.max(ks.getLastColumn(), 1);
  const head = ks.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) { return String(h).trim().toUpperCase(); });
  const map = {};
  Object.keys(FIELDS).forEach(function (k) {
    let i = head.indexOf(FIELDS[k]);
    if (i < 0) {
      if (ks.getMaxColumns() <= lastCol) ks.insertColumnsAfter(lastCol, 1);
      lastCol++;
      ks.getRange(1, lastCol).setValue(FIELDS[k]);
      head.push(FIELDS[k]);
      i = lastCol - 1;
    }
    map[k] = i + 1;
  });
  map._width = lastCol;
  return map;
}
function typeId_(v) {
  const s = String(v || '').toLowerCase().replace(/[^a-z]/g, '');
  if (s.indexOf('carousel') === 0 || s === 'crs' || s === 'desain') return 'carousel';
  if (s.indexOf('story') === 0 || s === 'str') return 'story';
  if (s.indexOf('panjang') >= 0 || s === 'pjg' || s === 'long') return 'long';
  if (s.indexOf('pendek') >= 0 || s === 'pdk' || s === 'short') return 'short';
  return 'carousel';
}
function date_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/); // dd/mm/yyyy
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}
function newId_() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

function read_() {
  const ks = sh_('KONTEN'), bs = sh_('BRAND');
  const c = cols_();
  const items = [];
  const last = ks.getLastRow();
  if (last > 1) {
    const rng = ks.getRange(2, 1, last - 1, c._width);
    const raw = rng.getValues(), disp = rng.getDisplayValues();
    disp.forEach(function (r, i) {
      const title = String(r[c.title - 1] || '').trim();
      const date = date_(raw[i][c.date - 1]);
      if (!title || !date) return;
      const it = {};
      Object.keys(FIELDS).forEach(function (k) { if (k !== 'updated') it[k] = String(r[c[k] - 1] || '').trim(); });
      it.title = title; it.date = date; it.type = typeId_(it.type);
      if (!it.id) { it.id = newId_(); ks.getRange(i + 2, c.id).setValue(it.id); } // baris yang diisi manual di sheet
      items.push(it);
    });
  }
  const brandColors = {};
  const bl = bs.getLastRow();
  if (bl > 1) {
    bs.getRange(2, 1, bl - 1, 2).getValues().forEach(function (r) {
      const b = String(r[0] || '').trim().toLowerCase(), col = String(r[1] || '').trim();
      if (b && /^#[0-9a-f]{6}$/i.test(col)) brandColors[b] = col;
    });
  }
  items.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return { items: items, brandColors: brandColors };
}

function rowOf_(id, c) {
  const ks = sh_('KONTEN'), last = ks.getLastRow();
  if (last < 2) return -1;
  const ids = ks.getRange(2, c.id, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function upsert_(it) {
  if (!it || !it.id || !it.title || !it.date) return;
  const ks = sh_('KONTEN'), c = cols_();
  const r = rowOf_(it.id, c);
  const row = r > 0 ? r : ks.getLastRow() + 1;
  const range = ks.getRange(row, 1, 1, c._width);
  const vals = r > 0 ? range.getValues()[0] : new Array(c._width).fill('');
  Object.keys(FIELDS).forEach(function (k) {
    let v;
    if (k === 'updated') v = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    else if (k === 'type') v = JENIS[it.type] || 'Carousel';
    else v = it[k] == null ? '' : String(it[k]);
    vals[c[k] - 1] = v;
  });
  range.setNumberFormat('@'); // POST ID panjang jangan jadi angka
  range.setValues([vals]);
}
function remove_(id) { const c = cols_(); const r = rowOf_(id, c); if (r > 0) sh_('KONTEN').deleteRow(r); }
function color_(brand, color) {
  const b = String(brand || '').trim().toLowerCase();
  if (!b || !/^#[0-9a-f]{6}$/i.test(color)) return;
  const bs = sh_('BRAND'), last = bs.getLastRow();
  if (last > 1) {
    const names = bs.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim().toLowerCase() === b) { bs.getRange(i + 2, 2).setValue(color); return; }
    }
  }
  bs.appendRow([b, color]);
}

function doGet() { return out_(read_()); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const op = JSON.parse(e.postData.contents || '{}');
    if (op.op === 'upsert') upsert_(op.item);
    else if (op.op === 'delete') remove_(op.id);
    else if (op.op === 'color') color_(op.brand, op.color);
    else if (op.op === 'bulk') { // pindahan data lama dari browser, cuma kalau sheet masih kosong
      if (read_().items.length === 0) {
        (op.items || []).forEach(upsert_);
        Object.keys(op.brandColors || {}).forEach(function (b) { color_(b, op.brandColors[b]); });
      }
    }
    SpreadsheetApp.flush();
    return out_({ ok: true, data: read_() });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
