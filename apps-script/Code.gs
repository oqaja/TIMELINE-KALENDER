/**
 * TIMELINE KALENDER — penghubung Google Sheet <-> kalender (GitHub Pages)
 * Pasang di: Google Sheet > Ekstensi > Apps Script, lalu Deploy sebagai Web App.
 *
 * Tab KONTEN : ID | TANGGAL | JUDUL KONTEN | JENIS KONTEN | BRAND TAP IN | DIUBAH
 * Tab BRAND  : BRAND | WARNA
 */
const TZ = 'Asia/Jakarta';
const JENIS = { carousel: 'Carousel', short: 'Vid.Pendek', long: 'Vid.Panjang', story: 'Story' };

function sh_(name) { return SpreadsheetApp.getActive().getSheetByName(name); }
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function typeId_(v) {
  const s = String(v || '').toLowerCase().replace(/[^a-z]/g, '');
  if (s.indexOf('carousel') === 0 || s === 'crs') return 'carousel';
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
  const items = [];
  const last = ks.getLastRow();
  if (last > 1) {
    const vals = ks.getRange(2, 1, last - 1, 6).getValues();
    vals.forEach(function (r, i) {
      const title = String(r[2] || '').trim();
      const date = date_(r[1]);
      if (!title || !date) return;
      let id = String(r[0] || '').trim();
      if (!id) { id = newId_(); ks.getRange(i + 2, 1).setValue(id); } // baris yang diisi manual di sheet
      items.push({ id: id, date: date, title: title, type: typeId_(r[3]), brand: String(r[4] || '').trim() });
    });
  }
  const brandColors = {};
  const bl = bs.getLastRow();
  if (bl > 1) {
    bs.getRange(2, 1, bl - 1, 2).getValues().forEach(function (r) {
      const b = String(r[0] || '').trim().toLowerCase(), c = String(r[1] || '').trim();
      if (b && /^#[0-9a-f]{6}$/i.test(c)) brandColors[b] = c;
    });
  }
  items.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return { items: items, brandColors: brandColors };
}

function rowOf_(id) {
  const ks = sh_('KONTEN'), last = ks.getLastRow();
  if (last < 2) return -1;
  const ids = ks.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function rowData_(it) {
  return [it.id, it.date, it.title, JENIS[it.type] || 'Carousel', it.brand || '',
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')];
}
function upsert_(it) {
  if (!it || !it.id || !it.title || !it.date) return;
  const ks = sh_('KONTEN');
  const r = rowOf_(it.id);
  const row = r > 0 ? r : ks.getLastRow() + 1;
  ks.getRange(row, 2).setNumberFormat('@');
  ks.getRange(row, 1, 1, 6).setValues([rowData_(it)]);
}
function remove_(id) { const r = rowOf_(id); if (r > 0) sh_('KONTEN').deleteRow(r); }
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
