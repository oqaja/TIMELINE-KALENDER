/**
 * TIMELINE KALENDER — penghubung Google Sheet <-> kalender (GitHub Pages)
 * Pasang di: Google Sheet > Ekstensi > Apps Script, lalu Deploy sebagai Web App.
 *
 * CATATAN: sekarang Apps Script cuma berisi apps-script/Loader.gs. File INI diambil otomatis dari GitHub
 * oleh loader (±5–10 menit setelah push), jadi update cukup di-push ke repo — tanpa copas/deploy ulang.
 * Wajib tetap ada teks "TIMELINE KALENDER" dan "function doPost" di file ini (dicek loader).
 *
 * Kolom dibaca berdasarkan NAMA HEADER di baris 1 (urutan kolom bebas).
 * Kolom tambahan di luar daftar FIELDS tetap ikut terbaca/tersimpan lewat item.cols {HEADER: nilai},
 * jadi kolom baru cukup ditambah dari sisi kalender — script tidak perlu di-deploy ulang.
 * Tab KONTEN : ID | TANGGAL | JUDUL KONTEN | JENIS KONTEN | BRAND TAP IN | AKUN | SEGMEN | ... | DIUBAH
 * Tab KONTEN : sekarang khusus PREVIEW (tidak pernah masuk ke KALENDER KONTEN)
 * Tab BRAND  : BRAND | WARNA
 * Tab TODO   : ID | TANGGAL | ORANG | TUGAS | SELESAI | CATATAN | DIUBAH   (to-do manual per orang — dibuat otomatis)
 *
 * Konten beneran dibaca & ditulis langsung ke DASHBOARD KULTURAMA V1 > KALENDER KONTEN.
 * Di sana struktur kolom TIDAK diubah: baris dicari lewat EVENT ID, atau judul + tanggal,
 * dan yang ditulis cuma sel yang berubah (biar catatan/status dari otomasi lain nggak ketimpa).
 *
 * PIN EDIT: semua perubahan (tambah/edit/hapus) wajib bawa PIN. PIN TIDAK ditulis di kode ini —
 * isi lewat Apps Script > Setelan project (⚙) > Properti skrip:  EDIT_PIN = <pin kamu>
 * Tanpa EDIT_PIN, semua perubahan ditolak (kalender cuma bisa dilihat).
 * Tab OPSI   : DAFTAR | NILAI   (pilihan tambahan: akun, segmen, tim, status — dibuat otomatis;
 *              daftar '<nama>_hapus' = pilihan yang disembunyiin, mis. tim_hapus = orang yang dihapus dari daftar)
 */
const TZ = 'Asia/Jakarta';
const SCRIPT_VER = 3; // kalender pakai fitur baru (cek bentrok, preview per kolom) cuma kalau versi ini >= 2
const JENIS = { carousel: 'Desain', short: 'Video Pendek', long: 'Video Panjang', story: 'Story' };

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
function ensureCol_(map, header) {
  const h = String(header).trim().toUpperCase();
  let i = map._head.indexOf(h);
  if (i >= 0) return i + 1;
  const ks = sh_('KONTEN');
  if (ks.getMaxColumns() <= map._width) ks.insertColumnsAfter(map._width, 1);
  map._width++;
  ks.getRange(1, map._width).setValue(h);
  map._head.push(h);
  return map._width;
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
  map._head = head;
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
  m = s.match(/^(\d{1,2})[\s\-]([A-Za-z]{3})[A-Za-z]*[\s\-](\d{4})$/); // 02-Nov-2026 / 30-Okt-2026
  if (m) {
    const mon = { jan: 1, feb: 2, mar: 3, apr: 4, mei: 5, may: 5, jun: 6, jul: 7, agu: 8, ags: 8, aug: 8, sep: 9, okt: 10, oct: 10, nov: 11, des: 12, dec: 12 }[m[2].toLowerCase()];
    if (mon) return m[3] + '-' + ('0' + mon).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  return '';
}
function opsiSheet_() {
  let s = sh_('OPSI');
  if (!s) { s = SpreadsheetApp.getActive().insertSheet('OPSI'); s.getRange(1, 1, 1, 2).setValues([['DAFTAR', 'NILAI']]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function readOpsi_() {
  const s = opsiSheet_(), last = s.getLastRow(), o = {};
  if (last > 1) s.getRange(2, 1, last - 1, 2).getDisplayValues().forEach(function (r) {
    const l = String(r[0]).trim().toLowerCase(), v = String(r[1]).trim();
    if (l && v) { o[l] = o[l] || []; if (o[l].indexOf(v) < 0) o[l].push(v); }
  });
  return o;
}
function addOpsi_(list, value) {
  const l = String(list || '').trim().toLowerCase(), v = String(value || '').trim();
  if (!l || !v) return;
  const cur = readOpsi_()[l] || [];
  if (cur.some(function (x) { return x.toLowerCase() === v.toLowerCase(); })) return;
  opsiSheet_().appendRow([l, v]);
}
// ===================== TO-DO MANUAL (tab TODO) =====================
const TODO_HEAD = ['ID', 'TANGGAL', 'ORANG', 'TUGAS', 'SELESAI', 'CATATAN', 'DIUBAH'];
function todoSheet_() {
  let s = sh_('TODO');
  if (!s) {
    s = SpreadsheetApp.getActive().insertSheet('TODO');
    s.getRange(1, 1, 1, TODO_HEAD.length).setValues([TODO_HEAD]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function readTodo_() {
  const s = todoSheet_(), last = s.getLastRow(), out = [];
  if (last < 2) return out;
  const raw = s.getRange(2, 1, last - 1, 6).getValues(), disp = s.getRange(2, 1, last - 1, 6).getDisplayValues();
  for (let i = 0; i < raw.length; i++) {
    const text = String(disp[i][3] || '').trim(), date = date_(raw[i][1]) || date_(disp[i][1]);
    if (!text || !date) continue;
    let id = String(disp[i][0] || '').trim();
    if (!id) { id = newId_(); s.getRange(i + 2, 1).setValue(id); }
    out.push({ id: id, date: date, who: String(disp[i][2] || '').trim(), text: text, done: String(disp[i][4] || '').trim() === '✅', note: String(disp[i][5] || '').trim() });
  }
  return out;
}
function todoRow_(id) {
  const s = todoSheet_(), last = s.getLastRow();
  if (last < 2) return -1;
  const ids = s.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function todoUpsert_(t) {
  if (!t || !t.id || !t.text || !t.date) throw new Error('to-do butuh tugas & tanggal');
  const s = todoSheet_(), r = todoRow_(t.id), row = r > 0 ? r : s.getLastRow() + 1;
  const rng = s.getRange(row, 1, 1, TODO_HEAD.length);
  rng.setNumberFormat('@');
  rng.setValues([[t.id, t.date, t.who || '', t.text, t.done ? '✅' : '', t.note || '', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')]]);
}
function todoDelete_(id) { const r = todoRow_(id); if (r > 0) todoSheet_().deleteRow(r); }
function delOpsi_(list, value) {
  const l = String(list || '').trim().toLowerCase(), v = String(value || '').trim().toLowerCase();
  const s = opsiSheet_();
  for (let r = s.getLastRow(); r >= 2; r--) {
    const row = s.getRange(r, 1, 1, 2).getDisplayValues()[0];
    if (String(row[0]).trim().toLowerCase() === l && String(row[1]).trim().toLowerCase() === v) s.deleteRow(r);
  }
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
      const known = Object.keys(FIELDS).map(function (k) { return FIELDS[k]; });
      it.cols = {};
      c._head.forEach(function (h, j) { if (h && known.indexOf(h) < 0) it.cols[h] = String(r[j] || '').trim(); });
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
  return { items: items, brandColors: brandColors, options: readOpsi_() };
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
  Object.keys(it.cols || {}).forEach(function (h) { ensureCol_(c, h); });
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
  Object.keys(it.cols || {}).forEach(function (h) {
    const j = c._head.indexOf(String(h).trim().toUpperCase());
    if (j >= 0) vals[j] = it.cols[h] == null ? '' : String(it.cols[h]);
  });
  range.setNumberFormat('@'); // POST ID panjang jangan jadi angka
  range.setValues([vals]);
}
function pvUpdate_(id, changes, cols, prev, force) {
  const ks = sh_('KONTEN'), c = cols_();
  Object.keys(cols || {}).forEach(function (h) { ensureCol_(c, h); });
  const r = rowOf_(id, c);
  if (r < 0) throw new Error('notfound: preview ini sudah dihapus orang lain');
  const range = ks.getRange(r, 1, 1, c._width), disp = range.getDisplayValues()[0];
  const set = []; // [kolom, nilai baru, nilai lama yang dilihat pengguna, nama]
  Object.keys(changes || {}).forEach(function (k) {
    if (!FIELDS[k] || k === 'updated' || k === 'id') return;
    set.push([c[k], k === 'type' ? (JENIS[changes[k]] || 'Desain') : String(changes[k] == null ? '' : changes[k]), prev ? prev[k] : undefined, FIELDS[k], k]);
  });
  Object.keys(cols || {}).forEach(function (h) {
    const j = c._head.indexOf(String(h).trim().toUpperCase());
    if (j >= 0) set.push([j + 1, String(cols[h] == null ? '' : cols[h]), prev && prev.cols ? prev.cols[h] : undefined, h, 'col']);
  });
  if (prev && !force) {
    const clash = [];
    set.forEach(function (x) {
      if (x[2] === undefined) return;
      const k = x[4] === 'col' ? 'x' : x[4];
      if (conflicts_(k, disp[x[0] - 1], k === 'type' ? (JENIS[x[2]] || x[2]) : x[2], x[1])) clash.push({ k: x[4], col: x[3], cur: String(disp[x[0] - 1]).trim() });
    });
    if (clash.length) return { conflict: clash };
  }
  set.forEach(function (x) { ks.getRange(r, x[0]).setNumberFormat('@').setValue(x[1]); });
  ks.getRange(r, c.updated).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
  return r;
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

// ===================== KALENDER KONTEN (DASHBOARD KULTURAMA V1) =====================
const DASH_ID = '1raiIO1HccW7IxN9bh9BqUJ7DOHDVBN05GKRQ5xrrfeI';
const DASH_TAB = 'KALENDER KONTEN';
const KK = {
  date: 'TANGGAL', title: 'JUDUL KONTEN', akun: 'AKUN', segmen: 'SEGMEN',
  jamYT: 'JAM UP YT', jamIG: 'JAM UP IG', jamTT: 'JAM UP TT', jamFB: 'JAM UP FB',
  type: 'JENIS KONTEN', producer: 'PRODUCER', writing: 'WRITING', shooting: 'SHOOTING',
  designer: 'DESIGNER', editor: 'EDITOR', talent1: 'TALENT 1', talent2: 'TALENT 2',
  production: 'PRODUCTION', cover: 'COVER/THUMB',
  stYT: 'STATUS YT', stIG: 'STATUS IG', stTT: 'STATUS TT', stFB: 'STATUS FB',
  idYT: 'POST ID YT', idIG: 'POST ID IG', idTT: 'POST ID TT', idFB: 'POST ID FB',
  brand: 'BRAND TAP IN', catatan: 'CATATAN', eventId: 'EVENT ID'
};
const KK_TEXT = ['jamYT', 'jamIG', 'jamTT', 'jamFB', 'idYT', 'idIG', 'idTT', 'idFB']; // simpan sebagai teks
let kkSs_ = null;
function kkBook_() { return kkSs_ || (kkSs_ = SpreadsheetApp.openById(DASH_ID)); }
function kkSheet_() { return kkBook_().getSheetByName(DASH_TAB); }
// urutan kolom standar KALENDER KONTEN — dipakai cadangan kalau judul kolomnya kehapus
const KK_DEFAULT = ['date', 'title', 'akun', 'segmen', 'jamYT', 'jamIG', 'jamTT', 'jamFB', 'type', 'producer', 'writing', 'shooting',
  'designer', 'editor', 'talent1', 'talent2', 'production', 'cover', 'stYT', 'stIG', 'stTT', 'stFB', 'idYT', 'idIG', 'idTT', 'idFB',
  'brand', 'catatan', 'eventId'];
let kkWarn_ = [];
function kkCols_(sh) {
  const w = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, w).getDisplayValues()[0].map(function (h) { return String(h).trim().toUpperCase(); });
  const c = { _width: w }, warn = [];
  Object.keys(KK).forEach(function (k) {
    let i = head.indexOf(KK[k]);
    if (i < 0) { // judul hilang: pakai posisi standar, tapi cuma kalau sel judul di posisi itu kosong
      const d = KK_DEFAULT.indexOf(k);
      if (d >= 0 && d < w && !head[d]) { i = d; warn.push(KK[k]); }
    }
    c[k] = i + 1; // 0 = kolom tidak ada
  });
  kkWarn_ = warn;
  if (!c.date || !c.title) throw new Error('Kolom TANGGAL / JUDUL KONTEN tidak ketemu di ' + DASH_TAB);
  return c;
}
function kkDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, kkBook_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return date_(v);
}
function kkRead_(from, to) {
  const sh = kkSheet_(), c = kkCols_(sh), last = sh.getLastRow(), items = [];
  if (last < 2) return items;
  const rng = sh.getRange(2, 1, last - 1, c._width), raw = rng.getValues(), disp = rng.getDisplayValues();
  for (let i = 0; i < raw.length; i++) {
    const title = String(disp[i][c.title - 1] || '').trim();
    const date = kkDate_(raw[i][c.date - 1]) || date_(disp[i][c.date - 1]);
    if (!title || !date) continue;
    if ((from && date < from) || (to && date > to)) continue;
    const it = { src: 'kk', row: i + 2 };
    Object.keys(KK).forEach(function (k) { it[k] = c[k] ? String(disp[i][c[k] - 1] || '').trim() : ''; });
    it.title = title; it.date = date;
    it.jenis = it.type; it.type = typeId_(it.type);
    it.id = it.eventId ? 'ev:' + it.eventId : 'row:' + it.row;
    items.push(it);
  }
  return items;
}
// pilihan dropdown langsung dari aturan validasi di KALENDER KONTEN
function kkLists_() {
  const sh = kkSheet_(), c = kkCols_(sh), row = Math.max(2, sh.getLastRow());
  const want = { akun: 'akun', segmen: 'segmen', producer: 'tim', stYT: 'status', brand: 'brand' }, out = {};
  Object.keys(want).forEach(function (k) {
    if (!c[k]) return;
    try {
      const rule = sh.getRange(row, c[k]).getDataValidation();
      if (rule && String(rule.getCriteriaType()) === 'VALUE_IN_LIST') out[want[k]] = (rule.getCriteriaValues()[0] || []).map(String);
    } catch (e) {}
  });
  return out;
}
// cari baris: EVENT ID dulu, lalu baris lama (dicek judul+tanggal), lalu judul+tanggal yang unik
function kkFind_(sh, c, ref) {
  const last = sh.getLastRow();
  if (!ref || last < 2) return -1;
  if (ref.eventId && c.eventId) {
    const ev = sh.getRange(2, c.eventId, last - 1, 1).getDisplayValues();
    for (let i = 0; i < ev.length; i++) if (String(ev[i][0]).trim() === String(ref.eventId)) return i + 2;
  }
  const titles = sh.getRange(2, c.title, last - 1, 1).getDisplayValues();
  const dates = sh.getRange(2, c.date, last - 1, 1).getValues();
  const same = function (i) { return String(titles[i][0]).trim() === String(ref.title) && kkDate_(dates[i][0]) === ref.date; };
  if (ref.row >= 2 && ref.row <= last && same(ref.row - 2)) return ref.row;
  let found = -1;
  for (let i = 0; i < titles.length; i++) if (same(i)) { if (found > 0) return -2; found = i + 2; }
  return found;
}
function kkVal_(k, v) {
  v = v == null ? '' : String(v);
  if (k === 'type') return JENIS[v] || v;
  if (KK_TEXT.indexOf(k) >= 0 && v) return "'" + v;
  return v; // tanggal yyyy-MM-dd dibaca Sheets sebagai tanggal
}
// bandingkan isi sel sekarang dengan yang dilihat pengguna terakhir kali (prev)
function same_(k, cur, prev) {
  cur = String(cur == null ? '' : cur).trim(); prev = String(prev == null ? '' : prev).trim();
  if (k === 'type') return typeId_(cur) === typeId_(prev);
  if (k === 'date') return (date_(cur) || cur) === (date_(prev) || prev);
  return cur === prev;
}
function conflicts_(k, cur, prev, next) { return !same_(k, cur, prev) && !same_(k, cur, next); }
function kkUpdate_(ref, changes, prev, force) {
  const sh = kkSheet_(), c = kkCols_(sh), row = kkFind_(sh, c, ref);
  if (row === -2) throw new Error('dobel: ada lebih dari satu baris "' + ref.title + '" di tanggal ' + ref.date);
  if (row < 0) throw new Error('notfound: baris "' + ref.title + '" (' + ref.date + ') tidak ketemu, mungkin sudah diubah/dihapus');
  const keys = Object.keys(changes || {}).filter(function (k) { return c[k] && k !== 'eventId'; });
  if (prev && !force) {
    const disp = sh.getRange(row, 1, 1, c._width).getDisplayValues()[0], raw = sh.getRange(row, 1, 1, c._width).getValues()[0];
    const clash = [];
    keys.forEach(function (k) {
      if (!(k in prev)) return;
      const cur = k === 'date' ? kkDate_(raw[c[k] - 1]) : disp[c[k] - 1];
      if (conflicts_(k, cur, prev[k], changes[k])) clash.push({ k: k, col: KK[k], cur: String(cur).trim() });
    });
    if (clash.length) return { conflict: clash, row: row };
  }
  keys.forEach(function (k) { sh.getRange(row, c[k]).setValue(kkVal_(k, changes[k])); });
  return row;
}
function kkAdd_(it) {
  if (!it || !it.title || !it.date) throw new Error('judul & tanggal wajib');
  const sh = kkSheet_(), c = kkCols_(sh);
  const vals = new Array(c._width).fill('');
  Object.keys(KK).forEach(function (k) { if (c[k] && k !== 'eventId') vals[c[k] - 1] = kkVal_(k, it[k]); });
  sh.appendRow(vals);
  return sh.getLastRow();
}
function kkDelete_(ref) {
  const sh = kkSheet_(), c = kkCols_(sh), row = kkFind_(sh, c, ref);
  if (row === -2) throw new Error('dobel: ada lebih dari satu baris "' + ref.title + '" di tanggal ' + ref.date);
  if (row > 0) sh.deleteRow(row);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const data = read_();
  data.items.forEach(function (it) { it.src = 'pv'; });
  data.ver = SCRIPT_VER;
  try { data.todos = readTodo_(); } catch (err) { data.todos = null; }
  try { data.kk = kkRead_(p.from || '', p.to || ''); data.lists = kkLists_(); if (kkWarn_.length) data.kkWarn = kkWarn_; }
  catch (err) { data.kk = null; data.kkError = String(err); }
  data.range = { from: p.from || '', to: p.to || '' };
  return out_(data);
}

// ---- PIN: cek + kunci sementara kalau salah terus (anti tebak-tebakan)
function pinCheck_(pin) {
  const want = PropertiesService.getScriptProperties().getProperty('EDIT_PIN');
  if (!want) throw new Error('pin_not_set: PIN belum diatur di Properti skrip (EDIT_PIN)');
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pin_fails') || 0);
  if (fails >= 10) throw new Error('pin_locked: kebanyakan salah PIN, coba lagi 15 menit lagi');
  if (String(pin || '') !== String(want)) {
    cache.put('pin_fails', String(fails + 1), 900);
    throw new Error('pin_wrong: PIN salah');
  }
}

function doPost(e) {
  let op0 = {};
  try { op0 = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  try { pinCheck_(op0.pin); } catch (err) { return out_({ ok: false, error: String(err) }); }
  if (op0.op === 'checkPin') return out_({ ok: true });
  if (op0.op === 'invoiceSave') { try { return out_(invoiceSave_(op0)); } catch (err) { return out_({ ok: false, error: String(err) }); } }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const op = JSON.parse(e.postData.contents || '{}');
    if (op.op === 'upsert') upsert_(op.item);
    else if (op.op === 'delete') remove_(op.id);
    else if (op.op === 'color') color_(op.brand, op.color);
    else if (op.op === 'option') addOpsi_(op.list, op.value);
    else if (op.op === 'optionDel') delOpsi_(op.list, op.value);
    else if (op.op === 'todoUpsert') todoUpsert_(op.todo);
    else if (op.op === 'todoDelete') todoDelete_(op.id);
    else if (op.op === 'kkAdd') { const row = kkAdd_(op.item); SpreadsheetApp.flush(); return out_({ ok: true, row: row }); }
    else if (op.op === 'kkUpdate') {
      const res = kkUpdate_(op.ref, op.changes, op.prev, op.force); SpreadsheetApp.flush();
      return out_(res && res.conflict ? { ok: false, conflict: res.conflict, error: 'conflict' } : { ok: true, row: res });
    }
    else if (op.op === 'pvUpdate') {
      const res = pvUpdate_(op.id, op.changes, op.cols, op.prev, op.force); SpreadsheetApp.flush();
      return out_(res && res.conflict ? { ok: false, conflict: res.conflict, error: 'conflict' } : { ok: true });
    }
    else if (op.op === 'kkDelete') { kkDelete_(op.ref); SpreadsheetApp.flush(); return out_({ ok: true }); }
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

// ---- INVOICE: salinan PDF dari invoice.html disimpan ke folder Drive SEBELUM boleh di-download
// Folder bisa diganti lewat Properti skrip: INVOICE_FOLDER_ID = <id folder>
const INVOICE_FOLDER_ID = '1pUNqU1aVozohDB_cRLr7EmmBSt0BjSrF';
function invoiceFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty('INVOICE_FOLDER_ID') || INVOICE_FOLDER_ID;
  return DriveApp.getFolderById(id);
}
function invoiceSave_(op) {
  const name = String(op.name || '').replace(/[\\/:*?"<>|]+/g, '_').trim();
  if (!/\.pdf$/i.test(name)) throw new Error('nama file harus .pdf');
  const b64 = String(op.data || '').replace(/^data:[^,]*,/, '');
  if (!b64) throw new Error('file kosong');
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length < 1000) throw new Error('file PDF rusak/kekecilan');
  const folder = invoiceFolder_();
  // invoice dengan nama yang sama (mis. di-download ulang setelah diedit) -> versi lama dibuang ke Sampah
  const old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  const f = folder.createFile(Utilities.newBlob(bytes, 'application/pdf', name));
  return { ok: true, fileId: f.getId(), url: f.getUrl(), name: f.getName() };
}
// Jalankan SEKALI dari editor (pilih fungsi ini > Run) buat ngasih izin akses Drive
function izinDrive() { Logger.log('Folder invoice: ' + invoiceFolder_().getName()); }
