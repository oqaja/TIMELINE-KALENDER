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
const SCRIPT_VER = 11; // kalender pakai fitur baru (cek bentrok, preview per kolom) cuma kalau versi ini >= 2
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
  if (l === 'brand') { try { kkListAdd_('brand', v); } catch (e) {} } // biar dropdown BRAND TAP IN di KALENDER KONTEN ikut nambah
  const cur = readOpsi_()[l] || [];
  if (cur.some(function (x) { return x.toLowerCase() === v.toLowerCase(); })) return;
  opsiSheet_().appendRow([l, v]);
}
// ===================== BANK KONTEN (tab BANK) =====================
// konten tanpa tanggal, disimpan buat dipakai kapan aja. Kalau dijadwalin → jadi konten di KALENDER KONTEN & hilang dari bank.
const BANK_HEAD = ['ID', 'JUDUL', 'JENIS', 'SEGMEN', 'AKUN', 'BRAND TAP IN', 'CATATAN', 'ASAL', 'DIBUAT', 'DIUBAH'];
function bankSheet_() {
  let s = sh_('BANK');
  if (!s) {
    s = SpreadsheetApp.getActive().insertSheet('BANK');
    s.getRange(1, 1, 1, BANK_HEAD.length).setValues([BANK_HEAD]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function readBank_() {
  const s = bankSheet_(), last = s.getLastRow(), out = [];
  if (last < 2) return out;
  const v = s.getRange(2, 1, last - 1, BANK_HEAD.length).getDisplayValues();
  for (let i = 0; i < v.length; i++) {
    const title = String(v[i][1] || '').trim();
    if (!title) continue;
    let id = String(v[i][0] || '').trim();
    if (!id) { id = newId_(); s.getRange(i + 2, 1).setValue(id); }
    out.push({ id: id, title: title, type: typeId_(v[i][2]), segmen: String(v[i][3] || '').trim(), akun: String(v[i][4] || '').trim(),
      brand: String(v[i][5] || '').trim(), catatan: String(v[i][6] || ''), from: String(v[i][7] || '').trim(), created: String(v[i][8] || '').trim() });
  }
  return out;
}
function bankRow_(id) {
  const s = bankSheet_(), last = s.getLastRow();
  if (last < 2) return -1;
  const ids = s.getRange(2, 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}
function bankUpsert_(b) {
  if (!b || !b.id || !b.title) throw new Error('bank konten butuh judul');
  const s = bankSheet_(), r = bankRow_(b.id), row = r > 0 ? r : s.getLastRow() + 1, now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  const created = r > 0 ? s.getRange(r, 9).getDisplayValue() : now;
  const rng = s.getRange(row, 1, 1, BANK_HEAD.length);
  rng.setNumberFormat('@');
  rng.setValues([[b.id, b.title, JENIS[b.type] || b.type || '', b.segmen || '', b.akun || '', b.brand || '', b.catatan || '', b.from || '', created || now, now]]);
}
function bankDelete_(id) { const r = bankRow_(id); if (r > 0) bankSheet_().deleteRow(r); }
// dijadwalin: pindah ke arsip "BANK TERPAKAI" (riset/brief nggak masuk CATATAN KALENDER KONTEN, catatan itu buat otomasi upload)
const BANKU_HEAD = BANK_HEAD.concat(['DIJADWALIN', 'TANGGAL KONTEN', 'JUDUL KONTEN']);
function bankUsedSheet_() {
  let s = sh_('BANK TERPAKAI');
  if (!s) {
    s = SpreadsheetApp.getActive().insertSheet('BANK TERPAKAI');
    s.getRange(1, 1, 1, BANKU_HEAD.length).setValues([BANKU_HEAD]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}
function bankUse_(op) {
  const s = bankSheet_(), r = bankRow_(op.id), now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), b = op.bank || {};
  let row = r > 0 ? s.getRange(r, 1, 1, BANK_HEAD.length).getDisplayValues()[0]
    : [op.id, b.title || '', JENIS[b.type] || b.type || '', b.segmen || '', b.akun || '', b.brand || '', b.catatan || '', b.from || '', now, now];
  const u = bankUsedSheet_(), ur = u.getLastRow() + 1, rng = u.getRange(ur, 1, 1, BANKU_HEAD.length);
  rng.setNumberFormat('@');
  rng.setValues([row.concat([now, op.date || '', op.title || row[1]])]);
  if (r > 0) s.deleteRow(r);
}
function readBankUsed_() {
  const s = sh_('BANK TERPAKAI');
  if (!s || s.getLastRow() < 2) return [];
  const v = s.getRange(2, 1, s.getLastRow() - 1, BANKU_HEAD.length).getDisplayValues(), out = [];
  v.forEach(function (x) { if (x[1] || x[12]) out.push({ id: x[0], title: x[1], catatan: x[6], from: x[7], used: x[10], date: x[11], cTitle: x[12] }); });
  return out;
}

// ===================== IDE KONTEN SP, NSP & FP (DASHBOARD KULTURAMA, diisi agent riset) =====================
// dibaca semua kolomnya; dari web cuma kolom Status & kolom Brief (minta brief, SP & FP) yang diubah
// src 'sp'  -> tab "Ide Konten SP" (dulu "Ide Konten SMO"; nama lama masih dikenali)
// src 'nsp' -> tab "Ide Konten NSP"
// src 'fp'  -> tab "Ide Konten FP" (femme_police; punya agent Content Writer FP & Doc brief sendiri)
const IDE_TABS = { sp: ['Ide Konten SP', 'Ide Konten SMO'], nsp: ['Ide Konten NSP'], fp: ['Ide Konten FP'] };
const IDE_SHEET_ID = { sp: 2044770584, fp: 276736585 }; // ID tab tetap walau namanya diganti
const BRIEF_SRC = ['sp', 'fp']; // sumber ide yang punya agent Content Writer
const PEL_SRC = { sp: 'SP', nsp: 'NSP', fp: 'FP' }; // label SUMBER di tab Pelajaran Agent
const SMO_COLS = { date: 'Tanggal', title: 'Judul', angle: 'Angle/POV', trigger: 'Pemicu', research: 'Bahan Riset', flow: 'Alur Isi',
  hook: 'Hook', format: 'Format', urgency: 'Urgensi', sources: 'Sumber', status: 'Status', brief: 'Brief', rev: 'Revisi' };
function smoSrc_(src) { return src === 'nsp' || src === 'fp' ? src : 'sp'; }
function smoSheet_(src) {
  const names = IDE_TABS[smoSrc_(src)], b = kkBook_();
  for (let i = 0; i < names.length; i++) { const sh = b.getSheetByName(names[i]); if (sh) return sh; }
  const id = IDE_SHEET_ID[smoSrc_(src)];
  if (id != null) { const byId = b.getSheets().filter(function (x) { return x.getSheetId() === id; })[0]; if (byId) return byId; }
  throw new Error('Tab "' + names[0] + '" nggak ketemu');
}
function smoCols_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return String(h).trim().toLowerCase(); }), c = {};
  Object.keys(SMO_COLS).forEach(function (k) { const i = head.indexOf(SMO_COLS[k].toLowerCase()); if (i >= 0) c[k] = i; });
  if (c.title == null) throw new Error('Kolom Judul nggak ketemu di ' + sh.getName());
  return c;
}
function smoRead_(from, src) {
  const sh = smoSheet_(src), last = sh.getLastRow();
  if (last < 2) return [];
  const c = smoCols_(sh), w = sh.getLastColumn(), rng = sh.getRange(2, 1, last - 1, w), raw = rng.getValues(), disp = rng.getDisplayValues(), out = [];
  for (let i = 0; i < raw.length; i++) {
    const g = function (k) { return c[k] == null ? '' : String(disp[i][c[k]] || '').trim(); };
    const title = g('title');
    if (!title) continue;
    const d = c.date == null ? '' : (date_(raw[i][c.date]) || date_(disp[i][c.date]) || '');
    if (from && d && d < from) continue;
    out.push({ row: i + 2, date: d, title: title, angle: g('angle'), trigger: g('trigger'), research: g('research'), flow: g('flow'),
      hook: g('hook'), format: g('format'), urgency: g('urgency'), sources: g('sources'), status: g('status'), brief: g('brief'), rev: g('rev') });
  }
  return out;
}
function smoStatus_(ref, status, src) {
  const sh = smoSheet_(src), c = smoCols_(sh);
  if (c.status == null) throw new Error('Kolom Status nggak ada di ' + sh.getName());
  let row = Number(ref && ref.row) || 0;
  const titleAt = function (r) { return r >= 2 && r <= sh.getLastRow() ? String(sh.getRange(r, c.title + 1).getDisplayValue()).trim() : ''; };
  if (titleAt(row) !== String(ref.title || '').trim()) { // baris geser → cari pakai judul
    const t = sh.getRange(2, c.title + 1, Math.max(1, sh.getLastRow() - 1), 1).getDisplayValues();
    row = 0;
    for (let i = 0; i < t.length; i++) if (String(t[i][0]).trim() === String(ref.title || '').trim()) { row = i + 2; break; }
    if (!row) throw new Error('notfound: ide "' + ref.title + '" nggak ketemu di ' + sh.getName());
  }
  sh.getRange(row, c.status + 1).setValue(status || '');
}
// cari baris ide (pakai nomor baris, kalau geser cari pakai judul)
function smoFindRow_(sh, c, ref) {
  let row = Number(ref && ref.row) || 0;
  const titleAt = function (r) { return r >= 2 && r <= sh.getLastRow() ? String(sh.getRange(r, c.title + 1).getDisplayValue()).trim() : ''; };
  if (titleAt(row) === String(ref.title || '').trim()) return row;
  const t = sh.getRange(2, c.title + 1, Math.max(1, sh.getLastRow() - 1), 1).getDisplayValues();
  for (let i = 0; i < t.length; i++) if (String(t[i][0]).trim() === String(ref.title || '').trim()) return i + 2;
  throw new Error('notfound: ide "' + ref.title + '" nggak ketemu di ' + sh.getName());
}
// Edit ide (koreksi arah konten) dari web: cuma buat sumber yang punya Content Writer (SP & FP).
// Judul cuma boleh diganti selama brief belum ada (brief dicocokkan ke Doc pakai judul).
// Perubahan dicatet di Pelajaran Agent (JENIS "ide") biar agent riset belajar arah yang dimau.
const SMO_EDITABLE = ['title', 'angle', 'hook', 'flow', 'format', 'research'];
const SMO_LBL = { title: 'Judul', angle: 'Angle/POV', hook: 'Hook', flow: 'Alur Isi', format: 'Format', research: 'Bahan Riset' };
function smoEdit_(ref, fields, src) {
  src = smoSrc_(src);
  if (BRIEF_SRC.indexOf(src) < 0) return { ok: false, error: 'edit_src: edit ide belum bisa buat sumber ini' };
  const sh = smoSheet_(src), c = smoCols_(sh), row = smoFindRow_(sh, c, ref), w = sh.getLastColumn();
  const cur = sh.getRange(row, 1, 1, w).getDisplayValues()[0], get = function (k) { return c[k] == null ? '' : String(cur[c[k]] || ''); };
  const briefNow = String(get('brief')).trim(), changed = [];
  fields = fields || {};
  SMO_EDITABLE.forEach(function (k) {
    if (fields[k] == null || c[k] == null) return;
    const nv = String(fields[k]).replace(/\r/g, '').trim(), ov = get(k).trim();
    if (nv === ov) return;
    if (k === 'title') {
      if (!nv) throw new Error('edit_title: judul nggak boleh kosong');
      if (briefNow) throw new Error('edit_title: judul nggak bisa diganti karena brief-nya udah ada/diminta');
    }
    const cell = sh.getRange(row, c[k] + 1);
    if (cell.getDataValidation()) cell.clearDataValidations();
    cell.setValue(nv);
    changed.push({ k: k, o: ov, n: nv });
  });
  if (!changed.length) return { ok: true, changed: 0 };
  SpreadsheetApp.flush();
  const title = changed.filter(function (x) { return x.k === 'title'; }).map(function (x) { return x.n; })[0] || get('title');
  const detail = changed.map(function (x) { return SMO_LBL[x.k] + ':\n- ' + (x.o || '(kosong)').slice(0, 400) + '\n+ ' + x.n.slice(0, 400); }).join('\n');
  try { pelLog_(PEL_SRC[src], 'ide', title, detail); } catch (e) {}
  return { ok: true, changed: changed.length, title: title, summary: changed.map(function (x) { return SMO_LBL[x.k]; }).join(', ') };
}
// Tombol "Bikin Brief": nulis "antri · <waktu>" di kolom Brief (L). Agent Content Writer cuma ngerjain baris "antri".
// Isi kolom Brief: kosong | antri · … | proses · … | gagal: … | <link brief> — <Jenis>
function smoBrief_(ref, action, src) {
  src = smoSrc_(src);
  if (BRIEF_SRC.indexOf(src) < 0) return { ok: false, error: 'brief_status: Bikin Brief cuma buat Ide SP & Ide FP' };
  const sh = smoSheet_(src), c = smoCols_(sh);
  if (c.brief == null) { // kolom Brief belum ada → pakai kolom L
    const L = 12, h = String(sh.getRange(1, L).getDisplayValue()).trim();
    if (h && h.toLowerCase() !== 'brief') throw new Error('Kolom L di ' + sh.getName() + ' udah dipakai "' + h + '", bukan Brief');
    sh.getRange(1, L).setValue('Brief');
    c.brief = L - 1;
  }
  const row = smoFindRow_(sh, c, ref), cell = sh.getRange(row, c.brief + 1), cur = String(cell.getDisplayValue()).trim(), low = cur.toLowerCase();
  // kolom Brief teks bebas: buang dropdown yang kebawa dari kolom Status (bikin setValue ditolak)
  if (cell.getDataValidation()) cell.clearDataValidations();
  if (action === 'cancel') {
    if (low.indexOf('antri') !== 0) return { ok: false, error: 'brief_busy: brief udah nggak di antrian (' + (cur || 'kosong') + ')', brief: cur };
    cell.setValue('');
    return { ok: true, brief: '' };
  }
  const st = c.status == null ? '' : String(sh.getRange(row, c.status + 1).getDisplayValue()).trim().toLowerCase();
  if (st !== 'ambil' && st !== 'bank') return { ok: false, error: 'brief_status: ide ini belum "ambil"', brief: cur };
  if (cur && low.indexOf('gagal') !== 0) return { ok: true, brief: cur }; // udah antri / proses / jadi → biarin
  const v = 'antri · ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  cell.setValue(v);
  SpreadsheetApp.flush();
  return { ok: true, brief: v, fired: briefFire_(src) };
}
// Colek agent Content Writer biar langsung jalan (API trigger routine Claude).
// Isi di Properti skrip: BRIEF_FIRE_URL (URL /fire dari claude.ai/code/routines) + BRIEF_FIRE_TOKEN (sk-ant-oat01-…).
// Agent FP: BRIEF_FIRE_URL_FP (+ BRIEF_FIRE_TOKEN_FP; kalau kosong pakai BRIEF_FIRE_TOKEN).
// Kosong → nggak nyolek, antrian tetap dikerjain run per jam. Jeda min. 2 menit antar colekan biar agent nggak jalan dobel.
function briefFire_(src) {
  const fp = smoSrc_(src) === 'fp', pr = PropertiesService.getScriptProperties();
  const url = pr.getProperty(fp ? 'BRIEF_FIRE_URL_FP' : 'BRIEF_FIRE_URL'), tok = (fp && pr.getProperty('BRIEF_FIRE_TOKEN_FP')) || pr.getProperty('BRIEF_FIRE_TOKEN');
  if (!url || !tok) return 'off';
  const cache = CacheService.getScriptCache(), ck = fp ? 'brief_fired_fp' : 'brief_fired';
  if (cache.get(ck)) return 'soon'; // baru dicolek, run itu bakal ngambil antrian ini juga / run berikutnya
  try {
    const r = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + tok, 'anthropic-version': '2023-06-01', 'anthropic-beta': 'experimental-cc-routine-2026-04-01' },
      payload: JSON.stringify({ text: 'Dipicu tombol Bikin Brief di TIMELINE KALENDER.' }) });
    const code = r.getResponseCode();
    if (code >= 200 && code < 300) { cache.put(ck, '1', 120); return 'ok'; }
    console.warn('briefFire_ HTTP ' + code + ': ' + r.getContentText().slice(0, 300));
    return 'err';
  } catch (err) { console.warn('briefFire_: ' + err); return 'err'; }
}

// ===================== BRIEF (Doc "Brief Konten SMO Shoe Police") =====================
// Brief = 1 Heading 2 (judul) + semua paragraf di bawahnya sampai heading 1/2 berikutnya, di tab sesuai link kolom Brief.
// Isinya dikirim/diterima sebagai teks: **tebal**, link polos / [teks](url). Edit dari app nulis ulang paragrafnya di Doc.
const BRIEF_DOC_ID = '108W3MvSjq_1ySIomf8xUeq5x7DrFxFLFymzRlDt_YHY';
const BRIEF_DOC_FP_ID = '1PTcngtRnlN-UBSZEkTukWPnwbO58hiP4JkdVvDpuJBI'; // "Brief Konten FP femme_police" (ID Doc diambil dari link di kolom Brief)
function izinDocs() { Logger.log('Doc brief: ' + DocumentApp.openById(BRIEF_DOC_ID).getName() + ' · ' + DocumentApp.openById(BRIEF_DOC_FP_ID).getName()); }
function brNorm_(t) { return String(t || '').toLowerCase().replace(/[“”"'‘’`]/g, '').replace(/[^a-z0-9]+/g, ' ').trim(); }
function briefCell_(ref, src) {
  src = smoSrc_(src);
  if (BRIEF_SRC.indexOf(src) < 0) throw new Error('brief_none: sumber ide ini nggak punya brief');
  const sh = smoSheet_(src), c = smoCols_(sh);
  if (c.brief == null) throw new Error('brief_none: kolom Brief belum ada');
  const row = smoFindRow_(sh, c, ref), v = String(sh.getRange(row, c.brief + 1).getDisplayValue()).trim();
  const m = v.match(/https?:\/\/docs\.google\.com\/document\/d\/([\w-]+)[^\s|]*/);
  if (!m) throw new Error('brief_none: ide ini belum punya brief (' + (v || 'kolom Brief kosong') + ')');
  const tab = (m[0].match(/[?&]tab=([\w.]+)/) || [])[1] || '';
  return { src: src, sh: sh, c: c, row: row, val: v, url: m[0], docId: m[1], tabId: tab, title: String(sh.getRange(row, c.title + 1).getDisplayValue()).trim() };
}
function briefBody_(docId, tabId) {
  const doc = DocumentApp.openById(docId);
  let tab = null;
  if (tabId && doc.getTab) { try { tab = doc.getTab(tabId); } catch (e) {} }
  if (tab) return { body: tab.asDocumentTab().getBody(), tabName: tab.getTitle() };
  return { body: doc.getBody(), tabName: '' };
}
function briefHead_(el) {
  const t = el.getType();
  if (t !== DocumentApp.ElementType.PARAGRAPH) return '';
  const h = el.asParagraph().getHeading(), H = DocumentApp.ParagraphHeading;
  return h === H.HEADING1 ? 'h1' : h === H.HEADING2 ? 'h2' : '';
}
// cari heading 2 brief: judul sama persis (abaikan huruf besar/kecil & tanda kutip) → judul mirip
function briefFind_(body, title) {
  const n = body.getNumChildren(), want = brNorm_(title), wt = want.split(' ');
  let best = -1, bestS = 0;
  for (let i = 0; i < n; i++) {
    const el = body.getChild(i);
    if (briefHead_(el) !== 'h2') continue;
    const h = brNorm_(el.asParagraph().getText());
    if (!h) continue;
    if (h === want) return i;
    const ht = h.split(' '), inter = wt.filter(function (w) { return ht.indexOf(w) >= 0; }).length, sc = inter / Math.max(wt.length, ht.length);
    if (sc > bestS) { bestS = sc; best = i; }
  }
  return bestS >= 0.6 ? best : -1;
}
function briefEnd_(body, start) {
  const n = body.getNumChildren();
  let i = start + 1;
  while (i < n && !briefHead_(body.getChild(i))) i++;
  return i; // indeks pertama SETELAH brief
}
function brRunsMd_(txt) { // Text → md (**tebal**, link)
  const s = txt.getText();
  if (!s) return '';
  const idx = txt.getTextAttributeIndices();
  let out = '';
  for (let k = 0; k < idx.length; k++) {
    const a = idx[k], b = k + 1 < idx.length ? idx[k + 1] : s.length, part = s.slice(a, b);
    if (!part) continue;
    const url = txt.getLinkUrl(a), bold = txt.isBold(a);
    if (url && part.trim()) out += (part.trim() === url ? part : '[' + part + '](' + url + ')');
    else if (bold && part.trim()) { const lead = part.match(/^\s*/)[0], trail = part.match(/\s*$/)[0]; out += lead + '**' + part.trim() + '**' + trail; }
    else out += part;
  }
  return out.replace(/\*\*\*\*/g, '');
}
function brElMd_(el) {
  const T = DocumentApp.ElementType, t = el.getType();
  if (t === T.PARAGRAPH || t === T.LIST_ITEM) return [brRunsMd_(el.editAsText())];
  if (t === T.TABLE) { const tb = el.asTable(), o = []; for (let r = 0; r < tb.getNumRows(); r++) { const row = tb.getRow(r), cells = []; for (let c = 0; c < row.getNumCells(); c++) cells.push(row.getCell(c).getText().replace(/\n/g, ' ')); o.push(cells.join(' | ')); } return o; }
  return [];
}
function brHash_(md) { return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, md, Utilities.Charset.UTF_8).map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); }).join('').slice(0, 16); }
function briefLocate_(ref, src) {
  const bc = briefCell_(ref, src), bb = briefBody_(bc.docId, bc.tabId), i = briefFind_(bb.body, bc.title);
  if (i < 0) throw new Error('brief_notfound: brief "' + bc.title + '" nggak ketemu di tab ' + (bb.tabName || 'Doc') + ' (judulnya mungkin diganti di Doc)');
  const end = briefEnd_(bb.body, i);
  let lines = [];
  for (let k = i + 1; k < end; k++) lines = lines.concat(brElMd_(bb.body.getChild(k)));
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  while (lines.length && !lines[0].trim()) lines.shift();
  const md = lines.join('\n');
  return { bc: bc, body: bb.body, tabName: bb.tabName, at: i, end: end, md: md, hash: brHash_(md), heading: bb.body.getChild(i).asParagraph().getText().trim() };
}
function briefRead_(ref, src) {
  const L = briefLocate_(ref, src);
  return { title: L.heading, tab: L.tabName, md: L.md, hash: L.hash, url: L.bc.url };
}
// satu baris md → teks polos + rentang tebal & link
function brParse_(line) {
  let text = '';
  const bold = [], link = [], re = /\*\*([^*]+?)\*\*|\[([^\]]+)\]\((https?:[^)\s]+)\)|(https?:\/\/[^\s)]+)/g;
  let last = 0, m;
  while ((m = re.exec(line))) {
    text += line.slice(last, m.index);
    const s0 = text.length;
    if (m[1] != null) { text += m[1]; bold.push([s0, text.length - 1]); }
    else if (m[2] != null) { text += m[2]; link.push([s0, text.length - 1, m[3]]); }
    else { text += m[4]; link.push([s0, text.length - 1, m[4]]); }
    last = re.lastIndex;
  }
  text += line.slice(last);
  return { text: text, bold: bold, link: link };
}
function briefSave_(ref, md, base, src) {
  md = String(md || '').replace(/\r/g, '');
  if (!md.trim()) throw new Error('brief_empty: brief nggak boleh kosong');
  const L = briefLocate_(ref, src);
  if (base && base !== L.hash) return { ok: false, error: 'brief_conflict: brief di Doc baru aja berubah', hash: L.hash };
  const body = L.body;
  if (L.end >= body.getNumChildren()) body.appendParagraph(''); // Doc wajib diakhiri paragraf
  for (let k = L.end - 1; k > L.at; k--) body.removeChild(body.getChild(k));
  const lines = md.split('\n');
  for (let k = 0; k < lines.length; k++) {
    const r = brParse_(lines[k]), para = body.insertParagraph(L.at + 1 + k, r.text);
    para.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    if (!r.text) continue;
    const t = para.editAsText();
    t.setBold(false); t.setLinkUrl(null);
    r.bold.forEach(function (b) { if (b[1] >= b[0]) t.setBold(b[0], b[1], true); });
    r.link.forEach(function (b) { if (b[1] >= b[0]) t.setLinkUrl(b[0], b[1], b[2]); });
  }
  body.insertParagraph(L.at + 1 + lines.length, ''); // jarak ke brief berikutnya
  try { pelLog_(PEL_SRC[L.bc.src], 'edit', L.bc.title, brDiff_(L.md, md)); } catch (e) {}
  let h = '';
  try { h = briefLocate_(ref, src).hash; } catch (e) {}
  return { ok: true, hash: h };
}
function brDiff_(a, b) {
  const A = a.split('\n').map(function (x) { return x.trim(); }).filter(Boolean), B = b.split('\n').map(function (x) { return x.trim(); }).filter(Boolean);
  const del = A.filter(function (x) { return B.indexOf(x) < 0; }), add = B.filter(function (x) { return A.indexOf(x) < 0; });
  const cut = function (arr, mark) { let o = '', n = 0; for (let i = 0; i < arr.length && o.length < 1400; i++) { o += mark + ' ' + arr[i].slice(0, 300) + '\n'; n++; } return o + (arr.length > n ? '(' + (arr.length - n) + ' baris lagi)\n' : ''); };
  return ('DIHAPUS:\n' + (del.length ? cut(del, '-') : '(nggak ada)\n') + 'DITAMBAH / DIGANTI JADI:\n' + (add.length ? cut(add, '+') : '(nggak ada)\n')).trim();
}
// Minta revisi: kolom Brief jadi "revisi · <waktu> | <link> — <Jenis>", catatannya di kolom Revisi, terus colek agent
function briefRevise_(ref, note, src) {
  note = String(note || '').trim();
  if (!note) throw new Error('brief_note: catatan revisi kosong');
  const bc = briefCell_(ref, src), low = bc.val.toLowerCase();
  if (/^(proses revisi|revisi)/.test(low)) return { ok: false, error: 'brief_busy: revisi sebelumnya belum beres', brief: bc.val };
  const linkPart = bc.val.slice(bc.val.indexOf('http'));
  let revCol = bc.c.rev;
  if (revCol == null) { revCol = bc.sh.getLastColumn(); bc.sh.getRange(1, revCol + 1).setValue('Revisi'); }
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), v = 'revisi · ' + now + ' | ' + linkPart;
  const rc = bc.sh.getRange(bc.row, revCol + 1);
  if (rc.getDataValidation()) rc.clearDataValidations();
  rc.setValue(note);
  bc.sh.getRange(bc.row, bc.c.brief + 1).setValue(v);
  SpreadsheetApp.flush();
  bumpRev_();
  try { pelLog_(PEL_SRC[bc.src], 'revisi', bc.title, note); } catch (e) {}
  return { ok: true, brief: v, fired: briefFire_(bc.src) };
}

// ===================== PELAJARAN AGENT =====================
// Tab "Pelajaran Agent" (DASHBOARD): log feedback mentah (edit brief, minta revisi, alasan skip). STATUS baru → dipakai (diisi agent).
// Tab "Aturan Agent": aturan singkat hasil rangkuman agent. AKTIF ya | dihapus (dihapus dari app = agent nggak boleh bikin lagi).
const PEL_TAB = 'Pelajaran Agent', PEL_HEAD = ['TANGGAL', 'SUMBER', 'JENIS', 'JUDUL', 'DETAIL', 'STATUS'];
const ATR_TAB = 'Aturan Agent', ATR_HEAD = ['ID', 'UNTUK', 'ATURAN', 'ALASAN', 'DIBUAT', 'AKTIF'];
function pelTab_(name, head) {
  const b = kkBook_();
  let sh = b.getSheetByName(name);
  if (!sh) { sh = b.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function pelLog_(src, kind, title, detail) {
  kind = String(kind || '').toLowerCase();
  if (['edit', 'revisi', 'skip', 'ambil', 'ide'].indexOf(kind) < 0) throw new Error('jenis feedback nggak dikenal');
  const sh = pelTab_(PEL_TAB, PEL_HEAD);
  sh.appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), src, kind, String(title || '').slice(0, 200), String(detail || '').slice(0, 3000), 'baru']);
}
function pelRead_() {
  const b = kkBook_(), P = b.getSheetByName(PEL_TAB), A = b.getSheetByName(ATR_TAB), log = [], rules = [];
  if (P && P.getLastRow() > 1) {
    const n = P.getLastRow() - 1, from = Math.max(2, P.getLastRow() - 59), v = P.getRange(from, 1, P.getLastRow() - from + 1, 6).getDisplayValues();
    v.forEach(function (r) { log.push({ at: r[0], src: r[1], kind: r[2], title: r[3], detail: r[4], status: r[5] }); });
    log.reverse();
    log.total = n;
  }
  if (A && A.getLastRow() > 1) A.getRange(2, 1, A.getLastRow() - 1, 6).getDisplayValues().forEach(function (r, i) {
    if (String(r[2]).trim()) rules.push({ row: i + 2, id: r[0], untuk: r[1], text: r[2], why: r[3], at: r[4], aktif: String(r[5]).trim().toLowerCase() || 'ya' });
  });
  return { log: log, rules: rules, pending: log.filter(function (x) { return x.status === 'baru'; }).length };
}
function ruleAdd_(untuk, text) {
  text = String(text || '').trim();
  if (!text) throw new Error('aturan kosong');
  const ok = ['brief', 'riset SP', 'riset NSP', 'brief FP', 'riset FP', 'semua'];
  if (ok.indexOf(untuk) < 0) untuk = 'semua';
  const sh = pelTab_(ATR_TAB, ATR_HEAD), id = 'u' + Date.now().toString(36);
  sh.appendRow([id, untuk, text.slice(0, 300), 'Ditambah manual dari TIMELINE KALENDER', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), 'ya']);
  return { ok: true, id: id };
}
function ruleDel_(id, text) {
  const sh = kkBook_().getSheetByName(ATR_TAB);
  if (!sh || sh.getLastRow() < 2) return;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues();
  for (let i = 0; i < v.length; i++) {
    if ((id && String(v[i][0]).trim() === String(id)) || (!id && String(v[i][2]).trim() === String(text || '').trim())) { sh.getRange(i + 2, 6).setValue('dihapus'); return; }
  }
}

// ===================== BRAND DEAL (tab SOW BRAND DEAL: 1 baris = 1 deal) =====================
const SOW_TAB = 'SOW BRAND DEAL';
const SOW_MAP = { bulan: 'BULAN', brand: 'BRAND', artikel: 'ARTIKEL', jml: 'JUMLAH SEPATU', barang: 'STATUS BARANG',
  car: 'CAROUSEL', stCar: 'STATUS CAROUSEL', vp: 'VIDEO PENDEK', stVp: 'STATUS VIDEO PENDEK', vpj: 'VIDEO PANJANG', stVpj: 'STATUS VIDEO PANJANG',
  aff: 'AFFILIATE', stAff: 'STATUS AFFILIATE', ket: 'KETERANGAN', id: 'ID' };
function sowSheet_() { const sh = kkBook_().getSheetByName(SOW_TAB); if (!sh) throw new Error('sow_tab: tab "' + SOW_TAB + '" nggak ketemu'); return sh; }
function sowCols_(sh) {
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (x) { return String(x).trim().toUpperCase(); }), c = {};
  Object.keys(SOW_MAP).forEach(function (k) { const i = h.indexOf(SOW_MAP[k]); if (i >= 0) c[k] = i; });
  if (c.id == null || c.brand == null) throw new Error('sow_head: header ID / BRAND nggak ada di tab ' + SOW_TAB);
  c._w = h.length; return c;
}
function sowMonth_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM');
  const t = String(v || '').trim(), m = t.match(/^(\d{4})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  const B = ['januari', 'februari', 'maret', 'april', 'mei', 'juni', 'juli', 'agustus', 'september', 'oktober', 'november', 'desember'], w = t.toLowerCase().split(/\s+/);
  const bi = B.findIndex(function (b) { return w[0] && b.indexOf(w[0].slice(0, 3)) === 0; });
  return bi >= 0 && /^\d{4}$/.test(w[1] || '') ? w[1] + '-' + ('0' + (bi + 1)).slice(-2) : '';
}
function sowRead_() {
  const sh = sowSheet_(), c = sowCols_(sh), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const raw = sh.getRange(2, 1, n, c._w).getValues(), disp = sh.getRange(2, 1, n, c._w).getDisplayValues(), out = [];
  for (let i = 0; i < n; i++) {
    const g = function (k) { return c[k] == null ? '' : String(disp[i][c[k]]).trim(); };
    if (!g('brand') && !g('id')) continue;
    const o = { row: i + 2 };
    Object.keys(SOW_MAP).forEach(function (k) { o[k] = g(k); });
    o.bulan = c.bulan == null ? '' : sowMonth_(raw[i][c.bulan]);
    out.push(o);
  }
  return out;
}
function sowRow_(sh, c, id) {
  const n = sh.getLastRow() - 1; if (n < 1 || !id) return -1;
  const v = sh.getRange(2, c.id + 1, n, 1).getDisplayValues();
  for (let i = 0; i < n; i++) if (String(v[i][0]).trim() === String(id)) return i + 2;
  return -1;
}
function sowVal_(k, v) {
  v = v == null ? '' : String(v).trim();
  if (k === 'bulan') { const m = sowMonth_(v); return m ? new Date(Number(m.slice(0, 4)), Number(m.slice(5, 7)) - 1, 1) : ''; }
  if (['jml', 'car', 'vp', 'vpj', 'aff'].indexOf(k) >= 0) return v === '' ? '' : (isNaN(Number(v)) ? v : Number(v));
  return v;
}
function sowAdd_(rec) {
  rec = rec || {};
  if (!String(rec.brand || '').trim()) throw new Error('sow_brand: brand wajib diisi');
  const sh = sowSheet_(), c = sowCols_(sh), id = String(rec.id || ('sow' + Date.now().toString(36)));
  if (sowRow_(sh, c, id) > 0) return { ok: true, id: id, dup: true };
  const row = new Array(c._w).fill('');
  Object.keys(SOW_MAP).forEach(function (k) { if (c[k] != null && k !== 'id') row[c[k]] = sowVal_(k, rec[k]); });
  row[c.id] = id;
  sh.appendRow(row);
  return { ok: true, id: id };
}
function sowUpdate_(id, patch) {
  const sh = sowSheet_(), c = sowCols_(sh), r = sowRow_(sh, c, id);
  if (r < 0) throw new Error('sow_notfound: deal ' + id + ' nggak ketemu (mungkin dihapus di Sheet)');
  Object.keys(patch || {}).forEach(function (k) { if (c[k] != null && k !== 'id') sh.getRange(r, c[k] + 1).setValue(sowVal_(k, patch[k])); });
}
function sowDelete_(id) {
  const sh = sowSheet_(), c = sowCols_(sh), r = sowRow_(sh, c, id);
  if (r > 0) sh.deleteRow(r);
}

// ===================== UPLOAD SEKARANG (status Gagal → coba upload lagi) =====================
// Balikin STATUS <platform> ke "Acc", terus colek GitHub Actions di repo automation (ops-pipeline-core)
// bawa nomor baris + judul. Automation-nya ngeloloskan baris itu walau tanggalnya udah lewat.
// Isi di Properti skrip: GH_DISPATCH_TOKEN = token GitHub (fine-grained, repo ops-pipeline-core, izin Contents: Read and write)
//                        GH_DISPATCH_REPO  = oqaja/ops-pipeline-core (boleh kosong, default itu)
// CATATAN nggak disentuh — itu tempat automation nulis hasil/errornya.
const UPLOAD_PF = { YT: 'content-acc-yt', TT: 'content-acc-tt', IG: { SP: 'content-acc-ig-sp', NSP: 'content-acc-ig-nsp', FP: 'content-acc-ig-fp' } };
function uploadNow_(ref, pf) {
  pf = String(pf || '').toUpperCase();
  if (!UPLOAD_PF[pf]) throw new Error('upload_now: ' + (pf === 'FB' ? 'Facebook belum diurus automation, upload manual dulu' : 'platform tidak dikenal'));
  const pr = PropertiesService.getScriptProperties(), tok = pr.getProperty('GH_DISPATCH_TOKEN'), repo = pr.getProperty('GH_DISPATCH_REPO') || 'oqaja/ops-pipeline-core';
  if (!tok) throw new Error('upload_now: GH_DISPATCH_TOKEN belum diisi di Properti skrip Apps Script');
  const sh = kkSheet_(), c = kkCols_(sh), row = kkFind_(sh, c, ref);
  if (row === -2) throw new Error('upload_now: ada lebih dari satu baris "' + ref.title + '" di tanggal ' + ref.date);
  if (row < 0) throw new Error('upload_now: baris "' + ref.title + '" nggak ketemu, tarik ulang data dulu');
  if (!c['st' + pf]) throw new Error('upload_now: kolom STATUS ' + pf + ' nggak ada');
  const v = sh.getRange(row, 1, 1, c._width).getDisplayValues()[0], g = function (k) { return c[k] ? String(v[c[k] - 1]).trim() : ''; };
  const st = g('st' + pf), title = g('title');
  if (st.toLowerCase() !== 'gagal') return { ok: false, error: 'upload_now: status ' + pf + ' sekarang "' + (st || 'kosong') + '", bukan Gagal', st: st };
  if (g('id' + pf)) throw new Error('upload_now: POST ID ' + pf + ' udah keisi (' + g('id' + pf) + '), kosongin dulu kalau mau upload ulang');
  if (pf !== 'IG' && g('production') !== '✅') throw new Error('upload_now: PRODUCTION belum ✅');
  let ev = UPLOAD_PF[pf];
  if (pf === 'IG') { ev = ev[g('akun').toUpperCase()]; if (!ev) throw new Error('upload_now: akun "' + g('akun') + '" nggak punya automation IG'); }
  sh.getRange(row, c['st' + pf]).setValue('Acc');
  SpreadsheetApp.flush();
  bumpRev_();
  let fired = 'err', info = '';
  try {
    const r = UrlFetchApp.fetch('https://api.github.com/repos/' + repo + '/dispatches', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + tok, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
      payload: JSON.stringify({ event_type: ev, client_payload: { row: String(row), title: title, pf: pf, from: 'timeline-kalender' } }) });
    const code = r.getResponseCode();
    if (code === 204) fired = 'ok'; else { info = 'HTTP ' + code; console.warn('uploadNow_ ' + info + ': ' + r.getContentText().slice(0, 300)); }
  } catch (err) { info = String(err).slice(0, 120); console.warn('uploadNow_: ' + err); }
  if (fired !== 'ok') { // GitHub nggak kesolek → balikin ke Gagal biar tombolnya bisa diklik lagi (tanpa colekan, IG tanggal lewat nggak bakal keambil)
    sh.getRange(row, c['st' + pf]).setValue('Gagal');
    SpreadsheetApp.flush();
    bumpRev_();
    return { ok: false, error: 'upload_now: gagal nyolek GitHub (' + (info || 'error') + '), cek GH_DISPATCH_TOKEN. Status tetap Gagal.' };
  }
  return { ok: true, st: 'Acc', row: row, fired: fired };
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
  const rng = sh.getRange(2, 1, last - 1, c._width), raw = rng.getValues(), disp = rng.getDisplayValues(), bankOf = kkBankRows_(sh);
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
    if (bankOf[it.row]) it.bankId = bankOf[it.row];
    items.push(it);
  }
  return items;
}
// penanda tersembunyi di baris KALENDER KONTEN (developer metadata, nggak kelihatan & nggak ngisi sel apa pun):
// konten ini asalnya dari bank konten mana. Ikut pindah kalau baris digeser/di-sort, tetap nyambung walau judul/tanggal diganti.
const KK_BANK_KEY = 'tl_bank';
function kkMarkBank_(sh, row, bankId) {
  if (!bankId) return;
  try { sh.getRange(row + ':' + row).addDeveloperMetadata(KK_BANK_KEY, String(bankId)); } catch (e) {}
}
function kkBankRows_(sh) {
  const out = {};
  try {
    sh.createDeveloperMetadataFinder().withKey(KK_BANK_KEY).find().forEach(function (md) {
      const loc = md.getLocation();
      if (String(loc.getLocationType()) === 'ROW') out[loc.getRow().getRow()] = md.getValue();
    });
  } catch (e) {}
  return out;
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
// tambah pilihan ke dropdown kolom KALENDER KONTEN (aturan validasinya di-copy, cuma daftarnya yang ditambah)
function kkListAdd_(key, value) {
  const sh = kkSheet_(), c = kkCols_(sh);
  if (!c[key]) return false;
  const rule = sh.getRange(2, c[key]).getDataValidation();
  if (!rule || String(rule.getCriteriaType()) !== 'VALUE_IN_LIST') return false;
  const args = rule.getCriteriaValues(), list = (args[0] || []).map(String);
  if (list.some(function (x) { return x.toLowerCase() === String(value).toLowerCase(); })) return false;
  const nr = rule.copy().withCriteria(rule.getCriteriaType(), [list.concat([value]), args[1] !== false]).build();
  sh.getRange(2, c[key], sh.getMaxRows() - 1, 1).setDataValidation(nr);
  return true;
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
  if (changes && changes.brand) { try { kkListAdd_('brand', String(changes.brand)); } catch (e) {} }
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
  if (keys.indexOf('date') >= 0) kkDateFmt_(sh, row, c);
  return row;
}
// tanggal ditulis "yyyy-MM-dd" → Sheets ikut pakai format itu; samain ke format kolom TANGGAL (01-Nov-2026)
const KK_DATE_FMT = 'dd"-"mmm"-"yyyy';
function kkDateFmt_(sh, row, c) { if (c.date) sh.getRange(row, c.date).setNumberFormat(KK_DATE_FMT); }
function kkAdd_(it) {
  if (!it || !it.title || !it.date) throw new Error('judul & tanggal wajib');
  const sh = kkSheet_(), c = kkCols_(sh);
  if (it.brand) { try { kkListAdd_('brand', String(it.brand)); } catch (e) {} }
  const vals = new Array(c._width).fill('');
  Object.keys(KK).forEach(function (k) { if (c[k] && k !== 'eventId') vals[c[k] - 1] = kkVal_(k, it[k]); });
  sh.appendRow(vals);
  const row = sh.getLastRow();
  kkDateFmt_(sh, row, c);
  kkMarkBank_(sh, row, it.bankId);
  return row;
}
function kkDelete_(ref) {
  const sh = kkSheet_(), c = kkCols_(sh), row = kkFind_(sh, c, ref);
  if (row === -2) throw new Error('dobel: ada lebih dari satu baris "' + ref.title + '" di tanggal ' + ref.date);
  if (row > 0) sh.deleteRow(row);
}

function doGet(e) {
  const t0 = Date.now(), p = (e && e.parameter) || {};
  if (p.update === '1' && typeof updateSekarang === 'function') { try { updateSekarang(); } catch (err) {} }
  const useCache = !p.cek && !p.nocache && p.update !== '1';
  const key = (p.smo === '1' ? 's' + smoSrc_(p.src) + '|' : p.ins === '1' ? 'i|' : 'g|') + dataRev_() + '|' + (p.from || '') + '|' + (p.to || '');
  if (useCache) {
    const hit = cacheGet_(key);
    if (hit) { const d = JSON.parse(hit); d.cached = true; d.ms = Date.now() - t0; return out_(d); }
  }
  if (p.brief === '1') { // isi 1 brief dari Doc (nggak di-cache: bisa diedit kapan aja)
    try { return out_({ ok: true, ver: SCRIPT_VER, brief: briefRead_({ row: p.row, title: p.title }, p.src) }); }
    catch (err) { return out_({ ok: false, ver: SCRIPT_VER, error: String(err) }); }
  }
  if (p.pel === '1') { // pelajaran & aturan agent
    try { return out_({ ok: true, ver: SCRIPT_VER, pel: pelRead_() }); }
    catch (err) { return out_({ ok: false, ver: SCRIPT_VER, error: String(err) }); }
  }
  if (p.smo === '1') { // tab Ide SMO
    const d = { ok: true, ver: SCRIPT_VER };
    try { d.smo = smoRead_(p.from || '', p.src); d.src = smoSrc_(p.src); } catch (err) { d.smo = null; d.smoError = String(err); }
    if (useCache && d.smo) { try { cachePut_(key, JSON.stringify(d), READ_TTL); } catch (err) {} }
    d.cached = false; d.ms = Date.now() - t0;
    return out_(d);
  }
  if (p.sow === '1') { // tab Brand Deal (SOW BRAND DEAL)
    const k2 = 'w|' + dataRev_();
    if (useCache) { const hit = cacheGet_(k2); if (hit) { const d = JSON.parse(hit); d.cached = true; return out_(d); } }
    const d = { ok: true, ver: SCRIPT_VER };
    try { d.sow = sowRead_(); } catch (err) { d.sow = null; d.sowError = String(err); }
    if (useCache && d.sow) { try { cachePut_(k2, JSON.stringify(d), READ_TTL); } catch (err) {} }
    return out_(d);
  }
  if (p.ins === '1') { // tab Insight: cuma data INSIGHT, terpisah biar tarikan kalender tetap ringan
    const d = { ok: true, ver: SCRIPT_VER };
    try { d.ins = insightRead_(p.from || '', p.to || ''); } catch (err) { d.ins = null; d.insError = String(err); }
    if (useCache && d.ins) { try { cachePut_(key, JSON.stringify(d), READ_TTL); } catch (err) {} }
    d.cached = false; d.ms = Date.now() - t0;
    return out_(d);
  }
  const data = read_();
  data.items.forEach(function (it) { it.src = 'pv'; });
  data.ver = SCRIPT_VER;
  try { data.todos = readTodo_(); } catch (err) { data.todos = null; }
  try { data.bank = readBank_(); } catch (err) { data.bank = null; }
  try { data.bankUsed = readBankUsed_(); } catch (err) { data.bankUsed = null; }
  try { data.kk = kkRead_(p.from || '', p.to || ''); data.lists = kkLists_(); if (kkWarn_.length) data.kkWarn = kkWarn_; }
  catch (err) { data.kk = null; data.kkError = String(err); }
  data.range = { from: p.from || '', to: p.to || '' };
  data.loader = typeof LOADER_SRC !== 'undefined';
  if (p.cek === 'drive') data.drive = driveCek_();
  if (useCache && data.kk) { try { cachePut_(key, JSON.stringify(data), READ_TTL); } catch (err) {} }
  data.cached = false; data.ms = Date.now() - t0;
  return out_(data);
}

// ---- INSIGHT (DASHBOARD KULTURAMA): cuma dibaca, nggak pernah ditulis dari web
const INS_TAB = 'INSIGHT';
const INS_COLS = { date: 'Tanggal Upload', title: 'Judul', jenis: 'Jenis Konten', platform: 'Platform', akun: 'Akun', link: 'Link Video',
  views: 'Views', subs: 'Subscriber Gained', likes: 'Likes', comments: 'Comments', shares: 'Shares', saved: 'Saved',
  updated: 'Terakhir Update', postId: 'POST ID', brand: 'BRAND TAP IN' };
const INS_PF = { youtube: 'YT', instagram: 'IG', tiktok: 'TT', facebook: 'FB' };
function insNum_(raw, disp) {
  if (typeof raw === 'number') return raw;
  const s = String(disp || '').trim();
  if (!s || s === '-') return null;
  const n = Number(s.replace(/\./g, '').replace(',', '.'));
  return isNaN(n) ? null : n;
}
function insDate_(raw, disp, tz) { // -> ['yyyy-MM-dd', 'HH:mm']
  if (raw instanceof Date) return [Utilities.formatDate(raw, tz, 'yyyy-MM-dd'), Utilities.formatDate(raw, tz, 'HH:mm')];
  const m = String(disp || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return ['', ''];
  return [m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2), m[4] ? ('0' + m[4]).slice(-2) + ':' + m[5] : ''];
}
function insightRead_(from, to) {
  const book = kkBook_(), sh = book.getSheetByName(INS_TAB);
  if (!sh) throw new Error('Tab INSIGHT nggak ketemu');
  const last = sh.getLastRow(), width = sh.getLastColumn();
  if (last < 2) return [];
  const head = sh.getRange(1, 1, 1, width).getDisplayValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
  const c = {};
  Object.keys(INS_COLS).forEach(function (k) { const i = head.indexOf(INS_COLS[k].toLowerCase()); if (i >= 0) c[k] = i; });
  if (c.date == null || c.title == null) throw new Error('Kolom Tanggal Upload / Judul nggak ketemu di INSIGHT');
  const rng = sh.getRange(2, 1, last - 1, width), raw = rng.getValues(), disp = rng.getDisplayValues(), tz = book.getSpreadsheetTimeZone(), out = [];
  const g = function (i, k) { return c[k] == null ? '' : String(disp[i][c[k]] || '').trim(); };
  const n = function (i, k) { return c[k] == null ? null : insNum_(raw[i][c[k]], disp[i][c[k]]); };
  for (let i = 0; i < raw.length; i++) {
    const dt = insDate_(raw[i][c.date], disp[i][c.date], tz);
    if (!dt[0] || (from && dt[0] < from) || (to && dt[0] > to)) continue;
    const pf = g(i, 'platform');
    out.push({ d: dt[0], tm: dt[1], ti: g(i, 'title'), j: g(i, 'jenis'), pf: INS_PF[pf.toLowerCase()] || pf, ak: g(i, 'akun'), url: g(i, 'link'),
      v: n(i, 'views'), sg: n(i, 'subs'), l: n(i, 'likes'), c: n(i, 'comments'), sh: n(i, 'shares'), sv: n(i, 'saved'),
      up: g(i, 'updated'), id: g(i, 'postId'), br: g(i, 'brand') });
  }
  return out;
}

// ---- Simpanan sementara hasil baca (biar narik data cepat). Dianggap basi tiap ada yang nyimpen lewat app.
// Edit langsung di Google Sheet kebaca paling lambat READ_TTL detik.
const READ_TTL = 60;
function dataRev_() {
  const c = CacheService.getScriptCache();
  let r = c.get('data_rev');
  if (!r) { r = String(Date.now()); c.put('data_rev', r, 21600); }
  return r;
}
function bumpRev_() { CacheService.getScriptCache().put('data_rev', String(Date.now()) + Math.random().toString(36).slice(2, 6), 21600); }
// CacheService maks 100KB per nilai -> dipecah
function cachePut_(key, str, ttl) {
  const size = 90000, n = Math.ceil(str.length / size), o = {};
  if (n > 20) return;
  for (let i = 0; i < n; i++) o[key + '#' + i] = str.slice(i * size, (i + 1) * size);
  o[key + '#n'] = String(n);
  CacheService.getScriptCache().putAll(o, ttl);
}
function cacheGet_(key) {
  const c = CacheService.getScriptCache(), n = Number(c.get(key + '#n') || 0);
  if (!n) return '';
  const ks = []; for (let i = 0; i < n; i++) ks.push(key + '#' + i);
  const got = c.getAll(ks); let s = '';
  for (let i = 0; i < n; i++) { if (got[ks[i]] == null) return ''; s += got[ks[i]]; }
  return s;
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
  const t0 = Date.now();
  let op0 = {};
  try { op0 = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  try { pinCheck_(op0.pin); } catch (err) { return out_({ ok: false, error: String(err) }); }
  if (op0.op === 'checkPin') return out_({ ok: true });
  if (op0.op === 'invoiceSave') { try { return out_(invoiceSave_(op0)); } catch (err) { return out_({ ok: false, error: String(err) }); } }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (op0.op === 'batch') {
      // banyak perubahan sekaligus dalam 1 request; berhenti di perubahan pertama yang gagal
      const ops = op0.ops || [], results = [];
      for (let i = 0; i < ops.length; i++) {
        let r;
        try { r = runOp_(ops[i]); } catch (err) { r = { ok: false, error: String(err) }; }
        if (!r.ok) { if (results.length) bumpRev_(); return out_({ ok: true, done: results.length, results: results, fail: r, ms: Date.now() - t0 }); }
        results.push(r);
      }
      SpreadsheetApp.flush(); bumpRev_();
      return out_({ ok: true, done: results.length, results: results, ms: Date.now() - t0 });
    }
    const r = runOp_(op0);
    SpreadsheetApp.flush(); if (r.ok) bumpRev_();
    r.ms = Date.now() - t0;
    return out_(r);
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
function runOp_(op) {
  if (op.op === 'upsert') upsert_(op.item);
  else if (op.op === 'delete') remove_(op.id);
  else if (op.op === 'color') color_(op.brand, op.color);
  else if (op.op === 'option') addOpsi_(op.list, op.value);
  else if (op.op === 'optionDel') delOpsi_(op.list, op.value);
  else if (op.op === 'todoUpsert') todoUpsert_(op.todo);
  else if (op.op === 'todoDelete') todoDelete_(op.id);
  else if (op.op === 'bankUpsert') bankUpsert_(op.bank);
  else if (op.op === 'bankDelete') bankDelete_(op.id);
  else if (op.op === 'bankUse') bankUse_(op);
  else if (op.op === 'smoStatus') smoStatus_(op.ref, op.status, op.src);
  else if (op.op === 'sowAdd') return sowAdd_(op.rec);
  else if (op.op === 'sowUpdate') sowUpdate_(op.id, op.patch);
  else if (op.op === 'sowDelete') sowDelete_(op.id);
  else if (op.op === 'smoBrief') return smoBrief_(op.ref, op.action, op.src);
  else if (op.op === 'smoEdit') return smoEdit_(op.ref, op.fields, op.src);
  else if (op.op === 'uploadNow') return uploadNow_(op.ref, op.pf);
  else if (op.op === 'briefSave') return briefSave_(op.ref, op.md, op.base, op.src);
  else if (op.op === 'briefRevise') return briefRevise_(op.ref, op.note, op.src);
  else if (op.op === 'agentFb') pelLog_(PEL_SRC[smoSrc_(op.src)], op.kind, op.title, op.detail);
  else if (op.op === 'ruleAdd') return ruleAdd_(op.untuk, op.text);
  else if (op.op === 'ruleDel') ruleDel_(op.id, op.text);
  else if (op.op === 'kkAdd') {
    if (op.dedupe) { // dikirim ulang setelah app ditutup: jangan bikin baris dobel kalau ternyata udah masuk
      const sh = kkSheet_(), row = kkFind_(sh, kkCols_(sh), { title: op.item && op.item.title, date: op.item && op.item.date });
      if (row > 0) return { ok: true, row: row };
    }
    return { ok: true, row: kkAdd_(op.item) };
  }
  else if (op.op === 'kkUpdate') {
    const res = kkUpdate_(op.ref, op.changes, op.prev, op.force);
    return res && res.conflict ? { ok: false, conflict: res.conflict, error: 'conflict' } : { ok: true, row: res };
  }
  else if (op.op === 'pvUpdate') {
    const res = pvUpdate_(op.id, op.changes, op.cols, op.prev, op.force);
    return res && res.conflict ? { ok: false, conflict: res.conflict, error: 'conflict' } : { ok: true };
  }
  else if (op.op === 'kkDelete') kkDelete_(op.ref);
  else if (op.op === 'bulk') { // pindahan data lama dari browser, cuma kalau sheet masih kosong
    if (read_().items.length === 0) {
      (op.items || []).forEach(upsert_);
      Object.keys(op.brandColors || {}).forEach(function (b) { color_(b, op.brandColors[b]); });
    }
  }
  else return { ok: false, error: 'op tidak dikenal: ' + op.op };
  return { ok: true };
}

// ---- INVOICE: salinan PDF dari invoice.html disimpan ke folder Drive SEBELUM boleh di-download
// Folder bisa diganti lewat Properti skrip: INVOICE_FOLDER_ID = <id folder>
const INVOICE_FOLDER_ID = '1pUNqU1aVozohDB_cRLr7EmmBSt0BjSrF';
function invoiceFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty('INVOICE_FOLDER_ID') || INVOICE_FOLDER_ID;
  try { return DriveApp.getFolderById(id); }
  catch (err) {
    // bedain: izin Drive belum ada vs folder nggak bisa diakses akun pemilik script
    let me = '';
    try { me = DriveApp.getRootFolder().getOwner().getEmail(); } catch (e2) {
      throw new Error('drive_scope: Apps Script belum dapet izin Drive (' + String(e2) + ')');
    }
    throw new Error('drive_folder: folder invoice (' + id + ') nggak bisa dibuka akun ' + me + '. Share folder itu ke akun ini sebagai Editor. (' + String(err) + ')');
  }
}
function driveCek_() {
  try { const f = invoiceFolder_(); return { ok: true, folder: f.getName() }; }
  catch (err) { return { ok: false, error: String(err) }; }
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
