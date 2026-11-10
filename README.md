# Timeline Produksi Konten — Shoe Police

Kalender produksi konten. Klik tanggal untuk menambah konten, klik konten untuk edit/hapus, dan drag & drop konten untuk memindahkan tanggal.

- Field: semua kolom KALENDER KONTEN (DASHBOARD KULTURAMA) — Judul, Jenis, Brand Tap In, Akun, Segmen, Jam Up & Status & Post ID per platform, Producer, Designer, Editor, Talent, Writing/Shooting/Production, Catatan, Event ID
- Warna per brand tap in (bisa diganti), atau warna per jenis konten
- Satu file `index.html`, tanpa build

Catatan: versi GitHub Pages menyimpan data di browser masing-masing (localStorage), jadi data tidak otomatis sinkron antar perangkat.

## Sinkron ke Google Sheet

Data disimpan di Google Sheet "TIMELINE KALENDER — Shoe Police" (tab KONTEN dan BRAND), jadi bisa dibuka dari device mana pun.

Setup sekali:
1. Buka Sheet-nya → **Ekstensi → Apps Script**.
2. Hapus isi `Code.gs`, tempel isi file [`apps-script/Code.gs`](apps-script/Code.gs), simpan.
3. **Deploy → Deployment baru** → jenis **Aplikasi web**, Jalankan sebagai **Saya**, Akses **Siapa saja** → Deploy, izinkan akses.
4. Salin URL yang berakhiran `/exec`, tempel di tombol ⚙ di kalender (atau isi `SHEET_API` di `index.html`).

Konten juga bisa ditambah langsung di Sheet: isi TANGGAL (yyyy-mm-dd), JUDUL KONTEN, JENIS KONTEN, BRAND TAP IN, kosongkan ID (terisi otomatis).

## Konten vs Preview (sejak 2 Nov 2026)

- **Konten** dibaca & ditulis langsung ke DASHBOARD KULTURAMA V1 → KALENDER KONTEN (struktur kolom di sana tidak diubah; baris dikenali lewat EVENT ID atau judul + tanggal, dan yang ditulis hanya sel yang berubah).
- **Preview** cuma disimpan di tab KONTEN pada Sheet TIMELINE KALENDER — tidak pernah masuk KALENDER KONTEN. Tampil bergaris dengan label PV.
- Kalender menarik KALENDER KONTEN per rentang bulan (bulan sebelum s/d sesudah yang dibuka).
- Akun yang men-deploy Apps Script harus punya akses edit ke DASHBOARD KULTURAMA V1.

## Auto-nyambung + PIN edit

- URL Web App sudah ditanam di `index.html` (`SHEET_API`), jadi kalender langsung nyambung di browser/device mana pun.
- Semua perubahan (tambah/edit/hapus/warna/opsi) wajib PIN. PIN **tidak** ada di kode — atur di Apps Script → ⚙ Setelan project → Properti skrip → `EDIT_PIN`.
- Tanpa `EDIT_PIN`, script menolak semua perubahan (kalender cuma bisa dilihat). Salah PIN 10x → dikunci 15 menit.
- PIN diminta sekali per browser, lalu diingat. Hapus lewat ⚙ → "Lupain PIN".

## Invoice Maker (`invoice.html`)

Bikin invoice PT Kulturama Virtua Solusindo / Shoe Police sesuai template, langsung jadi PDF A4. Buka dari tab **Invoice ↗** di kalender.

- No. Ref otomatis `YYYYMM` + urutan 3 digit (naik tiap download PDF), bisa diedit manual.
- Klien bisa disimpan & dipilih ulang. Item bisa ditambah / dihapus / diurutkan; harga boleh kosong untuk item satu paket.
- Diskon & PPN 11%/12% opsional (muncul baris Subtotal/Diskon/PPN).
- Item banyak → otomatis lanjut ke halaman berikutnya; Grand Total, tanda tangan & bank di halaman terakhir.
- Nama marketing, rekening, NPWP, kontak diingat di browser. Semua data invoice disimpan di browser (localStorage), tidak ke Sheet.
- Gambar template (kop, stempel, tanda tangan, dll) ada di folder `invoice-assets/`.

### Wajib tersimpan di Drive dulu

Tombol **Simpan & Download PDF** ngirim salinan PDF ke folder Drive invoice lewat Apps Script. File baru di-download **setelah** Drive bilang berhasil. Kalau gagal, nggak ada download (ada tombol Coba lagi). PIN sama dengan PIN edit kalender.

- Folder: `INVOICE_FOLDER_ID` di `apps-script/Code.gs` (bisa ditimpa lewat Properti skrip `INVOICE_FOLDER_ID`).
- Nama file sama (download ulang setelah diedit) → versi lama dipindah ke Sampah Drive, yang baru disimpan.

Setup sekali setelah update script:
1. Buka Apps Script, ganti isi `Code.gs` dengan isi [`apps-script/Code.gs`](apps-script/Code.gs), simpan.
2. Pilih fungsi **izinDrive** di toolbar, klik **Run**, lalu izinkan akses Google Drive.
3. **Deploy → Kelola deployment → ✏ Edit → Versi: Versi baru → Deploy** (URL `/exec` tetap sama).

## Loader Apps Script (update otomatis dari GitHub)

Isi Apps Script cukup [`apps-script/Loader.gs`](apps-script/Loader.gs) (ditempel sekali). Loader ngambil `apps-script/Code.gs` terbaru dari repo ini dan ngejalaninnya, jadi tiap perubahan di `Code.gs` yang di-push ke `main` kepakai sendiri dalam ±5–10 menit — **tanpa copas & tanpa Deploy ulang**.

Setup sekali:
1. Apps Script → hapus semua isi `Code.gs`, tempel isi `Loader.gs`, simpan.
2. Pilih fungsi **izinSemua** → **Run** → izinkan semua akses.
3. **Deploy → Kelola deployment → ✏ Edit → Versi baru → Deploy** (URL `/exec` tetap).

- Butuh update saat itu juga? Run **updateSekarang**.
- GitHub down → dipakai salinan terakhir (disimpan di Properti skrip `LOADER_SRC_*`; jangan dihapus).
- Fitur baru yang butuh izin Google jenis baru (Gmail, Calendar, dll) tetap perlu Run **izinSemua** sekali lagi.
