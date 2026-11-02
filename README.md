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
