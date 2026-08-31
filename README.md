# Contributing Factor Coding Review — website

Isi folder ini:

| File | Fungsi |
|------|--------|
| `index.html` | Halaman web (tabel + 6 grafik). Sudah lengkap, tidak perlu diedit tangan. |
| `data.json` | Data 376 faktor. **File inilah yang diganti setiap ada data baru.** |
| `build-data.ps1` | Skrip yang membuat ulang `data.json` dari workbook Excel. |
| `executive-summary.html` | Ringkasan eksekutif dwibahasa (tombol Bahasa / English) untuk anggota IPI. Ditaut dari halaman utama ("Executive summary →"). Angkanya diketik manual — perbarui sendiri bila data berubah. |
| `.nojekyll` | Penanda kecil supaya GitHub Pages menyajikan file apa adanya. |
| `README.md` | Dokumen ini. |

Halaman membaca `data.json` saat dibuka. Ada juga salinan cadangan data di dalam
`index.html` supaya file tetap berfungsi bila dibuka langsung tanpa server
(mis. dikirim lewat email). `build-data.ps1` memperbarui keduanya sekaligus.

---

## Bagian 1 — Publikasi ke GitHub Pages (sekali saja)

Hasil akhir: alamat publik seperti `https://<user>.github.io/factor-coding-review/`
yang bisa dibuka siapa saja yang punya link.

1. Buat akun di <https://github.com> (gratis) bila belum punya.
2. Klik **New repository**.
   - **Repository name:** `factor-coding-review` (bebas, tanpa spasi)
   - **Public** (wajib — Pages gratis hanya untuk repo publik)
   - centang **Add a README file** boleh, boleh juga tidak
   - klik **Create repository**
3. Di halaman repo, klik **Add file → Upload files**, lalu tarik masuk
   **semua isi folder ini**: `index.html`, `data.json`, `build-data.ps1`,
   `executive-summary.html`, `.nojekyll`, `README.md`. Klik **Commit changes**.
4. Buka **Settings → Pages** (menu kiri).
   - **Source:** *Deploy from a branch*
   - **Branch:** `main`, folder `/ (root)` → **Save**
5. Tunggu 1–2 menit, muat ulang halaman Settings → Pages. Akan muncul:
   *"Your site is live at https://…"* — itu alamat yang dibagikan ke anggota IPI.

> Catatan privasi: repo publik berarti `data.json` bisa diunduh siapa saja dan
> halaman bisa terindeks mesin pencari. Data ini berasal dari laporan akhir KNKT
> yang memang sudah dipublikasikan. Bila suatu saat perlu dibatasi hanya untuk
> anggota, pindahkan hosting ke Cloudflare Pages + Cloudflare Access.

---

## Bagian 2 — Update saat ada data accident baru

### Langkah 1 — tambah baris di Excel

Buka `Document\Investigation Database - Contributing Factors Update 2026.xlsx`
dan tambahkan satu baris **per contributing factor** (satu kecelakaan bisa
menghasilkan beberapa baris):

| Kolom | Isi |
|-------|-----|
| A | Tanggal peristiwa |
| B | Tahun peristiwa |
| E | Operator |
| G | Tipe pesawat |
| I | Occurrence type (pakai istilah CICTT, mis. `Runway Excursion (RE)`) |
| K | Teks contributing factor |
| L | Identification (Human / Technical / Environment / Facility) |
| **N** | **Domain** |
| **O** | **Discipline** |
| **P** | **Element** |

Kolom **N–P** dikode mengikuti taksonomi *Air Traffic Causal & Contributory
Factors* — aturan ringkasnya ada di bagian **"How these codes were assigned"**
pada halaman web. Kalau ragu, kirim teks faktornya untuk dibantu dikode.

Simpan file Excel.

### Langkah 2 — buat ulang data.json

Di folder ini, klik kanan `build-data.ps1` → **Run with PowerShell**.
(atau di jendela PowerShell: `cd` ke folder ini lalu `.\build-data.ps1`)

Skrip akan:
- membaca workbook,
- menulis ulang `data.json`,
- memperbarui salinan cadangan di `index.html`,
- menampilkan ringkasan (jumlah record, sebaran domain, rentang tahun) untuk dicek.

Butuh Microsoft Excel terpasang di komputer. Tutup file workbook dulu sebelum menjalankan.

### Langkah 3 — unggah ke GitHub

Di repo GitHub: **Add file → Upload files**, tarik masuk `data.json` **dan**
`index.html` (yang baru), **Commit changes**.

GitHub Pages otomatis men-deploy ulang dalam ~1 menit. Muat ulang halaman —
tabel, grafik, dan semua hitungan menyesuaikan sendiri.

---

## Catatan

- **Baris untuk satu kecelakaan sebaiknya berurutan di workbook.** Bila baris faktor
  dari satu kecelakaan terpisah, halaman menampilkannya sebagai dua pita warna
  (hitungan "event" tetap sekali). Dekatkan baris-barisnya di Excel untuk menyatukan pita.
- **Font** diambil dari Google Fonts saat online; offline otomatis memakai font
  bawaan sistem.
- Halaman menyimpan pilihan mode terang/gelap di browser masing-masing pengunjung.
