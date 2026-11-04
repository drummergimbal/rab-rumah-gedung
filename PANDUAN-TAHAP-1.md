# RAB Rumah & Gedung — Panduan Tahap 1 (by DG)

Versi 1.0.0 · Tahap 1: RAB + AHSP/HSD + Master + Import + Profil Perusahaan + Penawaran + PDF ke Drive.

---

## Isi paket

```
RAB-Rumah-Gedung-Tahap1/
├─ pwa/                  ← semua file ini di-upload ke GitHub / Cloudflare Pages
│   ├─ index.html, style.css, app.js, calc.js, sw.js, manifest.json, favicon.ico
│   ├─ icons/  logo/     ← aset yang sudah disediakan
│   └─ vendor/           ← library PDF & Excel (disimpan lokal agar bisa offline)
├─ apps-script/
│   ├─ Code.gs           ← backend (disalin ke Apps Script)
│   └─ Calc.gs           ← mesin hitung (isinya SAMA dengan pwa/calc.js)
└─ PANDUAN-TAHAP-1.md    ← file ini
```

---

## Langkah 1 — Buat Google Sheet & pasang Apps Script

1. Buka **sheets.google.com**, lalu buat spreadsheet kosong.
2. Ganti namanya menjadi **DATABASE RAB RUMAH & GEDUNG**.
3. Buka menu **Ekstensi → Apps Script**. Tab baru akan terbuka.
4. Di panel kiri ada file **Code.gs**. Hapus semua isinya, lalu tempel seluruh isi `apps-script/Code.gs`.
5. Klik **+** (Tambah file) → **Skrip**, lalu beri nama **Calc**. Hapus isinya, lalu tempel seluruh isi `apps-script/Calc.gs`.
6. Klik ikon **💾 Simpan**.
7. Kembali ke tab Google Sheet, lalu **muat ulang halaman** (F5). Setelah beberapa detik, menu **RAB Rumah & Gedung** muncul di sebelah kanan menu Bantuan.
8. Klik **RAB Rumah & Gedung → 1. Setup Awal**.
9. Google akan meminta izin:
   - **Lanjutkan** → pilih akun kamu.
   - Kalau muncul "Google hasn't verified this app", klik **Advanced / Lanjutan** → **Go to … (unsafe)**. Ini normal karena script buatan kamu sendiri.
   - Klik **Allow / Izinkan**.
10. Klik **1. Setup Awal** lagi kalau prosesnya berhenti setelah pemberian izin. Tunggu sampai muncul pesan **"Setup selesai ✅"** (±30–60 detik).

Setelah setup:
- 30 sheet sudah dibuat dengan tampilan rapi (header navy, baris selang-seling, format Rupiah).
- Folder **RAB Rumah & Gedung** sudah ada di Google Drive kamu, berisi `00_MASTER` dan `99_Backup`. Spreadsheet juga otomatis dipindah ke folder ini.
- Backup mingguan sudah aktif (Minggu 23:00 WIB, menyimpan 8 backup terakhir).
- Ada data **CONTOH** berisi 5 AHSP + 12 harga dasar untuk uji coba. **Ini bukan data resmi.**

## Langkah 2 — Deploy Web App (API)

1. Di Apps Script, klik **Deploy → New deployment** (Terapkan → Deployment baru).
2. Klik ikon ⚙️ di sebelah "Select type", lalu pilih **Web app**.
3. Isi:
   - Description: `RAB v1`
   - Execute as: **Me (email kamu)**
   - Who has access: **Anyone**
4. Klik **Deploy**, lalu salin **Web app URL** (berakhiran `/exec`).
5. Kembali ke Google Sheet, lalu klik **RAB Rumah & Gedung → 2. Tampilkan Token API** dan salin tokennya.

> "Anyone" artinya URL bisa dipanggil, tapi setiap permintaan wajib membawa **token**. Tanpa token yang benar, server menolak. Jangan bagikan token. Kalau token bocor, pakai menu **Buat Token Baru**.

## Langkah 3 — Upload PWA ke Cloudflare Pages

Caranya sama dengan PWA kamu yang lain. Pilih salah satu.

**A. Lewat GitHub (disarankan, mudah update nanti)**
1. Di github.com, klik **New repository**, beri nama `rab-rumah-gedung`, lalu pilih Private atau Public.
2. Klik **Add file → Upload files**, lalu seret **isi folder `pwa/`** (bukan foldernya). Pastikan `index.html` ada di paling atas repo, sejajar dengan folder `icons/`, `logo/`, dan `vendor/`. Klik **Commit changes**.
3. Di dash.cloudflare.com, buka **Workers & Pages → Create → Pages → Connect to Git**, lalu pilih repo `rab-rumah-gedung`.
4. Isi:
   - Framework preset: **None**
   - Build command: *(kosongkan)*
   - Build output directory: `/`
5. Klik **Save and Deploy**. Kamu akan mendapat alamat seperti `rab-rumah-gedung.pages.dev`.

**B. Upload langsung (tanpa GitHub)**
Buka **Workers & Pages → Create → Pages → Upload assets**, lalu seret folder `pwa/` dan klik **Deploy**.

## Langkah 4 — Hubungkan aplikasi ke Sheet

1. Buka alamat `….pages.dev`.
2. Pilih menu **Master → Pengaturan**.
3. Tempel **URL Web App** dan **Token API**, lalu klik **Simpan & tes koneksi**.
4. Kalau muncul "Terhubung ke RAB Rumah & Gedung v1.0.0", indikator di kanan atas akan menjadi **Tersinkron** (titik hijau).

## Langkah 5 — Install di HP

- **iPhone (Safari):** buka alamatnya → tombol **Bagikan** → **Tambah ke Layar Utama**.
- **Android (Chrome):** buka alamatnya → menu **⋮** → **Instal aplikasi** / **Tambahkan ke layar utama**.

Data cukup diisi di satu perangkat. Perangkat lain akan ikut sinkron setelah URL & token diisi di perangkat tersebut.

---

## Langkah 6 — Uji: rumah 1 lantai 6×10 m

1. Buat profil perusahaan: **Master → Profil Perusahaan → Profil baru**. Isi nama, alamat, dan kontak, lalu upload logo dan klik Simpan.
2. Buat proyek: **RAB → Proyek Baru**.
   - Nama: Rumah Uji 6×10
   - Luas bangunan: 60
   - Wilayah: Kabupaten Bogor, tahun 2026
   - Template: **Rumah 1 lantai**
3. Di tab **Item RAB**, klik **Cocokkan AHSP otomatis**. Saran yang kuat sudah dicentang. Klik **Terapkan**.
4. Isi volume tiga item ini dengan mengetuk item → kolom Volume → Simpan:

| Item | Volume |
|---|---|
| Galian tanah pondasi | 15,36 m³ — atau pakai 🧮 kalkulator: P 32 × L 0,6 × T 0,8 |
| Urugan pasir bawah pondasi | 1,92 m³ |
| Plesteran 1PC:4PP tebal 15 mm | 240 m² |

### Hitung manual sebagai bukti rumus

Pakai data **CONTOH** dengan overhead 10%. Rumusnya: **HSP = E + F**, dengan **E = Σ(koefisien × harga dasar)** dan **F = 10% × E**.

**1) Galian tanah (CTH-01), per m³**

| Komponen | Koef | Harga dasar | Jumlah |
|---|---|---|---|
| Pekerja | 0,75 OH | 120.000 | 90.000 |
| Mandor | 0,025 OH | 180.000 | 4.500 |
| **E** | | | **94.500** |
| **F** (10%) | | | 9.450 |
| **HSP** | | | **103.950** |

Jumlah: 15,36 × 103.950 = **Rp 1.596.672**

**2) Urugan pasir (CTH-02), per m³**

| Komponen | Koef | Harga dasar | Jumlah |
|---|---|---|---|
| Pasir urug | 1,2 m³ | 250.000 | 300.000 |
| Pekerja | 0,3 OH | 120.000 | 36.000 |
| Mandor | 0,01 OH | 180.000 | 1.800 |
| **E** | | | **337.800** |
| **F** (10%) | | | 33.780 |
| **HSP** | | | **371.580** |

Jumlah: 1,92 × 371.580 = 713.433,6 → **Rp 713.434**

**3) Plesteran (CTH-04), per m²**

| Komponen | Koef | Harga dasar | Jumlah |
|---|---|---|---|
| Semen | 6,24 kg | 1.500 | 9.360 |
| Pasir pasang | 0,024 m³ | 300.000 | 7.200 |
| Pekerja | 0,3 OH | 120.000 | 36.000 |
| Tukang batu | 0,15 OH | 150.000 | 22.500 |
| Kepala tukang | 0,015 OH | 170.000 | 2.550 |
| Mandor | 0,015 OH | 180.000 | 2.700 |
| **E** | | | **80.310** |
| **F** (10%) | | | 8.031 |
| **HSP** | | | **88.341** |

Jumlah: 240 × 88.341 = **Rp 21.201.840**

Angka di aplikasi (tab Item / Rekap) harus sama persis dengan angka di atas. Ketiga item ini sudah saya uji dan hasilnya cocok.

5. Lanjutkan uji fitur lain:
   - Tab **Rekap**: cek subtotal, SMKK terpisah, biaya tak terduga, PPN, terbilang, dan harga/m². Coba **PDF RAB** dan **Excel**.
   - Tab **Penawaran**: buat penawaran → atur markup/pembulatan → **Preview** → **Cetak & simpan ke Drive**. PDF otomatis masuk ke folder Drive `PRJ001 - Rumah Uji 6×10/02_Penawaran`.
   - Buka Google Sheet. Data proyek dan item muncul di sheet **PROYEK** dan **RAB_ITEM**. Kalau kamu mengubah satu angka langsung di Sheet, perubahan itu ikut muncul di aplikasi setelah sinkron (± 1 menit, atau ketuk indikator sinkron).

---

## Mengganti data CONTOH dengan AHSP resmi

1. **Master → Import → Unduh template** untuk melihat format kolom:
   `KODE_AHSP | URAIAN | SATUAN | DIVISI | JENIS_KOMPONEN | KODE_HSD | NAMA_KOMPONEN | SATUAN_KOMPONEN | KOEFISIEN | SUMBER`
   Satu baris berisi satu komponen.
2. Isi template dari lampiran Permen PUPR 8/2023 atau SE terbaru. Bagian ini bisa saya bantu ekstrak per divisi.
3. Klik **Import AHSP**, lalu periksa pratinjau (AHSP baru, koefisien berubah, dan HSD yang perlu harga). Beri nama versi, lalu import. Arsip CSV-nya otomatis tersimpan di `00_MASTER`.
4. Klik **Import HSD** untuk harga dari HSPK/SSH Pemda atau survei toko. Petakan kolomnya, lalu cocokkan baris yang belum terpetakan.
5. Cek sampel per divisi. Kalau sudah benar, klik **Tandai terverifikasi**.
6. Buka **Master → Pengaturan → Hapus semua data CONTOH**.

---

## Kalau nanti `Code.gs` diperbarui

Tempel kode baru → Simpan → **Deploy → Manage deployments** → ✏️ edit → Version: **New version** → Deploy. URL tetap sama, jadi aplikasi tidak perlu diubah.

Kalau `calc.js` berubah, salin juga isinya ke **Calc.gs** supaya hasil hitung di HP dan server tetap identik.

---

## Asumsi & batasan Tahap 1

**Fakta yang dipakai**
- Metode mengacu pada Permen PUPR 8/2023:
  - HSP = (Upah + Bahan + Alat) + biaya umum & keuntungan.
  - SMKK dihitung sebagai divisi terpisah (baris B di rekap).
  - Koefisien diperbarui lewat SE Dirjen Bina Konstruksi.

**Asumsi (bisa diubah di Master → Pengaturan atau per proyek)**
- Biaya umum & keuntungan default 10%.
- PPN default 11%. Tarif efektif bisa berubah, jadi sesuaikan bila aturan berubah.
- Biaya tak terduga 5%. Ini hanya ada di RAB internal dan tidak dibebankan ke penawaran.
- Pembulatan total RAB ke atas per Rp 1.000.
- Pembulatan penawaran ke atas: harga satuan per Rp 100, total per Rp 1.000.
- Harga dasar dicari dengan urutan berikut, dan setiap fallback diberi tanda peringatan:
  1. wilayah & tahun proyek;
  2. tahun terdekat di wilayah yang sama;
  3. wilayah lain.
- Kontinjensi tidak ikut dihitung dalam "margin internal" penawaran. Margin dibandingkan terhadap RAB sebelum biaya tak terduga & PPN.

**Batasan**
- Data AHSP resmi belum terisi. Yang ada hanya 5 AHSP **CONTOH** dengan koefisien ilustrasi. Jangan dipakai untuk RAB sungguhan.
- Template RAB hanya mengisi uraian & satuan. Pencocokan ke AHSP memakai kata kunci, sehingga hasilnya perlu dicek, terutama yang bertanda "saran lemah".
- Sinkron otomatis berjalan tiap ±1 menit saat aplikasi terbuka, ditambah setiap kali ada perubahan.
  - Kalau satu baris diubah di dua tempat bersamaan, yang tersimpan adalah versi terakhir, dan konfliknya dicatat di sheet **LOG**.
- PDF dibuat di HP/laptop (bisa offline), lalu otomatis diunggah ke Drive saat online.
- Peran Owner/Pengawas/Viewer, tenaga kerja (Tahap 2), jadwal/Kurva S (Tahap 3), serta realisasi/opname/CCO (Tahap 4) belum aktif. Sheet untuk modul-modul itu sudah disiapkan supaya struktur tidak perlu diubah nanti.

**Penyesuaian kecil dari prompt**
- Ada sheet tambahan **REVISI_ITEM**, karena snapshot revisi bisa melebihi batas karakter satu sel.
- Library PDF/Excel disimpan lokal di folder `vendor/`, bukan dari CDN, supaya cetak PDF tetap bisa saat offline.
- Ada file tambahan **Calc.gs** (mesin hitung) supaya validasi di server sama persis dengan hitungan di HP.

**Yang sudah saya uji di lingkungan uji**
Pengujian memakai backend tiruan Apps Script dan Chrome headless, bukan Google asli. Yang diuji:
- setup sheet;
- sinkron dua arah dan offline → online;
- template;
- cocokkan AHSP;
- kalkulator volume;
- tambah item opsi A dan C;
- AHSP custom;
- rekap;
- PDF RAB & penawaran (dicek tidak memuat kata AHSP/koefisien/margin);
- Excel;
- revisi;
- import AHSP & HSD;
- service worker offline.

Yang **belum** teruji adalah izin Google, Drive, dan deployment yang sebenarnya, karena itu terjadi di akun kamu. Kalau ada error saat Setup atau Deploy, kirim screenshot pesan errornya.
