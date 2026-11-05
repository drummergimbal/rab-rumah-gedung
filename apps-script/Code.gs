/* ============================================================
 * RAB Rumah & Gedung — Backend Google Apps Script (by DG)
 * Versi 1.2.0 (Tahap 1–3: RAB, Tenaga Kerja, Jadwal & Kurva S)
 *
 * File di project Apps Script:
 *   Code.gs  → file ini
 *   Calc.gs  → salin isi calc.js (mesin hitung, sama dengan PWA)
 *
 * Menu "RAB Rumah & Gedung" muncul di Google Sheet setelah file disimpan
 * dan Sheet dibuka ulang. Jalankan "1. Setup Awal" sekali saja.
 * ============================================================ */

var APP = { name: 'RAB Rumah & Gedung', version: '1.3.0', tz: 'Asia/Jakarta' };

var THEME = {
  primary: '#1F4E79', primaryDark: '#0E2A45', primary50: '#EAF2FA',
  accent: '#F5A524', accent50: '#FFF4DD', bg: '#F7F9FC', surface: '#FFFFFF',
  zebra: '#F4F8FC', border: '#E3E8EF', text: '#1A2B3C', muted: '#5B6B7C',
  success: '#1E9E6A', successBg: '#E6F6EF', warning: '#E08A00', warningBg: '#FFF4DD',
  danger: '#D64545', dangerBg: '#FDECEC', info: '#2F7DD1', infoBg: '#EAF2FA',
  tab: { RAB: '#1F4E79', MASTER: '#5B6B7C', PELAKSANAAN: '#1E9E6A', PENAWARAN: '#F5A524', SISTEM: '#E3E8EF' }
};

/* Warna status — SAMA dengan badge di PWA */
var STATUS_COLORS = {
  'TERVERIFIKASI': 'success', 'OK': 'success', 'DITERIMA': 'success', 'SELESAI': 'success', 'AMAN': 'success', 'SESUAI': 'success', 'DISETUJUI': 'success',
  'PERLU CEK': 'warning', 'PERLU HARGA': 'warning', 'DRAFT': 'warning', 'NEGOSIASI': 'warning', 'WASPADA': 'warning', 'DIAJUKAN': 'warning',
  'DITOLAK': 'danger', 'KEDALUWARSA': 'danger', 'BATAL': 'danger', 'TERLAMBAT': 'danger', 'OVER BUDGET': 'danger',
  'TERKIRIM': 'info', 'BERJALAN': 'info', 'PENAWARAN': 'info', 'RENCANA': 'info', 'TERCAPAI': 'success',
  'DIBAYAR': 'success', 'TAMBAH': 'info', 'KURANG': 'warning', 'GANTI': 'info', 'TANPA ANGGARAN': 'warning', 'TIDAK': 'danger'
};

/* ---------- Skema sheet ----------
 * Format kolom "NAMA:tipe" → rp | num | coef | pct | date | ts | json | text(default)
 * Kolom ID selalu pertama; UPDATED_AT & DELETED selalu terakhir. */
var SCHEMA = {
  // Proyek & RAB
  PROYEK: { group: 'RAB', cols: ['KODE', 'NAMA', 'PEMILIK', 'LOKASI', 'JENIS_BANGUNAN', 'JUMLAH_LANTAI:num', 'LUAS_TANAH:num', 'LUAS_BANGUNAN:num', 'LUAS_PER_LANTAI', 'TAHUN_ANGGARAN:num', 'WILAYAH', 'VERSI_AHSP', 'PROFIL_ID', 'MODE', 'OVERHEAD_PCT:pct', 'PPN_PCT:pct', 'KONTINJENSI_PCT:pct', 'PEMBULATAN:num', 'TOTAL_RAB:rp', 'NILAI_KONTRAK:rp', 'STATUS', 'FOLDER_ID', 'CATATAN', 'TANGGAL_MULAI:date', 'ESKALASI_PCT:pct'] },
  DIVISI: { group: 'RAB', cols: ['KODE', 'NAMA', 'URUTAN:num', 'MODE', 'KETERANGAN'] },
  RAB_ITEM: { group: 'RAB', freezeCols: 6, cols: ['PROYEK_ID', 'PARENT_ID', 'TIPE_BARIS', 'URUTAN:num', 'KODE', 'URAIAN', 'SATUAN', 'VOLUME:num', 'JENIS_HARGA', 'KODE_AHSP', 'HARGA_PASAR:rp', 'SUMBER_HARGA', 'HARGA_SATUAN:rp', 'JUMLAH:rp', 'UPAH:rp', 'BAHAN:rp', 'ALAT:rp', 'BOBOT:pct', 'SPESIFIKASI', 'CATATAN'] },
  VOLUME_DETAIL: { group: 'RAB', cols: ['PROYEK_ID', 'ITEM_ID', 'URUTAN:num', 'KETERANGAN', 'PANJANG:num', 'LEBAR:num', 'TINGGI:num', 'JUMLAH:num', 'PENGURANG', 'HASIL:num'] },
  REVISI: { group: 'RAB', cols: ['PROYEK_ID', 'NAMA_REV', 'TANGGAL:date', 'TOTAL:rp', 'JUMLAH_ITEM:num', 'CATATAN'] },
  REVISI_ITEM: { group: 'RAB', cols: ['REVISI_ID', 'PROYEK_ID', 'ITEM_ID', 'TIPE_BARIS', 'DIVISI', 'KODE', 'URAIAN', 'SATUAN', 'VOLUME:num', 'HARGA_SATUAN:rp', 'JUMLAH:rp'] },
  // Master
  AHSP_HEADER: { group: 'MASTER', cols: ['KODE_AHSP', 'URAIAN', 'SATUAN', 'DIVISI', 'JENIS', 'KODE_REFERENSI', 'SUMBER', 'DASAR_PERHITUNGAN', 'TAG', 'SINONIM', 'SCOPE', 'PROYEK_ID', 'VERSI', 'STATUS', 'DIBUAT_OLEH', 'TANGGAL:date'] },
  AHSP_KOMPONEN: { group: 'MASTER', cols: ['KODE_AHSP', 'URUTAN:num', 'JENIS_KOMPONEN', 'KODE_HSD', 'NAMA_KOMPONEN', 'SATUAN', 'KOEFISIEN:coef'] },
  AHSP_VERSI: { group: 'MASTER', cols: ['VERSI', 'SUMBER', 'TANGGAL:date', 'JUMLAH_AHSP:num', 'JUMLAH_KOMPONEN:num', 'FILE_ARSIP', 'CATATAN'] },
  HSD_UPAH: { group: 'MASTER', cols: ['KODE_HSD', 'NAMA', 'SATUAN', 'WILAYAH', 'TAHUN:num', 'HARGA:rp', 'SUMBER', 'SINONIM', 'STATUS', 'TANGGAL_UPDATE:date'] },
  HSD_BAHAN: { group: 'MASTER', cols: ['KODE_HSD', 'NAMA', 'SATUAN', 'WILAYAH', 'TAHUN:num', 'HARGA:rp', 'SUMBER', 'SINONIM', 'STATUS', 'TANGGAL_UPDATE:date'] },
  HSD_ALAT: { group: 'MASTER', cols: ['KODE_HSD', 'NAMA', 'SATUAN', 'WILAYAH', 'TAHUN:num', 'HARGA:rp', 'SUMBER', 'SINONIM', 'STATUS', 'TANGGAL_UPDATE:date'] },
  WILAYAH: { group: 'MASTER', cols: ['NAMA', 'PROVINSI', 'AKTIF'] },
  PENGATURAN: { group: 'MASTER', cols: ['NILAI', 'KETERANGAN'] },
  PROFIL_USAHA: { group: 'MASTER', cols: ['NAMA', 'BENTUK', 'TAGLINE', 'LOGO_FILE_ID', 'ALAMAT', 'KOTA', 'PROVINSI', 'KODE_POS', 'TELEPON', 'WHATSAPP', 'EMAIL', 'WEBSITE', 'INSTAGRAM', 'NPWP', 'PKP', 'NIB', 'BANK', 'NO_REKENING', 'ATAS_NAMA', 'PJ_NAMA', 'PJ_JABATAN', 'TTD_FILE_ID', 'STEMPEL_FILE_ID', 'TAMPILKAN_TTD', 'WARNA_AKSEN', 'IS_DEFAULT'] },
  TEMPLATE_SYARAT: { group: 'MASTER', cols: ['JENIS', 'JUDUL', 'ISI', 'URUTAN:num'] },
  IMPORT_MAPPING: { group: 'MASTER', cols: ['NAMA', 'TIPE', 'MAPPING_JSON:json'] },
  // Tenaga & Jadwal (tahap 2–3, header disiapkan)
  TENAGA_REKAP: { group: 'PELAKSANAAN', cols: ['PROYEK_ID', 'JENIS_TENAGA', 'TOTAL_OH:num', 'UPAH_HARIAN:rp', 'TOTAL_UPAH:rp', 'ORANG_PUNCAK:num'] },
  BORONGAN_UPAH: { group: 'PELAKSANAAN', cols: ['PROYEK_ID', 'MODE', 'ITEM_ID', 'SATUAN', 'HARGA:rp', 'JUMLAH_TUKANG:num', 'HARI:num', 'TOTAL:rp', 'CATATAN'] },
  JADWAL: { group: 'PELAKSANAAN', cols: ['PROYEK_ID', 'ITEM_ID', 'DURASI:num', 'MULAI:date', 'SELESAI:date', 'PREDECESSOR', 'JENIS_RELASI', 'LAG:num', 'JUMLAH_TUKANG:num'] },
  TERMIN_UPAH: { group: 'PELAKSANAAN', cols: ['PROYEK_ID', 'URUTAN:num', 'LABEL', 'BASIS', 'PERSEN:pct', 'NOMINAL:rp', 'MINGGU_KE:num', 'STATUS', 'CATATAN'] },
  PROGRES: { group: 'PELAKSANAAN', cols: ['PROYEK_ID', 'MINGGU_KE:num', 'ITEM_ID', 'PERSEN_AKTUAL:pct', 'TANGGAL:date', 'CATATAN', 'FOTO'] },
  LIBUR: { group: 'SISTEM', cols: ['TANGGAL:date', 'KETERANGAN'] },
  MILESTONE: { group: 'SISTEM', cols: ['PROYEK_ID', 'NAMA', 'TANGGAL:date', 'STATUS'] },
  // Kontrol proyek (tahap 4)
  REALISASI: { group: 'PELAKSANAAN', cols: ['PROYEK_ID', 'TANGGAL:date', 'ITEM_ID', 'KATEGORI', 'VENDOR', 'URAIAN', 'NOMINAL:rp', 'FOTO_NOTA', 'CATATAN', 'FOTO_URL', 'DIINPUT_OLEH'] },
  CCO: { group: 'PELAKSANAAN', cols: ['PROYEK_ID', 'NOMOR', 'JENIS', 'ITEM_ID', 'URAIAN', 'SATUAN', 'VOLUME:num', 'HARGA_SATUAN:rp', 'JUMLAH:rp', 'ALASAN', 'STATUS', 'TANGGAL:date', 'DIVISI_ID', 'KODE_AHSP', 'HARGA_LAMA:rp', 'SPESIFIKASI', 'CATATAN', 'FILE_ID', 'FILE_URL'] },
  OPNAME: { group: 'PELAKSANAAN', cols: ['PROYEK_ID', 'PERIODE', 'TANGGAL:date', 'PERSEN_PROGRES:pct', 'NILAI_PROGRES:rp', 'TERMIN_SEBELUMNYA:rp', 'RETENSI:rp', 'NILAI_TAGIHAN:rp', 'STATUS', 'FILE_ID', 'NOMOR', 'JENIS', 'MINGGU_KE:num', 'NILAI_KONTRAK:rp', 'RETENSI_PCT:pct', 'NILAI_PERIODE:rp', 'FILE_URL', 'CATATAN', 'DETAIL_JSON:json'] },
  // Penawaran
  PENAWARAN: { group: 'PENAWARAN', cols: ['PROYEK_ID', 'NOMOR', 'REV:num', 'TANGGAL:date', 'KLIEN_NAMA', 'KLIEN_ALAMAT', 'PERIHAL', 'PROFIL_ID', 'DETAIL', 'MARKUP_PCT:pct', 'MARKUP_DIVISI_JSON:json', 'DIVISI_SETTING_JSON:json', 'BULAT_SATUAN:num', 'BULAT_TOTAL:num', 'DISKON_TIPE', 'DISKON_NILAI:num', 'PPN_MODE', 'PPN_PCT:pct', 'MASA_BERLAKU_HARI:num', 'WAKTU_PELAKSANAAN', 'TERMIN', 'GARANSI', 'LINGKUP', 'EXCLUSIONS', 'SYARAT', 'SPESIFIKASI', 'TAMPIL_REKENING', 'TAMPIL_HARGA_M2', 'TAMPIL_CREDIT', 'SUBTOTAL:rp', 'DISKON:rp', 'PPN:rp', 'TOTAL:rp', 'TOTAL_RAB_INTERNAL:rp', 'STATUS', 'FILE_ID', 'FILE_URL'] },
  PENAWARAN_ITEM: { group: 'PENAWARAN', cols: ['PENAWARAN_ID', 'URUTAN:num', 'TIPE_BARIS', 'KODE', 'URAIAN', 'SATUAN', 'VOLUME:num', 'HARGA_SATUAN:rp', 'JUMLAH:rp', 'SPESIFIKASI'] },
  // Sistem
  AKSES: { group: 'SISTEM', cols: ['NAMA', 'PERAN', 'PIN_HASH', 'PROYEK_IDS', 'AKTIF', 'TERAKHIR_LOGIN:ts', 'CATATAN'] },
  FILES: { group: 'SISTEM', cols: ['PROYEK_ID', 'JENIS', 'NAMA_FILE', 'FILE_ID', 'URL', 'TERKAIT', 'TANGGAL:date', 'UKURAN:num'] },
  LOG: { group: 'SISTEM', cols: ['WAKTU:ts', 'AKSI', 'TABEL', 'RECORD_ID', 'DETAIL', 'SUMBER'] }
};

/* Dropdown (data validation) per kolom */
var DROPDOWN = {
  'PROYEK.MODE': ['RUMAH', 'GEDUNG'],
  'PROYEK.STATUS': ['DRAFT', 'PENAWARAN', 'BERJALAN', 'SELESAI', 'BATAL'],
  'DIVISI.MODE': ['SEMUA', 'GEDUNG'],
  'RAB_ITEM.TIPE_BARIS': ['DIVISI', 'ITEM', 'SUBTOTAL', 'TOTAL'],
  'RAB_ITEM.JENIS_HARGA': ['AHSP', 'PASAR'],
  'VOLUME_DETAIL.PENGURANG': ['YA', 'TIDAK'],
  'AHSP_HEADER.JENIS': ['RESMI', 'MODIFIKASI', 'CUSTOM', 'HARGA PASAR'],
  'AHSP_HEADER.STATUS': ['TERVERIFIKASI', 'PERLU CEK'],
  'AHSP_HEADER.SCOPE': ['UMUM', 'PROYEK'],
  'AHSP_KOMPONEN.JENIS_KOMPONEN': ['Upah', 'Bahan', 'Alat'],
  'HSD_UPAH.STATUS': ['OK', 'PERLU HARGA', 'PERLU CEK'],
  'HSD_BAHAN.STATUS': ['OK', 'PERLU HARGA', 'PERLU CEK'],
  'HSD_ALAT.STATUS': ['OK', 'PERLU HARGA', 'PERLU CEK'],
  'PROFIL_USAHA.PKP': ['YA', 'TIDAK'],
  'PROFIL_USAHA.TAMPILKAN_TTD': ['YA', 'TIDAK'],
  'PENAWARAN.STATUS': ['DRAFT', 'TERKIRIM', 'NEGOSIASI', 'DITERIMA', 'DITOLAK', 'KEDALUWARSA'],
  'PENAWARAN.DETAIL': ['RINGKAS', 'STANDAR', 'LUMPSUM'],
  'PENAWARAN.PPN_MODE': ['BELUM', 'TERMASUK', 'TIDAK'],
  'CCO.STATUS': ['DIAJUKAN', 'DISETUJUI', 'DITOLAK'],
  'CCO.JENIS': ['TAMBAH', 'KURANG', 'GANTI'],
  'REALISASI.KATEGORI': ['BAHAN', 'UPAH', 'ALAT', 'LAINNYA'],
  'OPNAME.JENIS': ['TERMIN', 'RETENSI'],
  'OPNAME.STATUS': ['DRAFT', 'DIAJUKAN', 'DISETUJUI', 'DIBAYAR'],
  'AKSES.PERAN': ['PENGAWAS', 'VIEWER'],
  'AKSES.AKTIF': ['YA', 'TIDAK'],
  'BORONGAN_UPAH.MODE': ['M2', 'ITEM', 'HARIAN_TUKANG', 'HARIAN_PEKERJA'],
  'TERMIN_UPAH.BASIS': ['AHSP', 'M2', 'ITEM', 'HARIAN'],
  'TERMIN_UPAH.STATUS': ['DRAFT', 'DISETUJUI'],
  'JADWAL.JENIS_RELASI': ['FS', 'SS'],
  'MILESTONE.STATUS': ['RENCANA', 'TERCAPAI', 'TERLAMBAT']
};

var FORMAT = {
  rp: '_("Rp"* #,##0_);[Red]_("Rp"* -#,##0_);_("Rp"* "-"_);_(@_)',
  num: '#,##0.00', coef: '#,##0.0000', pct: '0.00"%"', date: 'dd/mm/yyyy', text: '@'
};

var DEFAULT_SETTINGS = [
  ['OVERHEAD_PCT', 10, 'Biaya umum & keuntungan default (%), maksimal sesuai Permen yang berlaku'],
  ['PPN_PCT', 11, 'PPN efektif default (%), sesuaikan bila aturan berubah'],
  ['KONTINJENSI_PCT', 5, 'Biaya tak terduga default (%), isi 0 untuk mematikan'],
  ['RETENSI_PCT', 5, 'Retensi default (%) yang ditahan dari setiap tagihan termin'],
  ['SERAPAN_WASPADA', 90, 'Kontrol biaya: status waspada bila serapan anggaran divisi > … %'],
  ['SELISIH_BIAYA_PROGRES', 15, 'Peringatan bila % biaya terpakai melebihi % progres fisik lebih dari … poin'],
  ['ESKALASI_PCT_DEFAULT', 5, 'Eskalasi harga material default (%/tahun) — Mode Gedung, proyek > 6 bulan'],
  ['PORSI_BAHAN_PASAR', 60, 'Asumsi porsi bahan pada item harga pasar untuk hitung eskalasi (%)'],
  ['PEMBULATAN', 1000, 'Pembulatan total RAB ke atas (1 = tanpa pembulatan)'],
  ['WILAYAH_DEFAULT', 'Kabupaten Bogor', 'Wilayah harga default proyek baru'],
  ['VERSI_AHSP_AKTIF', 'CONTOH', 'Versi library AHSP yang aktif'],
  ['MODE_DEFAULT', 'RUMAH', 'RUMAH atau GEDUNG'],
  ['AMBANG_DEVIASI_HSP', 30, 'Peringatan bila HSP custom menyimpang dari rata-rata sejenis (±%)'],
  ['FORMAT_NOMOR_PENAWARAN', '{NO}/PNW/RAB/{BLN}/{THN}', 'Format nomor penawaran'],
  ['MASA_BERLAKU_HARI', 14, 'Masa berlaku penawaran default (hari)'],
  ['TERMIN_DEFAULT', 'DP 30% – Progres 30% – Progres 30% – Retensi 10%', 'Skema termin default penawaran'],
  ['TAMPIL_CREDIT', 'YA', 'Tampilkan tulisan kecil "dibuat dengan RAB Rumah & Gedung by DG" di dokumen'],
  ['JAM_KERJA_HARI', 8, 'Jam kerja per hari (1 OH = 8 jam)'],
  ['HARI_KERJA_MINGGU', 6, 'Hari kerja per minggu'],
  ['RASIO_KEPALA_TUKANG', 10, 'Peringatan bila 1 kepala tukang membawahi lebih dari … tukang'],
  ['RASIO_MANDOR_PEKERJA', 20, 'Peringatan bila 1 mandor membawahi lebih dari … pekerja + tukang'],
  ['DEVIASI_WASPADA', 5, 'Kurva S: status waspada bila aktual tertinggal ≥ … % dari rencana'],
  ['DEVIASI_KRITIS', 10, 'Kurva S: status kritis bila aktual tertinggal ≥ … % dari rencana'],
  ['ROOT_FOLDER_ID', '', 'Diisi otomatis saat setup'],
  ['MASTER_FOLDER_ID', '', 'Diisi otomatis saat setup'],
  ['PROFIL_FOLDER_ID', '', 'Diisi otomatis saat setup'],
  ['BACKUP_FOLDER_ID', '', 'Diisi otomatis saat setup']
];

/* ============================================================
 * MENU
 * ============================================================ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('RAB Rumah & Gedung')
    .addItem('1. Setup Awal', 'setupSheets')
    .addItem('2. Tampilkan Token API', 'showToken')
    .addSeparator()
    .addItem('Perbarui Struktur Sheet (setelah update)', 'upgradeSheets')
    .addItem('Rapikan Tampilan (Apply Theme)', 'applyThemeAll')
    .addItem('Backup Sekarang', 'weeklyBackup')
    .addItem('Buat Token Baru', 'resetToken')
    .addToUi();
}

/* ============================================================
 * SETUP
 * ============================================================ */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try { ss.setSpreadsheetLocale('id_ID'); } catch (e) {}
  ss.setSpreadsheetTimeZone(APP.tz);
  Object.keys(SCHEMA).forEach(function (name) { ensureSheet_(ss, name); });
  // hapus sheet bawaan yang kosong
  ['Sheet1', 'Lembar1', 'Sheet 1'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() <= 1 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  seedSettings_();
  seedMaster_();
  setupDrive_();
  getToken_();
  setupTriggers_();
  applyThemeAll();
  orderSheets_(ss);
  log_('SETUP', '', '', 'Setup awal selesai', 'SCRIPT');
  try {
    SpreadsheetApp.getUi().alert('Setup selesai ✅',
      'Semua sheet, folder Drive, dan backup mingguan sudah dibuat.\n\n' +
      'Langkah berikutnya: Deploy → New deployment → Web app.\n' +
      'Lalu buka menu "RAB Rumah & Gedung → 2. Tampilkan Token API" untuk menyalin token ke aplikasi.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
}

/* Dipakai setelah update versi: tambah sheet/kolom/pengaturan baru tanpa menyentuh data */
function upgradeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var before = {};
  ss.getSheets().forEach(function (s) { before[s.getName()] = s.getLastColumn(); });
  Object.keys(SCHEMA).forEach(function (name) { ensureSheet_(ss, name); });
  seedSettings_();
  seedLibur_();
  Object.keys(SCHEMA).forEach(function (name) { var sh = sheet_(name); if (!(name in before) || sh.getLastColumn() !== before[name]) applyTheme(name); });
  orderSheets_(ss);
  log_('UPGRADE', '', '', 'Struktur diperbarui ke v' + APP.version, 'SCRIPT');
  try { SpreadsheetApp.getUi().alert('Struktur sheet sudah diperbarui ke v' + APP.version + ' ✅'); } catch (e) {}
}

function headersOf_(name) {
  return ['ID'].concat(SCHEMA[name].cols.map(function (c) { return c.split(':')[0]; })).concat(['UPDATED_AT', 'DELETED']);
}
function typesOf_(name) {
  var t = { ID: 'text', UPDATED_AT: 'ts', DELETED: 'text' };
  SCHEMA[name].cols.forEach(function (c) { var p = c.split(':'); t[p[0]] = p[1] || 'text'; });
  return t;
}

function ensureSheet_(ss, name) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var want = headersOf_(name);
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var have = sh.getRange(1, 1, 1, lastCol).getValues()[0].filter(String);
  if (!have.length) {
    sh.getRange(1, 1, 1, want.length).setValues([want]);
  } else {
    // tambahkan kolom baru di akhir tanpa merusak data lama
    var missing = want.filter(function (h) { return have.indexOf(h) < 0; });
    if (missing.length) sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

function orderSheets_(ss) {
  var order = Object.keys(SCHEMA);
  order.forEach(function (n, i) {
    var sh = ss.getSheetByName(n);
    if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 1); }
  });
  ss.setActiveSheet(ss.getSheetByName('PROYEK'));
}

function seedSettings_() {
  var rows = readTable_('PENGATURAN');
  var have = {};
  rows.forEach(function (r) { have[r.ID] = 1; });
  var add = DEFAULT_SETTINGS.filter(function (s) { return !have[s[0]]; }).map(function (s) {
    return { ID: s[0], NILAI: s[1], KETERANGAN: s[2] };
  });
  if (add.length) appendRows_('PENGATURAN', add);
}

function seedMaster_() {
  if (!readTable_('WILAYAH').length) {
    appendRows_('WILAYAH', [{ ID: 'W001', NAMA: 'Kabupaten Bogor', PROVINSI: 'Jawa Barat', AKTIF: 'YA' }]);
  }
  if (!readTable_('DIVISI').length) {
    var div = [
      ['I', 'PEKERJAAN PERSIAPAN', 'SEMUA'], ['II', 'SMKK (SISTEM MANAJEMEN KESELAMATAN KONSTRUKSI)', 'SEMUA'],
      ['III', 'PEKERJAAN TANAH', 'SEMUA'], ['IV', 'PEKERJAAN PONDASI', 'SEMUA'], ['V', 'PEKERJAAN STRUKTUR BETON', 'SEMUA'],
      ['VI', 'PEKERJAAN STRUKTUR BAJA / ATAP', 'SEMUA'], ['VII', 'PEKERJAAN DINDING', 'SEMUA'],
      ['VIII', 'PEKERJAAN LANTAI & DINDING KERAMIK', 'SEMUA'], ['IX', 'PEKERJAAN PLAFON', 'SEMUA'],
      ['X', 'PEKERJAAN KUSEN, PINTU, JENDELA', 'SEMUA'], ['XI', 'PEKERJAAN SANITAIR', 'SEMUA'],
      ['XII', 'PEKERJAAN PLUMBING', 'SEMUA'], ['XIII', 'PEKERJAAN MEKANIKAL', 'SEMUA'],
      ['XIV', 'PEKERJAAN ELEKTRIKAL', 'SEMUA'], ['XV', 'PEKERJAAN ELEKTRONIK / LOW CURRENT', 'SEMUA'],
      ['XVI', 'PEKERJAAN PROTEKSI KEBAKARAN', 'GEDUNG'], ['XVII', 'PEKERJAAN FINISHING', 'SEMUA'],
      ['XVIII', 'PEKERJAAN LUAR / LANSEKAP', 'SEMUA'], ['XIX', 'PEKERJAAN LAIN-LAIN', 'SEMUA']
    ];
    appendRows_('DIVISI', div.map(function (d, i) {
      return { ID: 'DIV' + ('0' + (i + 1)).slice(-2), KODE: d[0], NAMA: d[1], URUTAN: i + 1, MODE: d[2] };
    }));
  }
  if (!readTable_('TEMPLATE_SYARAT').length) {
    var t = [
      ['LINGKUP', 'Lingkup termasuk', 'Pengadaan material dan tenaga kerja sesuai rincian penawaran\nPekerjaan persiapan dan pembersihan akhir\nPenerapan keselamatan kerja (SMKK)', 1],
      ['EXCLUSION', 'Tidak termasuk', 'Biaya perizinan (PBG/SLF) dan retribusi\nPenyambungan daya PLN dan PDAM baru\nFurnitur lepas, gorden, dan perlengkapan elektronik\nPekerjaan di luar gambar kerja/rincian penawaran', 2],
      ['SYARAT', 'Syarat & ketentuan', 'Pekerjaan tambah/kurang dihitung berdasarkan harga satuan penawaran dan disepakati tertulis\nJadwal pelaksanaan dihitung sejak DP diterima dan lahan siap\nPerubahan spesifikasi atas permintaan pemilik dihitung sebagai pekerjaan tambah/kurang', 3],
      ['SPESIFIKASI', 'Spesifikasi material utama', 'Semen setara Tiga Roda/Holcim\nBesi beton SNI\nKeramik lantai 60×60 setara Roman/Platinum\nPipa air bersih PVC AW setara Rucika/Wavin\nKabel NYM setara Supreme/Eterna', 4]
    ];
    appendRows_('TEMPLATE_SYARAT', t.map(function (x, i) { return { ID: 'TS' + (i + 1), JENIS: x[0], JUDUL: x[1], ISI: x[2], URUTAN: x[3] }; }));
  }
  if (!readTable_('AHSP_HEADER').length) seedContoh_();
  seedLibur_();
}

/* Libur nasional & cuti bersama (SKB 3 Menteri) 2026–2027. Bisa diedit di sheet LIBUR / aplikasi. */
var LIBUR_SKB = [
  ['2026-01-01', 'Tahun Baru 2026 Masehi'], ['2026-01-16', 'Isra Mikraj Nabi Muhammad saw.'], ['2026-02-16', 'Cuti bersama Tahun Baru Imlek'], ['2026-02-17', 'Tahun Baru Imlek 2577 Kongzili'],
  ['2026-03-18', 'Cuti bersama Nyepi'], ['2026-03-19', 'Hari Suci Nyepi (Tahun Baru Saka 1948)'], ['2026-03-20', 'Cuti bersama Idulfitri'], ['2026-03-21', 'Idulfitri 1447 H'], ['2026-03-22', 'Idulfitri 1447 H'],
  ['2026-03-23', 'Cuti bersama Idulfitri'], ['2026-03-24', 'Cuti bersama Idulfitri'], ['2026-04-03', 'Wafat Yesus Kristus'], ['2026-04-05', 'Kebangkitan Yesus Kristus (Paskah)'], ['2026-05-01', 'Hari Buruh Internasional'],
  ['2026-05-14', 'Kenaikan Yesus Kristus'], ['2026-05-15', 'Cuti bersama Kenaikan Yesus Kristus'], ['2026-05-27', 'Iduladha 1447 H'], ['2026-05-28', 'Cuti bersama Iduladha'], ['2026-05-31', 'Hari Raya Waisak 2570 BE'],
  ['2026-06-01', 'Hari Lahir Pancasila'], ['2026-06-16', '1 Muharam Tahun Baru Islam 1448 H'], ['2026-08-17', 'Proklamasi Kemerdekaan'], ['2026-08-25', 'Maulid Nabi Muhammad saw.'],
  ['2026-12-24', 'Cuti bersama Natal'], ['2026-12-25', 'Kelahiran Yesus Kristus'],
  ['2027-01-01', 'Tahun Baru 2027 Masehi'], ['2027-01-05', 'Isra Mikraj Nabi Muhammad saw.'], ['2027-02-05', 'Cuti bersama Tahun Baru Imlek'], ['2027-02-06', 'Tahun Baru Imlek 2578 Kongzili'],
  ['2027-03-08', 'Hari Suci Nyepi (Tahun Baru Saka 1949)'], ['2027-03-09', 'Cuti bersama Idulfitri'], ['2027-03-10', 'Idulfitri 1448 H'], ['2027-03-11', 'Idulfitri 1448 H'], ['2027-03-12', 'Cuti bersama Idulfitri'],
  ['2027-03-15', 'Cuti bersama Idulfitri'], ['2027-03-25', 'Cuti bersama Wafat Yesus Kristus'], ['2027-03-26', 'Wafat Yesus Kristus'], ['2027-03-28', 'Kebangkitan Yesus Kristus (Paskah)'], ['2027-05-01', 'Hari Buruh Internasional'],
  ['2027-05-06', 'Kenaikan Yesus Kristus'], ['2027-05-17', 'Iduladha 1448 H'], ['2027-05-18', 'Cuti bersama Iduladha'], ['2027-05-19', 'Cuti bersama Waisak'], ['2027-05-20', 'Hari Raya Waisak 2571 BE'],
  ['2027-06-01', 'Hari Lahir Pancasila'], ['2027-06-06', '1 Muharam Tahun Baru Islam 1449 H'], ['2027-08-15', 'Maulid Nabi Muhammad saw.'], ['2027-08-17', 'Proklamasi Kemerdekaan'],
  ['2027-12-24', 'Cuti bersama Natal'], ['2027-12-25', 'Kelahiran Yesus Kristus'], ['2027-12-26', 'Isra Mikraj Nabi Muhammad saw.']
];
function seedLibur_() {
  if (readTable_('LIBUR').length) return;
  appendRows_('LIBUR', LIBUR_SKB.map(function (l) { return { ID: 'L' + l[0].replace(/-/g, ''), TANGGAL: l[0], KETERANGAN: l[1] + ' (SKB 3 Menteri)' }; }));
}

/* Data CONTOH untuk testing saja. BUKAN data resmi. Hapus dari aplikasi:
 * Master → Pengaturan → "Hapus semua data CONTOH" setelah library resmi di-import. */
function seedContoh_() {
  var w = 'Kabupaten Bogor', th = new Date().getFullYear(), today = new Date();
  var src = 'CONTOH – BUKAN DATA RESMI';
  var hsd = function (id, kode, nama, sat, harga) {
    return { ID: id, KODE_HSD: kode, NAMA: nama, SATUAN: sat, WILAYAH: w, TAHUN: th, HARGA: harga, SUMBER: src, STATUS: 'PERLU CEK', TANGGAL_UPDATE: today };
  };
  appendRows_('HSD_UPAH', [
    hsd('CTH-L1', 'CTH.L.01', 'Pekerja', 'OH', 120000), hsd('CTH-L2', 'CTH.L.02', 'Tukang batu', 'OH', 150000),
    hsd('CTH-L3', 'CTH.L.03', 'Kepala tukang', 'OH', 170000), hsd('CTH-L4', 'CTH.L.04', 'Mandor', 'OH', 180000)
  ]);
  appendRows_('HSD_BAHAN', [
    hsd('CTH-M1', 'CTH.M.01', 'Semen Portland (PC)', 'kg', 1500), hsd('CTH-M2', 'CTH.M.02', 'Pasir pasang', 'm3', 300000),
    hsd('CTH-M3', 'CTH.M.03', 'Bata merah', 'bh', 900), hsd('CTH-M4', 'CTH.M.04', 'Pasir beton', 'm3', 350000),
    hsd('CTH-M5', 'CTH.M.05', 'Kerikil/split', 'm3', 380000), hsd('CTH-M6', 'CTH.M.06', 'Air', 'liter', 50),
    hsd('CTH-M7', 'CTH.M.07', 'Pasir urug', 'm3', 250000)
  ]);
  appendRows_('HSD_ALAT', [hsd('CTH-E1', 'CTH.E.01', 'Molen 0,35 m3', 'jam', 60000)]);
  var H = function (k, u, s, d, tag) {
    return { ID: 'H-' + k, KODE_AHSP: k, URAIAN: u, SATUAN: s, DIVISI: d, JENIS: 'CUSTOM', SUMBER: src, DASAR_PERHITUNGAN: 'Angka ilustrasi untuk uji aplikasi', TAG: tag, SCOPE: 'UMUM', VERSI: 'CONTOH', STATUS: 'PERLU CEK', DIBUAT_OLEH: 'SETUP', TANGGAL: today };
  };
  appendRows_('AHSP_HEADER', [
    H('CTH-01', 'Galian tanah biasa sedalam 1 m', 'm3', 'PEKERJAAN TANAH', 'galian, tanah, pondasi'),
    H('CTH-02', 'Urugan pasir bawah pondasi/lantai', 'm3', 'PEKERJAAN TANAH', 'urugan, pasir, urug'),
    H('CTH-03', 'Pasangan dinding bata merah 1PC:4PP', 'm2', 'PEKERJAAN DINDING', 'dinding, bata, pasangan'),
    H('CTH-04', 'Plesteran 1PC:4PP tebal 15 mm', 'm2', 'PEKERJAAN DINDING', 'plester, plesteran, dinding'),
    H('CTH-05', 'Beton mutu fc 20 MPa site mix', 'm3', 'PEKERJAAN STRUKTUR BETON', 'beton, cor, k-250')
  ]);
  var K = [];
  var k = function (a, n, j, h, nama, sat, koef) { K.push({ ID: a + '-' + n, KODE_AHSP: a, URUTAN: n, JENIS_KOMPONEN: j, KODE_HSD: h, NAMA_KOMPONEN: nama, SATUAN: sat, KOEFISIEN: koef }); };
  k('CTH-01', 1, 'Upah', 'CTH.L.01', 'Pekerja', 'OH', 0.75); k('CTH-01', 2, 'Upah', 'CTH.L.04', 'Mandor', 'OH', 0.025);
  k('CTH-02', 1, 'Bahan', 'CTH.M.07', 'Pasir urug', 'm3', 1.2); k('CTH-02', 2, 'Upah', 'CTH.L.01', 'Pekerja', 'OH', 0.3); k('CTH-02', 3, 'Upah', 'CTH.L.04', 'Mandor', 'OH', 0.01);
  k('CTH-03', 1, 'Bahan', 'CTH.M.03', 'Bata merah', 'bh', 70); k('CTH-03', 2, 'Bahan', 'CTH.M.01', 'Semen', 'kg', 11.5); k('CTH-03', 3, 'Bahan', 'CTH.M.02', 'Pasir pasang', 'm3', 0.043);
  k('CTH-03', 4, 'Upah', 'CTH.L.01', 'Pekerja', 'OH', 0.3); k('CTH-03', 5, 'Upah', 'CTH.L.02', 'Tukang batu', 'OH', 0.1); k('CTH-03', 6, 'Upah', 'CTH.L.03', 'Kepala tukang', 'OH', 0.01); k('CTH-03', 7, 'Upah', 'CTH.L.04', 'Mandor', 'OH', 0.015);
  k('CTH-04', 1, 'Bahan', 'CTH.M.01', 'Semen', 'kg', 6.24); k('CTH-04', 2, 'Bahan', 'CTH.M.02', 'Pasir pasang', 'm3', 0.024);
  k('CTH-04', 3, 'Upah', 'CTH.L.01', 'Pekerja', 'OH', 0.3); k('CTH-04', 4, 'Upah', 'CTH.L.02', 'Tukang batu', 'OH', 0.15); k('CTH-04', 5, 'Upah', 'CTH.L.03', 'Kepala tukang', 'OH', 0.015); k('CTH-04', 6, 'Upah', 'CTH.L.04', 'Mandor', 'OH', 0.015);
  k('CTH-05', 1, 'Bahan', 'CTH.M.01', 'Semen', 'kg', 352); k('CTH-05', 2, 'Bahan', 'CTH.M.04', 'Pasir beton', 'm3', 0.52); k('CTH-05', 3, 'Bahan', 'CTH.M.05', 'Kerikil/split', 'm3', 0.76); k('CTH-05', 4, 'Bahan', 'CTH.M.06', 'Air', 'liter', 215);
  k('CTH-05', 5, 'Upah', 'CTH.L.01', 'Pekerja', 'OH', 1.65); k('CTH-05', 6, 'Upah', 'CTH.L.02', 'Tukang batu', 'OH', 0.275); k('CTH-05', 7, 'Upah', 'CTH.L.03', 'Kepala tukang', 'OH', 0.028); k('CTH-05', 8, 'Upah', 'CTH.L.04', 'Mandor', 'OH', 0.083);
  k('CTH-05', 9, 'Alat', 'CTH.E.01', 'Molen', 'jam', 0.25);
  appendRows_('AHSP_KOMPONEN', K);
  appendRows_('AHSP_VERSI', [{ ID: 'V-CONTOH', VERSI: 'CONTOH', SUMBER: src, TANGGAL: today, JUMLAH_AHSP: 5, JUMLAH_KOMPONEN: K.length, CATATAN: 'Data uji. Ganti dengan import lampiran resmi.' }]);
}

/* ---------- Folder Drive ---------- */
function setupDrive_() {
  var S = settingsMap_();
  var root = folderById_(S.ROOT_FOLDER_ID);
  if (!root) {
    var it = DriveApp.getFoldersByName(APP.name);
    root = it.hasNext() ? it.next() : DriveApp.createFolder(APP.name);
    setSetting_('ROOT_FOLDER_ID', root.getId());
  }
  // pindahkan spreadsheet ke folder utama agar rapi
  try {
    var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    if (!isInFolder_(file, root.getId())) file.moveTo(root);
  } catch (e) {}
  var master = folderById_(S.MASTER_FOLDER_ID) || childFolder_(root, '00_MASTER');
  setSetting_('MASTER_FOLDER_ID', master.getId());
  var profil = folderById_(S.PROFIL_FOLDER_ID) || childFolder_(master, 'Profil');
  setSetting_('PROFIL_FOLDER_ID', profil.getId());
  var backup = folderById_(S.BACKUP_FOLDER_ID) || childFolder_(root, '99_Backup');
  setSetting_('BACKUP_FOLDER_ID', backup.getId());
}
var PROJECT_SUBFOLDERS = ['01_RAB', '02_Penawaran', '03_Jadwal', '04_Opname-BAP', '05_Nota', '06_Progres', '07_CCO'];

function projectFolder_(proyekId) {
  var S = settingsMap_();
  var root = folderById_(S.ROOT_FOLDER_ID);
  if (!root) { setupDrive_(); root = folderById_(settingsMap_().ROOT_FOLDER_ID); }
  var sh = sheet_('PROYEK'), data = sh.getDataRange().getValues(), H = data[0];
  var iId = H.indexOf('ID'), iF = H.indexOf('FOLDER_ID'), iK = H.indexOf('KODE'), iN = H.indexOf('NAMA'), iU = H.indexOf('UPDATED_AT');
  for (var r = 1; r < data.length; r++) {
    if (data[r][iId] !== proyekId) continue;
    var f = folderById_(data[r][iF]);
    if (!f) {
      var name = (data[r][iK] || proyekId) + ' - ' + (data[r][iN] || 'Proyek');
      f = childFolder_(root, safeName_(name));
      sh.getRange(r + 1, iF + 1).setValue(f.getId());
      sh.getRange(r + 1, iU + 1).setValue(nowIso_());
    }
    PROJECT_SUBFOLDERS.forEach(function (s) { childFolder_(f, s); });
    return f;
  }
  throw new Error('Proyek tidak ditemukan: ' + proyekId);
}
function childFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function folderById_(id) {
  if (!id) return null;
  try { var f = DriveApp.getFolderById(id); return f.isTrashed() ? null : f; } catch (e) { return null; }
}
function isInFolder_(file, folderId) {
  var p = file.getParents();
  while (p.hasNext()) if (p.next().getId() === folderId) return true;
  return false;
}
/* Pastikan file ada di dalam folder aplikasi (keamanan getFile/share) */
function isInRoot_(fileOrFolder) {
  var rootId = settingsMap_().ROOT_FOLDER_ID;
  var queue = [fileOrFolder], depth = 0;
  while (queue.length && depth < 6) {
    var next = [];
    for (var i = 0; i < queue.length; i++) {
      var ps = queue[i].getParents();
      while (ps.hasNext()) {
        var p = ps.next();
        if (p.getId() === rootId) return true;
        next.push(p);
      }
    }
    queue = next; depth++;
  }
  return false;
}
function safeName_(s) { return String(s).replace(/[\\/:*?"<>|]+/g, '-').substring(0, 120); }

/* ---------- Token & trigger ---------- */
function getToken_() {
  var p = PropertiesService.getScriptProperties();
  var t = p.getProperty('API_TOKEN');
  if (!t) { t = Utilities.getUuid().replace(/-/g, '').substring(0, 24); p.setProperty('API_TOKEN', t); }
  return t;
}
function showToken() {
  SpreadsheetApp.getUi().alert('Token API', 'Salin token ini ke aplikasi (Master → Pengaturan → Koneksi):\n\n' + getToken_() +
    '\n\nJangan bagikan token ke orang lain.', SpreadsheetApp.getUi().ButtonSet.OK);
}
function resetToken() {
  PropertiesService.getScriptProperties().deleteProperty('API_TOKEN');
  showToken();
}
function setupTriggers_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'weeklyBackup'; });
  if (!has) ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).inTimezone(APP.tz).create();
}

/* ---------- Backup mingguan (simpan 8 terakhir) ---------- */
function weeklyBackup() {
  var S = settingsMap_();
  var folder = folderById_(S.BACKUP_FOLDER_ID);
  if (!folder) { setupDrive_(); folder = folderById_(settingsMap_().BACKUP_FOLDER_ID); }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'BACKUP_' + ss.getName() + '_' + Utilities.formatDate(new Date(), APP.tz, 'yyyyMMdd_HHmm');
  DriveApp.getFileById(ss.getId()).makeCopy(name, folder);
  var files = [], it = folder.getFiles();
  while (it.hasNext()) { var f = it.next(); if (f.getName().indexOf('BACKUP_') === 0) files.push(f); }
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(8).forEach(function (f) { f.setTrashed(true); });
  log_('BACKUP', '', '', name, 'SCRIPT');
}

/* ============================================================
 * TAMPILAN SHEET (identik dengan PWA)
 * ============================================================ */
function applyThemeAll() {
  Object.keys(SCHEMA).forEach(function (n) { applyTheme(n); });
}
function applyTheme(name) {
  var sh = sheet_(name);
  if (!sh) return;
  var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var T = typesOf_(name), nCols = H.length;
  var maxRows = Math.max(sh.getMaxRows(), 200);
  if (sh.getMaxRows() < 200) sh.insertRowsAfter(sh.getMaxRows(), 200 - sh.getMaxRows());
  var all = sh.getRange(1, 1, maxRows, nCols);
  all.setFontFamily('Inter').setFontSize(10).setFontColor(THEME.text).setVerticalAlignment('middle');
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(1);
  if (SCHEMA[name].freezeCols) sh.setFrozenColumns(SCHEMA[name].freezeCols);
  sh.setTabColor(THEME.tab[SCHEMA[name].group] || THEME.muted);
  // header
  sh.getRange(1, 1, 1, nCols).setBackground(THEME.primary).setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sh.setRowHeight(1, 32);
  // banding
  sh.getBandings().forEach(function (b) { b.remove(); });
  var band = all.applyRowBanding(SpreadsheetApp.BandingTheme.BLUE, true, false);
  band.setHeaderRowColor(THEME.primary).setFirstRowColor('#FFFFFF').setSecondRowColor(THEME.zebra);
  sh.getRange(2, 1, maxRows - 1, nCols).setBorder(null, null, true, null, null, true, THEME.border, SpreadsheetApp.BorderStyle.SOLID);
  // format kolom
  H.forEach(function (h, i) {
    var t = T[h] || 'text', col = sh.getRange(2, i + 1, maxRows - 1, 1);
    if (FORMAT[t]) col.setNumberFormat(FORMAT[t]);
    if (t === 'ts' || t === 'json') col.setNumberFormat('@');
    if (t === 'date' || t === 'rp' || t === 'num' || t === 'coef' || t === 'pct') col.setHorizontalAlignment('right');
    if (h === 'ID' || h === 'UPDATED_AT' || h === 'DELETED' || /_ID$/.test(h)) col.setFontColor(THEME.muted);
    var w = colWidth_(h, t);
    sh.setColumnWidth(i + 1, w);
    if (/^(URAIAN|ISI|LINGKUP|EXCLUSIONS|SYARAT|SPESIFIKASI|CATATAN|DASAR_PERHITUNGAN|ALAMAT|KETERANGAN)$/.test(h)) col.setWrap(true);
    var dd = DROPDOWN[name + '.' + h];
    if (dd) col.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(dd, true).setAllowInvalid(true).build());
  });
  // conditional formatting: status & tipe baris
  var rules = [];
  H.forEach(function (h, i) {
    if (h !== 'STATUS') return;
    var col = sh.getRange(2, i + 1, maxRows - 1, 1);
    Object.keys(STATUS_COLORS).forEach(function (v) {
      var c = STATUS_COLORS[v];
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v)
        .setFontColor(THEME[c]).setBackground(THEME[c + 'Bg']).setBold(true).setRanges([col]).build());
    });
  });
  if (name === 'RAB_ITEM') {
    var iT = H.indexOf('TIPE_BARIS') + 1;
    var L = columnLetter_(iT), body = sh.getRange(2, 1, maxRows - 1, nCols);
    rules.unshift(
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + L + '2="TOTAL"').setBackground(THEME.primaryDark).setFontColor('#FFFFFF').setBold(true).setRanges([body]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + L + '2="SUBTOTAL"').setBackground(THEME.accent50).setBold(true).setRanges([body]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + L + '2="DIVISI"').setBackground(THEME.primary50).setBold(true).setRanges([body]).build()
    );
  }
  var iDel = H.indexOf('DELETED') + 1;
  if (iDel) {
    rules.unshift(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + columnLetter_(iDel) + '2="YA"')
      .setFontColor('#B0B8C1').setStrikethrough(true).setRanges([sh.getRange(2, 1, maxRows - 1, nCols)]).build());
  }
  sh.setConditionalFormatRules(rules);
  // proteksi kolom sistem (peringatan saja)
  sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) { if (p.getDescription() === 'Kolom sistem') p.remove(); });
  ['ID', 'UPDATED_AT'].forEach(function (h) {
    var i = H.indexOf(h);
    if (i >= 0) sh.getRange(1, i + 1, maxRows, 1).protect().setDescription('Kolom sistem').setWarningOnly(true);
  });
}
function colWidth_(h, t) {
  if (/^(URAIAN|NAMA|ISI|LINGKUP|EXCLUSIONS|SYARAT|SPESIFIKASI|DASAR_PERHITUNGAN|ALAMAT|KLIEN_ALAMAT)$/.test(h)) return 280;
  if (/(CATATAN|KETERANGAN|SUMBER|TAG|SINONIM|DETAIL|PERIHAL|TERMIN|URL)/.test(h)) return 200;
  if (t === 'rp') return 130;
  if (t === 'ts') return 170;
  if (t === 'json') return 160;
  if (h === 'ID' || /_ID$/.test(h)) return 110;
  if (t === 'num' || t === 'coef' || t === 'pct') return 90;
  return 120;
}
function columnLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}

/* Edit langsung di Sheet → beri ID & UPDATED_AT agar ikut tersinkron ke PWA */
function onEdit(e) {
  try {
    var sh = e.range.getSheet(), name = sh.getName();
    if (!SCHEMA[name] || e.range.getRow() < 2) return;
    var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var iId = H.indexOf('ID') + 1, iU = H.indexOf('UPDATED_AT') + 1;
    if (e.range.getNumColumns() === 1 && e.range.getColumn() === iU) return;
    var r0 = e.range.getRow(), n = e.range.getNumRows();
    var block = sh.getRange(r0, 1, n, H.length).getValues();
    var stamp = nowIso_(), outIds = [], outTs = [];
    for (var i = 0; i < n; i++) {
      var rowVals = block[i];
      var empty = rowVals.every(function (v, j) { return j === iId - 1 || j === iU - 1 || v === ''; });
      outIds.push([rowVals[iId - 1] || (empty ? '' : Utilities.getUuid())]);
      outTs.push([empty ? '' : stamp]);
    }
    sh.getRange(r0, iId, n, 1).setValues(outIds);
    sh.getRange(r0, iU, n, 1).setValues(outTs);
  } catch (err) { /* abaikan */ }
}

/* ============================================================
 * API (Web App)
 * ============================================================ */
function doGet(e) {
  return json_({ ok: true, app: APP.name, version: APP.version, time: nowIso_() });
}
function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Format request tidak valid' }); }
  if (body && body.action === 'login') { try { return json_(apiLogin_(body)); } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); } }
  var who = body ? authOf_(body.token) : null;
  if (!who) return json_({ ok: false, error: 'Token salah, belum diisi, atau akses sudah dicabut' });
  var owner = who.role === 'OWNER';
  try {
    switch (body.action) {
      case 'ping': return json_({ ok: true, app: APP.name, version: APP.version, time: nowIso_(), role: who.role, user: who.nama || '', projects: who.projects || [] });
      case 'pull': return json_(apiPull_(body, who));
      case 'push': return json_(withLock_(function () { return apiPush_(body, who); }));
      case 'upload': return json_(apiUpload_(body, who));
      case 'getFile': return json_(apiGetFile_(body));
      case 'share': if (!owner) throw new Error('Hanya Owner yang bisa membuat link'); return json_(apiShare_(body));
      case 'projectFolder': if (!owner) throw new Error('Hanya Owner'); return json_({ ok: true, folderId: projectFolder_(body.proyekId).getId() });
      case 'backup': if (!owner) throw new Error('Hanya Owner'); weeklyBackup(); return json_({ ok: true });
      default: return json_({ ok: false, error: 'Aksi tidak dikenal: ' + body.action });
    }
  } catch (err) {
    log_('ERROR', '', '', String(err && err.stack || err), 'API');
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ============================================================
 * AKSES TIM — Owner (token utama), Pengawas & Viewer (nama + PIN)
 * ============================================================ */
var PENGAWAS_TABLES = ['PROGRES', 'REALISASI'];
var PENGAWAS_FOLDERS = ['05_Nota', '06_Progres'];
function sha256Hex_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function sessionSig_(u) {
  var raw = Utilities.computeHmacSha256Signature(u.ID + '|' + u.PIN_HASH + '|' + Calc.norm(u.PERAN) + '|' + u.PROYEK_IDS, getToken_());
  return raw.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('').substring(0, 32);
}
function projList_(s) { return String(s || '').split(/[,;\s]+/).filter(String); }
function authOf_(token) {
  if (!token) return null;
  if (token === getToken_()) return { role: 'OWNER' };
  var m = /^s\.([^.]+)\.([0-9a-f]{32})$/.exec(String(token));
  if (!m) return null;
  var u = readTable_('AKSES').filter(function (r) { return r.ID === m[1]; })[0];
  if (!u || Calc.isDel(u) || Calc.norm(u.AKTIF) === 'TIDAK' || !u.PIN_HASH) return null;
  if (sessionSig_(u) !== m[2]) return null;
  var role = Calc.norm(u.PERAN) === 'PENGAWAS' ? 'PENGAWAS' : 'VIEWER';
  return { role: role, id: u.ID, nama: u.NAMA, projects: projList_(u.PROYEK_IDS) };
}
function apiLogin_(b) {
  var nama = String(b.nama || '').trim(), pin = String(b.pin || '').trim();
  if (!nama || !pin) throw new Error('Isi nama & PIN');
  var cache = CacheService.getScriptCache(), ck = 'LOGIN_' + nama.toUpperCase();
  var fails = +(cache.get(ck) || 0);
  if (fails >= 5) throw new Error('Terlalu banyak percobaan. Coba lagi 15 menit lagi.');
  var cand = readTable_('AKSES').filter(function (r) { return !Calc.isDel(r) && Calc.norm(r.AKTIF) !== 'TIDAK' && Calc.norm(r.NAMA) === Calc.norm(nama); });
  var u = cand.filter(function (r) { return r.PIN_HASH && sha256Hex_('RABPIN|' + r.ID + '|' + pin) === r.PIN_HASH; })[0];
  if (!u) { cache.put(ck, String(fails + 1), 900); log_('LOGIN GAGAL', 'AKSES', '', nama, 'API'); throw new Error('Nama atau PIN salah'); }
  cache.remove(ck);
  setCell_('AKSES', u.ID, 'TERAKHIR_LOGIN', new Date());
  log_('LOGIN', 'AKSES', u.ID, u.NAMA + ' (' + u.PERAN + ')', 'API');
  var role = Calc.norm(u.PERAN) === 'PENGAWAS' ? 'PENGAWAS' : 'VIEWER';
  return { ok: true, token: 's.' + u.ID + '.' + sessionSig_(u), role: role, user: u.NAMA, projects: projList_(u.PROYEK_IDS), app: APP.name, version: APP.version };
}
function setCell_(name, id, col, val) {
  var sh = sheet_(name); if (!sh) return;
  var data = sh.getDataRange().getValues(), H = data[0], iId = H.indexOf('ID'), iC = H.indexOf(col);
  if (iC < 0) return;
  for (var r = 1; r < data.length; r++) if (data[r][iId] === id) { sh.getRange(r + 1, iC + 1).setValue(val); return; }
}
/* tabel yang terikat proyek → disaring sesuai daftar proyek anggota tim */
function allowedProject_(who, pid) { return who.role === 'OWNER' || !who.projects.length || who.projects.indexOf(pid) >= 0; }

/* pull: semua baris yang berubah setelah `since` (ISO). since kosong = semua data. */
function apiPull_(body, who) {
  who = who || { role: 'OWNER' };
  var since = body.since || '';
  var only = body.tables || Object.keys(SCHEMA);
  var out = {}, pnAllowed = null;
  only.forEach(function (name) {
    if (!SCHEMA[name] || name === 'LOG') return;
    if (name === 'AKSES' && who.role !== 'OWNER') return;
    var list = readTable_(name).filter(function (r) { return !since || String(r.UPDATED_AT || '') > since || !r.UPDATED_AT; });
    if (name === 'AKSES') list = list.map(function (r) { var c = Object.assign({}, r); delete c.PIN_HASH; c.ADA_PIN = r.PIN_HASH ? 'YA' : 'TIDAK'; return c; });
    if (who.role !== 'OWNER' && who.projects.length) {
      if (name === 'PROYEK') list = list.filter(function (r) { return allowedProject_(who, r.ID); });
      else if (name === 'PENAWARAN_ITEM') {
        if (!pnAllowed) { pnAllowed = {}; readTable_('PENAWARAN').forEach(function (p) { if (allowedProject_(who, p.PROYEK_ID)) pnAllowed[p.ID] = 1; }); }
        list = list.filter(function (r) { return pnAllowed[r.PENAWARAN_ID]; });
      } else if (headersOf_(name).indexOf('PROYEK_ID') >= 0) list = list.filter(function (r) { return !r.PROYEK_ID || allowedProject_(who, r.PROYEK_ID); });
    }
    out[name] = list;
  });
  return { ok: true, serverTime: nowIso_(), tables: out, role: who.role };
}

/* push: changes = [{table, row}] → upsert berdasarkan ID, last-write-wins via UPDATED_AT */
function apiPush_(body, who) {
  who = who || { role: 'OWNER' };
  if (who.role === 'VIEWER') throw new Error('Akses Viewer hanya untuk melihat');
  var changes = body.changes || [];
  var byTable = {}, rejected = [];
  changes.forEach(function (c) {
    if (!SCHEMA[c.table] || c.table === 'LOG' || !c.row || !c.row.ID) return;
    if (who.role === 'PENGAWAS') {
      if (PENGAWAS_TABLES.indexOf(c.table) < 0 || !allowedProject_(who, c.row.PROYEK_ID)) { rejected.push({ table: c.table, id: c.row.ID }); return; }
    }
    (byTable[c.table] = byTable[c.table] || []).push(c.row);
  });
  if (who.role === 'PENGAWAS') guardPengawas_(byTable, who, rejected);
  var applied = [], conflicts = [], stamp = nowIso_();
  // validasi hitung ulang di server
  if (byTable.RAB_ITEM) validateRabItems_(byTable.RAB_ITEM, byTable.PROYEK || []);
  validateKontrol_(byTable);
  Object.keys(byTable).forEach(function (name) {
    var sh = sheet_(name) || ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), name);
    var data = sh.getDataRange().getValues();
    var H = data[0], T = typesOf_(name);
    var iId = H.indexOf('ID'), iU = H.indexOf('UPDATED_AT');
    var rowOf = {};
    for (var r = 1; r < data.length; r++) if (data[r][iId] !== '') rowOf[data[r][iId]] = r;
    var appends = [];
    byTable[name].forEach(function (inc) {
      var r = rowOf[inc.ID];
      if (r !== undefined) {
        var serverTs = String(data[r][iU] || '');
        if (inc.BASE_UPDATED_AT !== undefined && serverTs && inc.BASE_UPDATED_AT && serverTs > inc.BASE_UPDATED_AT && serverTs > String(inc.UPDATED_AT || '')) {
          conflicts.push({ table: name, id: inc.ID, server: rowToObj_(H, data[r], T) });
          log_('KONFLIK', name, inc.ID, 'Versi Sheet lebih baru, perubahan PWA diabaikan', 'API');
          return;
        }
      }
      inc.UPDATED_AT = stamp;
      var arr = H.map(function (h) { return toCell_(inc[h], T[h]); });
      if (r !== undefined) {
        // pertahankan nilai kolom yang tidak dikirim
        H.forEach(function (h, j) { if (!(h in inc)) arr[j] = data[r][j]; });
        sh.getRange(r + 1, 1, 1, H.length).setValues([arr]);
        data[r] = arr;
      } else {
        appends.push(arr);
        rowOf[inc.ID] = -1;
      }
      var clean = {}; H.forEach(function (h, j) { clean[h] = fromCell_(arr[j], T[h]); });
      applied.push({ table: name, row: clean });
    });
    if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, H.length).setValues(appends);
  });
  if (changes.length) log_('PUSH', Object.keys(byTable).join(','), '', applied.length + ' baris, ' + conflicts.length + ' konflik' + (rejected.length ? ', ' + rejected.length + ' ditolak' : '') + (who.nama ? ' · ' + who.nama : ''), who.role === 'OWNER' ? 'PWA' : who.role);
  return { ok: true, serverTime: stamp, applied: applied, conflicts: conflicts, rejected: rejected };
}
/* Pengawas: baris lama harus milik proyek yang diizinkan; realisasi hanya boleh diubah oleh penginputnya */
function guardPengawas_(byTable, who, rejected) {
  Object.keys(byTable).forEach(function (name) {
    var cur = {}; readTable_(name).forEach(function (r) { cur[r.ID] = r; });
    byTable[name] = byTable[name].filter(function (row) {
      var old = cur[row.ID];
      var ok = !old || (allowedProject_(who, old.PROYEK_ID) && (name !== 'REALISASI' || !old.DIINPUT_OLEH || old.DIINPUT_OLEH === who.nama));
      if (!ok) rejected.push({ table: name, id: row.ID });
      if (ok && name === 'REALISASI') row.DIINPUT_OLEH = old && old.DIINPUT_OLEH ? old.DIINPUT_OLEH : who.nama;
      return ok;
    });
  });
}
/* hitung ulang nilai CCO & opname di server (rumus sama dengan PWA, lewat Calc.gs) */
function validateKontrol_(byTable) {
  (byTable.CCO || []).forEach(function (c) {
    if (Calc.isDel(c) || c.VOLUME === undefined || c.HARGA_SATUAN === undefined) return;
    var j = Calc.ccoJumlah(c);
    if (Math.abs(j - Calc.num(c.JUMLAH)) > 1) { log_('KOREKSI', 'CCO', c.ID, 'JUMLAH ' + c.JUMLAH + ' → ' + j, 'VALIDASI'); c.JUMLAH = j; }
  });
  (byTable.OPNAME || []).forEach(function (o) {
    if (Calc.isDel(o) || Calc.norm(o.JENIS) === 'RETENSI' || o.PERSEN_PROGRES === undefined || o.NILAI_KONTRAK === undefined) return;
    var r = Calc.opname({ persen: o.PERSEN_PROGRES, kontrak: o.NILAI_KONTRAK, prev: o.TERMIN_SEBELUMNYA, retensiPct: o.RETENSI_PCT });
    if (Math.abs(r.tagihan - Calc.num(o.NILAI_TAGIHAN)) > 1 || Math.abs(r.retensi - Calc.num(o.RETENSI)) > 1) log_('KOREKSI', 'OPNAME', o.ID, 'TAGIHAN ' + o.NILAI_TAGIHAN + ' → ' + r.tagihan, 'VALIDASI');
    o.NILAI_PROGRES = r.nilai; o.NILAI_PERIODE = r.periode; o.RETENSI = r.retensi; o.NILAI_TAGIHAN = r.tagihan;
  });
  (byTable.REALISASI || []).forEach(function (r) { if (r.NOMINAL !== undefined && Calc.num(r.NOMINAL) < 0) r.NOMINAL = 0; });
}

function validateRabItems_(items, incomingProjects) {
  var tables = {};
  ['HSD_UPAH', 'HSD_BAHAN', 'HSD_ALAT', 'AHSP_HEADER', 'AHSP_KOMPONEN'].forEach(function (t) { tables[t] = readTable_(t); });
  var idx = Calc.buildIndex(tables);
  var projects = {};
  readTable_('PROYEK').forEach(function (p) { projects[p.ID] = p; });
  incomingProjects.forEach(function (p) { projects[p.ID] = Object.assign({}, projects[p.ID] || {}, p); });
  var S = settingsMap_();
  items.forEach(function (it) {
    if (Calc.norm(it.TIPE_BARIS) === 'DIVISI' || Calc.isDel(it)) return;
    var p = projects[it.PROYEK_ID];
    if (!p) return;
    var c = Calc.itemCalc(idx, it, Calc.ctxFromProject(p, S));
    if (Math.abs(c.hs - Calc.num(it.HARGA_SATUAN)) > 1 || Math.abs(c.jumlah - Calc.num(it.JUMLAH)) > 1) {
      log_('KOREKSI', 'RAB_ITEM', it.ID, 'HS ' + it.HARGA_SATUAN + ' → ' + c.hs, 'VALIDASI');
      it.HARGA_SATUAN = c.hs; it.JUMLAH = c.jumlah;
      it.UPAH = Math.round(c.upah); it.BAHAN = Math.round(c.bahan); it.ALAT = Math.round(c.alat);
    }
  });
}

/* upload file (base64) ke folder proyek/master */
function apiUpload_(b, who) {
  who = who || { role: 'OWNER' };
  if (who.role === 'VIEWER') throw new Error('Akses Viewer hanya untuk melihat');
  if (who.role === 'PENGAWAS' && (PENGAWAS_FOLDERS.indexOf(b.kategori) < 0 || !allowedProject_(who, b.proyekId))) throw new Error('Pengawas hanya bisa mengunggah foto nota & progres');
  var folder;
  if (b.kategori === 'MASTER') folder = folderById_(settingsMap_().MASTER_FOLDER_ID);
  else if (b.kategori === 'PROFIL') folder = folderById_(settingsMap_().PROFIL_FOLDER_ID);
  else {
    if (PROJECT_SUBFOLDERS.indexOf(b.kategori) < 0) throw new Error('Kategori folder tidak valid');
    folder = childFolder_(projectFolder_(b.proyekId), b.kategori);
  }
  if (!folder) { setupDrive_(); return apiUpload_(b); }
  var bytes = Utilities.base64Decode(b.base64);
  var blob = Utilities.newBlob(bytes, b.mime || 'application/octet-stream', safeName_(b.name));
  // timpa file dengan nama sama di folder yang sama (versi baru → nama baru oleh PWA)
  var file = folder.createFile(blob);
  var rec = { ID: Utilities.getUuid(), PROYEK_ID: b.proyekId || '', JENIS: b.jenis || b.kategori, NAMA_FILE: file.getName(), FILE_ID: file.getId(), URL: file.getUrl(), TERKAIT: b.terkait || '', TANGGAL: new Date(), UKURAN: bytes.length, UPDATED_AT: nowIso_() };
  appendRows_('FILES', [rec]);
  return { ok: true, fileId: file.getId(), url: file.getUrl(), name: file.getName(), record: fromRecord_('FILES', rec) };
}
function apiGetFile_(b) {
  var f = DriveApp.getFileById(b.fileId);
  if (!isInRoot_(f)) throw new Error('File di luar folder aplikasi');
  var blob = f.getBlob();
  return { ok: true, mime: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()), name: f.getName() };
}
function apiShare_(b) {
  var f = DriveApp.getFileById(b.fileId);
  if (!isInRoot_(f)) throw new Error('File di luar folder aplikasi');
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  log_('SHARE', 'FILES', b.fileId, 'Link view-only dibuat', 'PWA');
  return { ok: true, url: f.getUrl() };
}

/* ============================================================
 * Helper data
 * ============================================================ */
function sheet_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function nowIso_() { return new Date().toISOString(); }

function readTable_(name) {
  var sh = sheet_(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getDataRange().getValues(), H = data[0], T = typesOf_(name);
  var out = [];
  for (var r = 1; r < data.length; r++) {
    if (data[r].every(function (v) { return v === ''; })) continue;
    var o = rowToObj_(H, data[r], T);
    if (!o.ID) continue;
    out.push(o);
  }
  return out;
}
function rowToObj_(H, row, T) {
  var o = {};
  H.forEach(function (h, j) { if (h) o[h] = fromCell_(row[j], T[h]); });
  return o;
}
function fromRecord_(name, rec) {
  var T = typesOf_(name), o = {};
  headersOf_(name).forEach(function (h) { o[h] = fromCell_(toCell_(rec[h], T[h]), T[h]); });
  return o;
}
function appendRows_(name, objs) {
  if (!objs.length) return;
  var sh = sheet_(name) || ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(), name);
  var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], T = typesOf_(name), stamp = nowIso_();
  var rows = objs.map(function (o) {
    if (!o.UPDATED_AT) o.UPDATED_AT = stamp;
    return H.map(function (h) { return toCell_(o[h], T[h]); });
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, H.length).setValues(rows);
}
function toCell_(v, t) {
  if (v === undefined || v === null) return '';
  if (t === 'date') {
    if (v instanceof Date) return v;
    if (/^\d{4}-\d{2}-\d{2}/.test(String(v))) { var p = String(v).substring(0, 10).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
    return v;
  }
  if (t === 'rp' || t === 'num' || t === 'coef' || t === 'pct') {
    if (v === '') return '';
    var n = Number(v);
    return isFinite(n) ? n : v;
  }
  if (t === 'json' && typeof v === 'object') return JSON.stringify(v);
  return v;
}
function fromCell_(v, t) {
  if (v instanceof Date) {
    if (t === 'ts') return v.toISOString();
    return Utilities.formatDate(v, APP.tz, 'yyyy-MM-dd');
  }
  return v;
}
function settingsMap_() {
  var m = {};
  readTable_('PENGATURAN').forEach(function (r) { m[r.ID] = r.NILAI; });
  return m;
}
function setSetting_(key, val) {
  var sh = sheet_('PENGATURAN'), data = sh.getDataRange().getValues(), H = data[0];
  var iId = H.indexOf('ID'), iN = H.indexOf('NILAI'), iU = H.indexOf('UPDATED_AT');
  for (var r = 1; r < data.length; r++) {
    if (data[r][iId] === key) {
      sh.getRange(r + 1, iN + 1).setValue(val);
      sh.getRange(r + 1, iU + 1).setValue(nowIso_());
      return;
    }
  }
  appendRows_('PENGATURAN', [{ ID: key, NILAI: val }]);
}
function log_(aksi, tabel, id, detail, sumber) {
  try {
    var sh = sheet_('LOG');
    if (!sh) return;
    sh.appendRow([Utilities.getUuid(), nowIso_(), aksi, tabel, id, String(detail).substring(0, 500), sumber, nowIso_(), '']);
  } catch (e) {}
}
