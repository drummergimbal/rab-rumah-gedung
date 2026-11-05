/* ============================================================
 * RAB Rumah & Gedung — PWA (by DG) · v1.2.0 · Tahap 1–3
 * RAB + AHSP/HSD + Master + Import + Profil + Penawaran + PDF ke Drive
 * ============================================================ */
'use strict';
const APP_VERSION = '1.2.0';
const TABLES = ['PROYEK', 'DIVISI', 'RAB_ITEM', 'VOLUME_DETAIL', 'REVISI', 'REVISI_ITEM', 'AHSP_HEADER', 'AHSP_KOMPONEN', 'AHSP_VERSI',
  'HSD_UPAH', 'HSD_BAHAN', 'HSD_ALAT', 'WILAYAH', 'PENGATURAN', 'PROFIL_USAHA', 'TEMPLATE_SYARAT', 'IMPORT_MAPPING',
  'PENAWARAN', 'PENAWARAN_ITEM', 'FILES', 'JADWAL', 'TENAGA_REKAP', 'BORONGAN_UPAH', 'TERMIN_UPAH', 'PROGRES', 'LIBUR', 'MILESTONE'];
const MASTER_TABLES = ['AHSP_HEADER', 'AHSP_KOMPONEN', 'HSD_UPAH', 'HSD_BAHAN', 'HSD_ALAT'];
const HSD_TABLE = { UPAH: 'HSD_UPAH', BAHAN: 'HSD_BAHAN', ALAT: 'HSD_ALAT' };
const THEME = {
  primary: '#1F4E79', primaryDark: '#0E2A45', primary50: '#EAF2FA', accent: '#F5A524', accent50: '#FFF4DD',
  zebra: '#F4F8FC', border: '#E3E8EF', text: '#1A2B3C', muted: '#5B6B7C',
  success: '#1E9E6A', warning: '#E08A00', danger: '#D64545', info: '#2F7DD1'
};
const STATUS_CLASS = {
  'TERVERIFIKASI': 'success', 'OK': 'success', 'DITERIMA': 'success', 'SELESAI': 'success', 'RESMI': 'success',
  'PERLU CEK': 'warning', 'PERLU HARGA': 'warning', 'DRAFT': 'warning', 'NEGOSIASI': 'warning',
  'DITOLAK': 'danger', 'KEDALUWARSA': 'danger', 'BATAL': 'danger',
  'DISETUJUI': 'success', 'TERCAPAI': 'success', 'SESUAI': 'success', 'LEBIH CEPAT': 'success', 'WASPADA': 'warning', 'TERLAMBAT': 'danger', 'RENCANA': 'info', 'TERKIRIM': 'info', 'BERJALAN': 'info', 'PENAWARAN': 'info', 'MODIFIKASI': 'primary', 'CUSTOM': 'accent', 'HARGA PASAR': 'accent'
};
const DEFAULTS = {
  OVERHEAD_PCT: 10, PPN_PCT: 11, KONTINJENSI_PCT: 5, PEMBULATAN: 1000, WILAYAH_DEFAULT: 'Kabupaten Bogor',
  VERSI_AHSP_AKTIF: 'CONTOH', MODE_DEFAULT: 'RUMAH', AMBANG_DEVIASI_HSP: 30, FORMAT_NOMOR_PENAWARAN: '{NO}/PNW/RAB/{BLN}/{THN}',
  MASA_BERLAKU_HARI: 14, JAM_KERJA_HARI: 8, HARI_KERJA_MINGGU: 6, RASIO_KEPALA_TUKANG: 10, RASIO_MANDOR_PEKERJA: 20, DEVIASI_WASPADA: 5, DEVIASI_KRITIS: 10, TERMIN_DEFAULT: 'DP 30% – Progres 30% – Progres 30% – Retensi 10%', TAMPIL_CREDIT: 'YA'
};
const VENDOR = {
  jspdf: 'vendor/jspdf.umd.min.js', autotable: 'vendor/jspdf.plugin.autotable.min.js', xlsx: 'vendor/xlsx.full.min.js'
};

/* ---------- Utils ---------- */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
const esc = s => String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = (p = '') => p + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 7).toUpperCase();
const nowIso = () => new Date().toISOString();
const todayStr = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
const num = Calc.num;
const N = (n, d = 0) => new Intl.NumberFormat('id-ID', { minimumFractionDigits: d, maximumFractionDigits: d }).format(num(n));
const Nmax = (n, d = 4) => new Intl.NumberFormat('id-ID', { maximumFractionDigits: d }).format(num(n));
const rp = n => { const v = Math.round(num(n)); return (v < 0 ? '-Rp ' : 'Rp ') + N(Math.abs(v)); };
const rpCell = n => { const v = Math.round(num(n)); return `<span class="rp-cell${v < 0 ? ' neg' : ''}"><span class="cur">Rp</span><span>${v === 0 ? '-' : (v < 0 ? '-' : '') + N(Math.abs(v))}</span></span>`; };
const pct = (n, d = 2) => N(n, d) + '%';
const fmtDate = s => { if (!s) return '-'; const p = String(s).slice(0, 10).split('-'); return p.length === 3 ? `${p[2]}/${p[1]}/${p[0]}` : s; };
const ROMAWI = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
const BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const fmtDateLong = s => { const p = String(s || todayStr()).slice(0, 10).split('-'); return `${+p[2]} ${BULAN[+p[1] - 1]} ${p[0]}`; };
const norm = Calc.norm;
/* parse angka input Indonesia: "1.250.000" / "1250000" / "2,5" / "2.5" */
function parseNum(s, money) {
  if (typeof s === 'number') return s;
  s = String(s || '').trim().replace(/\s|Rp/gi, '');
  if (!s) return '';
  if (money) { s = s.replace(/\./g, '').replace(',', '.'); }
  else if (s.indexOf(',') >= 0) { s = s.replace(/\./g, '').replace(',', '.'); }
  else if ((s.match(/\./g) || []).length > 1) { s = s.replace(/\./g, ''); }
  const n = Number(s);
  return isFinite(n) ? n : '';
}
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
function badge(text, cls) { return `<span class="badge ${cls || STATUS_CLASS[norm(text)] || ''}">${esc(text)}</span>`; }
const ICON = {
  plus: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>',
  search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>',
  close: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12M18 6 6 18"/></svg>',
  up: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="m6 15 6-6 6 6"/></svg>',
  down: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="m6 9 6 6 6-6"/></svg>',
  back: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="m15 18-6-6 6-6"/></svg>',
  warn: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>',
  pdf: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M12 12v6M9 15l3 3 3-3"/></svg>',
  share: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><path d="m8.6 13.5 6.8 4M15.4 6.5l-6.8 4"/></svg>',
  calc: '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="4" y="2" width="16" height="20" rx="2"/><path d="M8 6h8M8 11h.01M12 11h.01M16 11h.01M8 15h.01M12 15h.01M16 15h.01M8 19h8"/></svg>'
};

/* ---------- Toast & modal ---------- */
function toast(msg, type = '') {
  const el = document.createElement('div');
  el.className = 'toast ' + type; el.textContent = msg;
  $('#toasts').appendChild(el);
  setTimeout(() => el.remove(), type === 'err' ? 6000 : 3200);
}
let modalStack = [];
function openModal({ title, body, foot = '', wide = false, onOpen, onClose }) {
  const back = document.createElement('div');
  back.className = 'modal-back';
  back.innerHTML = `<div class="modal ${wide ? 'wide' : ''}" role="dialog" aria-modal="true">
    <div class="modal-head"><h3>${esc(title)}</h3><button class="btn ghost icon" data-close aria-label="Tutup">${ICON.close}</button></div>
    <div class="modal-body">${body}</div>${foot ? `<div class="modal-foot">${foot}</div>` : ''}</div>`;
  const m = { el: back, close: () => { back.remove(); modalStack = modalStack.filter(x => x !== m); onClose && onClose(); } };
  back.addEventListener('click', e => { if (e.target === back || e.target.closest('[data-close]')) m.close(); });
  $('#modalRoot').appendChild(back);
  modalStack.push(m);
  onOpen && onOpen(back, m);
  const f = back.querySelector('input:not([type=hidden]):not([readonly]), select, textarea');
  if (f && window.innerWidth > 760) setTimeout(() => f.focus(), 50);
  return m;
}
function confirmBox(msg, { ok = 'Ya, lanjutkan', danger = false, title = 'Konfirmasi' } = {}) {
  return new Promise(res => {
    let done = false;
    openModal({
      title, body: `<p style="margin:0;white-space:pre-line">${esc(msg)}</p>`,
      foot: `<button class="btn" data-close>Batal</button><button class="btn ${danger ? 'primary' : 'primary'}" data-ok style="${danger ? 'background:var(--danger);border-color:var(--danger)' : ''}">${esc(ok)}</button>`,
      onOpen: (el, m) => { $('[data-ok]', el).onclick = () => { done = true; m.close(); res(true); }; },
      onClose: () => { if (!done) res(false); }
    });
  });
}
function formData(el) {
  const o = {};
  $$('[name]', el).forEach(i => {
    if (i.type === 'checkbox') o[i.name] = i.checked ? 'YA' : 'TIDAK';
    else if (i.dataset.num !== undefined) o[i.name] = parseNum(i.value, i.dataset.num === 'rp');
    else o[i.name] = i.value.trim();
  });
  return o;
}
function inp(name, label, val, { type = 'text', numType, hint = '', ph = '', cls = '', req = false, list = '', ro = false } = {}) {
  const isNum = numType !== undefined;
  let v = val === undefined || val === null ? '' : val;
  if (isNum && v !== '') v = numType === 'rp' ? N(v) : String(v).replace('.', ',');
  return `<div class="field ${cls}"><label>${esc(label)}${req ? ' *' : ''}</label>
    <input class="input ${isNum ? 'num' : ''}" name="${name}" type="${type}" value="${esc(v)}" placeholder="${esc(ph)}" ${isNum ? `inputmode="decimal" data-num="${numType}"` : ''} ${list ? `list="${list}"` : ''} ${req ? 'required' : ''} ${ro ? 'readonly' : ''}>
    ${hint ? `<span class="hint">${hint}</span>` : ''}</div>`;
}
function sel(name, label, val, opts, { cls = '', hint = '' } = {}) {
  return `<div class="field ${cls}"><label>${esc(label)}</label><select class="input" name="${name}">${opts.map(o => {
    const [v, t] = Array.isArray(o) ? o : [o, o];
    return `<option value="${esc(v)}" ${String(v) === String(val) ? 'selected' : ''}>${esc(t)}</option>`;
  }).join('')}</select>${hint ? `<span class="hint">${hint}</span>` : ''}</div>`;
}
function txt(name, label, val, { cls = '', rows = 4, hint = '' } = {}) {
  return `<div class="field ${cls}"><label>${esc(label)}</label><textarea class="input" name="${name}" rows="${rows}">${esc(val || '')}</textarea>${hint ? `<span class="hint">${hint}</span>` : ''}</div>`;
}
function bindMoneyInputs(root) {
  $$('input[data-num="rp"]', root).forEach(i => i.addEventListener('blur', () => { const v = parseNum(i.value, true); i.value = v === '' ? '' : N(v); }));
}

/* ---------- IndexedDB ---------- */
const DB = (() => {
  let db;
  const open = () => new Promise((res, rej) => {
    const r = indexedDB.open('rab-rumah-gedung', 1);
    r.onupgradeneeded = () => { const d = r.result; ['tables', 'meta', 'blobs'].forEach(s => d.objectStoreNames.contains(s) || d.createObjectStore(s)); };
    r.onsuccess = () => { db = r.result; res(); };
    r.onerror = () => rej(r.error);
  });
  const tx = (store, mode, fn) => new Promise((res, rej) => {
    const t = db.transaction(store, mode), s = t.objectStore(store);
    let out; const q = fn(s);
    if (q) q.onsuccess = () => { out = q.result; };
    t.oncomplete = () => res(out); t.onerror = () => rej(t.error);
  });
  return {
    open, get: (store, k) => tx(store, 'readonly', s => s.get(k)),
    set: (store, k, v) => tx(store, 'readwrite', s => s.put(v, k)),
    del: (store, k) => tx(store, 'readwrite', s => s.delete(k)),
    keys: store => tx(store, 'readonly', s => s.getAllKeys()),
    clear: store => tx(store, 'readwrite', s => s.clear())
  };
})();

/* ---------- State & data ---------- */
const S = { t: {}, queue: {}, uploads: [], meta: { lastSync: '', apiUrl: '', token: '', activeProject: '' }, idx: null, syncing: false, online: navigator.onLine, lastError: '' };
TABLES.forEach(t => S.t[t] = []);

const dirtyTables = new Set();
const flushTables = debounce(async () => {
  const names = Array.from(dirtyTables); dirtyTables.clear();
  for (const n of names) await DB.set('tables', n, S.t[n]);
  await DB.set('meta', 'queue', S.queue);
  await DB.set('meta', 'uploads', S.uploads);
}, 300);
function markDirty(t) { dirtyTables.add(t); if (MASTER_TABLES.includes(t) || t === 'RAB_ITEM') S.idx = null; flushTables(); }
async function saveMeta() { await DB.set('meta', 'meta', S.meta); }

function rows(t, includeDeleted = false) { return (S.t[t] || []).filter(r => includeDeleted || !Calc.isDel(r)); }
function getRow(t, id) { return (S.t[t] || []).find(r => r.ID === id); }
/* simpan baris lokal + masuk antrian sync */
function save(t, row, { silent = false } = {}) {
  const arr = S.t[t];
  const i = arr.findIndex(r => r.ID === row.ID);
  const base = i >= 0 ? arr[i].UPDATED_AT : '';
  const merged = Object.assign({}, i >= 0 ? arr[i] : {}, row, { UPDATED_AT: nowIso() });
  if (!merged.ID) merged.ID = uid();
  if (i >= 0) arr[i] = merged; else arr.push(merged);
  const key = t + '|' + merged.ID;
  const prev = S.queue[key];
  S.queue[key] = { table: t, row: merged, base: prev ? prev.base : base };
  markDirty(t);
  if (!silent) scheduleSync();
  return merged;
}
function saveMany(t, list) { list.forEach(r => save(t, r, { silent: true })); scheduleSync(); }
function remove(t, id) { const r = getRow(t, id); if (r) save(t, { ID: id, DELETED: 'YA' }); }
function idx() { if (!S.idx) S.idx = Calc.buildIndex(S.t); return S.idx; }
function setting(k) {
  const r = getRow('PENGATURAN', k);
  const v = r && r.NILAI !== '' && r.NILAI !== undefined ? r.NILAI : DEFAULTS[k];
  return v === undefined ? '' : v;
}
function settingsObj() { const o = Object.assign({}, DEFAULTS); rows('PENGATURAN').forEach(r => { if (r.NILAI !== '') o[r.ID] = r.NILAI; }); return o; }
function setSetting(k, v) { save('PENGATURAN', { ID: k, NILAI: v }); }
function pendingCount() { return Object.keys(S.queue).length + S.uploads.length; }

/* ---------- Sinkronisasi ---------- */
async function api(action, payload = {}) {
  if (!S.meta.apiUrl || !S.meta.token) throw new Error('Koneksi belum diatur (Master → Pengaturan)');
  const res = await fetch(S.meta.apiUrl, {
    method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, redirect: 'follow',
    body: JSON.stringify(Object.assign({ action, token: S.meta.token }, payload))
  });
  const txt = await res.text();
  let j; try { j = JSON.parse(txt); } catch (e) { throw new Error('Respon server tidak valid. Pastikan URL Web App benar & akses "Anyone".'); }
  if (!j.ok) throw new Error(j.error || 'Gagal');
  return j;
}
const scheduleSync = debounce(() => sync(), 2500);
function applyServerRow(t, row) {
  const arr = S.t[t] || (S.t[t] = []);
  const key = t + '|' + row.ID;
  const q = S.queue[key];
  if (q && q.row.UPDATED_AT > (q.sentAt || '')) return; // ada perubahan lokal yang lebih baru, tunggu push berikutnya
  const i = arr.findIndex(r => r.ID === row.ID);
  if (Calc.isDel(row)) { if (i >= 0) arr.splice(i, 1); }
  else if (i >= 0) arr[i] = row; else arr.push(row);
  markDirty(t);
}
async function sync({ full = false, quiet = true } = {}) {
  if (S.syncing) return;
  if (!S.meta.apiUrl || !S.meta.token) { renderSync(); return; }
  if (!navigator.onLine) { S.online = false; renderSync(); return; }
  S.syncing = true; renderSync();
  try {
    // 1) upload file tertunda
    await processUploads();
    // 2) push perubahan
    const keys = Object.keys(S.queue);
    for (let i = 0; i < keys.length; i += 200) {
      const chunk = keys.slice(i, i + 200).filter(k => S.queue[k]);
      const sentAt = nowIso();
      const changes = chunk.map(k => { S.queue[k].sentAt = sentAt; return { table: S.queue[k].table, row: Object.assign({}, S.queue[k].row, { BASE_UPDATED_AT: S.queue[k].base || '' }) }; });
      const r = await api('push', { changes });
      chunk.forEach(k => { const q = S.queue[k]; if (q && q.row.UPDATED_AT <= q.sentAt) delete S.queue[k]; });
      (r.applied || []).forEach(a => applyServerRow(a.table, a.row));
      (r.conflicts || []).forEach(c => { delete S.queue[c.table + '|' + c.id]; applyServerRow(c.table, c.server); });
      if ((r.conflicts || []).length) toast(`${r.conflicts.length} data diperbarui dari Sheet (versi Sheet lebih baru)`, '');
    }
    // 3) pull
    const since = full ? '' : S.meta.lastSync;
    const p = await api('pull', { since });
    if (full) { TABLES.forEach(t => { if (p.tables[t]) { const pend = Object.values(S.queue).filter(q => q.table === t).map(q => q.row); S.t[t] = p.tables[t].filter(r => !Calc.isDel(r)); pend.forEach(r => applyLocal(t, r)); markDirty(t); } }); }
    else Object.keys(p.tables || {}).forEach(t => { if (S.t[t]) p.tables[t].forEach(r => applyServerRow(t, r)); });
    S.meta.lastSync = p.serverTime;
    S.lastError = ''; S.online = true;
    await saveMeta(); flushTables();
    if (!quiet) toast('Sinkron selesai', 'ok');
    if (!CUR.dirty) rerenderSoft(); // jangan hapus isian yang sedang diketik
    else CUR.pendingRefresh = true;
  } catch (e) {
    S.lastError = e.message;
    if (!quiet) toast('Sinkron gagal: ' + e.message, 'err');
  } finally { S.syncing = false; renderSync(); }
}
function applyLocal(t, r) { const arr = S.t[t]; const i = arr.findIndex(x => x.ID === r.ID); if (i >= 0) arr[i] = r; else arr.push(r); }
function renderSync() {
  const chip = $('#syncChip'); if (!chip) return;
  const n = pendingCount();
  let cls = 'ok', lbl = 'Tersinkron';
  if (!S.meta.apiUrl || !S.meta.token) { cls = 'off'; lbl = 'Belum terhubung'; }
  else if (S.syncing) { cls = 'busy'; lbl = 'Menyinkron…'; }
  else if (!navigator.onLine) { cls = 'off'; lbl = n ? `Offline · ${n} tertunda` : 'Offline'; }
  else if (S.lastError) { cls = 'off'; lbl = n ? `Gagal · ${n} tertunda` : 'Gagal sinkron'; }
  else if (n) { cls = 'pending'; lbl = `${n} menunggu`; }
  chip.className = 'sync-chip ' + cls;
  chip.querySelector('.lbl').textContent = lbl;
  chip.title = S.lastError || (S.meta.lastSync ? 'Sinkron terakhir: ' + new Date(S.meta.lastSync).toLocaleString('id-ID') : 'Ketuk untuk sinkron');
}

/* ---------- Upload file ke Drive (antrian offline) ---------- */
function blobToBase64(blob) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = rej; r.readAsDataURL(blob); });
}
async function queueUpload({ blob, name, mime, kategori, proyekId, jenis, terkait, onDone }) {
  const key = uid('UP');
  await DB.set('blobs', key, blob);
  S.uploads.push({ key, name, mime, kategori, proyekId, jenis, terkait, onDone, createdAt: nowIso() });
  await DB.set('meta', 'uploads', S.uploads);
  renderSync();
  if (navigator.onLine && S.meta.apiUrl) sync();
  return key;
}
const UPLOAD_HOOKS = {};
async function processUploads() {
  while (S.uploads.length) {
    const u = S.uploads[0];
    const blob = await DB.get('blobs', u.key);
    if (!blob) { S.uploads.shift(); continue; }
    const r = await api('upload', { kategori: u.kategori, proyekId: u.proyekId, name: u.name, mime: u.mime, jenis: u.jenis, terkait: u.terkait, base64: await blobToBase64(blob) });
    if (r.record) applyServerRow('FILES', r.record);
    if (u.onDone && UPLOAD_HOOKS[u.onDone.type]) UPLOAD_HOOKS[u.onDone.type](u.onDone, r);
    await DB.del('blobs', u.key);
    S.uploads.shift();
    await DB.set('meta', 'uploads', S.uploads);
  }
}
/* setelah upload: tulis file id ke baris terkait */
UPLOAD_HOOKS.setField = (d, r) => { save(d.table, { ID: d.id, [d.field]: r.fileId }); if (d.urlField) save(d.table, { ID: d.id, [d.urlField]: r.url }); };

/* Ambil gambar (logo/ttd) dari cache lokal atau Drive */
async function imageData(fileId) {
  if (!fileId) return '';
  const cached = await DB.get('blobs', 'IMG_' + fileId);
  if (cached) return cached;
  if (String(fileId).startsWith('LOCAL_')) return '';
  try {
    const r = await api('getFile', { fileId });
    const d = `data:${r.mime};base64,${r.base64}`;
    await DB.set('blobs', 'IMG_' + fileId, d);
    return d;
  } catch (e) { return ''; }
}

function loadScript(src) {
  return new Promise((res, rej) => {
    if ($(`script[src="${src}"]`)) return res();
    const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Gagal memuat ' + src));
    document.head.appendChild(s);
  });
}
async function needPdf() { await loadScript(VENDOR.jspdf); await loadScript(VENDOR.autotable); return window.jspdf.jsPDF; }
async function needXlsx() { await loadScript(VENDOR.xlsx); return window.XLSX; }
function downloadBlob(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}
function readFile(file, as = 'text') {
  return new Promise((res, rej) => {
    const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej;
    if (as === 'buffer') r.readAsArrayBuffer(file); else if (as === 'dataurl') r.readAsDataURL(file); else r.readAsText(file);
  });
}
/* kompres gambar → dataURL (maks lebar, JPEG/PNG) */
function compressImage(file, maxW = 1600, quality = 0.8, keepPng = false) {
  return new Promise(async (res, rej) => {
    try {
      const url = await readFile(file, 'dataurl');
      if (/svg/.test(file.type)) return res({ dataUrl: url, w: 1000, h: 1000, mime: 'image/svg+xml' });
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxW / img.width);
        const c = document.createElement('canvas'); c.width = Math.round(img.width * scale); c.height = Math.round(img.height * scale);
        const ctx = c.getContext('2d');
        const png = keepPng || /png/.test(file.type);
        if (!png) { ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); }
        ctx.drawImage(img, 0, 0, c.width, c.height);
        const mime = png ? 'image/png' : 'image/jpeg';
        res({ dataUrl: c.toDataURL(mime, quality), w: img.width, h: img.height, mime });
      };
      img.onerror = () => rej(new Error('Gambar tidak bisa dibaca'));
      img.src = url;
    } catch (e) { rej(e); }
  });
}
function dataUrlToBlob(d) {
  const [h, b] = d.split(','); const mime = h.match(/:(.*?);/)[1]; const bin = atob(b);
  const arr = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return new Blob([arr], { type: mime });
}
/* ============================================================
 * Domain: proyek, item RAB, hitung ulang, template, AHSP
 * ============================================================ */
const projects = () => rows('PROYEK').sort((a, b) => String(b.UPDATED_AT).localeCompare(String(a.UPDATED_AT)));
const projItems = pid => rows('RAB_ITEM').filter(r => r.PROYEK_ID === pid).sort((a, b) => num(a.URUTAN) - num(b.URUTAN));
const projDivs = pid => projItems(pid).filter(r => norm(r.TIPE_BARIS) === 'DIVISI');
const divItems = (pid, did) => projItems(pid).filter(r => r.PARENT_ID === did && norm(r.TIPE_BARIS) !== 'DIVISI');
const hsdAll = () => [].concat(rows('HSD_UPAH').map(r => Object.assign({ _t: 'HSD_UPAH', _j: 'Upah' }, r)), rows('HSD_BAHAN').map(r => Object.assign({ _t: 'HSD_BAHAN', _j: 'Bahan' }, r)), rows('HSD_ALAT').map(r => Object.assign({ _t: 'HSD_ALAT', _j: 'Alat' }, r)));
const ahspKomp = kode => rows('AHSP_KOMPONEN').filter(r => norm(r.KODE_AHSP) === norm(kode)).sort((a, b) => num(a.URUTAN) - num(b.URUTAN));
const ahspHead = kode => rows('AHSP_HEADER').find(r => norm(r.KODE_AHSP) === norm(kode));
const isOfficial = h => h && norm(h.JENIS) === 'RESMI';
const ctxOf = p => Calc.ctxFromProject(p, settingsObj());
function rekapOf(p) { return Calc.rekap(idx(), p, rows('RAB_ITEM'), settingsObj()); }
function profilOf(id) { return getRow('PROFIL_USAHA', id) || rows('PROFIL_USAHA').find(r => norm(r.IS_DEFAULT) === 'YA') || rows('PROFIL_USAHA')[0] || null; }
function nextProjectCode() {
  const n = rows('PROYEK', true).map(p => parseInt(String(p.KODE || '').replace(/\D/g, ''), 10) || 0);
  return 'PRJ' + String((n.length ? Math.max(...n) : 0) + 1).padStart(3, '0');
}
function divisiMaster(mode) {
  return rows('DIVISI').filter(d => mode === 'GEDUNG' || norm(d.MODE) !== 'GEDUNG').sort((a, b) => num(a.URUTAN) - num(b.URUTAN));
}

/* Hitung ulang proyek → simpan HS, jumlah, bobot, subtotal, kode urut, total */
function recompute(pid) {
  const p = getRow('PROYEK', pid); if (!p) return null;
  const rk = rekapOf(p);
  const changed = [];
  let order = 0;
  rk.groups.forEach(g => {
    if (g.row.ID === '_') return;
    order += 1;
    const d = g.row;
    const dKode = d.KODE || ROMAWI[order - 1] || String(order);
    const dPatch = { ID: d.ID, URUTAN: order * 1000, JUMLAH: Math.round(g.subtotal), BOBOT: Calc.round(g.bobot, 4), UPAH: Math.round(g.upah), BAHAN: Math.round(g.bahan), ALAT: Math.round(g.alat), KODE: dKode };
    if (diff(d, dPatch)) changed.push(dPatch);
    g.items.forEach((it, i) => {
      const c = rk.computed[it.ID];
      const patch = { ID: it.ID, URUTAN: order * 1000 + i + 1, KODE: dKode + '.' + (i + 1), HARGA_SATUAN: c.hs, JUMLAH: c.jumlah, UPAH: Math.round(c.upah), BAHAN: Math.round(c.bahan), ALAT: Math.round(c.alat), BOBOT: Calc.round(c.bobot, 4) };
      if (diff(it, patch)) changed.push(patch);
    });
  });
  changed.forEach(c => save('RAB_ITEM', c, { silent: true }));
  if (Math.round(num(p.TOTAL_RAB)) !== Math.round(rk.total)) save('PROYEK', { ID: pid, TOTAL_RAB: rk.total }, { silent: true });
  if (changed.length) scheduleSync();
  return rekapOf(getRow('PROYEK', pid));
}
function diff(a, patch) { return Object.keys(patch).some(k => k !== 'ID' && String(a[k] ?? '') !== String(patch[k] ?? '')); }

const WARN_TEXT = {
  TANPA_AHSP: ['Belum terhubung AHSP', 'warning'], HARGA_KOSONG: ['Harga kosong', 'danger'], AHSP_TIDAK_ADA: ['AHSP tidak ditemukan', 'danger'],
  HSD_KOSONG: ['Ada harga dasar kosong', 'danger'], HSD_FALLBACK: ['Harga dasar dari wilayah/tahun lain', 'warning'], PERLU_CEK: ['AHSP perlu cek', 'warning']
};

/* ---------- Template RAB ----------
 * Daftar uraian standar per divisi. AHSP dicocokkan otomatis dari library (bisa dikonfirmasi). */
const TPL_ITEMS = {
  'I': [['Pembersihan lahan', 'm2'], ['Pengukuran & pemasangan bouwplank', "m'"], ['Direksi keet / gudang material', 'ls'], ['Air & listrik kerja', 'ls'], ['Pagar sementara', "m'"], ['Mobilisasi & demobilisasi', 'ls'], ['Papan nama proyek', 'bh'], ['Dokumentasi & administrasi', 'ls']],
  'II': [['Penyiapan RKK', 'ls'], ['Sosialisasi, promosi & pelatihan K3', 'ls'], ['Alat pelindung diri (APD) & alat pelindung kerja (APK)', 'ls'], ['Asuransi & perizinan', 'ls'], ['Personel K3 konstruksi', 'ls'], ['Fasilitas sarana kesehatan (P3K)', 'ls'], ['Rambu-rambu keselamatan', 'ls'], ['Konsultasi dengan ahli K3', 'ls'], ['Kegiatan lain pengendalian risiko K3', 'ls']],
  'III': [['Galian tanah pondasi', 'm3'], ['Urugan tanah kembali', 'm3'], ['Urugan pasir bawah pondasi', 'm3'], ['Urugan pasir bawah lantai', 'm3'], ['Urugan tanah peninggi lantai', 'm3'], ['Pemadatan tanah', 'm2'], ['Buang tanah sisa galian', 'm3']],
  'IV': [['Aanstamping batu kosong', 'm3'], ['Pondasi batu kali 1PC:4PP', 'm3'], ['Pondasi footplat beton bertulang', 'm3'], ['Sloof beton bertulang', 'm3'], ['Bore pile / strauss pile', "m'"]],
  'V': [['Bekisting kolom', 'm2'], ['Pembesian kolom', 'kg'], ['Beton kolom fc 20 MPa', 'm3'], ['Bekisting balok', 'm2'], ['Pembesian balok', 'kg'], ['Beton balok fc 20 MPa', 'm3'], ['Bekisting plat lantai', 'm2'], ['Pembesian plat lantai', 'kg'], ['Beton plat lantai fc 20 MPa', 'm3'], ['Ring balk beton bertulang', 'm3'], ['Tangga beton bertulang', 'm3']],
  'VI': [['Rangka atap baja ringan', 'm2'], ['Penutup atap genteng', 'm2'], ['Bubungan / nok', "m'"], ['Lisplang', "m'"], ['Talang', "m'"], ['Waterproofing dak beton', 'm2']],
  'VII': [['Pasangan dinding bata merah 1PC:4PP', 'm2'], ['Pasangan dinding bata ringan (hebel)', 'm2'], ['Kolom praktis 11×11', "m'"], ['Balok praktis 11×11', "m'"], ['Plesteran 1PC:4PP tebal 15 mm', 'm2'], ['Acian', 'm2'], ['Benangan / sudut', "m'"]],
  'VIII': [['Lantai kerja beton tumbuk', 'm3'], ['Screed lantai', 'm2'], ['Keramik lantai 60×60', 'm2'], ['Keramik lantai KM 25×25', 'm2'], ['Keramik dinding KM 25×40', 'm2'], ['Plint keramik', "m'"]],
  'IX': [['Rangka plafon hollow', 'm2'], ['Penutup plafon gypsum 9 mm', 'm2'], ['Plafon GRC teras', 'm2'], ['List profil plafon', "m'"]],
  'X': [['Kusen aluminium', "m'"], ['Daun pintu panel', 'bh'], ['Daun jendela kaca', 'bh'], ['Kaca 5 mm', 'm2'], ['Kunci pintu & handle', 'set'], ['Engsel pintu & jendela', 'bh']],
  'XI': [['Kloset duduk', 'bh'], ['Wastafel', 'bh'], ['Shower set', 'bh'], ['Kran air', 'bh'], ['Floor drain', 'bh'], ['Kitchen sink', 'bh']],
  'XII': [['Pipa air bersih PVC/PPR ½"–¾"', "m'"], ['Fitting & aksesoris air bersih', 'ls'], ['Pompa air', 'unit'], ['Tandon / toren air', 'unit'], ['Pipa air kotor PVC 4"', "m'"], ['Pipa air bekas PVC 2"–3"', "m'"], ['Pipa vent', "m'"], ['Bak kontrol', 'bh'], ['Septic tank / biofilter', 'unit'], ['Sumur resapan', 'unit'], ['Talang tegak air hujan', "m'"], ['Saluran drainase keliling', "m'"]],
  'XIII': [['AC split (unit & instalasi)', 'unit'], ['Instalasi pipa refrigerant & drain AC', "m'"], ['Exhaust fan', 'bh'], ['Pompa booster', 'unit']],
  'XIV': [['Penyambungan daya PLN', 'ls'], ['Panel listrik & MCB', 'unit'], ['Kabel feeder', "m'"], ['Titik lampu', 'titik'], ['Titik stop kontak', 'titik'], ['Titik saklar', 'titik'], ['Armatur lampu', 'bh'], ['Grounding / pembumian', 'titik'], ['Penangkal petir', 'ls']],
  'XV': [['Titik data / LAN', 'titik'], ['Titik CCTV', 'titik'], ['Bel / intercom', 'titik'], ['Titik TV', 'titik']],
  'XVI': [['APAR', 'unit'], ['Hidran box', 'unit'], ['Instalasi sprinkler', 'titik'], ['Detektor & fire alarm', 'titik'], ['Pompa kebakaran', 'unit']],
  'XVII': [['Plamir dinding', 'm2'], ['Cat dinding interior', 'm2'], ['Cat dinding eksterior', 'm2'], ['Cat plafon', 'm2'], ['Cat kayu / besi', 'm2'], ['Waterproofing KM', 'm2'], ['Nat keramik', 'm2'], ['Pembersihan akhir', 'ls']],
  'XVIII': [['Lantai carport', 'm2'], ['Paving block', 'm2'], ['Pagar & gerbang', "m'"], ['Taman', 'm2'], ['Saluran luar', "m'"]],
  'XIX': []
};
const TEMPLATES = {
  RUMAH1: { nama: 'Rumah 1 lantai', mode: 'RUMAH', skip: ['V:Bekisting plat lantai', 'V:Pembesian plat lantai', 'V:Beton plat lantai fc 20 MPa', 'V:Tangga beton bertulang', 'XIII:Pompa booster', 'IV:Bore pile / strauss pile'], divisi: ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVII', 'XVIII'] },
  RUMAH2: { nama: 'Rumah 2 lantai', mode: 'RUMAH', skip: ['XIII:Pompa booster'], divisi: ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVII', 'XVIII'] },
  RUKO: { nama: 'Ruko', mode: 'RUMAH', skip: ['XVIII:Taman'], divisi: ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII'] },
  GEDUNG: { nama: 'Gedung', mode: 'GEDUNG', skip: [], divisi: ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX'] }
};
function applyTemplate(pid, key) {
  const tpl = TEMPLATES[key]; if (!tpl) return;
  const master = rows('DIVISI');
  let urut = (projItems(pid).length ? Math.max(...projItems(pid).map(r => num(r.URUTAN))) : 0) + 1000;
  const list = [];
  tpl.divisi.forEach(code => {
    const m = master.find(d => d.KODE === code) || { KODE: code, NAMA: 'DIVISI ' + code };
    const did = uid('D');
    list.push({ ID: did, PROYEK_ID: pid, TIPE_BARIS: 'DIVISI', KODE: m.KODE, URAIAN: m.NAMA, URUTAN: urut });
    (TPL_ITEMS[code] || []).forEach((it, i) => {
      if (tpl.skip.includes(code + ':' + it[0])) return;
      list.push({ ID: uid('I'), PROYEK_ID: pid, PARENT_ID: did, TIPE_BARIS: 'ITEM', URAIAN: it[0], SATUAN: it[1], VOLUME: 0, JENIS_HARGA: 'AHSP', KODE_AHSP: '', URUTAN: urut + i + 1 });
    });
    urut += 1000;
  });
  saveMany('RAB_ITEM', list);
  recompute(pid);
}

/* Cocokkan AHSP otomatis untuk item yang belum terhubung */
function autoMatchSuggestions(pid) {
  const I = idx();
  const divs = {}; projDivs(pid).forEach(d => divs[d.ID] = d);
  return projItems(pid).filter(r => norm(r.TIPE_BARIS) !== 'DIVISI' && !r.KODE_AHSP && norm(r.JENIS_HARGA) !== 'PASAR').map(it => {
    const res = Calc.search(I, it.URAIAN, { satuan: it.SATUAN, divisi: (divs[it.PARENT_ID] || {}).URAIAN, limit: 3 });
    const best = res[0];
    if (!(best && best.score >= 4 && norm(best.header.SATUAN) === norm(it.SATUAN))) return null;
    // kuat bila ≥60% kata kunci uraian item ada di uraian AHSP
    const own = Calc.tokens(it.URAIAN).filter(t => t.length > 2), ur = Calc.tokens(best.header.URAIAN);
    const hit = own.filter(t => ur.some(u => u === t || (t.length > 3 && u.indexOf(t) === 0))).length;
    return { item: it, best, alt: res.slice(1), strong: own.length ? hit / own.length >= 0.6 : false };
  }).filter(Boolean);
}

/* ---------- AHSP custom / modifikasi ---------- */
function nextCustomCode(base) {
  const all = rows('AHSP_HEADER', true).map(h => h.KODE_AHSP);
  if (base) { let n = 1; while (all.includes(base + '-C' + String(n).padStart(2, '0'))) n++; return base + '-C' + String(n).padStart(2, '0'); }
  let n = 1; while (all.includes('CUS-' + String(n).padStart(3, '0'))) n++; return 'CUS-' + String(n).padStart(3, '0');
}
function duplicateAhsp(kode, { proyekId = '' } = {}) {
  const h = ahspHead(kode); if (!h) return null;
  const nk = nextCustomCode(h.KODE_AHSP);
  const nh = Object.assign({}, h, {
    ID: uid('H'), KODE_AHSP: nk, JENIS: 'MODIFIKASI', KODE_REFERENSI: h.KODE_AHSP, SUMBER: 'Modifikasi dari ' + h.KODE_AHSP + (h.SUMBER ? ' (' + h.SUMBER + ')' : ''),
    STATUS: 'PERLU CEK', SCOPE: proyekId ? 'PROYEK' : 'UMUM', PROYEK_ID: proyekId, DIBUAT_OLEH: 'PWA', TANGGAL: todayStr(), VERSI: 'CUSTOM', UPDATED_AT: '', DELETED: ''
  });
  save('AHSP_HEADER', nh, { silent: true });
  ahspKomp(kode).forEach(k => save('AHSP_KOMPONEN', Object.assign({}, k, { ID: uid('K'), KODE_AHSP: nk, UPDATED_AT: '', DELETED: '' }), { silent: true }));
  scheduleSync();
  return nk;
}
function ahspDeviation(kode, p) {
  const ctx = p ? ctxOf(p) : { wilayah: setting('WILAYAH_DEFAULT'), tahun: new Date().getFullYear(), overhead: num(setting('OVERHEAD_PCT')) };
  const a = Calc.analyze(idx(), kode, ctx);
  const sim = Calc.similar(idx(), kode, ctx, 3);
  const dev = sim.avg ? (a.HSP - sim.avg) / sim.avg * 100 : 0;
  return { a, sim, dev, over: sim.avg && Math.abs(dev) > num(setting('AMBANG_DEVIASI_HSP')) };
}
function libraryHealth() {
  const out = { tanpaSumber: [], hsdKosong: [], hsdLama: [], perluCek: [], tanpaKomponen: [] };
  const komp = {}; rows('AHSP_KOMPONEN').forEach(k => komp[norm(k.KODE_AHSP)] = 1);
  const year = new Date();
  rows('AHSP_HEADER').forEach(h => {
    if (!String(h.SUMBER || '').trim()) out.tanpaSumber.push(h);
    if (norm(h.STATUS) === 'PERLU CEK') out.perluCek.push(h);
    if (!komp[norm(h.KODE_AHSP)] && norm(h.JENIS) !== 'HARGA PASAR') out.tanpaKomponen.push(h);
  });
  hsdAll().forEach(r => {
    if (!(num(r.HARGA) > 0)) out.hsdKosong.push(r);
    const d = r.TANGGAL_UPDATE ? new Date(r.TANGGAL_UPDATE) : null;
    if (d && (year - d) / 864e5 > 365) out.hsdLama.push(r);
  });
  return out;
}
function penawaranNumber(p) {
  const year = new Date().getFullYear();
  const count = rows('PENAWARAN', true).filter(x => String(x.TANGGAL || '').startsWith(String(year)) && num(x.REV) === 0).length + 1;
  return String(setting('FORMAT_NOMOR_PENAWARAN') || DEFAULTS.FORMAT_NOMOR_PENAWARAN)
    .replace('{NO}', String(count).padStart(3, '0')).replace('{BLN}', ROMAWI[new Date().getMonth()]).replace('{THN}', year)
    .replace('{KODE}', p ? p.KODE : '');
}
function syaratText(jenis) { return rows('TEMPLATE_SYARAT').filter(t => norm(t.JENIS) === jenis).sort((a, b) => num(a.URUTAN) - num(b.URUTAN)).map(t => t.ISI).join('\n'); }
/* ============================================================
 * Router & halaman utama
 * ============================================================ */
const ROUTES = [];
function route(pattern, fn) { ROUTES.push({ re: new RegExp('^' + pattern.replace(/:\w+/g, '([^/]+)') + '$'), fn }); }
let CUR = { hash: '' };
document.addEventListener('input', e => { if (e.target.closest && e.target.closest('#app')) CUR.dirty = true; });
function render() {
  CUR.dirty = false; CUR.pendingRefresh = false;
  const hash = location.hash.replace(/^#/, '') || '/dashboard';
  CUR.hash = hash;
  const nav = hash.split('/')[1];
  $$('#nav a').forEach(a => a.classList.toggle('active', a.dataset.nav === nav));
  for (const r of ROUTES) {
    const m = hash.match(r.re);
    if (m) { try { r.fn(...m.slice(1).map(decodeURIComponent)); } catch (e) { console.error(e); $('#app').innerHTML = `<div class="notice danger">Terjadi kesalahan: ${esc(e.message)}</div>`; } window.scrollTo(0, CUR.keepScroll || 0); CUR.keepScroll = 0; return; }
  }
  location.hash = '#/dashboard';
}
function rerenderSoft() { CUR.keepScroll = window.scrollY; render(); }
function setTitle(t, sub) { $('#pageTitle').textContent = t; $('#pageSub').textContent = sub || 'by DG'; document.title = t + ' · RAB Rumah & Gedung'; }
const view = html => { $('#app').innerHTML = html + `<div class="footer-credit">RAB Rumah & Gedung v${APP_VERSION} · by DG</div>`; };

/* ---------- Dashboard ---------- */
route('/dashboard', () => {
  setTitle('Dashboard');
  const list = projects();
  if (!list.length) {
    view(`<div class="page-head"><h2>Selamat datang</h2></div>
      ${!S.meta.apiUrl ? `<div class="notice warning" style="margin-bottom:12px">${ICON.warn}<div>Aplikasi belum terhubung ke Google Sheet. Buka <a href="#/master/pengaturan">Master → Pengaturan</a> untuk mengisi URL Web App & token. Data tetap tersimpan di perangkat sampai terhubung.</div></div>` : ''}
      <div class="card empty"><p>Belum ada proyek. Mulai dengan membuat proyek pertama.</p><button class="btn primary" id="newP">${ICON.plus} Proyek Baru</button></div>`);
    $('#newP').onclick = () => projectForm();
    return;
  }
  let pid = S.meta.activeProject;
  if (!getRow('PROYEK', pid)) pid = list[0].ID;
  const p = getRow('PROYEK', pid);
  const rk = rekapOf(p);
  const top = Object.entries(rk.computed).map(([id, c]) => ({ it: getRow('RAB_ITEM', id), c })).filter(x => x.it).sort((a, b) => b.c.jumlah - a.c.jumlah).slice(0, 10);
  const warns = dashboardWarnings(p, rk);
  view(`<div class="page-head"><h2>Dashboard</h2>
      <div class="actions"><select class="input" id="pSel" style="min-width:220px">${list.map(x => `<option value="${x.ID}" ${x.ID === pid ? 'selected' : ''}>${esc(x.KODE)} · ${esc(x.NAMA)}</option>`).join('')}</select></div></div>
    <div class="grid k4">
      <div class="card kpi primary"><div class="label">Total RAB</div><div class="value">${rp(rk.total)}</div><div class="hint">${rk.ppnPct ? 'termasuk PPN ' + pct(rk.ppnPct, 0) : 'tanpa PPN'}</div></div>
      <div class="card kpi"><div class="label">Harga per m²</div><div class="value">${rk.perM2 ? rp(rk.perM2) : '-'}</div><div class="hint">${rk.luas ? N(rk.luas, 2) + ' m² bangunan' : 'isi luas bangunan'}</div></div>
      ${(() => { const t = tenagaOf(p, rk); return `<a class="card kpi clickable" href="#/rab/${pid}/tenaga" style="text-decoration:none;color:inherit"><div class="label">Tenaga kerja</div><div class="value">${t.totalOH ? N(t.totalOH, 0) + ' OH' : '-'}</div><div class="hint">${t.totalDurasi ? N(t.totalDurasi) + ' hari kerja · puncak ' + N(Math.max(0, ...t.weeks.map(w => w.orang))) + ' org' : 'isi durasi di tab Tenaga'}</div></a>`; })()}
      ${(() => { const M = jadwalModel(p); if (!M.scheduled) return `<a class="card kpi clickable" href="#/jadwal/${pid}/gantt" style="text-decoration:none;color:inherit"><div class="label">Progres</div><div class="value muted">-</div><div class="hint">susun jadwal dulu</div></a>`; const w = Math.max(1, Math.min(M.kurva.length, M.lastData || M.weekNow)); const a = M.actualCum(w), pl = M.planAt(w), st = devStatus(a - pl); return `<a class="card kpi clickable" href="#/jadwal/${pid}/kurva" style="text-decoration:none;color:inherit"><div class="label">Progres minggu ${w}</div><div class="value">${M.lastData ? pct(a, 1) : '-'}</div><div class="hint">rencana ${pct(pl, 1)} ${M.lastData ? badge(st[0], st[1]) : ''}</div></a>`; })()}
    </div>
    <div class="grid md2" style="margin-top:12px">
      <div class="card"><h3>Peringatan</h3>${warns.length ? `<div class="list warn-list">${warns.slice(0, 3).map(w => `<div class="li">${ICON.warn}<div class="main"><div class="t1">${esc(w.t)}</div><div class="t2">${esc(w.d)}</div></div>${w.href ? `<a class="btn sm" href="${w.href}">Buka</a>` : ''}</div>`).join('')}</div>${warns.length > 3 ? `<div class="small muted" style="margin-top:6px">+${warns.length - 3} peringatan lain</div>` : ''}` : '<div class="muted">Tidak ada peringatan. 👍</div>'}</div>
      <div class="card"><h3>Bobot per divisi</h3>${rk.groups.filter(g => g.subtotal).length ? rk.groups.filter(g => g.subtotal).map(g => `<div style="margin-bottom:10px"><div class="row between small"><span>${esc(g.row.KODE || '')} ${esc(g.row.URAIAN)}</span><span class="num">${pct(g.bobot)}</span></div><div class="bar ${g.smkk ? 'accent' : ''}"><span style="width:${Math.min(100, g.bobot)}%"></span></div></div>`).join('') : '<div class="muted">Belum ada nilai pekerjaan.</div>'}</div>
    </div>
    ${minggiIni(p)}
    <div class="card" style="margin-top:12px"><div class="row between"><h3>10 item termahal</h3><a class="btn sm" href="#/rab/${pid}">Buka RAB</a></div>
      ${top.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Uraian</th><th class="num">Volume</th><th class="num">Jumlah</th><th class="num">Bobot</th></tr></thead><tbody>${top.map(x => `<tr><td>${esc(x.it.URAIAN)}</td><td class="num">${Nmax(x.it.VOLUME, 2)} ${esc(x.it.SATUAN)}</td><td class="num">${rpCell(x.c.jumlah)}</td><td class="num">${pct(x.c.bobot)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="muted">Belum ada item.</div>'}</div>`);
  $('#pSel').onchange = e => { S.meta.activeProject = e.target.value; saveMeta(); render(); };
});
function dashboardWarnings(p, rk) {
  const w = [];
  const counts = {};
  rk.warnings.filter(x => num(x.item.VOLUME) > 0).forEach(x => counts[x.code] = (counts[x.code] || 0) + 1);
  const zero = Object.values(rk.computed).length - Object.keys(rk.computed).filter(id => num((getRow('RAB_ITEM', id) || {}).VOLUME) > 0).length;
  if (counts.HARGA_KOSONG) w.push({ t: `${counts.HARGA_KOSONG} item belum ada harga`, d: 'Hubungkan AHSP atau isi harga pasar.', href: `#/rab/${p.ID}` });
  if (counts.TANPA_AHSP) w.push({ t: `${counts.TANPA_AHSP} item belum terhubung AHSP`, d: 'Gunakan "Cocokkan AHSP otomatis" di tab Item.', href: `#/rab/${p.ID}` });
  if (counts.HSD_KOSONG) w.push({ t: `${counts.HSD_KOSONG} item memakai harga dasar kosong`, d: 'Lengkapi harga di Master → HSD.', href: '#/master/hsd' });
  if (counts.HSD_FALLBACK) w.push({ t: `${counts.HSD_FALLBACK} item memakai harga wilayah/tahun lain`, d: `Belum ada harga ${p.WILAYAH} ${p.TAHUN_ANGGARAN} untuk sebagian komponen.`, href: '#/master/hsd' });
  if (counts.PERLU_CEK) w.push({ t: `${counts.PERLU_CEK} item memakai AHSP berstatus PERLU CEK`, d: 'Verifikasi koefisien terhadap lampiran resmi.', href: '#/master/ahsp' });
  if (zero) w.push({ t: `${zero} item volumenya masih 0`, d: 'Isi volume, atau hapus item template yang tidak dipakai.', href: `#/rab/${p.ID}` });
  if (!num(p.LUAS_BANGUNAN)) w.push({ t: 'Luas bangunan belum diisi', d: 'Diperlukan untuk harga per m².', href: `#/rab/${p.ID}/info` });
  if (!S.meta.apiUrl) w.push({ t: 'Belum terhubung ke Google Sheet', d: 'Data baru tersimpan di perangkat ini.', href: '#/master/pengaturan' });
  if (pendingCount()) w.push({ t: `${pendingCount()} perubahan belum tersinkron`, d: S.lastError || 'Akan dikirim saat online.', href: '' });
  if (rows('AHSP_HEADER').some(h => norm(h.VERSI) === 'CONTOH')) w.push({ t: 'Library masih memakai data CONTOH', d: 'Import AHSP & HSD resmi lalu hapus data contoh.', href: '#/master/import' });
  return w;
}

/* ---------- Daftar proyek ---------- */
route('/rab', () => {
  setTitle('RAB', 'Daftar proyek');
  const list = projects();
  view(`<div class="page-head"><h2>Proyek</h2><div class="actions"><button class="btn primary" id="newP">${ICON.plus} Proyek Baru</button></div></div>
    ${list.length ? `<div class="search" style="margin-bottom:12px">${ICON.search}<input class="input" id="q" placeholder="Cari proyek, pemilik, lokasi…"></div>
    <div class="grid md2" id="plist">${list.map(projectCard).join('')}</div>` : `<div class="card empty"><p>Belum ada proyek.</p><button class="btn primary" id="newP2">${ICON.plus} Proyek Baru</button></div>`}`);
  $('#newP').onclick = () => projectForm();
  $('#newP2') && ($('#newP2').onclick = () => projectForm());
  $('#q') && ($('#q').oninput = e => {
    const q = e.target.value.toLowerCase();
    $('#plist').innerHTML = list.filter(p => [p.KODE, p.NAMA, p.PEMILIK, p.LOKASI].join(' ').toLowerCase().includes(q)).map(projectCard).join('');
  });
});
function projectCard(p) {
  return `<a class="card clickable" href="#/rab/${p.ID}" style="text-decoration:none;color:inherit;display:block">
    <div class="row between"><span class="small muted">${esc(p.KODE)} · ${esc(p.JENIS_BANGUNAN || '')}</span>${badge(p.STATUS || 'DRAFT')}</div>
    <div style="font-weight:700;font-size:16px;margin:4px 0">${esc(p.NAMA)}</div>
    <div class="small muted">${esc(p.PEMILIK || '-')} · ${esc(p.LOKASI || '-')}</div>
    <div class="row between" style="margin-top:10px"><span class="small muted">${N(p.LUAS_BANGUNAN, 2)} m² · ${esc(p.MODE || 'RUMAH')}</span><strong class="num">${rp(p.TOTAL_RAB)}</strong></div></a>`;
}
function projectForm(p) {
  const isNew = !p;
  p = p || { KODE: nextProjectCode(), TAHUN_ANGGARAN: new Date().getFullYear(), WILAYAH: setting('WILAYAH_DEFAULT'), MODE: setting('MODE_DEFAULT'), STATUS: 'DRAFT', JUMLAH_LANTAI: 1, VERSI_AHSP: setting('VERSI_AHSP_AKTIF') };
  const wil = rows('WILAYAH').map(w => w.NAMA);
  const profs = rows('PROFIL_USAHA');
  const body = `<form class="form-grid c2" id="pf">
    ${inp('KODE', 'Kode proyek', p.KODE, { req: true })}
    ${inp('NAMA', 'Nama proyek', p.NAMA, { req: true, ph: 'mis. Rumah Tinggal Bp. Andi' })}
    ${inp('PEMILIK', 'Pemilik / klien', p.PEMILIK)}
    ${inp('LOKASI', 'Lokasi', p.LOKASI)}
    ${sel('JENIS_BANGUNAN', 'Jenis bangunan', p.JENIS_BANGUNAN || 'Rumah 1 lantai', ['Rumah 1 lantai', 'Rumah 2 lantai', 'Rumah 3 lantai', 'Ruko', 'Gedung kantor', 'Gedung sekolah', 'Gudang', 'Renovasi', 'Lainnya'])}
    ${sel('MODE', 'Mode', p.MODE || 'RUMAH', [['RUMAH', 'Mode Rumah'], ['GEDUNG', 'Mode Gedung']], { hint: 'Mode Rumah menyembunyikan divisi khusus gedung' })}
    ${inp('JUMLAH_LANTAI', 'Jumlah lantai', p.JUMLAH_LANTAI, { numType: 'num' })}
    ${inp('LUAS_TANAH', 'Luas tanah (m²)', p.LUAS_TANAH, { numType: 'num' })}
    ${inp('LUAS_BANGUNAN', 'Luas bangunan total (m²)', p.LUAS_BANGUNAN, { numType: 'num', hint: 'Dasar harga per m²' })}
    ${inp('LUAS_PER_LANTAI', 'Luas per lantai', p.LUAS_PER_LANTAI, { ph: 'mis. 60; 45' })}
    ${inp('TAHUN_ANGGARAN', 'Tahun harga', p.TAHUN_ANGGARAN, { numType: 'num' })}
    <div class="field"><label>Wilayah harga</label><input class="input" name="WILAYAH" list="wilList" value="${esc(p.WILAYAH || '')}"><datalist id="wilList">${wil.map(w => `<option value="${esc(w)}">`).join('')}</datalist></div>
    ${inp('OVERHEAD_PCT', 'Biaya umum & keuntungan (%)', p.OVERHEAD_PCT, { numType: 'num', ph: String(setting('OVERHEAD_PCT')), hint: 'Kosong = pakai default ' + setting('OVERHEAD_PCT') + '%' })}
    ${inp('KONTINJENSI_PCT', 'Biaya tak terduga (%)', p.KONTINJENSI_PCT, { numType: 'num', ph: String(setting('KONTINJENSI_PCT')), hint: 'Isi 0 untuk mematikan' })}
    ${inp('PPN_PCT', 'PPN (%)', p.PPN_PCT, { numType: 'num', ph: String(setting('PPN_PCT')), hint: 'Isi 0 bila tanpa PPN' })}
    ${sel('PEMBULATAN', 'Pembulatan total (ke atas)', p.PEMBULATAN || setting('PEMBULATAN'), [['1', 'Tanpa pembulatan'], ['1000', 'Ribuan'], ['10000', 'Puluh ribuan'], ['100000', 'Ratus ribuan']])}
    ${sel('PROFIL_ID', 'Profil perusahaan (kop)', p.PROFIL_ID || (profilOf() || {}).ID || '', [['', '— pilih —']].concat(profs.map(x => [x.ID, x.NAMA])))}
    ${sel('STATUS', 'Status', p.STATUS || 'DRAFT', ['DRAFT', 'PENAWARAN', 'BERJALAN', 'SELESAI', 'BATAL'])}
    ${isNew ? sel('_TPL', 'Isi dari template', 'RUMAH1', [['', 'Kosong (tanpa template)']].concat(Object.entries(TEMPLATES).map(([k, t]) => [k, t.nama]))) : ''}
    ${txt('CATATAN', 'Catatan', p.CATATAN, { cls: 'span2', rows: 2 })}
  </form>`;
  openModal({
    title: isNew ? 'Proyek baru' : 'Ubah data proyek', body, wide: true,
    foot: `${!isNew ? `<button class="btn danger" id="delP">Hapus proyek</button><span class="spacer"></span>` : ''}<button class="btn" data-close>Batal</button><button class="btn primary" id="saveP">Simpan</button>`,
    onOpen: (el, m) => {
      bindMoneyInputs(el);
      $('#saveP', el).onclick = () => {
        const d = formData($('#pf', el));
        if (!d.NAMA) return toast('Nama proyek wajib diisi', 'err');
        const tpl = d._TPL; delete d._TPL;
        const row = save('PROYEK', Object.assign({ ID: p.ID || uid('P') }, d));
        S.meta.activeProject = row.ID; saveMeta();
        if (isNew && tpl) applyTemplate(row.ID, tpl);
        recompute(row.ID);
        m.close();
        toast(isNew ? 'Proyek dibuat' : 'Proyek disimpan', 'ok');
        if (isNew) location.hash = `#/rab/${row.ID}`; else rerenderSoft();
      };
      $('#delP', el) && ($('#delP', el).onclick = async () => {
        if (!await confirmBox(`Hapus proyek "${p.NAMA}" beserta semua itemnya?\nData ditandai terhapus di Sheet (tidak hilang permanen).`, { danger: true, ok: 'Hapus' })) return;
        projItems(p.ID).forEach(r => save('RAB_ITEM', { ID: r.ID, DELETED: 'YA' }, { silent: true }));
        remove('PROYEK', p.ID); m.close(); location.hash = '#/rab';
      });
    }
  });
}

/* ---------- Placeholder tahap berikutnya ---------- */
route('/lapangan', () => { setTitle('Lapangan'); view(`<div class="page-head"><h2>Lapangan</h2></div><div class="card empty"><p>Realisasi biaya, opname/BAP, dan pekerjaan tambah-kurang dibangun di <strong>Tahap 4</strong>.</p><p class="small">Input progres mingguan & foto dokumentasi sudah tersedia di menu <a href="#/jadwal">Jadwal → Input Progres</a>.</p><p class="small">Tenaga kerja & jumlah tukang sudah tersedia di tab <strong>Tenaga</strong> pada tiap proyek.</p></div>`); });

function minggiIni(p) {
  const M = jadwalModel(p); if (!M.scheduled || M.weekNow < 1) return '';
  const R = weekReport(M, M.weekNow);
  if (!R.mulai.length && !R.selesai.length && !R.late.length) return '';
  return `<div class="card" style="margin-top:12px"><div class="row between"><h3>Minggu ini (ke-${M.weekNow})</h3><a class="btn sm" href="#/jadwal/${p.ID}/laporan">Laporan</a></div>
    <div class="small"><strong>Mulai:</strong> ${R.mulai.map(g => esc(g.row.URAIAN)).join(', ') || '-'}<br><strong>Selesai:</strong> ${R.selesai.map(g => esc(g.row.URAIAN)).join(', ') || '-'}</div>
    ${R.late.length ? `<div class="notice danger" style="margin-top:8px">${ICON.warn}<div>Terlambat: ${R.late.map(x => `${esc(x.g.row.URAIAN)} (${N(x.ap, 0)}% dari rencana ${N(x.pp, 0)}%)`).join(', ')}</div></div>` : ''}</div>`;
}
/* ============================================================
 * Halaman proyek: Item | Rekap | Material | Penawaran | Revisi | Info
 * ============================================================ */
const PTABS = [['item', 'Item RAB'], ['rekap', 'Rekap'], ['material', 'Material'], ['tenaga', 'Tenaga'], ['jadwal', 'Jadwal'], ['penawaran', 'Penawaran'], ['revisi', 'Revisi'], ['info', 'Info & File']];
route('/rab/:id', id => projectView(id, 'item'));
route('/rab/:id/:tab', (id, tab) => projectView(id, tab));
function projectView(id, tab) {
  if (tab === 'jadwal') { location.replace(`#/jadwal/${id}/gantt`); return; }
  const p = getRow('PROYEK', id);
  if (!p) { view('<div class="card empty"><p>Proyek tidak ditemukan.</p><a class="btn" href="#/rab">Kembali</a></div>'); return; }
  if (S.meta.activeProject !== id) { S.meta.activeProject = id; saveMeta(); }
  setTitle(p.NAMA, `${p.KODE} · ${p.WILAYAH || '-'} ${p.TAHUN_ANGGARAN || ''}`);
  const rk = recompute(id);
  const head = `<div class="page-head"><div><a href="#/rab" class="small muted" style="text-decoration:none">${ICON.back.replace('class="ico"', 'class="ico" style="width:14px;height:14px;vertical-align:-2px"')} Semua proyek</a><h2 style="margin-top:4px">${esc(p.NAMA)}</h2></div>
    <div class="actions"><div class="card tight kpi" style="min-width:200px"><div class="label">Total RAB</div><div class="value" style="font-size:18px">${rp(rk.total)}</div></div></div></div>
    <div class="tabs">${PTABS.map(([k, t]) => `<button data-tab="${k}" class="${k === tab ? 'active' : ''}">${t}</button>`).join('')}</div><div id="tabBody"></div>`;
  view(head);
  $$('.tabs button').forEach(b => b.onclick = () => { location.hash = `#/rab/${id}/${b.dataset.tab}`; });
  const fn = { item: tabItems, rekap: tabRekap, material: tabMaterial, tenaga: tabTenaga, penawaran: tabPenawaran, revisi: tabRevisi, info: tabInfo }[tab] || tabItems;
  fn(p, rk, $('#tabBody'));
}

/* ---------- Tab Item ---------- */
function tabItems(p, rk, el) {
  const divs = projDivs(p.ID);
  const sugg = autoMatchSuggestions(p.ID);
  el.innerHTML = `<div class="row" style="margin-bottom:12px">
      <button class="btn primary" id="addItem">${ICON.plus} Item</button>
      <button class="btn" id="addDiv">${ICON.plus} Divisi</button>
      ${sugg.length ? `<button class="btn accent" id="autoMatch">Cocokkan AHSP otomatis (${sugg.length})</button>` : ''}
      <span class="spacer"></span><span class="small muted">${rk.itemCount} item · ${divs.length} divisi</span></div>
    ${divs.length ? `<div class="rab-list">${rk.groups.map(g => divBlock(p, g, rk)).join('')}</div>` :
      `<div class="card empty"><p>Belum ada divisi pekerjaan.</p><div class="row" style="justify-content:center">${Object.entries(TEMPLATES).map(([k, t]) => `<button class="btn sm" data-tpl="${k}">Template ${esc(t.nama)}</button>`).join('')}<button class="btn primary sm" id="addDiv2">${ICON.plus} Divisi manual</button></div></div>`}`;
  $('#addItem', el).onclick = () => { if (!divs.length) return toast('Tambahkan divisi dulu', 'err'); addItemFlow(p, divs[divs.length - 1].ID); };
  $('#addDiv', el).onclick = () => addDivisi(p);
  $('#addDiv2', el) && ($('#addDiv2', el).onclick = () => addDivisi(p));
  $$('[data-tpl]', el).forEach(b => b.onclick = () => { applyTemplate(p.ID, b.dataset.tpl); toast('Template diterapkan. Isi volume tiap item.', 'ok'); rerenderSoft(); });
  $('#autoMatch', el) && ($('#autoMatch', el).onclick = () => autoMatchModal(p));
  el.onclick = e => {
    const t = e.target.closest('[data-act]'); if (!t) return;
    const { act, id } = t.dataset;
    if (act === 'item') itemEditor(p, getRow('RAB_ITEM', id));
    if (act === 'additem') addItemFlow(p, id);
    if (act === 'divup' || act === 'divdown') moveDiv(p, id, act === 'divup' ? -1 : 1);
    if (act === 'divedit') divEditor(p, getRow('RAB_ITEM', id));
  };
}
function divBlock(p, g, rk) {
  const d = g.row;
  const items = g.items.map(it => {
    const c = rk.computed[it.ID];
    const h = it.KODE_AHSP ? ahspHead(it.KODE_AHSP) : null;
    const tags = [];
    if (norm(it.JENIS_HARGA) === 'PASAR') tags.push(badge('HARGA PASAR', 'accent'));
    else if (h) tags.push(`<span class="badge ${STATUS_CLASS[norm(h.JENIS)] || ''}">${esc(h.KODE_AHSP)}</span>`);
    const v0 = !(num(it.VOLUME) > 0);
    c.warn.filter(w => !(v0 && w === 'HARGA_KOSONG')).forEach(w => tags.push(badge(WARN_TEXT[w][0], WARN_TEXT[w][1])));
    if (v0) tags.push(badge('volume 0'));
    return `<div class="item-row" data-act="item" data-id="${it.ID}">
      <div class="main"><div class="ur"><span class="muted small">${esc(it.KODE || '')}</span> ${esc(it.URAIAN)}</div><div class="meta">${tags.join('')}</div></div>
      <div class="vol">${Nmax(it.VOLUME, 3)} ${esc(it.SATUAN)}</div>
      <div class="calc"><span class="vx">${Nmax(it.VOLUME, 3)} ${esc(it.SATUAN)} × </span>${rp(c.hs)}</div>
      <div class="amt">${rp(c.jumlah)}</div></div>`;
  }).join('');
  if (d.ID === '_') return `<div class="div-block"><div class="div-head"><span class="code">–</span><span class="name">Item tanpa divisi</span></div>${items}</div>`;
  return `<div class="div-block">
    <div class="div-head"><span class="code">${esc(d.KODE || '')}</span><span class="name">${esc(d.URAIAN)}${g.smkk ? ' ' + badge('terpisah', 'accent') : ''}</span>
      <span class="tot">${rp(g.subtotal)} <span class="small muted">· ${pct(g.bobot, 1)}</span></span>
      <span class="btns"><button class="btn ghost icon" data-act="divup" data-id="${d.ID}" title="Naik">${ICON.up}</button><button class="btn ghost icon" data-act="divdown" data-id="${d.ID}" title="Turun">${ICON.down}</button>
      <button class="btn ghost sm" data-act="divedit" data-id="${d.ID}">Ubah</button></span></div>
    ${items || '<div class="item-row muted small" style="cursor:default">Belum ada item.</div>'}
    <div class="div-foot"><button class="btn sm ghost" data-act="additem" data-id="${d.ID}">${ICON.plus} Tambah item</button></div></div>`;
}
function moveDiv(p, id, dir) {
  const divs = projDivs(p.ID), i = divs.findIndex(d => d.ID === id), j = i + dir;
  if (j < 0 || j >= divs.length) return;
  const a = divs[i], b = divs[j];
  save('RAB_ITEM', { ID: a.ID, URUTAN: b.URUTAN }, { silent: true }); save('RAB_ITEM', { ID: b.ID, URUTAN: a.URUTAN });
  rerenderSoft();
}
function addDivisi(p) {
  const have = projDivs(p.ID).map(d => d.KODE);
  const master = divisiMaster(p.MODE).filter(d => !have.includes(d.KODE));
  openModal({
    title: 'Tambah divisi', body: `${master.length ? `<div class="list">${master.map(d => `<label class="check li"><input type="checkbox" value="${d.ID}"><div class="main"><div class="t1">${esc(d.KODE)}. ${esc(d.NAMA)}</div></div></label>`).join('')}</div>` : '<div class="muted">Semua divisi standar sudah ada.</div>'}
      <div class="field" style="margin-top:14px"><label>Atau divisi custom</label><input class="input" id="dvCustom" placeholder="mis. PEKERJAAN KOLAM RENANG"></div>`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="dvOk">Tambahkan</button>`,
    onOpen: (el, m) => {
      $('#dvOk', el).onclick = () => {
        const all = rows('DIVISI');
        let urut = (projItems(p.ID).length ? Math.max(...projItems(p.ID).map(r => num(r.URUTAN))) : 0) + 1000;
        const picks = $$('input[type=checkbox]:checked', el).map(c => all.find(d => d.ID === c.value)).sort((a, b) => num(a.URUTAN) - num(b.URUTAN));
        picks.forEach(d => { save('RAB_ITEM', { ID: uid('D'), PROYEK_ID: p.ID, TIPE_BARIS: 'DIVISI', KODE: d.KODE, URAIAN: d.NAMA, URUTAN: urut }, { silent: true }); urut += 1000; });
        const c = $('#dvCustom', el).value.trim();
        if (c) save('RAB_ITEM', { ID: uid('D'), PROYEK_ID: p.ID, TIPE_BARIS: 'DIVISI', KODE: '', URAIAN: c.toUpperCase(), URUTAN: urut }, { silent: true });
        scheduleSync(); m.close(); rerenderSoft();
      };
    }
  });
}
function divEditor(p, d) {
  openModal({
    title: 'Ubah divisi', body: `<form class="form-grid" id="df">${inp('KODE', 'Kode', d.KODE)}${inp('URAIAN', 'Nama divisi', d.URAIAN)}</form>`,
    foot: `<button class="btn danger" id="dDel">Hapus divisi</button><span class="spacer"></span><button class="btn" data-close>Batal</button><button class="btn primary" id="dOk">Simpan</button>`,
    onOpen: (el, m) => {
      $('#dOk', el).onclick = () => { save('RAB_ITEM', Object.assign({ ID: d.ID }, formData($('#df', el)))); m.close(); rerenderSoft(); };
      $('#dDel', el).onclick = async () => {
        const n = divItems(p.ID, d.ID).length;
        if (!await confirmBox(`Hapus divisi "${d.URAIAN}"${n ? ` beserta ${n} item di dalamnya` : ''}?`, { danger: true, ok: 'Hapus' })) return;
        divItems(p.ID, d.ID).forEach(it => save('RAB_ITEM', { ID: it.ID, DELETED: 'YA' }, { silent: true }));
        remove('RAB_ITEM', d.ID); m.close(); rerenderSoft();
      };
    }
  });
}

/* ---------- Tambah item + referensi AHSP (section 10) ---------- */
function addItemFlow(p, divId, presetText = '') {
  const divs = projDivs(p.ID);
  const ctx = ctxOf(p);
  const body = `<div class="form-grid c3">
      ${inp('_q', 'Uraian pekerjaan', presetText, { cls: 'span2', ph: 'mis. pasang roster beton 20x20' })}
      ${inp('_s', 'Satuan (opsional)', '', { ph: 'm2 / m3 / bh …' })}
      ${sel('_d', 'Masuk ke divisi', divId, divs.map(d => [d.ID, (d.KODE ? d.KODE + '. ' : '') + d.URAIAN]), { cls: 'span3' })}
    </div>
    <div id="refRes" style="margin-top:14px"></div>
    <div class="card tight" style="margin-top:14px;background:var(--zebra)"><div class="small muted" style="margin-bottom:8px">Tidak ada yang cocok?</div>
      <div class="row"><button class="btn sm" id="optC">C. Buat analisa baru</button><button class="btn sm" id="optD">D. Harga pasar / LS</button></div></div>`;
  openModal({
    title: 'Tambah item', body, wide: true,
    onOpen: (el, m) => {
      const q = $('[name=_q]', el), s = $('[name=_s]', el), dv = $('[name=_d]', el);
      const run = () => {
        const div = getRow('RAB_ITEM', dv.value);
        const res = q.value.trim().length >= 2 ? Calc.search(idx(), q.value, { satuan: s.value, divisi: div && div.URAIAN, limit: 8 }) : [];
        $('#refRes', el).innerHTML = !q.value.trim() ? '<div class="muted small">Ketik uraian untuk melihat AHSP terkait dari library.</div>' :
          res.length ? `<div class="small muted" style="margin-bottom:6px">${res.length} AHSP terkait, urut paling relevan:</div>` + res.map(r => {
            const a = Calc.analyze(idx(), r.kode, ctx);
            return `<div class="ref-card"><div class="row between"><div><strong>${esc(r.kode)}</strong> ${badge(r.header.JENIS || '-')} ${norm(r.header.STATUS) === 'PERLU CEK' ? badge('PERLU CEK') : ''}</div><strong class="num">${rp(a.HSP)} / ${esc(r.header.SATUAN)}</strong></div>
              <div style="margin:4px 0">${esc(r.header.URAIAN)}</div>
              <div class="small muted">${esc(r.header.SUMBER || 'tanpa sumber')} · cocok karena ${esc(r.reasons.join(' · '))}${a.missing.length ? ' · <span class="neg">ada harga dasar kosong</span>' : ''}</div>
              <div class="row" style="margin-top:8px"><button class="btn sm primary" data-use="${esc(r.kode)}">A. Pakai AHSP ini</button><button class="btn sm" data-dup="${esc(r.kode)}">B. Duplikat & modifikasi</button><button class="btn sm ghost" data-cmp="${esc(r.kode)}">Bandingkan sejenis</button></div>
              <div data-cmpbox="${esc(r.kode)}"></div></div>`;
          }).join('') : '<div class="notice warning">Tidak ada AHSP yang cocok di library. Gunakan opsi C atau D, atau tambah tag/sinonim di Master → AHSP.</div>';
      };
      q.oninput = debounce(run, 200); s.oninput = debounce(run, 200); dv.onchange = run; run();
      const mk = (extra) => {
        const urut = Math.max(0, ...divItems(p.ID, dv.value).map(r => num(r.URUTAN)), num((getRow('RAB_ITEM', dv.value) || {}).URUTAN)) + 1;
        return save('RAB_ITEM', Object.assign({ ID: uid('I'), PROYEK_ID: p.ID, PARENT_ID: dv.value, TIPE_BARIS: 'ITEM', URAIAN: q.value.trim(), SATUAN: s.value.trim(), VOLUME: 0, URUTAN: urut }, extra));
      };
      $('#refRes', el).onclick = e => {
        const b = e.target.closest('button'); if (!b) return;
        if (b.dataset.use) { const h = ahspHead(b.dataset.use); const it = mk({ JENIS_HARGA: 'AHSP', KODE_AHSP: h.KODE_AHSP, URAIAN: q.value.trim() || h.URAIAN, SATUAN: h.SATUAN }); m.close(); recompute(p.ID); rerenderSoft(); itemEditor(p, it); }
        if (b.dataset.dup) {
          const nk = duplicateAhsp(b.dataset.dup, {});
          m.close();
          ahspEditor(nk, { project: p, onSaved: kode => { const h = ahspHead(kode); const it = mk({ JENIS_HARGA: 'AHSP', KODE_AHSP: kode, URAIAN: q.value.trim() || h.URAIAN, SATUAN: h.SATUAN }); recompute(p.ID); rerenderSoft(); itemEditor(p, it); } });
        }
        if (b.dataset.cmp) {
          const box = $(`[data-cmpbox="${CSS.escape(b.dataset.cmp)}"]`, el);
          const sim = Calc.similar(idx(), b.dataset.cmp, ctx, 3);
          box.innerHTML = sim.list.length ? `<div class="small" style="margin-top:8px;border-top:1px solid var(--border);padding-top:8px">${sim.list.map(x => `<div class="row between"><span>${esc(x.header.KODE_AHSP)} · ${esc(x.header.URAIAN)}</span><span class="num">${rp(x.HSP)}</span></div>`).join('')}<div class="row between muted"><span>Rata-rata sejenis</span><span class="num">${rp(sim.avg)}</span></div></div>` : '<div class="small muted" style="margin-top:8px">Belum ada AHSP sejenis (divisi & satuan sama) untuk pembanding.</div>';
        }
      };
      $('#optC', el).onclick = () => {
        if (!q.value.trim()) return toast('Isi uraian pekerjaan dulu', 'err');
        const kode = nextCustomCode();
        const div = getRow('RAB_ITEM', dv.value);
        save('AHSP_HEADER', { ID: uid('H'), KODE_AHSP: kode, URAIAN: q.value.trim(), SATUAN: s.value.trim(), DIVISI: div ? div.URAIAN : '', JENIS: 'CUSTOM', STATUS: 'PERLU CEK', SCOPE: 'UMUM', DIBUAT_OLEH: 'PWA', TANGGAL: todayStr(), VERSI: 'CUSTOM' });
        m.close();
        ahspEditor(kode, { project: p, isNew: true, onSaved: k => { const h = ahspHead(k); const it = mk({ JENIS_HARGA: 'AHSP', KODE_AHSP: k, SATUAN: h.SATUAN }); recompute(p.ID); rerenderSoft(); itemEditor(p, it); }, onCancel: () => { const h = ahspHead(kode); if (h && !ahspKomp(kode).length) remove('AHSP_HEADER', h.ID); } });
      };
      $('#optD', el).onclick = () => {
        if (!q.value.trim()) return toast('Isi uraian pekerjaan dulu', 'err');
        const it = mk({ JENIS_HARGA: 'PASAR', KODE_AHSP: '', SATUAN: s.value.trim() || 'ls' });
        m.close(); recompute(p.ID); rerenderSoft(); itemEditor(p, it);
      };
    }
  });
}

/* ---------- Editor item ---------- */
function itemEditor(p, it) {
  if (!it) return;
  const ctx = ctxOf(p);
  const divs = projDivs(p.ID);
  let jenis = norm(it.JENIS_HARGA) === 'PASAR' ? 'PASAR' : 'AHSP';
  let kode = it.KODE_AHSP || '';
  const body = `<form id="itf"><div class="form-grid c3">
      ${inp('URAIAN', 'Uraian pekerjaan', it.URAIAN, { cls: 'span2', req: true })}
      ${inp('SATUAN', 'Satuan', it.SATUAN)}
      <div class="field"><label>Volume</label><div class="row" style="flex-wrap:nowrap"><input class="input num" name="VOLUME" inputmode="decimal" data-num="num" value="${esc(String(num(it.VOLUME)).replace('.', ','))}"><button type="button" class="btn icon" id="vCalc" title="Kalkulator volume">${ICON.calc}</button></div><span class="hint" id="vHint"></span></div>
      ${sel('PARENT_ID', 'Divisi', it.PARENT_ID, divs.map(d => [d.ID, (d.KODE ? d.KODE + '. ' : '') + d.URAIAN]), { cls: 'span2' })}
    </div>
    <div class="field" style="margin-top:14px"><label>Dasar harga</label><div class="seg" id="jSeg"><button type="button" data-j="AHSP">AHSP</button><button type="button" data-j="PASAR">Harga pasar / LS</button></div></div>
    <div id="jBody" style="margin-top:12px"></div>
    <div class="form-grid c2" style="margin-top:12px">
      ${txt('SPESIFIKASI', 'Spesifikasi (tampil di penawaran)', it.SPESIFIKASI, { rows: 2 })}
      ${txt('CATATAN', 'Catatan internal', it.CATATAN, { rows: 2 })}
    </div></form>`;
  openModal({
    title: 'Item pekerjaan', body, wide: true,
    foot: `<button class="btn danger" id="iDel">Hapus</button><button class="btn ghost" id="iDup">Duplikat</button><span class="spacer"></span><button class="btn" data-close>Batal</button><button class="btn primary" id="iOk">Simpan</button>`,
    onOpen: (el, m) => {
      const volIn = $('[name=VOLUME]', el);
      const dets = rows('VOLUME_DETAIL').filter(d => d.ITEM_ID === it.ID);
      $('#vHint', el).textContent = dets.length ? `dari ${dets.length} baris rincian` : '';
      const drawJ = () => {
        $$('#jSeg button', el).forEach(b => b.classList.toggle('active', b.dataset.j === jenis));
        const jb = $('#jBody', el);
        if (jenis === 'PASAR') {
          jb.innerHTML = `<div class="form-grid c2">${inp('HARGA_PASAR', 'Harga satuan', it.HARGA_PASAR, { numType: 'rp', req: true })}${inp('SUMBER_HARGA', 'Sumber harga (wajib)', it.SUMBER_HARGA, { req: true, ph: 'mis. Penawaran CV Maju Jaya, 02/11/2026' })}</div>
            <div class="notice" style="margin-top:8px">Item harga pasar ditandai <strong>HARGA PASAR</strong> di RAB internal dan masuk laporan AHSP non-resmi.</div>`;
          bindMoneyInputs(jb); return;
        }
        const h = kode ? ahspHead(kode) : null;
        if (!h) {
          jb.innerHTML = `<div class="notice warning">${ICON.warn}<div>Item belum terhubung AHSP.</div></div><div class="row" style="margin-top:8px"><button type="button" class="btn sm primary" id="pickA">Pilih AHSP dari library</button></div>`;
        } else {
          const dv = ahspDeviation(kode, p), a = dv.a;
          jb.innerHTML = `<div class="ref-card sel"><div class="row between"><div><strong>${esc(h.KODE_AHSP)}</strong> ${badge(h.JENIS || '-')} ${badge(h.STATUS || '-')}</div><div class="row"><button type="button" class="btn sm" id="pickA">Ganti</button>${isOfficial(h) ? `<button type="button" class="btn sm" id="dupA">Duplikat & modifikasi</button>` : `<button type="button" class="btn sm" id="editA">Ubah analisa</button>`}</div></div>
            <div style="margin:4px 0">${esc(h.URAIAN)} <span class="muted">/ ${esc(h.SATUAN)}</span></div>
            <div class="small muted">${esc(h.SUMBER || '')}${h.KODE_REFERENSI ? ' · referensi ' + esc(h.KODE_REFERENSI) : ''}${h.DASAR_PERHITUNGAN ? ' · dasar: ' + esc(h.DASAR_PERHITUNGAN) : ''}</div></div>
            ${analisaTable(a, ctx)}
            ${dv.over ? `<div class="notice warning" style="margin-top:8px">${ICON.warn}<div>HSP menyimpang ${pct(dv.dev, 1)} dari rata-rata AHSP sejenis (${rp(dv.sim.avg)}). Ambang ±${setting('AMBANG_DEVIASI_HSP')}%.</div></div>` : ''}
            ${h.SATUAN && $('[name=SATUAN]', el).value && norm(h.SATUAN) !== norm($('[name=SATUAN]', el).value) ? `<div class="notice warning" style="margin-top:8px">${ICON.warn}<div>Satuan item (${esc($('[name=SATUAN]', el).value)}) berbeda dengan satuan AHSP (${esc(h.SATUAN)}).</div></div>` : ''}
            <div class="form-grid c2" style="margin-top:10px">${inp('HARGA_PASAR', 'Harga pembanding pasar (opsional)', it.HARGA_PASAR, { numType: 'rp', hint: 'Untuk perbandingan Mode AHSP vs Pasar di tab Rekap' })}${inp('SUMBER_HARGA', 'Sumber harga pembanding', it.SUMBER_HARGA)}</div>`;
          bindMoneyInputs(jb);
        }
        $('#pickA', jb) && ($('#pickA', jb).onclick = () => pickAhsp(p, $('[name=URAIAN]', el).value, $('[name=SATUAN]', el).value, k => { kode = k; const hh = ahspHead(k); if (!$('[name=SATUAN]', el).value) $('[name=SATUAN]', el).value = hh.SATUAN; drawJ(); }));
        $('#dupA', jb) && ($('#dupA', jb).onclick = () => { const nk = duplicateAhsp(kode); ahspEditor(nk, { project: p, onSaved: k => { kode = k; drawJ(); } }); });
        $('#editA', jb) && ($('#editA', jb).onclick = () => ahspEditor(kode, { project: p, onSaved: () => drawJ() }));
      };
      $$('#jSeg button', el).forEach(b => b.onclick = () => { jenis = b.dataset.j; drawJ(); });
      drawJ();
      $('#vCalc', el).onclick = () => volumeCalc(p, it, v => { volIn.value = String(v).replace('.', ','); $('#vHint', el).textContent = 'dari rincian kalkulator'; });
      $('#iOk', el).onclick = () => {
        const d = formData($('#itf', el));
        if (!d.URAIAN) return toast('Uraian wajib diisi', 'err');
        if (jenis === 'PASAR' && !(num(d.HARGA_PASAR) > 0)) return toast('Isi harga satuan', 'err');
        if (jenis === 'PASAR' && !d.SUMBER_HARGA) return toast('Sumber harga wajib diisi untuk harga pasar', 'err');
        save('RAB_ITEM', Object.assign({ ID: it.ID }, d, { VOLUME: d.VOLUME === '' ? 0 : d.VOLUME, JENIS_HARGA: jenis, KODE_AHSP: jenis === 'AHSP' ? kode : '' }));
        recompute(p.ID); m.close(); rerenderSoft();
      };
      $('#iDel', el).onclick = async () => { if (!await confirmBox(`Hapus item "${it.URAIAN}"?`, { danger: true, ok: 'Hapus' })) return; remove('RAB_ITEM', it.ID); recompute(p.ID); m.close(); rerenderSoft(); };
      $('#iDup', el).onclick = () => { const c = Object.assign({}, it, { ID: uid('I'), URAIAN: it.URAIAN + ' (salinan)', URUTAN: num(it.URUTAN) + 0.5, UPDATED_AT: '' }); save('RAB_ITEM', c); recompute(p.ID); m.close(); rerenderSoft(); toast('Item diduplikat', 'ok'); };
    }
  });
}
function analisaTable(a, ctx) {
  if (!a || !a.ok) return '';
  const st = { OK: '', TAHUN_LAIN: badge('tahun lain', 'warning'), WILAYAH_LAIN: badge('wilayah lain', 'warning'), KOSONG: badge('harga kosong', 'danger') };
  const grp = j => a.rows.filter(r => r.jenis === j);
  const sect = (j, label, sum) => grp(j).length ? `<tr class="div"><td colspan="5">${label}</td><td class="num">${rpCell(sum)}</td></tr>` + grp(j).map(r => `<tr><td>${esc(r.kodeHsd)}</td><td>${esc(r.nama)} ${st[r.status] || ''}</td><td>${esc(r.satuan)}</td><td class="num">${Nmax(r.koef, 4)}</td><td class="num">${rpCell(r.hargaDasar)}</td><td class="num">${rpCell(r.jumlah)}</td></tr>`).join('') : '';
  return `<div class="table-wrap" style="margin-top:10px"><table class="t"><thead><tr><th>Kode</th><th>Komponen</th><th>Sat</th><th class="num">Koefisien</th><th class="num">Harga dasar</th><th class="num">Jumlah</th></tr></thead><tbody>
    ${sect('UPAH', 'A. Tenaga kerja', a.upah)}${sect('BAHAN', 'B. Bahan', a.bahan)}${sect('ALAT', 'C. Peralatan', a.alat)}
    <tr class="sub"><td colspan="5">D. Jumlah (A+B+C)</td><td class="num">${rpCell(a.E)}</td></tr>
    <tr><td colspan="5">E. Biaya umum & keuntungan ${pct(ctx.overhead, 1)} × D</td><td class="num">${rpCell(a.F)}</td></tr>
    <tr class="tot"><td colspan="5">F. Harga satuan pekerjaan (D+E)</td><td class="num">${rpCell(a.HSP)}</td></tr></tbody></table></div>
    <div class="small muted" style="margin-top:4px">Harga dasar: ${esc(ctx.wilayah)} ${ctx.tahun}</div>`;
}
function pickAhsp(p, text, satuan, onPick) {
  const ctx = ctxOf(p);
  openModal({
    title: 'Pilih AHSP', wide: true,
    body: `<div class="search">${ICON.search}<input class="input" id="pq" value="${esc(text)}" placeholder="Cari kode / uraian / tag"></div><div id="pres" style="margin-top:12px"></div>`,
    onOpen: (el, m) => {
      const run = () => {
        const qv = $('#pq', el).value.trim();
        let res = qv ? Calc.search(idx(), qv, { satuan, limit: 20 }) : rows('AHSP_HEADER').slice(0, 30).map(h => ({ kode: h.KODE_AHSP, header: h, reasons: [] }));
        const exact = rows('AHSP_HEADER').find(h => norm(h.KODE_AHSP) === norm(qv));
        if (exact && !res.some(r => r.kode === exact.KODE_AHSP)) res.unshift({ kode: exact.KODE_AHSP, header: exact, reasons: ['kode sama'] });
        $('#pres', el).innerHTML = res.length ? res.map(r => `<div class="ref-card" style="cursor:pointer" data-k="${esc(r.kode)}"><div class="row between"><strong>${esc(r.kode)}</strong><span class="num">${rp(Calc.analyze(idx(), r.kode, ctx).HSP)} / ${esc(r.header.SATUAN)}</span></div><div>${esc(r.header.URAIAN)}</div><div class="small muted">${badge(r.header.JENIS || '-')} ${esc(r.reasons.join(' · '))}</div></div>`).join('') : '<div class="muted">Tidak ada hasil.</div>';
      };
      $('#pq', el).oninput = debounce(run, 200); run();
      $('#pres', el).onclick = e => { const c = e.target.closest('[data-k]'); if (c) { m.close(); onPick(c.dataset.k); } };
    }
  });
}

/* ---------- Kalkulator volume ---------- */
function volumeCalc(p, it, onUse) {
  let det = rows('VOLUME_DETAIL').filter(d => d.ITEM_ID === it.ID).sort((a, b) => num(a.URUTAN) - num(b.URUTAN)).map(d => Object.assign({}, d));
  if (!det.length) det = [{ ID: uid('V'), KETERANGAN: '', PANJANG: '', LEBAR: '', TINGGI: '', JUMLAH: 1, PENGURANG: 'TIDAK' }];
  const removed = [];
  openModal({
    title: 'Kalkulator volume · ' + it.URAIAN, wide: true,
    body: `<div class="small muted" style="margin-bottom:8px">Hasil per baris = P × L × T × Jumlah (kolom kosong diabaikan). Centang "Kurangi" untuk bukaan pintu/jendela.</div><div id="vrows"></div>
      <button class="btn sm" id="vAdd" style="margin-top:8px">${ICON.plus} Baris</button>
      <div class="card tight" style="margin-top:12px"><div class="row between"><strong>Total volume</strong><strong id="vTot" class="num"></strong></div></div>`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="vUse">Pakai hasil</button>`,
    onOpen: (el, m) => {
      const draw = () => {
        $('#vrows', el).innerHTML = `<div class="table-wrap"><table class="t"><thead><tr><th>Keterangan</th><th class="num">P</th><th class="num">L</th><th class="num">T</th><th class="num">Jml</th><th>Kurangi</th><th class="num">Hasil</th><th></th></tr></thead><tbody>${det.map((d, i) => {
          const h = Calc.volumeFromDetails([Object.assign({}, d, { PENGURANG: 'TIDAK' })]);
          return `<tr data-i="${i}"><td><input class="input" data-f="KETERANGAN" value="${esc(d.KETERANGAN)}" placeholder="mis. Dinding as A" style="min-width:140px"></td>${['PANJANG', 'LEBAR', 'TINGGI', 'JUMLAH'].map(f => `<td><input class="input num" data-f="${f}" inputmode="decimal" value="${d[f] === '' || d[f] === undefined ? '' : String(d[f]).replace('.', ',')}" style="width:76px"></td>`).join('')}
            <td><input type="checkbox" data-f="PENGURANG" ${norm(d.PENGURANG) === 'YA' ? 'checked' : ''} style="width:20px;height:20px"></td><td class="num ${norm(d.PENGURANG) === 'YA' ? 'neg' : ''}">${norm(d.PENGURANG) === 'YA' ? '-' : ''}${Nmax(h, 3)}</td><td><button class="btn ghost icon" data-rm="${i}">${ICON.close}</button></td></tr>`;
        }).join('')}</tbody></table></div>`;
        $('#vTot', el).textContent = Nmax(Calc.volumeFromDetails(det), 3) + ' ' + (it.SATUAN || '');
      };
      draw();
      $('#vrows', el).addEventListener('change', e => {
        const tr = e.target.closest('tr[data-i]'); if (!tr) return;
        const d = det[+tr.dataset.i], f = e.target.dataset.f;
        if (f === 'PENGURANG') d[f] = e.target.checked ? 'YA' : 'TIDAK';
        else if (f === 'KETERANGAN') d[f] = e.target.value;
        else d[f] = parseNum(e.target.value);
        setTimeout(draw, 0);
      });
      $('#vrows', el).addEventListener('click', e => { const b = e.target.closest('[data-rm]'); if (b) { removed.push(det[+b.dataset.rm]); det.splice(+b.dataset.rm, 1); draw(); } });
      $('#vAdd', el).onclick = () => { det.push({ ID: uid('V'), KETERANGAN: '', PANJANG: '', LEBAR: '', TINGGI: '', JUMLAH: 1, PENGURANG: 'TIDAK' }); draw(); };
      $('#vUse', el).onclick = () => {
        const tot = Calc.volumeFromDetails(det);
        det.forEach((d, i) => save('VOLUME_DETAIL', Object.assign({}, d, { PROYEK_ID: p.ID, ITEM_ID: it.ID, URUTAN: i + 1, HASIL: Calc.volumeFromDetails([Object.assign({}, d, { PENGURANG: 'TIDAK' })]) * (norm(d.PENGURANG) === 'YA' ? -1 : 1) }), { silent: true }));
        removed.forEach(d => { if (getRow('VOLUME_DETAIL', d.ID)) save('VOLUME_DETAIL', { ID: d.ID, DELETED: 'YA' }, { silent: true }); });
        scheduleSync(); m.close(); onUse(tot);
      };
    }
  });
}

/* ---------- Cocokkan AHSP otomatis ---------- */
function autoMatchModal(p) {
  const list = autoMatchSuggestions(p.ID);
  const ctx = ctxOf(p);
  openModal({
    title: 'Cocokkan AHSP otomatis', wide: true,
    body: `<div class="notice" style="margin-bottom:12px">Saran berdasarkan kata kunci, divisi, dan satuan. Saran yang kuat sudah dicentang; saran lemah perlu kamu periksa dulu. Hanya yang dicentang yang dihubungkan.</div>
      <div class="list">${list.map((s, i) => `<label class="li check" style="align-items:flex-start"><input type="checkbox" data-i="${i}" ${s.strong ? 'checked' : ''} style="margin-top:3px"><div class="main"><div class="t1">${esc(s.item.URAIAN)} <span class="muted">(${esc(s.item.SATUAN)})</span></div>
        <div class="t2">→ <strong>${esc(s.best.kode)}</strong> ${esc(s.best.header.URAIAN)} · ${rp(Calc.analyze(idx(), s.best.kode, ctx).HSP)} ${norm(s.best.header.STATUS) === 'PERLU CEK' ? badge('PERLU CEK') : ''} ${s.strong ? '' : badge('saran lemah', 'warning')}</div></div></label>`).join('')}</div>`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="amOk">Terapkan</button>`,
    onOpen: (el, m) => {
      $('#amOk', el).onclick = () => {
        let n = 0;
        $$('input[data-i]:checked', el).forEach(c => { const s = list[+c.dataset.i]; save('RAB_ITEM', { ID: s.item.ID, KODE_AHSP: s.best.kode, JENIS_HARGA: 'AHSP' }, { silent: true }); n++; });
        scheduleSync(); recompute(p.ID); m.close(); rerenderSoft(); toast(`${n} item dihubungkan ke AHSP`, 'ok');
      };
    }
  });
}

/* ---------- Tab Rekap ---------- */
function tabRekap(p, rk, el) {
  const s = rk.split;
  // Mode AHSP vs Pasar
  let cmpA = 0, cmpP = 0, cmpN = 0;
  rk.groups.forEach(g => g.items.forEach(it => { if (it.KODE_AHSP && num(it.HARGA_PASAR) > 0) { cmpN++; cmpA += rk.computed[it.ID].jumlah; cmpP += num(it.HARGA_PASAR) * num(it.VOLUME); } }));
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><button class="btn primary" id="pdfRab">${ICON.pdf} PDF RAB</button><button class="btn" id="xlsRab">Excel</button><button class="btn ghost" id="waRab">${ICON.share} WhatsApp</button><span class="spacer"></span><button class="btn ghost" id="editP">Ubah parameter</button></div>
    <div class="grid md2">
    <div class="card"><h3>Rekapitulasi</h3><div class="table-wrap"><table class="t"><thead><tr><th>No</th><th>Uraian pekerjaan</th><th class="num">Jumlah</th><th class="num">Bobot</th></tr></thead><tbody>
      ${rk.groups.map(g => `<tr><td>${esc(g.row.KODE || '')}</td><td>${esc(g.row.URAIAN)}</td><td class="num">${rpCell(g.subtotal)}</td><td class="num">${pct(g.bobot)}</td></tr>`).join('')}
      <tr class="sub"><td colspan="2">A. Jumlah biaya pekerjaan</td><td class="num">${rpCell(rk.A)}</td><td></td></tr>
      <tr class="sub"><td colspan="2">B. Biaya penerapan SMKK</td><td class="num">${rpCell(rk.B)}</td><td></td></tr>
      <tr><td colspan="2">C. Total (A+B)</td><td class="num">${rpCell(rk.C)}</td><td></td></tr>
      ${rk.kontPct ? `<tr><td colspan="2">D. Biaya tak terduga ${pct(rk.kontPct, 1)}</td><td class="num">${rpCell(rk.D)}</td><td></td></tr>` : ''}
      <tr><td colspan="2">E. Total sebelum PPN</td><td class="num">${rpCell(rk.E)}</td><td></td></tr>
      ${rk.ppnPct ? `<tr><td colspan="2">PPN ${pct(rk.ppnPct, 1)}</td><td class="num">${rpCell(rk.ppn)}</td><td></td></tr>` : ''}
      ${rk.selisihBulat ? `<tr><td colspan="2">Pembulatan</td><td class="num">${rpCell(rk.selisihBulat)}</td><td></td></tr>` : ''}
      <tr class="tot"><td colspan="2">TOTAL RAB</td><td class="num">${rpCell(rk.total)}</td><td></td></tr></tbody></table></div>
      <div class="small" style="margin-top:8px"><em>Terbilang: ${esc(Calc.terbilang(rk.total))}</em></div>
      ${rk.perM2 ? `<div class="small muted" style="margin-top:4px">Harga per m²: <strong>${rp(rk.perM2)}</strong> (${N(rk.luas, 2)} m²)</div>` : ''}</div>
    <div class="stack">
      <div class="card"><h3>Komposisi biaya</h3><dl class="kv">
        <dt>Bahan</dt><dd>${rp(s.bahan)}</dd><dt>Upah</dt><dd>${rp(s.upah)}</dd><dt>Alat</dt><dd>${rp(s.alat)}</dd>
        <dt>Biaya umum & keuntungan</dt><dd>${rp(s.overhead)}</dd><dt>Harga pasar / LS</dt><dd>${rp(s.lain)}</dd></dl>
        ${rk.C ? `<div class="bar" style="margin-top:12px;display:flex;height:12px">${[['bahan', THEME.primary], ['upah', THEME.accent], ['alat', THEME.info], ['overhead', THEME.muted], ['lain', THEME.success]].map(([k, c]) => `<span style="width:${s[k] / rk.C * 100}%;background:${c};border-radius:0"></span>`).join('')}</div><div class="small muted" style="margin-top:6px">Bahan · Upah · Alat · Overhead · Pasar</div>` : ''}</div>
      <div class="card"><h3>Mode AHSP vs Harga Pasar</h3>${cmpN ? `<dl class="kv"><dt>${cmpN} item dibandingkan · AHSP</dt><dd>${rp(cmpA)}</dd><dt>Harga pasar</dt><dd>${rp(cmpP)}</dd><dt>Selisih</dt><dd class="${cmpP - cmpA < 0 ? 'neg' : ''}">${rp(cmpP - cmpA)} (${pct(cmpA ? (cmpP - cmpA) / cmpA * 100 : 0, 1)})</dd></dl>` : '<div class="small muted">Isi "Harga pembanding pasar" di item AHSP untuk membandingkan.</div>'}</div>
      <div class="card"><h3>Parameter</h3><dl class="kv"><dt>Wilayah / tahun harga</dt><dd>${esc(rk.ctx.wilayah)} ${rk.ctx.tahun}</dd><dt>Biaya umum & keuntungan</dt><dd>${pct(rk.ctx.overhead, 1)}</dd><dt>Biaya tak terduga</dt><dd>${pct(rk.kontPct, 1)}</dd><dt>PPN</dt><dd>${pct(rk.ppnPct, 1)}</dd><dt>Pembulatan</dt><dd>${rk.pembulatan > 1 ? 'ke ' + N(rk.pembulatan) : '-'}</dd></dl></div>
    </div></div>`;
  $('#editP', el).onclick = () => projectForm(p);
  $('#pdfRab', el).onclick = () => exportRabPdf(p);
  $('#xlsRab', el).onclick = () => exportRabXlsx(p);
  $('#waRab', el).onclick = () => shareWa(`*RAB ${p.NAMA}*\n${p.LOKASI || ''}\nTotal: ${rp(rk.total)}${rk.perM2 ? `\nHarga/m²: ${rp(rk.perM2)}` : ''}\n\n_dibuat dengan RAB Rumah & Gedung_`);
}
function shareWa(text) { window.open('https://wa.me/?text=' + encodeURIComponent(text), '_blank'); }

/* ---------- Tab Material ---------- */
function tabMaterial(p, rk, el) {
  const list = Calc.takeoff(idx(), rk);
  let f = 'BAHAN';
  const draw = () => {
    const l = list.filter(x => x.jenis === f), tot = l.reduce((a, b) => a + b.jumlah, 0);
    $('#mt', el).innerHTML = l.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Kode</th><th>${f === 'UPAH' ? 'Tenaga' : f === 'ALAT' ? 'Alat' : 'Bahan'}</th><th class="num">Kebutuhan</th><th>Sat</th><th class="num">Harga dasar</th><th class="num">Jumlah</th></tr></thead><tbody>${l.map(x => `<tr><td>${esc(x.kode)}</td><td>${esc(x.nama)}</td><td class="num">${Nmax(x.qty, 3)}</td><td>${esc(x.satuan)}</td><td class="num">${rpCell(x.harga)}</td><td class="num">${rpCell(x.jumlah)}</td></tr>`).join('')}<tr class="tot"><td colspan="5">Total ${f.toLowerCase()} (sebelum overhead)</td><td class="num">${rpCell(tot)}</td></tr></tbody></table></div>
      ${f === 'UPAH' ? `<div class="small muted" style="margin-top:6px">Kebutuhan tenaga dalam OH (orang-hari). Jumlah tukang & durasi ada di tab <a href="#/rab/${p.ID}/tenaga">Tenaga</a>.</div>` : ''}` : '<div class="card empty"><p>Belum ada kebutuhan dari item ber-AHSP.</p></div>';
  };
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><div class="seg" id="mSeg"><button data-f="BAHAN" class="active">Bahan</button><button data-f="UPAH">Tenaga (OH)</button><button data-f="ALAT">Alat</button></div><span class="spacer"></span><button class="btn" id="mX">Excel</button></div><div id="mt"></div>
    <div class="small muted" style="margin-top:8px">Item harga pasar/LS tidak masuk daftar kebutuhan karena tidak punya rincian komponen.</div>`;
  $$('#mSeg button', el).forEach(b => b.onclick = () => { f = b.dataset.f; $$('#mSeg button', el).forEach(x => x.classList.toggle('active', x === b)); draw(); });
  $('#mX', el).onclick = async () => {
    const X = await needXlsx();
    const data = [['Jenis', 'Kode', 'Nama', 'Kebutuhan', 'Satuan', 'Harga dasar', 'Jumlah']].concat(list.map(x => [x.jenis, x.kode, x.nama, Calc.round(x.qty, 4), x.satuan, x.harga, Math.round(x.jumlah)]));
    const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(data), 'Kebutuhan');
    X.writeFile(wb, `${p.KODE}_MATERIAL_${todayStr().replace(/-/g, '')}.xlsx`);
  };
  draw();
}

/* ---------- Tab Revisi ---------- */
function tabRevisi(p, rk, el) {
  const revs = rows('REVISI').filter(r => r.PROYEK_ID === p.ID).sort((a, b) => String(b.TANGGAL + b.UPDATED_AT).localeCompare(String(a.TANGGAL + a.UPDATED_AT)));
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><button class="btn primary" id="rvNew">Simpan snapshot revisi</button><span class="small muted">Snapshot menyimpan volume & harga saat ini sebagai pembanding.</span></div>
    ${revs.length ? `<div class="card"><div class="list">${revs.map(r => `<div class="li"><div class="main"><div class="t1">${esc(r.NAMA_REV)}</div><div class="t2">${fmtDate(r.TANGGAL)} · ${N(r.JUMLAH_ITEM)} item · ${esc(r.CATATAN || '')}</div></div><strong class="num">${rp(r.TOTAL)}</strong><button class="btn sm" data-cmp="${r.ID}">Bandingkan</button></div>`).join('')}</div></div><div id="rvCmp" style="margin-top:12px"></div>` : '<div class="card empty"><p>Belum ada revisi tersimpan.</p></div>'}`;
  $('#rvNew', el).onclick = () => {
    const name = 'Rev-' + revs.length;
    openModal({
      title: 'Simpan revisi', body: `<div class="form-grid">${inp('NAMA_REV', 'Nama revisi', name)}${txt('CATATAN', 'Catatan perubahan', '', { rows: 2 })}</div>`,
      foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="rvOk">Simpan</button>`,
      onOpen: (m2, m) => {
        $('#rvOk', m2).onclick = () => {
          const d = formData(m2), rid = uid('R');
          save('REVISI', { ID: rid, PROYEK_ID: p.ID, NAMA_REV: d.NAMA_REV, TANGGAL: todayStr(), TOTAL: rk.total, JUMLAH_ITEM: rk.itemCount, CATATAN: d.CATATAN }, { silent: true });
          rk.groups.forEach(g => {
            if (g.row.ID !== '_') save('REVISI_ITEM', { ID: uid('RI'), REVISI_ID: rid, PROYEK_ID: p.ID, ITEM_ID: g.row.ID, TIPE_BARIS: 'DIVISI', DIVISI: g.row.URAIAN, KODE: g.row.KODE, URAIAN: g.row.URAIAN, JUMLAH: Math.round(g.subtotal) }, { silent: true });
            g.items.forEach(it => { const c = rk.computed[it.ID]; save('REVISI_ITEM', { ID: uid('RI'), REVISI_ID: rid, PROYEK_ID: p.ID, ITEM_ID: it.ID, TIPE_BARIS: 'ITEM', DIVISI: g.row.URAIAN, KODE: it.KODE, URAIAN: it.URAIAN, SATUAN: it.SATUAN, VOLUME: num(it.VOLUME), HARGA_SATUAN: c.hs, JUMLAH: c.jumlah }, { silent: true }); });
          });
          scheduleSync(); m.close(); toast('Revisi disimpan', 'ok'); rerenderSoft();
        };
      }
    });
  };
  $$('[data-cmp]', el).forEach(b => b.onclick = () => {
    const rv = getRow('REVISI', b.dataset.cmp);
    const old = rows('REVISI_ITEM').filter(x => x.REVISI_ID === rv.ID && norm(x.TIPE_BARIS) === 'ITEM');
    const cur = []; rk.groups.forEach(g => g.items.forEach(it => cur.push({ it, c: rk.computed[it.ID] })));
    const lines = [];
    cur.forEach(({ it, c }) => {
      const o = old.find(x => x.ITEM_ID === it.ID);
      if (!o) lines.push({ t: 'Baru', u: it.URAIAN, a: 0, b: c.jumlah, d: `${Nmax(it.VOLUME, 3)} ${it.SATUAN}` });
      else if (Math.round(num(o.JUMLAH)) !== Math.round(c.jumlah)) lines.push({ t: 'Berubah', u: it.URAIAN, a: num(o.JUMLAH), b: c.jumlah, d: `vol ${Nmax(o.VOLUME, 3)} → ${Nmax(it.VOLUME, 3)} · HS ${rp(o.HARGA_SATUAN)} → ${rp(c.hs)}` });
    });
    old.forEach(o => { if (!cur.some(x => x.it.ID === o.ITEM_ID)) lines.push({ t: 'Dihapus', u: o.URAIAN, a: num(o.JUMLAH), b: 0, d: '' }); });
    $('#rvCmp', el).innerHTML = `<div class="card"><h3>${esc(rv.NAMA_REV)} → saat ini</h3><dl class="kv" style="margin-bottom:10px"><dt>Total ${esc(rv.NAMA_REV)}</dt><dd>${rp(rv.TOTAL)}</dd><dt>Total saat ini</dt><dd>${rp(rk.total)}</dd><dt>Selisih</dt><dd class="${rk.total - num(rv.TOTAL) < 0 ? 'neg' : ''}">${rp(rk.total - num(rv.TOTAL))}</dd></dl>
      ${lines.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Status</th><th>Uraian</th><th class="num">Sebelum</th><th class="num">Sesudah</th><th class="num">Selisih</th></tr></thead><tbody>${lines.map(l => `<tr><td>${badge(l.t, l.t === 'Baru' ? 'success' : l.t === 'Dihapus' ? 'danger' : 'warning')}</td><td>${esc(l.u)}<div class="small muted">${esc(l.d)}</div></td><td class="num">${rpCell(l.a)}</td><td class="num">${rpCell(l.b)}</td><td class="num">${rpCell(l.b - l.a)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="muted">Tidak ada perubahan item.</div>'}</div>`;
  });
}

/* ---------- Tab Info & File ---------- */
function tabInfo(p, rk, el) {
  const files = rows('FILES').filter(f => f.PROYEK_ID === p.ID).sort((a, b) => String(b.UPDATED_AT).localeCompare(String(a.UPDATED_AT)));
  const prof = profilOf(p.PROFIL_ID);
  el.innerHTML = `<div class="grid md2"><div class="card"><div class="row between"><h3>Data proyek</h3><button class="btn sm" id="eP">Ubah</button></div><dl class="kv">
      <dt>Kode</dt><dd>${esc(p.KODE)}</dd><dt>Pemilik</dt><dd>${esc(p.PEMILIK || '-')}</dd><dt>Lokasi</dt><dd>${esc(p.LOKASI || '-')}</dd>
      <dt>Jenis</dt><dd>${esc(p.JENIS_BANGUNAN || '-')}</dd><dt>Mode</dt><dd>${esc(p.MODE || 'RUMAH')}</dd><dt>Lantai</dt><dd>${N(p.JUMLAH_LANTAI)}</dd>
      <dt>Luas tanah</dt><dd>${N(p.LUAS_TANAH, 2)} m²</dd><dt>Luas bangunan</dt><dd>${N(p.LUAS_BANGUNAN, 2)} m²</dd><dt>Wilayah & tahun harga</dt><dd>${esc(p.WILAYAH)} ${esc(p.TAHUN_ANGGARAN)}</dd>
      <dt>Profil (kop)</dt><dd>${esc(prof ? prof.NAMA : '-')}</dd><dt>Nilai kontrak</dt><dd>${num(p.NILAI_KONTRAK) ? rp(p.NILAI_KONTRAK) : '-'}</dd><dt>Status</dt><dd>${badge(p.STATUS || 'DRAFT')}</dd></dl>
      ${p.FOLDER_ID ? `<a class="btn sm" style="margin-top:12px" target="_blank" rel="noopener" href="https://drive.google.com/drive/folders/${esc(p.FOLDER_ID)}">Buka folder Drive</a>` : '<div class="small muted" style="margin-top:12px">Folder Drive dibuat otomatis saat file pertama diunggah.</div>'}</div>
    <div class="card"><h3>File tersimpan di Drive</h3>${files.length ? `<div class="list">${files.map(f => `<div class="li"><div class="main"><div class="t1" style="word-break:break-all">${esc(f.NAMA_FILE)}</div><div class="t2">${esc(f.JENIS)} · ${fmtDate(f.TANGGAL)} · ${N(num(f.UKURAN) / 1024)} KB</div></div><a class="btn sm" href="${esc(f.URL)}" target="_blank" rel="noopener">Buka</a><button class="btn sm ghost" data-share="${esc(f.FILE_ID)}">Link</button></div>`).join('')}</div>` : '<div class="muted small">Belum ada file. PDF RAB & penawaran otomatis tersimpan di sini.</div>'}
      ${S.uploads.filter(u => u.proyekId === p.ID).length ? `<div class="notice warning" style="margin-top:10px">${S.uploads.filter(u => u.proyekId === p.ID).length} file menunggu diunggah (offline).</div>` : ''}</div></div>`;
  $('#eP', el).onclick = () => projectForm(p);
  $$('[data-share]', el).forEach(b => b.onclick = () => shareDriveLink(b.dataset.share));
}
async function shareDriveLink(fileId) {
  if (!await confirmBox('Buat link view-only untuk file ini?\nSiapa pun yang menerima link bisa melihat file (tidak bisa mengubah).', { ok: 'Buat link' })) return;
  try { const r = await api('share', { fileId }); await navigator.clipboard.writeText(r.url).catch(() => {}); toast('Link disalin ke clipboard', 'ok'); return r.url; }
  catch (e) { toast('Gagal: ' + e.message, 'err'); }
}
/* ============================================================
 * Master: AHSP | HSD | Import | Profil | Pengaturan | Laporan
 * ============================================================ */
const MTABS = [['ahsp', 'AHSP'], ['hsd', 'Harga Dasar'], ['import', 'Import'], ['profil', 'Profil Perusahaan'], ['pengaturan', 'Pengaturan'], ['laporan', 'Laporan Library']];
route('/master', () => masterView('ahsp'));
route('/master/:tab', tab => masterView(tab));
function masterView(tab) {
  setTitle('Master', 'Library & pengaturan');
  view(`<div class="page-head"><h2>Master</h2></div><div class="tabs">${MTABS.map(([k, t]) => `<button data-tab="${k}" class="${k === tab ? 'active' : ''}">${t}</button>`).join('')}</div><div id="tabBody"></div>`);
  $$('.tabs button').forEach(b => b.onclick = () => { location.hash = '#/master/' + b.dataset.tab; });
  ({ ahsp: mAhsp, hsd: mHsd, import: mImport, profil: mProfil, pengaturan: mSetting, laporan: mLaporan }[tab] || mAhsp)($('#tabBody'));
}
const defaultCtx = () => {
  const p = getRow('PROYEK', S.meta.activeProject);
  return p ? ctxOf(p) : { wilayah: setting('WILAYAH_DEFAULT'), tahun: new Date().getFullYear(), overhead: num(setting('OVERHEAD_PCT')) };
};

/* ---------- AHSP ---------- */
function mAhsp(el) {
  const ctx = defaultCtx();
  const st = { q: '', jenis: '', status: '', lim: 60 };
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><div class="search" style="flex:1;min-width:220px">${ICON.search}<input class="input" id="aq" placeholder="Cari kode, uraian, tag…"></div>
      <select class="input" id="aj" style="width:auto">${[['', 'Semua jenis'], 'RESMI', 'MODIFIKASI', 'CUSTOM', 'HARGA PASAR'].map(o => { const [v, t] = Array.isArray(o) ? o : [o, o]; return `<option value="${v}">${t}</option>`; }).join('')}</select>
      <select class="input" id="as" style="width:auto"><option value="">Semua status</option><option>TERVERIFIKASI</option><option>PERLU CEK</option></select>
      <button class="btn primary" id="aNew">${ICON.plus} AHSP custom</button></div>
    <div class="small muted" style="margin-bottom:8px">Harga dihitung untuk ${esc(ctx.wilayah)} ${ctx.tahun}, overhead ${pct(ctx.overhead, 0)} (proyek aktif).</div>
    <div id="alist"></div>`;
  const draw = () => {
    let list = rows('AHSP_HEADER');
    if (st.q) { const res = Calc.search(idx(), st.q, { limit: 200 }); const ex = list.filter(h => norm(h.KODE_AHSP).includes(norm(st.q))); const set = new Set(ex.map(h => h.KODE_AHSP).concat(res.map(r => r.kode))); list = list.filter(h => set.has(h.KODE_AHSP)); }
    if (st.jenis) list = list.filter(h => norm(h.JENIS) === st.jenis);
    if (st.status) list = list.filter(h => norm(h.STATUS) === st.status);
    list.sort((a, b) => String(a.DIVISI + a.KODE_AHSP).localeCompare(String(b.DIVISI + b.KODE_AHSP)));
    const shown = list.slice(0, st.lim);
    $('#alist', el).innerHTML = list.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Kode</th><th>Uraian</th><th>Sat</th><th>Jenis</th><th>Status</th><th class="num">HSP</th></tr></thead><tbody>${shown.map(h => {
      const a = Calc.analyze(idx(), h.KODE_AHSP, ctx);
      return `<tr class="clickable" data-k="${esc(h.KODE_AHSP)}"><td>${esc(h.KODE_AHSP)}</td><td>${esc(h.URAIAN)}<div class="small muted">${esc(h.DIVISI || '')}</div></td><td>${esc(h.SATUAN)}</td><td>${badge(h.JENIS || '-')}</td><td>${badge(h.STATUS || '-')}</td><td class="num">${rpCell(a.HSP)}${a.missing.length ? '<div class="small neg">harga dasar kosong</div>' : ''}</td></tr>`;
    }).join('')}</tbody></table></div>${list.length > st.lim ? `<div class="row" style="justify-content:center;margin-top:10px"><button class="btn sm" id="aMore">Tampilkan lebih banyak (${list.length - st.lim} lagi)</button></div>` : ''}` : '<div class="card empty"><p>Tidak ada AHSP. Import library di tab Import.</p></div>';
    $('#aMore', el) && ($('#aMore', el).onclick = () => { st.lim += 100; draw(); });
  };
  $('#aq', el).oninput = debounce(e => { st.q = e.target.value.trim(); st.lim = 60; draw(); }, 250);
  $('#aj', el).onchange = e => { st.jenis = e.target.value; draw(); };
  $('#as', el).onchange = e => { st.status = e.target.value; draw(); };
  $('#aNew', el).onclick = () => { const k = nextCustomCode(); save('AHSP_HEADER', { ID: uid('H'), KODE_AHSP: k, URAIAN: '', SATUAN: '', JENIS: 'CUSTOM', STATUS: 'PERLU CEK', SCOPE: 'UMUM', DIBUAT_OLEH: 'PWA', TANGGAL: todayStr(), VERSI: 'CUSTOM' }); ahspEditor(k, { isNew: true, onSaved: () => draw(), onCancel: () => { const h = ahspHead(k); if (h && !h.URAIAN) remove('AHSP_HEADER', h.ID); } }); };
  $('#alist', el).onclick = e => { const tr = e.target.closest('[data-k]'); if (tr) ahspDetail(tr.dataset.k, draw); };
  draw();
}
function ahspDetail(kode, onChange) {
  const h = ahspHead(kode); if (!h) return;
  const ctx = defaultCtx(), dv = ahspDeviation(kode, getRow('PROYEK', S.meta.activeProject));
  openModal({
    title: h.KODE_AHSP + ' · ' + h.URAIAN, wide: true,
    body: `<div class="row" style="margin-bottom:8px">${badge(h.JENIS || '-')} ${badge(h.STATUS || '-')} ${badge(h.SCOPE || 'UMUM', 'primary')} <span class="small muted">versi ${esc(h.VERSI || '-')} · ${esc(h.SATUAN)}</span></div>
      <dl class="kv" style="grid-template-columns:auto 1fr;margin-bottom:8px"><dt>Divisi</dt><dd style="text-align:left">${esc(h.DIVISI || '-')}</dd><dt>Sumber</dt><dd style="text-align:left">${esc(h.SUMBER || '-')}</dd>${h.KODE_REFERENSI ? `<dt>Referensi</dt><dd style="text-align:left">${esc(h.KODE_REFERENSI)}</dd>` : ''}${h.DASAR_PERHITUNGAN ? `<dt>Dasar perhitungan</dt><dd style="text-align:left">${esc(h.DASAR_PERHITUNGAN)}</dd>` : ''}<dt>Tag / sinonim</dt><dd style="text-align:left">${esc([h.TAG, h.SINONIM].filter(Boolean).join(' · ') || '-')}</dd></dl>
      ${analisaTable(dv.a, ctx)}
      ${dv.sim.list.length ? `<div class="small" style="margin-top:10px"><strong>Pembanding sejenis</strong> (rata-rata ${rp(dv.sim.avg)}${dv.over ? `, <span class="neg">menyimpang ${pct(dv.dev, 1)}</span>` : ''})${dv.sim.list.map(x => `<div class="row between"><span>${esc(x.header.KODE_AHSP)} ${esc(x.header.URAIAN)}</span><span class="num">${rp(x.HSP)}</span></div>`).join('')}</div>` : ''}`,
    foot: `${!isOfficial(h) ? `<button class="btn danger" id="hDel">Hapus</button>` : ''}<span class="spacer"></span>
      <button class="btn" id="hTag">Tag & sinonim</button>
      ${norm(h.STATUS) !== 'TERVERIFIKASI' ? `<button class="btn" id="hVer">Tandai terverifikasi</button>` : ''}
      <button class="btn" id="hDup">Duplikat & modifikasi</button>
      ${!isOfficial(h) ? `<button class="btn primary" id="hEdit">Ubah analisa</button>` : ''}`,
    onOpen: (el, m) => {
      $('#hDup', el).onclick = () => { const nk = duplicateAhsp(kode); m.close(); ahspEditor(nk, { onSaved: onChange }); };
      $('#hEdit', el) && ($('#hEdit', el).onclick = () => { m.close(); ahspEditor(kode, { onSaved: onChange }); });
      $('#hVer', el) && ($('#hVer', el).onclick = async () => { if (!await confirmBox('Tandai AHSP ini TERVERIFIKASI?\nPastikan koefisien sudah dicek terhadap lampiran resmi / dasar perhitungan.')) return; save('AHSP_HEADER', { ID: h.ID, STATUS: 'TERVERIFIKASI' }); m.close(); onChange && onChange(); });
      $('#hTag', el).onclick = () => openModal({
        title: 'Tag & sinonim ' + h.KODE_AHSP,
        body: `<div class="form-grid">${inp('TAG', 'Tag (pisahkan koma)', h.TAG, { ph: 'roster, ventilasi, bata hias' })}${inp('SINONIM', 'Sinonim (grup "a = b", pisahkan ;)', h.SINONIM, { ph: 'hebel = bata ringan = aac; ubin = keramik' })}</div><div class="small muted" style="margin-top:8px">Tag & sinonim membantu pencarian referensi saat menambah item. Boleh diisi juga untuk AHSP resmi (tidak mengubah koefisien).</div>`,
        foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="tOk">Simpan</button>`,
        onOpen: (e2, m2) => { $('#tOk', e2).onclick = () => { save('AHSP_HEADER', Object.assign({ ID: h.ID }, formData(e2))); m2.close(); m.close(); onChange && onChange(); }; }
      });
      $('#hDel', el) && ($('#hDel', el).onclick = async () => {
        const used = rows('RAB_ITEM').filter(r => norm(r.KODE_AHSP) === norm(kode)).length;
        if (!await confirmBox(`Hapus AHSP ${kode}?${used ? `\nMasih dipakai ${used} item RAB — item tersebut akan kehilangan harga.` : ''}`, { danger: true, ok: 'Hapus' })) return;
        ahspKomp(kode).forEach(k => save('AHSP_KOMPONEN', { ID: k.ID, DELETED: 'YA' }, { silent: true }));
        remove('AHSP_HEADER', h.ID); m.close(); onChange && onChange();
      });
    }
  });
}
/* Editor AHSP custom/modifikasi (AHSP resmi read-only) */
function ahspEditor(kode, { project, onSaved, onCancel, isNew } = {}) {
  const h = ahspHead(kode); if (!h) return;
  if (isOfficial(h)) return toast('AHSP resmi tidak bisa diubah. Gunakan Duplikat & modifikasi.', 'err');
  const ctx = project ? ctxOf(project) : defaultCtx();
  let comps = ahspKomp(kode).map(c => Object.assign({}, c));
  const removed = [];
  const hsd = hsdAll();
  let saved = false;
  openModal({
    title: (isNew ? 'AHSP baru ' : 'Ubah AHSP ') + kode, wide: true,
    body: `<form id="ah" class="form-grid c3">
        ${inp('URAIAN', 'Uraian pekerjaan', h.URAIAN, { cls: 'span2', req: true })}${inp('SATUAN', 'Satuan', h.SATUAN, { req: true })}
        ${inp('DIVISI', 'Divisi', h.DIVISI, { list: 'divList', cls: 'span2' })}${sel('JENIS', 'Jenis', h.JENIS || 'CUSTOM', ['MODIFIKASI', 'CUSTOM'])}
        ${txt('DASAR_PERHITUNGAN', 'Dasar perhitungan (wajib)', h.DASAR_PERHITUNGAN, { cls: 'span3', rows: 2, hint: 'mis. produktivitas tukang 6 m²/hari; kebutuhan bahan dari brosur pabrik; modifikasi dari ' + (h.KODE_REFERENSI || 'AHSP referensi') })}
        ${inp('TAG', 'Tag', h.TAG, { cls: 'span2' })}${sel('SCOPE', 'Simpan untuk', h.SCOPE || 'UMUM', [['UMUM', 'Library umum'], ['PROYEK', 'Hanya proyek ini']])}
      </form><datalist id="divList">${rows('DIVISI').map(d => `<option value="${esc(d.NAMA)}">`).join('')}</datalist>
      <datalist id="hsdList">${hsd.map(x => `<option value="${esc(x.KODE_HSD)}">${esc(x.NAMA)} · ${esc(x.SATUAN)} · ${esc(x.WILAYAH)} ${esc(x.TAHUN)}</option>`).join('')}</datalist>
      <h3 style="margin:16px 0 8px;font-size:15px">Komponen</h3><div id="crows"></div>
      <div class="row" style="margin-top:8px"><button type="button" class="btn sm" data-add="Upah">${ICON.plus} Upah</button><button type="button" class="btn sm" data-add="Bahan">${ICON.plus} Bahan</button><button type="button" class="btn sm" data-add="Alat">${ICON.plus} Alat</button></div>
      <div id="apv" style="margin-top:12px"></div>`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="ahOk">Simpan AHSP</button>`,
    onClose: () => { if (!saved && onCancel) onCancel(); },
    onOpen: (el, m) => {
      const hsdName = k => { const x = hsd.find(r => norm(r.KODE_HSD) === norm(k)); return x ? x : null; };
      const draw = () => {
        $('#crows', el).innerHTML = comps.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Jenis</th><th>Kode HSD</th><th>Nama</th><th>Sat</th><th class="num">Koefisien</th><th></th></tr></thead><tbody>${comps.map((c, i) => `<tr data-i="${i}">
          <td><select class="input" data-f="JENIS_KOMPONEN" style="width:96px">${['Upah', 'Bahan', 'Alat'].map(j => `<option ${norm(c.JENIS_KOMPONEN) === norm(j) ? 'selected' : ''}>${j}</option>`).join('')}</select></td>
          <td><input class="input" data-f="KODE_HSD" list="hsdList" value="${esc(c.KODE_HSD)}" style="min-width:110px"></td>
          <td><input class="input" data-f="NAMA_KOMPONEN" value="${esc(c.NAMA_KOMPONEN)}" style="min-width:140px"></td>
          <td><input class="input" data-f="SATUAN" value="${esc(c.SATUAN)}" style="width:70px"></td>
          <td><input class="input num" data-f="KOEFISIEN" inputmode="decimal" value="${esc(String(c.KOEFISIEN ?? '').replace('.', ','))}" style="width:100px"></td>
          <td><button type="button" class="btn ghost icon" data-rm="${i}">${ICON.close}</button></td></tr>`).join('')}</tbody></table></div>` : '<div class="muted small">Belum ada komponen.</div>';
        preview();
      };
      const preview = () => {
        // hitung sementara pakai data form
        const tmp = Object.assign({}, S.t);
        const hk = 'TMP-PREVIEW';
        tmp.AHSP_HEADER = rows('AHSP_HEADER').concat([{ ID: 'tmp', KODE_AHSP: hk, URAIAN: 'tmp', SATUAN: $('[name=SATUAN]', el).value, DIVISI: $('[name=DIVISI]', el).value }]);
        tmp.AHSP_KOMPONEN = rows('AHSP_KOMPONEN').concat(comps.map((c, i) => Object.assign({}, c, { KODE_AHSP: hk, URUTAN: i + 1 })));
        const I = Calc.buildIndex(tmp);
        const a = Calc.analyze(I, hk, ctx);
        const sim = Calc.similar(I, hk, ctx, 3);
        const dev = sim.avg ? (a.HSP - sim.avg) / sim.avg * 100 : 0;
        $('#apv', el).innerHTML = analisaTable(a, ctx) + (sim.avg && Math.abs(dev) > num(setting('AMBANG_DEVIASI_HSP')) ? `<div class="notice warning" style="margin-top:8px">${ICON.warn}<div>HSP ${rp(a.HSP)} menyimpang ${pct(dev, 1)} dari rata-rata ${sim.list.length} AHSP sejenis (${rp(sim.avg)}).</div></div>` : '') + (a.missing.length ? `<div class="notice danger" style="margin-top:8px">Harga dasar kosong: ${esc(a.missing.join(', '))}. Lengkapi di Master → Harga Dasar.</div>` : '');
      };
      draw();
      $('#crows', el).addEventListener('change', e => {
        const tr = e.target.closest('tr[data-i]'); if (!tr) return;
        const c = comps[+tr.dataset.i], f = e.target.dataset.f;
        c[f] = f === 'KOEFISIEN' ? parseNum(e.target.value) : e.target.value;
        if (f === 'KODE_HSD') { const x = hsdName(c.KODE_HSD); if (x) { c.NAMA_KOMPONEN = x.NAMA; c.SATUAN = x.SATUAN; c.JENIS_KOMPONEN = x._j; } }
        setTimeout(draw, 0);
      });
      $('#crows', el).addEventListener('click', e => { const b = e.target.closest('[data-rm]'); if (b) { removed.push(comps[+b.dataset.rm]); comps.splice(+b.dataset.rm, 1); draw(); } });
      $$('[data-add]', el).forEach(b => b.onclick = () => { comps.push({ ID: uid('K'), KODE_AHSP: kode, JENIS_KOMPONEN: b.dataset.add, KODE_HSD: '', NAMA_KOMPONEN: '', SATUAN: '', KOEFISIEN: '' }); draw(); });
      ['SATUAN', 'DIVISI'].forEach(n => $(`[name=${n}]`, el).addEventListener('input', debounce(preview, 300)));
      $('#ahOk', el).onclick = () => {
        const d = formData($('#ah', el));
        if (!d.URAIAN || !d.SATUAN) return toast('Uraian & satuan wajib diisi', 'err');
        if (!d.DASAR_PERHITUNGAN) return toast('Dasar perhitungan wajib diisi', 'err');
        if (!comps.length) return toast('Tambahkan minimal 1 komponen', 'err');
        if (comps.some(c => !c.KODE_HSD || !(num(c.KOEFISIEN) > 0))) return toast('Setiap komponen wajib punya kode HSD & koefisien > 0', 'err');
        const missingHsd = comps.filter(c => !hsdName(c.KODE_HSD));
        missingHsd.forEach(c => {
          const t = HSD_TABLE[norm(c.JENIS_KOMPONEN)] || 'HSD_BAHAN';
          save(t, { ID: uid('S'), KODE_HSD: c.KODE_HSD, NAMA: c.NAMA_KOMPONEN || c.KODE_HSD, SATUAN: c.SATUAN, WILAYAH: ctx.wilayah, TAHUN: ctx.tahun, HARGA: '', STATUS: 'PERLU HARGA', SUMBER: '', TANGGAL_UPDATE: todayStr() }, { silent: true });
        });
        save('AHSP_HEADER', Object.assign({ ID: h.ID }, d, { PROYEK_ID: d.SCOPE === 'PROYEK' && project ? project.ID : '', STATUS: 'PERLU CEK', TANGGAL: todayStr() }), { silent: true });
        comps.forEach((c, i) => save('AHSP_KOMPONEN', Object.assign({}, c, { KODE_AHSP: kode, URUTAN: i + 1 }), { silent: true }));
        removed.forEach(c => { if (getRow('AHSP_KOMPONEN', c.ID)) save('AHSP_KOMPONEN', { ID: c.ID, DELETED: 'YA' }, { silent: true }); });
        scheduleSync(); saved = true; m.close();
        if (missingHsd.length) toast(`${missingHsd.length} HSD baru dibuat dengan status PERLU HARGA`, '');
        toast('AHSP disimpan', 'ok');
        onSaved && onSaved(kode);
      };
    }
  });
}

/* ---------- HSD ---------- */
function mHsd(el) {
  const st = { t: 'HSD_UPAH', q: '', w: '', th: '' };
  const wil = [...new Set(hsdAll().map(r => r.WILAYAH).concat(rows('WILAYAH').map(w => w.NAMA)).filter(Boolean))];
  const yrs = [...new Set(hsdAll().map(r => r.TAHUN).filter(Boolean))].sort().reverse();
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><div class="seg" id="hSeg"><button data-t="HSD_UPAH" class="active">Upah</button><button data-t="HSD_BAHAN">Bahan</button><button data-t="HSD_ALAT">Alat</button></div>
      <div class="search" style="flex:1;min-width:180px">${ICON.search}<input class="input" id="hq" placeholder="Cari kode / nama"></div>
      <select class="input" id="hw" style="width:auto"><option value="">Semua wilayah</option>${wil.map(w => `<option>${esc(w)}</option>`).join('')}</select>
      <select class="input" id="ht" style="width:auto"><option value="">Semua tahun</option>${yrs.map(y => `<option>${esc(y)}</option>`).join('')}</select>
      <button class="btn primary" id="hNew">${ICON.plus} Harga dasar</button></div><div id="hlist"></div>`;
  const draw = () => {
    let l = rows(st.t);
    if (st.q) { const q = st.q.toLowerCase(); l = l.filter(r => (r.KODE_HSD + ' ' + r.NAMA + ' ' + (r.SINONIM || '')).toLowerCase().includes(q)); }
    if (st.w) l = l.filter(r => r.WILAYAH === st.w);
    if (st.th) l = l.filter(r => String(r.TAHUN) === st.th);
    l.sort((a, b) => String(a.KODE_HSD).localeCompare(String(b.KODE_HSD)) || num(b.TAHUN) - num(a.TAHUN));
    $('#hlist', el).innerHTML = l.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Kode</th><th>Nama</th><th>Sat</th><th>Wilayah</th><th>Tahun</th><th class="num">Harga</th><th>Status</th></tr></thead><tbody>${l.slice(0, 300).map(r => `<tr class="clickable" data-id="${r.ID}"><td>${esc(r.KODE_HSD)}</td><td>${esc(r.NAMA)}<div class="small muted">${esc(r.SUMBER || '')}</div></td><td>${esc(r.SATUAN)}</td><td>${esc(r.WILAYAH)}</td><td>${esc(r.TAHUN)}</td><td class="num">${num(r.HARGA) > 0 ? rpCell(r.HARGA) : '<span class="neg">kosong</span>'}</td><td>${badge(r.STATUS || '-')}</td></tr>`).join('')}</tbody></table></div>${l.length > 300 ? `<div class="small muted" style="margin-top:6px">Menampilkan 300 dari ${l.length}. Persempit dengan pencarian.</div>` : ''}` : '<div class="card empty"><p>Belum ada harga dasar.</p></div>';
  };
  $$('#hSeg button', el).forEach(b => b.onclick = () => { st.t = b.dataset.t; $$('#hSeg button', el).forEach(x => x.classList.toggle('active', x === b)); draw(); });
  $('#hq', el).oninput = debounce(e => { st.q = e.target.value.trim(); draw(); }, 200);
  $('#hw', el).onchange = e => { st.w = e.target.value; draw(); };
  $('#ht', el).onchange = e => { st.th = e.target.value; draw(); };
  $('#hNew', el).onclick = () => hsdEditor(st.t, null, draw);
  $('#hlist', el).onclick = e => { const tr = e.target.closest('[data-id]'); if (tr) hsdEditor(st.t, getRow(st.t, tr.dataset.id), draw); };
  draw();
}
function hsdEditor(t, r, onDone) {
  const isNew = !r;
  r = r || { WILAYAH: setting('WILAYAH_DEFAULT'), TAHUN: new Date().getFullYear(), STATUS: 'OK' };
  openModal({
    title: (isNew ? 'Harga dasar baru · ' : 'Ubah harga dasar · ') + t.replace('HSD_', ''),
    body: `<form class="form-grid c2" id="hf">${inp('KODE_HSD', 'Kode HSD', r.KODE_HSD, { req: true })}${inp('NAMA', 'Nama', r.NAMA, { req: true })}${inp('SATUAN', 'Satuan', r.SATUAN)}
      <div class="field"><label>Wilayah</label><input class="input" name="WILAYAH" list="wl" value="${esc(r.WILAYAH || '')}"><datalist id="wl">${rows('WILAYAH').map(w => `<option value="${esc(w.NAMA)}">`).join('')}</datalist></div>
      ${inp('TAHUN', 'Tahun', r.TAHUN, { numType: 'num' })}${inp('HARGA', 'Harga', r.HARGA, { numType: 'rp' })}
      ${inp('SUMBER', 'Sumber harga', r.SUMBER, { cls: 'span2', ph: 'mis. HSPK Kab. Bogor 2026 / survei TB Sinar, 03/11/2026' })}
      ${inp('SINONIM', 'Sinonim nama (untuk pencocokan import)', r.SINONIM, { cls: 'span2', ph: 'mis. PC; semen portland; semen 50kg' })}
      ${sel('STATUS', 'Status', r.STATUS || 'OK', ['OK', 'PERLU HARGA', 'PERLU CEK'])}</form>
      ${!isNew ? '<div class="small muted" style="margin-top:8px">Untuk tahun/wilayah berbeda, gunakan "Simpan sebagai harga baru" agar riwayat harga tetap tersimpan.</div>' : ''}`,
    foot: `${!isNew ? '<button class="btn danger" id="hfDel">Hapus</button><span class="spacer"></span><button class="btn" id="hfNew">Simpan sebagai harga baru</button>' : ''}<button class="btn" data-close>Batal</button><button class="btn primary" id="hfOk">Simpan</button>`,
    onOpen: (el, m) => {
      bindMoneyInputs(el);
      const go = asNew => {
        const d = formData($('#hf', el));
        if (!d.KODE_HSD || !d.NAMA) return toast('Kode & nama wajib diisi', 'err');
        if (num(d.HARGA) > 0 && d.STATUS === 'PERLU HARGA') d.STATUS = 'OK';
        d.TANGGAL_UPDATE = todayStr();
        const dup = rows(t).find(x => x.ID !== (asNew ? '' : r.ID) && norm(x.KODE_HSD) === norm(d.KODE_HSD) && norm(x.WILAYAH) === norm(d.WILAYAH) && String(x.TAHUN) === String(d.TAHUN));
        if (dup) return toast('Sudah ada harga untuk kode, wilayah & tahun yang sama', 'err');
        save(t, Object.assign({ ID: asNew || isNew ? uid('S') : r.ID }, d)); m.close(); onDone && onDone();
      };
      $('#hfOk', el).onclick = () => go(false);
      $('#hfNew', el) && ($('#hfNew', el).onclick = () => go(true));
      $('#hfDel', el) && ($('#hfDel', el).onclick = async () => { if (!await confirmBox('Hapus harga dasar ini?', { danger: true, ok: 'Hapus' })) return; remove(t, r.ID); m.close(); onDone && onDone(); });
    }
  });
}

/* ---------- Import ---------- */
const AHSP_COLS = ['KODE_AHSP', 'URAIAN', 'SATUAN', 'DIVISI', 'JENIS_KOMPONEN', 'KODE_HSD', 'NAMA_KOMPONEN', 'SATUAN_KOMPONEN', 'KOEFISIEN', 'SUMBER'];
function mImport(el) {
  el.innerHTML = `<div class="grid md2">
    <div class="card"><h3>Import library AHSP</h3><p class="small muted" style="margin-top:0">CSV / Excel dengan kolom:<br><code style="font-size:11px">${AHSP_COLS.join(' | ')}</code><br>Satu baris = satu komponen. Header AHSP diambil dari baris pertama tiap KODE_AHSP. JENIS_KOMPONEN: Upah / Bahan / Alat.</p>
      <div class="row"><label class="btn primary">Pilih file<input type="file" id="fA" accept=".csv,.xlsx,.xls" hidden></label><button class="btn ghost" id="tA">Unduh template</button></div></div>
    <div class="card"><h3>Import harga dasar (HSD)</h3><p class="small muted" style="margin-top:0">Dari HSPK/SSH Pemda, survei toko, atau penawaran vendor. Kolom bebas, nanti dipetakan. Nama bahan dicocokkan otomatis ke kode HSD lewat kode/nama/sinonim.</p>
      <div class="row"><label class="btn primary">Pilih file<input type="file" id="fH" accept=".csv,.xlsx,.xls" hidden></label><button class="btn ghost" id="tH">Unduh template</button></div></div>
    <div class="card"><h3>Export library</h3><p class="small muted" style="margin-top:0">Unduh seluruh AHSP + komponen dan harga dasar sebagai Excel (cadangan / dibagikan ke tim).</p><button class="btn" id="xL">Export Excel</button></div>
    <div class="card"><h3>Versi library</h3>${rows('AHSP_VERSI').length ? `<div class="list">${rows('AHSP_VERSI').sort((a, b) => String(b.TANGGAL).localeCompare(String(a.TANGGAL))).map(v => `<div class="li"><div class="main"><div class="t1">${esc(v.VERSI)} ${norm(v.VERSI) === norm(setting('VERSI_AHSP_AKTIF')) ? badge('aktif', 'success') : ''}</div><div class="t2">${fmtDate(v.TANGGAL)} · ${N(v.JUMLAH_AHSP)} AHSP · ${esc(v.SUMBER || '')}</div></div></div>`).join('')}</div>` : '<div class="muted small">Belum ada.</div>'}</div></div>`;
  $('#tA', el).onclick = async () => { const X = await needXlsx(); const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([AHSP_COLS, ['A.2.2.1.1', 'Galian tanah biasa sedalam 1 m', 'm3', 'PEKERJAAN TANAH', 'Upah', 'L.01', 'Pekerja', 'OH', 0.75, 'Lampiran Permen PUPR 8/2023 hal. …'], ['A.2.2.1.1', '', '', '', 'Upah', 'L.04', 'Mandor', 'OH', 0.025, '']]), 'AHSP'); X.writeFile(wb, 'TEMPLATE_IMPORT_AHSP.xlsx'); };
  $('#tH', el).onclick = async () => { const X = await needXlsx(); const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['KODE_HSD', 'NAMA', 'SATUAN', 'HARGA', 'JENIS'], ['L.01', 'Pekerja', 'OH', 120000, 'Upah'], ['M.15', 'Semen Portland', 'kg', 1500, 'Bahan']]), 'HSD'); X.writeFile(wb, 'TEMPLATE_IMPORT_HSD.xlsx'); };
  $('#fA', el).onchange = async e => { const f = e.target.files[0]; e.target.value = ''; if (f) importAhsp(await readTable(f), f.name); };
  $('#fH', el).onchange = async e => { const f = e.target.files[0]; e.target.value = ''; if (f) importHsd(await readTable(f), f.name); };
  $('#xL', el).onclick = exportLibrary;
}
async function readTable(file) {
  const X = await needXlsx();
  const buf = await readFile(file, 'buffer');
  const wb = X.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return X.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true }).filter(r => r.some(c => String(c).trim() !== ''));
}
function importAhsp(aoa, fname) {
  if (aoa.length < 2) return toast('File kosong', 'err');
  const H = aoa[0].map(h => norm(h).replace(/\s+/g, '_'));
  const miss = ['KODE_AHSP', 'JENIS_KOMPONEN', 'KOEFISIEN'].filter(c => !H.includes(c));
  if (miss.length) return toast('Kolom wajib tidak ada: ' + miss.join(', '), 'err');
  const g = (r, c) => { const i = H.indexOf(c); return i < 0 ? '' : r[i]; };
  const heads = {}, comps = {};
  let lastKode = '';
  aoa.slice(1).forEach(r => {
    const kode = String(g(r, 'KODE_AHSP') || lastKode).trim(); if (!kode) return; lastKode = kode;
    if (!heads[kode]) heads[kode] = { KODE_AHSP: kode, URAIAN: String(g(r, 'URAIAN')).trim(), SATUAN: String(g(r, 'SATUAN')).trim(), DIVISI: String(g(r, 'DIVISI')).trim(), SUMBER: String(g(r, 'SUMBER')).trim() };
    else ['URAIAN', 'SATUAN', 'DIVISI', 'SUMBER'].forEach(c => { if (!heads[kode][c] && g(r, c)) heads[kode][c] = String(g(r, c)).trim(); });
    if (String(g(r, 'KODE_HSD')).trim() || num(parseNum(g(r, 'KOEFISIEN')))) (comps[kode] = comps[kode] || []).push({ JENIS_KOMPONEN: String(g(r, 'JENIS_KOMPONEN')).trim() || 'Bahan', KODE_HSD: String(g(r, 'KODE_HSD')).trim(), NAMA_KOMPONEN: String(g(r, 'NAMA_KOMPONEN')).trim(), SATUAN: String(g(r, 'SATUAN_KOMPONEN')).trim(), KOEFISIEN: num(parseNum(g(r, 'KOEFISIEN'))) });
  });
  const kodes = Object.keys(heads);
  const existing = {}; rows('AHSP_HEADER').forEach(h => existing[norm(h.KODE_AHSP)] = h);
  const baru = [], ubah = [], sama = [];
  kodes.forEach(k => {
    const e = existing[norm(k)];
    if (!e) return baru.push(k);
    const old = ahspKomp(k).map(c => norm(c.KODE_HSD) + ':' + num(c.KOEFISIEN)).sort().join('|');
    const nw = (comps[k] || []).map(c => norm(c.KODE_HSD) + ':' + num(c.KOEFISIEN)).sort().join('|');
    (old === nw ? sama : ubah).push(k);
  });
  const official = rows('AHSP_HEADER').filter(h => isOfficial(h));
  const hilang = official.filter(h => !heads[h.KODE_AHSP]).map(h => h.KODE_AHSP);
  const hsdKnown = {}; hsdAll().forEach(x => hsdKnown[norm(x.KODE_HSD)] = 1);
  const hsdBaru = {}; kodes.forEach(k => (comps[k] || []).forEach(c => { if (c.KODE_HSD && !hsdKnown[norm(c.KODE_HSD)]) hsdBaru[norm(c.KODE_HSD)] = c; }));
  const diffList = ubah.slice(0, 20).map(k => {
    const o = ahspKomp(k), n = comps[k] || [];
    const keys = [...new Set(o.map(c => norm(c.KODE_HSD)).concat(n.map(c => norm(c.KODE_HSD))))];
    return `<div class="small" style="margin-top:6px"><strong>${esc(k)}</strong> ${keys.map(kk => { const a = o.find(c => norm(c.KODE_HSD) === kk), b = n.find(c => norm(c.KODE_HSD) === kk); const av = a ? num(a.KOEFISIEN) : null, bv = b ? num(b.KOEFISIEN) : null; return av === bv ? '' : `<span class="badge warning">${esc(kk)}: ${av === null ? '–' : Nmax(av)} → ${bv === null ? '–' : Nmax(bv)}</span>`; }).join(' ')}</div>`;
  }).join('');
  openModal({
    title: 'Pratinjau import AHSP · ' + fname, wide: true,
    body: `<div class="grid k4"><div class="card tight kpi"><div class="label">AHSP baru</div><div class="value">${baru.length}</div></div><div class="card tight kpi"><div class="label">Berubah koefisien</div><div class="value">${ubah.length}</div></div><div class="card tight kpi"><div class="label">Sama</div><div class="value">${sama.length}</div></div><div class="card tight kpi"><div class="label">HSD baru (perlu harga)</div><div class="value">${Object.keys(hsdBaru).length}</div></div></div>
      ${diffList ? `<div style="margin-top:12px"><strong>Perubahan koefisien</strong>${diffList}${ubah.length > 20 ? `<div class="small muted">+${ubah.length - 20} lainnya</div>` : ''}</div>` : ''}
      ${hilang.length ? `<div class="notice warning" style="margin-top:12px">${hilang.length} AHSP resmi lama tidak ada di file ini (tidak dihapus, tetap tersimpan): ${esc(hilang.slice(0, 10).join(', '))}${hilang.length > 10 ? '…' : ''}</div>` : ''}
      <form class="form-grid c2" id="imf" style="margin-top:14px">${inp('VERSI', 'Nama versi', 'SE ' + new Date().getFullYear(), { req: true, hint: 'mis. "Permen 8-2023" atau "SE 2025"' })}${sel('JENIS', 'Jenis data', 'RESMI', [['RESMI', 'RESMI (lampiran Permen/SE)'], ['CUSTOM', 'CUSTOM (analisa sendiri)']])}
        ${inp('SUMBER', 'Sumber default (jika kolom SUMBER kosong)', 'Lampiran Permen PUPR 8/2023', { cls: 'span2' })}
        ${sel('STATUS', 'Status awal', 'PERLU CEK', [['PERLU CEK', 'PERLU CEK (disarankan, verifikasi sampel dulu)'], ['TERVERIFIKASI', 'TERVERIFIKASI']])}
        <label class="check"><input type="checkbox" name="AKTIF" checked> Jadikan versi aktif</label></form>`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="imOk">Import ${kodes.length} AHSP</button>`,
    onOpen: (el, m) => {
      $('#imOk', el).onclick = async () => {
        const d = formData($('#imf', el));
        if (!d.VERSI) return toast('Isi nama versi', 'err');
        const today = todayStr();
        kodes.forEach(k => {
          const e = existing[norm(k)], h = heads[k];
          if (e && !isOfficial(e) && d.JENIS === 'RESMI') { /* timpa custom bernama sama */ }
          save('AHSP_HEADER', Object.assign({ ID: e ? e.ID : uid('H') }, e || {}, { KODE_AHSP: k, URAIAN: h.URAIAN || (e && e.URAIAN) || k, SATUAN: h.SATUAN || (e && e.SATUAN), DIVISI: h.DIVISI || (e && e.DIVISI) || '', SUMBER: h.SUMBER || d.SUMBER, JENIS: d.JENIS, VERSI: d.VERSI, STATUS: d.STATUS, SCOPE: 'UMUM', DIBUAT_OLEH: 'IMPORT', TANGGAL: today }), { silent: true });
          if (e) ahspKomp(k).forEach(c => save('AHSP_KOMPONEN', { ID: c.ID, DELETED: 'YA' }, { silent: true }));
          (comps[k] || []).forEach((c, i) => save('AHSP_KOMPONEN', Object.assign({ ID: uid('K'), KODE_AHSP: k, URUTAN: i + 1 }, c), { silent: true }));
        });
        Object.values(hsdBaru).forEach(c => save(HSD_TABLE[norm(c.JENIS_KOMPONEN)] || 'HSD_BAHAN', { ID: uid('S'), KODE_HSD: c.KODE_HSD, NAMA: c.NAMA_KOMPONEN || c.KODE_HSD, SATUAN: c.SATUAN, WILAYAH: setting('WILAYAH_DEFAULT'), TAHUN: new Date().getFullYear(), HARGA: '', STATUS: 'PERLU HARGA', TANGGAL_UPDATE: today }, { silent: true }));
        const vid = uid('V');
        save('AHSP_VERSI', { ID: vid, VERSI: d.VERSI, SUMBER: d.SUMBER, TANGGAL: today, JUMLAH_AHSP: kodes.length, JUMLAH_KOMPONEN: Object.values(comps).reduce((a, b) => a + b.length, 0), CATATAN: fname }, { silent: true });
        if (d.AKTIF === 'YA') setSetting('VERSI_AHSP_AKTIF', d.VERSI);
        scheduleSync(); m.close();
        toast(`${kodes.length} AHSP diimport · ${Object.keys(hsdBaru).length} HSD baru perlu harga`, 'ok');
        // arsip CSV ke Drive 00_MASTER
        try {
          const X = await needXlsx();
          const csv = X.utils.sheet_to_csv(X.utils.aoa_to_sheet(aoa));
          queueUpload({ blob: new Blob([csv], { type: 'text/csv' }), name: `AHSP_${d.VERSI.replace(/\W+/g, '-')}_${today.replace(/-/g, '')}.csv`, mime: 'text/csv', kategori: 'MASTER', jenis: 'ARSIP_AHSP', terkait: vid, onDone: { type: 'setField', table: 'AHSP_VERSI', id: vid, field: 'FILE_ARSIP' } });
        } catch (e) { }
        rerenderSoft();
      };
    }
  });
}
function importHsd(aoa, fname) {
  if (aoa.length < 2) return toast('File kosong', 'err');
  const H = aoa[0].map(h => String(h).trim());
  const maps = rows('IMPORT_MAPPING').filter(x => norm(x.TIPE) === 'HSD');
  const guess = k => { const re = { KODE: /kode/i, NAMA: /nama|uraian|bahan|item/i, SATUAN: /sat/i, HARGA: /harga|nilai|price/i, JENIS: /jenis|kategori|kelompok/i }[k]; const i = H.findIndex(h => re.test(h)); return i; };
  let map = { KODE: guess('KODE'), NAMA: guess('NAMA'), SATUAN: guess('SATUAN'), HARGA: guess('HARGA'), JENIS: guess('JENIS') };
  const colSel = (k, label) => `<div class="field"><label>${label}</label><select class="input" data-map="${k}"><option value="-1">— tidak ada —</option>${H.map((h, i) => `<option value="${i}" ${map[k] === i ? 'selected' : ''}>${esc(h || 'Kolom ' + (i + 1))}</option>`).join('')}</select></div>`;
  openModal({
    title: 'Import harga dasar · ' + fname, wide: true,
    body: `${maps.length ? `<div class="field" style="margin-bottom:12px"><label>Pakai mapping tersimpan</label><select class="input" id="mSaved"><option value="">—</option>${maps.map(x => `<option value="${x.ID}">${esc(x.NAMA)}</option>`).join('')}</select></div>` : ''}
      <div class="form-grid c3" id="mapf">${colSel('KODE', 'Kolom kode')}${colSel('NAMA', 'Kolom nama *')}${colSel('SATUAN', 'Kolom satuan')}${colSel('HARGA', 'Kolom harga *')}${colSel('JENIS', 'Kolom jenis (Upah/Bahan/Alat)')}
        <div class="field"><label>Jenis default</label><select class="input" id="jDef"><option>Bahan</option><option>Upah</option><option>Alat</option></select></div></div>
      <div class="form-grid c3" style="margin-top:12px"><div class="field"><label>Wilayah</label><input class="input" id="iW" list="wl2" value="${esc(setting('WILAYAH_DEFAULT'))}"><datalist id="wl2">${rows('WILAYAH').map(w => `<option value="${esc(w.NAMA)}">`).join('')}</datalist></div>
        ${inp('_th', 'Tahun', new Date().getFullYear(), { numType: 'num' })}${inp('_src', 'Sumber', 'HSPK ' + new Date().getFullYear(), {})}</div>
      <div class="row" style="margin-top:12px"><button class="btn" id="mPrev">Cocokkan & pratinjau</button><span class="spacer"></span>${inp('_mname', '', '', { ph: 'Nama mapping (opsional, untuk disimpan)' }).replace('<label></label>', '')}</div>
      <div id="mRes" style="margin-top:12px"></div>`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="mOk" disabled>Import</button>`,
    onOpen: (el, m) => {
      let plan = null;
      $('#mSaved', el) && ($('#mSaved', el).onchange = e => { const x = getRow('IMPORT_MAPPING', e.target.value); if (!x) return; try { const mm = JSON.parse(x.MAPPING_JSON); Object.keys(mm.cols || {}).forEach(k => { const i = H.indexOf(mm.cols[k]); const s = $(`[data-map="${k}"]`, el); if (s) s.value = i; }); if (mm.jenis) $('#jDef', el).value = mm.jenis; if (mm.sumber) $('[name=_src]', el).value = mm.sumber; } catch (er) { } });
      $('#mPrev', el).onclick = () => {
        $$('[data-map]', el).forEach(s => map[s.dataset.map] = +s.value);
        if (map.NAMA < 0 || map.HARGA < 0) return toast('Pilih kolom nama & harga', 'err');
        const W = $('#iW', el).value.trim(), TH = num(parseNum($('[name=_th]', el).value)), SRC = $('[name=_src]', el).value.trim(), JD = $('#jDef', el).value;
        const all = hsdAll();
        const byKode = {}, byName = {};
        all.forEach(x => { byKode[norm(x.KODE_HSD)] = x; [x.NAMA].concat(String(x.SINONIM || '').split(/[;,]/)).forEach(n => { const k = Calc.tokens(n).join(' '); if (k) byName[k] = byName[k] || x; }); });
        const matched = [], unmatched = [];
        aoa.slice(1).forEach(r => {
          const kode = map.KODE >= 0 ? String(r[map.KODE]).trim() : '', nama = String(r[map.NAMA]).trim(), harga = num(parseNum(r[map.HARGA], true));
          if (!nama && !kode) return;
          const jenis = map.JENIS >= 0 && String(r[map.JENIS]).trim() ? String(r[map.JENIS]).trim() : JD;
          const sat = map.SATUAN >= 0 ? String(r[map.SATUAN]).trim() : '';
          const ref = (kode && byKode[norm(kode)]) || byName[Calc.tokens(nama).join(' ')];
          const rec = { kode, nama, harga, sat, jenis, ref };
          (ref ? matched : unmatched).push(rec);
        });
        plan = { matched, unmatched, W, TH, SRC, JD };
        $('#mRes', el).innerHTML = `<div class="row"><span class="badge success">${matched.length} cocok</span><span class="badge warning">${unmatched.length} belum terpetakan</span></div>
          ${matched.length ? `<div class="table-wrap" style="margin-top:8px;max-height:240px;overflow:auto"><table class="t"><thead><tr><th>Baris file</th><th>→ HSD</th><th class="num">Harga</th></tr></thead><tbody>${matched.slice(0, 100).map(x => `<tr><td>${esc(x.kode)} ${esc(x.nama)}</td><td>${esc(x.ref.KODE_HSD)} · ${esc(x.ref.NAMA)}</td><td class="num">${rpCell(x.harga)}</td></tr>`).join('')}</tbody></table></div>` : ''}
          ${unmatched.length ? `<div style="margin-top:10px"><strong>Belum terpetakan</strong> <span class="small muted">— pilih tindakan per baris</span><div class="table-wrap" style="margin-top:6px;max-height:260px;overflow:auto"><table class="t"><thead><tr><th>Baris file</th><th class="num">Harga</th><th>Tindakan</th></tr></thead><tbody>${unmatched.slice(0, 200).map((x, i) => `<tr><td>${esc(x.kode)} ${esc(x.nama)} <span class="muted">${esc(x.sat)}</span></td><td class="num">${rpCell(x.harga)}</td><td><select class="input" data-u="${i}" style="min-width:200px"><option value="SKIP">Lewati</option><option value="NEW">Buat HSD baru</option>${candidates(x.nama, all).map(a => `<option value="${esc(a.KODE_HSD)}">→ ${esc(a.KODE_HSD)} ${esc(a.NAMA)}</option>`).join('')}</select></td></tr>`).join('')}</tbody></table></div></div>` : ''}`;
        $('#mOk', el).disabled = false;
      };
      $('#mOk', el).onclick = () => {
        if (!plan) return;
        const { W, TH, SRC } = plan, today = todayStr();
        let n = 0, created = 0;
        const put = (kode, nama, sat, jenis, harga) => {
          const t = HSD_TABLE[norm(jenis)] || 'HSD_BAHAN';
          const ex = rows(t).find(x => norm(x.KODE_HSD) === norm(kode) && norm(x.WILAYAH) === norm(W) && String(x.TAHUN) === String(TH));
          save(t, Object.assign({ ID: ex ? ex.ID : uid('S') }, ex || {}, { KODE_HSD: kode, NAMA: nama, SATUAN: sat, WILAYAH: W, TAHUN: TH, HARGA: harga, SUMBER: SRC, STATUS: harga > 0 ? 'OK' : 'PERLU HARGA', TANGGAL_UPDATE: today }), { silent: true });
          n++;
        };
        plan.matched.forEach(x => put(x.ref.KODE_HSD, x.ref.NAMA, x.ref.SATUAN || x.sat, x.ref._j, x.harga));
        $$('[data-u]', el).forEach(s => {
          const x = plan.unmatched[+s.dataset.u];
          if (s.value === 'SKIP') return;
          if (s.value === 'NEW') { put(x.kode || ('HSD-' + uid().slice(-6)), x.nama, x.sat, x.jenis, x.harga); created++; }
          else { const ref = hsdAll().find(a => a.KODE_HSD === s.value); put(ref.KODE_HSD, ref.NAMA, ref.SATUAN || x.sat, ref._j, x.harga); const srow = getRow(ref._t, ref.ID); if (srow && !String(srow.SINONIM || '').toLowerCase().includes(x.nama.toLowerCase())) save(ref._t, { ID: ref.ID, SINONIM: [srow.SINONIM, x.nama].filter(Boolean).join('; ') }, { silent: true }); }
        });
        const mname = $('[name=_mname]', el).value.trim();
        if (mname) save('IMPORT_MAPPING', { ID: uid('MP'), NAMA: mname, TIPE: 'HSD', MAPPING_JSON: JSON.stringify({ cols: Object.fromEntries(Object.entries(map).filter(([, v]) => v >= 0).map(([k, v]) => [k, H[v]])), jenis: plan.JD, sumber: SRC }) }, { silent: true });
        if (W && !rows('WILAYAH').some(w => norm(w.NAMA) === norm(W))) save('WILAYAH', { ID: uid('W'), NAMA: W, AKTIF: 'YA' }, { silent: true });
        scheduleSync(); m.close(); toast(`${n} harga dasar diimport (${created} HSD baru)`, 'ok'); rerenderSoft();
      };
    }
  });
}
/* kandidat HSD terdekat berdasarkan kesamaan kata */
function candidates(nama, all) {
  const t = new Set(Calc.tokens(nama));
  const seen = {};
  return all.map(a => { const at = Calc.tokens(a.NAMA + ' ' + (a.SINONIM || '')); const sc = at.filter(x => t.has(x)).length / Math.max(1, Math.min(t.size, at.length)); return { a, sc }; })
    .filter(x => x.sc > 0 && !seen[x.a.KODE_HSD] && (seen[x.a.KODE_HSD] = 1)).sort((x, y) => y.sc - x.sc).slice(0, 6).map(x => x.a);
}
async function exportLibrary() {
  const X = await needXlsx();
  const wb = X.utils.book_new();
  const A = [AHSP_COLS.concat(['JENIS', 'STATUS', 'VERSI'])];
  rows('AHSP_HEADER').forEach(h => { const ks = ahspKomp(h.KODE_AHSP); (ks.length ? ks : [{}]).forEach((c, i) => A.push([h.KODE_AHSP, i ? '' : h.URAIAN, i ? '' : h.SATUAN, i ? '' : h.DIVISI, c.JENIS_KOMPONEN || '', c.KODE_HSD || '', c.NAMA_KOMPONEN || '', c.SATUAN || '', num(c.KOEFISIEN), i ? '' : h.SUMBER, h.JENIS, h.STATUS, h.VERSI])); });
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(A), 'AHSP');
  const Hs = [['JENIS', 'KODE_HSD', 'NAMA', 'SATUAN', 'WILAYAH', 'TAHUN', 'HARGA', 'SUMBER', 'STATUS']].concat(hsdAll().map(r => [r._j, r.KODE_HSD, r.NAMA, r.SATUAN, r.WILAYAH, r.TAHUN, num(r.HARGA), r.SUMBER, r.STATUS]));
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(Hs), 'HSD');
  X.writeFile(wb, `LIBRARY_AHSP_HSD_${todayStr().replace(/-/g, '')}.xlsx`);
}

/* ---------- Profil perusahaan ---------- */
function mProfil(el) {
  const list = rows('PROFIL_USAHA');
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><button class="btn primary" id="prNew">${ICON.plus} Profil baru</button><span class="small muted">Dipakai sebagai kop di RAB, penawaran, dan dokumen lain.</span></div>
    ${list.length ? `<div class="grid md2">${list.map(p => `<div class="card clickable" data-id="${p.ID}"><div class="row"><div class="logo-prev" style="width:56px;height:56px" data-logo="${esc(p.LOGO_FILE_ID || '')}"></div><div style="flex:1;min-width:0"><div style="font-weight:700">${esc(p.NAMA)} ${norm(p.IS_DEFAULT) === 'YA' ? badge('default', 'primary') : ''}</div><div class="small muted">${esc(p.BENTUK || '')} · ${esc(p.KOTA || '')}</div><div class="small muted">${esc([p.TELEPON, p.EMAIL].filter(Boolean).join(' · '))}</div></div></div>${profilIssues(p).length ? `<div class="small" style="margin-top:8px;color:var(--warning)">${esc(profilIssues(p).join(' · '))}</div>` : ''}</div>`).join('')}</div>` : '<div class="card empty"><p>Belum ada profil perusahaan.</p></div>'}`;
  $('#prNew', el).onclick = () => profilEditor(null);
  $$('[data-id]', el).forEach(c => c.onclick = () => profilEditor(getRow('PROFIL_USAHA', c.dataset.id)));
  $$('[data-logo]', el).forEach(async d => { const src = await imageData(d.dataset.logo); if (src) d.innerHTML = `<img src="${src}" alt="">`; });
}
function profilIssues(p) {
  const out = [];
  if (!p.NAMA) out.push('nama wajib'); if (!p.ALAMAT) out.push('alamat wajib'); if (!p.TELEPON && !p.WHATSAPP && !p.EMAIL) out.push('kontak wajib');
  if (!p.LOGO_FILE_ID) out.push('belum ada logo');
  return out;
}
function profilEditor(p) {
  const isNew = !p;
  p = p || { BENTUK: 'Perorangan', PKP: 'TIDAK', TAMPILKAN_TTD: 'TIDAK', IS_DEFAULT: rows('PROFIL_USAHA').length ? 'TIDAK' : 'YA' };
  const img = { LOGO_FILE_ID: null, TTD_FILE_ID: null, STEMPEL_FILE_ID: null };
  const imgBox = (f, label) => `<div class="field"><label>${label}</label><div class="row"><div class="logo-prev" data-prev="${f}"></div><div class="stack" style="flex:1"><label class="btn sm">Pilih gambar<input type="file" accept="image/png,image/jpeg,image/svg+xml" data-img="${f}" hidden></label><span class="hint" data-hint="${f}">${f === 'LOGO_FILE_ID' ? 'PNG transparan disarankan, minimal 300 px' : 'PNG transparan'}</span></div></div></div>`;
  openModal({
    title: isNew ? 'Profil perusahaan baru' : 'Ubah profil · ' + p.NAMA, wide: true,
    body: `<form id="prf"><div class="form-grid c3">
      ${inp('NAMA', 'Nama perusahaan / usaha', p.NAMA, { cls: 'span2', req: true })}${sel('BENTUK', 'Bentuk usaha', p.BENTUK, ['Perorangan', 'CV', 'PT', 'Lainnya'])}
      ${inp('TAGLINE', 'Tagline (opsional)', p.TAGLINE, { cls: 'span3' })}
      ${txt('ALAMAT', 'Alamat (jalan)', p.ALAMAT, { cls: 'span3', rows: 2 })}
      ${inp('KOTA', 'Kota/Kabupaten', p.KOTA)}${inp('PROVINSI', 'Provinsi', p.PROVINSI)}${inp('KODE_POS', 'Kode pos', p.KODE_POS)}
      ${inp('TELEPON', 'Telepon', p.TELEPON, { type: 'tel' })}${inp('WHATSAPP', 'WhatsApp', p.WHATSAPP, { type: 'tel' })}${inp('EMAIL', 'Email', p.EMAIL, { type: 'email' })}
      ${inp('WEBSITE', 'Website', p.WEBSITE)}${inp('INSTAGRAM', 'Instagram', p.INSTAGRAM, { ph: '@akun' })}${inp('NIB', 'NIB (opsional)', p.NIB)}
      ${inp('NPWP', 'NPWP', p.NPWP)}${sel('PKP', 'Status PKP', p.PKP, [['TIDAK', 'Bukan PKP (tanpa PPN)'], ['YA', 'PKP (kenakan PPN)']])}<div></div>
      ${inp('BANK', 'Bank', p.BANK)}${inp('NO_REKENING', 'No. rekening', p.NO_REKENING)}${inp('ATAS_NAMA', 'Atas nama', p.ATAS_NAMA)}
      ${inp('PJ_NAMA', 'Penanggung jawab', p.PJ_NAMA)}${inp('PJ_JABATAN', 'Jabatan', p.PJ_JABATAN, { ph: 'Direktur / Pemilik' })}<div></div>
      ${imgBox('LOGO_FILE_ID', 'Logo')}${imgBox('TTD_FILE_ID', 'Tanda tangan (opsional)')}${imgBox('STEMPEL_FILE_ID', 'Stempel (opsional)')}
      <label class="check span2"><input type="checkbox" name="TAMPILKAN_TTD" ${norm(p.TAMPILKAN_TTD) === 'YA' ? 'checked' : ''}> Tampilkan tanda tangan & stempel di dokumen</label>
      <label class="check"><input type="checkbox" name="IS_DEFAULT" ${norm(p.IS_DEFAULT) === 'YA' ? 'checked' : ''}> Profil default</label>
    </div></form>
    <div class="notice" style="margin-top:12px">Nomor rekening & tanda tangan tersimpan di Google Sheet/Drive kamu. Tanda tangan hanya muncul di dokumen bila dicentang.</div>`,
    foot: `${!isNew ? '<button class="btn danger" id="prDel">Hapus</button><span class="spacer"></span>' : ''}<button class="btn" data-close>Batal</button><button class="btn primary" id="prOk">Simpan</button>`,
    onOpen: (el, m) => {
      ['LOGO_FILE_ID', 'TTD_FILE_ID', 'STEMPEL_FILE_ID'].forEach(async f => { const src = await imageData(p[f]); if (src) $(`[data-prev="${f}"]`, el).innerHTML = `<img src="${src}" alt="">`; });
      $$('[data-img]', el).forEach(i => i.onchange = async () => {
        const f = i.files[0]; if (!f) return;
        try {
          const r = await compressImage(f, i.dataset.img === 'LOGO_FILE_ID' ? 800 : 600, 0.9, true);
          img[i.dataset.img] = r;
          $(`[data-prev="${i.dataset.img}"]`, el).innerHTML = `<img src="${r.dataUrl}" alt="">`;
          $(`[data-hint="${i.dataset.img}"]`, el).innerHTML = r.w < 300 && i.dataset.img === 'LOGO_FILE_ID' ? `<span class="neg">Logo kecil (${r.w}px) — hasil cetak bisa buram</span>` : `${r.w}×${r.h}px`;
        } catch (e) { toast(e.message, 'err'); }
      });
      $('#prOk', el).onclick = async () => {
        const d = formData($('#prf', el));
        if (!d.NAMA) return toast('Nama wajib diisi', 'err');
        const id = p.ID || uid('PR');
        if (d.IS_DEFAULT === 'YA') rows('PROFIL_USAHA').filter(x => x.ID !== id && norm(x.IS_DEFAULT) === 'YA').forEach(x => save('PROFIL_USAHA', { ID: x.ID, IS_DEFAULT: 'TIDAK' }, { silent: true }));
        const row = save('PROFIL_USAHA', Object.assign({ ID: id }, d));
        for (const f of Object.keys(img)) {
          const r = img[f]; if (!r) continue;
          const localId = 'LOCAL_' + uid();
          await DB.set('blobs', 'IMG_' + localId, r.dataUrl);
          save('PROFIL_USAHA', { ID: id, [f]: localId }, { silent: true });
          const ext = r.mime.includes('png') ? 'png' : r.mime.includes('svg') ? 'svg' : 'jpg';
          await queueUpload({ blob: dataUrlToBlob(r.dataUrl), name: `${f.replace('_FILE_ID', '')}_${(d.NAMA || 'profil').replace(/\W+/g, '-')}_${todayStr().replace(/-/g, '')}.${ext}`, mime: r.mime, kategori: 'PROFIL', jenis: 'PROFIL', terkait: id, onDone: { type: 'profilImg', id, field: f, localId } });
        }
        const issues = profilIssues(getRow('PROFIL_USAHA', id));
        m.close(); toast(issues.filter(x => x !== 'belum ada logo').length ? 'Disimpan. Lengkapi: ' + issues.join(', ') : 'Profil disimpan', issues.length ? '' : 'ok'); rerenderSoft();
      };
      $('#prDel', el) && ($('#prDel', el).onclick = async () => { if (!await confirmBox('Hapus profil ini?', { danger: true, ok: 'Hapus' })) return; remove('PROFIL_USAHA', p.ID); m.close(); rerenderSoft(); });
    }
  });
}
UPLOAD_HOOKS.profilImg = async (d, r) => {
  const cached = await DB.get('blobs', 'IMG_' + d.localId);
  if (cached) await DB.set('blobs', 'IMG_' + r.fileId, cached);
  const cur = getRow('PROFIL_USAHA', d.id);
  if (cur && cur[d.field] === d.localId) save('PROFIL_USAHA', { ID: d.id, [d.field]: r.fileId }, { silent: true });
};

/* ---------- Pengaturan ---------- */
function mSetting(el) {
  const s = settingsObj();
  el.innerHTML = `<div class="grid md2">
    <div class="card"><h3>Koneksi Google Sheet</h3><form id="cf" class="form-grid">${inp('apiUrl', 'URL Web App (Apps Script)', S.meta.apiUrl, { ph: 'https://script.google.com/macros/s/…/exec' })}${inp('token', 'Token API', S.meta.token, { hint: 'Dari menu Sheet: RAB Rumah & Gedung → Tampilkan Token API' })}</form>
      <div class="row" style="margin-top:12px"><button class="btn primary" id="cSave">Simpan & tes koneksi</button><button class="btn" id="cSync">Sinkron penuh</button></div>
      <div class="small muted" style="margin-top:8px">Sinkron terakhir: ${S.meta.lastSync ? new Date(S.meta.lastSync).toLocaleString('id-ID') : '-'} · ${pendingCount()} tertunda${S.lastError ? `<br><span class="neg">${esc(S.lastError)}</span>` : ''}</div></div>
    <div class="card"><h3>Default perhitungan</h3><form id="sf" class="form-grid c2">
      ${inp('OVERHEAD_PCT', 'Biaya umum & keuntungan (%)', s.OVERHEAD_PCT, { numType: 'num' })}${inp('PPN_PCT', 'PPN (%)', s.PPN_PCT, { numType: 'num' })}
      ${inp('KONTINJENSI_PCT', 'Biaya tak terduga (%)', s.KONTINJENSI_PCT, { numType: 'num' })}${sel('PEMBULATAN', 'Pembulatan total', s.PEMBULATAN, [['1', 'Tanpa'], ['1000', 'Ribuan'], ['10000', 'Puluh ribuan'], ['100000', 'Ratus ribuan']])}
      ${inp('WILAYAH_DEFAULT', 'Wilayah default', s.WILAYAH_DEFAULT, { list: 'wl3' })}${sel('MODE_DEFAULT', 'Mode default', s.MODE_DEFAULT, [['RUMAH', 'Mode Rumah'], ['GEDUNG', 'Mode Gedung']])}
      ${inp('AMBANG_DEVIASI_HSP', 'Ambang deviasi HSP custom (±%)', s.AMBANG_DEVIASI_HSP, { numType: 'num' })}${inp('MASA_BERLAKU_HARI', 'Masa berlaku penawaran (hari)', s.MASA_BERLAKU_HARI, { numType: 'num' })}
      ${inp('JAM_KERJA_HARI', 'Jam kerja per hari', s.JAM_KERJA_HARI, { numType: 'num', hint: '1 OH = 8 jam' })}${inp('HARI_KERJA_MINGGU', 'Hari kerja per minggu', s.HARI_KERJA_MINGGU, { numType: 'num' })}
      ${inp('RASIO_KEPALA_TUKANG', 'Maks. tukang per kepala tukang', s.RASIO_KEPALA_TUKANG, { numType: 'num' })}${inp('RASIO_MANDOR_PEKERJA', 'Maks. pekerja+tukang per mandor', s.RASIO_MANDOR_PEKERJA, { numType: 'num' })}
      ${inp('DEVIASI_WASPADA', 'Kurva S: waspada bila tertinggal (%)', s.DEVIASI_WASPADA, { numType: 'num' })}${inp('DEVIASI_KRITIS', 'Kurva S: terlambat bila tertinggal (%)', s.DEVIASI_KRITIS, { numType: 'num' })}
      ${inp('FORMAT_NOMOR_PENAWARAN', 'Format nomor penawaran', s.FORMAT_NOMOR_PENAWARAN, { cls: 'span2', hint: '{NO} {BLN} {THN} {KODE}' })}
      ${inp('TERMIN_DEFAULT', 'Skema termin default', s.TERMIN_DEFAULT, { cls: 'span2' })}
      <label class="check span2"><input type="checkbox" name="TAMPIL_CREDIT" ${norm(s.TAMPIL_CREDIT) === 'YA' ? 'checked' : ''}> Tampilkan "dibuat dengan RAB Rumah & Gedung by DG" kecil di dokumen</label>
      </form><datalist id="wl3">${rows('WILAYAH').map(w => `<option value="${esc(w.NAMA)}">`).join('')}</datalist><button class="btn primary" id="sSave" style="margin-top:12px">Simpan default</button></div>
    <div class="card"><h3>Wilayah harga</h3><div class="list">${rows('WILAYAH').map(w => `<div class="li"><div class="main"><div class="t1">${esc(w.NAMA)}</div><div class="t2">${esc(w.PROVINSI || '')}</div></div><button class="btn sm ghost" data-delw="${w.ID}">Hapus</button></div>`).join('')}</div>
      <div class="row" style="margin-top:8px"><input class="input" id="wNew" placeholder="mis. Kota Depok" style="flex:1"><input class="input" id="wProv" placeholder="Provinsi" style="flex:1"><button class="btn" id="wAdd">Tambah</button></div></div>
    <div class="card"><h3>Template syarat & ketentuan</h3><div class="list">${rows('TEMPLATE_SYARAT').sort((a, b) => num(a.URUTAN) - num(b.URUTAN)).map(t => `<div class="li"><div class="main"><div class="t1">${esc(t.JUDUL)} ${badge(t.JENIS, 'primary')}</div><div class="t2" style="white-space:pre-line">${esc(String(t.ISI).slice(0, 120))}${String(t.ISI).length > 120 ? '…' : ''}</div></div><button class="btn sm" data-ts="${t.ID}">Ubah</button></div>`).join('')}</div></div>
    <div class="card"><h3>Data contoh</h3><p class="small muted" style="margin-top:0">${rows('AHSP_HEADER').filter(h => norm(h.VERSI) === 'CONTOH').length} AHSP & ${hsdAll().filter(r => String(r.SUMBER || '').includes('CONTOH')).length} harga dasar berlabel CONTOH (bukan data resmi). Hapus setelah library resmi di-import.</p><button class="btn danger" id="delC" style="border:1px solid var(--danger)">Hapus semua data CONTOH</button></div>
    <div class="card"><h3>Perangkat ini</h3><p class="small muted" style="margin-top:0">Versi aplikasi ${APP_VERSION}. Data lokal hanya cache — sumber utama tetap Google Sheet.</p><div class="row"><button class="btn" id="dExp">Backup lokal (JSON)</button><button class="btn danger" id="dReset">Reset cache lokal</button></div></div>
  </div>`;
  $('#cSave', el).onclick = async () => {
    const d = formData($('#cf', el));
    S.meta.apiUrl = d.apiUrl.trim(); S.meta.token = d.token.trim(); await saveMeta();
    try { const r = await api('ping'); toast(`Terhubung ke ${r.app} v${r.version}`, 'ok'); await sync({ full: !S.meta.lastSync, quiet: false }); }
    catch (e) { toast('Gagal terhubung: ' + e.message, 'err'); }
    renderSync(); rerenderSoft();
  };
  $('#cSync', el).onclick = () => sync({ full: true, quiet: false });
  $('#sSave', el).onclick = () => { const d = formData($('#sf', el)); Object.entries(d).forEach(([k, v]) => setSetting(k, v)); toast('Default disimpan', 'ok'); };
  $('#wAdd', el).onclick = () => { const n = $('#wNew', el).value.trim(); if (!n) return; save('WILAYAH', { ID: uid('W'), NAMA: n, PROVINSI: $('#wProv', el).value.trim(), AKTIF: 'YA' }); rerenderSoft(); };
  $$('[data-delw]', el).forEach(b => b.onclick = () => { remove('WILAYAH', b.dataset.delw); rerenderSoft(); });
  $$('[data-ts]', el).forEach(b => b.onclick = () => {
    const t = getRow('TEMPLATE_SYARAT', b.dataset.ts);
    openModal({ title: 'Ubah template · ' + t.JUDUL, body: `<div class="form-grid">${inp('JUDUL', 'Judul', t.JUDUL)}${txt('ISI', 'Isi (satu poin per baris)', t.ISI, { rows: 8 })}</div>`, foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="tsOk">Simpan</button>`, onOpen: (e2, m2) => { $('#tsOk', e2).onclick = () => { save('TEMPLATE_SYARAT', Object.assign({ ID: t.ID }, formData(e2))); m2.close(); rerenderSoft(); }; } });
  });
  $('#delC', el).onclick = async () => {
    const used = rows('RAB_ITEM').filter(r => r.KODE_AHSP && norm((ahspHead(r.KODE_AHSP) || {}).VERSI) === 'CONTOH').length;
    if (!await confirmBox(`Hapus semua AHSP & harga dasar CONTOH?${used ? `\n${used} item RAB masih memakai AHSP contoh dan akan kehilangan harga.` : ''}`, { danger: true, ok: 'Hapus data contoh' })) return;
    rows('AHSP_HEADER').filter(h => norm(h.VERSI) === 'CONTOH').forEach(h => { ahspKomp(h.KODE_AHSP).forEach(k => save('AHSP_KOMPONEN', { ID: k.ID, DELETED: 'YA' }, { silent: true })); save('AHSP_HEADER', { ID: h.ID, DELETED: 'YA' }, { silent: true }); });
    hsdAll().filter(r => String(r.SUMBER || '').includes('CONTOH')).forEach(r => save(r._t, { ID: r.ID, DELETED: 'YA' }, { silent: true }));
    rows('AHSP_VERSI').filter(v => norm(v.VERSI) === 'CONTOH').forEach(v => save('AHSP_VERSI', { ID: v.ID, DELETED: 'YA' }, { silent: true }));
    scheduleSync(); toast('Data contoh dihapus', 'ok'); rerenderSoft();
  };
  $('#dExp', el).onclick = () => downloadBlob(new Blob([JSON.stringify({ app: 'RAB Rumah & Gedung', version: APP_VERSION, exported: nowIso(), tables: S.t }, null, 1)], { type: 'application/json' }), `BACKUP_RAB_${todayStr().replace(/-/g, '')}.json`);
  $('#dReset', el).onclick = async () => {
    if (pendingCount() && !await confirmBox(`Masih ada ${pendingCount()} perubahan belum tersinkron dan akan hilang. Lanjutkan?`, { danger: true, ok: 'Tetap reset' })) return;
    if (!await confirmBox('Hapus cache lokal lalu muat ulang data dari Google Sheet?', { danger: true, ok: 'Reset' })) return;
    await DB.clear('tables'); S.queue = {}; S.uploads = []; await DB.set('meta', 'queue', {}); await DB.set('meta', 'uploads', []);
    S.meta.lastSync = ''; await saveMeta(); location.reload();
  };
}

/* ---------- Laporan library ---------- */
function mLaporan(el) {
  const h = libraryHealth();
  const non = rows('AHSP_HEADER').filter(x => !isOfficial(x));
  const sec = (title, list, fn) => `<div class="card"><div class="row between"><h3>${title}</h3>${badge(String(list.length), list.length ? 'warning' : 'success')}</div>${list.length ? `<div class="list">${list.slice(0, 50).map(fn).join('')}</div>${list.length > 50 ? `<div class="small muted">+${list.length - 50} lainnya</div>` : ''}` : '<div class="small muted">Aman.</div>'}</div>`;
  const ah = x => `<div class="li"><div class="main"><div class="t1">${esc(x.KODE_AHSP)}</div><div class="t2">${esc(x.URAIAN)}</div></div></div>`;
  const hs = x => `<div class="li"><div class="main"><div class="t1">${esc(x.KODE_HSD)} · ${esc(x.NAMA)}</div><div class="t2">${esc(x._j)} · ${esc(x.WILAYAH)} ${esc(x.TAHUN)} · update ${fmtDate(x.TANGGAL_UPDATE)}</div></div></div>`;
  el.innerHTML = `<h3 style="margin:0 0 10px">Kesehatan library</h3><div class="grid md2">
    ${sec('AHSP tanpa sumber', h.tanpaSumber, ah)}${sec('AHSP status PERLU CEK', h.perluCek, ah)}${sec('HSD harga kosong', h.hsdKosong, hs)}${sec('HSD harga > 12 bulan', h.hsdLama, hs)}${sec('AHSP tanpa komponen', h.tanpaKomponen, ah)}</div>
    <div class="row between" style="margin:20px 0 10px"><h3 style="margin:0">Daftar AHSP non-resmi</h3><button class="btn sm" id="xNon">Excel</button></div>
    ${non.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Kode</th><th>Uraian</th><th>Jenis</th><th>Referensi asal</th><th>Dasar perhitungan</th><th>Dibuat</th></tr></thead><tbody>${non.map(x => `<tr><td>${esc(x.KODE_AHSP)}</td><td>${esc(x.URAIAN)}</td><td>${badge(x.JENIS || '-')}</td><td>${esc(x.KODE_REFERENSI || '-')}</td><td>${esc(x.DASAR_PERHITUNGAN || '-')}</td><td>${esc(x.DIBUAT_OLEH || '')} ${fmtDate(x.TANGGAL)}</td></tr>`).join('')}</tbody></table></div>` : '<div class="card empty"><p>Semua AHSP resmi.</p></div>'}
    <h3 style="margin:20px 0 10px">Item RAB harga pasar / LS</h3>
    ${(() => { const l = rows('RAB_ITEM').filter(r => norm(r.JENIS_HARGA) === 'PASAR' && norm(r.TIPE_BARIS) !== 'DIVISI'); return l.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Proyek</th><th>Uraian</th><th class="num">Harga</th><th>Sumber harga</th></tr></thead><tbody>${l.map(r => `<tr><td>${esc((getRow('PROYEK', r.PROYEK_ID) || {}).KODE || '')}</td><td>${esc(r.URAIAN)}</td><td class="num">${rpCell(r.HARGA_PASAR)}</td><td>${esc(r.SUMBER_HARGA || '-')}</td></tr>`).join('')}</tbody></table></div>` : '<div class="small muted">Tidak ada.</div>'; })()}`;
  $('#xNon', el).onclick = async () => { const X = await needXlsx(); const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['KODE', 'URAIAN', 'SATUAN', 'JENIS', 'REFERENSI', 'DASAR PERHITUNGAN', 'SUMBER', 'DIBUAT OLEH', 'TANGGAL']].concat(non.map(x => [x.KODE_AHSP, x.URAIAN, x.SATUAN, x.JENIS, x.KODE_REFERENSI, x.DASAR_PERHITUNGAN, x.SUMBER, x.DIBUAT_OLEH, x.TANGGAL]))), 'Non-resmi'); X.writeFile(wb, `AHSP_NON_RESMI_${todayStr().replace(/-/g, '')}.xlsx`); };
}
/* ============================================================
 * Penawaran harga (untuk klien) + PDF + Excel
 * ============================================================ */
route('/rab/:id/penawaran/:pn', (id, pn) => penawaranEditor(id, pn));
const PN_STATUS = ['DRAFT', 'TERKIRIM', 'NEGOSIASI', 'DITERIMA', 'DITOLAK', 'KEDALUWARSA'];
const pnItems = pnId => rows('PENAWARAN_ITEM').filter(r => r.PENAWARAN_ID === pnId).sort((a, b) => num(a.URUTAN) - num(b.URUTAN));
const pnLocked = pn => pnItems(pn.ID).length > 0;
function pnExpiry(pn) { const d = new Date(String(pn.TANGGAL || todayStr()).slice(0, 10)); d.setDate(d.getDate() + num(pn.MASA_BERLAKU_HARI || 14)); return d; }
function checkExpired(pn) {
  if (['TERKIRIM', 'NEGOSIASI'].includes(norm(pn.STATUS)) && pnExpiry(pn) < new Date(todayStr())) { save('PENAWARAN', { ID: pn.ID, STATUS: 'KEDALUWARSA' }); return Object.assign({}, pn, { STATUS: 'KEDALUWARSA' }); }
  return pn;
}
function jparse(s, d) { try { return s ? (typeof s === 'string' ? JSON.parse(s) : s) : d; } catch (e) { return d; } }
function pnOpt(pn) {
  return {
    markupPct: num(pn.MARKUP_PCT), markupDivisi: jparse(pn.MARKUP_DIVISI_JSON, {}), divisiSetting: jparse(pn.DIVISI_SETTING_JSON, {}),
    bulatSatuan: num(pn.BULAT_SATUAN), bulatTotal: num(pn.BULAT_TOTAL), diskonTipe: pn.DISKON_TIPE || 'NOMINAL', diskonNilai: num(pn.DISKON_NILAI),
    ppnMode: pn.PPN_MODE || 'TIDAK', ppnPct: num(pn.PPN_PCT)
  };
}

function tabPenawaran(p, rk, el) {
  const list = rows('PENAWARAN').filter(x => x.PROYEK_ID === p.ID).map(checkExpired).sort((a, b) => String(b.NOMOR).localeCompare(String(a.NOMOR)) || num(b.REV) - num(a.REV));
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><button class="btn primary" id="pnNew">${ICON.plus} Buat penawaran</button><span class="small muted">Penawaran terpisah dari RAB internal — AHSP, koefisien & margin tidak pernah tercetak.</span></div>
    ${list.length ? `<div class="card"><div class="list">${list.map(x => `<a class="li" href="#/rab/${p.ID}/penawaran/${x.ID}" style="text-decoration:none;color:inherit"><div class="main"><div class="t1">${esc(x.NOMOR)} · Rev-${num(x.REV)}</div><div class="t2">${fmtDate(x.TANGGAL)} · ${esc(x.KLIEN_NAMA || '-')} · ${esc(x.DETAIL || '')}${pnLocked(x) ? ' · tercetak' : ' · draf'}</div></div>${badge(x.STATUS || 'DRAFT')}<strong class="num">${rp(x.TOTAL)}</strong></a>`).join('')}</div></div>` : '<div class="card empty"><p>Belum ada penawaran untuk proyek ini.</p></div>'}
    ${num(p.NILAI_KONTRAK) ? `<div class="notice success" style="margin-top:12px">Nilai kontrak: <strong>${rp(p.NILAI_KONTRAK)}</strong></div>` : ''}`;
  $('#pnNew', el).onclick = () => {
    const prof = profilOf(p.PROFIL_ID);
    const pkp = prof && norm(prof.PKP) === 'YA';
    const specs = [...new Set(projItems(p.ID).map(i => i.SPESIFIKASI).filter(Boolean))].join('\n');
    const row = save('PENAWARAN', {
      ID: uid('PN'), PROYEK_ID: p.ID, NOMOR: penawaranNumber(p), REV: 0, TANGGAL: todayStr(), KLIEN_NAMA: p.PEMILIK || '', KLIEN_ALAMAT: p.LOKASI || '',
      PERIHAL: 'Penawaran Harga ' + (p.NAMA || 'Pekerjaan'), PROFIL_ID: prof ? prof.ID : '', DETAIL: 'STANDAR', MARKUP_PCT: 0, MARKUP_DIVISI_JSON: '{}', DIVISI_SETTING_JSON: '{}',
      BULAT_SATUAN: 100, BULAT_TOTAL: 1000, DISKON_TIPE: 'NOMINAL', DISKON_NILAI: 0, PPN_MODE: pkp ? 'BELUM' : 'TIDAK', PPN_PCT: setting('PPN_PCT'),
      MASA_BERLAKU_HARI: setting('MASA_BERLAKU_HARI'), WAKTU_PELAKSANAAN: '', TERMIN: setting('TERMIN_DEFAULT'), GARANSI: 'Masa pemeliharaan 3 bulan setelah serah terima',
      LINGKUP: syaratText('LINGKUP'), EXCLUSIONS: syaratText('EXCLUSION'), SYARAT: syaratText('SYARAT'), SPESIFIKASI: specs || syaratText('SPESIFIKASI'),
      TAMPIL_REKENING: 'YA', TAMPIL_HARGA_M2: 'TIDAK', TAMPIL_CREDIT: setting('TAMPIL_CREDIT'), STATUS: 'DRAFT', TOTAL_RAB_INTERNAL: rk.C
    });
    location.hash = `#/rab/${p.ID}/penawaran/${row.ID}`;
  };
}

/* Bangun data penawaran: dari snapshot (bila sudah dicetak) atau live dari RAB */
function pnBuild(p, pn) {
  if (pnLocked(pn)) {
    const its = pnItems(pn.ID), groups = [];
    let g = null;
    its.forEach(r => {
      if (norm(r.TIPE_BARIS) === 'DIVISI') { g = { kode: r.KODE, uraian: r.URAIAN, items: [], subtotal: num(r.JUMLAH) }; groups.push(g); }
      else if (g) g.items.push({ uraian: r.URAIAN, satuan: r.SATUAN, volume: num(r.VOLUME), hs: num(r.HARGA_SATUAN), jumlah: num(r.JUMLAH), spesifikasi: r.SPESIFIKASI });
    });
    const ppnMode = pn.PPN_MODE || 'TIDAK', ppnPct = num(pn.PPN_PCT);
    const dpp = num(pn.SUBTOTAL) - num(pn.DISKON);
    return { locked: true, groups, subtotal: num(pn.SUBTOTAL), diskon: num(pn.DISKON), ppn: num(pn.PPN), ppnMode, ppnPct, total: num(pn.TOTAL), ppnInfo: ppnMode === 'TERMASUK' ? Math.round(dpp * ppnPct / (100 + ppnPct)) : 0, internal: num(pn.TOTAL_RAB_INTERNAL), margin: (ppnMode === 'BELUM' ? num(pn.TOTAL) - num(pn.PPN) : num(pn.TOTAL) - (ppnMode === 'TERMASUK' ? Math.round(dpp * ppnPct / (100 + ppnPct)) : 0)) - num(pn.TOTAL_RAB_INTERNAL) };
  }
  const rk = rekapOf(p);
  const r = Calc.penawaran(rk, pnOpt(pn));
  r.locked = false;
  r.groups.forEach(g => g.kode = g.kode || '');
  return r;
}

function penawaranEditor(pid, pnId) {
  const p = getRow('PROYEK', pid); let pn = getRow('PENAWARAN', pnId);
  if (!p || !pn) { view('<div class="card empty"><p>Penawaran tidak ditemukan.</p></div>'); return; }
  pn = checkExpired(pn);
  setTitle('Penawaran ' + pn.NOMOR, p.NAMA);
  const locked = pnLocked(pn);
  const rk = rekapOf(p);
  const divs = rk.groups.filter(g => g.row.ID !== '_');
  const ds = jparse(pn.DIVISI_SETTING_JSON, {}), md = jparse(pn.MARKUP_DIVISI_JSON, {});
  const profs = rows('PROFIL_USAHA');
  const form = locked ? '' : `<form id="pnf" class="stack">
    <div class="card"><h3>Data surat</h3><div class="form-grid c2">
      ${inp('NOMOR', 'Nomor', pn.NOMOR)}${inp('TANGGAL', 'Tanggal', pn.TANGGAL, { type: 'date' })}
      ${inp('KLIEN_NAMA', 'Kepada (nama klien)', pn.KLIEN_NAMA)}${inp('KLIEN_ALAMAT', 'Alamat klien', pn.KLIEN_ALAMAT)}
      ${inp('PERIHAL', 'Perihal', pn.PERIHAL, { cls: 'span2' })}
      ${sel('PROFIL_ID', 'Profil perusahaan (kop)', pn.PROFIL_ID, [['', '— pilih —']].concat(profs.map(x => [x.ID, x.NAMA])))}
      ${sel('DETAIL', 'Tingkat detail', pn.DETAIL, [['RINGKAS', 'Ringkas — rekap per divisi'], ['STANDAR', 'Standar — per item'], ['LUMPSUM', 'Lump sum — total + lingkup']])}</div></div>
    <div class="card"><h3>Harga</h3><div class="form-grid c3">
      ${inp('MARKUP_PCT', 'Markup penawaran (%)', pn.MARKUP_PCT, { numType: 'num', hint: 'di luar overhead AHSP' })}
      ${sel('BULAT_SATUAN', 'Pembulatan harga satuan', pn.BULAT_SATUAN, [['1', 'Tanpa'], ['100', 'Ratusan'], ['1000', 'Ribuan']])}
      ${sel('BULAT_TOTAL', 'Pembulatan total', pn.BULAT_TOTAL, [['1', 'Tanpa'], ['1000', 'Ribuan'], ['10000', 'Puluh ribuan'], ['100000', 'Ratus ribuan']])}
      ${sel('DISKON_TIPE', 'Jenis diskon', pn.DISKON_TIPE, [['NOMINAL', 'Nominal (Rp)'], ['PERSEN', 'Persen (%)']])}
      ${inp('DISKON_NILAI', 'Nilai diskon', pn.DISKON_NILAI, { numType: 'num' })}
      ${sel('PPN_MODE', 'PPN', pn.PPN_MODE, [['BELUM', 'Belum termasuk (ditambahkan)'], ['TERMASUK', 'Sudah termasuk'], ['TIDAK', 'Tidak dikenakan']])}
      ${inp('PPN_PCT', 'Tarif PPN (%)', pn.PPN_PCT, { numType: 'num' })}</div>
      <h3 style="margin-top:16px;font-size:14px">Pengaturan per divisi</h3>
      <div class="table-wrap"><table class="t"><thead><tr><th>Divisi</th><th>Tampil</th><th>Gabung ke</th><th class="num">Markup khusus %</th></tr></thead><tbody>${divs.map(g => {
        const s = ds[g.row.ID] || {};
        return `<tr data-div="${g.row.ID}"><td>${esc(g.row.KODE || '')} ${esc(g.row.URAIAN)}<div class="small muted">${rp(g.subtotal)}</div></td>
          <td><select class="input" data-f="mode" style="width:130px"><option value="TAMPIL" ${!s.mode || s.mode === 'TAMPIL' ? 'selected' : ''}>Tampil</option><option value="GABUNG" ${s.mode === 'GABUNG' ? 'selected' : ''}>Gabungkan</option><option value="SEMBUNYI" ${s.mode === 'SEMBUNYI' ? 'selected' : ''}>Keluarkan</option></select></td>
          <td><select class="input" data-f="ke" style="min-width:150px"><option value="">—</option>${divs.filter(x => x.row.ID !== g.row.ID).map(x => `<option value="${x.row.ID}" ${s.ke === x.row.ID ? 'selected' : ''}>${esc(x.row.KODE || '')} ${esc(x.row.URAIAN.slice(0, 28))}</option>`).join('')}</select></td>
          <td><input class="input num" data-f="mk" inputmode="decimal" value="${md[g.row.ID] !== undefined && md[g.row.ID] !== '' ? esc(String(md[g.row.ID]).replace('.', ',')) : ''}" placeholder="global" style="width:90px"></td></tr>`;
      }).join('')}</tbody></table></div>
      <div class="small muted" style="margin-top:6px">"Keluarkan" = tidak termasuk penawaran. "Gabungkan" = nilainya masuk ke divisi tujuan (mis. SMKK ke Persiapan).</div></div>
    <div class="card"><h3>Syarat & isi dokumen</h3><div class="form-grid c2">
      ${inp('MASA_BERLAKU_HARI', 'Masa berlaku (hari)', pn.MASA_BERLAKU_HARI, { numType: 'num' })}${inp('WAKTU_PELAKSANAAN', 'Estimasi waktu pelaksanaan', pn.WAKTU_PELAKSANAAN, { ph: 'mis. 120 hari kalender' })}
      ${inp('TERMIN', 'Skema termin pembayaran', pn.TERMIN, { cls: 'span2' })}${inp('GARANSI', 'Masa pemeliharaan / garansi', pn.GARANSI, { cls: 'span2' })}
      ${txt('SPESIFIKASI', 'Spesifikasi material utama', pn.SPESIFIKASI, { cls: 'span2', rows: 4 })}
      ${txt('LINGKUP', 'Lingkup pekerjaan (termasuk)', pn.LINGKUP, { rows: 4 })}${txt('EXCLUSIONS', 'Tidak termasuk', pn.EXCLUSIONS, { rows: 4 })}
      ${txt('SYARAT', 'Syarat & ketentuan', pn.SYARAT, { cls: 'span2', rows: 4 })}
      <label class="check"><input type="checkbox" name="TAMPIL_REKENING" ${norm(pn.TAMPIL_REKENING) === 'YA' ? 'checked' : ''}> Tampilkan rekening pembayaran</label>
      <label class="check"><input type="checkbox" name="TAMPIL_HARGA_M2" ${norm(pn.TAMPIL_HARGA_M2) === 'YA' ? 'checked' : ''}> Tampilkan harga per m²</label>
      <label class="check"><input type="checkbox" name="TAMPIL_CREDIT" ${norm(pn.TAMPIL_CREDIT) === 'YA' ? 'checked' : ''}> Tulisan kecil "dibuat dengan RAB Rumah & Gedung"</label></div></div></form>`;
  view(`<div class="page-head"><div><a href="#/rab/${pid}/penawaran" class="small muted" style="text-decoration:none">${ICON.back.replace('class="ico"', 'class="ico" style="width:14px;height:14px;vertical-align:-2px"')} ${esc(p.NAMA)}</a><h2 style="margin-top:4px">${esc(pn.NOMOR)} · Rev-${num(pn.REV)}</h2></div>
      <div class="actions">${badge(pn.STATUS || 'DRAFT')}<select class="input" id="pnSt" style="width:auto">${PN_STATUS.map(s => `<option ${norm(pn.STATUS) === s ? 'selected' : ''}>${s}</option>`).join('')}</select></div></div>
    ${locked ? `<div class="notice" style="margin-bottom:12px">Penawaran ini sudah dicetak — harganya terkunci (snapshot ${fmtDate(pn.TANGGAL)}) dan tidak ikut berubah saat AHSP/HSD diperbarui. Untuk mengubah, buat revisi.</div>` : ''}
    <div class="grid" style="grid-template-columns:minmax(0,1fr)">${form}<div id="pnSum"></div></div>`);
  $('#pnSt').onchange = e => { save('PENAWARAN', { ID: pn.ID, STATUS: e.target.value }); if (e.target.value === 'DITERIMA') toast('Status Diterima — gunakan "Jadikan nilai kontrak" di bawah', ''); rerenderSoft(); };
  const collect = () => {
    if (locked) return pn;
    const d = formData($('#pnf'));
    const nds = {}, nmd = {};
    $$('tr[data-div]').forEach(tr => {
      const mode = $('[data-f=mode]', tr).value, ke = $('[data-f=ke]', tr).value, mk = parseNum($('[data-f=mk]', tr).value);
      if (mode !== 'TAMPIL') nds[tr.dataset.div] = { mode, ke };
      if (mk !== '') nmd[tr.dataset.div] = mk;
    });
    return Object.assign({}, pn, d, { DIVISI_SETTING_JSON: JSON.stringify(nds), MARKUP_DIVISI_JSON: JSON.stringify(nmd) });
  };
  const drawSum = () => {
    const cur = collect();
    const r = pnBuild(p, cur);
    const bad = Object.entries(jparse(cur.DIVISI_SETTING_JSON, {})).filter(([, v]) => v.mode === 'GABUNG' && !v.ke).length;
    $('#pnSumBody').innerHTML = `<div class="table-wrap"><table class="t"><thead><tr><th>Uraian</th><th class="num">Jumlah</th></tr></thead><tbody>${r.groups.map(g => `<tr><td>${esc(g.kode || '')} ${esc(g.uraian)} <span class="small muted">(${g.items.length} item)</span></td><td class="num">${rpCell(g.subtotal)}</td></tr>`).join('')}
        <tr class="sub"><td>Subtotal</td><td class="num">${rpCell(r.subtotal)}</td></tr>
        ${r.diskon ? `<tr><td>Diskon</td><td class="num">${rpCell(-r.diskon)}</td></tr>` : ''}
        ${r.ppnMode === 'BELUM' ? `<tr><td>PPN ${pct(r.ppnPct, 1)}</td><td class="num">${rpCell(r.ppn)}</td></tr>` : ''}
        <tr class="tot"><td>TOTAL PENAWARAN</td><td class="num">${rpCell(r.total)}</td></tr></tbody></table></div>
      <div class="small" style="margin-top:6px"><em>${esc(Calc.terbilang(r.total))}</em>${r.ppnMode === 'TERMASUK' ? ` · sudah termasuk PPN ${rp(r.ppnInfo)}` : ''}</div>
      ${bad ? `<div class="notice danger" style="margin-top:8px">${bad} divisi diset "Gabungkan" tanpa divisi tujuan.</div>` : ''}
      <div class="notice warning" style="margin-top:12px"><div><strong>Margin internal (tidak tercetak):</strong> ${rp(r.margin)} · ${pct(r.internal ? r.margin / r.internal * 100 : 0, 1)} terhadap RAB internal ${rp(r.internal)} (sebelum biaya tak terduga & PPN)</div></div>`;
    return { cur, r, bad };
  };
  $('#pnSum').innerHTML = `<div class="card"><h3>Ringkasan penawaran</h3><div id="pnSumBody"></div>
      <div class="row" style="margin-top:12px">
        ${locked ? `<button class="btn primary" id="pnPdf">${ICON.pdf} Unduh PDF</button><button class="btn" id="pnRev">Buat revisi</button>` : `<button class="btn" id="pnSave">Simpan draf</button><button class="btn primary" id="pnPdf">${ICON.pdf} Cetak & simpan ke Drive</button>`}
        <button class="btn ghost" id="pnPrev">Preview</button>
        ${!locked ? `<button class="btn danger" id="pnDel">Hapus draf</button>` : ''}
        ${norm(pn.STATUS) === 'DITERIMA' ? `<button class="btn accent" id="pnKon">Jadikan nilai kontrak</button>` : ''}</div>
      ${pn.FILE_URL ? `<div class="small muted" style="margin-top:8px">PDF di Drive: <a href="${esc(pn.FILE_URL)}" target="_blank" rel="noopener">buka</a></div>` : ''}</div>`;
  $('#pnSave') && ($('#pnSave').onclick = () => { const { cur, r } = drawSum(); save('PENAWARAN', Object.assign(cur, { SUBTOTAL: r.subtotal, DISKON: r.diskon, PPN: r.ppn, TOTAL: r.total, TOTAL_RAB_INTERNAL: r.internal })); toast('Draf disimpan', 'ok'); });
  $('#pnPrev').onclick = () => pnPreview(p, collect());
  $('#pnPdf').onclick = async () => {
    let { cur: cur2, bad } = drawSum();
    if (bad) return toast('Perbaiki pengaturan "Gabungkan" dulu', 'err');
    if (!locked) {
      if (!await confirmBox('Cetak penawaran?\nHarga akan dikunci (snapshot). Perubahan berikutnya harus lewat revisi.', { ok: 'Cetak' })) return;
      cur2 = pnSnapshot(p, cur2);
    }
    await makePenawaranPdf(p, cur2, { upload: !locked });
    if (!locked) rerenderSoft();
  };
  $('#pnRev') && ($('#pnRev').onclick = () => {
    const nr = save('PENAWARAN', Object.assign({}, pn, { ID: uid('PN'), REV: num(pn.REV) + 1, TANGGAL: todayStr(), STATUS: 'DRAFT', FILE_ID: '', FILE_URL: '', UPDATED_AT: '' }));
    toast('Revisi Rev-' + nr.REV + ' dibuat (draf, harga mengikuti RAB terbaru)', 'ok');
    location.hash = `#/rab/${pid}/penawaran/${nr.ID}`;
  });
  $('#pnDel') && ($('#pnDel').onclick = async () => { if (!await confirmBox('Hapus draf penawaran ini?', { danger: true, ok: 'Hapus' })) return; remove('PENAWARAN', pn.ID); location.hash = `#/rab/${pid}/penawaran`; });
  $('#pnKon') && ($('#pnKon').onclick = async () => {
    const nilai = drawSum().r.total;
    if (!await confirmBox(`Jadikan ${rp(nilai)} sebagai nilai kontrak proyek?\nNilai ini menjadi acuan opname, termin & CCO (Tahap 4).`, { ok: 'Jadikan nilai kontrak' })) return;
    save('PROYEK', { ID: p.ID, NILAI_KONTRAK: nilai, STATUS: 'BERJALAN' }); toast('Nilai kontrak disimpan', 'ok'); rerenderSoft();
  });
  if (!locked) { bindMoneyInputs($('#pnf')); $('#pnf').addEventListener('input', debounce(drawSum, 300)); $('#pnf').addEventListener('change', drawSum); }
  drawSum();
}
/* Kunci harga: tulis PENAWARAN_ITEM + total */
function pnSnapshot(p, pn) {
  const r = Calc.penawaran(rekapOf(p), pnOpt(pn));
  pnItems(pn.ID).forEach(x => save('PENAWARAN_ITEM', { ID: x.ID, DELETED: 'YA' }, { silent: true }));
  let n = 0;
  r.groups.forEach(g => {
    save('PENAWARAN_ITEM', { ID: uid('PI'), PENAWARAN_ID: pn.ID, URUTAN: ++n, TIPE_BARIS: 'DIVISI', KODE: g.kode || '', URAIAN: g.uraian, JUMLAH: g.subtotal }, { silent: true });
    g.items.forEach(it => save('PENAWARAN_ITEM', { ID: uid('PI'), PENAWARAN_ID: pn.ID, URUTAN: ++n, TIPE_BARIS: 'ITEM', URAIAN: it.uraian, SATUAN: it.satuan, VOLUME: it.volume, HARGA_SATUAN: it.hs, JUMLAH: it.jumlah, SPESIFIKASI: it.spesifikasi }, { silent: true }));
  });
  const row = save('PENAWARAN', Object.assign({}, pn, { SUBTOTAL: r.subtotal, DISKON: r.diskon, PPN: r.ppn, TOTAL: r.total, TOTAL_RAB_INTERNAL: r.internal }));
  if (norm(p.STATUS) === 'DRAFT') save('PROYEK', { ID: p.ID, STATUS: 'PENAWARAN' }, { silent: true });
  return row;
}
function pnPreview(p, pn) {
  const r = pnBuild(p, pn), prof = profilOf(pn.PROFIL_ID) || {};
  const det = pn.DETAIL || 'STANDAR';
  openModal({
    title: 'Preview penawaran (tampilan klien)', wide: true,
    body: `<div class="preview-doc"><div style="border-bottom:2px solid ${THEME.accent};padding-bottom:8px;margin-bottom:12px"><strong style="font-size:15px">${esc(prof.NAMA || '(profil belum dipilih)')}</strong><div class="small muted">${esc([prof.ALAMAT, prof.KOTA].filter(Boolean).join(', '))} · ${esc([prof.TELEPON, prof.EMAIL].filter(Boolean).join(' · '))}</div></div>
      <div>Nomor: ${esc(pn.NOMOR)} · ${fmtDateLong(pn.TANGGAL)}<br>Perihal: ${esc(pn.PERIHAL)}<br>Kepada: ${esc(pn.KLIEN_NAMA)}</div>
      ${det === 'LUMPSUM' ? `<p><strong>Total lump sum: ${rp(r.total)}</strong></p>` : `<table class="t" style="margin:10px 0"><thead><tr><th>Uraian</th>${det === 'STANDAR' ? '<th class="num">Vol</th><th>Sat</th><th class="num">Harga satuan</th>' : ''}<th class="num">Jumlah</th></tr></thead><tbody>${r.groups.map(g => `<tr class="div"><td colspan="${det === 'STANDAR' ? 4 : 1}">${esc(g.uraian)}</td><td class="num">${rp(g.subtotal)}</td></tr>${det === 'STANDAR' ? g.items.map(i => `<tr><td>${esc(i.uraian)}</td><td class="num">${Nmax(i.volume, 2)}</td><td>${esc(i.satuan)}</td><td class="num">${rp(i.hs)}</td><td class="num">${rp(i.jumlah)}</td></tr>`).join('') : ''}`).join('')}<tr class="tot"><td colspan="${det === 'STANDAR' ? 4 : 1}">TOTAL</td><td class="num">${rp(r.total)}</td></tr></tbody></table>`}
      <div class="small muted">Preview ringkas. Dokumen PDF lengkap memuat spesifikasi, lingkup, syarat, dan kolom tanda tangan.</div></div>`
  });
}

/* ============================================================
 * PDF
 * ============================================================ */
const FORBIDDEN = /\b(AHSP|koefisien|HSD|overhead|biaya umum|keuntungan|margin|MODIFIKASI|CUSTOM|PERLU CEK|harga dasar)\b/i;
function hexRgb(h) { const n = parseInt(h.slice(1), 16); return [n >> 16 & 255, n >> 8 & 255, n & 255]; }
async function pdfStart(prof) {
  const jsPDF = await needPdf();
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  doc._texts = [];
  const origText = doc.text.bind(doc);
  const fix = x => String(x).replace(/²/g, '2').replace(/³/g, '3').replace(/[“”]/g, '"').replace(/[‘’]/g, "'");
  doc.text = function (t, ...a) { t = Array.isArray(t) ? t.map(fix) : fix(t); doc._texts.push(Array.isArray(t) ? t.join(' ') : t); return origText(t, ...a); };
  const logo = prof ? await imageData(prof.LOGO_FILE_ID) : '';
  return { doc, logo, prof: prof || {} };
}
function pdfKop(ctx, y = 14, W = 210) {
  const { doc, logo, prof } = ctx, M = 15;
  let x = M;
  if (logo) {
    try { const ip = doc.getImageProperties(logo); const h = 20, w = Math.min(40, ip.width / ip.height * h); doc.addImage(logo, M, y, w, h); x = M + w + 5; } catch (e) { }
  }
  doc.setTextColor(...hexRgb(THEME.primaryDark)); doc.setFont('helvetica', 'bold'); doc.setFontSize(14);
  doc.text(prof.NAMA || '', x, y + 6);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(...hexRgb(THEME.muted));
  const lines = [prof.TAGLINE, [prof.ALAMAT, prof.KOTA, prof.PROVINSI, prof.KODE_POS].filter(Boolean).join(', '),
    [prof.TELEPON && 'Telp ' + prof.TELEPON, prof.WHATSAPP && 'WA ' + prof.WHATSAPP, prof.EMAIL, prof.WEBSITE, prof.INSTAGRAM].filter(Boolean).join('  ·  '),
    [prof.NPWP && 'NPWP ' + prof.NPWP, prof.NIB && 'NIB ' + prof.NIB].filter(Boolean).join('  ·  ')].filter(Boolean);
  let ly = y + 11;
  lines.forEach(l => { doc.text(doc.splitTextToSize(l, W - x - M), x, ly); ly += 4; });
  const bottom = Math.max(y + 22, ly);
  doc.setDrawColor(...hexRgb(prof.WARNA_AKSEN || THEME.accent)); doc.setLineWidth(0.8); doc.line(M, bottom, W - M, bottom);
  doc.setTextColor(...hexRgb(THEME.text));
  return bottom + 6;
}
function pdfFooter(ctx, docNo, credit) {
  const { doc, prof } = ctx, n = doc.getNumberOfPages();
  for (let i = 1; i <= n; i++) {
    doc.setPage(i); doc.setFontSize(7.5); doc.setTextColor(...hexRgb(THEME.muted));
    doc.text([prof.NAMA, docNo].filter(Boolean).join('  ·  '), 15, 289);
    doc.text(`Halaman ${i} dari ${n}`, 195, 289, { align: 'right' });
    if (credit) doc.text('dibuat dengan RAB Rumah & Gedung by DG', 105, 293, { align: 'center' });
  }
}
const AT = (extra = {}) => Object.assign({
  theme: 'grid', styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 1.8, lineColor: hexRgb(THEME.border), lineWidth: 0.1, textColor: hexRgb(THEME.text) },
  headStyles: { fillColor: hexRgb(THEME.primary), textColor: 255, fontStyle: 'bold' }, alternateRowStyles: { fillColor: hexRgb(THEME.zebra) },
  margin: { left: 15, right: 15, top: 16, bottom: 18 }
}, extra);
const rpS = n => { const v = Math.round(num(n)); return v === 0 ? '-' : (v < 0 ? '-' : '') + N(Math.abs(v)); };
function rowStyle(kind) {
  if (kind === 'DIV') return { fillColor: hexRgb(THEME.primary50), fontStyle: 'bold' };
  if (kind === 'SUB') return { fillColor: hexRgb(THEME.accent50), fontStyle: 'bold' };
  if (kind === 'TOT') return { fillColor: hexRgb(THEME.primaryDark), textColor: 255, fontStyle: 'bold' };
  return {};
}
function styled(row, kind) { return row.map(c => ({ content: c, styles: rowStyle(kind) })); }
function signBlock(doc, y, cols) {
  if (y > 240) { doc.addPage(); y = 20; }
  const W = 180 / cols.length;
  doc.setFontSize(9);
  cols.forEach((c, i) => {
    const x = 15 + W * i + W / 2;
    doc.text(c.title, x, y, { align: 'center' });
    if (c.img) { try { doc.addImage(c.img, x - 18, y + 3, 36, 18); } catch (e) { } }
    if (c.stamp) { try { doc.addImage(c.stamp, x - 6, y + 2, 24, 20); } catch (e) { } }
    doc.setFont('helvetica', 'bold'); doc.text(c.name || '(..............................)', x, y + 28, { align: 'center' });
    doc.setFont('helvetica', 'normal'); if (c.role) doc.text(c.role, x, y + 32, { align: 'center' });
  });
  return y + 36;
}
async function deliverPdf(doc, name, { proyekId, kategori, jenis, terkait, onDone, shareText }) {
  const blob = doc.output('blob');
  if (proyekId) await queueUpload({ blob, name, mime: 'application/pdf', kategori, proyekId, jenis, terkait, onDone });
  const file = new File([blob], name, { type: 'application/pdf' });
  openModal({
    title: 'PDF siap',
    body: `<p style="margin-top:0"><strong>${esc(name)}</strong></p><p class="small muted">${proyekId ? (navigator.onLine && S.meta.apiUrl ? 'Sedang disimpan ke Google Drive (folder proyek).' : 'Akan tersimpan ke Google Drive otomatis saat online.') : ''}</p>`,
    foot: `<button class="btn" id="dlPdf">Unduh</button>${navigator.canShare && navigator.canShare({ files: [file] }) ? '<button class="btn" id="shPdf">Bagikan…</button>' : ''}<button class="btn primary" id="waPdf">${ICON.share} WhatsApp</button>`,
    onOpen: (el) => {
      $('#dlPdf', el).onclick = () => downloadBlob(blob, name);
      $('#shPdf', el) && ($('#shPdf', el).onclick = () => navigator.share({ files: [file], title: name, text: shareText }).catch(() => { }));
      $('#waPdf', el).onclick = async () => {
        if (navigator.canShare && navigator.canShare({ files: [file] })) return navigator.share({ files: [file], text: shareText }).catch(() => { });
        downloadBlob(blob, name); shareWa(shareText + '\n\n(PDF terlampir terpisah)');
      };
    }
  });
}

/* ---------- PDF RAB internal ---------- */
async function exportRabPdf(p) {
  openModal({
    title: 'Cetak PDF RAB', body: `<div class="list">
      <label class="check li"><input type="checkbox" id="oR" checked> Rincian RAB per item</label>
      <label class="check li"><input type="checkbox" id="oA" checked> Analisa harga satuan (AHSP terpakai)</label>
      <label class="check li"><input type="checkbox" id="oH"> Daftar harga dasar terpakai</label></div>
      <div class="small muted" style="margin-top:8px">Dokumen internal: memuat AHSP & koefisien. Untuk klien gunakan tab Penawaran.</div>`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="oOk">${ICON.pdf} Buat PDF</button>`,
    onOpen: (el, m) => { $('#oOk', el).onclick = async () => { const o = { rinci: $('#oR', el).checked, ahsp: $('#oA', el).checked, hsd: $('#oH', el).checked }; m.close(); try { await buildRabPdf(p, o); } catch (e) { console.error(e); toast('Gagal membuat PDF: ' + e.message, 'err'); } }; }
  });
}
async function buildRabPdf(p, o) {
  const rk = recompute(p.ID), prof = profilOf(p.PROFIL_ID);
  const ctx = await pdfStart(prof), doc = ctx.doc;
  let y = pdfKop(ctx);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.text('RENCANA ANGGARAN BIAYA (RAB)', 105, y + 2, { align: 'center' }); y += 8;
  doc.autoTable(AT({
    startY: y, theme: 'plain', styles: { fontSize: 9, cellPadding: 1 }, columnStyles: { 0: { cellWidth: 38, textColor: hexRgb(THEME.muted) } },
    body: [['Pekerjaan', p.NAMA], ['Pemilik', p.PEMILIK || '-'], ['Lokasi', p.LOKASI || '-'], ['Jenis / luas', `${p.JENIS_BANGUNAN || '-'} · ${N(p.LUAS_BANGUNAN, 2)} m²`], ['Harga dasar', `${rk.ctx.wilayah} ${rk.ctx.tahun}`], ['Tanggal', fmtDateLong(todayStr())]]
  }));
  y = doc.lastAutoTable.finalY + 4;
  const rekBody = rk.groups.map(g => [g.row.KODE || '', g.row.URAIAN, rpS(g.subtotal), pct(g.bobot)]);
  const tail = [['', 'A. Jumlah biaya pekerjaan', rpS(rk.A), '', 'SUB'], ['', 'B. Biaya penerapan SMKK', rpS(rk.B), '', 'SUB'], ['', 'C. Total (A+B)', rpS(rk.C), '']];
  if (rk.kontPct) tail.push(['', `D. Biaya tak terduga ${pct(rk.kontPct, 1)}`, rpS(rk.D), '']);
  if (rk.ppnPct) tail.push(['', `PPN ${pct(rk.ppnPct, 1)}`, rpS(rk.ppn), '']);
  if (rk.selisihBulat) tail.push(['', 'Pembulatan', rpS(rk.selisihBulat), '']);
  tail.push(['', 'TOTAL RAB', rpS(rk.total), '', 'TOT']);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.text('REKAPITULASI', 15, y + 3);
  doc.autoTable(AT({ startY: y + 5, head: [['No', 'Uraian pekerjaan', 'Jumlah (Rp)', 'Bobot']], body: rekBody.concat(tail.map(t => styled(t.slice(0, 4), t[4]))), columnStyles: { 0: { cellWidth: 14 }, 2: { halign: 'right', cellWidth: 38 }, 3: { halign: 'right', cellWidth: 20 } } }));
  y = doc.lastAutoTable.finalY + 4;
  doc.setFont('helvetica', 'italic'); doc.setFontSize(8.5); doc.text(doc.splitTextToSize('Terbilang: ' + Calc.terbilang(rk.total), 180), 15, y + 2);
  if (rk.perM2) { doc.setFont('helvetica', 'normal'); doc.text(`Harga per m²: Rp ${N(rk.perM2)} (${N(rk.luas, 2)} m²)`, 15, y + 8); }
  if (o.rinci) {
    doc.addPage(); doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.text('RINCIAN RENCANA ANGGARAN BIAYA', 105, 18, { align: 'center' });
    const body = [];
    rk.groups.filter(g => g.items.some(it => num(it.VOLUME) > 0)).forEach(g => {
      body.push(styled([g.row.KODE || '', g.row.URAIAN, '', '', '', ''], 'DIV'));
      g.items.filter(it => num(it.VOLUME) > 0).forEach(it => { const c = rk.computed[it.ID]; body.push([it.KODE || '', it.URAIAN + (norm(it.JENIS_HARGA) === 'PASAR' ? ' *' : ''), Nmax(it.VOLUME, 3), it.SATUAN || '', rpS(c.hs), rpS(c.jumlah)]); });
      body.push(styled(['', 'Subtotal ' + g.row.URAIAN, '', '', '', rpS(g.subtotal)], 'SUB'));
    });
    body.push(styled(['', 'TOTAL BIAYA PEKERJAAN + SMKK', '', '', '', rpS(rk.C)], 'TOT'));
    const zeroN = rk.itemCount - rk.groups.reduce((a, g) => a + g.items.filter(it => num(it.VOLUME) > 0).length, 0);
    doc.autoTable(AT({ startY: 22, head: [['No', 'Uraian pekerjaan', 'Volume', 'Sat', 'Harga satuan', 'Jumlah (Rp)']], body, columnStyles: { 0: { cellWidth: 16 }, 2: { halign: 'right', cellWidth: 18 }, 3: { cellWidth: 12 }, 4: { halign: 'right', cellWidth: 28 }, 5: { halign: 'right', cellWidth: 32 } } }));
    doc.setFontSize(7.5); doc.setFont('helvetica', 'normal');
    doc.text(['* harga pasar / lump sum (lihat sumber harga di aplikasi)', zeroN ? `${zeroN} item dengan volume 0 tidak dicetak.` : ''].filter(Boolean).join('   ·   '), 15, doc.lastAutoTable.finalY + 4);
  }
  if (o.ahsp) {
    const used = [...new Set(Object.values(rk.computed).filter(c => c.analisa && c.analisa.ok).map(c => c.analisa.kode))];
    if (used.length) {
      doc.addPage(); doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.text('ANALISA HARGA SATUAN PEKERJAAN', 105, 18, { align: 'center' });
      let yy = 24;
      used.forEach(k => {
        const a = Calc.analyze(idx(), k, rk.ctx), h = a.header;
        const body = [];
        [['UPAH', 'A. Tenaga kerja', a.upah], ['BAHAN', 'B. Bahan', a.bahan], ['ALAT', 'C. Peralatan', a.alat]].forEach(([j, l, s]) => {
          const rr = a.rows.filter(r => r.jenis === j); if (!rr.length) return;
          body.push(styled(['', l, '', '', '', rpS(s)], 'DIV'));
          rr.forEach(r => body.push([r.kodeHsd, r.nama, r.satuan, Nmax(r.koef, 4), rpS(r.hargaDasar), rpS(r.jumlah)]));
        });
        body.push(styled(['', 'D. Jumlah (A+B+C)', '', '', '', rpS(a.E)], 'SUB'));
        body.push(['', `E. Biaya umum & keuntungan ${pct(rk.ctx.overhead, 1)}`, '', '', '', rpS(a.F)]);
        body.push(styled(['', 'F. Harga satuan pekerjaan', '', '', '', rpS(a.HSP)], 'TOT'));
        doc.autoTable(AT({ startY: yy, head: [[{ content: `${k} — ${h.URAIAN} (per ${h.SATUAN})`, colSpan: 6, styles: { halign: 'left' } }], ['Kode', 'Komponen', 'Sat', 'Koef', 'Harga', 'Jumlah']], body, pageBreak: 'avoid', columnStyles: { 0: { cellWidth: 20 }, 2: { cellWidth: 12 }, 3: { halign: 'right', cellWidth: 18 }, 4: { halign: 'right', cellWidth: 26 }, 5: { halign: 'right', cellWidth: 28 } } }));
        doc.setFontSize(7); doc.setTextColor(...hexRgb(THEME.muted)); doc.text(doc.splitTextToSize(`Sumber: ${h.SUMBER || '-'} · ${h.JENIS || ''} · ${h.STATUS || ''}`, 180), 15, doc.lastAutoTable.finalY + 3); doc.setTextColor(...hexRgb(THEME.text));
        yy = doc.lastAutoTable.finalY + 8;
      });
    }
  }
  if (o.hsd) {
    const list = Calc.takeoff(idx(), rk);
    doc.addPage(); doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.text('DAFTAR HARGA DASAR & KEBUTUHAN', 105, 18, { align: 'center' });
    doc.autoTable(AT({ startY: 22, head: [['Jenis', 'Kode', 'Nama', 'Sat', 'Harga dasar', 'Kebutuhan']], body: list.map(x => [x.jenis, x.kode, x.nama, x.satuan, rpS(x.harga), Nmax(x.qty, 3)]), columnStyles: { 4: { halign: 'right' }, 5: { halign: 'right' } } }));
  }
  let sy = (doc.lastAutoTable ? doc.lastAutoTable.finalY : 200) + 12;
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9); if (sy > 240) { doc.addPage(); sy = 24; }
  doc.text(`${(prof && prof.KOTA) || ''}${prof && prof.KOTA ? ', ' : ''}${fmtDateLong(todayStr())}`, 195, sy, { align: 'right' });
  const ttd = prof && norm(prof.TAMPILKAN_TTD) === 'YA' ? await imageData(prof.TTD_FILE_ID) : '';
  signBlock(doc, sy + 6, [{ title: 'Dibuat oleh', name: prof && prof.PJ_NAMA, role: prof && prof.PJ_JABATAN, img: ttd }, { title: 'Diperiksa oleh' }, { title: 'Disetujui oleh', name: p.PEMILIK }]);
  pdfFooter(ctx, `RAB ${p.KODE}`, norm(setting('TAMPIL_CREDIT')) === 'YA');
  const name = `${p.KODE}_RAB_${todayStr().replace(/-/g, '')}.pdf`;
  await deliverPdf(doc, name, { proyekId: p.ID, kategori: '01_RAB', jenis: 'RAB', terkait: p.ID, shareText: `RAB ${p.NAMA} — total ${rp(rk.total)}` });
}

/* ---------- PDF Penawaran (klien) ---------- */
async function makePenawaranPdf(p, pn, { upload }) {
  try {
    const prof = profilOf(pn.PROFIL_ID);
    if (!prof) return toast('Pilih profil perusahaan dulu (Master → Profil Perusahaan)', 'err');
    const r = pnBuild(p, pn);
    const ctx = await pdfStart(prof), doc = ctx.doc, det = pn.DETAIL || 'STANDAR';
    let y = pdfKop(ctx);
    doc.setFontSize(9.5); doc.setFont('helvetica', 'normal');
    doc.autoTable(AT({ startY: y, theme: 'plain', styles: { fontSize: 9.5, cellPadding: 0.8 }, columnStyles: { 0: { cellWidth: 22 } }, body: [['Nomor', ': ' + pn.NOMOR + (num(pn.REV) ? ` (Rev-${num(pn.REV)})` : '')], ['Perihal', ': ' + (pn.PERIHAL || '')], ['Lampiran', ': Rincian penawaran']] }));
    doc.text(`${prof.KOTA ? prof.KOTA + ', ' : ''}${fmtDateLong(pn.TANGGAL)}`, 195, y + 4, { align: 'right' });
    y = doc.lastAutoTable.finalY + 5;
    doc.text('Kepada Yth.', 15, y); doc.setFont('helvetica', 'bold'); doc.text(pn.KLIEN_NAMA || '-', 15, y + 5); doc.setFont('helvetica', 'normal');
    if (pn.KLIEN_ALAMAT) doc.text(doc.splitTextToSize(pn.KLIEN_ALAMAT, 120), 15, y + 10);
    y += pn.KLIEN_ALAMAT ? 20 : 15;
    const lokasi = p.LOKASI ? ` di ${p.LOKASI}` : '';
    const opening = `Dengan hormat,\nBersama ini kami sampaikan penawaran harga untuk pekerjaan ${p.NAMA}${lokasi}` + (num(p.LUAS_BANGUNAN) ? ` dengan luas bangunan ${N(p.LUAS_BANGUNAN, 2)} m²` : '') + (num(p.JUMLAH_LANTAI) ? ` (${N(p.JUMLAH_LANTAI)} lantai)` : '') + ', dengan rincian sebagai berikut:';
    doc.text(doc.splitTextToSize(opening, 180), 15, y); y += 16;
    if (det === 'LUMPSUM') {
      doc.autoTable(AT({ startY: y, head: [['Uraian', 'Jumlah (Rp)']], body: [[`Pekerjaan ${p.NAMA} (lump sum) sesuai lingkup pekerjaan terlampir`, rpS(r.subtotal)]], columnStyles: { 1: { halign: 'right', cellWidth: 40 } } }));
    } else if (det === 'RINGKAS') {
      doc.autoTable(AT({ startY: y, head: [['No', 'Uraian pekerjaan', 'Jumlah (Rp)']], body: r.groups.map((g, i) => [ROMAWI[i] || i + 1, g.uraian, rpS(g.subtotal)]), columnStyles: { 0: { cellWidth: 14 }, 2: { halign: 'right', cellWidth: 40 } } }));
    } else {
      const body = [];
      r.groups.forEach((g, gi) => {
        body.push(styled([ROMAWI[gi] || gi + 1, g.uraian, '', '', '', ''], 'DIV'));
        g.items.forEach((it, ii) => body.push([String(ii + 1), it.uraian, Nmax(it.volume, 2), it.satuan, rpS(it.hs), rpS(it.jumlah)]));
        body.push(styled(['', 'Subtotal', '', '', '', rpS(g.subtotal)], 'SUB'));
      });
      doc.autoTable(AT({ startY: y, head: [['No', 'Uraian pekerjaan', 'Volume', 'Sat', 'Harga satuan', 'Jumlah (Rp)']], body, columnStyles: { 0: { cellWidth: 12 }, 2: { halign: 'right', cellWidth: 18 }, 3: { cellWidth: 12 }, 4: { halign: 'right', cellWidth: 28 }, 5: { halign: 'right', cellWidth: 32 } } }));
    }
    const tot = [styled(['Subtotal', rpS(r.subtotal)], 'SUB')];
    if (r.diskon) tot.push(['Diskon', '-' + rpS(r.diskon)]);
    if (r.ppnMode === 'BELUM') tot.push([`PPN ${pct(r.ppnPct, 0)}`, rpS(r.ppn)]);
    if (r.total !== r.subtotal - r.diskon + r.ppn) tot.push(['Pembulatan', rpS(r.total - (r.subtotal - r.diskon + r.ppn))]);
    tot.push(styled(['TOTAL PENAWARAN', rpS(r.total)], 'TOT'));
    doc.autoTable(AT({ startY: doc.lastAutoTable.finalY + 2, body: tot, margin: { left: 95, right: 15 }, columnStyles: { 1: { halign: 'right', cellWidth: 40 } } }));
    y = doc.lastAutoTable.finalY + 5;
    doc.setFont('helvetica', 'italic'); doc.setFontSize(9);
    const tbl = 'Terbilang: ' + Calc.terbilang(r.total) + (r.ppnMode === 'TERMASUK' ? ' (sudah termasuk PPN)' : r.ppnMode === 'TIDAK' ? '' : ' (termasuk PPN)');
    doc.text(doc.splitTextToSize(tbl, 180), 15, y); y += 8;
    doc.setFont('helvetica', 'normal');
    if (norm(pn.TAMPIL_HARGA_M2) === 'YA' && num(p.LUAS_BANGUNAN)) { doc.text(`Harga per m² bangunan: Rp ${N(Math.round(r.total / num(p.LUAS_BANGUNAN)))}`, 15, y); y += 6; }
    const sect = (title, text) => {
      if (!String(text || '').trim()) return;
      const lines = String(text).split('\n').map(s => s.trim()).filter(Boolean);
      const wrapped = lines.map(l => doc.splitTextToSize('•  ' + l, 176));
      const h = 6 + wrapped.reduce((a, b) => a + b.length * 4.2, 0);
      if (y + h > 275) { doc.addPage(); y = 20; }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.text(title, 15, y); y += 5;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9);
      wrapped.forEach(w => { doc.text(w, 17, y); y += w.length * 4.2; });
      y += 2;
    };
    sect('Spesifikasi material utama', pn.SPESIFIKASI);
    sect('Lingkup pekerjaan', pn.LINGKUP);
    sect('Tidak termasuk dalam penawaran', pn.EXCLUSIONS);
    const sk = [`Penawaran berlaku ${num(pn.MASA_BERLAKU_HARI)} hari sejak tanggal surat (s.d. ${fmtDateLong(pnExpiry(pn).toISOString().slice(0, 10))}).`, pn.WAKTU_PELAKSANAAN ? 'Estimasi waktu pelaksanaan: ' + pn.WAKTU_PELAKSANAAN + '.' : '', pn.TERMIN ? 'Termin pembayaran: ' + pn.TERMIN + '.' : '', pn.GARANSI ? pn.GARANSI + '.' : ''].filter(Boolean).join('\n') + '\n' + (pn.SYARAT || '');
    sect('Syarat & ketentuan', sk);
    if (norm(pn.TAMPIL_REKENING) === 'YA' && prof.NO_REKENING) sect('Pembayaran ke rekening', `${prof.BANK || ''} ${prof.NO_REKENING} a.n. ${prof.ATAS_NAMA || prof.NAMA}`);
    if (y > 235) { doc.addPage(); y = 20; }
    doc.setFontSize(9); doc.text(doc.splitTextToSize('Demikian penawaran ini kami sampaikan. Atas perhatian dan kerja samanya kami ucapkan terima kasih.', 180), 15, y + 2);
    const showTtd = norm(prof.TAMPILKAN_TTD) === 'YA';
    const ttd = showTtd ? await imageData(prof.TTD_FILE_ID) : '', stamp = showTtd ? await imageData(prof.STEMPEL_FILE_ID) : '';
    signBlock(doc, y + 12, [{ title: 'Hormat kami,', name: prof.PJ_NAMA || prof.NAMA, role: prof.PJ_JABATAN || prof.NAMA, img: ttd, stamp }, { title: 'Disetujui oleh,', name: pn.KLIEN_NAMA, role: 'Tanggal: ..................' }]);
    pdfFooter(ctx, pn.NOMOR, norm(pn.TAMPIL_CREDIT) === 'YA');
    // cek kebocoran data internal
    let scan = doc._texts.join('\n');
    ['NAMA', 'TAGLINE', 'ALAMAT', 'KOTA', 'EMAIL', 'WEBSITE', 'INSTAGRAM', 'PJ_NAMA', 'PJ_JABATAN', 'ATAS_NAMA'].forEach(f => { if (prof[f]) scan = scan.split(String(prof[f])).join(' '); });
    const leak = scan.match(FORBIDDEN);
    if (leak && !await confirmBox(`Dokumen klien memuat kata "${leak[0]}" yang biasanya hanya untuk internal (mis. di uraian, spesifikasi, atau syarat).\nTetap cetak?`, { ok: 'Tetap cetak' })) return;
    const name = `${p.KODE}_PENAWARAN_${String(pn.NOMOR).replace(/[\\/:*?"<>|\s]+/g, '-')}_Rev${num(pn.REV)}_${String(pn.TANGGAL || todayStr()).replace(/-/g, '')}.pdf`;
    const shareText = `Yth. ${pn.KLIEN_NAMA || ''},\nBerikut penawaran harga ${p.NAMA} (${pn.NOMOR}) senilai ${rp(r.total)}, berlaku s.d. ${fmtDateLong(pnExpiry(pn).toISOString().slice(0, 10))}.\n\nTerima kasih.\n${prof.NAMA}`;
    await deliverPdf(doc, name, upload ? { proyekId: p.ID, kategori: '02_Penawaran', jenis: 'PENAWARAN', terkait: pn.ID, onDone: { type: 'setField', table: 'PENAWARAN', id: pn.ID, field: 'FILE_ID', urlField: 'FILE_URL' }, shareText } : { shareText });
  } catch (e) { console.error(e); toast('Gagal membuat PDF: ' + e.message, 'err'); }
}

/* ---------- Excel RAB ---------- */
async function exportRabXlsx(p) {
  const X = await needXlsx(), rk = recompute(p.ID);
  const wb = X.utils.book_new();
  const rek = [['REKAPITULASI RAB — ' + p.NAMA], [], ['No', 'Uraian', 'Jumlah', 'Bobot %']].concat(rk.groups.map(g => [g.row.KODE || '', g.row.URAIAN, Math.round(g.subtotal), Calc.round(g.bobot, 2)]))
    .concat([[], ['', 'A. Jumlah biaya pekerjaan', rk.A], ['', 'B. SMKK', rk.B], ['', 'C. Total', rk.C], ['', 'D. Biaya tak terduga', rk.D], ['', 'PPN', rk.ppn], ['', 'TOTAL RAB', rk.total], ['', 'Terbilang', Calc.terbilang(rk.total)], ['', 'Harga per m²', rk.perM2]]);
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(rek), 'Rekap');
  const rab = [['No', 'Uraian', 'Volume', 'Satuan', 'Harga satuan', 'Jumlah', 'Kode AHSP', 'Jenis harga', 'Bobot %']];
  rk.groups.forEach(g => { rab.push([g.row.KODE || '', g.row.URAIAN, '', '', '', Math.round(g.subtotal)]); g.items.forEach(it => { const c = rk.computed[it.ID]; rab.push([it.KODE, it.URAIAN, num(it.VOLUME), it.SATUAN, c.hs, c.jumlah, it.KODE_AHSP || '', it.JENIS_HARGA || 'AHSP', Calc.round(c.bobot, 3)]); }); });
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(rab), 'RAB');
  const an = [['Kode AHSP', 'Uraian', 'Jenis', 'Kode HSD', 'Komponen', 'Sat', 'Koefisien', 'Harga dasar', 'Jumlah']];
  [...new Set(Object.values(rk.computed).filter(c => c.analisa && c.analisa.ok).map(c => c.analisa.kode))].forEach(k => { const a = Calc.analyze(idx(), k, rk.ctx); a.rows.forEach(r => an.push([k, a.header.URAIAN, r.jenis, r.kodeHsd, r.nama, r.satuan, r.koef, r.hargaDasar, Math.round(r.jumlah)])); an.push([k, 'HSP (E+F)', '', '', '', '', '', '', a.HSP]); });
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(an), 'Analisa');
  X.writeFile(wb, `${p.KODE}_RAB_${todayStr().replace(/-/g, '')}.xlsx`);
}
/* ============================================================
 * Tahap 2 — Tenaga kerja: OH, durasi ↔ jumlah tukang, rekap,
 * histogram, perbandingan borongan, termin upah
 * ============================================================ */
const ROLE_LABEL = { MANDOR: 'Mandor', KEPALA: 'Kepala tukang', TUKANG: 'Tukang', PEKERJA: 'Pekerja', LAIN: 'Lainnya' };
const TWARN = { DURASI_KOSONG: ['Durasi belum diisi', 'warning'], RASIO_KEPALA: ['Kepala tukang kurang', 'warning'], RASIO_MANDOR: ['Mandor kurang', 'warning'] };
const tenagaOpt = () => ({ jamKerja: num(setting('JAM_KERJA_HARI')) || 8, hariMinggu: num(setting('HARI_KERJA_MINGGU')) || 6, rasioKepala: num(setting('RASIO_KEPALA_TUKANG')) || 10, rasioMandor: num(setting('RASIO_MANDOR_PEKERJA')) || 20 });
const jadwalRows = pid => rows('JADWAL').filter(r => r.PROYEK_ID === pid);
function tenagaPlan(pid) {
  const plan = {};
  jadwalRows(pid).forEach(r => { plan[r.ITEM_ID] = { durasi: num(r.DURASI), tukang: num(r.JUMLAH_TUKANG), rowId: r.ID }; });
  return plan;
}
function tenagaOf(p, rk) { return Calc.tenaga(idx(), rk || rekapOf(p), tenagaPlan(p.ID), tenagaOpt()); }
const borRows = pid => rows('BORONGAN_UPAH').filter(r => r.PROYEK_ID === pid);
const terminRows = pid => rows('TERMIN_UPAH').filter(r => r.PROYEK_ID === pid).sort((a, b) => num(a.URUTAN) - num(b.URUTAN));
function borongan(p, rk) {
  const B = borRows(p.ID), get = m => B.find(r => norm(r.MODE) === m) || {};
  const m2 = get('M2'), ht = get('HARIAN_TUKANG'), hp = get('HARIAN_PEKERJA');
  const totM2 = num(m2.HARGA) * num(p.LUAS_BANGUNAN);
  let totItem = 0, nItem = 0, ahspItem = 0;
  B.filter(r => norm(r.MODE) === 'ITEM' && num(r.HARGA) > 0).forEach(r => {
    const it = getRow('RAB_ITEM', r.ITEM_ID); if (!it || Calc.isDel(it)) return;
    totItem += num(r.HARGA) * num(it.VOLUME); nItem++;
    const c = rk.computed[it.ID]; if (c) ahspItem += c.upah;
  });
  const totHarian = num(ht.JUMLAH_TUKANG) * num(ht.HARI) * num(ht.HARGA) + num(hp.JUMLAH_TUKANG) * num(hp.HARI) * num(hp.HARGA);
  return { m2, ht, hp, totM2, totItem, nItem, ahspItem, totHarian };
}

/* Simpan rekap tenaga ke sheet TENAGA_REKAP (agar terlihat di Google Sheet) */
function syncTenagaRekap(p, t) {
  const want = {};
  t.roles.forEach(r => { want[p.ID + '|' + norm(r.kode)] = { ID: p.ID + '|' + norm(r.kode), PROYEK_ID: p.ID, JENIS_TENAGA: r.nama, TOTAL_OH: Calc.round(r.oh, 3), UPAH_HARIAN: r.harga, TOTAL_UPAH: Math.round(r.upah), ORANG_PUNCAK: r.puncak }; });
  let changed = false;
  rows('TENAGA_REKAP').filter(r => r.PROYEK_ID === p.ID).forEach(r => { if (!want[r.ID]) { save('TENAGA_REKAP', { ID: r.ID, DELETED: 'YA' }, { silent: true }); changed = true; } });
  Object.values(want).forEach(w => { const cur = getRow('TENAGA_REKAP', w.ID); if (!cur || diff(cur, w)) { save('TENAGA_REKAP', w, { silent: true }); changed = true; } });
  if (changed) scheduleSync();
}

function tabTenaga(p, rk, el) {
  const t = tenagaOf(p, rk);
  syncTenagaRekap(p, t);
  const o = tenagaOpt();
  const roleCol = t.roles;
  const bor = borongan(p, rk);
  const ahspUpah = t.totalUpah;
  const maxW = Math.max(1, ...t.weeks.map(w => w.orang));
  const divWithOh = t.divisi.filter(d => d.ohTotal);
  const anyDur = divWithOh.some(d => d.durasi);
  el.innerHTML = `
    ${!roleCol.length ? `<div class="notice warning" style="margin-bottom:12px">${ICON.warn}<div>Belum ada kebutuhan tenaga. Tenaga dihitung dari koefisien upah AHSP × volume — pastikan item sudah terhubung AHSP dan volumenya terisi.</div></div>` : ''}
    <div class="grid k4">
      <div class="card kpi primary"><div class="label">Total upah (AHSP)</div><div class="value">${rp(ahspUpah)}</div><div class="hint">sebelum biaya umum & keuntungan</div></div>
      <div class="card kpi"><div class="label">Total kebutuhan</div><div class="value">${N(t.totalOH, 1)} OH</div><div class="hint">orang-hari (1 OH = 8 jam)</div></div>
      <div class="card kpi"><div class="label">Durasi total</div><div class="value">${t.totalDurasi ? N(t.totalDurasi) + ' hari' : '-'}</div><div class="hint">${t.totalDurasi ? `± ${N(t.minggu)} minggu · ${o.hariMinggu} hari/minggu` : 'isi durasi per divisi'}</div></div>
      <div class="card kpi"><div class="label">Orang puncak</div><div class="value">${anyDur ? N(Math.max(0, ...t.weeks.map(w => w.orang))) + ' orang' : '-'}</div><div class="hint">jam kerja ${o.jamKerja} jam/hari</div></div>
    </div>

    <div class="card" style="margin-top:12px"><div class="row between"><h3>Durasi & jumlah tukang per divisi</h3><span class="small muted">Isi <strong>durasi</strong> (hari kerja) atau <strong>tukang tersedia</strong> — yang lain dihitung otomatis</span></div>
      ${divWithOh.length ? `<div class="table-wrap"><table class="t" id="tnTbl"><thead><tr><th>Divisi</th><th class="num">OH</th><th class="num">Durasi (hari)</th><th class="num">Tukang tersedia</th>${roleCol.map(r => `<th class="num" title="${esc(r.kode)}">${esc(r.nama)}</th>`).join('')}<th>Status</th></tr></thead><tbody>
        ${divWithOh.map(d => `<tr data-div="${d.id}"><td>${esc(d.kode)} ${esc(d.uraian)}<div class="small muted">${d.items.length} item</div></td><td class="num">${N(d.ohTotal, 1)}</td>
          <td class="num"><input class="input num" data-f="durasi" inputmode="numeric" value="${d.mode === 'TUKANG' ? '' : (d.durasi || '')}" placeholder="${d.mode === 'TUKANG' ? N(d.durasi) : 'hari'}" style="width:80px" ${d.mode === 'TUKANG' ? 'title="Dihitung dari jumlah tukang"' : ''}></td>
          <td class="num"><input class="input num" data-f="tukang" inputmode="numeric" value="${num((tenagaPlan(p.ID)[d.id] || {}).tukang) || ''}" placeholder="${d.ohTukang ? 'orang' : '–'}" ${d.ohTukang ? '' : 'disabled'} style="width:80px"></td>
          ${roleCol.map(r => { const k = norm(r.kode); return `<td class="num">${d.oh[k] ? `<strong>${d.durasi ? N(d.orang[k]) : '-'}</strong><div class="small muted">${N(d.oh[k], 1)} OH</div>` : '<span class="muted">·</span>'}</td>`; }).join('')}
          <td>${d.warn.map(w => badge(TWARN[w][0], TWARN[w][1])).join(' ') || (d.durasi ? badge('OK', 'success') : '')}${d.mode === 'TUKANG' ? `<div class="small muted">durasi ${N(d.durasi)} hari dari jumlah tukang</div>` : ''}</td></tr>`).join('')}
        <tr class="tot"><td>Total</td><td class="num">${N(t.totalOH, 1)}</td><td class="num">${N(t.totalDurasi)}</td><td></td>${roleCol.map(r => `<td class="num">${N(r.oh, 1)} OH</td>`).join('')}<td></td></tr></tbody></table></div>
        <div class="small muted" style="margin-top:6px">Jumlah orang = OH ÷ durasi, dibulatkan ke atas${o.jamKerja !== 8 ? ` (dikoreksi jam kerja ${o.jamKerja} jam/hari)` : ''}. Peringatan rasio: 1 kepala tukang ≤ ${o.rasioKepala} tukang, 1 mandor ≤ ${o.rasioMandor} pekerja+tukang. Ubah di Master → Pengaturan.</div>` : '<div class="muted">Belum ada divisi dengan kebutuhan tenaga.</div>'}</div>

    <div class="grid md2" style="margin-top:12px">
      <div class="card"><h3>Rekap per jenis tenaga</h3>${roleCol.length ? `<div class="table-wrap"><table class="t"><thead><tr><th>Jenis tenaga</th><th class="num">Total OH</th><th class="num">Upah/hari</th><th class="num">Total upah</th><th class="num">Puncak</th></tr></thead><tbody>
        ${roleCol.map(r => `<tr><td>${esc(r.nama)}<div class="small muted">${esc(r.kode)} · ${ROLE_LABEL[r.kelas]}</div></td><td class="num">${N(r.oh, 2)}</td><td class="num">${rpCell(r.harga)}</td><td class="num">${rpCell(r.upah)}</td><td class="num">${anyDur ? N(r.puncak) + ' org' : '-'}</td></tr>`).join('')}
        <tr class="tot"><td>Total</td><td class="num">${N(t.totalOH, 2)}</td><td></td><td class="num">${rpCell(ahspUpah)}</td><td></td></tr></tbody></table></div>` : '<div class="muted small">-</div>'}
        <div class="row" style="margin-top:10px"><button class="btn sm" id="tnX">Excel</button><button class="btn sm" id="tnP">${ICON.pdf} PDF rekap tenaga</button></div></div>
      <div class="card"><h3>Histogram tenaga per minggu</h3>${t.weeks.length ? `<div style="display:flex;align-items:flex-end;gap:4px;height:160px;padding-top:8px;overflow-x:auto">${t.weeks.map(w => `<div style="flex:1;min-width:22px;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%" title="Minggu ${w.minggu}: ${w.orang} orang · ${N(w.oh, 1)} OH · ${esc(Object.keys(w.divisi).join(', '))}"><span class="small num" style="font-size:10px">${w.orang}</span><div style="width:100%;background:var(--primary);border-radius:4px 4px 0 0;height:${Math.max(2, w.orang / maxW * 120)}px"></div><span class="small muted" style="font-size:10px">M${w.minggu}</span></div>`).join('')}</div>
        <div class="small muted" style="margin-top:6px">Perkiraan: divisi dikerjakan berurutan sesuai urutan RAB. Di Tahap 3 histogram mengikuti jadwal (pekerjaan paralel & ketergantungan).</div>` : '<div class="muted small">Isi durasi per divisi untuk melihat histogram.</div>'}</div>
    </div>

    <div class="card" style="margin-top:12px"><h3>Perbandingan upah: AHSP vs borongan vs harian</h3>
      <div class="form-grid c3" id="borF">
        ${inp('M2_HARGA', 'Borongan upah per m² bangunan', bor.m2.HARGA, { numType: 'rp', hint: `× ${N(p.LUAS_BANGUNAN, 2)} m² = ${rp(bor.totM2)}` })}
        <div class="field"><label>Harian — tukang</label><div class="row" style="flex-wrap:nowrap"><input class="input num" name="HT_N" inputmode="numeric" value="${esc(bor.ht.JUMLAH_TUKANG || '')}" placeholder="org" style="width:70px"><input class="input num" name="HT_HARI" inputmode="numeric" value="${esc(bor.ht.HARI || '')}" placeholder="hari" style="width:70px"><input class="input num" name="HT_UPAH" data-num="rp" inputmode="decimal" value="${num(bor.ht.HARGA) ? N(bor.ht.HARGA) : ''}" placeholder="upah/hari"></div></div>
        <div class="field"><label>Harian — pekerja</label><div class="row" style="flex-wrap:nowrap"><input class="input num" name="HP_N" inputmode="numeric" value="${esc(bor.hp.JUMLAH_TUKANG || '')}" placeholder="org" style="width:70px"><input class="input num" name="HP_HARI" inputmode="numeric" value="${esc(bor.hp.HARI || '')}" placeholder="hari" style="width:70px"><input class="input num" name="HP_UPAH" data-num="rp" inputmode="decimal" value="${num(bor.hp.HARGA) ? N(bor.hp.HARGA) : ''}" placeholder="upah/hari"></div></div>
      </div>
      <div class="row" style="margin-top:8px"><button class="btn sm" id="borItem">Borongan per item (${bor.nItem})</button>${t.totalDurasi ? `<button class="btn sm ghost" id="borFill">Isi harian dari hasil hitung</button>` : ''}<span class="spacer"></span><button class="btn sm primary" id="borSave">Simpan & bandingkan</button></div>
      <div class="table-wrap" style="margin-top:12px"><table class="t"><thead><tr><th>Metode</th><th class="num">Total upah</th><th class="num">Selisih vs AHSP</th></tr></thead><tbody>
        <tr><td><strong>AHSP</strong> (koefisien × upah dasar)</td><td class="num">${rpCell(ahspUpah)}</td><td class="num">-</td></tr>
        ${[['Borongan per m²', bor.totM2, ahspUpah], ['Borongan per item' + (bor.nItem ? ` (${bor.nItem} item)` : ''), bor.totItem, bor.ahspItem], ['Harian', bor.totHarian, ahspUpah]].map(([l, v, base]) => `<tr><td>${l}${l.startsWith('Borongan per item') && bor.nItem ? `<div class="small muted">dibanding upah AHSP item yang sama ${rp(base)}</div>` : ''}</td><td class="num">${v ? rpCell(v) : '<span class="muted">-</span>'}</td><td class="num">${v && base ? `<span class="${v - base < 0 ? 'neg' : ''}">${rp(v - base)} (${pct((v - base) / base * 100, 1)})</span>` : '-'}</td></tr>`).join('')}
      </tbody></table></div>
      <div class="small muted" style="margin-top:6px">Koefisien tenaga AHSP biasanya lebih besar dari praktik borongan rumah tinggal. Untuk proyek pemerintah tetap gunakan AHSP.</div></div>

    <div class="card" style="margin-top:12px"><div class="row between"><h3>Termin pembayaran upah</h3><div class="row"><button class="btn sm" id="tmGen">Buat termin</button></div></div>
      ${terminTable(p)}</div>`;

  // simpan durasi/tukang per divisi
  const tbl = $('#tnTbl', el);
  tbl && tbl.addEventListener('change', e => {
    const tr = e.target.closest('tr[data-div]'); if (!tr) return;
    const did = tr.dataset.div, plan = tenagaPlan(p.ID)[did] || {};
    const f = e.target.dataset.f, v = parseNum(e.target.value);
    const patch = { ID: plan.rowId || uid('J'), PROYEK_ID: p.ID, ITEM_ID: did };
    if (f === 'durasi') { patch.DURASI = v === '' ? '' : Math.max(0, Math.round(v)); patch.JUMLAH_TUKANG = ''; }
    else { patch.JUMLAH_TUKANG = v === '' ? '' : Math.max(0, Math.round(v)); if (v !== '') patch.DURASI = ''; }
    save('JADWAL', patch);
    setTimeout(() => rerenderSoft(), 0);
  });
  bindMoneyInputs(el);
  $('#borSave', el).onclick = () => {
    const d = formData($('#borF', el));
    const g = m => borRows(p.ID).find(r => norm(r.MODE) === m);
    const put = (m, patch) => save('BORONGAN_UPAH', Object.assign({ ID: (g(m) || {}).ID || uid('B'), PROYEK_ID: p.ID, MODE: m }, patch), { silent: true });
    put('M2', { SATUAN: 'm2', HARGA: d.M2_HARGA, TOTAL: Math.round(num(d.M2_HARGA) * num(p.LUAS_BANGUNAN)) });
    const n = x => parseNum(x);
    put('HARIAN_TUKANG', { JUMLAH_TUKANG: n(d.HT_N), HARI: n(d.HT_HARI), HARGA: parseNum(d.HT_UPAH, true), TOTAL: Math.round(num(n(d.HT_N)) * num(n(d.HT_HARI)) * num(parseNum(d.HT_UPAH, true))) });
    put('HARIAN_PEKERJA', { JUMLAH_TUKANG: n(d.HP_N), HARI: n(d.HP_HARI), HARGA: parseNum(d.HP_UPAH, true), TOTAL: Math.round(num(n(d.HP_N)) * num(n(d.HP_HARI)) * num(parseNum(d.HP_UPAH, true))) });
    scheduleSync(); toast('Perbandingan disimpan', 'ok'); rerenderSoft();
  };
  $('#borFill', el) && ($('#borFill', el).onclick = () => {
    const tk = t.roles.filter(r => r.kelas === 'TUKANG' || r.kelas === 'KEPALA'), pk = t.roles.filter(r => r.kelas === 'PEKERJA' || r.kelas === 'MANDOR' || r.kelas === 'LAIN');
    const avg = (list, f) => { const oh = list.reduce((a, r) => a + r.oh, 0); return oh ? list.reduce((a, r) => a + f(r), 0) / oh : 0; };
    const days = t.totalDurasi;
    const nT = days ? Math.ceil(tk.reduce((a, r) => a + r.oh, 0) * t.faktorJam / days) : 0, nP = days ? Math.ceil(pk.reduce((a, r) => a + r.oh, 0) * t.faktorJam / days) : 0;
    $('[name=HT_N]', el).value = nT; $('[name=HT_HARI]', el).value = days; $('[name=HT_UPAH]', el).value = N(Math.round(avg(tk, r => r.upah)));
    $('[name=HP_N]', el).value = nP; $('[name=HP_HARI]', el).value = days; $('[name=HP_UPAH]', el).value = N(Math.round(avg(pk, r => r.upah)));
    toast('Terisi dari hasil hitung (rata-rata orang & upah). Sesuaikan dengan harga lapangan.', '');
  });
  $('#borItem', el).onclick = () => boronganItemModal(p, rk);
  $('#tmGen', el).onclick = () => terminModal(p, rk, t, bor);
  $('#tnX', el).onclick = () => exportTenagaXlsx(p, rk, t, bor);
  $('#tnP', el).onclick = () => exportTenagaPdf(p, rk, t, bor).catch(e => toast('Gagal membuat PDF: ' + e.message, 'err'));
  $$('[data-tmst]', el).forEach(b => b.onchange = () => { save('TERMIN_UPAH', { ID: b.dataset.tmst, STATUS: b.value }); });
}
function terminTable(p) {
  const L = terminRows(p.ID);
  if (!L.length) return '<div class="muted small">Belum ada termin. Klik "Buat termin" untuk menyusun jadwal pembayaran upah (mingguan atau per termin).</div>';
  let cum = 0;
  return `<div class="table-wrap"><table class="t"><thead><tr><th>#</th><th>Termin</th><th class="num">%</th><th class="num">Nominal</th><th class="num">Kumulatif</th><th>Status</th></tr></thead><tbody>${L.map(r => { cum += num(r.NOMINAL); return `<tr><td>${num(r.URUTAN)}</td><td>${esc(r.LABEL)}${r.MINGGU_KE ? `<div class="small muted">minggu ke-${num(r.MINGGU_KE)}</div>` : ''}</td><td class="num">${pct(r.PERSEN, 1)}</td><td class="num">${rpCell(r.NOMINAL)}</td><td class="num">${rpCell(cum)}</td><td><select class="input" data-tmst="${r.ID}" style="width:130px"><option ${norm(r.STATUS) !== 'DISETUJUI' ? 'selected' : ''}>DRAFT</option><option ${norm(r.STATUS) === 'DISETUJUI' ? 'selected' : ''}>DISETUJUI</option></select></td></tr>`; }).join('')}
    <tr class="tot"><td colspan="3">Total (${esc(L[0].BASIS)})</td><td class="num">${rpCell(cum)}</td><td colspan="2"></td></tr></tbody></table></div>
    <div class="small muted" style="margin-top:6px">Status pembayaran aktual (dibayar/belum) dicatat di modul realisasi Tahap 4.</div>`;
}
function boronganItemModal(p, rk) {
  const items = []; rk.groups.forEach(g => g.items.forEach(it => { if (num(it.VOLUME) > 0) items.push({ it, g, c: rk.computed[it.ID] }); }));
  const B = borRows(p.ID).filter(r => norm(r.MODE) === 'ITEM');
  openModal({
    title: 'Borongan upah per item', wide: true,
    body: `<div class="small muted" style="margin-bottom:8px">Isi harga borongan upah per satuan (tanpa bahan) untuk item yang diborongkan. Kosongkan yang tidak diborongkan.</div>
      <div class="table-wrap"><table class="t"><thead><tr><th>Item</th><th class="num">Volume</th><th class="num">Upah AHSP/sat</th><th class="num">Borongan/sat</th></tr></thead><tbody>${items.map(({ it, c }) => {
        const r = B.find(x => x.ITEM_ID === it.ID) || {};
        const v = num(it.VOLUME);
        return `<tr data-it="${it.ID}" data-bid="${r.ID || ''}"><td>${esc(it.KODE || '')} ${esc(it.URAIAN)}</td><td class="num">${Nmax(v, 2)} ${esc(it.SATUAN)}</td><td class="num">${c && v ? rpCell(c.upah / v) : '-'}</td><td class="num"><input class="input num" data-num="rp" inputmode="decimal" value="${num(r.HARGA) ? N(r.HARGA) : ''}" style="width:120px"></td></tr>`;
      }).join('')}</tbody></table></div>`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="biOk">Simpan</button>`,
    onOpen: (el, m) => {
      bindMoneyInputs(el);
      $('#biOk', el).onclick = () => {
        $$('tr[data-it]', el).forEach(tr => {
          const h = parseNum($('input', tr).value, true), bid = tr.dataset.bid, it = getRow('RAB_ITEM', tr.dataset.it);
          if (h === '' || !(num(h) > 0)) { if (bid) save('BORONGAN_UPAH', { ID: bid, DELETED: 'YA' }, { silent: true }); return; }
          save('BORONGAN_UPAH', { ID: bid || uid('B'), PROYEK_ID: p.ID, MODE: 'ITEM', ITEM_ID: it.ID, SATUAN: it.SATUAN, HARGA: h, TOTAL: Math.round(h * num(it.VOLUME)), CATATAN: it.URAIAN }, { silent: true });
        });
        scheduleSync(); m.close(); rerenderSoft();
      };
    }
  });
}
function terminModal(p, rk, t, bor) {
  const bases = [['AHSP', 'Upah AHSP', t.totalUpah], ['M2', 'Borongan per m²', bor.totM2], ['ITEM', 'Borongan per item', bor.totItem], ['HARIAN', 'Harian', bor.totHarian]].filter(b => b[2] > 0);
  if (!bases.length) return toast('Belum ada total upah. Lengkapi AHSP/volume atau isi borongan dulu.', 'err');
  openModal({
    title: 'Buat termin pembayaran upah',
    body: `<form class="form-grid" id="tmF">${sel('BASIS', 'Dasar nilai upah', bases[0][0], bases.map(b => [b[0], `${b[1]} — ${rp(b[2])}`]))}
      ${sel('MODE', 'Pola termin', t.weeks.length ? 'MINGGUAN' : 'MANUAL', [['MINGGUAN', 'Mingguan — mengikuti bobot OH per minggu' + (t.weeks.length ? '' : ' (isi durasi dulu)')], ['MANUAL', 'Termin manual (%)']])}
      ${txt('MANUAL', 'Termin manual: satu baris = "Label; persen"', 'Termin 1 — pondasi & struktur; 30\nTermin 2 — dinding & atap; 30\nTermin 3 — finishing; 30\nTermin 4 — serah terima; 10', { rows: 5, hint: 'Total persen harus 100' })}
      ${inp('BULAT', 'Pembulatan per termin', 1000, { numType: 'num', hint: 'selisih pembulatan masuk ke termin terakhir' })}</form>
      ${terminRows(p.ID).length ? '<div class="notice warning" style="margin-top:10px">Termin lama proyek ini akan diganti.</div>' : ''}`,
    foot: `<button class="btn" data-close>Batal</button><button class="btn primary" id="tmOk">Buat termin</button>`,
    onOpen: (el, m) => {
      $('#tmOk', el).onclick = () => {
        const d = formData($('#tmF', el));
        const total = (bases.find(b => b[0] === d.BASIS) || bases[0])[2];
        const step = num(d.BULAT) || 1;
        let plan = [];
        if (d.MODE === 'MINGGUAN') {
          if (!t.weeks.length) return toast('Isi durasi per divisi dulu', 'err');
          const sumOh = t.weeks.reduce((a, w) => a + w.oh, 0);
          plan = t.weeks.map(w => ({ label: `Upah minggu ke-${w.minggu} (${Object.keys(w.divisi).join(', ')})`, persen: sumOh ? w.oh / sumOh * 100 : 0, minggu: w.minggu }));
        } else {
          plan = d.MANUAL.split('\n').map(l => l.trim()).filter(Boolean).map(l => { const i = l.lastIndexOf(';'); return { label: i > 0 ? l.slice(0, i).trim() : l, persen: i > 0 ? num(parseNum(l.slice(i + 1))) : 0 }; });
          const sp = plan.reduce((a, x) => a + x.persen, 0);
          if (Math.abs(sp - 100) > 0.01) return toast(`Total persen ${N(sp, 2)}% — harus 100%`, 'err');
        }
        terminRows(p.ID).forEach(r => save('TERMIN_UPAH', { ID: r.ID, DELETED: 'YA' }, { silent: true }));
        let acc = 0;
        plan.forEach((x, i) => {
          let nom = i === plan.length - 1 ? Math.round(total - acc) : Calc.roundTo(total * x.persen / 100, step, 'dekat');
          acc += nom;
          save('TERMIN_UPAH', { ID: uid('TU'), PROYEK_ID: p.ID, URUTAN: i + 1, LABEL: x.label, BASIS: d.BASIS, PERSEN: Calc.round(x.persen, 3), NOMINAL: nom, MINGGU_KE: x.minggu || '', STATUS: 'DRAFT' }, { silent: true });
        });
        scheduleSync(); m.close(); toast(`${plan.length} termin dibuat`, 'ok'); rerenderSoft();
      };
    }
  });
}
async function exportTenagaXlsx(p, rk, t, bor) {
  const X = await needXlsx(), wb = X.utils.book_new();
  const roles = t.roles;
  const A = [['Divisi', 'OH', 'Durasi (hari)'].concat(roles.map(r => r.nama + ' (org)'))].concat(t.divisi.filter(d => d.ohTotal).map(d => [`${d.kode} ${d.uraian}`, Calc.round(d.ohTotal, 2), d.durasi].concat(roles.map(r => d.orang[norm(r.kode)] || 0))));
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(A), 'Per divisi');
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['Jenis tenaga', 'Kode', 'Kelas', 'Total OH', 'Upah/hari', 'Total upah', 'Orang puncak']].concat(roles.map(r => [r.nama, r.kode, ROLE_LABEL[r.kelas], Calc.round(r.oh, 3), r.harga, Math.round(r.upah), r.puncak]))), 'Rekap tenaga');
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['Minggu', 'Orang (puncak)', 'OH', 'Divisi']].concat(t.weeks.map(w => [w.minggu, w.orang, Calc.round(w.oh, 2), Object.keys(w.divisi).join(', ')]))), 'Histogram');
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['Metode', 'Total upah'], ['AHSP', Math.round(t.totalUpah)], ['Borongan per m²', bor.totM2], ['Borongan per item', bor.totItem], ['Harian', bor.totHarian]]), 'Perbandingan');
  const tm = terminRows(p.ID);
  if (tm.length) X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['#', 'Termin', 'Basis', '%', 'Nominal', 'Status']].concat(tm.map(r => [num(r.URUTAN), r.LABEL, r.BASIS, num(r.PERSEN), num(r.NOMINAL), r.STATUS]))), 'Termin upah');
  X.writeFile(wb, `${p.KODE}_TENAGA_${todayStr().replace(/-/g, '')}.xlsx`);
}
async function exportTenagaPdf(p, rk, t, bor) {
  const prof = profilOf(p.PROFIL_ID), ctx = await pdfStart(prof), doc = ctx.doc;
  let y = pdfKop(ctx);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.text('REKAP KEBUTUHAN TENAGA KERJA', 105, y + 2, { align: 'center' });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.text(`${p.NAMA} · ${p.LOKASI || ''} · ${fmtDateLong(todayStr())}`, 105, y + 8, { align: 'center' });
  const o = tenagaOpt();
  doc.autoTable(AT({ startY: y + 12, head: [['Jenis tenaga', 'Total OH', 'Upah/hari', 'Total upah', 'Puncak']], body: t.roles.map(r => [r.nama, N(r.oh, 2), rpS(r.harga), rpS(r.upah), r.puncak ? r.puncak + ' org' : '-']).concat([styled(['Total', N(t.totalOH, 2), '', rpS(t.totalUpah), ''], 'TOT')]), columnStyles: { 1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' } } }));
  const roles = t.roles;
  doc.autoTable(AT({ startY: doc.lastAutoTable.finalY + 6, head: [['Divisi', 'OH', 'Hari'].concat(roles.map(r => r.nama))], body: t.divisi.filter(d => d.ohTotal).map(d => [`${d.kode} ${d.uraian}`, N(d.ohTotal, 1), d.durasi || '-'].concat(roles.map(r => d.orang[norm(r.kode)] || '-'))).concat([styled(['Total', N(t.totalOH, 1), t.totalDurasi].concat(roles.map(() => '')), 'TOT')]), styles: { fontSize: 7.5, cellPadding: 1.5 } }));
  let yy = doc.lastAutoTable.finalY + 6;
  doc.setFontSize(8); doc.setTextColor(...hexRgb(THEME.muted));
  doc.text(doc.splitTextToSize(`Jumlah orang = OH ÷ durasi, dibulatkan ke atas. Jam kerja ${o.jamKerja} jam/hari, ${o.hariMinggu} hari/minggu. Durasi total ${t.totalDurasi} hari kerja (± ${t.minggu} minggu, perkiraan divisi berurutan).`, 180), 15, yy);
  doc.setTextColor(...hexRgb(THEME.text));
  doc.autoTable(AT({ startY: yy + 8, head: [['Perbandingan metode upah', 'Total']], body: [['AHSP', rpS(t.totalUpah)], ['Borongan per m²', rpS(bor.totM2)], ['Borongan per item', rpS(bor.totItem)], ['Harian', rpS(bor.totHarian)]], columnStyles: { 1: { halign: 'right', cellWidth: 45 } } }));
  const tm = terminRows(p.ID);
  if (tm.length) doc.autoTable(AT({ startY: doc.lastAutoTable.finalY + 6, head: [['#', 'Termin upah', '%', 'Nominal']], body: tm.map(r => [num(r.URUTAN), r.LABEL, pct(r.PERSEN, 1), rpS(r.NOMINAL)]), columnStyles: { 2: { halign: 'right' }, 3: { halign: 'right' } } }));
  pdfFooter(ctx, `Tenaga ${p.KODE}`, norm(setting('TAMPIL_CREDIT')) === 'YA');
  await deliverPdf(doc, `${p.KODE}_TENAGA_${todayStr().replace(/-/g, '')}.pdf`, { proyekId: p.ID, kategori: '01_RAB', jenis: 'TENAGA', terkait: p.ID, shareText: `Rekap tenaga ${p.NAMA}: ${N(t.totalOH, 1)} OH, ${t.totalDurasi} hari kerja, upah ${rp(t.totalUpah)}` });
}
/* ============================================================
 * Tahap 3 — Jadwal: Gantt, jalur kritis, Kurva S, progres,
 * laporan mingguan, cash flow, libur & milestone
 * ============================================================ */
const JTABS = [['gantt', 'Gantt'], ['kurva', 'Kurva S'], ['progres', 'Input Progres'], ['laporan', 'Laporan'], ['kalender', 'Libur & Milestone']];
const progRows = pid => rows('PROGRES').filter(r => r.PROYEK_ID === pid);
const msRows = pid => rows('MILESTONE').filter(r => r.PROYEK_ID === pid);
/* Aturan urutan kerja default (kode divisi → [pendahulu, relasi, lag]) */
const JRULES = {
  III: ['I', 'FS'], IV: ['III', 'FS'], V: ['IV', 'FS'], VI: ['V', 'FS'], VII: ['VI', 'FS'],
  XII: ['VII', 'SS', 'HALF'], XIV: ['VII', 'SS', 'HALF'], X: ['VII', 'FS'], IX: ['X', 'FS'], VIII: ['IX', 'FS'], XI: ['VIII', 'FS'],
  XIII: ['XI', 'FS'], XV: ['XI', 'FS'], XVI: ['XI', 'FS'], XVII: ['XIII', 'FS'], XVIII: ['XVII', 'SS'], XIX: ['XVII', 'FS']
};
const JCHAIN = ['I', 'III', 'IV', 'V', 'VI', 'VII', 'X', 'IX', 'VIII', 'XI', 'XIII', 'XVII', 'XIX'];

function jadwalModel(p) {
  const rk = rekapOf(p), t = tenagaOf(p, rk);
  const start = String(p.TANGGAL_MULAI || '').slice(0, 10) || todayStr();
  const hm = num(setting('HARI_KERJA_MINGGU')) || 6;
  const cal = Calc.calendar(start, hm, rows('LIBUR').map(r => r.TANGGAL), 1500);
  const J = {}; jadwalRows(p.ID).forEach(r => J[r.ITEM_ID] = r);
  const td = {}; t.divisi.forEach(d => td[d.id] = d);
  const groups = rk.groups.filter(g => g.row.ID !== '_');
  const tasks = groups.map(g => {
    const j = J[g.row.ID] || {}, smkk = Calc.isSmkk(g.row);
    return { id: g.row.ID, dur: smkk ? 0 : num((td[g.row.ID] || {}).durasi), pred: j.PREDECESSOR || '', rel: j.JENIS_RELASI || 'FS', lag: num(j.LAG), minStart: j.MULAI || '', span: smkk };
  });
  const sched = Calc.schedule(tasks, cal);
  // jadwal item (opsional): mulai = mulai divisi + offset (LAG), durasi sendiri
  const itemSched = {};
  groups.forEach(g => {
    const r = sched.tasks[g.row.ID]; if (!r) return;
    g.items.forEach(it => {
      const j = J[it.ID]; if (!j || !(num(j.DURASI) > 0)) return;
      const es = r.es + Math.max(0, Math.round(num(j.LAG))), ef = es + Math.round(num(j.DURASI)) - 1;
      itemSched[it.ID] = { es, ef, dur: ef - es + 1, mulai: cal[es], selesai: cal[ef], over: ef > r.ef };
    });
  });
  const C = rk.C || 1;
  const bobot = {}; groups.forEach(g => bobot[g.row.ID] = g.subtotal / C);
  const kurva = Calc.kurvaS(groups.map(g => ({ id: g.row.ID, bobot: bobot[g.row.ID] })), sched, start, cal);
  const P = progRows(p.ID), byItem = {};
  P.forEach(r => (byItem[r.ITEM_ID] = byItem[r.ITEM_ID] || []).push(r));
  const latest = (id, w) => { let b = null; (byItem[id] || []).forEach(r => { if (num(r.MINGGU_KE) <= w && (!b || num(r.MINGGU_KE) > num(b.MINGGU_KE))) b = r; }); return b ? num(b.PERSEN_AKTUAL) : null; };
  const pctDiv = (g, w) => {
    const itemHas = g.items.some(it => latest(it.ID, w) !== null);
    if (itemHas && g.subtotal > 0) return g.items.reduce((a, it) => a + rk.computed[it.ID].jumlah * (latest(it.ID, w) || 0), 0) / g.subtotal;
    const v = latest(g.row.ID, w); return v === null ? 0 : v;
  };
  const actualCum = w => groups.reduce((a, g) => a + bobot[g.row.ID] * pctDiv(g, w), 0);
  const weeksWithData = [...new Set(P.filter(r => r.ITEM_ID !== '_MINGGU').map(r => num(r.MINGGU_KE)))].sort((a, b) => a - b);
  const lastData = weeksWithData.length ? weeksWithData[weeksWithData.length - 1] : 0;
  const weekNow = Calc.weekOf(start, todayStr());
  const planAt = w => w <= 0 ? 0 : (kurva[w - 1] ? kurva[w - 1].kum : (w > kurva.length && kurva.length ? 100 : 0));
  const scheduled = sched.durasiKerja > 0;
  return { p, rk, t, start, cal, groups, tasks, sched, itemSched, bobot, kurva, pctDiv, actualCum, latest, weekNow, lastData, planAt, scheduled, J, td };
}
function devStatus(dev) {
  const wa = num(setting('DEVIASI_WASPADA')) || 5, kr = num(setting('DEVIASI_KRITIS')) || 10;
  if (dev <= -kr) return ['TERLAMBAT', 'danger'];
  if (dev <= -wa) return ['WASPADA', 'warning'];
  if (dev >= wa) return ['LEBIH CEPAT', 'success'];
  return ['SESUAI', 'success'];
}
function autoMilestones(M) {
  const s = M.sched.tasks, byKode = {};
  M.groups.forEach(g => byKode[g.row.KODE] = s[g.row.ID]);
  const endOf = codes => { const e = codes.map(c => byKode[c]).filter(r => r && !r.span); return e.length ? e.reduce((a, r) => r.ef > a.ef ? r : a) : null; };
  const out = [];
  const st = endOf(['V']) || endOf(['IV']); if (st) out.push({ nama: 'Selesai struktur', tanggal: st.selesai, auto: true });
  const at = endOf(['VI']); if (at) out.push({ nama: 'Tutup atap', tanggal: at.selesai, auto: true });
  const mep = endOf(['XIII', 'XIV', 'XV', 'XVI']); if (mep) out.push({ nama: 'MEP terpasang', tanggal: mep.selesai, auto: true });
  if (M.scheduled) out.push({ nama: 'Serah terima (PHO)', tanggal: M.sched.selesai, auto: true });
  return out;
}

route('/jadwal', () => {
  const list = projects();
  if (!list.length) { setTitle('Jadwal'); view(`<div class="page-head"><h2>Jadwal & Kurva S</h2></div><div class="card empty"><p>Belum ada proyek.</p><a class="btn primary" href="#/rab">Buat proyek</a></div>`); return; }
  const pid = getRow('PROYEK', S.meta.activeProject) ? S.meta.activeProject : list[0].ID;
  location.replace(`#/jadwal/${pid}/gantt`);
});
route('/jadwal/:id/:tab', (id, tab) => jadwalView(id, tab));
function jadwalView(id, tab) {
  const p = getRow('PROYEK', id);
  if (!p) { location.hash = '#/jadwal'; return; }
  if (S.meta.activeProject !== id) { S.meta.activeProject = id; saveMeta(); }
  setTitle('Jadwal · ' + p.NAMA, p.KODE);
  const M = jadwalModel(p);
  const list = projects();
  view(`<div class="page-head"><div><h2>Jadwal & Kurva S</h2></div>
      <div class="actions"><select class="input" id="jpSel" style="min-width:200px">${list.map(x => `<option value="${x.ID}" ${x.ID === id ? 'selected' : ''}>${esc(x.KODE)} · ${esc(x.NAMA)}</option>`).join('')}</select></div></div>
    <div class="card tight" style="margin-bottom:12px"><div class="row">
      <div class="field" style="min-width:160px"><label>Tanggal mulai proyek</label><input class="input" type="date" id="jStart" value="${esc(String(p.TANGGAL_MULAI || '').slice(0, 10))}"></div>
      <div class="kv" style="flex:1;min-width:220px"><dt>Selesai (rencana)</dt><dd>${M.scheduled ? fmtDateLong(M.sched.selesai) : '-'}</dd><dt>Durasi</dt><dd>${M.scheduled ? `${N(M.sched.durasiKerja)} hari kerja · ${N(Calc.daysBetween(M.sched.mulai, M.sched.selesai) + 1)} hari kalender` : '-'}</dd><dt>Minggu berjalan</dt><dd>${M.weekNow > 0 ? 'ke-' + M.weekNow : 'belum mulai'}</dd></div>
      <div class="row"><button class="btn" id="jAuto">Susun urutan otomatis</button><button class="btn primary" id="jPdf">${ICON.pdf} PDF jadwal</button><button class="btn ghost" id="jXls">Excel</button></div></div>
      ${!p.TANGGAL_MULAI ? '<div class="notice warning" style="margin-top:10px">Tanggal mulai belum diisi — sementara memakai tanggal hari ini.</div>' : ''}
      ${M.sched.warn.length ? `<div class="notice warning" style="margin-top:10px">${ICON.warn}<div>${M.sched.warn.map(w => { const g = M.groups.find(x => x.row.ID === w.id); return `${esc(g ? g.row.URAIAN : w.id)}: ${w.code === 'SIKLUS' ? 'ketergantungan melingkar (diabaikan)' : 'pendahulu belum punya durasi'}`; }).join('<br>')}</div></div>` : ''}</div>
    <div class="tabs">${JTABS.map(([k, t]) => `<button data-tab="${k}" class="${k === tab ? 'active' : ''}">${t}</button>`).join('')}</div><div id="jBody"></div>`);
  $('#jpSel').onchange = e => { location.hash = `#/jadwal/${e.target.value}/${tab}`; };
  $$('.tabs button').forEach(b => b.onclick = () => { location.hash = `#/jadwal/${id}/${b.dataset.tab}`; });
  $('#jStart').onchange = e => { save('PROYEK', { ID: id, TANGGAL_MULAI: e.target.value }); rerenderSoft(); };
  $('#jAuto').onclick = () => autoArrange(M);
  $('#jPdf').onclick = () => exportJadwalPdf(M).catch(e => { console.error(e); toast('Gagal membuat PDF: ' + e.message, 'err'); });
  $('#jXls').onclick = () => exportJadwalXlsx(M);
  ({ gantt: tabGantt, kurva: tabKurva, progres: tabProgres, laporan: tabLaporan, kalender: tabKalender }[tab] || tabGantt)(M, $('#jBody'));
}

/* ---------- Susun urutan otomatis ---------- */
async function autoArrange(M) {
  if (!await confirmBox('Susun ulang urutan pekerjaan sesuai template standar?\nPersiapan → Tanah → Pondasi → Struktur → Atap → Dinding (+ instalasi plumbing & listrik tertanam) → Kusen → Plafon → Lantai → Sanitair → MEP → Finishing → Pekerjaan luar (paralel).\n\nPendahulu & lag yang sudah ada akan ditimpa. Durasi tidak berubah.', { ok: 'Susun otomatis' })) return;
  const byKode = {}; M.groups.forEach(g => { if (g.row.KODE) byKode[g.row.KODE] = g; });
  let prevId = '';
  M.groups.forEach(g => {
    if (Calc.isSmkk(g.row)) return;
    const code = g.row.KODE, rule = JRULES[code];
    let pred = '', rel = 'FS', lag = 0;
    if (rule) {
      let pc = rule[0];
      if (!byKode[pc]) { const i = JCHAIN.indexOf(pc); for (let k = i - 1; k >= 0; k--) if (byKode[JCHAIN[k]]) { pc = JCHAIN[k]; break; } }
      if (byKode[pc] && byKode[pc] !== g) {
        pred = byKode[pc].row.ID; rel = rule[1];
        if (rule[2] === 'HALF') lag = Math.ceil(num((M.td[pred] || {}).durasi) / 2);
      }
    } else if (code !== 'I' && prevId) pred = prevId;
    const j = M.J[g.row.ID];
    save('JADWAL', { ID: j ? j.ID : uid('J'), PROYEK_ID: M.p.ID, ITEM_ID: g.row.ID, PREDECESSOR: pred, JENIS_RELASI: rel, LAG: lag }, { silent: true });
    prevId = g.row.ID;
  });
  scheduleSync(); toast('Urutan disusun. Lengkapi durasi divisi yang masih kosong.', 'ok'); rerenderSoft();
}

/* ---------- Gantt ---------- */
let GZOOM = 'M', GEXP = {};
function tabGantt(M, el) {
  const pxd = { H: 26, M: 9, B: 3.5 }[GZOOM];
  const LW = window.innerWidth < 640 ? 130 : 230;
  const start = M.start;
  const end = M.scheduled ? M.sched.selesai : Calc.addDays(start, 27);
  const days = Math.max(28, Calc.daysBetween(start, end) + 8);
  const W = Math.ceil(days * pxd);
  const x = d => Calc.daysBetween(start, d) * pxd;
  const work = new Set(M.cal);
  const ms = autoMilestones(M).concat(msRows(M.p.ID).map(r => ({ nama: r.NAMA, tanggal: r.TANGGAL, status: r.STATUS, id: r.ID })));
  // header bulan
  let months = '', d0 = start;
  for (let i = 0; i < days;) {
    const dt = Calc.parseYmd(Calc.addDays(start, i)), y = dt.getUTCFullYear(), m = dt.getUTCMonth();
    const nextMonth = Calc.ymd(new Date(Date.UTC(y, m + 1, 1)));
    const span = Math.min(days, Calc.daysBetween(start, nextMonth)) - i;
    months += `<div class="g-mo" style="left:${i * pxd}px;width:${span * pxd}px">${span * pxd > 40 ? BULAN[m].slice(0, 3) + ' ' + String(y).slice(2) : ''}</div>`;
    i += span;
  }
  let sub = '';
  if (GZOOM === 'H') for (let i = 0; i < days; i++) { const s = Calc.addDays(start, i); sub += `<div class="g-dy ${work.has(s) ? '' : 'off'}" style="left:${i * pxd}px;width:${pxd}px">${+s.slice(8)}</div>`; }
  else if (GZOOM === 'M') for (let w = 0; w * 7 < days; w++) sub += `<div class="g-dy" style="left:${w * 7 * pxd}px;width:${7 * pxd}px">M${w + 1}</div>`;
  let shade = '';
  if (GZOOM === 'H') for (let i = 0; i < days; i++) { const s = Calc.addDays(start, i); if (!work.has(s)) shade += `<div class="g-off" style="left:${i * pxd}px;width:${pxd}px"></div>`; }
  const today = todayStr(), tx = Calc.daysBetween(start, today);
  const bar = (r, cls, label, title) => `<div class="g-bar ${cls}" style="left:${x(r.mulai)}px;width:${Math.max(pxd, (Calc.daysBetween(r.mulai, r.selesai) + 1) * pxd)}px" title="${esc(title)}">${(Calc.daysBetween(r.mulai, r.selesai) + 1) * pxd > 60 ? esc(label) : ''}</div>`;
  const rowsHtml = M.groups.map(g => {
    const r = M.sched.tasks[g.row.ID], td = M.td[g.row.ID] || {}, j = M.J[g.row.ID] || {};
    const predG = j.PREDECESSOR ? M.groups.find(x => x.row.ID === j.PREDECESSOR) : null;
    const pct = M.pctDiv(g, Math.max(1, M.weekNow));
    const info = r ? `${fmtDate(r.mulai)} – ${fmtDate(r.selesai)} · ${r.dur} hari${r.span ? ' (sepanjang proyek)' : ''}${!r.span && r.float > 0 ? ' · float ' + r.float + ' hari' : ''}${predG ? ` · setelah ${predG.row.KODE || predG.row.URAIAN} (${j.JENIS_RELASI || 'FS'}${num(j.LAG) ? (num(j.LAG) > 0 ? '+' : '') + num(j.LAG) : ''})` : ''}` : 'belum dijadwalkan';
    const cls = r ? (r.span ? 'span' : r.critical ? 'crit' : '') : '';
    const items = g.items.filter(it => num(it.VOLUME) > 0);
    const expanded = GEXP[g.row.ID];
    let html = `<div class="g-row"><div class="g-lab" data-gdiv="${g.row.ID}" style="width:${LW}px"><span class="g-tog" data-gexp="${g.row.ID}">${items.length ? (expanded ? '▾' : '▸') : ''}</span><div style="min-width:0"><div class="g-t">${esc(g.row.KODE || '')} ${esc(g.row.URAIAN)}</div><div class="g-s">${r ? (r.span ? 'sepanjang proyek' : r.dur + ' hari' + (r.critical ? ' · kritis' : '')) : '<span class="neg">durasi kosong</span>'}</div></div></div>
      <div class="g-tl" style="width:${W}px">${r ? bar(r, cls, `${g.row.KODE || ''} ${r.dur}h`, `${g.row.URAIAN}\n${info}\nProgres aktual: ${N(pct, 1)}%`) : ''}${r && pct > 0 && !r.span ? `<div class="g-prog" style="left:${x(r.mulai)}px;width:${Math.max(2, (Calc.daysBetween(r.mulai, r.selesai) + 1) * pxd * Math.min(100, pct) / 100)}px"></div>` : ''}</div></div>`;
    if (expanded) items.forEach(it => {
      const s = M.itemSched[it.ID] || (r && !r.span ? { mulai: r.mulai, selesai: r.selesai, inherit: true } : null);
      html += `<div class="g-row g-item"><div class="g-lab" style="width:${LW}px;padding-left:28px"><div style="min-width:0"><div class="g-t" style="font-weight:500">${esc(it.URAIAN)}</div><div class="g-s">${s && !s.inherit ? `${M.itemSched[it.ID].dur} hari${s.over ? ' · <span class="neg">melewati divisi</span>' : ''}` : 'ikut divisi'}</div></div></div>
        <div class="g-tl" style="width:${W}px">${s ? bar(s, 'item' + (s.inherit ? ' inherit' : ''), '', `${it.URAIAN}\n${fmtDate(s.mulai)} – ${fmtDate(s.selesai)}`) : ''}</div></div>`;
    });
    return html;
  }).join('');
  const msHtml = ms.filter(m => m.tanggal).map(m => `<div class="g-ms" style="left:${x(m.tanggal) + pxd / 2 - 6}px" title="${esc(m.nama)} · ${fmtDate(m.tanggal)}${m.auto ? ' (otomatis)' : ''}"></div>`).join('');
  el.innerHTML = `<div class="row" style="margin-bottom:10px"><div class="seg" id="gZoom">${[['H', 'Harian'], ['M', 'Mingguan'], ['B', 'Bulanan']].map(([k, t]) => `<button data-z="${k}" class="${GZOOM === k ? 'active' : ''}">${t}</button>`).join('')}</div>
      <button class="btn sm ghost" id="gExpAll">${Object.values(GEXP).some(Boolean) ? 'Tutup semua item' : 'Buka semua item'}</button><span class="spacer"></span>
      <span class="small muted"><span class="g-key crit"></span> kritis <span class="g-key"></span> normal <span class="g-key span"></span> SMKK <span class="g-key ms"></span> milestone <span class="g-key now"></span> hari ini</span></div>
    ${M.groups.length ? `<div class="gantt"><div class="g-inner" style="width:${LW + W}px">
      <div class="g-row g-head"><div class="g-lab" style="width:${LW}px"><strong>Pekerjaan</strong></div><div class="g-tl" style="width:${W}px">${months}</div></div>
      <div class="g-row g-head2"><div class="g-lab" style="width:${LW}px"><span class="small muted">Ketuk nama untuk atur</span></div><div class="g-tl" style="width:${W}px">${sub}</div></div>
      <div class="g-row g-msrow"><div class="g-lab" style="width:${LW}px"><span class="small muted">Milestone</span></div><div class="g-tl" style="width:${W}px">${msHtml}</div></div>
      <div class="g-body" style="position:relative">${shade ? `<div class="g-shade" style="left:${LW}px;width:${W}px">${shade}</div>` : ''}${rowsHtml}</div>
      ${tx >= 0 && tx <= days ? `<div class="g-now" style="left:${LW + tx * pxd + pxd / 2}px"></div>` : ''}
    </div></div>` : '<div class="card empty"><p>Belum ada divisi pekerjaan di RAB.</p></div>'}
    <div class="small muted" style="margin-top:8px">Jalur kritis = pekerjaan tanpa kelonggaran waktu; terlambat sehari di sini, proyek ikut mundur. Durasi divisi diambil dari tab Tenaga (bisa diubah di sini). Garis hijau tipis di bar = progres aktual.</div>`;
  $$('#gZoom button', el).forEach(b => b.onclick = () => { GZOOM = b.dataset.z; rerenderSoft(); });
  $('#gExpAll', el).onclick = () => { const open = !Object.values(GEXP).some(Boolean); GEXP = {}; if (open) M.groups.forEach(g => GEXP[g.row.ID] = true); rerenderSoft(); };
  el.onclick = e => {
    const tg = e.target.closest('[data-gexp]');
    if (tg) { GEXP[tg.dataset.gexp] = !GEXP[tg.dataset.gexp]; rerenderSoft(); return; }
    const lb = e.target.closest('[data-gdiv]');
    if (lb) taskEditor(M, lb.dataset.gdiv);
  };
  const gw = $('.gantt', el);
  if (gw && tx > 0) gw.scrollLeft = Math.max(0, tx * pxd - 120);
}
function taskEditor(M, did) {
  const g = M.groups.find(x => x.row.ID === did); if (!g) return;
  const j = M.J[did] || {}, td = M.td[did] || {}, r = M.sched.tasks[did];
  const smkk = Calc.isSmkk(g.row), tukangMode = td.mode === 'TUKANG';
  const f = M.t.faktorJam || 1;
  const sugg = td.ohTotal ? [3, 5, 8, 12].map(n => [n, Math.max(1, Math.ceil(td.ohTotal * f / n))]) : [];
  const items = g.items.filter(it => num(it.VOLUME) > 0);
  openModal({
    title: `${g.row.KODE || ''} ${g.row.URAIAN}`, wide: true,
    body: smkk ? `<div class="notice">Divisi SMKK berjalan sepanjang proyek (mulai hari pertama s.d. serah terima), otomatis mengikuti durasi total.</div>` : `<form id="tkF" class="form-grid c2">
        <div class="field"><label>Durasi (hari kerja)</label><input class="input num" name="DURASI" inputmode="numeric" value="${tukangMode ? '' : (num(td.durasi) || '')}" placeholder="${tukangMode ? N(td.durasi) + ' (dari jumlah tukang)' : 'hari'}" ${tukangMode ? 'disabled' : ''}>
          <span class="hint">${tukangMode ? `Dihitung dari ${num((tenagaPlan(M.p.ID)[did] || {}).tukang)} tukang tersedia di tab Tenaga. <a href="#/rab/${M.p.ID}/tenaga">Ubah di Tenaga</a>` : td.ohTotal ? `Kebutuhan ${N(td.ohTotal, 1)} OH. Saran: ${sugg.map(([n, d]) => `<a href="#" data-sg="${d}">${n} org → ${d} hari</a>`).join(' · ')}` : 'Divisi ini tidak punya koefisien tenaga (harga pasar/LS) — isi durasi manual.'}</span></div>
        <div class="field"><label>Mulai paling cepat (opsional)</label><input class="input" type="date" name="MULAI" value="${esc(String(j.MULAI || '').slice(0, 10))}"><span class="hint">Batas awal, misal menunggu izin/material</span></div>
        ${sel('PREDECESSOR', 'Dikerjakan setelah', j.PREDECESSOR || '', [['', '— tidak ada (mulai awal proyek) —']].concat(M.groups.filter(x => x.row.ID !== did && !Calc.isSmkk(x.row)).map(x => [x.row.ID, `${x.row.KODE || ''} ${x.row.URAIAN}`])), { cls: 'span2' })}
        ${sel('JENIS_RELASI', 'Jenis hubungan', j.JENIS_RELASI || 'FS', [['FS', 'FS — mulai setelah pendahulu selesai'], ['SS', 'SS — mulai bersamaan / setelah pendahulu mulai']])}
        ${inp('LAG', 'Lag (hari kerja)', j.LAG, { numType: 'num', hint: 'Jeda tambahan. Contoh: cor plat → bongkar bekisting 21–28 hari' })}
      </form>
      ${items.length ? `<h3 style="margin:16px 0 8px;font-size:14px">Jadwal per item (opsional)</h3><div class="small muted" style="margin-bottom:6px">Kosongkan durasi agar item mengikuti jadwal divisi. Offset = mulai berapa hari kerja setelah divisi mulai.</div>
      <div class="table-wrap"><table class="t"><thead><tr><th>Item</th><th class="num">Durasi</th><th class="num">Offset</th></tr></thead><tbody>${items.map(it => { const ji = M.J[it.ID] || {}; return `<tr data-it="${it.ID}" data-jid="${ji.ID || ''}"><td>${esc(it.URAIAN)}</td><td class="num"><input class="input num" data-f="DURASI" inputmode="numeric" value="${num(ji.DURASI) || ''}" style="width:70px"></td><td class="num"><input class="input num" data-f="LAG" inputmode="numeric" value="${num(ji.LAG) || ''}" style="width:70px"></td></tr>`; }).join('')}</tbody></table></div>` : ''}
      ${r && !r.span ? `<div class="small muted" style="margin-top:10px">Saat ini: ${fmtDate(r.mulai)} – ${fmtDate(r.selesai)} · float ${r.float} hari${r.critical ? ' · <strong class="neg">jalur kritis</strong>' : ''}</div>` : ''}`,
    foot: smkk ? '<button class="btn" data-close>Tutup</button>' : `<button class="btn" data-close>Batal</button><button class="btn primary" id="tkOk">Simpan</button>`,
    onOpen: (el, m) => {
      $$('[data-sg]', el).forEach(a => a.onclick = e => { e.preventDefault(); $('[name=DURASI]', el).value = a.dataset.sg; });
      $('#tkOk', el) && ($('#tkOk', el).onclick = () => {
        const d = formData($('#tkF', el));
        const patch = { ID: j.ID || uid('J'), PROYEK_ID: M.p.ID, ITEM_ID: did, MULAI: d.MULAI, PREDECESSOR: d.PREDECESSOR, JENIS_RELASI: d.JENIS_RELASI, LAG: d.LAG === '' ? 0 : d.LAG };
        if (!tukangMode) { patch.DURASI = d.DURASI === '' ? '' : Math.max(0, Math.round(num(d.DURASI))); }
        // cegah siklus sederhana
        if (d.PREDECESSOR) { let c = d.PREDECESSOR, guard = 0; while (c && guard++ < 50) { if (c === did) return toast('Hubungan melingkar — pilih pendahulu lain', 'err'); c = (M.J[c] || {}).PREDECESSOR; } }
        save('JADWAL', patch, { silent: true });
        $$('tr[data-it]', el).forEach(tr => {
          const du = parseNum($('[data-f=DURASI]', tr).value), lg = parseNum($('[data-f=LAG]', tr).value), jid = tr.dataset.jid;
          if (du === '' || !(num(du) > 0)) { if (jid) save('JADWAL', { ID: jid, DURASI: '', LAG: '' }, { silent: true }); return; }
          save('JADWAL', { ID: jid || uid('J'), PROYEK_ID: M.p.ID, ITEM_ID: tr.dataset.it, DURASI: Math.round(du), LAG: lg === '' ? 0 : Math.round(lg) }, { silent: true });
        });
        scheduleSync(); m.close(); rerenderSoft();
      });
    }
  });
}

/* ---------- Kurva S ---------- */
function kurvaSvg(M, w = 760, h = 300) {
  const K = M.kurva; if (!K.length) return '';
  const pad = { l: 40, r: 12, t: 14, b: 30 }, n = K.length;
  const X = i => pad.l + i / n * (w - pad.l - pad.r), Y = v => pad.t + (1 - v / 100) * (h - pad.t - pad.b);
  const plan = [[X(0), Y(0)]].concat(K.map((k, i) => [X(i + 1), Y(k.kum)]));
  const lastW = Math.min(n, M.lastData);
  const act = lastW ? [[X(0), Y(0)]].concat(Array.from({ length: lastW }, (_, i) => [X(i + 1), Y(Math.min(100, M.actualCum(i + 1)))])) : [];
  const maxInc = Math.max(1, ...K.map(k => k.rencana));
  const bars = K.map((k, i) => `<rect x="${X(i) + 1}" y="${Y(k.rencana / maxInc * 30)}" width="${Math.max(1, X(1) - X(0) - 2)}" height="${Y(0) - Y(k.rencana / maxInc * 30)}" fill="${THEME.primary50}"><title>Minggu ${k.minggu}: rencana +${N(k.rencana, 2)}%</title></rect>`).join('');
  const grid = [0, 25, 50, 75, 100].map(v => `<line x1="${pad.l}" x2="${w - pad.r}" y1="${Y(v)}" y2="${Y(v)}" stroke="${THEME.border}"/><text x="${pad.l - 6}" y="${Y(v) + 4}" font-size="10" text-anchor="end" fill="${THEME.muted}">${v}%</text>`).join('');
  const step = Math.ceil(n / 12);
  const xl = K.map((k, i) => (i + 1) % step === 0 || i === n - 1 ? `<text x="${X(i + 1)}" y="${h - 10}" font-size="10" text-anchor="middle" fill="${THEME.muted}">M${k.minggu}</text>` : '').join('');
  const nowX = M.weekNow > 0 && M.weekNow <= n ? `<line x1="${X(M.weekNow)}" x2="${X(M.weekNow)}" y1="${pad.t}" y2="${Y(0)}" stroke="${THEME.accent}" stroke-dasharray="4 3"/>` : '';
  const pts = a => a.map(p => p.join(',')).join(' ');
  return `<svg viewBox="0 0 ${w} ${h}" style="width:100%;height:auto;display:block" role="img" aria-label="Kurva S">${grid}${bars}${nowX}
    <polyline points="${pts(plan)}" fill="none" stroke="${THEME.primary}" stroke-width="2.5"/>
    ${act.length ? `<polyline points="${pts(act)}" fill="none" stroke="${THEME.accent}" stroke-width="2.5"/>${act.slice(1).map((p, i) => `<circle cx="${p[0]}" cy="${p[1]}" r="3" fill="${THEME.accent}"><title>Minggu ${i + 1}: aktual ${N(M.actualCum(i + 1), 1)}%</title></circle>`).join('')}` : ''}
    ${xl}</svg>`;
}
function tabKurva(M, el) {
  if (!M.kurva.length) { el.innerHTML = '<div class="card empty"><p>Kurva S muncul setelah divisi punya durasi di tab Gantt/Tenaga.</p></div>'; return; }
  const lastW = M.lastData;
  const cur = Math.max(0, Math.min(M.kurva.length, lastW || M.weekNow));
  const dev = cur ? M.actualCum(cur) - M.planAt(cur) : 0, st = devStatus(dev);
  el.innerHTML = `<div class="grid k4">
      <div class="card kpi"><div class="label">Rencana s.d. minggu ${cur || '-'}</div><div class="value">${cur ? pct(M.planAt(cur), 1) : '-'}</div></div>
      <div class="card kpi"><div class="label">Aktual s.d. minggu ${cur || '-'}</div><div class="value">${lastW ? pct(M.actualCum(cur), 1) : '-'}</div><div class="hint">${lastW ? 'data progres terakhir minggu ke-' + lastW : 'belum ada input progres'}</div></div>
      <div class="card kpi"><div class="label">Deviasi</div><div class="value ${dev < 0 ? 'neg' : ''}">${lastW ? (dev > 0 ? '+' : '') + N(dev, 1) + '%' : '-'}</div><div class="hint">${lastW ? badge(st[0], st[1]) : ''}</div></div>
      <div class="card kpi"><div class="label">Estimasi selesai</div><div class="value" style="font-size:16px">${estimasiSelesai(M) || '-'}</div><div class="hint">rencana ${fmtDate(M.sched.selesai)}</div></div></div>
    <div class="card" style="margin-top:12px"><div class="row between"><h3>Kurva S</h3><span class="small muted"><span class="g-key"></span> rencana <span class="g-key acc"></span> aktual <span class="g-key now"></span> minggu ini</span></div>${kurvaSvg(M)}</div>
    <div class="card" style="margin-top:12px"><h3>Tabel mingguan</h3><div class="table-wrap"><table class="t"><thead><tr><th>Minggu</th><th>Periode</th><th class="num">Rencana minggu ini</th><th class="num">Rencana kumulatif</th><th class="num">Aktual kumulatif</th><th class="num">Deviasi</th><th>Status</th></tr></thead><tbody>
      ${M.kurva.map(k => { const has = k.minggu <= lastW; const a = has ? M.actualCum(k.minggu) : null, d = has ? a - k.kum : null, s = has ? devStatus(d) : null; return `<tr${k.minggu === M.weekNow ? ' class="sub"' : ''}><td>${k.minggu}</td><td class="small">${fmtDate(k.mulai)} – ${fmtDate(k.selesai)}</td><td class="num">${pct(k.rencana, 2)}</td><td class="num">${pct(k.kum, 2)}</td><td class="num">${has ? pct(a, 2) : '-'}</td><td class="num ${d < 0 ? 'neg' : ''}">${has ? (d > 0 ? '+' : '') + N(d, 2) + '%' : '-'}</td><td>${has ? badge(s[0], s[1]) : ''}</td></tr>`; }).join('')}
    </tbody></table></div><div class="small muted" style="margin-top:6px">Ambang: waspada −${setting('DEVIASI_WASPADA')}%, terlambat −${setting('DEVIASI_KRITIS')}% (Master → Pengaturan). Bobot = nilai divisi ÷ total biaya pekerjaan + SMKK, disebar rata pada hari kerjanya.</div></div>`;
}
function estimasiSelesai(M) {
  // SPI (Schedule Performance Index) = aktual ÷ rencana; durasi perkiraan = durasi rencana ÷ SPI
  const w = M.lastData; if (!w || !M.scheduled) return '';
  const a = M.actualCum(w), pl = M.planAt(w); if (a <= 0 || pl <= 0) return '';
  if (a >= 99.99) return 'Selesai';
  const spi = a / pl, planDays = Calc.daysBetween(M.start, M.sched.selesai) + 1;
  return fmtDateLong(Calc.addDays(M.start, Math.max(0, Math.ceil(planDays / spi) - 1))) + ` (SPI ${N(spi, 2)})`;
}
/* ---------- Input progres ---------- */
let PWEEK = {};
function tabProgres(M, el) {
  if (!M.scheduled) { el.innerHTML = '<div class="card empty"><p>Atur durasi divisi di tab Gantt dulu.</p></div>'; return; }
  const maxW = Math.max(M.kurva.length, M.weekNow, 1);
  let w = PWEEK[M.p.ID] || Math.min(Math.max(1, M.weekNow), maxW);
  const head = progRows(M.p.ID).find(r => r.ITEM_ID === '_MINGGU' && num(r.MINGGU_KE) === w) || {};
  const fotos = String(head.FOTO || '').split('\n').filter(Boolean);
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><div class="field"><label>Minggu ke-</label><select class="input" id="pwSel" style="width:auto">${Array.from({ length: maxW + 2 }, (_, i) => i + 1).map(i => `<option value="${i}" ${i === w ? 'selected' : ''}>Minggu ${i} · ${fmtDate(Calc.addDays(M.start, (i - 1) * 7))} – ${fmtDate(Calc.addDays(M.start, i * 7 - 1))}${i === M.weekNow ? ' (berjalan)' : ''}</option>`).join('')}</select></div>
      <span class="spacer"></span><div class="small muted">Isi <strong>persen kumulatif</strong> (total selesai s.d. minggu ini, 0–100). Kosong = sama dengan minggu sebelumnya.</div></div>
    <div class="table-wrap"><table class="t" id="pgT"><thead><tr><th>Pekerjaan</th><th class="num">Bobot</th><th class="num">Rencana</th><th class="num">Aktual minggu lalu</th><th class="num">Aktual s.d. minggu ini (%)</th></tr></thead><tbody>
      ${M.groups.filter(g => M.bobot[g.row.ID] > 0).map(g => {
        const r = M.sched.tasks[g.row.ID], plan = Calc.planPct(r, w, M.start, M.cal), prev = M.pctDiv(g, w - 1);
        const own = progRows(M.p.ID).find(x => x.ITEM_ID === g.row.ID && num(x.MINGGU_KE) === w);
        const itemsV = g.items.filter(it => num(it.VOLUME) > 0);
        const itemMode = g.items.some(it => M.latest(it.ID, w) !== null);
        return `<tr data-pg="${g.row.ID}"><td><strong>${esc(g.row.KODE || '')} ${esc(g.row.URAIAN)}</strong>${itemsV.length ? `<div><a href="#" class="small" data-pexp="${g.row.ID}">${itemMode ? 'detail per item ▾' : 'isi per item ▸'}</a></div>` : ''}</td><td class="num">${pct(M.bobot[g.row.ID] * 100, 2)}</td><td class="num">${pct(plan, 0)}</td><td class="num">${pct(prev, 0)}</td>
          <td class="num"><input class="input num" data-pv inputmode="decimal" value="${own ? N(own.PERSEN_AKTUAL, 1).replace(',0', '') : ''}" placeholder="${itemMode ? 'per item' : N(prev, 0)}" ${itemMode ? 'disabled' : ''} style="width:90px"></td></tr>
          ${itemsV.map(it => { const o = progRows(M.p.ID).find(x => x.ITEM_ID === it.ID && num(x.MINGGU_KE) === w); const pv = M.latest(it.ID, w - 1); return `<tr data-pit="${it.ID}" data-of="${g.row.ID}" class="${itemMode ? '' : 'hidden'}"><td style="padding-left:28px">${esc(it.URAIAN)}</td><td class="num small">${pct(bobotItem(M, it) * 100, 2)}</td><td></td><td class="num small">${pv === null ? '-' : pct(pv, 0)}</td><td class="num"><input class="input num" data-pv inputmode="decimal" value="${o ? N(o.PERSEN_AKTUAL, 1).replace(',0', '') : ''}" placeholder="${pv === null ? '0' : N(pv, 0)}" style="width:90px"></td></tr>`; }).join('')}`;
      }).join('')}</tbody></table></div>
    <div class="grid md2" style="margin-top:12px">
      <div class="field"><label>Catatan minggu ini</label><textarea class="input" id="pgNote" rows="3" placeholder="cuaca, kendala, material datang, dll">${esc(head.CATATAN || '')}</textarea></div>
      <div class="field"><label>Foto dokumentasi</label><div class="row"><label class="btn sm">${ICON.plus} Tambah foto<input type="file" id="pgFoto" accept="image/*" multiple hidden></label><span class="small muted">dikompres & disimpan ke Drive 06_Progres</span></div>
        <div class="chips" style="margin-top:6px">${fotos.map((u, i) => `<a class="badge info" href="${esc(u)}" target="_blank" rel="noopener">Foto ${i + 1}</a>`).join('')}${S.uploads.filter(u => u.jenis === 'PROGRES' && u.proyekId === M.p.ID).length ? badge(S.uploads.filter(u => u.jenis === 'PROGRES' && u.proyekId === M.p.ID).length + ' menunggu upload', 'warning') : ''}</div></div></div>
    <div class="row" style="margin-top:12px"><span class="spacer"></span><button class="btn primary" id="pgSave">Simpan progres minggu ${w}</button></div>`;
  $('#pwSel', el).onchange = e => { PWEEK[M.p.ID] = +e.target.value; rerenderSoft(); };
  $$('[data-pexp]', el).forEach(a => a.onclick = e => { e.preventDefault(); const id = a.dataset.pexp; const trs = $$(`tr[data-of="${id}"]`, el); const show = trs.some(t => t.classList.contains('hidden')); trs.forEach(t => t.classList.toggle('hidden', !show)); $(`tr[data-pg="${id}"] [data-pv]`, el).disabled = show; a.textContent = show ? 'detail per item ▾' : 'isi per item ▸'; });
  $('#pgFoto', el).onchange = async e => {
    const files = Array.from(e.target.files || []); e.target.value = '';
    const hid = head.ID || `${M.p.ID}|${w}|_MINGGU`;
    if (!head.ID) save('PROGRES', { ID: hid, PROYEK_ID: M.p.ID, MINGGU_KE: w, ITEM_ID: '_MINGGU', TANGGAL: todayStr() });
    let n = fotos.length;
    for (const f of files) {
      try {
        const r = await compressImage(f, 1600, 0.8);
        n++;
        await queueUpload({ blob: dataUrlToBlob(r.dataUrl), name: `${M.p.KODE}_PROGRES_M${String(w).padStart(2, '0')}_${todayStr().replace(/-/g, '')}_${String(n).padStart(2, '0')}.jpg`, mime: 'image/jpeg', kategori: '06_Progres', proyekId: M.p.ID, jenis: 'PROGRES', terkait: hid, onDone: { type: 'appendFoto', table: 'PROGRES', id: hid } });
      } catch (er) { toast(er.message, 'err'); }
    }
    toast(`${files.length} foto masuk antrian upload — tersimpan otomatis ke Drive`, 'ok');
    const chips = $('#pgFoto', el).closest('.field').querySelector('.chips');
    if (chips) chips.insertAdjacentHTML('beforeend', badge(files.length + ' foto ditambahkan', 'warning'));
  };
  $('#pgSave', el).onclick = () => {
    let n = 0, bad = 0;
    $$('#pgT tr[data-pg], #pgT tr[data-pit]', el).forEach(tr => {
      const inpEl = $('[data-pv]', tr); if (!inpEl || inpEl.disabled || (tr.dataset.pit && tr.classList.contains('hidden'))) return;
      const id = tr.dataset.pg || tr.dataset.pit, v = parseNum(inpEl.value);
      const rid = `${M.p.ID}|${w}|${id}`;
      if (v === '') { if (getRow('PROGRES', rid)) save('PROGRES', { ID: rid, DELETED: 'YA' }, { silent: true }); return; }
      if (v < 0 || v > 100) { bad++; return; }
      save('PROGRES', { ID: rid, PROYEK_ID: M.p.ID, MINGGU_KE: w, ITEM_ID: id, PERSEN_AKTUAL: v, TANGGAL: todayStr() }, { silent: true }); n++;
    });
    const note = $('#pgNote', el).value.trim();
    if (note || head.ID) save('PROGRES', { ID: head.ID || `${M.p.ID}|${w}|_MINGGU`, PROYEK_ID: M.p.ID, MINGGU_KE: w, ITEM_ID: '_MINGGU', CATATAN: note, TANGGAL: todayStr() }, { silent: true });
    if (bad) toast(`${bad} nilai di luar 0–100 diabaikan`, 'err');
    scheduleSync(); toast(`Progres minggu ${w} disimpan (${n} baris)`, 'ok'); rerenderSoft();
  };
}
function bobotItem(M, it) { const c = M.rk.computed[it.ID]; return c && M.rk.C ? c.jumlah / M.rk.C : 0; }
UPLOAD_HOOKS.appendFoto = (d, r) => { const cur = getRow(d.table, d.id); const list = String((cur && cur.FOTO) || '').split('\n').filter(Boolean); list.push(r.url); save(d.table, { ID: d.id, FOTO: list.join('\n') }); };

/* ---------- Laporan mingguan + cash flow ---------- */
function weekReport(M, w) {
  const plan = M.planAt(w), act = M.actualCum(w), dev = act - plan, st = devStatus(dev);
  const wa = num(setting('DEVIASI_WASPADA')) || 5;
  const late = [], mulai = [], selesai = [];
  const ws = Calc.addDays(M.start, (w - 1) * 7), we = Calc.addDays(M.start, w * 7 - 1);
  M.groups.forEach(g => {
    const r = M.sched.tasks[g.row.ID]; if (!r || r.span) return;
    const pp = Calc.planPct(r, w, M.start, M.cal), ap = M.pctDiv(g, w);
    if (pp - ap >= wa) late.push({ g, pp, ap });
    if (r.mulai >= ws && r.mulai <= we) mulai.push(g);
    if (r.selesai >= ws && r.selesai <= we) selesai.push(g);
  });
  return { w, ws, we, plan, act, dev, st, late, mulai, selesai };
}
function cashflow(M) {
  const C = M.rk.C, tm = terminRows(M.p.ID).filter(r => num(r.MINGGU_KE) > 0);
  let kum = 0;
  return M.kurva.map(k => { const v = Math.round(k.rencana / 100 * C); kum += v; const up = tm.filter(r => num(r.MINGGU_KE) === k.minggu).reduce((a, r) => a + num(r.NOMINAL), 0); return { minggu: k.minggu, mulai: k.mulai, biaya: v, kum, upah: up }; });
}
function tabLaporan(M, el) {
  if (!M.scheduled) { el.innerHTML = '<div class="card empty"><p>Atur durasi divisi di tab Gantt dulu.</p></div>'; return; }
  const maxW = Math.max(M.kurva.length, M.weekNow, 1);
  let w = PWEEK[M.p.ID] || Math.min(Math.max(1, M.lastData || M.weekNow), maxW);
  const R = weekReport(M, w), head = progRows(M.p.ID).find(r => r.ITEM_ID === '_MINGGU' && num(r.MINGGU_KE) === w) || {};
  const cf = cashflow(M);
  el.innerHTML = `<div class="row" style="margin-bottom:12px"><div class="field"><label>Laporan minggu ke-</label><select class="input" id="lwSel" style="width:auto">${Array.from({ length: maxW }, (_, i) => i + 1).map(i => `<option value="${i}" ${i === w ? 'selected' : ''}>Minggu ${i}${i === M.weekNow ? ' (berjalan)' : ''}</option>`).join('')}</select></div><span class="spacer"></span><button class="btn primary" id="lwPdf">${ICON.pdf} PDF laporan</button><button class="btn ghost" id="lwWa">${ICON.share} WhatsApp</button></div>
    <div class="grid md2"><div class="card"><h3>Minggu ${w} · ${fmtDate(R.ws)} – ${fmtDate(R.we)}</h3><dl class="kv">
      <dt>Rencana kumulatif</dt><dd>${pct(R.plan, 2)}</dd><dt>Aktual kumulatif</dt><dd>${M.lastData >= w ? pct(R.act, 2) : '<span class="muted">belum diisi</span>'}</dd>
      <dt>Deviasi</dt><dd class="${R.dev < 0 ? 'neg' : ''}">${M.lastData >= w ? (R.dev > 0 ? '+' : '') + N(R.dev, 2) + '%' : '-'}</dd><dt>Status</dt><dd>${M.lastData >= w ? badge(R.st[0], R.st[1]) : '-'}</dd>
      <dt>Estimasi selesai</dt><dd>${estimasiSelesai(M) || '-'}</dd><dt>Rencana selesai</dt><dd>${fmtDateLong(M.sched.selesai)}</dd></dl>
      ${head.CATATAN ? `<div class="notice" style="margin-top:10px;white-space:pre-line">${esc(head.CATATAN)}</div>` : ''}
      ${String(head.FOTO || '').trim() ? `<div class="chips" style="margin-top:8px">${String(head.FOTO).split('\n').filter(Boolean).map((u, i) => `<a class="badge info" href="${esc(u)}" target="_blank" rel="noopener">Foto ${i + 1}</a>`).join('')}</div>` : ''}</div>
    <div class="card"><h3>Perlu perhatian</h3>
      ${R.late.length ? `<div class="list">${R.late.map(x => `<div class="li"><span class="ico" style="color:var(--danger)">${ICON.warn}</span><div class="main"><div class="t1">${esc(x.g.row.KODE || '')} ${esc(x.g.row.URAIAN)}</div><div class="t2">rencana ${pct(x.pp, 0)} · aktual ${pct(x.ap, 0)}</div></div></div>`).join('')}</div>` : '<div class="small muted">Tidak ada pekerjaan terlambat.</div>'}
      <div class="small" style="margin-top:10px"><strong>Mulai minggu ini:</strong> ${R.mulai.map(g => esc(g.row.KODE || g.row.URAIAN)).join(', ') || '-'}<br><strong>Selesai minggu ini:</strong> ${R.selesai.map(g => esc(g.row.KODE || g.row.URAIAN)).join(', ') || '-'}</div></div></div>
    <div class="card" style="margin-top:12px"><h3>Rencana cash flow</h3><div class="table-wrap"><table class="t"><thead><tr><th>Minggu</th><th class="num">Biaya rencana</th><th class="num">Kumulatif</th><th class="num">Termin upah</th></tr></thead><tbody>${cf.map(c => `<tr${c.minggu === w ? ' class="sub"' : ''}><td>M${c.minggu} <span class="small muted">${fmtDate(c.mulai)}</span></td><td class="num">${rpCell(c.biaya)}</td><td class="num">${rpCell(c.kum)}</td><td class="num">${c.upah ? rpCell(c.upah) : '-'}</td></tr>`).join('')}</tbody></table></div>
      <div class="small muted" style="margin-top:6px">Biaya rencana = bobot Kurva S × total biaya pekerjaan + SMKK (${rp(M.rk.C)}), sebelum biaya tak terduga & PPN. Termin upah diambil dari tab Tenaga (pola mingguan).</div></div>`;
  $('#lwSel', el).onchange = e => { PWEEK[M.p.ID] = +e.target.value; rerenderSoft(); };
  $('#lwPdf', el).onclick = () => exportLaporanPdf(M, w).catch(e => { console.error(e); toast('Gagal membuat PDF: ' + e.message, 'err'); });
  $('#lwWa', el).onclick = () => shareWa(`*Laporan Progres Minggu ${w}*\n${M.p.NAMA}\nPeriode ${fmtDate(R.ws)} – ${fmtDate(R.we)}\nRencana: ${pct(R.plan, 1)} · Aktual: ${pct(R.act, 1)} · Deviasi: ${(R.dev > 0 ? '+' : '') + N(R.dev, 1)}% (${R.st[0]})${R.late.length ? '\nTerlambat: ' + R.late.map(x => x.g.row.URAIAN).join(', ') : ''}${head.CATATAN ? '\nCatatan: ' + head.CATATAN : ''}`);
}

/* ---------- Libur & milestone ---------- */
function tabKalender(M, el) {
  const lib = rows('LIBUR').sort((a, b) => String(a.TANGGAL).localeCompare(String(b.TANGGAL)));
  const yr = new Set(lib.map(l => String(l.TANGGAL).slice(0, 4)));
  const auto = autoMilestones(M), man = msRows(M.p.ID).sort((a, b) => String(a.TANGGAL).localeCompare(String(b.TANGGAL)));
  el.innerHTML = `<div class="grid md2">
    <div class="card"><h3>Milestone</h3><div class="list">${auto.map(m => `<div class="li"><div class="main"><div class="t1">${esc(m.nama)} ${badge('otomatis', 'primary')}</div><div class="t2">${fmtDateLong(m.tanggal)}</div></div></div>`).join('')}
      ${man.map(m => `<div class="li"><div class="main"><div class="t1">${esc(m.NAMA)}</div><div class="t2">${fmtDateLong(m.TANGGAL)}</div></div><select class="input" data-mst="${m.ID}" style="width:130px">${['RENCANA', 'TERCAPAI', 'TERLAMBAT'].map(s => `<option ${norm(m.STATUS) === s ? 'selected' : ''}>${s}</option>`).join('')}</select><button class="btn sm ghost" data-mdel="${m.ID}">Hapus</button></div>`).join('')}</div>
      <div class="row" style="margin-top:10px"><input class="input" id="msN" placeholder="Nama milestone, mis. Pengecoran plat lt.2" style="flex:2;min-width:160px"><input class="input" type="date" id="msD" style="flex:1;min-width:140px"><button class="btn" id="msAdd">Tambah</button></div></div>
    <div class="card"><h3>Hari libur (${lib.length})</h3><div class="small muted" style="margin-bottom:8px">Hari libur tidak dihitung sebagai hari kerja. Data awal: libur nasional & cuti bersama SKB 3 Menteri ${[...yr].sort().join(', ') || '-'}. Hapus cuti bersama bila tukang tetap bekerja.</div>
      <div class="list" style="max-height:340px;overflow:auto">${lib.map(l => `<div class="li"><div class="main"><div class="t1">${fmtDateLong(l.TANGGAL)}</div><div class="t2">${esc(l.KETERANGAN || '')}</div></div><button class="btn sm ghost" data-ldel="${l.ID}">Hapus</button></div>`).join('') || '<div class="muted small">Belum ada.</div>'}</div>
      <div class="row" style="margin-top:10px"><input class="input" type="date" id="lbD" style="flex:1;min-width:140px"><input class="input" id="lbK" placeholder="Keterangan" style="flex:2;min-width:140px"><button class="btn" id="lbAdd">Tambah</button></div></div></div>`;
  $('#msAdd', el).onclick = () => { const n = $('#msN', el).value.trim(), d = $('#msD', el).value; if (!n || !d) return toast('Isi nama & tanggal', 'err'); save('MILESTONE', { ID: uid('MS'), PROYEK_ID: M.p.ID, NAMA: n, TANGGAL: d, STATUS: 'RENCANA' }); rerenderSoft(); };
  $$('[data-mst]', el).forEach(s => s.onchange = () => save('MILESTONE', { ID: s.dataset.mst, STATUS: s.value }));
  $$('[data-mdel]', el).forEach(b => b.onclick = () => { remove('MILESTONE', b.dataset.mdel); rerenderSoft(); });
  $('#lbAdd', el).onclick = () => { const d = $('#lbD', el).value; if (!d) return toast('Pilih tanggal', 'err'); save('LIBUR', { ID: 'L' + d.replace(/-/g, '') + uid().slice(-3), TANGGAL: d, KETERANGAN: $('#lbK', el).value.trim() }); rerenderSoft(); };
  $$('[data-ldel]', el).forEach(b => b.onclick = () => { remove('LIBUR', b.dataset.ldel); rerenderSoft(); });
}

/* ---------- Export ---------- */
async function exportJadwalXlsx(M) {
  const X = await needXlsx(), wb = X.utils.book_new();
  const A = [['Kode', 'Pekerjaan', 'Durasi (hari kerja)', 'Mulai', 'Selesai', 'Setelah', 'Relasi', 'Lag', 'Float', 'Kritis', 'Bobot %']];
  M.groups.forEach(g => {
    const r = M.sched.tasks[g.row.ID], j = M.J[g.row.ID] || {}, pg = M.groups.find(x => x.row.ID === j.PREDECESSOR);
    A.push([g.row.KODE || '', g.row.URAIAN, r ? r.dur : '', r ? fmtDate(r.mulai) : '', r ? fmtDate(r.selesai) : '', pg ? `${pg.row.KODE || ''} ${pg.row.URAIAN}` : '', j.JENIS_RELASI || '', num(j.LAG), r && !r.span ? r.float : '', r && r.critical ? 'YA' : '', Calc.round(M.bobot[g.row.ID] * 100, 3)]);
    g.items.forEach(it => { const s = M.itemSched[it.ID]; if (s) A.push(['', '   ' + it.URAIAN, s.dur, fmtDate(s.mulai), fmtDate(s.selesai), '', '', '', '', '', '']); });
  });
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet(A), 'Jadwal');
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['Minggu', 'Mulai', 'Selesai', 'Rencana %', 'Rencana kum %', 'Aktual kum %', 'Deviasi %']].concat(M.kurva.map(k => { const has = k.minggu <= M.lastData; const a = has ? M.actualCum(k.minggu) : ''; return [k.minggu, fmtDate(k.mulai), fmtDate(k.selesai), Calc.round(k.rencana, 3), Calc.round(k.kum, 3), has ? Calc.round(a, 3) : '', has ? Calc.round(a - k.kum, 3) : '']; }))), 'Kurva S');
  X.utils.book_append_sheet(wb, X.utils.aoa_to_sheet([['Minggu', 'Biaya rencana', 'Kumulatif', 'Termin upah']].concat(cashflow(M).map(c => [c.minggu, c.biaya, c.kum, c.upah]))), 'Cash flow');
  X.writeFile(wb, `${M.p.KODE}_JADWAL_${todayStr().replace(/-/g, '')}.xlsx`);
}
async function exportJadwalPdf(M) {
  if (!M.scheduled) return toast('Atur durasi divisi dulu', 'err');
  const prof = profilOf(M.p.PROFIL_ID), jsPDF = await needPdf();
  const doc = new jsPDF({ unit: 'mm', format: 'a3', orientation: 'landscape' });
  doc._texts = [];
  const logo = prof ? await imageData(prof.LOGO_FILE_ID) : '';
  const ctx = { doc, logo, prof: prof || {} };
  const PW = 420, PH = 297, M0 = 12;
  let y = pdfKop(ctx, 12, 420);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.text('JADWAL PELAKSANAAN (GANTT CHART)', PW / 2, y, { align: 'center' });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.text(`${M.p.NAMA} · ${fmtDateLong(M.sched.mulai)} – ${fmtDateLong(M.sched.selesai)} · ${M.sched.durasiKerja} hari kerja`, PW / 2, y + 5, { align: 'center' });
  y += 10;
  const LW = 95, TX = M0 + LW, TW = PW - M0 - TX, days = Calc.daysBetween(M.start, M.sched.selesai) + 1, sc = TW / days;
  const rowsList = []; M.groups.forEach(g => rowsList.push(g));
  const rh = Math.min(7, (PH - y - 30) / Math.max(1, rowsList.length + 1));
  // header bulan & minggu
  doc.setFontSize(7);
  for (let i = 0; i < days;) {
    const dt = Calc.parseYmd(Calc.addDays(M.start, i)), nm = Calc.ymd(new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth() + 1, 1)));
    const span = Math.min(days, Calc.daysBetween(M.start, nm)) - i;
    doc.setFillColor(...hexRgb(THEME.primary)); doc.rect(TX + i * sc, y, span * sc, 5, 'F');
    doc.setTextColor(255); if (span * sc > 12) doc.text(`${BULAN[dt.getUTCMonth()].slice(0, 3)} ${dt.getUTCFullYear()}`, TX + i * sc + 1.5, y + 3.5);
    i += span;
  }
  doc.setFillColor(...hexRgb(THEME.primary)); doc.rect(M0, y, LW, 5, 'F'); doc.text('Pekerjaan', M0 + 2, y + 3.5);
  y += 5; doc.setTextColor(...hexRgb(THEME.muted));
  for (let w = 0; w * 7 < days; w++) { doc.setDrawColor(...hexRgb(THEME.border)); doc.line(TX + w * 7 * sc, y, TX + w * 7 * sc, y + rh * rowsList.length + 4); if (7 * sc > 6) doc.text('M' + (w + 1), TX + w * 7 * sc + 0.8, y + 3); }
  y += 4;
  rowsList.forEach((g, i) => {
    const r = M.sched.tasks[g.row.ID], yy = y + i * rh;
    if (i % 2) { doc.setFillColor(...hexRgb(THEME.zebra)); doc.rect(M0, yy, PW - 2 * M0, rh, 'F'); }
    doc.setTextColor(...hexRgb(THEME.text)); doc.setFontSize(7.5);
    doc.text(doc.splitTextToSize(`${g.row.KODE || ''} ${g.row.URAIAN}`, LW - 22)[0], M0 + 1.5, yy + rh * 0.65);
    if (r) {
      doc.setTextColor(...hexRgb(THEME.muted)); doc.text(`${r.dur}h`, M0 + LW - 2, yy + rh * 0.65, { align: 'right' });
      const col = r.span ? THEME.accent : r.critical ? THEME.danger : THEME.primary;
      doc.setFillColor(...hexRgb(col)); doc.roundedRect(TX + Calc.daysBetween(M.start, r.mulai) * sc, yy + rh * 0.2, Math.max(0.8, (Calc.daysBetween(r.mulai, r.selesai) + 1) * sc), rh * 0.6, 0.6, 0.6, 'F');
      const pc = M.pctDiv(g, Math.max(1, M.lastData));
      if (pc > 0 && !r.span) { doc.setFillColor(...hexRgb(THEME.success)); doc.rect(TX + Calc.daysBetween(M.start, r.mulai) * sc, yy + rh * 0.7, (Calc.daysBetween(r.mulai, r.selesai) + 1) * sc * Math.min(100, pc) / 100, rh * 0.12, 'F'); }
    }
  });
  const yEnd = y + rowsList.length * rh;
  autoMilestones(M).concat(msRows(M.p.ID).map(r => ({ nama: r.NAMA, tanggal: r.TANGGAL }))).forEach(m => {
    if (!m.tanggal) return; const mx = TX + (Calc.daysBetween(M.start, m.tanggal) + 0.5) * sc;
    doc.setFillColor(...hexRgb(THEME.accent)); doc.triangle(mx, yEnd + 1, mx - 1.5, yEnd + 3.5, mx + 1.5, yEnd + 3.5, 'F');
    doc.setFontSize(6); doc.setTextColor(...hexRgb(THEME.text)); doc.text(m.nama, mx, yEnd + 6, { align: 'center' });
  });
  doc.setFontSize(7); doc.setTextColor(...hexRgb(THEME.muted));
  doc.text('Merah = jalur kritis · Biru = normal · Oranye = SMKK (sepanjang proyek) · Hijau tipis = progres aktual · Segitiga = milestone', M0, yEnd + 11);
  // halaman 2: Kurva S
  doc.addPage('a3', 'landscape');
  doc.setTextColor(...hexRgb(THEME.text)); doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.text('KURVA S — RENCANA vs AKTUAL', PW / 2, 18, { align: 'center' });
  const gx = 30, gy = 28, gw = 250, gh = 150, n = M.kurva.length;
  const X = i => gx + i / n * gw, Y = v => gy + (1 - v / 100) * gh;
  doc.setFont('helvetica', 'normal'); doc.setFontSize(7);
  [0, 25, 50, 75, 100].forEach(v => { doc.setDrawColor(...hexRgb(THEME.border)); doc.line(gx, Y(v), gx + gw, Y(v)); doc.setTextColor(...hexRgb(THEME.muted)); doc.text(v + '%', gx - 2, Y(v) + 1, { align: 'right' }); });
  M.kurva.forEach((k, i) => { if (n <= 30 || (i + 1) % Math.ceil(n / 30) === 0) doc.text('M' + k.minggu, X(i + 1), gy + gh + 5, { align: 'center' }); });
  doc.setLineWidth(0.7); doc.setDrawColor(...hexRgb(THEME.primary));
  let px = X(0), py = Y(0); M.kurva.forEach((k, i) => { doc.line(px, py, X(i + 1), Y(k.kum)); px = X(i + 1); py = Y(k.kum); });
  if (M.lastData) { doc.setDrawColor(...hexRgb(THEME.accent)); px = X(0); py = Y(0); for (let w = 1; w <= Math.min(n, M.lastData); w++) { const a = Math.min(100, M.actualCum(w)); doc.line(px, py, X(w), Y(a)); px = X(w); py = Y(a); } }
  doc.setLineWidth(0.2);
  doc.setTextColor(...hexRgb(THEME.text)); doc.setFontSize(8); doc.text('Biru = rencana · Oranye = aktual', gx, gy + gh + 11);
  doc.autoTable(AT({ startY: 28, margin: { left: 295, right: 12 }, styles: { fontSize: 7, cellPadding: 1 }, head: [['Mgg', 'Rencana', 'Kum', 'Aktual', 'Dev']], body: M.kurva.map(k => { const has = k.minggu <= M.lastData, a = has ? M.actualCum(k.minggu) : null; return [k.minggu, N(k.rencana, 2), N(k.kum, 2), has ? N(a, 2) : '-', has ? N(a - k.kum, 2) : '-']; }), columnStyles: { 1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' } } }));
  pdfFooterA3(ctx, `Jadwal ${M.p.KODE}`);
  await deliverPdf(doc, `${M.p.KODE}_JADWAL_${todayStr().replace(/-/g, '')}.pdf`, { proyekId: M.p.ID, kategori: '03_Jadwal', jenis: 'JADWAL', terkait: M.p.ID, shareText: `Jadwal ${M.p.NAMA}: ${fmtDateLong(M.sched.mulai)} – ${fmtDateLong(M.sched.selesai)}` });
}
function pdfFooterA3(ctx, docNo) {
  const { doc, prof } = ctx, n = doc.getNumberOfPages();
  for (let i = 1; i <= n; i++) { doc.setPage(i); doc.setFontSize(7.5); doc.setTextColor(...hexRgb(THEME.muted)); doc.text([prof.NAMA, docNo].filter(Boolean).join('  ·  '), 12, 290); doc.text(`Halaman ${i} dari ${n}`, 408, 290, { align: 'right' }); if (norm(setting('TAMPIL_CREDIT')) === 'YA') doc.text('dibuat dengan RAB Rumah & Gedung by DG', 210, 293, { align: 'center' }); }
}
async function exportLaporanPdf(M, w) {
  const R = weekReport(M, w), prof = profilOf(M.p.PROFIL_ID), ctx = await pdfStart(prof), doc = ctx.doc;
  const head = progRows(M.p.ID).find(r => r.ITEM_ID === '_MINGGU' && num(r.MINGGU_KE) === w) || {};
  let y = pdfKop(ctx);
  doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.text(`LAPORAN PROGRES MINGGU KE-${w}`, 105, y + 2, { align: 'center' });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.text(`${M.p.NAMA} · periode ${fmtDateLong(R.ws)} – ${fmtDateLong(R.we)}`, 105, y + 8, { align: 'center' });
  doc.autoTable(AT({ startY: y + 12, theme: 'plain', styles: { fontSize: 9.5, cellPadding: 1.2 }, columnStyles: { 0: { cellWidth: 55, textColor: hexRgb(THEME.muted) } }, body: [['Rencana kumulatif', pct(R.plan, 2)], ['Aktual kumulatif', pct(R.act, 2)], ['Deviasi', (R.dev > 0 ? '+' : '') + N(R.dev, 2) + '%'], ['Status', R.st[0]], ['Rencana selesai', fmtDateLong(M.sched.selesai)], ['Estimasi selesai (laju aktual)', estimasiSelesai(M) || '-']] }));
  doc.autoTable(AT({ startY: doc.lastAutoTable.finalY + 4, head: [['Pekerjaan', 'Bobot', 'Rencana', 'Aktual', 'Selisih']], body: M.groups.filter(g => M.bobot[g.row.ID] > 0).map(g => { const r = M.sched.tasks[g.row.ID], pp = r ? Calc.planPct(r, w, M.start, M.cal) : 0, ap = M.pctDiv(g, w); return [`${g.row.KODE || ''} ${g.row.URAIAN}`, pct(M.bobot[g.row.ID] * 100, 2), pct(pp, 0), pct(ap, 0), N(ap - pp, 0) + '%']; }), columnStyles: { 1: { halign: 'right' }, 2: { halign: 'right' }, 3: { halign: 'right' }, 4: { halign: 'right' } } }));
  let yy = doc.lastAutoTable.finalY + 6;
  const sect = (t, txtv) => { if (!txtv) return; if (yy > 260) { doc.addPage(); yy = 20; } doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.text(t, 15, yy); doc.setFont('helvetica', 'normal'); doc.setFontSize(9); const l = doc.splitTextToSize(txtv, 180); doc.text(l, 15, yy + 5); yy += 7 + l.length * 4.2; };
  sect('Pekerjaan terlambat', R.late.map(x => `• ${x.g.row.URAIAN} (rencana ${N(x.pp, 0)}%, aktual ${N(x.ap, 0)}%)`).join('\n') || 'Tidak ada.');
  sect('Mulai / selesai minggu ini', `Mulai: ${R.mulai.map(g => g.row.URAIAN).join(', ') || '-'}\nSelesai: ${R.selesai.map(g => g.row.URAIAN).join(', ') || '-'}`);
  sect('Catatan lapangan', head.CATATAN || '');
  const fotos = String(head.FOTO || '').split('\n').filter(Boolean);
  if (fotos.length) sect('Foto dokumentasi (Google Drive)', fotos.map((u, i) => `Foto ${i + 1}: ${u}`).join('\n'));
  signBlock(doc, Math.max(yy + 6, 220), [{ title: 'Dibuat oleh', name: prof && prof.PJ_NAMA, role: prof && prof.PJ_JABATAN }, { title: 'Diperiksa oleh' }, { title: 'Mengetahui', name: M.p.PEMILIK }]);
  pdfFooter(ctx, `Progres M${w} ${M.p.KODE}`, norm(setting('TAMPIL_CREDIT')) === 'YA');
  await deliverPdf(doc, `${M.p.KODE}_PROGRES_M${String(w).padStart(2, '0')}_${todayStr().replace(/-/g, '')}.pdf`, { proyekId: M.p.ID, kategori: '06_Progres', jenis: 'LAPORAN_PROGRES', terkait: M.p.ID, shareText: `Laporan progres minggu ${w} ${M.p.NAMA}: aktual ${pct(R.act, 1)} vs rencana ${pct(R.plan, 1)} (${R.st[0]})` });
}
/* ============================================================
 * Start aplikasi
 * ============================================================ */
async function boot() {
  try {
    await DB.open();
    const meta = await DB.get('meta', 'meta'); if (meta) Object.assign(S.meta, meta);
    S.queue = (await DB.get('meta', 'queue')) || {};
    S.uploads = (await DB.get('meta', 'uploads')) || [];
    for (const t of TABLES) { const v = await DB.get('tables', t); if (Array.isArray(v)) S.t[t] = v; }
  } catch (e) { console.error(e); toast('Penyimpanan lokal tidak tersedia (mode privat?). Data hanya tersimpan selama halaman terbuka.', 'err'); }
  window.addEventListener('hashchange', render);
  window.addEventListener('online', () => { S.online = true; renderSync(); sync(); });
  window.addEventListener('offline', () => { S.online = false; renderSync(); });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') sync(); });
  $('#syncChip').onclick = () => { if (!S.meta.apiUrl) location.hash = '#/master/pengaturan'; else sync({ quiet: false }); };
  setInterval(() => { if (document.visibilityState === 'visible') sync(); }, 60000);
  render(); renderSync();
  if (S.meta.apiUrl && S.meta.token) sync({ full: !S.meta.lastSync });
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').then(reg => {
      reg.addEventListener('updatefound', () => {
        const nw = reg.installing;
        nw && nw.addEventListener('statechange', () => { if (nw.state === 'installed' && navigator.serviceWorker.controller) toast('Versi baru tersedia — tutup & buka ulang aplikasi untuk memperbarui.'); });
      });
    }).catch(() => { });
  }
}
document.addEventListener('DOMContentLoaded', boot);
