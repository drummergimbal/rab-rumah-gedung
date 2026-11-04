/* ============================================================
 * RAB Rumah & Gedung — Mesin Hitung (by DG)
 * File ini dipakai DUA tempat dengan isi yang SAMA PERSIS:
 *   1. PWA  : calc.js
 *   2. Apps Script : Calc.gs  (copy-paste seluruh isi file ini)
 * Jangan ubah di satu tempat saja, supaya hasil hitung identik.
 * ============================================================ */
var Calc = (function () {
  'use strict';

  function num(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = Number(String(v).replace(',', '.'));
    return isFinite(n) ? n : 0;
  }
  function round(v, d) {
    var f = Math.pow(10, d || 0);
    return Math.round((num(v) + Number.EPSILON) * f) / f;
  }
  function isDel(r) { return r && String(r.DELETED || '').toUpperCase() === 'YA'; }
  function norm(s) { return String(s || '').trim().toUpperCase(); }

  /* ---------- Index data master ---------- */
  function buildIndex(tables) {
    var idx = { hsd: {}, ahsp: {}, komp: {}, usage: {} };
    ['HSD_UPAH', 'HSD_BAHAN', 'HSD_ALAT'].forEach(function (t) {
      (tables[t] || []).forEach(function (r) {
        if (isDel(r) || !r.KODE_HSD) return;
        var k = norm(r.KODE_HSD);
        (idx.hsd[k] = idx.hsd[k] || []).push(r);
      });
    });
    (tables.AHSP_HEADER || []).forEach(function (r) {
      if (isDel(r) || !r.KODE_AHSP) return;
      idx.ahsp[norm(r.KODE_AHSP)] = r;
    });
    (tables.AHSP_KOMPONEN || []).forEach(function (r) {
      if (isDel(r) || !r.KODE_AHSP) return;
      var k = norm(r.KODE_AHSP);
      (idx.komp[k] = idx.komp[k] || []).push(r);
    });
    Object.keys(idx.komp).forEach(function (k) {
      idx.komp[k].sort(function (a, b) { return num(a.URUTAN) - num(b.URUTAN); });
    });
    (tables.RAB_ITEM || []).forEach(function (r) {
      if (isDel(r) || !r.KODE_AHSP) return;
      var k = norm(r.KODE_AHSP);
      idx.usage[k] = (idx.usage[k] || 0) + 1;
    });
    return idx;
  }

  /* ---------- Harga dasar per wilayah & tahun ----------
   * Prioritas: wilayah & tahun sama → wilayah sama tahun terdekat (<= tahun, lalu terbaru)
   * → wilayah lain tahun terbaru (ditandai fallback). Harga 0/kosong = belum ada harga. */
  function hsdPrice(idx, kode, wilayah, tahun) {
    var rows = (idx.hsd[norm(kode)] || []).filter(function (r) { return num(r.HARGA) > 0; });
    if (!rows.length) return { harga: 0, row: (idx.hsd[norm(kode)] || [])[0] || null, status: 'KOSONG' };
    var w = norm(wilayah), t = num(tahun);
    var same = rows.filter(function (r) { return norm(r.WILAYAH) === w; });
    var pick = null, status = 'OK';
    if (same.length) {
      var exact = same.filter(function (r) { return num(r.TAHUN) === t; });
      if (exact.length) pick = exact[0];
      else {
        var older = same.filter(function (r) { return num(r.TAHUN) <= t; })
          .sort(function (a, b) { return num(b.TAHUN) - num(a.TAHUN); });
        pick = older[0] || same.sort(function (a, b) { return num(b.TAHUN) - num(a.TAHUN); })[0];
        status = 'TAHUN_LAIN';
      }
    } else {
      pick = rows.slice().sort(function (a, b) { return num(b.TAHUN) - num(a.TAHUN); })[0];
      status = 'WILAYAH_LAIN';
    }
    return { harga: num(pick.HARGA), row: pick, status: status };
  }

  /* ---------- Analisa AHSP ----------
   * E = Σ(koef × HSD) per Upah/Bahan/Alat ; F = E × overhead% ; HSP = E + F */
  function analyze(idx, kodeAhsp, ctx) {
    var k = norm(kodeAhsp);
    var head = idx.ahsp[k];
    var komp = idx.komp[k] || [];
    var res = {
      kode: kodeAhsp, header: head || null, ok: !!head, rows: [],
      upah: 0, bahan: 0, alat: 0, E: 0, F: 0, HSP: 0, missing: [], fallback: []
    };
    if (!head) return res;
    komp.forEach(function (c) {
      var p = hsdPrice(idx, c.KODE_HSD, ctx.wilayah, ctx.tahun);
      var koef = num(c.KOEFISIEN);
      var harga = koef * p.harga;
      var jenis = norm(c.JENIS_KOMPONEN);
      if (jenis === 'UPAH') res.upah += harga;
      else if (jenis === 'ALAT') res.alat += harga;
      else res.bahan += harga;
      if (p.status === 'KOSONG') res.missing.push(c.KODE_HSD);
      else if (p.status !== 'OK') res.fallback.push(c.KODE_HSD);
      res.rows.push({
        jenis: jenis, kodeHsd: c.KODE_HSD,
        nama: c.NAMA_KOMPONEN || (p.row && p.row.NAMA) || c.KODE_HSD,
        satuan: c.SATUAN || (p.row && p.row.SATUAN) || '',
        koef: koef, hargaDasar: p.harga, jumlah: harga, status: p.status
      });
    });
    res.E = res.upah + res.bahan + res.alat;
    res.F = res.E * num(ctx.overhead) / 100;
    res.HSP = round(res.E + res.F, 0);
    return res;
  }

  /* ---------- Hitung satu item RAB ---------- */
  function itemCalc(idx, item, ctx) {
    var vol = num(item.VOLUME);
    var out = { hs: 0, jumlah: 0, upah: 0, bahan: 0, alat: 0, overhead: 0, lain: 0, analisa: null, warn: [] };
    if (norm(item.JENIS_HARGA) === 'PASAR' || !item.KODE_AHSP) {
      out.hs = round(item.HARGA_PASAR, 0);
      out.jumlah = round(out.hs * vol, 0);
      out.lain = out.jumlah;
      if (!item.KODE_AHSP && norm(item.JENIS_HARGA) !== 'PASAR') out.warn.push('TANPA_AHSP');
      if (out.hs <= 0) out.warn.push('HARGA_KOSONG');
      return out;
    }
    var a = analyze(idx, item.KODE_AHSP, ctx);
    out.analisa = a;
    if (!a.ok) { out.warn.push('AHSP_TIDAK_ADA'); return out; }
    if (a.missing.length) out.warn.push('HSD_KOSONG');
    if (a.fallback.length) out.warn.push('HSD_FALLBACK');
    if (a.header && norm(a.header.STATUS) === 'PERLU CEK') out.warn.push('PERLU_CEK');
    out.hs = a.HSP;
    out.jumlah = round(out.hs * vol, 0);
    out.upah = a.upah * vol; out.bahan = a.bahan * vol; out.alat = a.alat * vol;
    out.overhead = out.jumlah - (out.upah + out.bahan + out.alat);
    return out;
  }

  function ctxFromProject(p, settings) {
    settings = settings || {};
    return {
      wilayah: p.WILAYAH || settings.WILAYAH_DEFAULT || '',
      tahun: num(p.TAHUN_ANGGARAN) || new Date().getFullYear(),
      overhead: p.OVERHEAD_PCT !== '' && p.OVERHEAD_PCT !== undefined && p.OVERHEAD_PCT !== null
        ? num(p.OVERHEAD_PCT) : num(settings.OVERHEAD_PCT || 10)
    };
  }

  /* ---------- Pembulatan ke kelipatan ---------- */
  function roundTo(v, step, mode) {
    step = num(step);
    if (!step || step <= 1) return round(v, 0);
    var x = num(v) / step;
    if (mode === 'atas') x = Math.ceil(x - 1e-9);
    else if (mode === 'bawah') x = Math.floor(x + 1e-9);
    else x = Math.round(x);
    return x * step;
  }

  /* ---------- Rekap proyek ---------- */
  function sortRows(rows) {
    return rows.slice().sort(function (a, b) { return num(a.URUTAN) - num(b.URUTAN); });
  }
  function rekap(idx, project, rows, settings) {
    settings = settings || {};
    var ctx = ctxFromProject(project, settings);
    var all = sortRows(rows.filter(function (r) { return !isDel(r) && r.PROYEK_ID === project.ID; }));
    var divs = all.filter(function (r) { return norm(r.TIPE_BARIS) === 'DIVISI'; });
    var items = all.filter(function (r) { return norm(r.TIPE_BARIS) !== 'DIVISI'; });
    var byDiv = {};
    divs.forEach(function (d) {
      byDiv[d.ID] = { row: d, items: [], subtotal: 0, upah: 0, bahan: 0, alat: 0, overhead: 0, lain: 0, smkk: isSmkk(d) };
    });
    var orphan = { row: { ID: '_', URAIAN: 'Tanpa Divisi', KODE: '-' }, items: [], subtotal: 0, upah: 0, bahan: 0, alat: 0, overhead: 0, lain: 0, smkk: false };
    var computed = {}, warnings = [];
    items.forEach(function (it) {
      var c = itemCalc(idx, it, ctx);
      computed[it.ID] = c;
      var g = byDiv[it.PARENT_ID] || orphan;
      g.items.push(it);
      g.subtotal += c.jumlah; g.upah += c.upah; g.bahan += c.bahan; g.alat += c.alat;
      g.overhead += c.overhead; g.lain += c.lain;
      c.warn.forEach(function (w) { warnings.push({ item: it, code: w }); });
    });
    var groups = divs.map(function (d) { return byDiv[d.ID]; });
    if (orphan.items.length) groups.push(orphan);
    var A = 0, B = 0, sum = { upah: 0, bahan: 0, alat: 0, overhead: 0, lain: 0 };
    groups.forEach(function (g) {
      if (g.smkk) B += g.subtotal; else A += g.subtotal;
      ['upah', 'bahan', 'alat', 'overhead', 'lain'].forEach(function (k) { sum[k] += g[k]; });
    });
    var C = A + B;
    groups.forEach(function (g) { g.bobot = C ? g.subtotal / C * 100 : 0; });
    items.forEach(function (it) { computed[it.ID].bobot = C ? computed[it.ID].jumlah / C * 100 : 0; });
    var kontPct = project.KONTINJENSI_PCT !== '' && project.KONTINJENSI_PCT !== undefined
      ? num(project.KONTINJENSI_PCT) : num(settings.KONTINJENSI_PCT);
    var ppnPct = project.PPN_PCT !== '' && project.PPN_PCT !== undefined
      ? num(project.PPN_PCT) : num(settings.PPN_PCT);
    var D = round(C * kontPct / 100, 0);
    var E = C + D;
    var ppn = round(E * ppnPct / 100, 0);
    var F = E + ppn;
    var step = num(project.PEMBULATAN || settings.PEMBULATAN || 0);
    var total = roundTo(F, step, 'atas');
    var luas = num(project.LUAS_BANGUNAN);
    return {
      ctx: ctx, groups: groups, computed: computed, warnings: warnings,
      A: A, B: B, C: C, kontPct: kontPct, D: D, E: E, ppnPct: ppnPct, ppn: ppn, F: F,
      pembulatan: step, total: total, selisihBulat: total - F,
      perM2: luas ? round(total / luas, 0) : 0, luas: luas, split: sum,
      itemCount: items.length
    };
  }
  function isSmkk(divRow) {
    var s = norm(divRow.URAIAN) + ' ' + norm(divRow.KODE_DIVISI || '');
    return s.indexOf('SMKK') >= 0 || s.indexOf('KESELAMATAN KONSTRUKSI') >= 0;
  }

  /* ---------- Volume dari rincian ---------- */
  function volumeFromDetails(details) {
    var total = 0;
    (details || []).forEach(function (d) {
      if (isDel(d)) return;
      var dims = [d.PANJANG, d.LEBAR, d.TINGGI, d.JUMLAH].filter(function (x) { return x !== '' && x !== null && x !== undefined; });
      if (!dims.length) return;
      var h = dims.reduce(function (a, b) { return a * num(b); }, 1);
      total += norm(d.PENGURANG) === 'YA' ? -h : h;
    });
    return round(total, 4);
  }

  /* ---------- Material & tenaga take-off ---------- */
  function takeoff(idx, rk) {
    var map = {};
    Object.keys(rk.computed).forEach(function (id) {
      var c = rk.computed[id];
      if (!c.analisa || !c.analisa.ok) return;
      var it = null;
      rk.groups.forEach(function (g) { g.items.forEach(function (x) { if (x.ID === id) it = x; }); });
      var vol = num(it && it.VOLUME);
      c.analisa.rows.forEach(function (r) {
        var k = norm(r.kodeHsd) + '|' + r.jenis;
        var m = map[k] || (map[k] = { kode: r.kodeHsd, nama: r.nama, satuan: r.satuan, jenis: r.jenis, qty: 0, harga: r.hargaDasar, jumlah: 0 });
        m.qty += r.koef * vol;
        m.jumlah += r.jumlah * vol;
      });
    });
    return Object.keys(map).map(function (k) { return map[k]; })
      .sort(function (a, b) { return a.jenis === b.jenis ? b.jumlah - a.jumlah : (a.jenis < b.jenis ? -1 : 1); });
  }

  /* ---------- Penawaran (harga klien) ----------
   * opt: { markupPct, markupDivisi:{divId:pct}, divisiSetting:{divId:{mode:'TAMPIL'|'SEMBUNYI'|'GABUNG', ke:divId}},
   *        bulatSatuan, bulatTotal, diskonTipe:'PERSEN'|'NOMINAL', diskonNilai, ppnMode:'BELUM'|'TERMASUK'|'TIDAK', ppnPct } */
  function penawaran(rk, opt) {
    opt = opt || {};
    var ds = opt.divisiSetting || {};
    var target = function (divId) {
      var s = ds[divId];
      if (s && s.mode === 'GABUNG' && s.ke && s.ke !== divId) return target(s.ke);
      return divId;
    };
    var out = [], gmap = {};
    rk.groups.forEach(function (g) {
      var id = g.row.ID, s = ds[id] || {};
      if (s.mode === 'SEMBUNYI') return;
      var tid = target(id);
      var tg = gmap[tid];
      if (!tg) {
        var trow = null;
        rk.groups.forEach(function (x) { if (x.row.ID === tid) trow = x.row; });
        tg = gmap[tid] = { divId: tid, uraian: (trow || g.row).URAIAN, kode: (trow || g.row).KODE, items: [], subtotal: 0, internal: 0 };
        out.push(tg);
      }
      var mk = opt.markupDivisi && opt.markupDivisi[id] !== undefined && opt.markupDivisi[id] !== ''
        ? num(opt.markupDivisi[id]) : num(opt.markupPct);
      g.items.forEach(function (it) {
        var c = rk.computed[it.ID];
        if (!(num(it.VOLUME) > 0) || !(c.jumlah > 0)) return; // item kosong tidak masuk penawaran
        var hs = roundTo(c.hs * (1 + mk / 100), opt.bulatSatuan, 'atas');
        var jml = round(hs * num(it.VOLUME), 0);
        tg.items.push({ itemId: it.ID, uraian: it.URAIAN, satuan: it.SATUAN, volume: num(it.VOLUME), hs: hs, jumlah: jml, spesifikasi: it.SPESIFIKASI || '' });
        tg.subtotal += jml;
        tg.internal += c.jumlah;
      });
    });
    // urut sesuai urutan divisi asli
    var order = {};
    rk.groups.forEach(function (g, i) { order[g.row.ID] = i; });
    out.sort(function (a, b) { return (order[a.divId] || 0) - (order[b.divId] || 0); });
    out = out.filter(function (g) { return g.items.length; });
    var subtotal = 0, internal = 0;
    out.forEach(function (g) { subtotal += g.subtotal; internal += g.internal; });
    var diskon = opt.diskonTipe === 'PERSEN' ? round(subtotal * num(opt.diskonNilai) / 100, 0) : round(opt.diskonNilai, 0);
    var dpp = subtotal - diskon;
    var ppnPct = num(opt.ppnPct), ppn = 0, total = dpp, ppnInfo = 0;
    if (opt.ppnMode === 'BELUM') { ppn = round(dpp * ppnPct / 100, 0); total = dpp + ppn; }
    else if (opt.ppnMode === 'TERMASUK') { ppnInfo = round(dpp * ppnPct / (100 + ppnPct), 0); }
    var totalBulat = roundTo(total, opt.bulatTotal, 'atas');
    var dppNet = opt.ppnMode === 'TERMASUK' ? totalBulat - ppnInfo : totalBulat - ppn;
    return {
      groups: out, subtotal: subtotal, diskon: diskon, dpp: dpp, ppnMode: opt.ppnMode, ppnPct: ppnPct,
      ppn: ppn, ppnInfo: ppnInfo, total: totalBulat, selisihBulat: totalBulat - total,
      internal: internal, margin: dppNet - internal,
      marginPct: internal ? (dppNet - internal) / internal * 100 : 0
    };
  }

  /* ---------- Terbilang ---------- */
  var SAT = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas'];
  function tb(n) {
    n = Math.floor(n);
    if (n < 12) return SAT[n];
    if (n < 20) return tb(n - 10) + ' belas';
    if (n < 100) return tb(Math.floor(n / 10)) + ' puluh' + (n % 10 ? ' ' + tb(n % 10) : '');
    if (n < 200) return 'seratus' + (n - 100 ? ' ' + tb(n - 100) : '');
    if (n < 1000) return tb(Math.floor(n / 100)) + ' ratus' + (n % 100 ? ' ' + tb(n % 100) : '');
    if (n < 2000) return 'seribu' + (n - 1000 ? ' ' + tb(n - 1000) : '');
    var units = [[1e12, ' triliun'], [1e9, ' miliar'], [1e6, ' juta'], [1e3, ' ribu']];
    for (var i = 0; i < units.length; i++) {
      if (n >= units[i][0]) {
        var q = Math.floor(n / units[i][0]), r = n % units[i][0];
        return tb(q) + units[i][1] + (r ? ' ' + tb(r) : '');
      }
    }
    return '';
  }
  function terbilang(n) {
    n = Math.round(num(n));
    if (n === 0) return 'Nol Rupiah';
    var s = (n < 0 ? 'minus ' : '') + tb(Math.abs(n)) + ' rupiah';
    s = s.replace(/\s+/g, ' ').trim();
    return s.replace(/(^|\s)(\S)/g, function (m, a, b) { return a + b.toUpperCase(); });
  }

  /* ---------- Pencarian referensi AHSP (offline, tanpa AI) ---------- */
  var DEFAULT_SINONIM = [
    'hebel = bata ringan = aac = celcon', 'keramik = ubin = tegel', 'granit = granite tile = homogeneous tile',
    'plester = plesteran', 'aci = acian = plamir semen', 'galian = gali = menggali', 'urugan = urug = timbunan',
    'pipa = paralon = pvc', 'cat = pengecatan = painting', 'bekisting = cetakan = begisting',
    'besi = pembesian = tulangan', 'beton = cor = pengecoran', 'kusen = kosen = frame', 'plafon = plafond = langit langit',
    'genteng = genting', 'roster = bata hias = loster', 'waterproofing = waterproof = pelapis kedap air',
    'titik lampu = instalasi lampu', 'stop kontak = stopkontak = kotak kontak', 'saklar = sakelar = switch',
    'kloset = closet = wc', 'wastafel = washtafel = lavatory', 'septic tank = septictank = tangki septik'
  ];
  function tokens(s) {
    return String(s || '').toLowerCase().replace(/[^a-z0-9À-ɏ\s.,x/]/g, ' ')
      .replace(/[.,/]/g, ' ').split(/\s+/).filter(function (t) { return t.length > 1; });
  }
  function synonymGroups(idx, extra) {
    var groups = [];
    DEFAULT_SINONIM.concat(extra || []).forEach(function (g) { groups.push(g); });
    Object.keys(idx.ahsp).forEach(function (k) {
      String(idx.ahsp[k].SINONIM || '').split(';').forEach(function (g) { if (g.indexOf('=') > 0) groups.push(g); });
    });
    return groups.map(function (g) {
      return g.split('=').map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean);
    });
  }
  function expandQuery(q, groups) {
    var base = tokens(q), lower = String(q || '').toLowerCase(), set = {};
    base.forEach(function (t) { set[t] = 1; });
    groups.forEach(function (g) {
      var hit = g.some(function (term) {
        return term.indexOf(' ') > 0 ? lower.indexOf(term) >= 0 : base.indexOf(term) >= 0;
      });
      if (hit) g.forEach(function (term) { tokens(term).forEach(function (t) { if (!set[t]) set[t] = 0.8; }); });
    });
    return set;
  }
  function search(idx, query, opt) {
    opt = opt || {};
    var groups = synonymGroups(idx, opt.sinonim);
    var q = expandQuery(query, groups);
    var qkeys = Object.keys(q);
    if (!qkeys.length) return [];
    var res = [];
    Object.keys(idx.ahsp).forEach(function (k) {
      var h = idx.ahsp[k];
      var ur = tokens(h.URAIAN), tg = tokens((h.TAG || '') + ' ' + (h.SINONIM || '') + ' ' + (h.DIVISI || ''));
      var score = 0, matched = [];
      qkeys.forEach(function (t) {
        var w = q[t];
        var inU = ur.some(function (u) { return u === t || (t.length > 3 && u.indexOf(t) === 0); });
        var inT = tg.some(function (u) { return u === t; });
        if (inU) { score += 2 * w; matched.push(t); }
        else if (inT) { score += 1.5 * w; matched.push(t); }
      });
      if (!score) return;
      var reasons = [];
      if (matched.length) reasons.push('kata kunci: ' + matched.slice(0, 4).join(', '));
      if (opt.divisi && norm(h.DIVISI) && norm(opt.divisi).indexOf(norm(h.DIVISI)) >= 0) { score += 2; reasons.push('divisi sama'); }
      if (opt.satuan && norm(h.SATUAN) === norm(opt.satuan)) { score += 1.5; reasons.push('satuan sama'); }
      var u = idx.usage[k] || 0;
      if (u) { score += 0.5 * Math.log(1 + u); reasons.push('dipakai ' + u + '×'); }
      if (norm(h.JENIS) === 'RESMI') score += 0.3;
      res.push({ kode: h.KODE_AHSP, header: h, score: score, reasons: reasons });
    });
    res.sort(function (a, b) { return b.score - a.score; });
    return res.slice(0, opt.limit || 10);
  }
  /* AHSP sejenis: divisi & satuan sama, untuk pembanding & validasi ±% */
  function similar(idx, kodeAhsp, ctx, limit) {
    var h = idx.ahsp[norm(kodeAhsp)];
    if (!h) return { list: [], avg: 0 };
    var list = Object.keys(idx.ahsp).map(function (k) { return idx.ahsp[k]; }).filter(function (x) {
      return norm(x.KODE_AHSP) !== norm(kodeAhsp) && norm(x.DIVISI) === norm(h.DIVISI) && norm(x.SATUAN) === norm(h.SATUAN);
    }).map(function (x) { return { header: x, HSP: analyze(idx, x.KODE_AHSP, ctx).HSP }; })
      .filter(function (x) { return x.HSP > 0; });
    var avg = list.length ? list.reduce(function (a, b) { return a + b.HSP; }, 0) / list.length : 0;
    return { list: list.slice(0, limit || 3), avg: avg };
  }

  /* ---------- Tenaga kerja (Tahap 2) ----------
   * OH per jenis tenaga = koefisien upah × volume.
   * plan: { divId: { durasi: hari kerja, tukang: jumlah tukang tersedia } }
   * opt : { jamKerja: 8, hariMinggu: 6, rasioKepala: 10, rasioMandor: 20 }
   * 1 OH = 8 jam kerja; bila jam kerja/hari ≠ 8, hari kerja = OH × 8 / jamKerja. */
  function roleClass(nama) {
    var n = String(nama || '').toLowerCase();
    if (/mandor|pengawas/.test(n)) return 'MANDOR';
    if (/kepala\s*tukang|kep\.?\s*tukang/.test(n)) return 'KEPALA';
    if (/tukang|operator|montir|instalatir|teknisi|juru|welder|las/.test(n)) return 'TUKANG';
    if (/pekerja|buruh|kenek|helper|laden/.test(n)) return 'PEKERJA';
    return 'LAIN';
  }
  function tenaga(idx, rk, plan, opt) {
    plan = plan || {}; opt = opt || {};
    var jam = num(opt.jamKerja) || 8, hm = num(opt.hariMinggu) || 6;
    var rK = num(opt.rasioKepala) || 10, rM = num(opt.rasioMandor) || 20;
    var f = 8 / jam;
    var roles = {}, divs = [], totalDurasi = 0, totalOH = 0, totalUpah = 0;
    rk.groups.forEach(function (g) {
      var d = { id: g.row.ID, kode: g.row.KODE || '', uraian: g.row.URAIAN, oh: {}, upah: 0, ohTotal: 0, ohTukang: 0, items: [], warn: [] };
      g.items.forEach(function (it) {
        var c = rk.computed[it.ID];
        if (!c || !c.analisa || !c.analisa.ok) return;
        var vol = num(it.VOLUME), ohItem = 0;
        c.analisa.rows.forEach(function (r) {
          if (r.jenis !== 'UPAH') return;
          var oh = r.koef * vol;
          if (!oh) return;
          var k = norm(r.kodeHsd);
          var R = roles[k] || (roles[k] = { kode: r.kodeHsd, nama: r.nama, kelas: roleClass(r.nama), harga: r.hargaDasar, oh: 0, upah: 0, puncak: 0 });
          R.oh += oh; R.upah += r.jumlah * vol;
          d.oh[k] = (d.oh[k] || 0) + oh;
          d.upah += r.jumlah * vol; d.ohTotal += oh; ohItem += oh;
          if (R.kelas === 'TUKANG') d.ohTukang += oh;
        });
        if (ohItem) d.items.push({ id: it.ID, uraian: it.URAIAN, oh: ohItem });
      });
      if (!d.ohTotal) { d.durasi = num((plan[d.id] || {}).durasi); d.orang = {}; d.orangTotal = 0; divs.push(d); return; }
      var p = plan[d.id] || {};
      var tk = num(p.tukang), du = num(p.durasi);
      d.mode = tk > 0 ? 'TUKANG' : 'DURASI';
      if (tk > 0) du = Math.ceil((d.ohTukang || d.ohTotal) * f / tk - 1e-9);
      d.durasi = du;
      d.orang = {}; d.orangTotal = 0;
      var cnt = { MANDOR: 0, KEPALA: 0, TUKANG: 0, PEKERJA: 0, LAIN: 0 };
      Object.keys(d.oh).forEach(function (k) {
        var o = du > 0 ? Math.ceil(d.oh[k] * f / du - 1e-9) : 0;
        if (tk > 0 && roles[k].kelas === 'TUKANG' && d.ohTukang) o = Math.max(1, Math.round(tk * d.oh[k] / d.ohTukang));
        d.orang[k] = o; d.orangTotal += o; cnt[roles[k].kelas] += o;
        if (o > roles[k].puncak) roles[k].puncak = o;
      });
      d.cnt = cnt;
      if (!du) d.warn.push('DURASI_KOSONG');
      if (cnt.KEPALA && cnt.TUKANG / cnt.KEPALA > rK) d.warn.push('RASIO_KEPALA');
      if (cnt.MANDOR && (cnt.PEKERJA + cnt.TUKANG) / cnt.MANDOR > rM) d.warn.push('RASIO_MANDOR');
      totalDurasi += du; totalOH += d.ohTotal; totalUpah += d.upah;
      divs.push(d);
    });
    // histogram mingguan (perkiraan: divisi dikerjakan berurutan)
    var weeks = [], day = 0;
    divs.forEach(function (d) {
      if (!d.durasi || !d.ohTotal) return;
      for (var i = 0; i < d.durasi; i++) {
        var w = Math.floor((day + i) / hm);
        var W = weeks[w] || (weeks[w] = { minggu: w + 1, orang: 0, oh: 0, divisi: {} });
        W.orang = Math.max(W.orang, d.orangTotal);
        W.oh += d.ohTotal / d.durasi;
        W.divisi[d.kode || d.uraian] = 1;
      }
      day += d.durasi;
    });
    for (var w = 0; w < weeks.length; w++) if (!weeks[w]) weeks[w] = { minggu: w + 1, orang: 0, oh: 0, divisi: {} };
    var list = Object.keys(roles).map(function (k) { return roles[k]; });
    var order = { MANDOR: 4, KEPALA: 3, TUKANG: 2, PEKERJA: 1, LAIN: 0 };
    list.sort(function (a, b) { return order[b.kelas] - order[a.kelas] || b.oh - a.oh; });
    return { roles: list, divisi: divs, totalDurasi: totalDurasi, totalOH: totalOH, totalUpah: totalUpah, weeks: weeks, minggu: Math.ceil(totalDurasi / hm), faktorJam: f };
  }

  return {
    roleClass: roleClass, tenaga: tenaga,
    num: num, round: round, isDel: isDel, norm: norm, buildIndex: buildIndex, hsdPrice: hsdPrice,
    analyze: analyze, itemCalc: itemCalc, ctxFromProject: ctxFromProject, roundTo: roundTo,
    rekap: rekap, isSmkk: isSmkk, volumeFromDetails: volumeFromDetails, takeoff: takeoff,
    penawaran: penawaran, terbilang: terbilang, search: search, similar: similar, tokens: tokens,
    DEFAULT_SINONIM: DEFAULT_SINONIM
  };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = Calc;
