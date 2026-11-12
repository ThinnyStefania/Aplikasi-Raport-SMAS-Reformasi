/* =====================================================================
   E-RAPOR WALI KELAS — app.js  (v2: banyak mata pelajaran)
   Aplikasi satu halaman tanpa server. Data disimpan di localStorage.

   Struktur data per MATA PELAJARAN mengikuti file leger Excel:
     Lingkup Materi (LM) -> Tujuan Pembelajaran (TP) -> 4 model KUDIC
     -> nilai per TP -> sumatif LM, STS, SAS -> Nilai Raport + deskripsi.
   Setiap SEMESTER (I Ganjil / II Genap) punya data sendiri.

   DAFTAR ISI
   1. Util & state          5. Perhitungan (meniru rumus Excel)
   2. Menu & router         6. Halaman (view)
   3. Handler event         7. Sampul, Rapor & Leger (cetak)
   4. Impor Excel           8. Pengaturan, backup, data contoh
   ===================================================================== */
'use strict';

/* =====================================================================
   1. UTIL & STATE
   ===================================================================== */
const KEY = 'erapor_v2';
const $  = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const num = v => { if (v === '' || v == null) return null; const n = Number(String(v).replace(',', '.')); return isNaN(n) ? null : n; };
const uid = () => Math.random().toString(36).slice(2, 9);
const mean = a => a.reduce((x, y) => x + y, 0) / a.length;
const r0 = v => (v == null ? '' : Math.round(v));
const short = n => String(n || '').split(' (')[0];                    // "Knowing (Unistruktural)" -> "Knowing"
const stripDot = s => String(s || '').trim().replace(/[.\s]+$/, '');
const KODE = ['K', 'C', 'B', 'SB'];                                    // urut dari nilai terendah
const BULAN = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const SEM = { '1': 'I (Ganjil)', '2': 'II (Genap)' };
const MODEL_DEFAULT = ['Knowing (Unistruktural)', 'Understanding (Multistruktural)', 'Doing (Relational)', 'Inovasi (Abstrak)'];
const DIM_DEFAULT = ['Keimanan dan Ketakwaan terhadap Tuhan YME', 'Kewargaan', 'Penalaran Kritis', 'Kreativitas', 'Kolaborasi', 'Kemandirian', 'Kesehatan', 'Komunikasi'];
/* Daftar awal 13 mata pelajaran — NAMA BISA DIUBAH di menu Mata Pelajaran */
const MAPEL_DEFAULT = [
  ['Pendidikan Agama dan Budi Pekerti', 'umum'], ['Pendidikan Pancasila', 'umum'], ['Bahasa Indonesia', 'umum'],
  ['Matematika', 'umum'], ['Ilmu Pengetahuan Alam dan Sosial (IPAS)', 'umum'], ['Pendidikan Jasmani, Olahraga, dan Kesehatan (PJOK)', 'umum'],
  ['Seni Rupa', 'umum'], ['Seni Musik', 'umum'], ['Bahasa Inggris', 'umum'], ['Informatika', 'umum'],
  ['Muatan Lokal 1', 'mulok'], ['Muatan Lokal 2', 'mulok'], ['Muatan Lokal 3', 'mulok']
];
const PRED_EKS = ['', 'Sangat Baik', 'Baik', 'Cukup', 'Perlu Bimbingan'];

const newModels = () => MODEL_DEFAULT.map(n => ({ nama: n, ind: '', desk: { SB: '', B: '', C: '', K: '' } }));
const newMapelData = () => ({
  lm: [1, 2, 3, 4].map(i => ({ id: 'LM' + i, nama: 'Lingkup Materi ' + i })),
  tp: [],        // { id, kode, lm, deskripsi, models[4] }
  nilai: {},     // nilai[tpId][siswaId] = { n:[k,u,d,i], cat:'' }
  sumatif: {},   // sumatif[siswaId] = { lm:{LM1:..}, sts, sas }
  hadir: { pertemuan: [], status: {}, tm: '' }   // daftar hadir MAPEL ini: status[siswaId][pertemuanId] = H|S|I|A
});
const newKoku = tema => ({
  id: uid(), tema: tema || 'Tema Kokurikuler',
  dims: DIM_DEFAULT.map(n => ({ nama: n, aspek: ['', '', ''] })),
  skor: {}       // skor[siswaId] = [24 angka 1-4]
});
const newSem = () => ({
  m: {},         // m[mapelId] = newMapelData()
  hari: '',      // jumlah hari efektif
  absen: {},     // absen[siswaId] = { S, I, A }
  eks: {},       // eks[siswaId][ekskulId] = predikat
  koku: [newKoku()],
  catatan: {}    // catatan[siswaId] = catatan wali kelas
});
function defaults() {
  return {
    v: 2,
    profil: {
      sekolah: 'SMA SWASTA REFORMASI PLUS', npsn: '50306573', nss: '321240202001 P', alamat: 'Jln. Timor Raya Km. 14,5 Noelbaki', kodePos: '85361', telp: '082340757072', desa: 'Noelbaki', kecamatan: 'Kupang Tengah', kabupaten: 'Kab. Kupang', provinsi: 'NTT',
      website: 'http://sekolahreformasiplus.weebly.com', email: 'smas.reformasi.plus@gmail.com', kepsek: '', nipKepsek: '',
      guru: '', nipGuru: '', jenjang: 'SMA', kelas: 'X (Sepuluh)', fase: 'E', tahun: '2025/2026', sem: '1',
      tempatRapor: '', tglRapor: '', tglAjaran: '', lambang: '', foto: '', yayasan: 'YAYASAN PENDIDIKAN REFORMASI NOELBAKI (YPRN)', logo: ''
    },
    siswa: [],     // { id, nama, nisn, nis, jk }
    mapel: MAPEL_DEFAULT.map(([nama, kat]) => ({ id: uid(), nama, kat, pem: 'semua', guru: { nama: '' } })),
    ekskul: ['Pramuka', 'Ekstrakurikuler 2', 'Ekstrakurikuler 3', 'Ekstrakurikuler 4'].map(nama => ({ id: uid(), nama })),
    rubrik: [      // berlaku untuk semua mapel (sheet ISI DISINI)
      { dari: 0,  predikat: 'Baru berkembang', tambahan: 'perlu bimbingan lebih lanjut' },
      { dari: 70, predikat: 'Layak',           tambahan: 'perlu penguatan' },
      { dari: 75, predikat: 'Cakap',           tambahan: '' },
      { dari: 90, predikat: 'Mahir',           tambahan: '' }
    ],
    kokuPred: [
      { dari: 0,  nama: 'Perlu Bimbingan', arti: 'Karakter baru mulai tampak; perlu bimbingan lebih lanjut.' },
      { dari: 50, nama: 'Cukup',            arti: 'Karakter sudah tampak tetapi perlu penguatan.' },
      { dari: 75, nama: 'Baik',             arti: 'Karakter tampak baik dan konsisten pada sebagian besar kesempatan.' },
      { dari: 90, nama: 'Sangat Baik',      arti: 'Karakter tampak sangat baik dan konsisten; dapat menjadi teladan bagi teman.' }
    ],
    sems: { '1': newSem(), '2': newSem() }
  };
}
/* ---- Multi kelas: tiap kelas (X, XI, XII) punya data sendiri (siswa, mapel, nilai); identitas sekolah dipakai bersama ---- */
const KEY_AKTIF = 'erapor_v2_aktif';
const KEYK = k => 'erapor_v2_k_' + k;
const KELAS_LIST = [['X', 'X (Sepuluh)', 'E'], ['XI', 'XI (Sebelas)', 'F'], ['XII', 'XII (Dua Belas)', 'F']];
const kelasKey = t => { const m = String(t || '').trim().toUpperCase().match(/^(XII|XI|X)\b/); return m ? m[1] : 'X'; };
const SHARED = ['yayasan', 'sekolah', 'npsn', 'nss', 'alamat', 'kodePos', 'telp', 'desa', 'kecamatan', 'kabupaten', 'provinsi', 'website', 'email', 'kepsek', 'nipKepsek',
  'jenjang', 'tahun', 'sem', 'tempatRapor', 'tglRapor', 'tglAjaran', 'foto', 'logo', 'lambang'];
function normalizeState(sv) {
  const d = defaults(), st = Object.assign(d, sv); st.profil = Object.assign({}, d.profil, sv.profil || {});
  ['yayasan', 'sekolah', 'alamat', 'kecamatan', 'kabupaten', 'provinsi', 'email', 'npsn', 'nss', 'kodePos', 'telp', 'desa', 'website'].forEach(k => { if (!st.profil[k]) st.profil[k] = defaults().profil[k]; });
  return st;
}
function loadClass(k) { try { const raw = localStorage.getItem(KEYK(k)); return raw ? normalizeState(JSON.parse(raw)) : null; } catch (e) { return null; } }
function newKelas(k, from) {          // kelas baru: daftar mapel & ekskul disalin dari kelas asal (tanpa siswa/nilai)
  const d = defaults(), kl = KELAS_LIST.find(x => x[0] === k) || KELAS_LIST[0];
  d.profil.kelas = kl[1]; d.profil.fase = kl[2];
  if (from) {
    d.mapel = from.mapel.map(m => ({ id: uid(), nama: m.nama, kat: m.kat, pem: m.pem || 'semua', guru: { nama: '' } }));
    d.ekskul = from.ekskul.map(e => ({ id: uid(), nama: e.nama }));
    d.rubrik = JSON.parse(JSON.stringify(from.rubrik)); d.kokuPred = JSON.parse(JSON.stringify(from.kokuPred));
    SHARED.forEach(f => { d.profil[f] = from.profil[f]; });
  }
  return d;
}
let KA = 'X';
function loadState() {
  let old = null;
  try {
    if (!localStorage.getItem(KEY_AKTIF)) {            // data lama (satu kelas) → pindahkan ke slot kelasnya
      old = localStorage.getItem(KEY);
      if (old) { const k = kelasKey((JSON.parse(old).profil || {}).kelas); localStorage.setItem(KEYK(k), old); localStorage.setItem(KEY_AKTIF, k); localStorage.removeItem(KEY); }
    }
  } catch (e) { /* kuota penuh: tetap pakai data lama di bawah */ }
  try {
    KA = localStorage.getItem(KEY_AKTIF) || (old ? kelasKey((JSON.parse(old).profil || {}).kelas) : 'X');
    const st = loadClass(KA); if (st) return st;
    if (old) return normalizeState(JSON.parse(old));
  } catch (e) { /* abaikan */ }
  return newKelas(KA);
}
let S = loadState();
const LOGO_DEFAULT = 'data:image/jpeg;base64,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';
const logoSrc = () => S.profil.logo || LOGO_DEFAULT;
const UI = { mid: '', tp: '', sid: '', kid: '', leger: 'akhir', allTp: false, tabKoku: 'skor', tpMid: '', lamp: false };

function gantiKelas(k) {
  if (k === KA) return;
  clearTimeout(saveTimer);
  try { localStorage.setItem(KEYK(KA), JSON.stringify(S)); } catch (e) { toast('Penyimpanan browser penuh. Buat backup lalu hapus foto.'); return; }
  const prev = S; let nx = loadClass(k), baru = false;
  if (!nx) { nx = newKelas(k, prev); baru = true; }
  SHARED.forEach(f => { nx.profil[f] = prev.profil[f]; });        // identitas sekolah sama untuk semua kelas
  S = nx; KA = k; UI.mid = UI.tp = UI.sid = UI.kid = UI.tpMid = '';
  save(); render(); toast(baru ? `Kelas ${k} dibuat: daftar mapel disalin dari kelas sebelumnya. Isi data siswa & wali kelas.` : `Kelas ${k}`);
}
let saveTimer;
function save() {
  $('#saveState').classList.add('dirty');
  $('#saveState').innerHTML = '<i class="fa-solid fa-pen"></i> Menyimpan…';
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    try { localStorage.setItem(KEYK(KA), JSON.stringify(S)); localStorage.setItem(KEY_AKTIF, KA); }
    catch (e) { toast('Penyimpanan browser penuh. Hapus foto profil atau ekspor backup.'); }
    $('#saveState').classList.remove('dirty');
    $('#saveState').innerHTML = '<i class="fa-solid fa-circle-check"></i> Tersimpan';
  }, 350);
}
const getPath = (o, p) => p.split('.').reduce((a, k) => (a == null ? a : a[k]), o);
function setPath(o, p, v) {
  const ks = p.split('.'); const last = ks.pop();
  const t = ks.reduce((a, k) => (a[k] == null ? (a[k] = {}) : a[k]), o);
  t[last] = v;
}
function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('show'), 2800);
}
function openModal(title, body, foot) {
  $('#modal').innerHTML = `<div class="box"><h3>${title}</h3><div class="mb">${body}</div>
    <div class="mf">${foot || '<button class="btn" data-act="closeModal">Tutup</button>'}</div></div>`;
  $('#modal').hidden = false;
}
const closeModal = () => { $('#modal').hidden = true; };
const opts = (list, sel) => list.map(o => {
  const [v, t] = Array.isArray(o) ? o : [o, o];
  return `<option value="${esc(v)}" ${String(v) === String(sel) ? 'selected' : ''}>${esc(t)}</option>`;
}).join('');

/* ---- akses data ---- */
const D = () => S.sems[S.profil.sem];                                  // data semester aktif
const MD = mid => {                                                    // data mapel pada semester aktif
  const md = (D().m[mid] ||= newMapelData());
  md.hadir ||= { pertemuan: [], status: {}, tm: '' };                  // data lama belum punya daftar hadir
  return md;
};
const mapelOf = id => S.mapel.find(m => m.id === id) || {};
const cur = () => MD(UI.mid);                                          // data mapel yang sedang dipilih
const siswaOf = id => S.siswa.find(s => s.id === id) || {};
const namaOf = id => siswaOf(id).nama || '';
/* MAPEL PILIHAN: tiap siswa memilih kelompok pilihan (IPA/IPS); tiap mapel berlaku untuk "semua" atau satu kelompok pilihan */
const berlaku = (m, sid) => !m.pem || m.pem === 'semua' || m.pem === (siswaOf(sid).pem || '');
const siswaMapel = mid => S.siswa.filter(s => berlaku(mapelOf(mid), s.id));   // siswa yang mengikuti mapel ini
const mapelSiswa = sid => S.mapel.filter(m => berlaku(m, sid));                // mapel yang diikuti siswa ini
const pemOpts = [['semua', 'Semua siswa (umum)'], ['IPA', 'Pilihan IPA'], ['IPS', 'Pilihan IPS']];
function siswaSummary() {
  const n = k => S.siswa.filter(s => (s.pem || '') === k).length;
  return `<div class="hint info"><b>Mata pelajaran pilihan:</b> IPA <b>${n('IPA')}</b> siswa · IPS <b>${n('IPS')}</b> siswa${n('') ? ` · <span style="color:#b45309"><b>${n('')}</b> belum diatur</span>` : ''}.
    Siswa yang belum diatur hanya mengikuti mata pelajaran “Semua siswa”. Mata pelajaran pilihan diatur di menu <a href="#/mapel">Mata Pelajaran</a>.</div>`;
}
const tpById = id => cur().tp.find(t => t.id === id);
function ensureUI() {
  if (!mapelOf(UI.mid).id) UI.mid = (S.mapel[0] || {}).id || '';
  if (UI.mid && !tpById(UI.tp)) UI.tp = (cur().tp[0] || {}).id || '';
  if (!siswaOf(UI.sid).id) UI.sid = (S.siswa[0] || {}).id || '';
  if (!D().koku.find(k => k.id === UI.kid)) UI.kid = (D().koku[0] || {}).id || '';
}
function tglID(iso) {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  return y ? `${d} ${BULAN[m - 1]} ${y}` : '';
}

/* =====================================================================
   2. MENU & ROUTER
   ===================================================================== */
const H = {}, C = {}, A = {};   // handler: input (H), change (C), klik (A) — lihat bagian 3

const MENU = [
  { id: 'dashboard', t: 'DASHBOARD', i: 'fa-bars', head: true },
  { id: 'siswa', t: 'Data Siswa', i: 'fa-users' },
  { id: 'mapel', t: 'Mata Pelajaran', i: 'fa-book' },
  { id: 'absen', t: 'Absen Wali Kelas', i: 'fa-calendar-check' },
  { id: 'hadir', t: 'Daftar Hadir Mapel', i: 'fa-clipboard-user' },
  { g: 'tp', t: 'Tujuan Pembelajaran', i: 'fa-bullseye', sub: [['tp-lm', 'Lingkup Materi & TP'], ['tp-rubrik', 'Rubrik Predikat'], ['tp-kudic', 'Deskripsi KUDIC']] },
  { id: 'ekskul', t: 'Ektrakurikuler', i: 'fa-futbol' },
  { g: 'nilai', t: 'Input Nilai', i: 'fa-pen-to-square', sub: [['nilai-tp', 'Nilai TP (KUDIC)'], ['nilai-sum', 'Sumatif LM, STS & SAS']] },
  { id: 'rekap', t: 'Rekap Nilai', i: 'fa-layer-group' },
  { id: 'sampul', t: 'Sampul & Identitas', i: 'fa-file-lines' },
  { g: 'koku', t: 'Input Kokurikuler', i: 'fa-puzzle-piece', sub: [['koku-tema', 'Kegiatan & Aspek'], ['koku-nilai', 'Input Skor'], ['koku-pred', 'Arti Skor & Predikat']] },
  { g: 'rpt', t: 'RAPORT', i: 'fa-award', sub: [['rapor1', 'Rapor I (Ganjil)'], ['rapor2', 'Rapor II (Genap)'], ['rapor-tp', 'Rapor Kemampuan per TP'], ['leger', 'Leger']] },
  { id: 'pengaturan', t: 'Pengaturan', i: 'fa-gear' }
];
const openGroups = new Set(['tp', 'nilai', 'koku', 'rpt']);
function renderNav(cur_) {
  $('#nav').innerHTML = MENU.map(m => {
    if (m.sub) {
      const act = m.sub.some(s => s[0] === cur_);
      return `<div class="nav-group ${act || openGroups.has(m.g) ? 'open' : ''}">
        <button data-act="toggleGroup" data-g="${m.g}"><i class="fa-solid ${m.i} ic"></i>${m.t}<i class="fa-solid fa-chevron-down caret"></i></button>
        <div class="nav-sub">${m.sub.map(s => `<a class="nav-link ${s[0] === cur_ ? 'active' : ''}" href="#/${s[0]}">${s[1]}</a>`).join('')}</div></div>`;
    }
    return `<a class="nav-link ${m.id === cur_ ? 'active' : ''}" href="#/${m.id}"><i class="fa-solid ${m.i}"></i>${m.t}</a>`;
  }).join('');
}
const VIEW = {};   // nama route -> fungsi pengembali HTML
const AFTER = {};  // (opsional) dijalankan setelah HTML dipasang
const route = () => (location.hash.replace(/^#\/?/, '') || 'dashboard');
function render() {
  let r = route(); if (!VIEW[r]) r = 'dashboard';
  ensureUI();
  renderNav(r);
  $('#view').innerHTML = VIEW[r]();
  if (AFTER[r]) AFTER[r]();
  updateChrome();
  $('#sidebar').classList.remove('open'); $('#backdrop').classList.remove('show');
  window.scrollTo(0, 0);
}
let _wm = '';
function updateChrome() {
  const p = S.profil;
  const wm = logoSrc(); if (wm !== _wm) { _wm = wm; const svg = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 216 329"><image xlink:href="${wm}" href="${wm}" x="45.5" y="100" width="125" height="125" preserveAspectRatio="xMidYMid meet"/></svg>`;
    document.documentElement.style.setProperty('--wm', `url("data:image/svg+xml,${encodeURIComponent(svg)}")`); }   // logo sekolah sebagai watermark
  $('#brandText').textContent = `RAPOR ${String(p.tahun).slice(0, 4)} ${p.jenjang}`;
  $('#schoolName').textContent = p.sekolah || 'Nama Sekolah (isi di Pengaturan)';
  $('#hdrYear').textContent = p.tahun;
  $('#semPick').value = p.sem; $('#kelasPick').value = KA;
  $('#userRole').textContent = `Guru Kelas ${p.kelas}`;
  $('#userName').textContent = p.guru || 'Nama Guru';
  $('#userNip').textContent = 'NIP/ ' + (p.nipGuru || '');
  document.title = `E-Rapor ${p.jenjang} Kelas ${p.kelas}`;
}
window.addEventListener('hashchange', render);

/* =====================================================================
   5. PERHITUNGAN — meniru rumus pada file Excel
   ===================================================================== */
/* index rubrik (0..3) dari sebuah nilai: Baru berkembang / Layak / Cakap / Mahir */
function idxRubrik(n) {
  let r = 0; S.rubrik.forEach((x, i) => { if (n >= (num(x.dari) ?? 0)) r = i; });
  return r;
}
const kodeOf = n => KODE[idxRubrik(n)];                    // SB / B / C / K
const predNama = n => S.rubrik[idxRubrik(n)].predikat;
function rangeOf(kode) {
  const r = S.rubrik.map(x => num(x.dari) ?? 0); const i = KODE.indexOf(kode);
  return `${r[i]} - ${i === 3 ? 100 : r[i + 1] - 1}`;
}

/* Hasil KUDIC satu siswa pada satu TP (= sheet "Leger TP x.y") */
function tpCalc(md, tp, sid) {
  const rec = (md.nilai[tp.id] || {})[sid]; if (!rec) return null;
  const vals = [0, 1, 2, 3].map(i => num((rec.n || [])[i]));
  const f = vals.filter(v => v != null); if (!f.length) return null;
  const avg = Math.round(mean(f));
  const mx = Math.max(...f), mn = Math.min(...f);
  const iMax = vals.indexOf(mx);
  const iMin = (f.length < 2 || mx === mn) ? -1 : vals.indexOf(mn);
  const nama = namaOf(sid);
  const dsc = i => i < 0 ? '' : stripDot((tp.models[i].desk || {})[kodeOf(vals[i])]);
  const tdesc = stripDot(tp.deskripsi) || tp.kode;
  const rMax = S.rubrik[idxRubrik(mx)];
  let T = `${nama} menunjukkan capaian ${rMax.predikat} dalam ${tdesc}, terutama pada tahap ${short(tp.models[iMax].nama)}`;
  if (dsc(iMax)) T += `: ${dsc(iMax)}`;
  if (rMax.tambahan) T += `, masih ${rMax.tambahan}`;
  T += '.';
  let U = '';
  if (iMin >= 0) {
    const rMin = S.rubrik[idxRubrik(mn)];
    U = rMin.tambahan
      ? `${nama} masih ${rMin.tambahan} dalam ${tdesc}, terutama pada tahap ${short(tp.models[iMin].nama)}`
      : `${nama} menunjukkan capaian ${rMin.predikat} dalam ${tdesc}, dengan tahap ${short(tp.models[iMin].nama)} sebagai yang relatif terendah`;
    if (dsc(iMin)) U += `: ${dsc(iMin)}`;
    U += '.';
  }
  return { vals, avg, iMax, iMin, T, U, cat: rec.cat || '' };
}

/* Rekap satu siswa pada satu mapel (= sheet SUMATIF) */
function calcSiswa(mid, sid) {
  const md = MD(mid);
  const form = {}; const fv = [];
  md.tp.forEach(tp => { const c = tpCalc(md, tp, sid); if (c) { form[tp.id] = c; fv.push(c.avg); } });
  const avgF = fv.length ? mean(fv) : null;
  const sm = md.sumatif[sid] || {};
  const lmVals = {}; const lv = [];
  md.lm.forEach(l => { const v = num((sm.lm || {})[l.id]); lmVals[l.id] = v; if (v != null) lv.push(v); });
  const avgLM = lv.length ? mean(lv) : null;
  const sts = num(sm.sts), sas = num(sm.sas);
  const parts = [avgLM, sts, sas, avgF].filter(v => v != null);
  const akhir = parts.length ? mean(parts) : null;

  const list = md.tp.filter(t => form[t.id]);
  let hi = null, lo = null, descHi = '', descLo = '';
  if (list.length) {
    const mx = Math.max(...list.map(t => form[t.id].avg));
    const mn = Math.min(...list.map(t => form[t.id].avg));
    hi = list.find(t => form[t.id].avg === mx);
    if (list.length >= 2 && mx !== mn) lo = list.find(t => form[t.id].avg === mn);
    descHi = form[hi.id].T;
    const nama = namaOf(sid);
    if (list.length >= 2 && mx === mn) {
      descLo = `${nama} menunjukkan capaian ${predNama(mx)} pada seluruh TP yang dinilai.`;
    } else if (lo) {
      const rM = S.rubrik[idxRubrik(mn)]; const td = stripDot(lo.deskripsi) || lo.kode;
      descLo = form[lo.id].U ||
        (rM.tambahan ? `${nama} masih ${rM.tambahan} dalam ${td}.` : `${nama} menunjukkan capaian ${rM.predikat} dalam ${td}.`);
    }
  }
  return { md, form, avgF, lmVals, avgLM, sts, sas, akhir, hi, lo, descHi, descLo,
           predikat: akhir == null ? '' : predNama(Math.round(akhir)) };
}

/* Kokurikuler / Profil Lulusan (= sheet "Profil Lulusan") */
function calcKoku(k, sid) {
  const sk = k.skor[sid] || [];
  const dimVal = k.dims.map((d, di) => {
    const v = [0, 1, 2].filter(x => (d.aspek[x] || '').trim()).map(x => num(sk[di * 3 + x])).filter(x => x != null);
    return v.length ? mean(v) / 4 * 100 : null;
  });
  const f = dimVal.filter(v => v != null);
  if (!f.length) return { dimVal, avg: null };
  const avg = mean(f);
  let pi = 0; S.kokuPred.forEach((p, i) => { if (avg >= (num(p.dari) ?? 0)) pi = i; });
  const pred = S.kokuPred[pi];
  const mx = Math.max(...f), mn = Math.min(...f);
  const hi = k.dims[dimVal.indexOf(mx)].nama;
  const lo = (f.length >= 2 && mx !== mn) ? k.dims[dimVal.indexOf(mn)].nama : '';
  let desc;
  if (f.length === 1) desc = `Ananda tergolong ${pred.nama.toLowerCase()} pada dimensi ${hi}.`;
  else if (mx === mn) desc = `Ananda tergolong ${pred.nama.toLowerCase()} dan konsisten pada seluruh dimensi yang dinilai.`;
  else desc = `Ananda tergolong ${pred.nama.toLowerCase()}, paling menonjol pada dimensi ${hi} dan perlu penguatan pada dimensi ${lo}.`;
  return { dimVal, avg, pred: pred.nama, hi, lo, desc };
}

/* Absen */
function absCalc(sid) {
  const a = D().absen[sid] || {}; const s = num(a.S) || 0, i = num(a.I) || 0, al = num(a.A) || 0;
  const hari = num(D().hari);
  return { S: s, I: i, A: al, total: s + i + al, pct: hari ? Math.max(0, (hari - s - i - al) / hari * 100) : null };
}
const eksKet = (pred, nama) => !pred ? '' :
  pred === 'Perlu Bimbingan' ? `Ananda perlu bimbingan dalam mengikuti kegiatan ${nama}.` : `Ananda ${pred.toLowerCase()} dalam mengikuti kegiatan ${nama}.`;
/* peringkat berdasarkan jumlah nilai akhir seluruh mapel */
function allAkhir(sid) {
  const v = S.mapel.map(m => berlaku(m, sid) ? calcSiswa(m.id, sid).akhir : null);
  const f = v.filter(x => x != null);
  return { v, jumlah: f.length ? f.reduce((a, b) => a + Math.round(b), 0) : null, rata: f.length ? mean(f.map(Math.round)) : null };
}

/* =====================================================================
   3. HANDLER EVENT  (data-h = input, data-c = change, data-act = klik)
   ===================================================================== */
document.addEventListener('input', e => {
  const el = e.target;
  if (el.dataset.bind) { setPath(S, el.dataset.bind, el.value); save(); updateChrome(); }
  else if (el.dataset.h && H[el.dataset.h]) H[el.dataset.h](el, e);
});
document.addEventListener('change', e => {
  const el = e.target;
  if (el.dataset.c && C[el.dataset.c]) C[el.dataset.c](el, e);
});
document.addEventListener('click', e => {
  const b = e.target.closest('[data-act]');
  if (b && A[b.dataset.act]) { if (b.tagName === 'A') e.preventDefault(); A[b.dataset.act](b, e); }
});
$('#burger').onclick = () => { $('#sidebar').classList.toggle('open'); $('#backdrop').classList.toggle('show'); };
$('#backdrop').onclick = () => { $('#sidebar').classList.remove('open'); $('#backdrop').classList.remove('show'); };
const toggleFs = () => { if (!document.fullscreenElement) document.documentElement.requestFullscreen?.(); else document.exitFullscreen?.(); };
$('#fsBtn').onclick = toggleFs; $('#fsBtn2').onclick = toggleFs;
$('#kelasPick').onchange = e => gantiKelas(e.target.value);
$('#semPick').onchange = e => { S.profil.sem = e.target.value; save(); render(); toast('Semester ' + SEM[S.profil.sem]); };
A.closeModal = closeModal;
A.toggleGroup = b => { const g = b.dataset.g; openGroups.has(g) ? openGroups.delete(g) : openGroups.add(g); b.parentElement.classList.toggle('open'); };

/* =====================================================================
   6. HALAMAN
   ===================================================================== */
const panel = (title, icon, body, tools = '') =>
  `<section class="panel"><div class="panel-h"><i class="fa-solid ${icon}"></i>${title}<div class="tools">${tools}</div></div><div class="panel-b">${body}</div></section>`;
const field = (label, bind, type = 'text', extra = '') =>
  `<div><label class="f">${label}</label><input type="${type}" data-bind="${bind}" value="${esc(getPath(S, bind))}" ${extra}></div>`;
const need = (cond, msg, link) => cond ? '' : `<div class="hint">${msg} ${link ? `<a href="#/${link[0]}">${link[1]}</a>` : ''}</div>`;
const noSiswa = t => panel(t, 'fa-users', need(false, 'Isi data siswa dulu.', ['siswa', 'Ke Data Siswa']));
const noMapel = t => panel(t, 'fa-book', need(false, 'Belum ada mata pelajaran.', ['mapel', 'Ke Mata Pelajaran']));
const mapelBar = (extra = '') => `<div class="row mapelbar">
  <b><i class="fa-solid fa-book"></i> Mata pelajaran:</b>
  <select data-c="pickMapel" style="max-width:420px">${opts(S.mapel.map(m => [m.id, m.nama]), UI.mid)}</select>
  <span class="pill">${SEM[S.profil.sem]}</span>${(mapelOf(UI.mid).pem && mapelOf(UI.mid).pem !== 'semua') ? `<span class="pill pem">Pilihan ${mapelOf(UI.mid).pem}</span>` : ''}
  <span class="muted small">${siswaMapel(UI.mid).length} siswa mengikuti</span>${extra}</div>`;
C.pickMapel = el => { UI.mid = el.value; UI.tp = ''; render(); };
C.pickTp = el => { UI.tp = el.value; render(); };

/* ---------------- DASHBOARD ---------------- */
VIEW.dashboard = () => {
  const p = S.profil, d = D(); const nSis = S.siswa.length;
  const nIPA = S.siswa.filter(s => s.pem === 'IPA').length, nIPS = S.siswa.filter(s => s.pem === 'IPS').length;
  const nUmum = S.mapel.filter(m => m.kat !== 'mulok').length, nMulok = S.mapel.length - nUmum;
  // kelengkapan: pasangan (mapel, siswa) yang sudah punya nilai
  let filled = 0, totalPair = 0; S.mapel.forEach(m => siswaMapel(m.id).forEach(s => { totalPair++; if (calcSiswa(m.id, s.id).akhir != null) filled++; }));
  const fAbs = S.siswa.filter(s => d.absen[s.id] && (d.absen[s.id].S !== undefined || d.absen[s.id].I !== undefined || d.absen[s.id].A !== undefined)).length;
  const fEks = S.siswa.filter(s => Object.values(d.eks[s.id] || {}).some(Boolean)).length;
  const fKoku = S.siswa.filter(s => d.koku.some(k => calcKoku(k, s.id).avg != null)).length;
  const pct = (a, b) => b ? Math.round(a / b * 100) : 0;
  const prog = (l, a, b) => `<div class="prog-row"><div class="lbl"><span>${l}</span><b>${pct(a, b)}%</b></div><div class="progress"><span style="width:${pct(a, b)}%"></span></div></div>`;
  const stat = (cls, icon, big, label, lIcon, to) => `<a class="stat ${cls}" href="#/${to}"><div class="top"><i class="fa-solid ${icon}"></i>${big}</div><div class="bot"><i class="fa-solid ${lIcon}"></i>${label}</div></a>`;
  const row = (k, v) => `<div class="k">${k}</div><div class="v">${esc(v)}</div>`;
  const tema = (d.koku[0] || {}).tema || '';
  return `
  <div class="stats">
    ${stat('s-green', 'fa-users', `<span class="num">${nSis}</span>`, 'Jumlah Siswa', 'fa-users', 'siswa')}
    ${stat('s-blue', 'fa-database', `<span class="num">${nUmum}</span>`, 'Jumlah Mapel', 'fa-database', 'mapel')}
    ${stat('s-pink', 'fa-heart-pulse', `<span class="num">${S.ekskul.length}</span>`, 'Jumlah Ekstra', 'fa-heart-pulse', 'ekskul')}
    ${stat('s-green2', 'fa-pen-to-square', `<span class="num">${nMulok}</span>`, 'Muatan Lokal', 'fa-pen-to-square', 'mapel')}
    ${stat('s-orange', 'fa-book-open-reader', `<span class="txt">${esc(tema)}</span>`, 'Kokurikuler', 'fa-book-open', 'koku-tema')}
  </div>
  <div class="cols">
    <div>
      <section class="panel">
        <div class="panel-h"><i class="fa-solid fa-gear"></i>Profil Sekolah<div class="tools"><a class="btn sm ghost" href="#/pengaturan"><i class="fa-solid fa-pen"></i>Ubah</a></div></div>
        <div class="panel-b prof-grid">
          <div class="kv">
            ${row('Nama Sekolah', p.sekolah)}${row('NPSN', p.npsn)}${row('NSS', p.nss)}${row('Alamat Sekolah', p.alamat)}
            ${row('Kode Pos', p.kodePos)}${row('Desa / Kelurahan', p.desa)}${row('Kecamatan', p.kecamatan)}
            ${row('Kabupaten / Kota', p.kabupaten)}${row('Provinsi', p.provinsi)}${row('Website', p.website || '-')}
            ${row('E-mail', p.email || '-')}${row('Nama Kepala Sekolah', p.kepsek)}${row('NIP', p.nipKepsek)}
          </div>
          <div class="photo">
            ${p.foto ? `<img src="${p.foto}" alt="Foto">` : '<i class="fa-solid fa-school ph"></i>'}
            <label><i class="fa-solid fa-camera"></i> Ganti foto<input type="file" accept="image/*" data-c="foto" hidden></label>
          </div>
        </div>
      </section>
    </div>
    <div>
      <section class="panel">
        <div class="panel-h"><i class="fa-regular fa-face-smile"></i>Profil Guru</div>
        <div class="panel-b"><div class="kv">
          ${row('Nama Guru kelas', p.guru)}${row('NIP/NIPPPK', p.nipGuru)}${row('Kelas', p.kelas)}${row('Fase', p.fase)}
          ${row('Tahun Pelajaran', p.tahun)}${row('Semester', SEM[p.sem])}
          ${row('Tempat/Tanggal Rapor', [p.tempatRapor, tglID(p.tglRapor)].filter(Boolean).join(', '))}
          ${row('Tanggal Tahun Ajaran', p.tglAjaran)}${row('Jumlah Siswa', nSis + (nIPA || nIPS ? ` (IPA ${nIPA}, IPS ${nIPS})` : ''))}
        </div></div>
        <div class="panel-foot"><i class="fa-solid fa-fingerprint"></i></div>
      </section>
      <section class="panel">
        <div class="panel-h" style="font-size:16px">Informasi</div>
        <div class="panel-b">
          <div class="kv" style="grid-template-columns:150px 1fr">
            ${row('Versi Aplikasi', 'v. 2.0 (HTML • CSS • JS)')}${row('Penyimpanan', 'Browser (localStorage)')}
          </div>
          <div class="mt"><b>Kelengkapan data — ${SEM[p.sem]}</b>
            ${prog('Nilai mapel (siswa × mapel)', filled, totalPair)}
            ${prog('Absen', fAbs, nSis)}
            ${prog('Ekstrakurikuler', fEks, nSis)}
            ${prog('Kokurikuler', fKoku, nSis)}
          </div>
        </div>
      </section>
    </div>
  </div>
  ${mapelStatusPanel()}`;
};
C.foto = el => {
  const f = el.files[0]; if (!f) return;
  const rd = new FileReader();
  rd.onload = () => {
    const img = new Image();
    img.onload = () => {            // perkecil agar hemat localStorage
      const k = Math.min(1, 600 / Math.max(img.width, img.height));
      const cv = document.createElement('canvas'); cv.width = img.width * k; cv.height = img.height * k;
      cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
      S.profil.foto = cv.toDataURL('image/jpeg', .82); save(); render();
    };
    img.src = rd.result;
  };
  rd.readAsDataURL(f);
};

/* ---------------- DATA SISWA ---------------- */
VIEW.siswa = () => panel('Data Siswa', 'fa-users', `${siswaSummary()}
    <div class="tw"><table class="grid">
      <thead><tr><th>No</th><th>Nama Peserta Didik</th><th>NISN</th><th>NIS</th><th>L/P</th><th>Mapel Pilihan</th><th>Identitas</th><th></th></tr></thead>
      <tbody>${S.siswa.map((s, i) => `<tr>
        <td class="c">${i + 1}</td>
        <td><input type="text" value="${esc(s.nama)}" data-h="siswa" data-i="${i}" data-f="nama"></td>
        <td><input type="text" value="${esc(s.nisn || '')}" data-h="siswa" data-i="${i}" data-f="nisn"></td>
        <td><input type="text" value="${esc(s.nis || '')}" data-h="siswa" data-i="${i}" data-f="nis"></td>
        <td><select data-c="siswaJk" data-i="${i}">${opts([['', '-'], 'L', 'P'], s.jk || '')}</select></td>
        <td><select data-c="siswaPem" data-i="${i}">${opts([['', '— belum —'], 'IPA', 'IPS'], s.pem || '')}</select></td>
        <td class="c"><button class="btn sm ghost" data-act="editIdent" data-i="${i}" title="Isi identitas & pas foto"><i class="fa-solid fa-id-card"></i>${s.ayah || s.tempatLahir || s.alamat ? '' : ' <span style="color:#b45309">kosong</span>'}</button></td>
        <td class="c"><button class="btn sm danger icon" data-act="delSiswa" data-i="${i}" title="Hapus"><i class="fa-solid fa-trash"></i></button></td>
      </tr>`).join('') || '<tr><td colspan="8" class="empty">Belum ada siswa. Klik “Tambah” atau “Tempel daftar nama”.</td></tr>'}</tbody>
    </table></div>`,
    `<button class="btn" data-act="addSiswa"><i class="fa-solid fa-plus"></i>Tambah</button>
     <button class="btn ghost" data-act="pasteSiswa"><i class="fa-solid fa-paste"></i>Tempel daftar nama</button>
     <button class="btn ghost" data-act="importIdent"><i class="fa-solid fa-file-excel"></i>Impor identitas (Excel)</button>
     <button class="btn ghost" data-act="tplIdent"><i class="fa-solid fa-download"></i>Unduh template</button>`);
H.siswa = el => { S.siswa[+el.dataset.i][el.dataset.f] = el.value; save(); };
C.siswaJk = el => { S.siswa[+el.dataset.i].jk = el.value; save(); };
C.siswaPem = el => { S.siswa[+el.dataset.i].pem = el.value; save(); render(); };
A.addSiswa = () => { S.siswa.push({ id: uid(), nama: '', nisn: '', nis: '', jk: '', pem: '' }); save(); render(); setTimeout(() => $$('[data-f=nama]').pop()?.focus(), 30); };
A.delSiswa = b => {
  const s = S.siswa[+b.dataset.i];
  if (!confirm(`Hapus ${s.nama || 'siswa ini'} beserta seluruh nilainya (semua mapel & semester)?`)) return;
  Object.values(S.sems).forEach(d => {
    Object.values(d.m).forEach(md => { Object.values(md.nilai).forEach(o => delete o[s.id]); delete md.sumatif[s.id]; if (md.hadir) delete md.hadir.status[s.id]; });
    delete d.absen[s.id]; delete d.eks[s.id]; delete d.catatan[s.id]; d.koku.forEach(k => delete k.skor[s.id]);
  });
  S.siswa.splice(+b.dataset.i, 1); save(); render();
};
A.pasteSiswa = () => openModal('Tempel daftar nama siswa',
  `<p class="small muted">Satu nama per baris (bisa disalin dari Excel). Kelompok mapel pilihan boleh ditulis setelah nama, contoh: <code>Budi Santoso;IPA</code> atau <code>Sari Dewi;IPS</code>.</p><textarea id="pasteBox" style="min-height:240px"></textarea>`,
  `<button class="btn ghost" data-act="closeModal">Batal</button><button class="btn ok" data-act="doPaste">Tambahkan</button>`);
A.doPaste = () => {
  const lines = $('#pasteBox').value.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  lines.forEach(l => {                       // "Nama" atau "Nama;IPA" / "Nama<TAB>IPS"
    const [nama, pem] = l.split(/[\t;|]/).map(x => x.trim());
    S.siswa.push({ id: uid(), nama, nisn: '', nis: '', jk: '', pem: /^IP[AS]$/i.test(pem || '') ? pem.toUpperCase() : '' });
  });
  closeModal(); save(); render(); toast(`${lines.length} siswa ditambahkan`);
};

/* ---------------- MATA PELAJARAN (13 mapel + tambah) ---------------- */
VIEW.mapel = () => `
  ${panel('Mata Pelajaran', 'fa-book', `
    <div class="hint info"><b>Mata pelajaran pilihan:</b> pada kolom <i>Berlaku untuk</i> pilih “Semua siswa (umum)”, “Pilihan IPA”, atau “Pilihan IPS”. Daftar nilai, daftar hadir, rapor, dan leger otomatis hanya memuat siswa yang mengikuti mapel itu (atur pilihan tiap siswa di menu Data Siswa).<br><b>Mapel pilihan yang diikuti semua siswa:</b> ubah kolom <i>Kategori</i> menjadi “Mapel pilihan” dan biarkan <i>Berlaku untuk</i> = “Semua siswa”. Mapel itu tampil di kelompok “Mata Pelajaran Pilihan” pada rapor untuk seluruh siswa.<br>Setiap mata pelajaran punya <b>Lingkup Materi, TP, deskripsi KUDIC, nilai, sumatif, STS & SAS</b> sendiri, persis seperti file leger Anda
      (dipakai per semester). Pilih <i>Kelola</i> untuk mengisi TP mapel tersebut, atau <i>Impor leger</i> untuk memasukkan file leger Excel mapel itu.</div>
    <div class="row" style="margin-bottom:12px">
      <input type="text" id="mapelBaru" placeholder="Nama mata pelajaran baru…" style="max-width:380px">
      <select id="mapelKat" style="width:170px">${opts([['umum', 'Mata pelajaran'], ['pilihan', 'Mapel pilihan'], ['mulok', 'Muatan lokal']], 'umum')}</select>
      <select id="mapelPem" style="width:170px">${opts(pemOpts, 'semua')}</select>
      <button class="btn ok" data-act="addMapel"><i class="fa-solid fa-plus"></i>Tambah mata pelajaran</button>
      <button class="btn ghost" data-act="pasteMapel"><i class="fa-solid fa-paste"></i>Tempel daftar</button>
    </div>
    <div class="tw"><table class="grid"><thead><tr><th>No</th><th>Nama Mata Pelajaran</th><th>Kategori</th><th>Berlaku untuk</th><th>Guru pengampu</th><th>Jumlah TP (${SEM[S.profil.sem]})</th><th>Aksi</th></tr></thead>
    <tbody>${S.mapel.map((m, i) => `<tr>
      <td class="c">${i + 1}</td>
      <td><input type="text" value="${esc(m.nama)}" data-h="mapel" data-i="${i}" data-f="nama"></td>
      <td style="width:170px"><select data-c="mapelKatSet" data-i="${i}">${opts([['umum', 'Mata pelajaran'], ['pilihan', 'Mapel pilihan'], ['mulok', 'Muatan lokal']], m.kat)}</select></td>
      <td style="width:170px"><select data-c="mapelPemSet" data-i="${i}">${opts(pemOpts, m.pem || 'semua')}</select></td>
      <td><input type="text" value="${esc((m.guru || {}).nama || '')}" data-h="mapelGuru" data-i="${i}" placeholder="Nama guru"></td>
      <td class="c">${MD(m.id).tp.length}</td>
      <td class="c" style="white-space:nowrap">
        <button class="btn sm" data-act="kelolaMapel" data-i="${i}"><i class="fa-solid fa-sliders"></i>Kelola</button>
        <button class="btn sm ghost" data-act="importLeger" data-i="${i}"><i class="fa-solid fa-file-import"></i>Impor leger</button>
        <button class="btn sm ghost icon" data-act="mapelUp" data-i="${i}" title="Naik"><i class="fa-solid fa-arrow-up"></i></button>
        <button class="btn sm ghost icon" data-act="mapelDown" data-i="${i}" title="Turun"><i class="fa-solid fa-arrow-down"></i></button>
        <button class="btn sm danger icon" data-act="delMapel" data-i="${i}" title="Hapus"><i class="fa-solid fa-trash"></i></button></td></tr>`).join('')
      || '<tr><td colspan="7" class="empty">Belum ada mata pelajaran.</td></tr>'}</tbody></table></div>`)}`;
H.mapel = el => { S.mapel[+el.dataset.i][el.dataset.f] = el.value; save(); };
H.mapelGuru = el => { (S.mapel[+el.dataset.i].guru ||= { nama: '' }).nama = el.value; save(); };
C.mapelPemSet = el => { S.mapel[+el.dataset.i].pem = el.value; save(); render(); };
C.mapelKatSet = el => { S.mapel[+el.dataset.i].kat = el.value; save(); render(); };
A.addMapel = () => {
  const n = $('#mapelBaru').value.trim(); if (!n) return toast('Isi nama mata pelajaran');
  S.mapel.push({ id: uid(), nama: n, kat: $('#mapelKat').value, pem: $('#mapelPem').value, guru: { nama: '' } }); save(); render(); toast('Mata pelajaran ditambahkan');
};
A.pasteMapel = () => openModal('Tempel daftar mata pelajaran',
  `<p class="small muted">Satu nama per baris. Awali <b>*</b> untuk muatan lokal; tambahkan <code>;IPA</code> atau <code>;IPS</code> untuk mapel pilihan. Contoh:<br><code>Matematika</code><br><code>Kimia;IPA</code><br><code>Ekonomi;IPS</code><br><code>*Bahasa Daerah</code></p><textarea id="pasteBox" style="min-height:220px"></textarea>`,
  `<button class="btn ghost" data-act="closeModal">Batal</button><button class="btn ok" data-act="doPasteMapel">Tambahkan</button>`);
A.doPasteMapel = () => {
  const L = $('#pasteBox').value.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  L.forEach(l => {
    const mulok = l.startsWith('*'); const [nama, pem] = l.replace(/^\*/, '').split(/[;|\t]/).map(x => x.trim());
    S.mapel.push({ id: uid(), nama, kat: mulok ? 'mulok' : 'umum', pem: /^IP[AS]$/i.test(pem || '') ? pem.toUpperCase() : 'semua', guru: { nama: '' } });
  });
  closeModal(); save(); render(); toast(`${L.length} mata pelajaran ditambahkan`);
};
A.kelolaMapel = b => { UI.mid = S.mapel[+b.dataset.i].id; UI.tp = ''; location.hash = '#/tp-lm'; };
const swapMapel = (i, j) => { if (j < 0 || j >= S.mapel.length) return; [S.mapel[i], S.mapel[j]] = [S.mapel[j], S.mapel[i]]; save(); render(); };
A.mapelUp = b => swapMapel(+b.dataset.i, +b.dataset.i - 1);
A.mapelDown = b => swapMapel(+b.dataset.i, +b.dataset.i + 1);
A.delMapel = b => {
  const m = S.mapel[+b.dataset.i];
  if (!confirm(`Hapus mata pelajaran "${m.nama}" beserta TP dan nilainya (semua semester)?`)) return;
  Object.values(S.sems).forEach(d => delete d.m[m.id]); S.mapel.splice(+b.dataset.i, 1); save(); render();
};

/* ---------------- ABSEN & CATATAN WALI KELAS ---------------- */
VIEW.absen = () => !S.siswa.length ? noSiswa('Absen') : panel('Rekap Ketidakhadiran & Catatan Wali Kelas', 'fa-calendar-check', `
  <div class="row" style="margin-bottom:12px"><span class="pill">${SEM[S.profil.sem]}</span>
    <label class="f" style="margin:0">Jumlah hari efektif</label>
    <input type="number" min="0" style="width:110px" value="${esc(D().hari)}" data-h="hari"></div>
  <div class="tw"><table class="grid" id="absTable"><thead><tr><th>No</th><th class="stick">Nama Peserta Didik</th><th>Sakit</th><th>Izin</th><th>Alpa</th><th>Jumlah</th><th>% Hadir</th><th>Catatan Wali Kelas</th></tr></thead>
  <tbody>${S.siswa.map((s, i) => { const a = D().absen[s.id] || {}; const c = absCalc(s.id); return `<tr data-sid="${s.id}">
    <td class="c">${i + 1}</td><td class="nm stick">${esc(s.nama)}</td>
    ${['S', 'I', 'A'].map(k => `<td class="c"><input type="number" min="0" class="n" value="${esc(a[k] ?? '')}" data-h="abs" data-s="${s.id}" data-f="${k}"></td>`).join('')}
    <td class="c ctot">${c.total}</td><td class="c cpct num-strong">${c.pct == null ? '' : Math.round(c.pct) + '%'}</td>
    <td class="wrap"><textarea style="min-height:36px" data-h="catat" data-s="${s.id}">${esc(D().catatan[s.id] || '')}</textarea></td></tr>`; }).join('')}</tbody></table></div>`);
H.hari = el => { D().hari = el.value; save(); S.siswa.forEach(s => refreshAbs(s.id)); };
function refreshAbs(sid) {
  const tr = $(`#absTable tr[data-sid="${sid}"]`); if (!tr) return; const c = absCalc(sid);
  tr.querySelector('.ctot').textContent = c.total; tr.querySelector('.cpct').textContent = c.pct == null ? '' : Math.round(c.pct) + '%';
}
H.abs = el => { ((D().absen[el.dataset.s] ||= {}))[el.dataset.f] = el.value; save(); refreshAbs(el.dataset.s); };
H.catat = el => { D().catatan[el.dataset.s] = el.value; save(); };

/* ---------------- STATUS PENGISIAN MAPEL (dashboard) ---------------- */
function mapelStatus(m) {
  const md = MD(m.id), SM = siswaMapel(m.id), n = SM.length; let a = 0, b = 0;
  SM.forEach(s => { const c = calcSiswa(m.id, s.id); if (c.akhir != null) a++; if (c.sts != null && c.sas != null) b++; });
  const used = md.tp.filter(tp => (tp.deskripsi || '').trim() || SM.some(s => tpCalc(md, tp, s.id)));
  const tpF = used.filter(tp => SM.some(s => tpCalc(md, tp, s.id))).length;
  // sudah = semua siswa punya nilai akhir, STS & SAS lengkap, dan setiap TP terpakai sudah ada nilainya
  const st = a === 0 ? 'belum' : (n > 0 && a === n && b === n && tpF === used.length ? 'sudah' : 'sebagian');
  return { n, a, b, tpF, tpU: used.length, hadir: md.hadir.pertemuan.length, st };
}
const ST_LABEL = { sudah: 'Sudah diisi', sebagian: 'Sebagian', belum: 'Belum diisi' };
function mapelStatusPanel() {
  const list = S.mapel;
  const st = list.map(m => ({ m, ...mapelStatus(m) }));
  const cnt = k => st.filter(x => x.st === k).length;
  const f = UI.statFilter || 'semua';
  const chip = (k, t) => `<button class="chip ${k} ${f === k ? 'on' : ''}" data-act="statFilter" data-f="${k}">${t} <b>${k === 'semua' ? st.length : cnt(k)}</b></button>`;
  const rows = st.filter(x => f === 'semua' || x.st === f);
  return panel(`Status Pengisian Mata Pelajaran — ${SEM[S.profil.sem]}`, 'fa-list-check', `
    <div class="chips">${chip('semua', 'Semua')}${chip('sudah', 'Sudah diisi')}${chip('sebagian', 'Sebagian')}${chip('belum', 'Belum diisi')}</div>
    <div class="tw"><table class="grid"><thead><tr><th>No</th><th>Mata Pelajaran</th><th>Guru Pengampu</th><th>Status</th><th>Siswa bernilai</th><th>TP terisi</th><th>STS & SAS</th><th>Daftar hadir</th><th></th></tr></thead>
    <tbody>${rows.map((x, i) => `<tr><td class="c">${i + 1}</td><td class="nm" style="white-space:normal">${esc(x.m.nama)}${x.m.kat === 'mulok' ? ' <span class="pill">Mulok</span>' : ''}${x.m.pem && x.m.pem !== 'semua' ? ` <span class="pill pem">${x.m.pem}</span>` : ''}</td>
      <td>${esc((x.m.guru || {}).nama || '—')}</td>
      <td class="c"><span class="stbadge ${x.st}">${ST_LABEL[x.st]}</span></td>
      <td class="c">${x.a} / ${x.n}</td><td class="c">${x.tpF} / ${x.tpU}</td><td class="c">${x.b} / ${x.n}</td>
      <td class="c">${x.hadir ? x.hadir + ' pertemuan' : '<span class="muted">belum</span>'}</td>
      <td class="c"><button class="btn sm ghost" data-act="bukaMapel" data-m="${x.m.id}"><i class="fa-solid fa-arrow-right"></i>Buka</button></td></tr>`).join('')
      || '<tr><td colspan="9" class="empty">Tidak ada mata pelajaran pada kategori ini.</td></tr>'}</tbody></table></div>
    <p class="small muted" style="margin:8px 0 0"><b>Sudah diisi</b> = semua siswa punya nilai akhir, STS & SAS lengkap, dan setiap TP yang dipakai sudah ada nilainya. <b>Belum diisi</b> = belum ada satu pun nilai.</p>`);
}
A.statFilter = b => { UI.statFilter = b.dataset.f; render(); };
A.bukaMapel = b => { UI.mid = b.dataset.m; UI.tp = ''; location.hash = '#/nilai-tp'; };

/* ---------------- DAFTAR HADIR PER MAPEL ---------------- */
function hadirCalc(md, sid) {
  const st = md.hadir.status[sid] || {}; const c = { S: 0, I: 0, A: 0, H: 0 };
  md.hadir.pertemuan.forEach(p => { const v = st[p.id]; if (c[v] != null) c[v]++; });
  const tm = num(md.hadir.tm) || md.hadir.pertemuan.length;
  c.pct = tm ? c.H / tm * 100 : null; return c;
}
VIEW.hadir = () => {
  if (!S.mapel.length) return noMapel('Daftar Hadir Mapel');
  if (!S.siswa.length) return noSiswa('Daftar Hadir Mapel');
  const md = cur(), h = md.hadir, P = h.pertemuan;
  const groups = []; P.forEach(p => { const g = groups[groups.length - 1]; g && g.b === p.bulan ? g.n++ : groups.push({ b: p.bulan, n: 1 }); });
  return mapelBar() + panel(`Daftar Hadir — ${esc(mapelOf(UI.mid).nama)}`, 'fa-clipboard-user', `
    <div class="hint info">Daftar hadir ini khusus pelajaran <b>${esc(mapelOf(UI.mid).nama)}</b> (terpisah per mapel dan per semester). Ketik <b>H</b> hadir, <b>S</b> sakit, <b>I</b> izin, <b>A</b> alpa.
      % hadir = jumlah H ÷ jumlah tatap muka.</div>
    <div class="row" style="margin-bottom:12px">
      <select id="pBulan" style="width:150px">${opts(BULAN, BULAN[new Date().getMonth()])}</select>
      <input type="number" id="pTgl" min="1" max="31" placeholder="Tanggal" style="width:110px">
      <button class="btn" data-act="addHadir"><i class="fa-solid fa-plus"></i>Tambah pertemuan</button>
      <span class="spacer"></span><label class="f" style="margin:0">Jumlah tatap muka</label>
      <input type="number" min="0" style="width:90px" value="${esc(h.tm)}" placeholder="${P.length}" data-h="htm">
    </div>
    <div class="tw"><table class="grid" id="hadirTable"><thead><tr><th rowspan="2">No</th><th rowspan="2" class="stick">Nama Peserta Didik</th>
      ${groups.map(g => `<th colspan="${g.n}">${g.b}</th>`).join('')}<th colspan="4">Jumlah</th><th rowspan="2">% Hadir</th></tr>
      <tr class="sub">${P.map(p => `<th>${esc(p.tgl)}<br><button class="btn sm ghost icon" title="Semua hadir" data-act="allHadir" data-p="${p.id}"><i class="fa-solid fa-check"></i></button>
        <button class="btn sm danger icon" title="Hapus pertemuan" data-act="delHadir" data-p="${p.id}"><i class="fa-solid fa-xmark"></i></button></th>`).join('')}
        <th>S</th><th>I</th><th>A</th><th>H</th></tr></thead>
    <tbody>${siswaMapel(UI.mid).map((s, i) => { const st = h.status[s.id] || {}; const c = hadirCalc(md, s.id); return `<tr data-sid="${s.id}">
      <td class="c">${i + 1}</td><td class="nm stick">${esc(s.nama)}</td>
      ${P.map(p => `<td class="c"><input class="abs" maxlength="1" value="${st[p.id] || ''}" data-h="hdr" data-s="${s.id}" data-p="${p.id}"></td>`).join('')}
      <td class="c cS">${c.S}</td><td class="c cI">${c.I}</td><td class="c cA">${c.A}</td><td class="c cH">${c.H}</td>
      <td class="c cP num-strong">${c.pct == null ? '' : Math.round(c.pct) + '%'}</td></tr>`; }).join('')}</tbody></table></div>`);
};
function refreshHadirRow(sid) {
  const tr = $(`#hadirTable tr[data-sid="${sid}"]`); if (!tr) return; const c = hadirCalc(cur(), sid);
  ['S', 'I', 'A', 'H'].forEach(k => tr.querySelector('.c' + k).textContent = c[k]);
  tr.querySelector('.cP').textContent = c.pct == null ? '' : Math.round(c.pct) + '%';
}
H.hdr = el => {
  let v = el.value.toUpperCase(); if (!'HSIA'.includes(v)) v = ''; el.value = v;
  const o = (cur().hadir.status[el.dataset.s] ||= {});
  v ? o[el.dataset.p] = v : delete o[el.dataset.p];
  save(); refreshHadirRow(el.dataset.s);
  if (v) el.closest('td').nextElementSibling?.querySelector('input.abs')?.focus();
};
H.htm = el => { cur().hadir.tm = el.value; save(); siswaMapel(UI.mid).forEach(s => refreshHadirRow(s.id)); };
A.addHadir = () => {
  const t = $('#pTgl').value.trim(); if (!t) return toast('Isi tanggal pertemuan');
  const P = cur().hadir.pertemuan; P.push({ id: uid(), bulan: $('#pBulan').value, tgl: t.padStart(2, '0') });
  const mo = [6, 7, 8, 9, 10, 11, 0, 1, 2, 3, 4, 5];            // urutan tahun ajaran: Juli … Juni
  P.sort((a, b) => mo.indexOf(BULAN.indexOf(a.bulan)) - mo.indexOf(BULAN.indexOf(b.bulan)) || a.tgl - b.tgl);
  save(); render();
};
A.delHadir = b => {
  if (!confirm('Hapus pertemuan ini?')) return; const h = cur().hadir;
  h.pertemuan = h.pertemuan.filter(p => p.id !== b.dataset.p); Object.values(h.status).forEach(o => delete o[b.dataset.p]); save(); render();
};
A.allHadir = b => { const h = cur().hadir; siswaMapel(UI.mid).forEach(s => ((h.status[s.id] ||= {})[b.dataset.p] = 'H')); save(); render(); };

/* ---------------- TUJUAN PEMBELAJARAN: LM & TP ---------------- */
VIEW['tp-lm'] = () => {
  if (!S.mapel.length) return noMapel('Tujuan Pembelajaran');
  const md = cur();
  return `${mapelBar()}
  ${panel('Lingkup Materi', 'fa-folder-tree', `
    <div class="form-grid">${md.lm.map((l, i) => `
      <div class="row" style="flex-wrap:nowrap"><b style="width:46px">LM ${i + 1}</b>
        <input type="text" value="${esc(l.nama)}" data-h="lm" data-i="${i}">
        <button class="btn sm danger icon" data-act="delLm" data-i="${i}"><i class="fa-solid fa-trash"></i></button></div>`).join('')}
    </div>`, `<button class="btn" data-act="addLm"><i class="fa-solid fa-plus"></i>Tambah LM</button>`)}
  ${panel('Tujuan Pembelajaran (TP)', 'fa-bullseye', `
    <div class="hint info">Deskripsi TP dipakai dalam kalimat capaian kompetensi pada rapor. Tulis dalam bentuk frasa kemampuan, contoh:
      <i>“Menghitung keliling dan luas bangun datar.”</i></div>
    <div class="tw"><table class="grid"><thead><tr><th>Kode TP</th><th>Lingkup Materi</th><th>Deskripsi Tujuan Pembelajaran</th><th></th></tr></thead>
    <tbody>${md.tp.map((t, i) => `<tr>
      <td style="width:110px"><input type="text" value="${esc(t.kode)}" data-h="tp" data-i="${i}" data-f="kode"></td>
      <td style="width:230px"><select data-c="tpLm" data-i="${i}">${opts(md.lm.map((l, k) => [l.id, `LM ${k + 1}`]), t.lm)}</select></td>
      <td><textarea data-h="tp" data-i="${i}" data-f="deskripsi" style="min-height:44px">${esc(t.deskripsi)}</textarea></td>
      <td class="c"><button class="btn sm danger icon" data-act="delTp" data-i="${i}"><i class="fa-solid fa-trash"></i></button></td></tr>`).join('')
      || '<tr><td colspan="4" class="empty">Belum ada TP. Klik “Tambah TP”.</td></tr>'}</tbody></table></div>`,
    `<button class="btn" data-act="addTp"><i class="fa-solid fa-plus"></i>Tambah TP</button>`)}`;
};
H.lm = el => { cur().lm[+el.dataset.i].nama = el.value; save(); };
A.addLm = () => { const md = cur(); md.lm.push({ id: 'LM' + uid(), nama: 'Lingkup Materi ' + (md.lm.length + 1) }); save(); render(); };
A.delLm = b => {
  const md = cur(); const l = md.lm[+b.dataset.i];
  if (md.tp.some(t => t.lm === l.id)) return toast('LM masih dipakai TP. Pindahkan/hapus TP-nya dulu.');
  md.lm.splice(+b.dataset.i, 1); save(); render();
};
A.addTp = () => {
  const md = cur(); const lm = md.lm[md.lm.length - 1]; const li = md.lm.length;
  const same = md.tp.filter(t => t.lm === lm.id).length;
  md.tp.push({ id: uid(), kode: `TP ${li}.${same + 1}`, lm: lm.id, deskripsi: '', models: newModels() });
  save(); render(); setTimeout(() => $$('[data-f=deskripsi]').pop()?.focus(), 30);
};
H.tp = el => { cur().tp[+el.dataset.i][el.dataset.f] = el.value; save(); };
C.tpLm = el => { cur().tp[+el.dataset.i].lm = el.value; save(); };
A.delTp = b => {
  const md = cur(); const t = md.tp[+b.dataset.i]; if (!confirm(`Hapus ${t.kode} beserta nilainya?`)) return;
  delete md.nilai[t.id]; md.tp.splice(+b.dataset.i, 1); save(); render();
};

/* ---------------- RUBRIK PREDIKAT (berlaku untuk semua mapel) ---------------- */
VIEW['tp-rubrik'] = () => panel('Rubrik Predikat (dipakai di deskripsi Leger & Rapor semua mapel)', 'fa-ranking-star', `
  <div class="hint info">Nilai 0–100 diubah menjadi predikat KUDIC <b>K / C / B / SB</b> menurut batas “Nilai mulai dari”. Urutkan dari nilai terkecil.
    Kolom <i>tambahan kalimat</i> boleh kosong; bila diisi akan muncul sebagai “…, masih perlu penguatan”.</div>
  <div class="tw"><table class="grid"><thead><tr><th>Kode</th><th>Nilai mulai dari</th><th>Predikat</th><th>Rentang</th><th>Tambahan kalimat (boleh kosong)</th></tr></thead><tbody>
  ${S.rubrik.map((r, i) => `<tr><td class="c"><span class="badge ${KODE[i]}">${KODE[i]}</span></td>
    <td style="width:130px"><input type="number" value="${esc(r.dari)}" data-h="rub" data-i="${i}" data-f="dari"></td>
    <td><input type="text" value="${esc(r.predikat)}" data-h="rub" data-i="${i}" data-f="predikat"></td>
    <td class="c" id="rg${i}">${rangeOf(KODE[i])}</td>
    <td><input type="text" value="${esc(r.tambahan)}" data-h="rub" data-i="${i}" data-f="tambahan"></td></tr>`).join('')}
  </tbody></table></div>`);
H.rub = el => {
  S.rubrik[+el.dataset.i][el.dataset.f] = el.dataset.f === 'dari' ? num(el.value) ?? 0 : el.value; save();
  S.rubrik.forEach((_, i) => { const c = $('#rg' + i); if (c) c.textContent = rangeOf(KODE[i]); });
};

/* ---------------- DESKRIPSI KUDIC PER TP ---------------- */
VIEW['tp-kudic'] = () => {
  if (!S.mapel.length) return noMapel('Deskripsi KUDIC');
  const md = cur(); const tp = tpById(UI.tp);
  if (!tp) return mapelBar() + panel('Deskripsi KUDIC per TP', 'fa-diagram-project', need(false, 'Mapel ini belum punya TP.', ['tp-lm', 'Ke Lingkup Materi & TP']));
  return mapelBar() + panel('Deskripsi KUDIC per TP', 'fa-diagram-project', `
    <div class="row" style="margin-bottom:12px">
      <select data-c="pickTp" style="max-width:420px">${opts(md.tp.map(t => [t.id, `${t.kode} — ${stripDot(t.deskripsi).slice(0, 60)}`]), UI.tp)}</select>
      <span class="muted small">Salin model dari TP:</span>
      <select id="copyFrom" style="max-width:200px">${opts(md.tp.filter(t => t.id !== UI.tp).map(t => [t.id, t.kode]), '')}</select>
      <button class="btn ghost sm" data-act="copyModels"><i class="fa-solid fa-copy"></i>Salin</button>
    </div>
    <div class="hint info">Untuk setiap model KUDIC (Knowing, Understanding, Doing, Inovasi) tulis <b>indikator</b> dan deskripsi kemampuan siswa untuk tiap predikat.
      Guru tinggal mengetik nilai; predikat & deskripsi muncul otomatis.</div>
    ${tp.models.map((m, mi) => `<div class="kudic-card">
      <h4>Model ${mi + 1}</h4>
      <div class="form-grid" style="grid-template-columns:1fr 2fr">
        <div><label class="f">Nama model</label><input type="text" value="${esc(m.nama)}" data-h="mdl" data-m="${mi}" data-f="nama"></div>
        <div><label class="f">Indikator</label><input type="text" value="${esc(m.ind)}" data-h="mdl" data-m="${mi}" data-f="ind" placeholder="mis. Menyebutkan pengertian konsep"></div>
      </div>
      <div class="kudic-desc">${['SB', 'B', 'C', 'K'].map(k => `<div>
        <label class="f"><span class="badge ${k}">${k}</span> nilai ${rangeOf(k)} — ${esc(S.rubrik[KODE.indexOf(k)].predikat)}</label>
        <textarea data-h="mdl" data-m="${mi}" data-f="desk.${k}">${esc((m.desk || {})[k])}</textarea></div>`).join('')}</div>
    </div>`).join('')}`);
};
H.mdl = el => {
  const m = tpById(UI.tp).models[+el.dataset.m]; const f = el.dataset.f;
  if (f.startsWith('desk.')) (m.desk ||= {})[f.slice(5)] = el.value; else m[f] = el.value; save();
};
A.copyModels = () => {
  const src = tpById($('#copyFrom').value); if (!src) return;
  if (!confirm(`Timpa model KUDIC ${tpById(UI.tp).kode} dengan salinan dari ${src.kode}?`)) return;
  tpById(UI.tp).models = JSON.parse(JSON.stringify(src.models)); save(); render();
};

/* ---------------- EKSTRAKURIKULER ---------------- */
VIEW.ekskul = () => `
  ${panel('Daftar Ekstrakurikuler', 'fa-futbol', `
    <div class="form-grid">${S.ekskul.map((e, i) => `<div class="row" style="flex-wrap:nowrap"><b style="width:30px">${i + 1}.</b>
      <input type="text" value="${esc(e.nama)}" data-h="eksNama" data-i="${i}">
      <button class="btn sm danger icon" data-act="delEks" data-i="${i}"><i class="fa-solid fa-trash"></i></button></div>`).join('')}</div>`,
    `<button class="btn" data-act="addEks"><i class="fa-solid fa-plus"></i>Tambah ekstrakurikuler</button>`)}
  ${!S.siswa.length ? noSiswa('Penilaian Ekstrakurikuler') : panel(`Penilaian Ekstrakurikuler — ${SEM[S.profil.sem]}`, 'fa-heart-pulse', `
    <div class="hint info">Pilih predikat tiap siswa pada kegiatan yang diikuti; kosongkan bila tidak mengikuti. Keterangan di rapor dibuat otomatis
      (“Ananda baik dalam mengikuti kegiatan Pramuka.”).</div>
    <div class="tw"><table class="grid"><thead><tr><th>No</th><th class="stick">Nama</th>${S.ekskul.map(e => `<th>${esc(e.nama)}</th>`).join('')}</tr></thead>
    <tbody>${S.siswa.map((s, i) => `<tr><td class="c">${i + 1}</td><td class="nm stick">${esc(s.nama)}</td>
      ${S.ekskul.map(e => `<td><select data-c="eksVal" data-s="${s.id}" data-e="${e.id}">${opts(PRED_EKS.map(p => [p, p || '—']), (D().eks[s.id] || {})[e.id] || '')}</select></td>`).join('')}</tr>`).join('')}
    </tbody></table></div>`)}`;
H.eksNama = el => { S.ekskul[+el.dataset.i].nama = el.value; save(); };
A.addEks = () => { S.ekskul.push({ id: uid(), nama: 'Ekstrakurikuler ' + (S.ekskul.length + 1) }); save(); render(); };
A.delEks = b => {
  const e = S.ekskul[+b.dataset.i]; if (!confirm(`Hapus "${e.nama}"?`)) return;
  Object.values(S.sems).forEach(d => Object.values(d.eks).forEach(o => delete o[e.id])); S.ekskul.splice(+b.dataset.i, 1); save(); render();
};
C.eksVal = el => { ((D().eks[el.dataset.s] ||= {}))[el.dataset.e] = el.value; save(); };

/* ---------------- INPUT NILAI TP (KUDIC) ---------------- */
VIEW['nilai-tp'] = () => {
  if (!S.mapel.length) return noMapel('Nilai TP (KUDIC)');
  if (!S.siswa.length) return noSiswa('Nilai TP (KUDIC)');
  const md = cur(); const tp = tpById(UI.tp);
  if (!tp) return mapelBar() + panel('Nilai TP (KUDIC)', 'fa-pen-to-square', need(false, 'Mapel ini belum punya TP.', ['tp-lm', 'Tambah TP']));
  return mapelBar() + panel('Nilai TP (KUDIC)', 'fa-pen-to-square', `
    <div class="row" style="margin-bottom:12px">
      <select data-c="pickTp" style="max-width:560px">${opts(md.tp.map(t => [t.id, `${t.kode} — ${stripDot(t.deskripsi).slice(0, 70)}`]), UI.tp)}</select>
    </div>
    <div class="hint info">Ketik nilai <b>0–100</b> untuk tiap model. Predikat (SB/B/C/K) dan Nilai TP (rata-rata) terisi otomatis dan masuk ke Rekap Nilai & Rapor.</div>
    <div class="tw"><table class="grid" id="nilaiTable"><thead><tr><th>No</th><th class="stick">Nama</th>
      ${tp.models.map(m => `<th>${esc(short(m.nama))}</th>`).join('')}<th>Nilai TP</th><th>Tertinggi</th><th>Terendah</th><th>Catatan guru</th><th></th></tr></thead>
    <tbody>${siswaMapel(UI.mid).map((s, i) => nilaiRow(md, s, i)).join('')}</tbody></table></div>`);
};
function nilaiRow(md, s, i) {
  const rec = ((md.nilai[UI.tp] || {})[s.id]) || { n: [], cat: '' };
  return `<tr data-sid="${s.id}"><td class="c">${i + 1}</td><td class="nm stick">${esc(s.nama)}</td>
    ${[0, 1, 2, 3].map(k => `<td class="c"><input type="number" class="n" min="0" max="100" value="${esc((rec.n || [])[k] ?? '')}" data-h="nilai" data-s="${s.id}" data-k="${k}"><br><span class="b${k}"></span></td>`).join('')}
    <td class="c num-strong cavg"></td><td class="small chi"></td><td class="small clo"></td>
    <td class="wrap"><textarea style="min-height:36px" data-h="cat" data-s="${s.id}">${esc(rec.cat || '')}</textarea></td>
    <td class="c"><button class="btn sm ghost icon" title="Lihat deskripsi otomatis" data-act="showDesc" data-s="${s.id}"><i class="fa-solid fa-eye"></i></button></td></tr>`;
}
function refreshNilaiRow(sid) {
  const md = cur(); const tp = tpById(UI.tp); const tr = $(`#nilaiTable tr[data-sid="${sid}"]`); if (!tr) return;
  const c = tpCalc(md, tp, sid);
  [0, 1, 2, 3].forEach(k => {
    const v = c ? c.vals[k] : null;
    tr.querySelector('.b' + k).innerHTML = v == null ? '' : `<span class="badge ${kodeOf(v)}">${kodeOf(v)}</span>`;
  });
  tr.querySelector('.cavg').textContent = c ? c.avg : '';
  tr.querySelector('.chi').textContent = c ? short(tp.models[c.iMax].nama) : '';
  tr.querySelector('.clo').textContent = c && c.iMin >= 0 ? short(tp.models[c.iMin].nama) : '';
}
AFTER['nilai-tp'] = () => { if (tpById(UI.tp)) siswaMapel(UI.mid).forEach(s => refreshNilaiRow(s.id)); };
const nilaiRec = sid => ((cur().nilai[UI.tp] ||= {})[sid] ||= { n: ['', '', '', ''], cat: '' });
H.nilai = el => {
  let v = el.value; if (v !== '' && (num(v) == null || num(v) < 0 || num(v) > 100)) { v = ''; el.value = ''; }
  nilaiRec(el.dataset.s).n[+el.dataset.k] = v; save(); refreshNilaiRow(el.dataset.s);
};
H.cat = el => { nilaiRec(el.dataset.s).cat = el.value; save(); };
A.showDesc = b => {
  const tp = tpById(UI.tp); const c = tpCalc(cur(), tp, b.dataset.s);
  openModal(`${esc(namaOf(b.dataset.s))} — ${esc(tp.kode)}`, c ? `
    <p><b>Capaian tertinggi</b><br>${esc(c.T)}</p>
    <p><b>Capaian terendah</b><br>${c.U ? esc(c.U) : '<i class="muted">Tidak ada (nilai semua model sama / hanya satu nilai).</i>'}</p>` :
    '<p class="muted">Belum ada nilai.</p>');
};

/* ---------------- INPUT SUMATIF ---------------- */
VIEW['nilai-sum'] = () => {
  if (!S.mapel.length) return noMapel('Sumatif');
  if (!S.siswa.length) return noSiswa('Sumatif');
  const md = cur();
  return mapelBar() + panel('Sumatif Lingkup Materi, STS & SAS', 'fa-pen-to-square', `
    <div class="hint info"><b>Nilai Raport</b> = rata-rata dari (Rata-rata Formatif, Rata-rata Sumatif LM, STS, SAS) yang terisi.
      Nilai formatif per TP diambil otomatis dari menu <i>Nilai TP</i>.</div>
    <div class="tw"><table class="grid" id="sumTable"><thead><tr><th>No</th><th class="stick">Nama</th>
      ${md.lm.map((l, i) => `<th title="${esc(l.nama)}">Sumatif LM ${i + 1}</th>`).join('')}<th>Rata-rata SLM</th><th>STS</th><th>SAS</th><th>Rata Formatif</th><th>Nilai Raport</th></tr></thead>
    <tbody>${siswaMapel(UI.mid).map((s, i) => { const x = md.sumatif[s.id] || { lm: {} }; return `<tr data-sid="${s.id}"><td class="c">${i + 1}</td><td class="nm stick">${esc(s.nama)}</td>
      ${md.lm.map(l => `<td class="c"><input type="number" class="n" min="0" max="100" value="${esc((x.lm || {})[l.id] ?? '')}" data-h="sum" data-s="${s.id}" data-f="lm.${l.id}"></td>`).join('')}
      <td class="c cslm"></td>
      <td class="c"><input type="number" class="n" min="0" max="100" value="${esc(x.sts ?? '')}" data-h="sum" data-s="${s.id}" data-f="sts"></td>
      <td class="c"><input type="number" class="n" min="0" max="100" value="${esc(x.sas ?? '')}" data-h="sum" data-s="${s.id}" data-f="sas"></td>
      <td class="c cform"></td><td class="c num-strong cakhir"></td></tr>`; }).join('')}</tbody></table></div>`);
};
function refreshSumRow(sid) {
  const tr = $(`#sumTable tr[data-sid="${sid}"]`); if (!tr) return; const c = calcSiswa(UI.mid, sid);
  tr.querySelector('.cslm').textContent = r0(c.avgLM); tr.querySelector('.cform').textContent = r0(c.avgF); tr.querySelector('.cakhir').textContent = r0(c.akhir);
}
AFTER['nilai-sum'] = () => siswaMapel(UI.mid).forEach(s => refreshSumRow(s.id));
H.sum = el => {
  let v = el.value; if (v !== '' && (num(v) == null || num(v) < 0 || num(v) > 100)) { v = ''; el.value = ''; }
  const x = (cur().sumatif[el.dataset.s] ||= { lm: {}, sts: '', sas: '' }); x.lm ||= {};
  const f = el.dataset.f; f.startsWith('lm.') ? x.lm[f.slice(3)] = v : x[f] = v;
  save(); refreshSumRow(el.dataset.s);
};

/* ---------------- REKAP NILAI (per mapel) ---------------- */
VIEW.rekap = () => {
  if (!S.mapel.length) return noMapel('Rekap Nilai');
  if (!S.siswa.length) return noSiswa('Rekap Nilai');
  const md = cur(); const SM = siswaMapel(UI.mid); const cs = SM.map(s => calcSiswa(UI.mid, s.id));
  const tps = md.tp.filter(t => UI.allTp || cs.some(c => c.form[t.id]));
  return mapelBar() + panel(`Rekap Nilai Asesmen — ${esc(mapelOf(UI.mid).nama)}`, 'fa-layer-group', `
    <div class="row" style="margin-bottom:10px"><label class="row small"><input type="checkbox" ${UI.allTp ? 'checked' : ''} data-c="allTp"> Tampilkan semua TP (termasuk yang belum ada nilai)</label></div>
    <div class="tw" style="max-height:72vh"><table class="grid"><thead>
      <tr><th rowspan="2">No</th><th rowspan="2" class="stick">Peserta Didik</th>
        ${tps.length ? `<th colspan="${tps.length}">Formatif (nilai per TP)</th>` : ''}<th rowspan="2">Rata Formatif</th>
        <th colspan="${md.lm.length}">Sumatif LM</th><th rowspan="2">Rata SLM</th><th rowspan="2">STS</th><th rowspan="2">SAS</th>
        <th rowspan="2">Nilai Raport</th><th rowspan="2">Predikat</th><th rowspan="2">TP Tertinggi</th><th rowspan="2">TP Terendah</th>
        <th rowspan="2">Deskripsi Kompetensi Tertinggi</th><th rowspan="2">Deskripsi Kompetensi Terendah</th></tr>
      <tr class="sub">${tps.map(t => `<th title="${esc(t.deskripsi)}">${esc(t.kode)}</th>`).join('')}${md.lm.map((l, i) => `<th>LM ${i + 1}</th>`).join('')}</tr></thead>
    <tbody>${SM.map((s, i) => { const c = cs[i]; return `<tr><td class="c">${i + 1}</td><td class="nm stick">${esc(s.nama)}</td>
      ${tps.map(t => `<td class="c">${c.form[t.id] ? c.form[t.id].avg : ''}</td>`).join('')}
      <td class="c">${r0(c.avgF)}</td>${md.lm.map(l => `<td class="c">${c.lmVals[l.id] ?? ''}</td>`).join('')}
      <td class="c">${r0(c.avgLM)}</td><td class="c">${c.sts ?? ''}</td><td class="c">${c.sas ?? ''}</td>
      <td class="c num-strong">${r0(c.akhir)}</td><td class="c">${c.predikat ? `<span class="pill">${esc(c.predikat)}</span>` : ''}</td>
      <td class="c">${c.hi ? esc(c.hi.kode) : ''}</td><td class="c">${c.lo ? esc(c.lo.kode) : ''}</td>
      <td class="wrap small">${esc(c.descHi)}</td><td class="wrap small">${esc(c.descLo)}</td></tr>`; }).join('')}</tbody></table></div>`,
    `<button class="btn ghost" data-act="exportRekap"><i class="fa-solid fa-file-excel"></i>Ekspor Excel</button>`);
};
C.allTp = el => { UI.allTp = el.checked; render(); };

/* ---------------- KOKURIKULER (Profil Lulusan) ---------------- */
const kokuBar = () => `<div class="row mapelbar"><b><i class="fa-solid fa-puzzle-piece"></i> Kegiatan kokurikuler:</b>
  <select data-c="pickKoku" style="max-width:420px">${opts(D().koku.map(k => [k.id, k.tema || '(tanpa tema)']), UI.kid)}</select>
  <span class="pill">${SEM[S.profil.sem]}</span></div>`;
C.pickKoku = el => { UI.kid = el.value; render(); };
const curKoku = () => D().koku.find(k => k.id === UI.kid);

VIEW['koku-tema'] = () => `
  ${panel(`Kegiatan Kokurikuler — ${SEM[S.profil.sem]}`, 'fa-puzzle-piece', `
    <div class="hint info">Tiap kegiatan/tema kokurikuler punya aspek yang diamati pada 8 dimensi Profil Lulusan (maks. 3 aspek per dimensi).
      Dimensi yang dikosongkan tidak mengurangi nilai siswa.</div>
    ${D().koku.map((k, ki) => `<div class="kudic-card">
      <div class="row" style="margin-bottom:8px"><b>Kegiatan ${ki + 1}</b>
        <input type="text" value="${esc(k.tema)}" data-h="kTema" data-k="${k.id}" placeholder="Tema kokurikuler, mis. Aku Cinta Indonesia" style="max-width:480px">
        <button class="btn sm danger icon" data-act="delKoku" data-k="${k.id}"><i class="fa-solid fa-trash"></i></button></div>
      <div class="tw"><table class="grid"><thead><tr><th>Dimensi</th><th>Aspek 1</th><th>Aspek 2</th><th>Aspek 3</th></tr></thead><tbody>
      ${k.dims.map((d, di) => `<tr><td style="min-width:200px;white-space:normal"><input type="text" value="${esc(d.nama)}" data-h="kDim" data-k="${k.id}" data-d="${di}"></td>
        ${[0, 1, 2].map(x => `<td class="wrap"><textarea style="min-height:44px" data-h="kAspek" data-k="${k.id}" data-d="${di}" data-x="${x}">${esc(d.aspek[x] || '')}</textarea></td>`).join('')}</tr>`).join('')}
      </tbody></table></div></div>`).join('')}`,
    `<button class="btn" data-act="addKoku"><i class="fa-solid fa-plus"></i>Tambah kegiatan</button>`)}`;
const kokuById = id => D().koku.find(k => k.id === id);
H.kTema = el => { kokuById(el.dataset.k).tema = el.value; save(); };
H.kDim = el => { kokuById(el.dataset.k).dims[+el.dataset.d].nama = el.value; save(); };
H.kAspek = el => { kokuById(el.dataset.k).dims[+el.dataset.d].aspek[+el.dataset.x] = el.value; save(); };
A.addKoku = () => { D().koku.push(newKoku('Tema Kokurikuler ' + (D().koku.length + 1))); save(); render(); };
A.delKoku = b => { if (!confirm('Hapus kegiatan kokurikuler ini beserta skornya?')) return; D().koku = D().koku.filter(k => k.id !== b.dataset.k); save(); render(); };

VIEW['koku-nilai'] = () => {
  const k = curKoku(); if (!k) return panel('Input Skor Kokurikuler', 'fa-puzzle-piece', need(false, 'Belum ada kegiatan kokurikuler.', ['koku-tema', 'Tambah kegiatan']));
  if (!S.siswa.length) return noSiswa('Input Skor Kokurikuler');
  const cols = []; k.dims.forEach((d, di) => [0, 1, 2].forEach(x => { if ((d.aspek[x] || '').trim()) cols.push({ di, x, t: d.aspek[x] }); }));
  return kokuBar() + panel(`Skor Profil Lulusan — ${esc(k.tema)}`, 'fa-user-graduate', !cols.length ?
    need(false, 'Belum ada aspek yang diisi.', ['koku-tema', 'Isi aspek']) : `
    <div class="hint info">Isi skor <b>1–4</b> (4 Konsisten · 3 Tampak · 2 Mulai tampak · 1 Belum tampak). Arahkan kursor ke judul kolom untuk membaca aspeknya.</div>
    <div class="tw" style="max-height:70vh"><table class="grid" id="kokuTable"><thead><tr><th rowspan="2">No</th><th rowspan="2" class="stick">Nama</th>
      ${k.dims.map((d, di) => { const n = cols.filter(c => c.di === di).length; return n ? `<th colspan="${n}">${esc(d.nama)}</th>` : ''; }).join('')}
      <th rowspan="2">Rata-rata</th><th rowspan="2">Predikat</th><th rowspan="2">Paling menonjol</th><th rowspan="2">Perlu penguatan</th></tr>
      <tr class="sub">${cols.map(c => `<th title="${esc(c.t)}">${c.x + 1}</th>`).join('')}</tr></thead>
    <tbody>${S.siswa.map((s, i) => { const sk = k.skor[s.id] || []; return `<tr data-sid="${s.id}"><td class="c">${i + 1}</td><td class="nm stick">${esc(s.nama)}</td>
      ${cols.map(c => `<td class="c"><input type="number" class="n" style="width:50px" min="1" max="4" value="${esc(sk[c.di * 3 + c.x] ?? '')}" data-h="skor" data-s="${s.id}" data-i="${c.di * 3 + c.x}"></td>`).join('')}
      <td class="c num-strong kavg"></td><td class="c kpred"></td><td class="small khi"></td><td class="small klo"></td></tr>`; }).join('')}</tbody></table></div>`);
};
function refreshKokuRow(sid) {
  const tr = $(`#kokuTable tr[data-sid="${sid}"]`); if (!tr) return; const c = calcKoku(curKoku(), sid);
  tr.querySelector('.kavg').textContent = r0(c.avg); tr.querySelector('.kpred').innerHTML = c.pred ? `<span class="pill">${esc(c.pred)}</span>` : '';
  tr.querySelector('.khi').textContent = c.hi || ''; tr.querySelector('.klo').textContent = c.lo || '';
}
AFTER['koku-nilai'] = () => { if (curKoku()) S.siswa.forEach(s => refreshKokuRow(s.id)); };
H.skor = el => {
  let v = el.value; if (v !== '' && !['1', '2', '3', '4'].includes(v)) { v = ''; el.value = ''; }
  const a = (curKoku().skor[el.dataset.s] ||= Array(24).fill('')); a[+el.dataset.i] = v; save(); refreshKokuRow(el.dataset.s);
};
VIEW['koku-pred'] = () => panel('Arti Skor & Predikat Karakter', 'fa-ranking-star', `
  <div class="hint info"><b>Skor:</b> 4 = Konsisten · 3 = Tampak · 2 = Mulai tampak · 1 = Belum tampak.<br>
    Nilai dimensi = rata-rata skor ÷ 4 × 100. Predikat ditentukan dari rata-rata semua dimensi menurut tabel berikut (boleh diubah).</div>
  <div class="tw"><table class="grid"><thead><tr><th>Nilai mulai dari</th><th>Predikat</th><th>Artinya</th></tr></thead><tbody>
  ${S.kokuPred.map((p, i) => `<tr><td style="width:120px"><input type="number" value="${esc(p.dari)}" data-h="kpred" data-i="${i}" data-f="dari"></td>
    <td style="width:220px"><input type="text" value="${esc(p.nama)}" data-h="kpred" data-i="${i}" data-f="nama"></td>
    <td><input type="text" value="${esc(p.arti)}" data-h="kpred" data-i="${i}" data-f="arti"></td></tr>`).join('')}</tbody></table></div>`);
H.kpred = el => { S.kokuPred[+el.dataset.i][el.dataset.f] = el.dataset.f === 'dari' ? num(el.value) ?? 0 : el.value; save(); };

/* =====================================================================
   7. SAMPUL, RAPOR & LEGER (cetak)
   ===================================================================== */
function setPage(css) { let st = $('#pageStyle'); if (!st) { st = document.createElement('style'); st.id = 'pageStyle'; document.head.appendChild(st); } st.textContent = css; }
const PAGE_PORTRAIT = '@page{size:216mm 330mm;margin:0}';
function printNodes(html, css) {
  setPage(css || PAGE_PORTRAIT);
  $('#printAll').innerHTML = html; document.body.classList.add('print-all'); window.print();
  setTimeout(() => { document.body.classList.remove('print-all'); $('#printAll').innerHTML = ''; }, 500);
}
const sidPicker = (extra = '') => `<div class="panel noprint"><div class="panel-b row">
  <b>Siswa:</b><select data-c="pickSid" style="max-width:340px">${opts(S.siswa.map(s => [s.id, s.nama + (s.pem ? ` (${s.pem})` : '')]), UI.sid)}</select>
  <button class="btn ghost" data-act="prevSid"><i class="fa-solid fa-chevron-left"></i></button>
  <button class="btn ghost" data-act="nextSid"><i class="fa-solid fa-chevron-right"></i></button>
  <div class="spacer"></div>${extra}</div></div>`;
C.pickSid = el => { UI.sid = el.value; render(); };
const stepSid = d => { const i = S.siswa.findIndex(s => s.id === UI.sid); UI.sid = S.siswa[(i + d + S.siswa.length) % S.siswa.length].id; render(); };
A.prevSid = () => stepSid(-1); A.nextSid = () => stepSid(1);

/* ---- Sampul ---- */
const JENJANG_NAMA = { SD: 'SEKOLAH DASAR', SMP: 'SEKOLAH MENENGAH PERTAMA', SMA: 'SEKOLAH MENENGAH ATAS', SMK: 'SEKOLAH MENENGAH KEJURUAN', MI: 'MADRASAH IBTIDAIYAH', MTs: 'MADRASAH TSANAWIYAH', MA: 'MADRASAH ALIYAH' };
function namaSekolahPenuh() {          // "SMA SWASTA REFORMASI PLUS" → "SEKOLAH MENENGAH ATAS SWASTA REFORMASI PLUS"
  const nama = String(S.profil.sekolah || '').trim().toUpperCase();
  const m = nama.match(/^(SD|SMP|SMA|SMK|MI|MTS|MA)\b\s*(.*)$/);
  const jj = m ? (m[1] === 'MTS' ? 'MTs' : m[1]) : '';
  return m && JENJANG_NAMA[jj] ? (JENJANG_NAMA[jj] + ' ' + m[2]).trim() : (nama || JENJANG_NAMA[S.profil.jenjang] || '');
}
function coverHTML(sid) {
  const p = S.profil, s = siswaOf(sid);
  return `<div class="sheet cover"><div class="cv-frame">
    <div class="cv-top">${p.lambang ? `<img class="cv-lambang" src="${p.lambang}" alt="Lambang">` : '<div class="cv-lambang"></div>'}
      <div class="cv-jenjang">${esc(namaSekolahPenuh())}</div></div>
    <div class="cv-mid"><img class="cv-logoimg" src="${logoSrc()}" alt="Logo Sekolah"></div>
    <div class="cv-who"><div class="cv-lbl">Nama Peserta Didik</div><div class="cv-name">${esc(String(s.nama || '').toUpperCase())}</div>
      <div class="cv-lbl" style="margin-top:12mm">NISN / NIS</div><div class="cv-box">${esc([s.nisn, s.nis].filter(Boolean).join(' / ') || '-')}</div></div>
    <div class="cv-bot">${esc(String(p.yayasan || '').toUpperCase())}</div></div></div>`;
}

/* ---- Identitas Sekolah & Identitas Peserta Didik (dicetak setelah sampul) ---- */
const dash = v => (v == null || String(v).trim() === '') ? '-' : esc(String(v).trim());
const fmtTgl = v => /^\d{4}-\d{2}-\d{2}$/.test(v || '') ? tglID(v) : (v || '');
const stripKab = v => String(v || '').trim().replace(/^(Kab(upaten)?\.?)\s+/i, '').trim();
const stripKec = v => String(v || '').trim().replace(/^(Kec(amatan)?\.?)\s+/i, '').trim();
const provPenuh = v => String(v || '').trim().replace(/^NTT$/i, 'Nusa Tenggara Timur');

function identSekolahHTML() {
  const p = S.profil;
  const row = (l, v) => `<tr><td class="lb">${l}</td><td class="cl">:</td><td>${v}</td></tr>`;
  return `<div class="sheet ident ident-sek"><h2 class="big">IDENTITAS SEKOLAH</h2>
    <table class="idtab big"><tbody>
      ${row('Nama Sekolah', dash(p.sekolah))}${row('NIPD/NISS', dash(p.nss))}${row('N P S N', dash(p.npsn))}
      ${row('Alamat', dash(p.alamat))}
      <tr><td></td><td></td><td>Kode Pos: ${dash(p.kodePos)} &nbsp;&nbsp;&nbsp; Telp. ${dash(p.telp)}</td></tr>
      ${row('Desa', dash(p.desa))}${row('Kecamatan', dash(stripKec(p.kecamatan)))}${row('Kab/Kota', dash(stripKab(p.kabupaten)))}
      ${row('Provinsi', dash(provPenuh(p.provinsi)))}${row('Website', dash(p.website))}
    </tbody></table></div>`;
}

function identSiswaHTML(sid) {
  const s = siswaOf(sid), p = S.profil;
  const r = (no, l, v, sub) => `<tr class="${sub ? 'sub' : ''}"><td class="no">${no}</td><td class="lb">${l}</td><td class="cl">:</td><td>${v === null ? '' : dash(v)}</td></tr>`;
  const ttl = [s.tempatLahir, fmtTgl(s.tglLahir)].filter(Boolean).join(', ');
  const tglTtd = [p.tempatRapor, tglID(p.tglRapor)].filter(Boolean).join(', ');
  return `<div class="sheet ident ident-sis"><h2>IDENTITAS PESERTA DIDIK</h2>
    <table class="idtab"><tbody>
      ${r(1, 'Nama Peserta Didik', s.nama)}${r(2, 'Nomor Induk Siswa (NIS)', s.nis)}${r(3, 'N I S Nasional (NISN)', s.nisn)}
      ${r(4, 'Jenis Kelamin', s.jk)}${r(5, 'Tempat, Tanggal Lahir', ttl)}${r(6, 'Agama', s.agama)}${r(7, 'Status dalam Keluarga', s.statusKel)}
      ${r(8, 'Anak ke', s.anakKe)}${r(9, 'Alamat Peserta Didik', s.alamat)}${r(10, 'Telepon', s.telp)}
      ${r(11, 'Identitas Sekolah Asal', null)}${r('', 'a. Nama Sekolah', s.sekolahAsal, 1)}${r('', 'b. Alamat', s.sekolahAsalAlamat, 1)}
      ${r('', 'c. Tahun Ijazah', s.thnIjazah, 1)}${r('', 'd. Nomor Ijazah', s.noIjazah, 1)}
      ${r(12, 'Diterima di Sekolah Ini', null)}${r('', 'a. Di Kelas', s.diKelas, 1)}${r('', 'b. Pada Tanggal', fmtTgl(s.tglDiterima), 1)}
      ${r(13, 'Identitas Orang Tua', null)}${r('', 'a. Nama Ayah', s.ayah, 1)}${r('', 'b. Nama Ibu', s.ibu, 1)}${r('', 'c. Alamat Orang Tua', s.alamatOrtu, 1)}
      ${r('', 'd. Telepon Orang Tua', s.telpOrtu, 1)}${r('', 'e. Pekerjaan Orang Tua', s.kerjaOrtu, 1)}
      ${r(14, 'Identitas Wali', null)}${r('', 'a. Nama Wali', s.waliNama, 1)}${r('', 'b. Alamat Wali', s.waliAlamat, 1)}
      ${r('', 'c. Nomor Telepon', s.waliTelp, 1)}${r('', 'd. Pekerjaan', s.waliKerja, 1)}
    </tbody></table>
    <div class="ident-sign">
      <div class="ph">${s.foto ? `<img src="${s.foto}" alt="Foto">` : '<span>Pas foto<br>3 x 4</span>'}</div>
      <div class="sg">${esc(tglTtd)}<br>Kepala Sekolah, <div class="sp"></div><b><u>${esc(p.kepsek)}</u></b>${p.nipKepsek ? '<br>NIP. ' + esc(p.nipKepsek) : ''}</div>
    </div></div>`;
}
const paketSampulHTML = sid => coverHTML(sid) + identSekolahHTML() + identSiswaHTML(sid);

/* ---- editor identitas per siswa ---- */
const IDF = [['tempatLahir', 'Tempat Lahir'], ['tglLahir', 'Tanggal Lahir', 'date'], ['agama', 'Agama'], ['statusKel', 'Status dalam Keluarga'], ['anakKe', 'Anak ke'],
  ['alamat', 'Alamat Peserta Didik'], ['telp', 'Telepon'], ['sekolahAsal', 'Nama Sekolah Asal'], ['sekolahAsalAlamat', 'Alamat Sekolah Asal'], ['thnIjazah', 'Tahun Ijazah'],
  ['noIjazah', 'Nomor Ijazah'], ['diKelas', 'Diterima di Kelas'], ['tglDiterima', 'Diterima pada Tanggal', 'date'], ['ayah', 'Nama Ayah'], ['ibu', 'Nama Ibu'],
  ['alamatOrtu', 'Alamat Orang Tua'], ['telpOrtu', 'Telepon Orang Tua'], ['kerjaOrtu', 'Pekerjaan Orang Tua'], ['waliNama', 'Nama Wali'], ['waliAlamat', 'Alamat Wali'],
  ['waliTelp', 'Nomor Telepon Wali'], ['waliKerja', 'Pekerjaan Wali']];
function openIdent(i) {
  const s = S.siswa[i]; if (!s) return toast('Pilih siswa dulu'); UI.identI = i;
  openModal('Identitas Peserta Didik — ' + esc(s.nama || ''), `
    <div class="row" style="margin-bottom:10px"><img id="identPrev" src="${s.foto || ''}" alt="" style="height:96px;width:72px;object-fit:cover;border:1px solid #ccc;background:#f3f4f6">
      <label class="btn ghost"><i class="fa-solid fa-camera"></i>Unggah pas foto<input type="file" accept="image/*" data-c="identFoto" hidden></label>
      <button class="btn ghost" data-act="identFotoDel"><i class="fa-solid fa-xmark"></i>Hapus foto</button></div>
    <div class="form-grid">${IDF.map(([k, l, t]) => `<div><label class="f">${l}</label><input type="${t || 'text'}" id="idf_${k}" value="${esc(s[k] || '')}"></div>`).join('')}</div>`,
    `<button class="btn ghost" data-act="closeModal">Batal</button><button class="btn ok" data-act="saveIdent"><i class="fa-solid fa-floppy-disk"></i>Simpan</button>`);
  $('#modal .box').style.maxWidth = '860px';
}
A.editIdent = b => openIdent(+b.dataset.i);
A.editIdentSid = () => openIdent(S.siswa.findIndex(x => x.id === UI.sid));
A.saveIdent = () => { const s = S.siswa[UI.identI]; IDF.forEach(([k]) => { s[k] = $('#idf_' + k).value.trim(); }); closeModal(); save(); render(); toast('Identitas disimpan'); };
A.identFotoDel = () => { S.siswa[UI.identI].foto = ''; $('#identPrev').removeAttribute('src'); save(); };
C.identFoto = el => {
  const f = el.files[0]; if (!f) return; const rd = new FileReader();
  rd.onload = () => { const img = new Image(); img.onload = () => {            // potong tengah rasio 3:4, lebar 300 px
    const W = 300, H = 400, cv = document.createElement('canvas'); cv.width = W; cv.height = H;
    const k = Math.max(W / img.width, H / img.height), w = img.width * k, h = img.height * k;
    cv.getContext('2d').drawImage(img, (W - w) / 2, (H - h) / 2, w, h);
    S.siswa[UI.identI].foto = cv.toDataURL('image/jpeg', .85); $('#identPrev').src = S.siswa[UI.identI].foto; save(); }; img.src = rd.result; };
  rd.readAsDataURL(f);
};

/* ---- impor identitas dari Excel (kolom dikenali dari judulnya) ---- */
const IDRULES = [
  [/alamat.*sekolah|sekolah.*alamat/, 'sekolahAsalAlamat'], [/tahun.*ijazah|thn.*ijazah/, 'thnIjazah'], [/(nomor|no).*ijazah/, 'noIjazah'],
  [/sekolah asal|asal sekolah|asal smp|nama sekolah/, 'sekolahAsal'],
  [/wali.*(telp|telepon|hp)|(telp|telepon|hp).*wali/, 'waliTelp'], [/wali.*alamat|alamat.*wali/, 'waliAlamat'],
  [/wali.*(pekerjaan|kerja)|(pekerjaan|kerja).*wali/, 'waliKerja'], [/nama wali|^wali$/, 'waliNama'],
  [/(telp|telepon|hp).*(orang tua|ortu)|(orang tua|ortu).*(telp|telepon|hp)/, 'telpOrtu'], [/alamat.*(orang tua|ortu)|(orang tua|ortu).*alamat/, 'alamatOrtu'],
  [/(pekerjaan|kerja).*(orang tua|ortu)|(orang tua|ortu).*(pekerjaan|kerja)|pekerjaan ayah/, 'kerjaOrtu'], [/nama ayah|^ayah$/, 'ayah'], [/nama ibu|^ibu$/, 'ibu'],
  [/nisn|nis nasional/, 'nisn'], [/^(nis|nomor induk siswa|nomor induk siswa nis|no induk)$/, 'nis'],
  [/tempat.*lahir.*(tanggal|tgl)|^ttl$|tempat tanggal lahir/, 'ttl'], [/^tempat lahir$/, 'tempatLahir'], [/(tanggal|tgl) lahir/, 'tglLahir'],
  [/jenis kelamin|^jk$|^l p$|^lp$/, 'jk'], [/agama/, 'agama'], [/status.*keluarga/, 'statusKel'], [/anak ke/, 'anakKe'],
  [/di kelas|diterima.*kelas/, 'diKelas'], [/pada tanggal|(tanggal|tgl).*diterima|diterima.*(tanggal|tgl)/, 'tglDiterima'],
  [/alamat (peserta didik|siswa|rumah)|^alamat$/, 'alamat'], [/^(telepon|telp|no hp|nomor hp|hp|no telp|telepon siswa)$/, 'telp'],
  [/nama (peserta didik|siswa|murid|lengkap)|^nama$/, 'nama']];
const normH = h => String(h || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
function pTgl(v) {
  if (v === '' || v == null) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'number' && window.XLSX) { const d = XLSX.SSF.parse_date_code(v); return d ? `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}` : String(v); }
  const t = String(v).trim(); let m;
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
  if ((m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/))) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  if ((m = t.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/))) { const bi = BULAN.findIndex(b => b.toLowerCase().startsWith(m[2].toLowerCase().slice(0, 3))); if (bi >= 0) return `${m[3]}-${String(bi + 1).padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
  return t;
}
function readIdent(f) {
  const file = f.files[0]; if (!file) return; if (!window.XLSX) return toast('Pustaka Excel belum termuat (cek koneksi internet).');
  const rd = new FileReader();
  rd.onload = () => { try {
    const wb = XLSX.read(rd.result, { type: 'array', cellDates: false });
    const ws = wb.Sheets[wb.SheetNames.find(n => /identitas|siswa|peserta/i.test(n)) || wb.SheetNames[0]];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    let hi = aoa.findIndex((row, i) => i < 20 && row.filter(c => String(c).trim()).length >= 3 && row.some(c => /^nama|nama (peserta|siswa)/.test(normH(c))));
    if (hi < 0) throw new Error('Baris judul kolom (Nama Peserta Didik, dst.) tidak ditemukan');
    const col = {}; aoa[hi].forEach((h, c) => { const n = normH(h); if (!n) return; const rule = IDRULES.find(([re]) => re.test(n)); if (rule && !(rule[1] in col)) col[rule[1]] = c; });
    if (!('nama' in col)) throw new Error('Kolom nama tidak ditemukan');
    const norm = s => normH(s); let upd = 0, add = 0; const newS = [];
    aoa.slice(hi + 1).forEach(row => {
      const g = k => (k in col ? row[col[k]] : ''); const nama = String(g('nama')).trim(); if (!nama) return;
      let s = S.siswa.find(x => (g('nisn') && String(x.nisn) === String(g('nisn')).trim()) || (g('nis') && String(x.nis) === String(g('nis')).trim()) || norm(x.nama) === norm(nama));
      if (!s) { s = { id: uid(), nama, nisn: '', nis: '', jk: '', pem: '' }; newS.push(s); add++; } else upd++;
      const set = (k, v) => { const t = String(v == null ? '' : v).trim(); if (t) s[k] = t; };
      ['nisn', 'nis', 'agama', 'statusKel', 'anakKe', 'alamat', 'telp', 'sekolahAsal', 'sekolahAsalAlamat', 'thnIjazah', 'noIjazah', 'diKelas', 'ayah', 'ibu', 'alamatOrtu', 'telpOrtu', 'kerjaOrtu', 'waliNama', 'waliAlamat', 'waliTelp', 'waliKerja'].forEach(k => set(k, g(k)));
      const jk = String(g('jk')).trim().toUpperCase(); if (jk) s.jk = /^(P|PEREMPUAN|W|WANITA)/.test(jk) ? 'P' : 'L';
      if ('ttl' in col) { const t = String(g('ttl')); const i = t.indexOf(','); if (i > 0) { s.tempatLahir = t.slice(0, i).trim(); s.tglLahir = pTgl(t.slice(i + 1).trim()); } else set('tempatLahir', t); }
      set('tempatLahir', g('tempatLahir')); if (g('tglLahir') !== '') s.tglLahir = pTgl(g('tglLahir')); if (g('tglDiterima') !== '') s.tglDiterima = pTgl(g('tglDiterima'));
    });
    if (!confirm(`Ditemukan ${upd} siswa yang cocok${add ? ` dan ${add} siswa baru` : ''}. Isi identitas dari file ini?`)) return;
    newS.forEach(s => S.siswa.push(s)); save(); render(); toast(`Identitas diimpor: ${upd} diperbarui, ${add} ditambahkan`);
  } catch (err) { console.error(err); toast('Gagal membaca file: ' + err.message); } };
  rd.readAsArrayBuffer(file);
}
A.importIdent = () => { const f = $('#filePick'); f.accept = '.xlsx,.xls'; f.multiple = false; f.onchange = () => readIdent(f); f.value = ''; f.click(); };
A.tplIdent = () => {
  if (!window.XLSX) return toast('Pustaka Excel belum termuat (cek koneksi internet).');
  const head = ['Nama Peserta Didik', 'NIS', 'NISN', 'Jenis Kelamin', 'Tempat Lahir', 'Tanggal Lahir', 'Agama', 'Status dalam Keluarga', 'Anak ke', 'Alamat Peserta Didik', 'Telepon',
    'Nama Sekolah Asal', 'Alamat Sekolah Asal', 'Tahun Ijazah', 'Nomor Ijazah', 'Diterima di Kelas', 'Diterima pada Tanggal', 'Nama Ayah', 'Nama Ibu', 'Alamat Orang Tua',
    'Telepon Orang Tua', 'Pekerjaan Orang Tua', 'Nama Wali', 'Alamat Wali', 'Telepon Wali', 'Pekerjaan Wali'];
  const rows = S.siswa.length ? S.siswa.map(s => [s.nama, s.nis || '', s.nisn || '', s.jk || '']) : [['Contoh Nama Siswa', '2298', '0077482859', 'P', 'Nisum', '18/04/2007', 'Kristen', 'Anak kandung', 3, 'Oebobo', '08xxxxxxxxxx']];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([head, ...rows]), 'Identitas'); XLSX.writeFile(wb, 'template-identitas-siswa.xlsx');
};

VIEW.sampul = () => !S.siswa.length ? noSiswa('Sampul & Identitas') : sidPicker(`
  <button class="btn ghost" data-act="editIdentSid"><i class="fa-solid fa-id-card"></i>Isi identitas siswa ini</button>
  <button class="btn" data-act="printCover"><i class="fa-solid fa-print"></i>Cetak sampul + identitas</button>
  <button class="btn warn" data-act="printCoverAll"><i class="fa-solid fa-print"></i>Cetak semua (${S.siswa.length})</button>`) + paketSampulHTML(UI.sid);
A.printCover = () => printNodes(paketSampulHTML(UI.sid));
A.printCoverAll = () => printNodes(S.siswa.map(s => paketSampulHTML(s.id)).join(''));

/* ---- Rapor ---- */
/* Kop surat: logo di kiri, teks di tengah, garis di bawah */
function kopHTML() {
  const p = S.profil;
  const strip = (v, re) => String(v || '').trim().replace(re, '').trim();
  const kec = strip(p.kecamatan, /^(Kec(amatan)?\.?)\s+/i);
  const kab = strip(p.kabupaten, /^(Kab(upaten)?\.?)\s+/i);
  const prov = String(p.provinsi || '').trim().replace(/^Nusa Tenggara Timur$/i, 'NTT');
  const alamat = [String(p.alamat || '').trim(), kec && 'Kec. ' + kec, kab && 'Kab. ' + kab, prov].filter(Boolean).join(', ');
  return `<div class="kop"><img class="kop-logo" src="${logoSrc()}" alt="Logo"><div class="kop-t">
    ${p.yayasan ? `<div class="y">${esc(p.yayasan)}</div>` : ''}<div class="s">${esc(p.sekolah || 'NAMA SEKOLAH')}</div>
    <div class="a">${esc(alamat)}</div>
    <div class="a">${p.email ? 'Email : ' + esc(p.email) : ''}</div></div></div>`;
}
function raporHTML(sid) {
  const p = S.profil, s = siswaOf(sid), d = D(), a = absCalc(sid);
  const idt = (l, v) => `<tr><td style="width:150px;white-space:nowrap">${l}</td><td style="width:10px">:</td><td>${esc(v)}</td></tr>`;
  const alamat = [p.alamat, p.desa, p.kecamatan, p.kabupaten, p.provinsi].filter(Boolean).join(', ');
  let n = 0;
  const mrow = m => {
    const c = calcSiswa(m.id, sid); n++;
    return `<tr><td class="ct">${n}</td><td>${esc(m.nama)}</td><td class="ct"><b>${r0(c.akhir)}</b></td>
      <td>${c.descHi ? `<p>${esc(c.descHi)}</p>` : ''}${c.descLo ? `<p>${esc(c.descLo)}</p>` : ''}</td></tr>`;
  };
  const MS = mapelSiswa(sid);                    // hanya mapel yang diikuti siswa (sesuai peminatan)
  const isPil = m => m.kat === 'pilihan' || (m.kat !== 'mulok' && m.pem && m.pem !== 'semua');   // mapel pilihan: kategori "Mapel pilihan" (diikuti semua siswa) atau khusus kelompok IPA/IPS
  const umum = MS.filter(m => m.kat !== 'mulok' && !isPil(m)), pilih = MS.filter(isPil), mulok = MS.filter(m => m.kat === 'mulok');
  const kokuRows = d.koku.map((k, i) => { const c = calcKoku(k, sid);
    return `<tr><td class="ct">${i + 1}</td><td>${esc(k.tema)}</td><td class="ct">${esc(c.pred || '')}</td><td>${esc(c.desc || '')}</td></tr>`; }).join('');
  const eksRows = S.ekskul.filter(e => (d.eks[sid] || {})[e.id]).map((e, i) =>
    `<tr><td class="ct">${i + 1}</td><td>${esc(e.nama)}</td><td>${esc(eksKet(d.eks[sid][e.id], e.nama))}</td></tr>`).join('');
  return `<div class="sheet">
    ${kopHTML()}
    <h1>LAPORAN HASIL BELAJAR (RAPOR)</h1>
    <table class="idt" style="margin:10px 0"><tr><td style="width:55%"><table>${idt('Nama Peserta Didik', s.nama)}${idt('NISN / NIS', [s.nisn, s.nis].filter(Boolean).join(' / '))}${idt('Sekolah', p.sekolah)}</table></td>
      <td><table>${idt('Kelas / Fase', `${p.kelas} / ${p.fase}`)}${s.pem ? idt('Mapel Pilihan', s.pem) : ''}${idt('Semester', SEM[p.sem])}${idt('Tahun Pelajaran', p.tahun)}</table></td></tr></table>
    <div class="sec">A. Intrakurikuler</div>
    <table class="bd"><thead><tr><th style="width:30px">No</th><th style="width:150px">Mata Pelajaran</th><th style="width:52px">Nilai Akhir</th><th>Capaian Kompetensi</th></tr></thead><tbody>
      ${pilih.length || mulok.length ? '<tr><td colspan="4"><b>Kelompok Mata Pelajaran Umum</b></td></tr>' : ''}${umum.map(mrow).join('')}
      ${pilih.length ? `<tr><td colspan="4"><b>Mata Pelajaran Pilihan${s.pem && pilih.some(m => m.pem && m.pem !== 'semua') ? ' (' + esc(s.pem) + ')' : ''}</b></td></tr>${pilih.map(mrow).join('')}` : ''}
      ${mulok.length ? `<tr><td colspan="4"><b>Muatan Lokal</b></td></tr>${mulok.map(mrow).join('')}` : ''}</tbody></table>
    <div class="sec">B. Kokurikuler</div>
    <table class="bd"><thead><tr><th style="width:30px">No</th><th style="width:150px">Tema / Kegiatan</th><th style="width:80px">Predikat</th><th>Deskripsi</th></tr></thead><tbody>${kokuRows || '<tr><td colspan="4" class="ct">-</td></tr>'}</tbody></table>
    <div class="sec">C. Ekstrakurikuler</div>
    <table class="bd"><thead><tr><th style="width:30px">No</th><th style="width:200px">Kegiatan Ekstrakurikuler</th><th>Keterangan</th></tr></thead><tbody>${eksRows || '<tr><td colspan="3" class="ct">-</td></tr>'}</tbody></table>
    <div class="keep"><div class="sec">D. Ketidakhadiran</div>
    <table class="bd" style="width:50%"><tbody>
      <tr><td>Sakit</td><td class="ct" style="width:80px">${a.S} hari</td></tr><tr><td>Izin</td><td class="ct">${a.I} hari</td></tr>
      <tr><td>Tanpa keterangan</td><td class="ct">${a.A} hari</td></tr></tbody></table>
    <div class="sec">E. Catatan Wali Kelas</div>
    <table class="bd"><tbody><tr><td style="height:46px">${esc(d.catatan[sid] || '')}</td></tr></tbody></table>
    <div class="sign">
      <div>Mengetahui,<br>Orang Tua/Wali<div class="sp"></div>(............................)</div>
      <div>Mengetahui,<br>Kepala Sekolah<div class="sp"></div><u>${esc(p.kepsek)}</u><br>${p.nipKepsek ? 'NIP. ' + esc(p.nipKepsek) : ''}</div>
      <div>${esc(p.tempatRapor)}${p.tempatRapor && p.tglRapor ? ', ' : ''}${esc(tglID(p.tglRapor))}<br>Wali Kelas<div class="sp"></div><u>${esc(p.guru)}</u><br>${p.nipGuru ? 'NIP. ' + esc(p.nipGuru) : ''}</div>
    </div></div></div>`;
}
/* ---- Rapor kemampuan siswa per TP (alur KUDIC: Knowing → Understanding → Doing → Inovasi) ---- */
const KUDIC = ['Knowing', 'Understanding', 'Doing', 'Inovasi'];
function tpRaporHTML(sid, onlyMid, lampiran) {
  const p = S.profil, s = siswaOf(sid);
  const ms = mapelSiswa(sid).filter(m => !onlyMid || m.id === onlyMid);
  const bdg = n => n == null ? '' : `<span class="bdg ${kodeOf(n)}">${kodeOf(n)}</span>`;
  const blocks = ms.map(m => {
    const md = MD(m.id);
    const rows = md.tp.map(tp => ({ tp, c: tpCalc(md, tp, sid) })).filter(x => x.c);
    if (!rows.length) return '';
    return `<div class="sec">${esc(m.nama)}${m.pem && m.pem !== 'semua' ? ' (Pilihan ' + m.pem + ')' : (m.kat === 'pilihan' ? ' (Pilihan)' : '')}</div>
    <table class="bd tpr"><thead><tr><th style="width:44px">TP</th><th>Tujuan Pembelajaran</th>${KUDIC.map(k => `<th style="width:40px" title="${k}">${k[0]}</th>`).join('')}<th style="width:46px">Nilai TP</th><th>Kemampuan siswa</th></tr></thead><tbody>
    ${rows.map(({ tp, c }) => `<tr><td class="ct">${esc(tp.kode || '')}</td><td>${esc(stripDot(tp.deskripsi))}</td>
      ${c.vals.map(v => `<td class="ct">${v == null ? '–' : r0(v) + '<br>' + bdg(v)}</td>`).join('')}
      <td class="ct"><b>${r0(c.avg)}</b><br>${bdg(c.avg)}</td><td><p>${esc(c.T)}</p>${c.U ? `<p>${esc(c.U)}</p>` : ''}</td></tr>`).join('')}</tbody></table>`;
  }).join('');
  const legend = KODE.slice().reverse().map(k => `<b>${k}</b> = ${esc(S.rubrik[KODE.indexOf(k)].predikat)} (${rangeOf(k)})`).join(' &nbsp;·&nbsp; ');
  return `<div class="sheet${lampiran ? ' lamp' : ''}">
    ${kopHTML()}
    <h1>${lampiran ? 'LAMPIRAN — ' : ''}RAPOR KEMAMPUAN PESERTA DIDIK PER TUJUAN PEMBELAJARAN</h1>
    <div class="cv-sub" style="text-align:center;font-size:10pt;margin-bottom:6px">Alur KUDIC: Knowing → Understanding → Doing → Inovasi</div>
    <table class="idt" style="margin:8px 0"><tr><td style="width:55%"><table><tr><td style="width:130px">Nama</td><td style="width:10px">:</td><td>${esc(s.nama)}</td></tr><tr><td>NISN / NIS</td><td>:</td><td>${esc([s.nisn, s.nis].filter(Boolean).join(' / '))}</td></tr></table></td>
      <td><table><tr><td style="width:110px">Kelas / Fase</td><td style="width:10px">:</td><td>${esc(p.kelas)} / ${esc(p.fase)}</td></tr>${s.pem ? `<tr><td>Mapel Pilihan</td><td>:</td><td>${s.pem}</td></tr>` : ''}<tr><td>Semester</td><td>:</td><td>${SEM[p.sem]} — ${esc(p.tahun)}</td></tr></table></td></tr></table>
    ${blocks || '<p class="ct">Belum ada nilai TP untuk siswa ini.</p>'}
    <div class="keep"><p style="font-size:9pt;margin-top:8px"><b>K</b> = Knowing, <b>U</b> = Understanding, <b>D</b> = Doing, <b>I</b> = Inovasi. Predikat: ${legend}</p>
    <div class="sign"><div></div><div></div><div>${esc(p.tempatRapor)}${p.tempatRapor && p.tglRapor ? ', ' : ''}${esc(tglID(p.tglRapor))}<br>Wali Kelas<div class="sp"></div><u>${esc(p.guru)}</u><br>${p.nipGuru ? 'NIP. ' + esc(p.nipGuru) : ''}</div></div></div></div>`;
}
VIEW['rapor-tp'] = () => {
  if (!S.siswa.length) return noSiswa('Rapor Kemampuan per TP');
  const mm = mapelSiswa(UI.sid);
  if (UI.tpMid && !mm.some(m => m.id === UI.tpMid)) UI.tpMid = '';
  return sidPicker(`<select data-c="pickTpMid" style="max-width:220px">${opts([['', 'Semua mata pelajaran'], ...mm.map(m => [m.id, m.nama])], UI.tpMid)}</select>
    <span class="pill">${SEM[S.profil.sem]}</span>
    <button class="btn" data-act="printTpOne"><i class="fa-solid fa-print"></i>Cetak</button>
    <button class="btn warn" data-act="printTpAll"><i class="fa-solid fa-print"></i>Cetak semua (${S.siswa.length})</button>`) + tpRaporHTML(UI.sid, UI.tpMid);
};
C.pickTpMid = el => { UI.tpMid = el.value; render(); };
A.printTpOne = () => printNodes(tpRaporHTML(UI.sid, UI.tpMid));
A.printTpAll = () => printNodes(S.siswa.map(x => tpRaporHTML(x.id, UI.tpMid)).join(''));
C.lamp = el => { UI.lamp = el.checked; };

const raporPage = () => !S.siswa.length ? noSiswa('Cetak Rapor') : sidPicker(`
  <span class="pill">Semester ${SEM[S.profil.sem]}</span>
  <label class="small"><input type="checkbox" data-c="lamp" ${UI.lamp ? 'checked' : ''}> Sertakan lampiran kemampuan per TP</label>
  <button class="btn" data-act="printOne"><i class="fa-solid fa-print"></i>Cetak rapor ini</button>
  <button class="btn warn" data-act="printAllRapor"><i class="fa-solid fa-print"></i>Cetak semua (${S.siswa.length})</button>`) + raporHTML(UI.sid);
/* Rapor I / II: membuka rapor sekaligus berpindah ke semester terkait */
const setSem = v => { if (S.profil.sem !== v) { S.profil.sem = v; save(); } };
VIEW.rapor1 = () => { setSem('1'); return raporPage(); };
VIEW.rapor2 = () => { setSem('2'); return raporPage(); };
A.printOne = () => printNodes(raporHTML(UI.sid) + (UI.lamp ? tpRaporHTML(UI.sid, '', true) : ''));
A.printAllRapor = () => printNodes(S.siswa.map(s => raporHTML(s.id) + (UI.lamp ? tpRaporHTML(s.id, '', true) : '')).join(''));

/* ---- Leger: data tabel dipakai untuk layar & ekspor Excel ---- */
function legerData(mode) {
  const p = S.profil;
  if (mode === 'akhir') {
    const pf = UI.legerPem || 'semua';
    const rows = S.siswa.filter(s => pf === 'semua' || (s.pem || '') === pf).map(s => ({ s, ...allAkhir(s.id), a: absCalc(s.id) }));
    // peringkat dihitung di dalam kelompok pilihan yang sama (mapel yang diikuti berbeda)
    const rankOf = r => { if (r.jumlah == null) return ''; const g = rows.filter(x => (x.s.pem || '') === (r.s.pem || '') && x.jumlah != null).map(x => x.jumlah).sort((x, y) => y - x); return g.indexOf(r.jumlah) + 1; };
    const head = ['No', 'Nama Peserta Didik', 'Mapel Pilihan', ...S.mapel.map(m => m.nama), 'Jumlah', 'Rata-rata', 'Peringkat (dalam kelompok pilihan)', 'S', 'I', 'A'];
    const body = rows.map((r, i) => [i + 1, r.s.nama, r.s.pem || '-', ...S.mapel.map((m, k) => berlaku(m, r.s.id) ? r0(r.v[k]) : '–'),
      r.jumlah ?? '', r.rata == null ? '' : r.rata.toFixed(1).replace('.', ','), rankOf(r), r.a.S, r.a.I, r.a.A]);
    return { title: `LEGER NILAI AKHIR${pf !== 'semua' ? ' — PILIHAN ' + pf : ''} — ${SEM[p.sem].toUpperCase()}`, head, rows: body };
  }
  const tp = tpById(UI.tp); const md = cur();
  if (!tp) return { title: 'LEGER PER TP', head: [], rows: [] };
  const head = ['No', 'Nama Peserta Didik', ...tp.models.map(m => short(m.nama)), 'Nilai TP', 'Predikat TP'];
  const rows = siswaMapel(UI.mid).map((s, i) => { const c = tpCalc(md, tp, s.id);
    return [i + 1, s.nama, ...[0, 1, 2, 3].map(k => c && c.vals[k] != null ? c.vals[k] + ' (' + kodeOf(c.vals[k]) + ')' : ''), c ? c.avg : '', c ? predNama(c.avg) : '']; });
  return { title: `LEGER ${mapelOf(UI.mid).nama.toUpperCase()} — ${tp.kode}: ${stripDot(tp.deskripsi)}`, head, rows };
}
VIEW.leger = () => {
  if (!S.siswa.length) return noSiswa('Leger');
  const d = legerData(UI.leger); const p = S.profil;
  return `<div class="panel noprint"><div class="panel-b row">
      <b>Jenis leger:</b><select data-c="pickLeger" style="max-width:260px">${opts([['akhir', 'Leger Nilai Akhir (semua mapel)'], ['tp', 'Leger per TP (KUDIC)']], UI.leger)}</select>
      ${UI.leger === 'tp' ? `<select data-c="pickMapel" style="max-width:280px">${opts(S.mapel.map(m => [m.id, m.nama]), UI.mid)}</select>
        <select data-c="pickTp" style="max-width:220px">${opts(cur().tp.map(t => [t.id, t.kode]), UI.tp)}</select>` : ''}
      ${UI.leger === 'akhir' ? `<select data-c="pickLegerPem" style="max-width:190px">${opts([['semua', 'Semua kelompok pilihan'], ['IPA', 'Hanya IPA'], ['IPS', 'Hanya IPS']], UI.legerPem || 'semua')}</select>` : ''}
      <span class="pill">${SEM[p.sem]}</span><div class="spacer"></div>
      <button class="btn ghost" data-act="exportLeger"><i class="fa-solid fa-file-excel"></i>Ekspor Excel</button>
      <button class="btn" data-act="printLeger"><i class="fa-solid fa-print"></i>Cetak</button></div></div>
    <div class="paper"><div class="legerhead"><h2>${esc(d.title)}</h2>
      <div>${esc(p.sekolah)} — Kelas ${esc(p.kelas)} — Tahun Pelajaran ${esc(p.tahun)}</div></div>
      <div class="tw"><table class="grid"><thead><tr>${d.head.map(h => `<th class="wrapth">${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${d.rows.map(r => `<tr>${r.map((v, i) => `<td class="${i === 1 ? 'nm' : 'c'}">${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;
};
C.pickLeger = el => { UI.leger = el.value; render(); };
C.pickLegerPem = el => { UI.legerPem = el.value; render(); };
A.printLeger = () => { setPage('@page{size:330mm 216mm;margin:0} @media print{.content{padding:10mm!important}}'); window.print(); };
function downloadXlsx(aoa, name, sheet) {
  if (!window.XLSX) return toast('Pustaka Excel belum termuat (cek koneksi internet).');
  const ws = XLSX.utils.aoa_to_sheet(aoa); const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheet.slice(0, 30)); XLSX.writeFile(wb, name);
}
A.exportLeger = () => { const d = legerData(UI.leger); downloadXlsx([[d.title], [], d.head, ...d.rows], `leger-${UI.leger}-kelas-${S.profil.kelas}.xlsx`.replace(/[\s()]+/g, '_'), 'Leger'); };
A.exportRekap = () => {
  const md = cur(); const SM = siswaMapel(UI.mid); const cs = SM.map(s => calcSiswa(UI.mid, s.id)); const tps = md.tp.filter(t => cs.some(c => c.form[t.id]));
  const head = ['No', 'Nama', ...tps.map(t => t.kode), 'Rata Formatif', ...md.lm.map((l, i) => 'LM ' + (i + 1)), 'Rata SLM', 'STS', 'SAS', 'Nilai Raport', 'Predikat', 'TP Tertinggi', 'TP Terendah', 'Deskripsi Tertinggi', 'Deskripsi Terendah'];
  const rows = SM.map((s, i) => { const c = cs[i];
    return [i + 1, s.nama, ...tps.map(t => c.form[t.id]?.avg ?? ''), r0(c.avgF), ...md.lm.map(l => c.lmVals[l.id] ?? ''), r0(c.avgLM), c.sts ?? '', c.sas ?? '', r0(c.akhir), c.predikat, c.hi?.kode || '', c.lo?.kode || '', c.descHi, c.descLo]; });
  downloadXlsx([head, ...rows], `rekap-${mapelOf(UI.mid).nama}.xlsx`.replace(/\s+/g, '_'), 'Rekap');
};

/* =====================================================================
   4. IMPOR EXCEL — membaca file leger (satu file = satu mata pelajaran)
      Sheet yang dibaca: DATA SISWA, ISI DISINI, Leger TP x.y, SUMATIF
   ===================================================================== */
const colName = n => { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
function importLeger(wb) {
  const get = (sh, addr) => { const c = wb.Sheets[sh] && wb.Sheets[sh][addr]; return c == null ? '' : c.v; };
  const txt = (sh, addr) => String(get(sh, addr) ?? '').trim();
  const md = newMapelData(); const rep = []; const newSiswa = [];
  // --- siswa: cocokkan nama; bila daftar siswa aplikasi masih kosong, buat dari file
  const rowSid = {}; let skipped = 0;
  if (wb.Sheets['DATA SISWA']) for (let r = 10; r <= 39; r++) {
    const nm = txt('DATA SISWA', 'B' + r); if (!nm) continue;
    let s = S.siswa.find(x => x.nama.trim().toLowerCase() === nm.toLowerCase());
    if (!s && !S.siswa.length) { s = { id: uid(), nama: nm, nisn: '', nis: '', jk: '' }; newSiswa.push(s); }
    if (s) rowSid[r - 10] = s.id; else skipped++;
  }
  // --- ISI DISINI: LM & TP
  if (wb.Sheets['ISI DISINI']) {
    md.lm = []; let curLm = null;
    for (let r = 5; r <= 60; r++) {
      const b = txt('ISI DISINI', 'B' + r), c = txt('ISI DISINI', 'C' + r), dd = txt('ISI DISINI', 'D' + r);
      if (/^petunjuk/i.test(b)) break;
      if (b && /^lingkup\s+materi/i.test(b)) { curLm = { id: 'LM' + (md.lm.length + 1), nama: b.replace(/\(isi[^)]*\)/i, '').replace(/[:\s]+$/, '').trim() }; md.lm.push(curLm); }
      if (c && /^TP/i.test(c)) {
        if (!curLm) { curLm = { id: 'LM1', nama: 'Lingkup Materi 1' }; md.lm.push(curLm); }
        md.tp.push({ id: uid(), kode: c, lm: curLm.id, deskripsi: /^\(isi/i.test(dd) ? '' : dd, models: newModels() });
      }
    }
    rep.push(`${md.tp.length} TP`);
  }
  // --- Leger TP x.y: model KUDIC, deskripsi & nilai
  let nNilai = 0;
  wb.SheetNames.forEach(sn => {
    const m = /^Leger\s+(TP\s*[\d.]+)/i.exec(sn); if (!m) return;
    const tp = md.tp.find(t => t.kode.replace(/\s+/g, '').toLowerCase() === m[1].replace(/\s+/g, '').toLowerCase()); if (!tp) return;
    [8, 12, 16, 20].forEach((row0, mi) => {
      const nm = txt(sn, 'B' + row0), ind = txt(sn, 'C' + row0);
      if (nm) tp.models[mi].nama = nm;
      if (ind) tp.models[mi].ind = ind.split('\n').slice(1).join(' ').trim() || ind;
      ['SB', 'B', 'C', 'K'].forEach((k, ki) => { tp.models[mi].desk[k] = txt(sn, 'F' + (row0 + ki)); });
    });
    Object.entries(rowSid).forEach(([i, sid]) => {
      const base = 25 + 4 * (+i);
      const n = [0, 1, 2, 3].map(k => { const v = get(sn, 'D' + (base + k)); return v === '' ? '' : String(v); });
      const cat = [0, 1, 2, 3].map(k => txt(sn, 'G' + (base + k))).filter(Boolean).join(' ');
      if (n.some(x => x !== '') || cat) { (md.nilai[tp.id] ||= {})[sid] = { n, cat }; nNilai++; }
    });
  });
  if (nNilai) rep.push(`${nNilai} baris nilai TP`);
  // --- SUMATIF: LM, STS, SAS
  if (wb.Sheets['SUMATIF']) Object.entries(rowSid).forEach(([i, sid]) => {
    const r = 12 + (+i); const lm = {};
    md.lm.slice(0, 4).forEach((l, k) => { const v = get('SUMATIF', colName(28 + k) + r); if (v !== '') lm[l.id] = String(v); });   // kolom AC..AF
    const sts = get('SUMATIF', 'AH' + r), sas = get('SUMATIF', 'AI' + r);
    md.sumatif[sid] = { lm, sts: sts === '' ? '' : String(sts), sas: sas === '' ? '' : String(sas) };
  });
  if (newSiswa.length) rep.push(`${newSiswa.length} siswa baru`);
  if (skipped) rep.push(`${skipped} nama tidak cocok (dilewati)`);
  return { md, newSiswa, rep };
}
A.importLeger = b => { UI.importMid = S.mapel[+b.dataset.i].id; const f = $('#filePick'); f.accept = '.xlsx,.xls'; f.multiple = false; f.onchange = () => readFile(f, 'leger'); f.value = ''; f.click(); };
A.importJson = () => { const f = $('#filePick'); f.accept = '.json'; f.multiple = false; f.onchange = () => readFile(f, 'json'); f.value = ''; f.click(); };
function readFile(f, kind) {
  const file = f.files[0]; if (!file) return;
  const rd = new FileReader();
  rd.onload = () => {
    try {
      if (kind === 'json') {
        const o = JSON.parse(rd.result);
        if (o.multi && o.kelas) {
          if (!confirm('Ganti data SEMUA kelas dengan isi backup?')) return;
          KELAS_LIST.forEach(([k]) => { if (o.kelas[k]) localStorage.setItem(KEYK(k), JSON.stringify(o.kelas[k])); });
          KA = o.kelas[o.aktif] ? o.aktif : (KELAS_LIST.find(([k]) => o.kelas[k]) || ['X'])[0];
          localStorage.setItem(KEY_AKTIF, KA); S = loadClass(KA) || newKelas(KA); UI.mid = UI.tp = UI.sid = UI.kid = UI.tpMid = ''; render(); toast('Backup semua kelas dipulihkan'); return;
        }
        if (!o.profil || !o.siswa || !o.sems) throw new Error('Format backup tidak dikenali');
        const kk = kelasKey(o.profil.kelas);
        if (!confirm(`Ganti data kelas ${kk} dengan isi backup?`)) return;
        S = normalizeState(o); KA = kk; UI.mid = UI.tp = UI.sid = UI.kid = UI.tpMid = ''; save(); render(); toast('Backup dipulihkan (kelas ' + kk + ')');
      } else {
        if (!window.XLSX) return toast('Pustaka Excel belum termuat (cek koneksi internet).');
        const wb = XLSX.read(rd.result, { type: 'array' });
        const { md, newSiswa, rep } = importLeger(wb);
        const nm = mapelOf(UI.importMid).nama;
        if (!confirm(`Ditemukan: ${rep.join(', ') || 'tidak ada data'}.\nGanti data "${nm}" semester ${SEM[S.profil.sem]} dengan isi file ini?`)) return;
        newSiswa.forEach(s => S.siswa.push(s)); D().m[UI.importMid] = md; UI.tp = ''; save(); render(); toast(`Impor ke ${nm} berhasil: ${rep.join(', ')}`);
      }
    } catch (err) { console.error(err); toast('Gagal membaca file: ' + err.message); }
  };
  kind === 'json' ? rd.readAsText(file) : rd.readAsArrayBuffer(file);
}

/* =====================================================================
   8. PENGATURAN, BACKUP, DATA CONTOH
   ===================================================================== */
VIEW.pengaturan = () => `
  ${panel('Profil Sekolah', 'fa-school', `<div class="form-grid">
    ${field('Nama Yayasan (baris atas kop)', 'profil.yayasan')}${field('Nama Sekolah', 'profil.sekolah')}${field('NPSN', 'profil.npsn')}${field('NIPD/NISS', 'profil.nss')}${field('Alamat Sekolah', 'profil.alamat')}
    ${field('Kode Pos', 'profil.kodePos')}${field('Telepon Sekolah', 'profil.telp')}${field('Desa / Kelurahan', 'profil.desa')}${field('Kecamatan', 'profil.kecamatan')}
    ${field('Kabupaten / Kota', 'profil.kabupaten')}${field('Provinsi', 'profil.provinsi')}${field('Website', 'profil.website')}${field('E-mail', 'profil.email')}
    ${field('Nama Kepala Sekolah', 'profil.kepsek')}${field('NIP Kepala Sekolah', 'profil.nipKepsek')}</div>`)}
  ${panel('Logo Sekolah (kop & sampul rapor)', 'fa-image', `<div class="row"><img src="${logoSrc()}" alt="Logo" style="height:90px">
    <label class="btn ghost"><i class="fa-solid fa-upload"></i>Ganti logo<input type="file" accept="image/*" data-c="logoUp" hidden></label>
    <button class="btn ghost" data-act="logoReset"><i class="fa-solid fa-rotate-left"></i>Pakai logo bawaan</button></div>`)}
  ${panel('Lambang Daerah (bagian atas sampul rapor)', 'fa-shield-halved', `<div class="row">${S.profil.lambang ? `<img src="${S.profil.lambang}" alt="Lambang" style="height:90px">` : '<span class="muted small">Belum ada lambang. Unggah lambang provinsi/kabupaten (gambar PNG/JPG).</span>'}
    <label class="btn ghost"><i class="fa-solid fa-upload"></i>Unggah lambang<input type="file" accept="image/*" data-c="lambangUp" hidden></label>
    <button class="btn ghost" data-act="lambangDel"><i class="fa-solid fa-xmark"></i>Hapus lambang</button></div>`)}
  ${panel('Profil Guru, Kelas & Rapor', 'fa-user-tie', `<div class="form-grid">
    ${field('Nama Guru kelas', 'profil.guru')}${field('NIP/NIPPPK', 'profil.nipGuru')}
    <div><label class="f">Jenjang</label><select data-bind="profil.jenjang">${opts(['SD', 'SMP', 'SMA', 'SMK', 'MI', 'MTs', 'MA'], S.profil.jenjang)}</select></div>
    ${field('Kelas', 'profil.kelas')}${field('Fase', 'profil.fase')}${field('Tahun Pelajaran', 'profil.tahun')}
    <div><label class="f">Semester aktif</label><select data-bind="profil.sem" data-c="semSet">${opts([['1', SEM['1']], ['2', SEM['2']]], S.profil.sem)}</select></div>
    ${field('Tempat Rapor', 'profil.tempatRapor')}${field('Tanggal Rapor', 'profil.tglRapor', 'date')}${field('Tanggal Tahun Ajaran', 'profil.tglAjaran')}</div>`)}
  ${panel('Impor & Backup Data', 'fa-database', `
    <div class="hint info">Data tersimpan otomatis di browser ini. Buat backup berkala, terutama sebelum membersihkan data browser atau berpindah komputer.
      Impor file leger Excel dilakukan <b>per mata pelajaran</b> di menu <a href="#/mapel">Mata Pelajaran</a> (tombol “Impor leger”).</div>
    <div class="row">
      <button class="btn" data-act="exportJson"><i class="fa-solid fa-download"></i>Backup (.json)</button>
      <button class="btn ghost" data-act="importJson"><i class="fa-solid fa-upload"></i>Pulihkan backup</button>
      <button class="btn ghost" data-act="demo"><i class="fa-solid fa-flask"></i>Muat data contoh</button>
      <button class="btn danger" data-act="reset"><i class="fa-solid fa-trash"></i>Hapus semua data</button>
    </div>`)}`;
C.semSet = () => render();
C.logoUp = el => {
  const f = el.files[0]; if (!f) return; const rd = new FileReader();
  rd.onload = () => { const img = new Image(); img.onload = () => {
    const k = Math.min(1, 400 / Math.max(img.width, img.height)); const cv = document.createElement('canvas'); cv.width = img.width * k; cv.height = img.height * k;
    cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height); S.profil.logo = cv.toDataURL('image/png'); save(); render(); toast('Logo diganti'); }; img.src = rd.result; };
  rd.readAsDataURL(f);
};
A.logoReset = () => { S.profil.logo = ''; save(); render(); };
C.lambangUp = el => {
  const f = el.files[0]; if (!f) return; const rd = new FileReader();
  rd.onload = () => { const img = new Image(); img.onload = () => {
    const k = Math.min(1, 500 / Math.max(img.width, img.height)); const cv = document.createElement('canvas'); cv.width = img.width * k; cv.height = img.height * k;
    cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height); S.profil.lambang = cv.toDataURL('image/png'); save(); render(); toast('Lambang disimpan'); }; img.src = rd.result; };
  rd.readAsDataURL(f);
};
A.lambangDel = () => { S.profil.lambang = ''; save(); render(); };
A.exportJson = () => {
  try { localStorage.setItem(KEYK(KA), JSON.stringify(S)); } catch (e) { /* abaikan */ }
  const all = { multi: true, aktif: KA, kelas: {} };
  KELAS_LIST.forEach(([k]) => { try { const r = localStorage.getItem(KEYK(k)); if (r) all.kelas[k] = JSON.parse(r); } catch (e) { /* abaikan */ } });
  all.kelas[KA] = S;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([JSON.stringify(all)], { type: 'application/json' }));
  a.download = `backup-erapor-semua-kelas-${new Date().toISOString().slice(0, 10)}.json`; a.click(); URL.revokeObjectURL(a.href);
};
A.reset = () => { if (confirm(`Hapus SEMUA data kelas ${KA} (siswa, mata pelajaran, nilai)? Data kelas lain tidak terpengaruh. Tindakan ini tidak bisa dibatalkan.`)) { S = newKelas(KA, S); UI.mid = ''; UI.tp = ''; UI.sid = ''; save(); render(); toast('Data dihapus'); } };
A.demo = () => {
  if (S.siswa.length && !confirm('Data saat ini akan diganti dengan data contoh. Lanjutkan?')) return;
  const N = defaults(); N.profil = { ...S.profil, sekolah: S.profil.sekolah || 'SD Contoh Negeri 1', guru: S.profil.guru || 'Nama Wali Kelas, S.Pd.', kepsek: S.profil.kepsek || 'Nama Kepala Sekolah, S.Pd.' };
  ['Amanda Ekalistra Selan', 'Galti Putriyani Mtbani', 'Jecicco G. F Nabut', 'Ranti Nenoliu', 'Riski Febrianto Nesimnasi', 'Shezya Migata Boys']
    .forEach((n, i) => N.siswa.push({ id: uid(), nama: n, nisn: '', nis: '', jk: '', pem: i % 2 ? 'IPS' : 'IPA' }));
  const rnd = (a, b) => String(a + Math.floor(Math.random() * (b - a + 1)));
  const d = N.sems['1'];
  N.mapel.forEach(m => {
    const md = (d.m[m.id] = newMapelData());
    [`Memahami konsep dasar ${m.nama}.`, `Menerapkan konsep ${m.nama} dalam kehidupan sehari-hari.`].forEach((ds, i) => {
      const tp = { id: uid(), kode: `TP 1.${i + 1}`, lm: 'LM1', deskripsi: ds, models: newModels() }; md.tp.push(tp);
      tp.models.forEach(mm => ['SB', 'B', 'C', 'K'].forEach(k => mm.desk[k] = `Siswa ${{ SB: 'sangat baik', B: 'baik', C: 'cukup', K: 'belum' }[k]} pada tahap ${short(mm.nama)}`));
      N.siswa.forEach(s => { (md.nilai[tp.id] ||= {})[s.id] = { n: [rnd(80, 95), rnd(72, 90), rnd(65, 85), rnd(60, 80)], cat: '' }; });
    });
    N.siswa.forEach(s => { md.sumatif[s.id] = { lm: { LM1: rnd(70, 92) }, sts: rnd(70, 90), sas: rnd(70, 90) }; });
  });
  d.hari = 90;
  N.siswa.forEach(s => {
    d.absen[s.id] = { S: rnd(0, 3), I: rnd(0, 2), A: '0' }; d.eks[s.id] = { [N.ekskul[0].id]: PRED_EKS[1 + Math.floor(Math.random() * 3)] };
    d.catatan[s.id] = 'Terus tingkatkan semangat belajar.';
    d.koku[0].skor[s.id] = Array(24).fill('').map((_, i) => (i < 3 ? rnd(3, 4) : ''));
  });
  d.koku[0].tema = 'Aku Cinta Indonesia'; d.koku[0].dims[0].aspek[0] = 'Berdoa sebelum belajar'; d.koku[0].dims[0].aspek[1] = 'Bersyukur atas ciptaan Tuhan';
  S = N; UI.mid = ''; UI.tp = ''; UI.sid = ''; save(); render(); toast('Data contoh dimuat');
};

/* ---- mulai ---- */
render();