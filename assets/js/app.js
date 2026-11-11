// ============================================================
//  ADMIN MBG — DATA LAYER v2.0
//  localStorage key: mbg_v2
// ============================================================
const STORE_KEY = 'mbg_v2';

function emptyDB() {
  return {
    meta: {
      sppgName: 'SPPG MBG',
      adminName: 'Administrator',
      adminRole: 'Admin MBG',
      createdAt: new Date().toISOString()
    },
    staf: [],
    absensi: [],
    kontrak_history: [],
    catatan: []
  };
}

function loadDB() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      // ensure all keys exist
      return { ...emptyDB(), ...parsed };
    }
  } catch(e) {}
  return emptyDB();
}

function saveDB() {
  localStorage.setItem(STORE_KEY, JSON.stringify(DB));
}

let DB = loadDB();

// ============================================================
//  UTILS
// ============================================================
const JABATANS = [
  'Kepala SPPG','Admin MBG','Ahli Gizi','Juru Masak',
  'Asisten Juru Masak','Petugas Distribusi','Kebersihan','Keamanan','Driver','Lainnya'
];
const STATUS_ABS = ['Hadir','Terlambat','Izin','Sakit','Alpha'];
const STATUS_ABS_COLOR = {
  Hadir: 'bg', Terlambat: 'by', Izin: 'bp', Sakit: 'br', Alpha: 'bgr', '': 'bgr'
};
const AVATAR_COLORS = [
  'linear-gradient(135deg,#2563eb,#7c3aed)',
  'linear-gradient(135deg,#059669,#0891b2)',
  'linear-gradient(135deg,#d97706,#ef4444)',
  'linear-gradient(135deg,#7c3aed,#ec4899)',
  'linear-gradient(135deg,#0891b2,#2563eb)',
];

function avatarColor(name) {
  let n = 0;
  for (let c of (name||'')) n += c.charCodeAt(0);
  return AVATAR_COLORS[n % AVATAR_COLORS.length];
}

function fmt(str) {
  if (!str) return '-';
  const d = new Date(str + 'T00:00:00');
  return d.toLocaleDateString('id-ID', { day:'2-digit', month:'long', year:'numeric' });
}
function fmtShort(str) {
  if (!str) return '-';
  const d = new Date(str + 'T00:00:00');
  return d.toLocaleDateString('id-ID', { day:'2-digit', month:'short', year:'numeric' });
}
function fmtCur(n) {
  return 'Rp ' + (n||0).toLocaleString('id-ID');
}
function daysUntil(str) {
  if (!str) return 9999;
  const t = new Date(str + 'T00:00:00');
  const now = new Date(); now.setHours(0,0,0,0);
  return Math.round((t - now) / 86400000);
}
function todayStr() { return new Date().toISOString().slice(0,10); }
function curYM() {
  const n = new Date();
  return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}`;
}
function ymLabel(ym) {
  const [y,m] = ym.split('-');
  return new Date(y, m-1, 1).toLocaleDateString('id-ID', { month:'long', year:'numeric' });
}
function workDaysInMonth(ym) {
  const [y,m] = ym.split('-').map(Number);
  const days = new Date(y,m,0).getDate();
  let w = 0;
  for (let d=1;d<=days;d++) {
    const day = new Date(y,m-1,d).getDay();
    if (day!==0 && day!==6) w++;
  }
  return w;
}
function esc(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function nextId(prefix) {
  const nums = DB.staf.filter(s=>s.id.startsWith(prefix)).map(s=>parseInt(s.id.replace(prefix,''))||0);
  const max = nums.length ? Math.max(...nums) : 0;
  return prefix + String(max+1).padStart(3,'0');
}
function csvDownload(rows, filename) {
  // Gunakan 'sep=;' dan pemisah titik koma agar Microsoft Excel (khususnya regional Indonesia/Windows) 100% membagi ke kolom A, B, C, D
  const csv = 'sep=;\r\n' + rows.map(r => r.map(v => `"${String(v ?? '').replace(/"/g, '""')}"`).join(';')).join('\r\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' }));
  a.download = filename;
  a.click();
}

function exportToExcelTable(sheetTitle, headers, rows, filename) {
  if (typeof XLSX !== 'undefined') {
    const data = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Auto-fit column widths so text is never truncated
    const colWidths = headers.map((h, colIdx) => {
      let maxLen = String(h).length;
      rows.forEach(r => {
        const val = r[colIdx];
        const cellLen = val !== null && val !== undefined ? String(val).length : 0;
        if (cellLen > maxLen) maxLen = cellLen;
      });
      return { wch: Math.min(Math.max(maxLen + 4, 12), 40) };
    });
    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetTitle.slice(0, 31));

    const fname = filename.endsWith('.xlsx') ? filename : filename.replace(/\.[^.]+$/, '') + '.xlsx';
    XLSX.writeFile(wb, fname);
    return;
  }

  // Fallback if XLSX library is not yet ready
  const csv = 'sep=;\r\n' + [headers, ...rows].map(r => r.map(v => `"${String(v ?? '').replace(/"/g, '""')}"`).join(';')).join('\r\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' }));
  a.download = filename.endsWith('.csv') ? filename : filename.replace(/\.[^.]+$/, '') + '.csv';
  a.click();
}

// ============================================================
//  TOAST
// ============================================================
function toast(msg, type='ok') {
  const icons = {ok:'✅',err:'❌',warn:'⚠️',info:'ℹ️'};
  const c = document.getElementById('toasts');
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${icons[type]||'ℹ️'}</span><span style="flex:1">${msg}</span>`;
  c.appendChild(el);
  setTimeout(()=>{ el.style.transition='all .3s'; el.style.opacity='0'; el.style.transform='translateY(8px)'; setTimeout(()=>el.remove(),300); }, 3000);
}

// ============================================================
//  MODAL HELPERS
// ============================================================
function openModal(id) { document.getElementById(id)?.classList.add('open'); }
function closeModal(id) { document.getElementById(id)?.classList.remove('open'); }

document.addEventListener('keydown', e => {
  if (e.key === 'Escape') document.querySelectorAll('.overlay.open').forEach(m=>m.classList.remove('open'));
});

// ============================================================
//  NAVIGATION
// ============================================================
let _page = 'dashboard';

const PAGE_META = {
  dashboard: ['🏠 Dashboard',       'Ringkasan administrasi kepegawaian MBG'],
  staf:      ['👥 Data Staf',        'Kelola database seluruh staf MBG'],
  absensi:   ['📋 Absensi Harian',   'Pencatatan kehadiran harian staf'],
  rekap:     ['📊 Rekap Bulanan',    'Rekapitulasi kehadiran & kinerja per bulan'],
  kontrak:   ['📄 Manajemen Kontrak','Pantau & kelola masa berlaku kontrak'],
  excel:     ['📗 Pengelolaan Excel', 'Pusat administrasi data Microsoft Excel MBG'],
  laporan:   ['📈 Laporan',          'Laporan administrasi & statistik'],
  pengaturan:['⚙️ Pengaturan',       'Konfigurasi aplikasi & profil'],
};

function navigate(page) {
  _page = page;
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  const targetPage = document.getElementById('page-'+page);
  if (targetPage) targetPage.classList.add('active');
  document.querySelectorAll(`.nav-item[data-page="${page}"]`).forEach(n=>n.classList.add('active'));
  const [title,sub] = PAGE_META[page]||['Admin MBG',''];
  const hTitle = document.getElementById('h-title');
  if (hTitle) hTitle.textContent = title;
  const hSub = document.getElementById('h-sub');
  if (hSub) hSub.textContent = sub;
  
  // close mobile off-canvas if open
  if (window.innerWidth <= 640) {
    document.getElementById('sidebar')?.classList.remove('open');
    document.getElementById('mob-overlay')?.classList.remove('open');
  }

  // scroll to top smoothly
  window.scrollTo({ top: 0, behavior: 'instant' });

  // render
  try {
    const renderers = {
      dashboard: renderDash,
      staf: renderStaf,
      absensi: renderAbsensi,
      rekap: renderRekap,
      kontrak: renderKontrak,
      excel: renderExcelPage,
      laporan: renderLaporan,
      pengaturan: renderPengaturan
    };
    renderers[page]?.();
  } catch (err) {
    console.error('Error rendering page:', page, err);
  }

  updateHeaderMeta();
}

function updateHeaderMeta() {
  const expiring = DB.staf.filter(s=>s.status==='Aktif'&&daysUntil(s.tglKontrakBerakhir)<=30&&daysUntil(s.tglKontrakBerakhir)>=0).length;
  const badge = document.getElementById('nb-kontrak');
  if (badge) { badge.textContent = expiring; badge.style.display = expiring ? '' : 'none'; }
  
  let uName = DB.meta.adminName || 'Admin MBG';
  let uRole = DB.meta.adminRole || 'Administrator';
  try {
    const rawSess = sessionStorage.getItem('mbg_session') || localStorage.getItem('mbg_session');
    if (rawSess) {
      const sess = JSON.parse(rawSess);
      if (sess.name) uName = sess.name;
      if (sess.role) uRole = sess.role;
    }
  } catch(e) {}

  const elName = document.getElementById('h-user-name');
  const elRole = document.getElementById('h-user-role');
  const elInit = document.getElementById('h-user-init');
  if (elName) elName.textContent = uName;
  if (elRole) elRole.textContent = uRole;
  if (elInit) elInit.textContent = (uName || 'A').charAt(0).toUpperCase();
}

function logoutMBG() {
  if (confirm('Keluar dari sistem Admin MBG?')) {
    sessionStorage.removeItem('mbg_session');
    localStorage.removeItem('mbg_session');
    window.location.href = 'index.html';
  }
}

// ============================================================
//  DASHBOARD
// ============================================================
function renderDash() {
  const aktif = DB.staf.filter(s=>s.status==='Aktif').length;
  const total = DB.staf.length;
  const nonaktif = DB.staf.filter(s=>s.status==='Nonaktif').length;
  const exp30 = DB.staf.filter(s=>s.status==='Aktif'&&daysUntil(s.tglKontrakBerakhir)<=30&&daysUntil(s.tglKontrakBerakhir)>=0).length;
  const exp90 = DB.staf.filter(s=>s.status==='Aktif'&&daysUntil(s.tglKontrakBerakhir)<=90&&daysUntil(s.tglKontrakBerakhir)>=0).length;

  const today = todayStr();
  const todayRecs = DB.absensi.filter(a=>a.tanggal===today);
  const todayHadir = todayRecs.filter(a=>a.status==='Hadir'||a.status==='Terlambat').length;

  set('d-total',total); set('d-aktif',aktif); set('d-nonaktif',nonaktif);
  set('d-hadir',todayHadir+'/'+aktif); set('d-exp30',exp30);

  if (!total) {
    setHtml('d-welcome', welcomeHtml());
    document.getElementById('d-charts')?.classList.add('hidden');
  } else {
    setHtml('d-welcome','');
    document.getElementById('d-charts')?.classList.remove('hidden');
    renderDashCharts();
    renderDashAlerts();
    renderDashJabatan();
  }
}

function welcomeHtml() {
  return `<div class="card">
    <div class="card-body" style="padding:40px;text-align:center">
      <div style="font-size:3rem;margin-bottom:12px">🍱</div>
      <div style="font-size:1.1rem;font-weight:700;color:var(--gray-800);margin-bottom:6px">Selamat Datang di Admin MBG!</div>
      <div class="text-sm text-muted mb-4">Belum ada data staf. Mulai dengan menambahkan data staf pertama.</div>
      <div class="flex gap-3" style="justify-content:center;flex-wrap:wrap">
        <button class="btn btn-primary" onclick="navigate('staf')">➕ Tambah Staf Pertama</button>
        <button class="btn btn-ghost" onclick="navigate('pengaturan')">⚙️ Atur Profil SPPG</button>
      </div>
    </div>
  </div>`;
}

function renderDashCharts() {
  renderJabatanChart();
  renderTrendChart();
}

function renderJabatanChart() {
  const ctx = document.getElementById('c-jabatan');
  if (!ctx || typeof Chart === 'undefined') return;
  try {
    if (ctx._ch) ctx._ch.destroy();
    const aktif = DB.staf.filter(s=>s.status==='Aktif');
    const jab = {};
    aktif.forEach(s=>{ jab[s.jabatan]=(jab[s.jabatan]||0)+1; });
    const labels = Object.keys(jab);
    const data = Object.values(jab);
    const colors = ['#1d4ed8','#16a34a','#d97706','#dc2626','#7c3aed','#0891b2','#ec4899','#f97316','#059669'];
    ctx._ch = new Chart(ctx, {
      type:'doughnut',
      data:{ labels, datasets:[{data, backgroundColor:colors.slice(0,labels.length), borderWidth:2, borderColor:'#fff'}] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{ position:'right', labels:{font:{size:11},boxWidth:12,padding:8} } } }
    });
  } catch (e) { console.warn(e); }
}

function renderTrendChart() {
  const ctx = document.getElementById('c-trend');
  if (!ctx || typeof Chart === 'undefined') return;
  try {
    if (ctx._ch) ctx._ch.destroy();
    const months=[]; const dH=[],dT=[],dI=[],dS=[];
    const now=new Date();
    for(let i=2;i>=0;i--) {
      const d=new Date(now.getFullYear(),now.getMonth()-i,1);
      const ym=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
      months.push(d.toLocaleDateString('id-ID',{month:'short'}));
      const r=DB.absensi.filter(a=>a.tanggal.startsWith(ym));
      dH.push(r.filter(a=>a.status==='Hadir').length);
      dT.push(r.filter(a=>a.status==='Terlambat').length);
      dI.push(r.filter(a=>a.status==='Izin').length);
      dS.push(r.filter(a=>a.status==='Sakit').length);
    }
    ctx._ch = new Chart(ctx, {
      type:'bar',
      data:{ labels:months, datasets:[
        {label:'Hadir',data:dH,backgroundColor:'#16a34a',borderRadius:4},
        {label:'Terlambat',data:dT,backgroundColor:'#d97706',borderRadius:4},
        {label:'Izin',data:dI,backgroundColor:'#7c3aed',borderRadius:4},
        {label:'Sakit',data:dS,backgroundColor:'#dc2626',borderRadius:4},
      ]},
      options:{ responsive:true, maintainAspectRatio:false, plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12}}},
        scales:{ x:{stacked:true,grid:{display:false}}, y:{stacked:true,beginAtZero:true} } }
    });
  } catch (e) { console.warn(e); }
}

function renderDashAlerts() {
  const el = document.getElementById('d-alerts');
  if (!el) return;
  const exp = DB.staf.filter(s=>s.status==='Aktif'&&daysUntil(s.tglKontrakBerakhir)<=90)
    .sort((a,b)=>daysUntil(a.tglKontrakBerakhir)-daysUntil(b.tglKontrakBerakhir));
  if (!exp.length) {
    el.innerHTML = emptyHtml('✅','Tidak ada kontrak yang akan berakhir dalam 90 hari');
    return;
  }
  el.innerHTML = exp.map(s=>{
    const d=daysUntil(s.tglKontrakBerakhir);
    const cls=d<=14?'err':'warn';
    return `<div class="flex items-center gap-3" style="padding:10px 0;border-bottom:1px solid var(--gray-100)">
      <div style="width:44px;height:44px;border-radius:9px;background:${d<=14?'var(--red-light)':'var(--yellow-light)'};display:flex;flex-direction:column;align-items:center;justify-content:center;flex-shrink:0">
        <span style="font-size:1rem;font-weight:800;color:${d<=14?'var(--red)':'var(--yellow)'};line-height:1">${d}</span>
        <span style="font-size:.58rem;color:var(--gray-400)">hari</span>
      </div>
      <div class="flex-1">
        <div class="font-semi text-sm">${esc(s.nama)}</div>
        <div class="text-xs text-muted">${s.jabatan} · ${fmtShort(s.tglKontrakBerakhir)}</div>
      </div>
      <button class="btn btn-xs btn-primary" onclick="openPerpanjang('${s.id}')">Perpanjang</button>
    </div>`;
  }).join('');
}

function renderDashJabatan() {
  const el = document.getElementById('d-jabatan');
  if (!el) return;
  const aktif = DB.staf.filter(s=>s.status==='Aktif');
  const jab = {};
  aktif.forEach(s=>{ jab[s.jabatan]=(jab[s.jabatan]||0)+1; });
  if (!Object.keys(jab).length) { el.innerHTML = emptyHtml('📋','Belum ada data staf'); return; }
  el.innerHTML = Object.entries(jab).map(([j,c])=>{
    const pct=aktif.length?Math.round(c/aktif.length*100):0;
    return `<div style="margin-bottom:10px">
      <div class="flex justify-between items-center mb-1">
        <span class="text-sm font-semi">${esc(j)}</span>
        <span class="text-xs text-muted">${c} orang · ${pct}%</span>
      </div>
      <div class="prog"><div class="prog-fill pf-blue" style="width:${pct}%"></div></div>
    </div>`;
  }).join('');
}

// ============================================================
//  STAF MODULE
// ============================================================
let _stafFilters = { q:'', jabatan:'', status:'' };
let _editingStafId = null;

function renderStaf() {
  renderStafStats();
  renderStafTable();
}

function renderStafStats() {
  const aktif = DB.staf.filter(s=>s.status==='Aktif').length;
  set('ss-total',DB.staf.length);
  set('ss-aktif',aktif);
  set('ss-nonaktif',DB.staf.filter(s=>s.status==='Nonaktif').length);
  set('ss-jab',[...new Set(DB.staf.map(s=>s.jabatan))].length);
}

function renderStafTable() {
  const tbody = document.getElementById('staf-tbody');
  let data = [...DB.staf];
  if (_stafFilters.q) {
    const q=_stafFilters.q.toLowerCase();
    data=data.filter(s=>s.nama.toLowerCase().includes(q)||s.id.toLowerCase().includes(q)||(s.nik||'').includes(q));
  }
  if (_stafFilters.jabatan) data=data.filter(s=>s.jabatan===_stafFilters.jabatan);
  if (_stafFilters.status) data=data.filter(s=>s.status===_stafFilters.status);

  if (!data.length) {
    tbody.innerHTML = `<tr><td colspan="9">${emptyHtml(!DB.staf.length?'👥':'🔍',!DB.staf.length?'Belum ada staf. Klik tombol "+ Tambah Staf" untuk memulai.':'Tidak ada hasil yang cocok.')}</td></tr>`;
    return;
  }

  tbody.innerHTML = data.map(s=>{
    const days = daysUntil(s.tglKontrakBerakhir);
    let kb;
    if (s.status==='Nonaktif') kb=`<span class="badge bgr">Nonaktif</span>`;
    else if (days<0) kb=`<span class="badge br">Berakhir</span>`;
    else if (days<=30) kb=`<span class="badge br">≤30 hari</span>`;
    else if (days<=90) kb=`<span class="badge by">≤90 hari</span>`;
    else kb=`<span class="badge bg">Aktif</span>`;
    return `<tr>
      <td><span class="td-id">${s.id}</span></td>
      <td>
        <div class="flex items-center gap-2">
          <div class="avatar" style="background:${avatarColor(s.nama)}">${s.nama.charAt(0)}</div>
          <div>
            <div class="td-name">${esc(s.nama)}</div>
            <div class="text-xs text-muted">${s.nik||'-'}</div>
          </div>
        </div>
      </td>
      <td>${esc(s.jabatan)}</td>
      <td>${s.pendidikan||'-'}</td>
      <td>${fmtShort(s.tglMasuk)}</td>
      <td>${fmtShort(s.tglKontrakBerakhir)}</td>
      <td>${kb}</td>
      <td><span class="badge ${s.status==='Aktif'?'bg':'br'}">${s.status}</span></td>
      <td>
        <div class="flex gap-1">
          <button class="btn btn-icon btn-ghost btn-xs" onclick="viewStaf('${s.id}')" title="Detail">👁️</button>
          <button class="btn btn-icon btn-ghost btn-xs" onclick="openStafModal('${s.id}')" title="Edit">✏️</button>
          <button class="btn btn-icon btn-xs" style="background:var(--red-light);color:var(--red)" onclick="deleteStaf('${s.id}')" title="Hapus">🗑️</button>
        </div>
      </td>
    </tr>`;
  }).join('');
}

function openStafModal(id=null) {
  _editingStafId = id;
  const f = document.getElementById('f-staf');
  f.reset();
  document.getElementById('staf-modal-title').textContent = id ? '✏️ Edit Data Staf' : '➕ Tambah Staf Baru';
  if (id) {
    const s = DB.staf.find(x=>x.id===id);
    if (!s) return;
    Object.keys(s).forEach(k => { const el=f.querySelector(`[name="${k}"]`); if(el) el.value=s[k]; });
  } else {
    f.querySelector('[name=status]').value = 'Aktif';
  }
  openModal('modal-staf');
}

function saveStaf() {
  const f = document.getElementById('f-staf');
  const v = k => f.querySelector(`[name="${k}"]`)?.value?.trim() || '';
  const data = {
    nama: v('nama'), nik: v('nik'), jabatan: v('jabatan'),
    tempatLahir: v('tempatLahir'), tglLahir: v('tglLahir'),
    jenisKelamin: v('jenisKelamin'), agama: v('agama'),
    pendidikan: v('pendidikan'), noKK: v('noKK'),
    tglMasuk: v('tglMasuk'), tglKontrakBerakhir: v('tglKontrakBerakhir'),
    gaji: parseInt(v('gaji'))||0, telp: v('telp'),
    alamat: v('alamat'), status: v('status'), catatan: v('catatan'),
  };
  if (!data.nama||!data.jabatan||!data.tglMasuk||!data.tglKontrakBerakhir) {
    toast('Lengkapi field yang wajib diisi (*)','err'); return;
  }
  if (_editingStafId) {
    const idx = DB.staf.findIndex(s=>s.id===_editingStafId);
    if (idx!==-1) DB.staf[idx] = { ...DB.staf[idx], ...data };
    toast('Data staf berhasil diperbarui ✅');
  } else {
    data.id = nextId('STF');
    DB.staf.push(data);
    toast('Staf baru berhasil ditambahkan ✅');
  }
  saveDB(); closeModal('modal-staf'); renderStaf();
  if (_page==='dashboard') renderDash();
}

function deleteStaf(id) {
  const s = DB.staf.find(x=>x.id===id);
  if (!s) return;
  if (!confirm(`Hapus data staf "${s.nama}"?\n\nData absensi terkait juga akan ikut terhapus.`)) return;
  DB.staf = DB.staf.filter(x=>x.id!==id);
  DB.absensi = DB.absensi.filter(a=>a.stafId!==id);
  saveDB(); toast('Data staf dihapus','warn'); renderStaf();
}

function viewStaf(id) {
  const s = DB.staf.find(x=>x.id===id);
  if (!s) return;
  const days = daysUntil(s.tglKontrakBerakhir);
  const ym = curYM();
  const recs = DB.absensi.filter(a=>a.stafId===id&&a.tanggal.startsWith(ym));
  const hadir = recs.filter(a=>a.status==='Hadir').length;
  const terlambat = recs.filter(a=>a.status==='Terlambat').length;
  const izin = recs.filter(a=>a.status==='Izin').length;
  const sakit = recs.filter(a=>a.status==='Sakit').length;
  document.getElementById('staf-detail-body').innerHTML = `
    <div style="text-align:center;padding:16px 0 20px">
      <div class="avatar avatar-lg" style="background:${avatarColor(s.nama)};margin:0 auto 10px">${s.nama.charAt(0)}</div>
      <div style="font-size:1rem;font-weight:700">${esc(s.nama)}</div>
      <div class="text-xs text-muted">${esc(s.jabatan)} · ${s.id}</div>
      <span class="badge ${s.status==='Aktif'?'bg':'br'}" style="margin-top:6px">${s.status}</span>
    </div>
    <div class="tabs mb-3" id="detail-tabs">
      <div class="tab active" onclick="switchDetailTab('info')">📋 Info</div>
      <div class="tab" onclick="switchDetailTab('kontrak')">📄 Kontrak</div>
      <div class="tab" onclick="switchDetailTab('absensi')">📊 Absensi Bulan Ini</div>
    </div>
    <div id="dt-info" class="tab-panel active">
      <div class="form-grid">
        ${infoRow('🪪 NIK',s.nik||'-')}
        ${infoRow('🚻 Jenis Kelamin',s.jenisKelamin||'-')}
        ${infoRow('🗓️ Tgl Lahir',s.tglLahir?`${s.tempatLahir||''}, ${fmt(s.tglLahir)}`:'-')}
        ${infoRow('🕌 Agama',s.agama||'-')}
        ${infoRow('📚 Pendidikan',s.pendidikan||'-')}
        ${infoRow('📞 Telepon',s.telp||'-')}
        ${infoRow('🏠 Alamat',s.alamat||'-')}
        ${infoRow('💰 Gaji',fmtCur(s.gaji))}
      </div>
    </div>
    <div id="dt-kontrak" class="tab-panel">
      <div class="form-grid">
        ${infoRow('📅 Tanggal Masuk',fmt(s.tglMasuk))}
        ${infoRow('📋 Kontrak Berakhir',fmt(s.tglKontrakBerakhir))}
        ${infoRow('⏳ Sisa Kontrak', days<0?`<span class="text-red font-bold">Sudah berakhir ${Math.abs(days)} hari lalu</span>`:days<=30?`<span class="text-red font-bold">${days} hari lagi</span>`:days<=90?`<span class="text-yellow font-bold">${days} hari lagi</span>`:`<span class="text-green">${days} hari lagi</span>`)}
        ${infoRow('📝 Catatan',s.catatan||'-')}
      </div>
      <div class="mt-3 flex gap-2">
        <button class="btn btn-primary btn-sm" onclick="closeModal('modal-detail');openPerpanjang('${id}')">🔄 Perpanjang Kontrak</button>
        <button class="btn btn-ghost btn-sm" onclick="cetakSurat('${id}')">🖨️ Cetak Surat</button>
      </div>
    </div>
    <div id="dt-absensi" class="tab-panel">
      <div class="stats" style="grid-template-columns:repeat(2,1fr);margin-bottom:12px">
        <div class="stat" style="padding:10px"><div class="stat-ic si-green" style="width:34px;height:34px;font-size:.9rem">✅</div><div class="stat-info"><div class="stat-val" style="font-size:1.2rem">${hadir}</div><div class="stat-lbl">Hadir</div></div></div>
        <div class="stat" style="padding:10px"><div class="stat-ic si-yellow" style="width:34px;height:34px;font-size:.9rem">⏰</div><div class="stat-info"><div class="stat-val" style="font-size:1.2rem">${terlambat}</div><div class="stat-lbl">Terlambat</div></div></div>
        <div class="stat" style="padding:10px"><div class="stat-ic si-purple" style="width:34px;height:34px;font-size:.9rem">📝</div><div class="stat-info"><div class="stat-val" style="font-size:1.2rem">${izin}</div><div class="stat-lbl">Izin</div></div></div>
        <div class="stat" style="padding:10px"><div class="stat-ic si-red" style="width:34px;height:34px;font-size:.9rem">🤒</div><div class="stat-info"><div class="stat-val" style="font-size:1.2rem">${sakit}</div><div class="stat-lbl">Sakit</div></div></div>
      </div>
    </div>
  `;
  openModal('modal-detail');
}

function switchDetailTab(tab) {
  document.querySelectorAll('#detail-tabs .tab').forEach((el,i)=>{
    const tabs = ['info','kontrak','absensi'];
    el.classList.toggle('active', tabs[i]===tab);
  });
  document.querySelectorAll('#staf-detail-body .tab-panel').forEach(p=>{
    p.classList.toggle('active', p.id==='dt-'+tab);
  });
}

function infoRow(label, value) {
  return `<div style="background:var(--gray-50);padding:9px 11px;border-radius:8px">
    <div class="text-xs text-muted">${label}</div>
    <div class="text-sm font-semi" style="margin-top:2px">${value}</div>
  </div>`;
}

function exportStafCSV() {
  if (!DB.staf.length) { toast('Tidak ada data untuk diekspor', 'warn'); return; }
  const h = ['ID', 'Nama', 'NIK', 'Jabatan', 'Jenis Kelamin', 'Pendidikan', 'Tgl Masuk', 'Kontrak Berakhir', 'Status', 'Gaji', 'Telepon', 'Alamat'];
  const rows = DB.staf.map(s => [s.id, s.nama, s.nik, s.jabatan, s.jenisKelamin, s.pendidikan, s.tglMasuk, s.tglKontrakBerakhir, s.status, s.gaji, s.telp, s.alamat]);
  csvDownload([h, ...rows], `data_staf_mbg_${todayStr()}.csv`);
  toast('Data staf diekspor ke CSV (Format Kolom Excel Rapi) ✅');
}

function exportStafExcel() {
  if (!DB.staf.length) { toast('Tidak ada data untuk diekspor', 'warn'); return; }
  const h = ['No', 'ID Staf', 'Nama Lengkap', 'NIK (KTP)', 'Jabatan', 'Pendidikan', 'Tgl Masuk', 'Kontrak Berakhir', 'Status', 'Gaji Pokok', 'No. Telepon', 'Alamat'];
  const rows = DB.staf.map((s, idx) => [idx + 1, s.id, s.nama, s.nik, s.jabatan, s.pendidikan, s.tglMasuk, s.tglKontrakBerakhir, s.status, s.gaji, s.telp, s.alamat]);
  exportToExcelTable('DATABASE SELURUH STAF MBG', h, rows, `Database_Staf_MBG_${todayStr()}.xls`);
  toast('Data staf berhasil diekspor ke Excel (.xls) 📗');
}

// ============================================================
//  ABSENSI MODULE
// ============================================================
let _absDate = todayStr();

function renderAbsensi() {
  document.getElementById('abs-date').value = _absDate;
  document.getElementById('abs-day-label').textContent = new Date(_absDate+'T00:00:00').toLocaleDateString('id-ID',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  renderAbsensiTable();
}

function renderAbsensiTable() {
  const aktif = DB.staf.filter(s=>s.status==='Aktif');
  const ex = {};
  DB.absensi.filter(a=>a.tanggal===_absDate).forEach(a=>{ ex[a.stafId]=a; });

  let h=0,t=0,i=0,sk=0,al=0,bl=0;
  document.getElementById('abs-tbody').innerHTML = aktif.map(s=>{
    const r = ex[s.id];
    const st = r?.status||'';
    if (st==='Hadir') h++;
    else if (st==='Terlambat') t++;
    else if (st==='Izin') i++;
    else if (st==='Sakit') sk++;
    else if (st==='Alpha') al++;
    else bl++;
    const disabled = !st||st==='Izin'||st==='Sakit'||st==='Alpha';
    return `<tr>
      <td><span class="td-id">${s.id}</span></td>
      <td>
        <div class="flex items-center gap-2">
          <div class="avatar" style="background:${avatarColor(s.nama)}">${s.nama.charAt(0)}</div>
          <span class="td-name">${esc(s.nama)}</span>
        </div>
      </td>
      <td class="text-sm text-muted">${esc(s.jabatan)}</td>
      <td>
        <select class="fselect" style="font-size:.78rem;padding:4px 8px" onchange="setAbs('${s.id}','status',this.value)">
          <option value="" ${!st?'selected':''}>— Belum —</option>
          ${STATUS_ABS.map(x=>`<option value="${x}" ${st===x?'selected':''}>${x}</option>`).join('')}
        </select>
      </td>
      <td><input type="time" class="ctrl" style="width:106px;padding:4px 8px;font-size:.78rem" value="${r?.jamMasuk||''}" onchange="setAbs('${s.id}','jamMasuk',this.value)" ${disabled?'disabled':''}></td>
      <td><input type="time" class="ctrl" style="width:106px;padding:4px 8px;font-size:.78rem" value="${r?.jamKeluar||''}" onchange="setAbs('${s.id}','jamKeluar',this.value)" ${disabled?'disabled':''}></td>
      <td><input type="text" class="ctrl" style="font-size:.78rem;padding:4px 8px" placeholder="Keterangan..." value="${esc(r?.keterangan||'')}" onchange="setAbs('${s.id}','keterangan',this.value)"></td>
    </tr>`;
  }).join('');

  if (!aktif.length) {
    document.getElementById('abs-tbody').innerHTML = `<tr><td colspan="7">${emptyHtml('👥','Belum ada staf aktif. Tambahkan staf terlebih dahulu.')}</td></tr>`;
  }

  set('abs-h',h); set('abs-t',t); set('abs-i',i); set('abs-sk',sk); set('abs-al',al); set('abs-bl',bl);
}

function setAbs(stafId, field, value) {
  const id = `${stafId}|${_absDate}`;
  let r = DB.absensi.find(a=>a.id===id);
  if (!r) { r={id,stafId,tanggal:_absDate,status:'',jamMasuk:'',jamKeluar:'',keterangan:''}; DB.absensi.push(r); }
  r[field] = value;
  if (field==='status' && (value==='Izin'||value==='Sakit'||value==='Alpha'||value==='')) { r.jamMasuk=''; r.jamKeluar=''; }
  saveDB();
  if (field==='status') renderAbsensiTable();
}

function saveAbsensi() { saveDB(); toast('Absensi tersimpan ✅'); }

function exportAbsensiCSV() {
  const aktif = DB.staf.filter(s => s.status === 'Aktif');
  if (!aktif.length) { toast('Tidak ada staf aktif', 'warn'); return; }
  const ex = {};
  DB.absensi.filter(a => a.tanggal === _absDate).forEach(a => { ex[a.stafId] = a; });
  const h = ['ID', 'Nama', 'Jabatan', 'Status', 'Jam Masuk', 'Jam Keluar', 'Keterangan'];
  const rows = aktif.map(s => { const r = ex[s.id] || {}; return [s.id, s.nama, s.jabatan, r.status || 'Belum', r.jamMasuk || '', r.jamKeluar || '', r.keterangan || '']; });
  csvDownload([h, ...rows], `absensi_${_absDate}.csv`);
  toast('Absensi diekspor ke CSV (Format Kolom Rapi) ✅');
}

function exportAbsensiExcel() {
  const aktif = DB.staf.filter(s => s.status === 'Aktif');
  if (!aktif.length) { toast('Tidak ada staf aktif', 'warn'); return; }
  const ex = {};
  DB.absensi.filter(a => a.tanggal === _absDate).forEach(a => { ex[a.stafId] = a; });
  const h = ['No', 'ID Staf', 'Nama Staf', 'Jabatan', 'Status Kehadiran', 'Jam Masuk', 'Jam Keluar', 'Keterangan Presensi'];
  const rows = aktif.map((s, idx) => { const r = ex[s.id] || {}; return [idx + 1, s.id, s.nama, s.jabatan, r.status || 'Belum Diisi', r.jamMasuk || '-', r.jamKeluar || '-', r.keterangan || '-']; });
  exportToExcelTable(`ABSENSI HARIAN MBG - TANGGAL ${fmt(_absDate).toUpperCase()}`, h, rows, `Absensi_MBG_${_absDate}.xls`);
  toast('Absensi harian diekspor ke Excel (.xls) 📗');
}

// ============================================================
//  REKAP BULANAN
// ============================================================
let _rekapYM = curYM();

function renderRekap() {
  document.getElementById('rekap-ym').value = _rekapYM;
  set('rekap-label', ymLabel(_rekapYM));
  renderRekapTable();
  renderRekapChart();
}

function renderRekapTable() {
  const aktif = DB.staf.filter(s=>s.status==='Aktif');
  const recs = DB.absensi.filter(a=>a.tanggal.startsWith(_rekapYM));
  const wd = workDaysInMonth(_rekapYM);

  if (!aktif.length) {
    document.getElementById('rekap-tbody').innerHTML = `<tr><td colspan="10">${emptyHtml('👥','Belum ada staf aktif')}</td></tr>`;
    document.getElementById('rekap-tfoot').innerHTML='';
    return;
  }

  let totH=0,totT=0,totI=0,totSk=0,totAl=0;
  document.getElementById('rekap-tbody').innerHTML = aktif.map(s=>{
    const sr = recs.filter(a=>a.stafId===s.id);
    const h=sr.filter(a=>a.status==='Hadir').length;
    const t=sr.filter(a=>a.status==='Terlambat').length;
    const iz=sr.filter(a=>a.status==='Izin').length;
    const sk=sr.filter(a=>a.status==='Sakit').length;
    const al=sr.filter(a=>a.status==='Alpha').length;
    const eff=h+t;
    const pct=wd?Math.round(eff/wd*100):0;
    totH+=h;totT+=t;totI+=iz;totSk+=sk;totAl+=al;
    const pc=pct>=90?'pf-green':pct>=75?'pf-yellow':'pf-red';
    const bc=pct>=90?'bg':pct>=75?'by':'br';
    const kinerja=pct>=90?'Baik':pct>=75?'Cukup':'Kurang';
    return `<tr>
      <td><span class="td-id">${s.id}</span></td>
      <td><span class="td-name">${esc(s.nama)}</span></td>
      <td class="text-sm text-muted">${esc(s.jabatan)}</td>
      <td class="text-center"><span class="badge bg">${h}</span></td>
      <td class="text-center"><span class="badge by">${t}</span></td>
      <td class="text-center"><span class="badge bp">${iz}</span></td>
      <td class="text-center"><span class="badge br">${sk}</span></td>
      <td class="text-center"><span class="badge bgr">${al}</span></td>
      <td style="min-width:120px">
        <div class="flex items-center gap-2">
          <div class="prog flex-1"><div class="prog-fill ${pc}" style="width:${pct}%"></div></div>
          <span class="text-xs font-semi">${pct}%</span>
        </div>
      </td>
      <td><span class="badge ${bc}">${kinerja}</span></td>
    </tr>`;
  }).join('');

  document.getElementById('rekap-tfoot').innerHTML = `<tr>
    <td colspan="3" style="padding:9px 13px">TOTAL</td>
    <td class="text-center">${totH}</td><td class="text-center">${totT}</td>
    <td class="text-center">${totI}</td><td class="text-center">${totSk}</td>
    <td class="text-center">${totAl}</td><td colspan="2"></td>
  </tr>`;
}

function renderRekapChart() {
  const ctx = document.getElementById('c-rekap');
  if (!ctx || typeof Chart === 'undefined') return;
  try {
    if (ctx._ch) ctx._ch.destroy();
    const aktif = DB.staf.filter(s=>s.status==='Aktif');
    const recs = DB.absensi.filter(a=>a.tanggal.startsWith(_rekapYM));
    if (!aktif.length) return;
    ctx._ch = new Chart(ctx, {
      type:'bar',
      data:{
        labels: aktif.map(s=>s.nama.split(' ')[0]),
        datasets:[
          {label:'Hadir',data:aktif.map(s=>recs.filter(a=>a.stafId===s.id&&a.status==='Hadir').length),backgroundColor:'#16a34a',borderRadius:4},
          {label:'Terlambat',data:aktif.map(s=>recs.filter(a=>a.stafId===s.id&&a.status==='Terlambat').length),backgroundColor:'#d97706',borderRadius:4},
          {label:'Izin',data:aktif.map(s=>recs.filter(a=>a.stafId===s.id&&a.status==='Izin').length),backgroundColor:'#7c3aed',borderRadius:4},
          {label:'Sakit',data:aktif.map(s=>recs.filter(a=>a.stafId===s.id&&a.status==='Sakit').length),backgroundColor:'#dc2626',borderRadius:4},
        ]
      },
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12}}},
        scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,beginAtZero:true,ticks:{stepSize:1}}}
      }
    });
  } catch (e) { console.warn(e); }
}

function exportRekapCSV() {
  const aktif = DB.staf.filter(s => s.status === 'Aktif');
  if (!aktif.length) { toast('Tidak ada data', 'warn'); return; }
  const recs = DB.absensi.filter(a => a.tanggal.startsWith(_rekapYM));
  const wd = workDaysInMonth(_rekapYM);
  const h = ['ID', 'Nama', 'Jabatan', 'Hadir', 'Terlambat', 'Izin', 'Sakit', 'Alpha', 'Hari Kerja', '% Kehadiran', 'Kinerja'];
  const rows = aktif.map(s => {
    const sr = recs.filter(a => a.stafId === s.id);
    const hd = sr.filter(a => a.status === 'Hadir').length;
    const t = sr.filter(a => a.status === 'Terlambat').length;
    const iz = sr.filter(a => a.status === 'Izin').length;
    const sk = sr.filter(a => a.status === 'Sakit').length;
    const al = sr.filter(a => a.status === 'Alpha').length;
    const pct = wd ? Math.round((hd + t) / wd * 100) : 0;
    return [s.id, s.nama, s.jabatan, hd, t, iz, sk, al, wd, pct + '%', pct >= 90 ? 'Baik' : pct >= 75 ? 'Cukup' : 'Kurang'];
  });
  csvDownload([h, ...rows], `rekap_${_rekapYM}.csv`);
  toast('Rekap diekspor ke CSV (Format Kolom Rapi) ✅');
}

function exportRekapExcel() {
  const aktif = DB.staf.filter(s => s.status === 'Aktif');
  if (!aktif.length) { toast('Tidak ada data', 'warn'); return; }
  const recs = DB.absensi.filter(a => a.tanggal.startsWith(_rekapYM));
  const wd = workDaysInMonth(_rekapYM);
  const h = ['No', 'ID Staf', 'Nama Staf', 'Jabatan', 'Hadir', 'Terlambat', 'Izin', 'Sakit', 'Alpha', 'Total Hari Kerja', '% Kehadiran', 'Kinerja'];
  const rows = aktif.map((s, idx) => {
    const sr = recs.filter(a => a.stafId === s.id);
    const hd = sr.filter(a => a.status === 'Hadir').length;
    const t = sr.filter(a => a.status === 'Terlambat').length;
    const iz = sr.filter(a => a.status === 'Izin').length;
    const sk = sr.filter(a => a.status === 'Sakit').length;
    const al = sr.filter(a => a.status === 'Alpha').length;
    const pct = wd ? Math.round((hd + t) / wd * 100) : 0;
    return [idx + 1, s.id, s.nama, s.jabatan, hd, t, iz, sk, al, wd, pct + '%', pct >= 90 ? 'Baik' : pct >= 75 ? 'Cukup' : 'Kurang'];
  });
  exportToExcelTable(`REKAPITULASI PRESENSI BULAN ${ymLabel(_rekapYM).toUpperCase()}`, h, rows, `Rekap_Presensi_MBG_${_rekapYM}.xls`);
  toast('Rekapitulasi bulanan diekspor ke Excel (.xls) 📗');
}

// ============================================================
//  KONTRAK MODULE
// ============================================================
let _perpanjangId = null;

function renderKontrak() {
  renderKontrakStats();
  renderKontrakTable();
  renderKontrakChart();
}

function renderKontrakStats() {
  const aktif = DB.staf.filter(s=>s.status==='Aktif');
  set('ks-7', aktif.filter(s=>daysUntil(s.tglKontrakBerakhir)<=7&&daysUntil(s.tglKontrakBerakhir)>=0).length);
  set('ks-30', aktif.filter(s=>daysUntil(s.tglKontrakBerakhir)<=30&&daysUntil(s.tglKontrakBerakhir)>=0).length);
  set('ks-90', aktif.filter(s=>daysUntil(s.tglKontrakBerakhir)<=90&&daysUntil(s.tglKontrakBerakhir)>=0).length);
  set('ks-exp', DB.staf.filter(s=>daysUntil(s.tglKontrakBerakhir)<0).length);
}

function renderKontrakTable() {
  const filter = document.getElementById('kf-select')?.value||'all';
  let data = [...DB.staf];
  if (filter==='7') data=data.filter(s=>s.status==='Aktif'&&daysUntil(s.tglKontrakBerakhir)<=7&&daysUntil(s.tglKontrakBerakhir)>=0);
  else if (filter==='30') data=data.filter(s=>s.status==='Aktif'&&daysUntil(s.tglKontrakBerakhir)<=30&&daysUntil(s.tglKontrakBerakhir)>=0);
  else if (filter==='90') data=data.filter(s=>s.status==='Aktif'&&daysUntil(s.tglKontrakBerakhir)<=90&&daysUntil(s.tglKontrakBerakhir)>=0);
  else if (filter==='exp') data=data.filter(s=>daysUntil(s.tglKontrakBerakhir)<0);
  else data=data.filter(s=>s.status==='Aktif');
  data.sort((a,b)=>daysUntil(a.tglKontrakBerakhir)-daysUntil(b.tglKontrakBerakhir));

  const tbody = document.getElementById('kontrak-tbody');
  if (!data.length) {
    tbody.innerHTML=`<tr><td colspan="7">${emptyHtml('✅','Tidak ada kontrak dalam kategori ini')}</td></tr>`; return;
  }
  tbody.innerHTML = data.map(s=>{
    const d=daysUntil(s.tglKontrakBerakhir);
    let badge;
    if (d<0) badge=`<span class="badge br">Sudah Berakhir</span>`;
    else if (d<=7) badge=`<span class="badge br">🚨 Kritis (${d} hari)</span>`;
    else if (d<=30) badge=`<span class="badge br">⚠️ Mendesak (${d} hari)</span>`;
    else if (d<=90) badge=`<span class="badge by">⏳ Perhatian (${d} hari)</span>`;
    else badge=`<span class="badge bg">✅ Aman</span>`;
    const tgl1=new Date(s.tglMasuk+'T00:00:00'),tgl2=new Date(s.tglKontrakBerakhir+'T00:00:00');
    const tahun=Math.round((tgl2-tgl1)/(365*86400000)*10)/10;
    return `<tr>
      <td><span class="td-id">${s.id}</span></td>
      <td>
        <div class="flex items-center gap-2">
          <div class="avatar" style="background:${avatarColor(s.nama)}">${s.nama.charAt(0)}</div>
          <span class="td-name">${esc(s.nama)}</span>
        </div>
      </td>
      <td class="text-sm">${esc(s.jabatan)}</td>
      <td class="text-sm">${fmtShort(s.tglMasuk)}</td>
      <td class="text-sm">${fmtShort(s.tglKontrakBerakhir)}</td>
      <td>${badge}</td>
      <td>
        <div class="flex gap-1">
          <button class="btn btn-primary btn-xs" onclick="openPerpanjang('${s.id}')">🔄 Perpanjang</button>
          <button class="btn btn-ghost btn-xs" onclick="cetakSurat('${s.id}')">🖨️ Cetak</button>
        </div>
      </td>
    </tr>`;
  }).join('');
}

function renderKontrakChart() {
  const ctx = document.getElementById('c-kontrak');
  if (!ctx || typeof Chart === 'undefined') return;
  try {
    if (ctx._ch) ctx._ch.destroy();
    const aktif = DB.staf.filter(s=>s.status==='Aktif');
    if (!aktif.length) return;
    const months={};
    aktif.forEach(s=>{ const ym=s.tglKontrakBerakhir?.slice(0,7)||''; if(ym) months[ym]=(months[ym]||0)+1; });
    const sorted=Object.entries(months).sort((a,b)=>a[0].localeCompare(b[0]));
    ctx._ch = new Chart(ctx,{
      type:'bar',
      data:{labels:sorted.map(([k])=>ymLabel(k)),datasets:[{label:'Kontrak Berakhir',data:sorted.map(([,v])=>v),
        backgroundColor:sorted.map(([k])=>daysUntil(k+'-01')<0?'#dc2626':daysUntil(k+'-01')<=90?'#d97706':'#16a34a'),borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{stepSize:1}}}}
    });
  } catch (e) { console.warn(e); }
}

function openPerpanjang(id) {
  _perpanjangId = id;
  const s = DB.staf.find(x=>x.id===id);
  if (!s) return;
  document.getElementById('pp-nama').value = s.nama;
  document.getElementById('pp-jabatan').value = s.jabatan;
  document.getElementById('pp-lama').value = s.tglKontrakBerakhir;
  const newEnd = new Date((s.tglKontrakBerakhir||todayStr())+'T00:00:00');
  newEnd.setFullYear(newEnd.getFullYear()+1);
  document.getElementById('pp-baru').value = newEnd.toISOString().slice(0,10);
  document.getElementById('pp-catatan').value = '';
  document.getElementById('pp-durasi').value = '1 Tahun';
  openModal('modal-perpanjang');
}

function hitungDurasiPerpanjang() {
  const dur = document.getElementById('pp-durasi').value;
  const lama = document.getElementById('pp-lama').value;
  if (!lama) return;
  const base = new Date(lama+'T00:00:00');
  if (dur==='6 Bulan') base.setMonth(base.getMonth()+6);
  else if (dur==='1 Tahun') base.setFullYear(base.getFullYear()+1);
  else if (dur==='2 Tahun') base.setFullYear(base.getFullYear()+2);
  else if (dur==='3 Tahun') base.setFullYear(base.getFullYear()+3);
  document.getElementById('pp-baru').value = base.toISOString().slice(0,10);
}

function savePerpanjang() {
  const newDate = document.getElementById('pp-baru').value;
  const catatan = document.getElementById('pp-catatan').value;
  if (!newDate) { toast('Masukkan tanggal kontrak baru','err'); return; }
  const idx = DB.staf.findIndex(s=>s.id===_perpanjangId);
  if (idx===-1) return;
  const lama = DB.staf[idx].tglKontrakBerakhir;
  DB.staf[idx].tglKontrakBerakhir = newDate;
  DB.staf[idx].status = 'Aktif';
  DB.kontrak_history.push({
    id: 'KH'+Date.now(), stafId: _perpanjangId, nama: DB.staf[idx].nama,
    jabatan: DB.staf[idx].jabatan, kontrakLama: lama, kontrakBaru: newDate,
    catatan, tanggal: todayStr()
  });
  saveDB(); closeModal('modal-perpanjang');
  toast(`Kontrak ${DB.staf[idx].nama} diperpanjang hingga ${fmtShort(newDate)} ✅`);
  if (_page==='kontrak') renderKontrak();
  if (_page==='dashboard') renderDash();
}

function cetakSurat(id) {
  const s = DB.staf.find(x=>x.id===id);
  if (!s) return;
  const days = daysUntil(s.tglKontrakBerakhir);
  document.getElementById('cetak-body').innerHTML = `
    <div style="font-family:Georgia,serif;padding:20px;border:2px solid #000;line-height:1.9;color:#000">
      <div style="text-align:center;margin-bottom:20px;border-bottom:2px solid #000;padding-bottom:14px">
        <div style="font-size:1.2rem;font-weight:bold;text-transform:uppercase;letter-spacing:.05em">Surat Pemberitahuan</div>
        <div style="font-size:1rem;font-weight:bold">Masa Berlaku Kontrak Kerja</div>
        <div style="font-size:.85rem">Program Makan Bergizi Gratis (MBG) — ${esc(DB.meta.sppgName||'SPPG MBG')}</div>
      </div>
      <table style="width:100%;font-size:.85rem;margin-bottom:14px">
        <tr><td style="width:160px">Nomor Surat</td><td>: ${s.id}-${new Date().getFullYear()}/MBG/${String(new Date().getMonth()+1).padStart(2,'0')}</td></tr>
        <tr><td>Tanggal</td><td>: ${fmt(todayStr())}</td></tr>
        <tr><td>Perihal</td><td>: Pemberitahuan Masa Kontrak Kerja</td></tr>
      </table>
      <p>Kepada Yth.,<br><strong>${esc(s.nama)}</strong><br>di tempat</p>
      <p style="margin-top:14px">Dengan hormat,</p>
      <p>Bersama surat ini kami sampaikan bahwa masa berlaku kontrak kerja Saudara/i tercatat sebagai berikut:</p>
      <table style="width:100%;font-size:.85rem;margin:14px 0;border:1px solid #ccc;border-collapse:collapse">
        ${[['Nama Lengkap',s.nama],['NIK',s.nik||'-'],['Jabatan',s.jabatan],['Tanggal Masuk',fmt(s.tglMasuk)],['Kontrak Berakhir',fmt(s.tglKontrakBerakhir)],['Sisa Waktu',days<0?`Sudah berakhir ${Math.abs(days)} hari lalu`:days+' hari lagi']].map(([l,v])=>`<tr><td style="padding:7px 10px;border:1px solid #ccc;font-weight:bold;width:180px">${l}</td><td style="padding:7px 10px;border:1px solid #ccc">${esc(String(v))}</td></tr>`).join('')}
      </table>
      <p>Sehubungan dengan hal tersebut, kami mohon kepada Saudara/i untuk segera menghubungi bagian administrasi guna membicarakan proses perpanjangan kontrak kerja sesuai dengan ketentuan yang berlaku.</p>
      <div style="margin-top:36px;text-align:right">
        <p>${esc(DB.meta.sppgName||'SPPG MBG')}, ${fmt(todayStr())}</p>
        <br><br><br>
        <p style="font-weight:bold;border-top:1px solid #000;display:inline-block;min-width:200px;padding-top:4px">${esc(DB.meta.adminName||'Admin MBG')}</p>
        <p style="font-size:.8rem">${esc(DB.meta.adminRole||'Administrator')}</p>
      </div>
    </div>`;
  openModal('modal-cetak');
}

function printSurat() {
  const c = document.getElementById('cetak-body').innerHTML;
  const w = window.open('','_blank');
  w.document.write(`<html><head><title>Surat Kontrak</title><style>body{padding:30px}@media print{button{display:none}}</style></head><body>${c}</body></html>`);
  w.document.close(); w.print();
}

// ============================================================
//  LAPORAN MODULE
// ============================================================
function renderLaporan() {
  renderLaporanCharts();
  renderRiwayatPerpanjang();
  renderSummaryLaporan();
}

function renderSummaryLaporan() {
  const aktif = DB.staf.filter(s=>s.status==='Aktif').length;
  const exp90 = DB.staf.filter(s=>s.status==='Aktif'&&daysUntil(s.tglKontrakBerakhir)<=90&&daysUntil(s.tglKontrakBerakhir)>=0).length;
  const ym = curYM();
  const recs = DB.absensi.filter(a=>a.tanggal.startsWith(ym));
  const hadir = recs.filter(a=>a.status==='Hadir').length;
  const terlambat = recs.filter(a=>a.status==='Terlambat').length;
  const pct = aktif&&workDaysInMonth(ym)?Math.round((hadir+terlambat)/(aktif*workDaysInMonth(ym))*100):0;
  set('lap-aktif',aktif);
  set('lap-exp90',exp90);
  set('lap-pct',pct+'%');
  set('lap-history',DB.kontrak_history.length);
}

function renderLaporanCharts() {
  const ctx = document.getElementById('c-laporan');
  if (!ctx || typeof Chart === 'undefined') return;
  try {
    if (ctx._ch) ctx._ch.destroy();
    const months=[],dH=[],dT=[],dI=[],dS=[];
    const now=new Date();
    for(let i=5;i>=0;i--){
      const d=new Date(now.getFullYear(),now.getMonth()-i,1);
      const ym=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
      months.push(d.toLocaleDateString('id-ID',{month:'short',year:'2-digit'}));
      const r=DB.absensi.filter(a=>a.tanggal.startsWith(ym));
      dH.push(r.filter(a=>a.status==='Hadir').length);
      dT.push(r.filter(a=>a.status==='Terlambat').length);
      dI.push(r.filter(a=>a.status==='Izin').length);
      dS.push(r.filter(a=>a.status==='Sakit').length);
    }
    if (!dH.some(v=>v>0)&&!dT.some(v=>v>0)&&!dI.some(v=>v>0)&&!dS.some(v=>v>0)) return;
    ctx._ch = new Chart(ctx,{
      type:'line',
      data:{labels:months,datasets:[
        {label:'Hadir',data:dH,borderColor:'#16a34a',backgroundColor:'rgba(22,163,74,.1)',tension:.3,fill:true},
        {label:'Terlambat',data:dT,borderColor:'#d97706',tension:.3},
        {label:'Izin',data:dI,borderColor:'#7c3aed',tension:.3},
        {label:'Sakit',data:dS,borderColor:'#dc2626',tension:.3},
      ]},
      options:{responsive:true,maintainAspectRatio:false,
        plugins:{legend:{position:'top',labels:{font:{size:11},boxWidth:12}}},scales:{y:{beginAtZero:true}}}
    });
  } catch (e) { console.warn(e); }
}

function renderRiwayatPerpanjang() {
  const el = document.getElementById('riwayat-body');
  if (!el) return;
  if (!DB.kontrak_history.length) {
    el.innerHTML=`<tr><td colspan="6">${emptyHtml('📋','Belum ada riwayat perpanjangan kontrak')}</td></tr>`; return;
  }
  el.innerHTML = [...DB.kontrak_history].reverse().map(h=>`<tr>
    <td><span class="td-id">${h.id}</span></td>
    <td><span class="td-name">${esc(h.nama)}</span></td>
    <td class="text-sm">${esc(h.jabatan)}</td>
    <td class="text-sm">${fmtShort(h.kontrakLama)}</td>
    <td class="text-sm">${fmtShort(h.kontrakBaru)}</td>
    <td>${fmtShort(h.tanggal)}</td>
  </tr>`).join('');
}

function exportLaporan() {
  if (!DB.kontrak_history.length) { toast('Tidak ada riwayat perpanjangan', 'warn'); return; }
  const h = ['ID Riwayat', 'Nama Staf', 'Jabatan', 'Kontrak Lama', 'Kontrak Baru', 'Tanggal Proses', 'Catatan'];
  const rows = DB.kontrak_history.map(h => [h.id, h.nama, h.jabatan, h.kontrakLama, h.kontrakBaru, h.tanggal, h.catatan || '']);
  exportToExcelTable('RIWAYAT PERPANJANGAN KONTRAK KERJA MBG', h, rows, `Riwayat_Perpanjangan_Kontrak_${todayStr()}.xls`);
  toast('Laporan riwayat perpanjangan diekspor ke Excel (.xls) 📗');
}

// ============================================================
//  PENGATURAN (SETTINGS)
// ============================================================
function renderPengaturan() {
  document.getElementById('set-sppg').value = DB.meta.sppgName||'';
  document.getElementById('set-admin').value = DB.meta.adminName||'';
  document.getElementById('set-role').value = DB.meta.adminRole||'';
}

function savePengaturan() {
  DB.meta.sppgName = document.getElementById('set-sppg').value.trim()||'SPPG MBG';
  DB.meta.adminName = document.getElementById('set-admin').value.trim()||'Administrator';
  DB.meta.adminRole = document.getElementById('set-role').value.trim()||'Admin MBG';
  saveDB(); updateHeaderMeta();
  toast('Pengaturan disimpan ✅');
}

function resetDB() {
  if (!confirm('⚠️ PERHATIAN!\n\nSemua data (staf, absensi, kontrak, laporan) akan dihapus permanen.\n\nLanjutkan?')) return;
  if (!confirm('Anda yakin ingin mereset semua data? Tindakan ini tidak dapat dibatalkan.')) return;
  localStorage.removeItem(STORE_KEY);
  DB = emptyDB();
  toast('Semua data berhasil direset','warn');
  navigate('dashboard');
}

function backupData() {
  const json = JSON.stringify(DB, null, 2);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([json],{type:'application/json'}));
  a.download = `backup_mbg_${todayStr()}.json`;
  a.click();
  toast('Backup data berhasil diunduh ✅');
}

function restoreData(input) {
  const file = input.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const parsed = JSON.parse(e.target.result);
      if (!parsed.staf || !parsed.absensi) { toast('Format file tidak valid','err'); return; }
      DB = { ...emptyDB(), ...parsed };
      saveDB();
      toast('Data berhasil dipulihkan ✅');
      navigate(_page);
    } catch(err) { toast('Gagal membaca file: '+err.message,'err'); }
  };
  reader.readAsText(file);
  input.value='';
}

// ============================================================
//  HELPERS
// ============================================================
function set(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

function setHtml(id, val) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = val;
}

function emptyHtml(icon, text, btnLabel='', btnFn='') {
  return `<div class="empty"><div class="ei">${icon}</div><div class="et">${text}</div>${btnLabel?`<div class="ea"><button class="btn btn-primary" onclick="${btnFn}">${btnLabel}</button></div>`:''}</div>`;
}

// ============================================================
//  EXCEL MANAGEMENT MODULE (BAGIAN B)
// ============================================================
let _excelTab = 'db';
let _excelRekapYM = curYM();

function switchExcelTab(tab) {
  _excelTab = tab;
  document.querySelectorAll('#excel-tabs .tab').forEach((t, i) => {
    const tabs = ['db', 'rekap', 'kontrak', 'pivot', 'rumus', 'grafik', 'audit'];
    t.classList.toggle('active', tabs[i] === tab);
  });
  document.querySelectorAll('#page-excel .tab-panel').forEach(p => {
    p.classList.toggle('active', p.id === 'ex-tab-' + tab);
  });
  renderExcelActiveSubTab();
}

function openExcelModule(tabKey, filterVal) {
  if (_page !== 'excel') {
    navigate('excel');
  }
  switchExcelTab(tabKey);
  if (filterVal && tabKey === 'kontrak') {
    const sel = document.getElementById('ex-filter-kontrak-periode');
    if (sel) {
      sel.value = filterVal;
      renderExcelKontrakTable();
    }
  }
  setTimeout(() => {
    const targetPanel = document.getElementById('ex-tab-' + tabKey) || document.getElementById('excel-tabs');
    if (targetPanel) {
      targetPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, 50);
}

function renderExcelPage() {
  const ymInput = document.getElementById('ex-rekap-ym');
  if (ymInput && !ymInput.value) {
    ymInput.value = _excelRekapYM;
    ymInput.addEventListener('change', e => {
      _excelRekapYM = e.target.value;
      renderExcelRekapTable();
      renderExcelLaporanSummary();
    });
  }

  // Populate dropdowns for simulator
  populateSimulatorStafIds();

  // Update Status Checklist 10 Tugas
  updateExcelChecklistStatus();

  // Render sub-tab
  renderExcelActiveSubTab();
}

function updateExcelChecklistStatus() {
  const totalStaf = DB.staf.length;
  const exp30 = DB.staf.filter(s => s.status === 'Aktif' && daysUntil(s.tglKontrakBerakhir) <= 30 && daysUntil(s.tglKontrakBerakhir) >= 0).length;

  setHtml('chk-b1', totalStaf ? `<span class="badge bg">${totalStaf} Staf</span>` : `<span class="badge by">Perlu Data</span>`);
  setHtml('chk-b2', totalStaf ? `<span class="badge bg">Tabel Lengkap</span>` : `<span class="badge by">Kosong</span>`);
  setHtml('chk-b3', DB.absensi.length ? `<span class="badge bg">${DB.absensi.length} Entri</span>` : `<span class="badge by">Belum Diisi</span>`);
  setHtml('chk-b4', `<span class="badge bg">Otomatis (SUM/COUNT)</span>`);
  setHtml('chk-b5', exp30 ? `<span class="badge br">${exp30} Jatuh Tempo</span>` : `<span class="badge bg">Terkontrol</span>`);
  setHtml('chk-b8', totalStaf ? `<span class="badge bg">${[...new Set(DB.staf.map(s => s.jabatan))].length} Divisi</span>` : `<span class="badge bgr">-</span>`);
}

function renderExcelActiveSubTab() {
  if (_excelTab === 'db') renderExcelStafTable();
  else if (_excelTab === 'rekap') renderExcelRekapTable();
  else if (_excelTab === 'kontrak') renderExcelKontrakTable();
  else if (_excelTab === 'pivot') renderExcelPivotTable();
  else if (_excelTab === 'rumus') runFormulaSimulator();
  else if (_excelTab === 'grafik') renderExcelLaporanSummary();
  else if (_excelTab === 'audit') runExcelDataAudit();
}

function renderExcelStafTable() {
  const tbody = document.getElementById('ex-tbody-staf');
  if (!tbody) return;
  if (!DB.staf.length) {
    tbody.innerHTML = `<tr><td colspan="11">${emptyHtml('👥', 'Belum ada data di database Excel staf.', 'Tambah Staf Sekarang', "navigate('staf')")}</td></tr>`;
    return;
  }
  tbody.innerHTML = DB.staf.map((s, idx) => {
    const days = daysUntil(s.tglKontrakBerakhir);
    let sisaLabel = days < 0 ? `<span class="text-red font-bold">Habis (${Math.abs(days)}h lalu)</span>` :
                    days <= 30 ? `<span class="text-red font-bold">${days} hari</span>` :
                    days <= 90 ? `<span class="text-yellow font-bold">${days} hari</span>` : `<span class="text-green">${days} hari</span>`;
    return `<tr>
      <td class="excel-col-head">${idx + 1}</td>
      <td><span class="td-id">${s.id}</span></td>
      <td class="td-name">${esc(s.nama)}</td>
      <td style="font-family:monospace">${s.nik || '-'}</td>
      <td>${esc(s.jabatan)}</td>
      <td>${s.pendidikan || '-'}</td>
      <td>${fmt(s.tglMasuk)}</td>
      <td>${fmt(s.tglKontrakBerakhir)}</td>
      <td>${sisaLabel}</td>
      <td><span class="badge ${s.status === 'Aktif' ? 'bg' : 'br'}">${s.status}</span></td>
      <td>${fmtCur(s.gaji)}</td>
    </tr>`;
  }).join('');
}

function renderExcelRekapTable() {
  const tbody = document.getElementById('ex-tbody-rekap');
  const tfoot = document.getElementById('ex-tfoot-rekap');
  if (!tbody || !tfoot) return;
  const aktif = DB.staf.filter(s => s.status === 'Aktif');
  const ym = _excelRekapYM || curYM();
  const recs = DB.absensi.filter(a => a.tanggal.startsWith(ym));
  const wd = workDaysInMonth(ym);

  if (!aktif.length) {
    tbody.innerHTML = `<tr><td colspan="11">${emptyHtml('📋', 'Belum ada staf aktif untuk direkapitulasi')}</td></tr>`;
    tfoot.innerHTML = '';
    return;
  }

  let sumH = 0, sumT = 0, sumI = 0, sumS = 0, sumA = 0, sumEff = 0;
  tbody.innerHTML = aktif.map((s, idx) => {
    const sr = recs.filter(a => a.stafId === s.id);
    const h = sr.filter(a => a.status === 'Hadir').length;
    const t = sr.filter(a => a.status === 'Terlambat').length;
    const i = sr.filter(a => a.status === 'Izin').length;
    const sk = sr.filter(a => a.status === 'Sakit').length;
    const al = sr.filter(a => a.status === 'Alpha').length;
    const eff = h + t;
    const pct = wd ? Math.round(eff / wd * 100) : 0;
    sumH += h; sumT += t; sumI += i; sumS += sk; sumA += al; sumEff += eff;

    return `<tr>
      <td class="excel-col-head">${idx + 1}</td>
      <td><span class="td-id">${s.id}</span></td>
      <td class="td-name">${esc(s.nama)}</td>
      <td class="text-sm text-muted">${esc(s.jabatan)}</td>
      <td class="text-center font-semi">${h}</td>
      <td class="text-center text-yellow font-semi">${t}</td>
      <td class="text-center text-blue font-semi">${i}</td>
      <td class="text-center text-red font-semi">${sk}</td>
      <td class="text-center text-muted">${al}</td>
      <td class="text-center font-bold">${eff} / ${wd} hari</td>
      <td>
        <div class="flex items-center gap-2">
          <div class="prog flex-1" style="min-width:60px">
            <div class="prog-fill ${pct >= 90 ? 'pf-green' : pct >= 75 ? 'pf-yellow' : 'pf-red'}" style="width:${pct}%"></div>
          </div>
          <span class="text-xs font-bold">${pct}%</span>
        </div>
      </td>
    </tr>`;
  }).join('');

  const avgPct = aktif.length && wd ? Math.round((sumEff / (aktif.length * wd)) * 100) : 0;
  tfoot.innerHTML = `<tr style="background:#f3f4f6;font-weight:700">
    <td colspan="4" style="text-align:right">TOTAL FORMULA =SUM()</td>
    <td class="text-center font-bold">${sumH}</td>
    <td class="text-center font-bold text-yellow">${sumT}</td>
    <td class="text-center font-bold text-blue">${sumI}</td>
    <td class="text-center font-bold text-red">${sumS}</td>
    <td class="text-center font-bold">${sumA}</td>
    <td class="text-center font-bold">${sumEff} Kehadiran</td>
    <td class="font-bold text-green">${avgPct}% Rata-rata</td>
  </tr>`;
}

function renderExcelKontrakTable() {
  const tbody = document.getElementById('ex-tbody-kontrak');
  if (!tbody) return;
  const filter = document.getElementById('ex-filter-kontrak-periode')?.value || 'all';

  let list = [...DB.staf];
  if (filter === '30') list = list.filter(s => s.status === 'Aktif' && daysUntil(s.tglKontrakBerakhir) <= 30 && daysUntil(s.tglKontrakBerakhir) >= 0);
  else if (filter === '60') list = list.filter(s => s.status === 'Aktif' && daysUntil(s.tglKontrakBerakhir) <= 60 && daysUntil(s.tglKontrakBerakhir) >= 0);
  else if (filter === '90') list = list.filter(s => s.status === 'Aktif' && daysUntil(s.tglKontrakBerakhir) <= 90 && daysUntil(s.tglKontrakBerakhir) >= 0);
  else if (filter === 'exp') list = list.filter(s => daysUntil(s.tglKontrakBerakhir) < 0);
  else list = list.filter(s => s.status === 'Aktif');

  list.sort((a, b) => daysUntil(a.tglKontrakBerakhir) - daysUntil(b.tglKontrakBerakhir));

  if (!list.length) {
    tbody.innerHTML = `<tr><td colspan="10">${emptyHtml('✅', 'Tidak ada kontrak staf dalam periode/kriteria ini.')}</td></tr>`;
    return;
  }

  tbody.innerHTML = list.map((s, idx) => {
    const days = daysUntil(s.tglKontrakBerakhir);
    let periodeTag = days < 0 ? '<span class="badge br">Kadaluarsa</span>' :
                     days <= 30 ? '<span class="badge br">Bulan Berjalan (≤30h)</span>' :
                     days <= 60 ? '<span class="badge by">Bulan Depan (≤60h)</span>' :
                     days <= 90 ? '<span class="badge by">Kuartal Ini (≤90h)</span>' : '<span class="badge bg">Masa Panjang (>90h)</span>';

    let rekomendasi = days < 0 ? 'Segera terbitkan SK baru atau nonaktifkan' :
                      days <= 30 ? 'Siapkan draf Surat Perpanjangan Kerja ke Kepala SPPG' :
                      days <= 90 ? 'Lakukan evaluasi kinerja untuk rekomendasi perpanjangan' : 'Kontrak masih berjalan aman';

    return `<tr>
      <td class="excel-col-head">${idx + 1}</td>
      <td><span class="td-id">${s.id}</span></td>
      <td class="td-name">${esc(s.nama)}</td>
      <td>${esc(s.jabatan)}</td>
      <td>${fmtShort(s.tglMasuk)}</td>
      <td>${fmtShort(s.tglKontrakBerakhir)}</td>
      <td class="font-bold ${days <= 30 ? 'text-red' : days <= 90 ? 'text-yellow' : 'text-green'}">${days < 0 ? 'Lewat ' + Math.abs(days) + ' hari' : days + ' hari lagi'}</td>
      <td>${periodeTag}</td>
      <td class="text-xs text-muted">${rekomendasi}</td>
      <td>
        <button class="btn btn-primary btn-xs" onclick="openPerpanjang('${s.id}')">🔄 Proses</button>
      </td>
    </tr>`;
  }).join('');
}

function renderExcelPivotTable() {
  const tbody = document.getElementById('ex-tbody-pivot-jab');
  const summaryBox = document.getElementById('ex-pivot-status-summary');
  if (!tbody) return;

  const total = DB.staf.length;
  const aktif = DB.staf.filter(s => s.status === 'Aktif').length;
  const nonaktif = DB.staf.filter(s => s.status === 'Nonaktif').length;

  // Pivot by Jabatan
  const groups = {};
  DB.staf.forEach(s => {
    if (!groups[s.jabatan]) groups[s.jabatan] = { aktif: 0, nonaktif: 0, total: 0 };
    if (s.status === 'Aktif') groups[s.jabatan].aktif++;
    else groups[s.jabatan].nonaktif++;
    groups[s.jabatan].total++;
  });

  const rows = Object.entries(groups).sort((a, b) => b[1].total - a[1].total);
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="5">${emptyHtml('👥', 'Belum ada data untuk dikelompokkan')}</td></tr>`;
  } else {
    tbody.innerHTML = rows.map(([jab, val]) => {
      const pct = total ? Math.round(val.total / total * 100) : 0;
      return `<tr>
        <td class="font-semi">${esc(jab)}</td>
        <td class="text-center text-green font-bold">${val.aktif}</td>
        <td class="text-center text-red font-bold">${val.nonaktif}</td>
        <td class="text-center font-bold">${val.total}</td>
        <td>
          <div class="flex items-center gap-2">
            <div class="prog flex-1"><div class="prog-fill pf-blue" style="width:${pct}%"></div></div>
            <span class="text-xs">${pct}%</span>
          </div>
        </td>
      </tr>`;
    }).join('');
  }

  // Summary status keaktifan
  if (summaryBox) {
    summaryBox.innerHTML = `
      <div class="grid-2" style="gap:10px">
        <div style="background:var(--green-light);padding:10px;border-radius:8px">
          <div class="text-xs text-muted">Staf Berstatus Aktif</div>
          <div class="text-green font-bold" style="font-size:1.4rem">${aktif} Orang</div>
          <div class="text-xs text-muted mt-1">${total ? Math.round(aktif / total * 100) : 0}% dari seluruh personil MBG</div>
        </div>
        <div style="background:var(--red-light);padding:10px;border-radius:8px">
          <div class="text-xs text-muted">Staf Berstatus Nonaktif</div>
          <div class="text-red font-bold" style="font-size:1.4rem">${nonaktif} Orang</div>
          <div class="text-xs text-muted mt-1">${total ? Math.round(nonaktif / total * 100) : 0}% dari seluruh personil MBG</div>
        </div>
      </div>
    `;
  }

  // Render canvas chart pivot
  const ctx = document.getElementById('c-ex-pivot');
  if (ctx && typeof Chart !== 'undefined') {
    try {
      if (ctx._ch) ctx._ch.destroy();
      ctx._ch = new Chart(ctx, {
        type: 'pie',
        data: {
          labels: ['Staf Aktif', 'Staf Nonaktif'],
          datasets: [{
            data: [aktif, nonaktif],
            backgroundColor: ['#16a34a', '#dc2626'],
            borderWidth: 2,
            borderColor: '#fff'
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { position: 'right' } }
        }
      });
    } catch (e) { console.warn(e); }
  }
}

function populateSimulatorStafIds() {
  const selV = document.getElementById('sim-vlookup-id');
  if (!selV) return;
  selV.innerHTML = DB.staf.length ?
    DB.staf.map(s => `<option value="${s.id}">${s.id} - ${s.nama}</option>`).join('') :
    '<option value="">(Belum ada staf)</option>';
}

function runFormulaSimulator() {
  // 1. SUM: Total gaji staf aktif
  const totalGaji = DB.staf.filter(s => s.status === 'Aktif').reduce((acc, s) => acc + (s.gaji || 0), 0);
  setHtml('res-sum', `Hasil Kalkulasi =SUM(): <strong>${fmtCur(totalGaji)}</strong> / bulan (${DB.staf.filter(s => s.status === 'Aktif').length} staf aktif)`);

  // 2. COUNTIF: Kriteria jabatan
  const crit = document.getElementById('sim-countif-crit')?.value || 'Juru Masak';
  const countJab = DB.staf.filter(s => s.jabatan === crit).length;
  const countAktifJab = DB.staf.filter(s => s.jabatan === crit && s.status === 'Aktif').length;
  const fxCountif = document.getElementById('fx-countif-str');
  if (fxCountif) fxCountif.value = `=COUNTIF(Jabatan, "${crit}")`;
  setHtml('res-countif', `Hasil =COUNTIF(): <strong>${countJab} Orang</strong> (${countAktifJab} Aktif, ${countJab - countAktifJab} Nonaktif)`);

  // 3. IF: Logika evaluasi sisa hari kontrak
  const sisaHariInput = parseInt(document.getElementById('sim-if-days')?.value ?? '25');
  let statusIf = '';
  if (sisaHariInput < 0) statusIf = '<span class="text-red font-bold">KONTRAK KADALUARSA (Wajib Diproses/Dinonaktifkan)</span>';
  else if (sisaHariInput <= 30) statusIf = '<span class="text-red font-bold">MENDESAK (≤30 Hari - Terbitkan Rekomendasi Perpanjangan)</span>';
  else if (sisaHariInput <= 90) statusIf = '<span class="text-yellow font-bold">PERHATIAN (≤90 Hari - Jadwalkan Evaluasi Kinerja)</span>';
  else statusIf = '<span class="text-green font-bold">AMAN (>90 Hari - Masa Kerja Masih Berjalan)</span>';
  setHtml('res-if', `Hasil Evaluasi =IF(): ${statusIf}`);

  // 4. VLOOKUP: Pencarian berdasarkan ID
  const selectedId = document.getElementById('sim-vlookup-id')?.value;
  const foundStaf = DB.staf.find(s => s.id === selectedId);
  const fxVlookup = document.getElementById('fx-vlookup-str');
  if (fxVlookup) fxVlookup.value = `=VLOOKUP("${selectedId || 'STF001'}", DataStaf!A:K, 2, FALSE)`;
  if (foundStaf) {
    setHtml('res-vlookup', `Hasil =VLOOKUP(): <strong>${esc(foundStaf.nama)}</strong> | Jabatan: <strong>${esc(foundStaf.jabatan)}</strong> | Gaji: <strong>${fmtCur(foundStaf.gaji)}</strong> | Status: <strong>${foundStaf.status}</strong>`);
  } else {
    setHtml('res-vlookup', `Hasil =VLOOKUP(): <em>#N/A (Data staf tidak ditemukan di tabel)</em>`);
  }

  // 5. XLOOKUP: Pencarian fleksibel berdasarkan nama
  const qNama = document.getElementById('sim-xlookup-input')?.value?.trim().toLowerCase() || '';
  const matchStaf = qNama ? DB.staf.find(s => s.nama.toLowerCase().includes(qNama)) : null;
  const fxXlookup = document.getElementById('fx-xlookup-str');
  if (fxXlookup) fxXlookup.value = `=XLOOKUP("${qNama || 'Nama'}", DataStaf[Nama], DataStaf[Jabatan], "#N/A")`;
  if (matchStaf) {
    setHtml('res-xlookup', `Hasil =XLOOKUP(): Ditemukan <strong>${esc(matchStaf.nama)}</strong> (${matchStaf.id}) - Jabatan: <strong>${esc(matchStaf.jabatan)}</strong> | Masuk: ${fmtShort(matchStaf.tglMasuk)} | Kontrak: ${fmtShort(matchStaf.tglKontrakBerakhir)}`);
  } else if (qNama) {
    setHtml('res-xlookup', `Hasil =XLOOKUP(): <span class="text-muted">"Tidak Ditemukan" (Kriteria nama "${esc(qNama)}" tidak cocok dengan staf manapun)</span>`);
  } else {
    setHtml('res-xlookup', `Ketikkan nama staf di atas untuk simulasi XLOOKUP secara langsung`);
  }
}

function renderExcelLaporanSummary() {
  const box = document.getElementById('ex-laporan-summary-box');
  const ym = _excelRekapYM || curYM();
  const aktif = DB.staf.filter(s => s.status === 'Aktif').length;
  const total = DB.staf.length;
  const recs = DB.absensi.filter(a => a.tanggal.startsWith(ym));
  const hadir = recs.filter(a => a.status === 'Hadir').length;
  const terlambat = recs.filter(a => a.status === 'Terlambat').length;
  const izin = recs.filter(a => a.status === 'Izin').length;
  const sakit = recs.filter(a => a.status === 'Sakit').length;
  const alpha = recs.filter(a => a.status === 'Alpha').length;
  const exp30 = DB.staf.filter(s => s.status === 'Aktif' && daysUntil(s.tglKontrakBerakhir) <= 30 && daysUntil(s.tglKontrakBerakhir) >= 0).length;

  if (box) {
    box.innerHTML = `
      <div style="font-weight:700;font-size:.9rem;color:#107c41;margin-bottom:8px">Laporan Bulanan Periode: ${ymLabel(ym)}</div>
      <table style="width:100%;font-size:.8rem;border-collapse:collapse">
        <tr><td style="padding:4px 0">Instansi / Unit SPPG</td><td>: <strong>${esc(DB.meta.sppgName || 'SPPG MBG')}</strong></td></tr>
        <tr><td style="padding:4px 0">Administrator Pengelola</td><td>: <strong>${esc(DB.meta.adminName || 'Admin MBG')}</strong></td></tr>
        <tr><td style="padding:4px 0">Total Personil Terdaftar</td><td>: <strong>${total} Orang</strong> (${aktif} Aktif, ${total - aktif} Nonaktif)</td></tr>
        <tr><td style="padding:4px 0">Akumulasi Hari Hadir</td><td>: <strong>${hadir} Presensi Hadir</strong></td></tr>
        <tr><td style="padding:4px 0">Akumulasi Keterlambatan</td><td>: <strong class="text-yellow">${terlambat} Kali</strong></td></tr>
        <tr><td style="padding:4px 0">Izin / Sakit / Alpha</td><td>: ${izin} Izin, ${sakit} Sakit, ${alpha} Alpha</td></tr>
        <tr><td style="padding:4px 0">Urgensi Kontrak (≤30 Hari)</td><td>: <strong class="${exp30 ? 'text-red' : 'text-green'}">${exp30} Orang Staf</strong></td></tr>
      </table>
    `;
  }

  // Chart Laporan
  const ctx = document.getElementById('c-ex-laporan');
  if (ctx && typeof Chart !== 'undefined') {
    try {
      if (ctx._ch) ctx._ch.destroy();
      ctx._ch = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: ['Hadir', 'Terlambat', 'Izin', 'Sakit', 'Alpha'],
          datasets: [{
            label: 'Total Akumulasi Bulan ' + ym,
            data: [hadir, terlambat, izin, sakit, alpha],
            backgroundColor: ['#16a34a', '#d97706', '#7c3aed', '#dc2626', '#6b7280'],
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
        }
      });
    } catch (e) { console.warn(e); }
  }
}

function runExcelDataAudit() {
  const container = document.getElementById('ex-audit-container');
  if (!container) return;

  const checks = [];

  // Check 1: NIK 16 Digit
  const invalidNik = DB.staf.filter(s => !s.nik || s.nik.length !== 16);
  checks.push({
    title: 'Validasi Format NIK (Wajib 16 Digit Angka)',
    desc: invalidNik.length ? `${invalidNik.length} staf memiliki format NIK tidak standar: ${invalidNik.map(s => s.nama).slice(0, 3).join(', ')}` : 'Semua staf memiliki NIK 16 digit lengkap dan valid.',
    status: invalidNik.length ? 'warn' : 'ok',
    badge: invalidNik.length ? `${invalidNik.length} Anomali` : 'Akurat'
  });

  // Check 2: Duplikasi NIK
  const nikCounts = {};
  DB.staf.forEach(s => { if (s.nik) nikCounts[s.nik] = (nikCounts[s.nik] || 0) + 1; });
  const dupNiks = Object.entries(nikCounts).filter(([, c]) => c > 1);
  checks.push({
    title: 'Pemeriksaan Data Ganda (Duplikasi NIK)',
    desc: dupNiks.length ? `Ditemukan ${dupNiks.length} NIK yang terdaftar ganda di database!` : 'Tidak ada data duplikasi NIK. Seluruh entri unik.',
    status: dupNiks.length ? 'err' : 'ok',
    badge: dupNiks.length ? 'Duplikat!' : 'Bersih'
  });

  // Check 3: Kelengkapan Tanggal Masuk & Kontrak Berakhir
  const missingDates = DB.staf.filter(s => !s.tglMasuk || !s.tglKontrakBerakhir);
  checks.push({
    title: 'Kelengkapan Periode Kontrak Kerja (Tgl Masuk & Selesai)',
    desc: missingDates.length ? `${missingDates.length} staf belum dilengkapi tanggal kontrak.` : 'Seluruh staf telah memiliki tanggal masuk dan berakhir yang lengkap.',
    status: missingDates.length ? 'warn' : 'ok',
    badge: missingDates.length ? `${missingDates.length} Kosong` : 'Lengkap'
  });

  // Check 4: Kontrak Lewat Batas (Kadaluarsa)
  const expiredStaff = DB.staf.filter(s => s.status === 'Aktif' && daysUntil(s.tglKontrakBerakhir) < 0);
  checks.push({
    title: 'Konsistensi Status Staf & Masa Berlaku Kontrak',
    desc: expiredStaff.length ? `Terdapat ${expiredStaff.length} staf aktif yang masa kontraknya telah kadaluarsa!` : 'Seluruh staf aktif memiliki masa kontrak yang masih berlaku.',
    status: expiredStaff.length ? 'err' : 'ok',
    badge: expiredStaff.length ? `${expiredStaff.length} Perlu SK Baru` : 'Konsisten'
  });

  container.innerHTML = `
    <div style="display:flex;flex-direction:column;gap:10px">
      ${checks.map(c => `
        <div class="audit-check-item">
          <div>
            <div class="font-semi text-sm">${c.title}</div>
            <div class="text-xs text-muted mt-1">${c.desc}</div>
          </div>
          <span class="badge ${c.status === 'ok' ? 'bg' : c.status === 'warn' ? 'by' : 'br'}">${c.badge}</span>
        </div>
      `).join('')}
    </div>
  `;
}

// ============================================================
//  EXCEL WORKBOOK & SHEET EXPORTER (.xls HTML Table standard)
// ============================================================
function downloadAllExcelWorkbook() {
  if (!DB.staf.length) {
    toast('Belum ada data staf untuk diekspor ke Excel', 'warn');
    return;
  }

  const sppg = DB.meta.sppgName || 'SPPG MBG';
  const admin = DB.meta.adminName || 'Admin MBG';
  const ym = _excelRekapYM || curYM();
  const recs = DB.absensi.filter(a => a.tanggal.startsWith(ym));
  const wd = workDaysInMonth(ym);

  if (typeof XLSX === 'undefined') {
    toast('Library Excel sedang disiapkan, mohon coba lagi.', 'warn');
    return;
  }

  const wb = XLSX.utils.book_new();

  // Helper auto-fit column width
  const fitCols = (headers, rows) => headers.map((h, i) => {
    let max = String(h).length;
    rows.forEach(r => {
      const v = r[i] !== null && r[i] !== undefined ? String(r[i]) : '';
      if (v.length > max) max = v.length;
    });
    return { wch: Math.min(Math.max(max + 4, 12), 40) };
  });

  // SHEET 1: DATABASE STAF
  const h1 = ['No', 'ID Staf', 'Nama Lengkap', 'NIK (KTP)', 'Jabatan', 'Pendidikan', 'Tgl Masuk', 'Kontrak Berakhir', 'Status', 'Gaji Pokok', 'No. Telepon', 'Alamat'];
  const r1 = DB.staf.map((s, idx) => [
    idx + 1, s.id, s.nama, String(s.nik || '-'), s.jabatan, s.pendidikan || '-', s.tglMasuk || '-', s.tglKontrakBerakhir || '-', s.status, s.gaji || 0, String(s.telp || '-'), s.alamat || '-'
  ]);
  const ws1 = XLSX.utils.aoa_to_sheet([h1, ...r1]);
  ws1['!cols'] = fitCols(h1, r1);
  XLSX.utils.book_append_sheet(wb, ws1, 'Database Staf');

  // SHEET 2: REKAP PRESENSI BULANAN
  const h2 = ['No', 'ID Staf', 'Nama Staf', 'Jabatan', 'Hadir', 'Terlambat', 'Izin', 'Sakit', 'Alpha', 'Total Hadir Efektif', 'Hari Kerja', '% Disiplin'];
  const r2 = DB.staf.filter(s => s.status === 'Aktif').map((s, idx) => {
    const sr = recs.filter(a => a.stafId === s.id);
    const h = sr.filter(a => a.status === 'Hadir').length;
    const t = sr.filter(a => a.status === 'Terlambat').length;
    const i = sr.filter(a => a.status === 'Izin').length;
    const sk = sr.filter(a => a.status === 'Sakit').length;
    const al = sr.filter(a => a.status === 'Alpha').length;
    const eff = h + t;
    const pct = wd ? Math.round(eff / wd * 100) : 0;
    return [idx + 1, s.id, s.nama, s.jabatan, h, t, i, sk, al, eff, wd, pct + '%'];
  });
  const ws2 = XLSX.utils.aoa_to_sheet([h2, ...r2]);
  ws2['!cols'] = fitCols(h2, r2);
  XLSX.utils.book_append_sheet(wb, ws2, 'Rekap Presensi ' + ym);

  // SHEET 3: MONITORING KONTRAK
  const h3 = ['No', 'ID Staf', 'Nama Staf', 'Jabatan', 'Tgl Masuk', 'Kontrak Berakhir', 'Sisa Hari', 'Status Urgensi'];
  const r3 = DB.staf.map((s, idx) => {
    const days = daysUntil(s.tglKontrakBerakhir);
    const urg = days < 0 ? 'Kadaluarsa' : days <= 30 ? 'Mendesak (<=30 Hari)' : days <= 90 ? 'Perhatian (<=90 Hari)' : 'Aman';
    return [idx + 1, s.id, s.nama, s.jabatan, s.tglMasuk || '-', s.tglKontrakBerakhir || '-', days, urg];
  });
  const ws3 = XLSX.utils.aoa_to_sheet([h3, ...r3]);
  ws3['!cols'] = fitCols(h3, r3);
  XLSX.utils.book_append_sheet(wb, ws3, 'Monitoring Kontrak');

  // SHEET 4: PENGELOMPOKAN JABATAN
  const groups = {};
  DB.staf.forEach(s => {
    if (!groups[s.jabatan]) groups[s.jabatan] = { aktif: 0, nonaktif: 0, total: 0 };
    if (s.status === 'Aktif') groups[s.jabatan].aktif++;
    else groups[s.jabatan].nonaktif++;
    groups[s.jabatan].total++;
  });
  const h4 = ['No', 'Jabatan / Divisi', 'Jumlah Aktif', 'Jumlah Nonaktif', 'Total Personil'];
  const r4 = Object.entries(groups).map(([jab, val], idx) => [idx + 1, jab, val.aktif, val.nonaktif, val.total]);
  const ws4 = XLSX.utils.aoa_to_sheet([h4, ...r4]);
  ws4['!cols'] = fitCols(h4, r4);
  XLSX.utils.book_append_sheet(wb, ws4, 'Pengelompokan Divisi');

  // WRITE FILE .xlsx ASLI
  XLSX.writeFile(wb, `Buku_Kerja_Excel_MBG_${todayStr()}.xlsx`);
  toast('Buku Kerja Excel multi-sheet berhasil diunduh (.xlsx asli) 📗');
}

function exportExcelSheet(sheetType) {
  if (sheetType === 'db') {
    exportStafExcel();
    return;
  }
  const ym = _excelRekapYM || curYM();
  const recs = DB.absensi.filter(a => a.tanggal.startsWith(ym));
  const wd = workDaysInMonth(ym);

  if (sheetType === 'rekap') {
    exportRekapExcel();
  } else if (sheetType === 'kontrak') {
    const h = ['No', 'ID Staf', 'Nama Staf', 'Jabatan', 'Tgl Masuk', 'Kontrak Berakhir', 'Sisa Hari', 'Status Urgensi'];
    const r = DB.staf.map((s, idx) => {
      const days = daysUntil(s.tglKontrakBerakhir);
      const urg = days < 0 ? 'Kadaluarsa' : days <= 30 ? 'Mendesak (<=30 Hari)' : days <= 90 ? 'Perhatian (<=90 Hari)' : 'Aman';
      return [idx + 1, s.id, s.nama, s.jabatan, s.tglMasuk || '-', s.tglKontrakBerakhir || '-', days, urg];
    });
    exportToExcelTable(`MONITORING KONTRAK KERJA MBG`, h, r, `Monitoring_Kontrak_MBG_${todayStr()}.xlsx`);
    toast('Data kontrak berhasil diekspor ke Excel (.xlsx) 📗');
  } else if (sheetType === 'pivot-jab') {
    const groups = {};
    DB.staf.forEach(s => {
      if (!groups[s.jabatan]) groups[s.jabatan] = { aktif: 0, nonaktif: 0, total: 0 };
      if (s.status === 'Aktif') groups[s.jabatan].aktif++;
      else groups[s.jabatan].nonaktif++;
      groups[s.jabatan].total++;
    });
    const h = ['No', 'Jabatan / Divisi', 'Jumlah Aktif', 'Jumlah Nonaktif', 'Total Personil'];
    const r = Object.entries(groups).map(([jab, val], idx) => [idx + 1, jab, val.aktif, val.nonaktif, val.total]);
    exportToExcelTable(`PENGELOMPOKAN JABATAN STAF MBG`, h, r, `Pengelompokan_Jabatan_MBG_${todayStr()}.xlsx`);
    toast('Data pivot jabatan diekspor ke Excel (.xlsx) 📗');
  } else {
    downloadAllExcelWorkbook();
  }
}

// ============================================================
//  INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  // Nav items
  document.querySelectorAll('.nav-item[data-page]').forEach(el=>{
    el.addEventListener('click',()=>navigate(el.dataset.page));
  });

  // Sidebar toggle (Supports collapse on desktop & off-canvas drawer on tablet/mobile)
  const toggleBtn = document.getElementById('mob-toggle');
  if (toggleBtn) {
    toggleBtn.addEventListener('click', () => {
      if (window.innerWidth > 640) {
        document.body.classList.toggle('sidebar-collapsed');
      } else {
        document.getElementById('sidebar')?.classList.toggle('open');
        document.getElementById('mob-overlay')?.classList.toggle('open');
      }
    });
  }
  document.getElementById('mob-overlay')?.addEventListener('click', () => {
    document.getElementById('sidebar')?.classList.remove('open');
    document.getElementById('mob-overlay')?.classList.remove('open');
  });

  // Modal backdrop close
  document.querySelectorAll('.overlay').forEach(el=>{
    el.addEventListener('click',e=>{ if(e.target===el) el.classList.remove('open'); });
  });

  // Header date
  const hDate = document.getElementById('h-date');
  if (hDate) hDate.textContent = new Date().toLocaleDateString('id-ID',{weekday:'long',day:'numeric',month:'long',year:'numeric'});

  // Staf filters
  document.getElementById('sf-q')?.addEventListener('input',e=>{ _stafFilters.q=e.target.value; renderStafTable(); });
  document.getElementById('sf-jab')?.addEventListener('change',e=>{ _stafFilters.jabatan=e.target.value; renderStafTable(); });
  document.getElementById('sf-status')?.addEventListener('change',e=>{ _stafFilters.status=e.target.value; renderStafTable(); });

  // Absensi date
  document.getElementById('abs-date')?.addEventListener('change',e=>{ _absDate=e.target.value; renderAbsensiTable(); });

  // Rekap month
  document.getElementById('rekap-ym')?.addEventListener('change',e=>{ _rekapYM=e.target.value; renderRekap(); });

  // Kontrak filter
  document.getElementById('kf-select')?.addEventListener('change',()=>renderKontrakTable());

  // Perpanjang durasi
  document.getElementById('pp-durasi')?.addEventListener('change',hitungDurasiPerpanjang);

  navigate('dashboard');
});

// Explicit window bindings
window.navigate = navigate;
window.openExcelModule = openExcelModule;
window.switchExcelTab = switchExcelTab;
window.downloadAllExcelWorkbook = downloadAllExcelWorkbook;
window.exportExcelSheet = exportExcelSheet;
window.exportCsvStaf = exportCsvStaf;
window.exportRekapExcel = exportRekapExcel;
window.openModal = openModal;
window.closeModal = closeModal;
window.saveStaf = saveStaf;
window.openAddStaf = openAddStaf;
window.editStaf = editStaf;
window.deleteStaf = deleteStaf;
window.detailStaf = detailStaf;
window.perpanjangKontrak = perpanjangKontrak;
window.savePerpanjang = savePerpanjang;
window.cetakSurat = cetakSurat;
window.printSurat = printSurat;
window.savePengaturan = savePengaturan;
window.backupData = backupData;
window.restoreData = restoreData;
window.resetDB = resetDB;
window.runFormulaSimulator = runFormulaSimulator;
window.runExcelDataAudit = runExcelDataAudit;
