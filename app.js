/* ป้ายบอกโรค */
window.addEventListener('error', function (e) {
  const d = document.createElement('div');
  d.style.cssText = 'position:fixed;left:1rem;right:1rem;bottom:1rem;z-index:999;background:#7f1d1d;color:#fff;padding:1rem 1.2rem;border-radius:1rem;font-size:14px;white-space:pre-wrap';
  d.textContent = 'ข้อผิดพลาด: ' + (e.message || e) + ' (บรรทัด ' + (e.lineno || '-') + ')';
  document.body.appendChild(d);
});

const sb = supabase.createClient(APP_CONFIG.supabaseUrl, APP_CONFIG.supabaseAnonKey);
let allItems = [], myRole = 'viewer', me = null, editingId = null, fs = 1;
const SC = {'รอตรวจรับ':'#3b82f6','ใช้งานปกติ':'#22c55e','ถูกยืม':'#f59e0b','ชำรุด/รอซ่อม':'#ef4444','รอจำหน่าย':'#a855f7','จำหน่ายแล้ว':'#64748b','สูญหาย':'#94a3b8'};

window.onload = async function () {
  initPrefs();
  const { data } = await sb.auth.getSession();
  if (data.session) await enterApp(data.session.user);
  else document.getElementById('loginBox').style.display = 'flex';
};

/* ----- ล็อกอิน / ออก ----- */
async function doLogin(e) {
  e.preventDefault();
  const { error } = await sb.auth.signInWithPassword({
    email: document.getElementById('loginEmail').value.trim(),
    password: document.getElementById('loginPass').value
  });
  if (error) { toast('เข้าสู่ระบบไม่สำเร็จ: ' + error.message); return; }
  const { data } = await sb.auth.getSession();
  document.getElementById('loginBox').style.display = 'none';
  await enterApp(data.session.user);
}
function doLogout() { sb.auth.signOut().then(() => location.reload()); }

async function enterApp(user) {
  me = user;
  const { data } = await sb.from('user_roles').select('role').eq('user_id', user.id).maybeSingle();
  myRole = data ? data.role : 'viewer';
  document.getElementById('app').style.display = 'flex';
  document.getElementById('roleBadge').textContent = myRole === 'admin' ? 'ผู้ดูแลหลัก' : 'ผู้เข้าชม';
  document.getElementById('whoami').innerHTML = '';
  document.getElementById('whoami').append(user.email + ' · ');
  const lo = document.createElement('a');
  lo.href = '#'; lo.textContent = 'ออกจากระบบ';
  lo.onclick = ev => { ev.preventDefault(); doLogout(); };
  document.getElementById('whoami').appendChild(lo);
  document.getElementById('btnAdd').style.display = myRole === 'admin' ? '' : 'none';
  document.getElementById('btnExport').style.display = myRole === 'admin' ? '' : 'none';
  await loadInitialData();
}

/* ----- ธีม + ฟอนต์ ----- */
function initPrefs() {
  let th = localStorage.getItem('kp-theme');
  if (!th) th = matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  document.documentElement.dataset.theme = th;
  fs = parseFloat(localStorage.getItem('kp-fs')) || 1;
  applyFs();
}
function toggleTheme() {
  const nx = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  document.documentElement.dataset.theme = nx;
  localStorage.setItem('kp-theme', nx);
}
function setFont(v) { fs = Math.min(1.5, Math.max(0.8, Math.round(v * 10) / 10)); applyFs(); localStorage.setItem('kp-fs', fs); }
function applyFs() {
  document.documentElement.style.setProperty('--fs', fs);
  document.getElementById('fontLabel').textContent = Math.round(fs * 100) + '%';
}

/* ----- โหลดและวาดข้อมูล ----- */
async function loadInitialData() {
  const { data, error } = await sb.from('assets').select('*');
  if (error) { toast('โหลดข้อมูลไม่สำเร็จ: ' + error.message); return; }
  allItems = data.map(r => ({
    id: r.id, num: r.asset_number || '', name: r.name || '', cat: r.category || '',
    status: r.status || '', date: (r.purchase_date || '').slice(0, 10), year: r.year || '',
    price: r.price || '', loc: r.location || '', img: r.image_path || '', notes: r.notes || ''
  }));
  document.getElementById('statusFilter').innerHTML =
    '<option value="">ทุกสถานะ</option>' + APP_CONFIG.statuses.map(s => '<option>' + s + '</option>').join('');
  document.getElementById('fStatus').innerHTML =
    APP_CONFIG.statuses.map(s => '<option>' + s + '</option>').join('');
  const years = [...new Set(allItems.map(i => i.year).filter(Boolean))].sort().reverse();
  document.getElementById('yearFilter').innerHTML =
    '<option value="">ทุกปี</option>' + years.map(y => '<option>' + y + '</option>').join('');
  renderStats(); loadItems();
}

function renderStats() {
  const by = s => allItems.filter(i => i.status === s).length;
  document.getElementById('stats').innerHTML =
    '<div class="stat glass"><b>' + allItems.length + '</b><span>ทั้งหมด</span></div>' +
    '<div class="stat glass"><b style="color:#22c55e">' + by('ใช้งานปกติ') + '</b><span>ใช้งานปกติ</span></div>' +
    '<div class="stat glass"><b style="color:#ef4444">' + by('ชำรุด/รอซ่อม') + '</b><span>ชำรุด/รอซ่อม</span></div>' +
    '<div class="stat glass"><b style="color:var(--ac2)">' + allItems.filter(i => i.year === String(new Date().getFullYear() + 543)).length + '</b><span>ปีปัจจุบัน</span></div>';
}

function loadItems() {
  const q = document.getElementById('searchInput').value.trim().toLowerCase();
  const y = document.getElementById('yearFilter').value;
  const s = document.getElementById('statusFilter').value;
  renderList(allItems.filter(i =>
    (!q || [i.num, i.name, i.loc, i.notes].some(v => v.toLowerCase().includes(q))) &&
    (!y || i.year === y) && (!s || i.status === s)));
}

function renderList(items) {
  const el = document.getElementById('list');
  const admin = myRole === 'admin';
  if (!items.length) { el.innerHTML = '<div class="empty glass">ไม่พบข้อมูล</div>'; return; }
  el.innerHTML = '<div class="cards">' + items.map(i => {
    const c = SC[i.status] || '#8888aa';
    return '<div class="card glass"><div class="thumb">' +
      (i.img ? '<img data-path="' + i.img + '" src="">' : '<div class="noimg">📦</div>') + '</div>' +
      '<div class="cbody"><b>' + (i.num || '-') + '</b><div>' + (i.name || '-') + '</div>' +
      '<div class="meta">ปีงบ ' + (i.year || '-') + ' · ' + (i.date || '-') + '</div>' +
      '<div class="meta">📍 ' + (i.loc || 'ไม่ระบุ') + '</div>' +
      '<span class="pill" style="color:' + c + ';border-color:' + c + '66;background:' + c + '1f">' + (i.status || '-') + '</span>' +
      (admin ? '<div class="cact"><button class="btn" onclick="openForm(\'' + i.id + '\')">แก้ไข</button>' +
        '<select onchange="updateStatus(\'' + i.id + '\', this.value)">' +
        APP_CONFIG.statuses.map(s => '<option ' + (s === i.status ? 'selected' : '') + '>' + s + '</option>').join('') +
        '</select></div>' : '') +
      '</div></div>';
  }).join('') + '</div>';
  el.querySelectorAll('img[data-path]').forEach(async img => {
    const u = await imgObjectUrl(img.dataset.path);
    if (u) img.src = u; else img.outerHTML = '<div class="noimg">📦</div>';
  });
}

/* ----- คลังรูป ----- */
const imgCache = {};
async function imgObjectUrl(path) {
  if (imgCache[path]) return imgCache[path];
  const { data, error } = await sb.storage.from(APP_CONFIG.bucket).download(path);
  if (error) return '';
  const url = URL.createObjectURL(data);
  imgCache[path] = url;
  return url;
}

/* ----- ฟอร์มเพิ่ม/แก้ ----- */
function openForm(id) {
  if (myRole !== 'admin') return;
  editingId = id || null;
  const it = id ? allItems.find(x => x.id === id) : null;
  document.getElementById('formTitle').textContent = it ? 'แก้ไขครุภัณฑ์' : 'เพิ่มครุภัณฑ์';
  const set = (k, v) => document.getElementById(k).value = v || '';
  set('fNumber', it && it.num); set('fName', it && it.name); set('fCategory', it && it.cat);
  set('fStatus', it && it.status); set('fDate', it && it.date); set('fYear', it && it.year);
  set('fPrice', it && it.price); set('fLocation', it && it.loc); set('fNotes', it && it.notes);
  document.getElementById('fImage').value = '';
  document.getElementById('overlay').classList.add('open');
}
function closeForm() { document.getElementById('overlay').classList.remove('open'); }
const val = id => document.getElementById(id).value.trim();

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('itemForm').addEventListener('submit', submitForm);
  document.getElementById('searchInput').addEventListener('keydown', e => { if (e.key === 'Enter') loadItems(); });
});

async function submitForm(e) {
  e.preventDefault();
  const p = { num: val('fNumber'), name: val('fName'), cat: val('fCategory'), status: val('fStatus'),
              date: val('fDate'), year: val('fYear'), price: val('fPrice'), loc: val('fLocation'), notes: val('fNotes') };
  try {
    let imgPath = null;
    const file = document.getElementById('fImage').files[0];
    const newId = 'AST-' + Date.now();
    if (file) {
      const blob = await compressImage(file, 1600, 0.82);
      /* ชื่อไฟล์ในคลังรูปใช้เฉพาะอักษรอังกฤษ ตัวเลข _ - . เท่านั้น */
      imgPath = 'images/' + (editingId || newId) + '_' + Date.now() + '.jpg';
      const { error } = await sb.storage.from(APP_CONFIG.bucket)
        .upload(imgPath, blob, { contentType: 'image/jpeg' });
      if (error) throw error;
    }
    if (editingId) {
      const up = { asset_number: p.num, name: p.name, category: p.cat, status: p.status,
        purchase_date: p.date, year: p.year, price: p.price, location: p.loc,
        notes: p.notes, updated_at: new Date().toISOString() };
      if (imgPath) up.image_path = imgPath;
      const { error } = await sb.from('assets').update(up).eq('id', editingId);
      if (error) throw error;
    } else {
      const { error } = await sb.from('assets').insert([{
        id: newId, asset_number: p.num, name: p.name, category: p.cat,
        status: p.status || 'รอตรวจรับ', purchase_date: p.date, year: p.year, price: p.price,
        location: p.loc, image_path: imgPath, notes: p.notes
      }]);
      if (error) throw error;
    }
    toast('บันทึกเรียบร้อย ✓'); closeForm(); await loadInitialData();
  } catch (err) { toast('ผิดพลาด: ' + (err.message || err)); }
}
async function updateStatus(id, status) {
  if (myRole !== 'admin') return;
  const { error } = await sb.from('assets').update({ status: status, updated_at: new Date().toISOString() }).eq('id', id);
  if (error) { toast('ผิดพลาด: ' + error.message); return; }
  toast('อัปเดตสถานะแล้ว ✓');
  await loadInitialData();
}

/* ----- ส่งออก CSV ให้พี่เขาเปิดใน Excel ----- */
function exportCSV() {
  const head = ['id','asset_number','name','category','status','purchase_date','year','price','location','image_path','notes'];
  const esc = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  const rows = allItems.map(i => [i.id, i.num, i.name, i.cat, i.status, i.date, i.year, i.price, i.loc, i.img, i.notes].map(esc).join(','));
  const blob = new Blob(['\ufeff' + head.join(',') + '\n' + rows.join('\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'ครุภัณฑ์_' + new Date().toISOString().slice(0, 10) + '.csv';
  a.click();
}

/* ----- ตัวช่วย ----- */
function compressImage(file, maxW, q) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      const s = Math.min(1, maxW / img.width);
      const c = document.createElement('canvas');
      c.width = Math.round(img.width * s); c.height = Math.round(img.height * s);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
      c.toBlob(b => { URL.revokeObjectURL(url); b ? resolve(b) : reject(new Error('แปลงรูปไม่สำเร็จ')); }, 'image/jpeg', q);
    };
    img.onerror = () => reject(new Error('อ่านไฟล์รูปไม่สำเร็จ'));
    img.src = url;
  });
}
let toastTimer;
function toast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg; el.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 3500);
}
