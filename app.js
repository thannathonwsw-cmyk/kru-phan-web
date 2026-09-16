/* ===== 1) ล็อกอิน Microsoft ===== */
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: APP_CONFIG.clientId,
    authority: 'https://login.microsoftonline.com/' + APP_CONFIG.tenant,
    redirectUri: window.location.origin + window.location.pathname
  },
  cache: { cacheLocation: 'localStorage' }
});

let allRows = [], editingId = null, fileIdCache = '', fs = 1;
const SCOPES = ['User.Read', 'Files.ReadWrite.All'];

window.onload = async function () {
  initPrefs();
  await msalInstance.initialize();
  await msalInstance.handleRedirectPromise();
  const acc = msalInstance.getAllAccounts();
  if (!acc.length) { document.getElementById('loginBox').style.display = 'flex'; return; }
  msalInstance.setActiveAccount(acc[0]);
  document.getElementById('whoami').textContent = acc[0].name + ' · ';
  const lo = document.createElement('a');
  lo.href = '#'; lo.textContent = 'ออกจากระบบ';
  lo.onclick = e => { e.preventDefault(); msalInstance.logoutRedirect(); };
  document.getElementById('whoami').appendChild(lo);
  document.getElementById('app').style.display = 'flex';
  try { await loadInitialData(); } catch (err) { toast('โหลดข้อมูลไม่สำเร็จ: ' + err.message); }
};

function login() { return msalInstance.loginRedirect({ scopes: SCOPES }); }

async function getToken() {
  const req = { scopes: SCOPES, account: msalInstance.getActiveAccount() };
  try { return (await msalInstance.acquireTokenSilent(req)).accessToken; }
  catch (e) { return (await msalInstance.acquireTokenPopup(req)).accessToken; }
}

/* ===== 2) ธีม + ขนาดตัวอักษร ===== */
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
function setFont(v) {
  fs = Math.min(1.5, Math.max(0.8, Math.round(v * 10) / 10));
  applyFs(); localStorage.setItem('kp-fs', fs);
}
function applyFs() {
  document.documentElement.style.setProperty('--fs', fs);
  document.getElementById('fontLabel').textContent = Math.round(fs * 100) + '%';
}

/* ===== 3) ตัวคุยกับ Microsoft Graph ===== */
async function graphFetch(path, opt = {}) {
  const token = await getToken();
  const res = await fetch('https://graph.microsoft.com/v1.0/' + path, {
    ...opt,
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json', ...(opt.headers || {}) }
  });
  if (res.status === 429) { await new Promise(r => setTimeout(r, 3000)); return graphFetch(path, opt); }
  if (!res.ok) throw new Error('Graph ' + res.status + ' ' + (await res.text()));
  return res.status === 204 ? null : res.json();
}

async function getFileId() {
  if (fileIdCache) return fileIdCache;
  const p = APP_CONFIG.filePath.split('/').map(encodeURIComponent).join('/');
  const j = await graphFetch('me/drive/root:/' + p);
  fileIdCache = j.id;
  return j.id;
}
async function wb(path, opt) { return graphFetch('me/drive/items/' + (await getFileId()) + '/workbook' + path, opt); }
async function wbReadAll() {
  const j = await wb('/tables/' + APP_CONFIG.table + '/rows');
  return (j.value || []).map(r => r.values);
}
async function wbAppend(values) {
  return wb('/tables/' + APP_CONFIG.table + '/rows/add',
    { method: 'POST', body: JSON.stringify({ index: null, values: [values] }) });
}
async function wbUpdate(i, values) {
  return wb('/tables/' + APP_CONFIG.table + '/rows/itemAt(index=' + i + ')',
    { method: 'PATCH', body: JSON.stringify({ values: [values] }) });
}
async function uploadImage(blob, name) {
  const token = await getToken();
  const p = APP_CONFIG.imagePath.split('/').map(encodeURIComponent).join('/');
  const res = await fetch('https://graph.microsoft.com/v1.0/me/drive/root:/' + p + '/' + encodeURIComponent(name) + ':/content', {
    method: 'PUT',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'image/jpeg' },
    body: blob
  });
  if (!res.ok) throw new Error('อัปโหลดรูปไม่สำเร็จ ' + res.status);
  return res.json();
}
const thumbCache = {};
async function thumbUrl(itemId) {
  if (thumbCache[itemId]) return thumbCache[itemId];
  const token = await getToken();
  const base = 'https://graph.microsoft.com/v1.0/me/drive/items/' + itemId;
  let res = await fetch(base + '/thumbnails/0/medium/content', { headers: { Authorization: 'Bearer ' + token } });
  if (!res.ok) res = await fetch(base + '/content', { headers: { Authorization: 'Bearer ' + token } });
  if (!res.ok) return '';
  const url = URL.createObjectURL(await res.blob());
  thumbCache[itemId] = url;
  return url;
}

/* ===== 4) วาดหน้าจอ ===== */
function mapRows(rows) {
  return rows.map(r => ({
    id: String(r[0] || ''), num: String(r[1] || ''), name: String(r[2] || ''),
    cat: String(r[3] || ''), status: String(r[4] || ''), date: String(r[5] || '').slice(0, 10),
    year: String(r[6] || ''), price: String(r[7] || ''), loc: String(r[8] || ''),
    imgUrl: String(r[9] || ''), fileId: String(r[10] || ''), notes: String(r[11] || '')
  }));
}
const SC = {'รอตรวจรับ':'#3b82f6','ใช้งานปกติ':'#22c55e','ถูกยืม':'#f59e0b','ชำรุด/รอซ่อม':'#ef4444','รอจำหน่าย':'#a855f7','จำหน่ายแล้ว':'#64748b','สูญหาย':'#94a3b8'};

async function loadInitialData() {
  allRows = await wbReadAll();
  const items = mapRows(allRows);
  document.getElementById('statusFilter').innerHTML =
    '<option value="">ทุกสถานะ</option>' + APP_CONFIG.statuses.map(s => '<option>' + s + '</option>').join('');
  document.getElementById('fStatus').innerHTML =
    APP_CONFIG.statuses.map(s => '<option>' + s + '</option>').join('');
  const years = [...new Set(items.map(i => i.year).filter(Boolean))].sort().reverse();
  document.getElementById('yearFilter').innerHTML =
    '<option value="">ทุกปี</option>' + years.map(y => '<option>' + y + '</option>').join('');
  renderStats(items);
  loadItems();
}

function renderStats(items) {
  const by = s => items.filter(i => i.status === s).length;
  document.getElementById('stats').innerHTML =
    '<div class="stat glass"><b>' + items.length + '</b><span>ทั้งหมด</span></div>' +
    '<div class="stat glass"><b style="color:#22c55e">' + by('ใช้งานปกติ') + '</b><span>ใช้งานปกติ</span></div>' +
    '<div class="stat glass"><b style="color:#ef4444">' + by('ชำรุด/รอซ่อม') + '</b><span>ชำรุด/รอซ่อม</span></div>' +
    '<div class="stat glass"><b style="color:var(--ac2)">' + items.filter(i => i.year === String(new Date().getFullYear() + 543)).length + '</b><span>ปีปัจจุบัน</span></div>';
}

function loadItems() {
  const q = document.getElementById('searchInput').value.trim().toLowerCase();
  const y = document.getElementById('yearFilter').value;
  const s = document.getElementById('statusFilter').value;
  renderList(mapRows(allRows).filter(i =>
    (!q || [i.num, i.name, i.loc, i.notes].some(v => v.toLowerCase().includes(q))) &&
    (!y || i.year === y) && (!s || i.status === s)));
}

function renderList(items) {
  const el = document.getElementById('list');
  if (!items.length) { el.innerHTML = '<div class="empty glass">ไม่พบข้อมูล</div>'; return; }
  el.innerHTML = '<div class="cards">' + items.map(i => {
    const c = SC[i.status] || '#8888aa';
    return '<div class="card glass">' +
      '<div class="thumb">' + (i.fileId.indexOf('OD:') === 0
        ? '<img data-od="' + i.fileId.slice(3) + '" src="">'
        : (i.imgUrl ? '<img src="' + i.imgUrl + '">' : '<div class="noimg">📦</div>')) + '</div>' +
      '<div class="cbody"><b>' + (i.num || '-') + '</b><div>' + (i.name || '-') + '</div>' +
      '<div class="meta">ปีงบ ' + (i.year || '-') + ' · ' + (i.date || '-') + '</div>' +
      '<div class="meta">📍 ' + (i.loc || 'ไม่ระบุ') + '</div>' +
      '<span class="pill" style="color:' + c + ';border-color:' + c + '66;background:' + c + '1f">' + (i.status || '-') + '</span>' +
      '<div class="cact"><button class="btn" onclick="openForm(\'' + i.id + '\')">แก้ไข</button>' +
      '<select onchange="updateStatus(\'' + i.id + '\', this.value)">' +
      APP_CONFIG.statuses.map(s => '<option ' + (s === i.status ? 'selected' : '') + '>' + s + '</option>').join('') +
      '</select></div></div></div>';
  }).join('') + '</div>';
  el.querySelectorAll('img[data-od]').forEach(async img => {
    const u = await thumbUrl(img.dataset.od);
    if (u) img.src = u; else img.outerHTML = '<div class="noimg">📦</div>';
  });
}

/* ===== 5) ฟอร์มเพิ่ม/แก้ ===== */
function openForm(id) {
  editingId = id || null;
  const it = id ? mapRows(allRows).find(x => x.id === id) : null;
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
  const file = document.getElementById('fImage').files[0];
  try {
    if (file) {
      const blob = await compressImage(file, 1600, 0.82);
      const up = await uploadImage(blob, (p.num || 'asset') + '_' + Date.now() + '.jpg');
      p.fileId = 'OD:' + up.id;
    }
    if (editingId) {
      const idx = allRows.findIndex(r => String(r[0]) === editingId);
      if (idx < 0) throw new Error('ไม่พบรายการ');
      const old = allRows[idx].slice();
      old[1] = p.num || old[1]; old[2] = p.name || old[2]; old[3] = p.cat || old[3];
      old[4] = p.status || old[4]; old[5] = p.date || old[5]; old[6] = p.year || old[6];
      old[7] = p.price || old[7]; old[8] = p.loc || old[8]; old[11] = p.notes || old[11];
      if (p.fileId) { old[10] = p.fileId; old[9] = ''; }
      old[13] = new Date().toISOString();
      await wbUpdate(idx, old);
    } else {
      await wbAppend(['AST-' + Date.now(), p.num, p.name, p.cat, p.status || 'รอตรวจรับ',
        p.date, p.year, p.price, p.loc, '', p.fileId || '', p.notes,
        new Date().toISOString(), new Date().toISOString()]);
    }
    toast('บันทึกเรียบร้อย ✓');
    closeForm();
    await loadInitialData();
  } catch (err) { toast('ผิดพลาด: ' + err.message); }
}

async function updateStatus(id, status) {
  try {
    const idx = allRows.findIndex(r => String(r[0]) === id);
    if (idx < 0) return;
    const old = allRows[idx].slice();
    old[4] = status; old[13] = new Date().toISOString();
    await wbUpdate(idx, old);
    toast('อัปเดตสถานะแล้ว ✓');
    await loadInitialData();
  } catch (err) { toast('ผิดพลาด: ' + err.message); }
}

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
