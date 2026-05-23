// === DIGIMAP ROXY SQUARE JEMBER — GOOGLE SHEETS BACKEND ===
// Versi: HtmlService + Google Sheets sebagai database (tanpa Firebase).
// Deploy: Extensions > Apps Script > Deploy > Web app
//   Execute as: Me  |  Who has access: Anyone (atau Anyone with the link)

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_MEMBERS  = 'Members';
const SHEET_SALES    = 'Sales';
const SHEET_SETTINGS = 'Settings';
const HEAD_MEMBERS  = ['key', 'displayName', 'pin', 'isSuperUser', 'photo'];
const HEAD_SALES    = ['id', 'nama', 'tanggal', 'waktu', 'timestamp',
                       'device', 'acc', 'qoala', 'tsel', 'isat', 'xl', 'airpods'];
const HEAD_SETTINGS = ['key', 'value'];
const SUPER_USER  = 'kiki';
const DEFAULT_PIN = '1234';

// ===== ENTRY POINT =====
function doGet() {
  ensureSetup_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Digimap Roxy Square Jember')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function getSpreadsheetUrl() {
  return SS.getUrl();
}

// ===== SETUP =====
function ensureSetup_() {
  ensureSheet_(SHEET_MEMBERS,  HEAD_MEMBERS);
  ensureSheet_(SHEET_SALES,    HEAD_SALES);
  ensureSheet_(SHEET_SETTINGS, HEAD_SETTINGS);

  const members = listMembers_();
  if (!members[SUPER_USER]) {
    const sh = SS.getSheetByName(SHEET_MEMBERS);
    sh.appendRow([SUPER_USER, 'Kiki (Admin)', DEFAULT_PIN, true, '']);
  } else if (!members[SUPER_USER].isSuperUser) {
    updateMemberRaw_(SUPER_USER, { isSuperUser: true });
  }
}

function ensureSheet_(name, header) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ===== MEMBERS =====
function listMembers_() {
  const sh = ensureSheet_(SHEET_MEMBERS, HEAD_MEMBERS);
  if (sh.getLastRow() < 2) return {};
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD_MEMBERS.length).getValues();
  const out = {};
  data.forEach(r => {
    const key = (r[0] || '').toString().trim();
    if (!key) return;
    out[key] = {
      key,
      displayName: r[1] || key,
      pin: (r[2] || '').toString(),
      isSuperUser: r[3] === true || r[3] === 'TRUE' || r[3] === 'true',
      photo: r[4] || ''
    };
  });
  return out;
}

function findMemberRow_(key) {
  const sh = ensureSheet_(SHEET_MEMBERS, HEAD_MEMBERS);
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if ((keys[i][0] || '').toString().trim() === key) return i + 2;
  }
  return -1;
}

function updateMemberRaw_(key, patch) {
  const sh = ensureSheet_(SHEET_MEMBERS, HEAD_MEMBERS);
  const row = findMemberRow_(key);
  if (row < 0) throw new Error('Member tidak ditemukan');
  if ('displayName' in patch) sh.getRange(row, 2).setValue(patch.displayName);
  if ('pin'         in patch) sh.getRange(row, 3).setValue(patch.pin);
  if ('isSuperUser' in patch) sh.getRange(row, 4).setValue(!!patch.isSuperUser);
  if ('photo'       in patch) sh.getRange(row, 5).setValue(patch.photo);
}

function getMembersPublic() {
  ensureSetup_();
  const m = listMembers_();
  return Object.values(m).map(x => ({
    key: x.key,
    displayName: x.displayName,
    isSuperUser: x.isSuperUser,
    photo: x.photo
  })).sort((a, b) => a.key.localeCompare(b.key));
}

function validateLogin(payload) {
  ensureSetup_();
  const nama = (payload.nama || '').toString().trim();
  const pin  = (payload.pin  || '').toString().trim();
  if (!nama || !pin) return { ok: false, message: 'Nama dan PIN wajib diisi.' };
  if (!/^\d{4}$/.test(pin)) return { ok: false, message: 'PIN harus 4 digit angka.' };
  const m = listMembers_();
  if (!m[nama]) return { ok: false, message: 'Nama tidak ditemukan.' };
  if (m[nama].pin !== pin) return { ok: false, message: 'PIN salah.' };
  return {
    ok: true,
    member: {
      key: nama,
      displayName: m[nama].displayName,
      isSuperUser: m[nama].isSuperUser,
      photo: m[nama].photo
    }
  };
}

function addMember(payload) {
  ensureSetup_();
  const key  = (payload.key  || '').toString().trim().toLowerCase();
  const name = (payload.name || '').toString().trim();
  const pin  = (payload.pin  || '').toString().trim();
  if (!/^[a-z0-9_]+$/.test(key))  throw new Error('Username hanya huruf kecil/angka/underscore');
  if (!name)                       throw new Error('Nama wajib diisi');
  if (!/^\d{4}$/.test(pin))        throw new Error('PIN harus 4 digit');
  const m = listMembers_();
  if (m[key]) throw new Error('Username sudah dipakai');
  const sh = ensureSheet_(SHEET_MEMBERS, HEAD_MEMBERS);
  sh.appendRow([key, name, pin, false, '']);
  return { ok: true };
}

function updateMemberField(payload) {
  ensureSetup_();
  const { key, field, value } = payload;
  if (!key) throw new Error('Key kosong');
  const m = listMembers_();
  if (!m[key]) throw new Error('Member tidak ditemukan');
  const patch = {};
  if (field === 'displayName') {
    if (!value || !value.toString().trim()) throw new Error('Nama tidak boleh kosong');
    patch.displayName = value.toString().trim();
  } else if (field === 'pin') {
    if (!/^\d{4}$/.test(value || '')) throw new Error('PIN harus 4 digit');
    patch.pin = value;
  } else if (field === 'photo') {
    patch.photo = value || '';
  } else {
    throw new Error('Field tidak valid');
  }
  updateMemberRaw_(key, patch);
  return { ok: true };
}

function deleteMember(key) {
  ensureSetup_();
  if (!key) throw new Error('Key kosong');
  if (key === SUPER_USER) throw new Error('Tidak bisa hapus super user');
  const sh = ensureSheet_(SHEET_MEMBERS, HEAD_MEMBERS);
  const row = findMemberRow_(key);
  if (row < 0) throw new Error('Member tidak ditemukan');
  sh.deleteRow(row);
  return { ok: true };
}

// ===== SALES =====
function _newId_() {
  return 's_' + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function submitSale(payload) {
  ensureSetup_();
  const nama = (payload.nama || '').toString().trim();
  if (!nama) throw new Error('Session tidak valid.');
  const m = listMembers_();
  if (!m[nama]) throw new Error('Member tidak ditemukan.');

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sh = ensureSheet_(SHEET_SALES, HEAD_SALES);
    const now = new Date();
    const tz = SS.getSpreadsheetTimeZone() || 'Asia/Jakarta';
    sh.appendRow([
      _newId_(), nama,
      Utilities.formatDate(now, tz, 'dd/MM/yyyy'),
      Utilities.formatDate(now, tz, 'HH:mm:ss'),
      now.getTime(),
      Number(payload.device)  || 0,
      Number(payload.acc)     || 0,
      Number(payload.qoala)   || 0,
      Number(payload.tsel)    || 0,
      Number(payload.isat)    || 0,
      Number(payload.xl)      || 0,
      Number(payload.airpods) || 0
    ]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function _readAllSales_() {
  const sh = ensureSheet_(SHEET_SALES, HEAD_SALES);
  if (sh.getLastRow() < 2) return [];
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD_SALES.length).getValues();
  return data.map((r, i) => ({
    row: i + 2,
    id: r[0], nama: r[1], tanggal: r[2], waktu: r[3],
    timestamp: Number(r[4]) || 0,
    device:  Number(r[5])  || 0,
    acc:     Number(r[6])  || 0,
    qoala:   Number(r[7])  || 0,
    tsel:    Number(r[8])  || 0,
    isat:    Number(r[9])  || 0,
    xl:      Number(r[10]) || 0,
    airpods: Number(r[11]) || 0
  }));
}

function getHistory(payload) {
  ensureSetup_();
  const limit = Math.min(parseInt((payload || {}).limit, 10) || 20, 500);
  const filterUser = (payload || {}).filterUser || '';
  let all = _readAllSales_();
  if (filterUser) all = all.filter(x => x.nama === filterUser);
  all.sort((a, b) => b.timestamp - a.timestamp);
  return all.slice(0, limit);
}

function _findSaleRow_(id) {
  const sh = ensureSheet_(SHEET_SALES, HEAD_SALES);
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  return -1;
}

function updateSaleField(payload) {
  ensureSetup_();
  const { id, field, value } = payload;
  const colMap = { device: 6, acc: 7, qoala: 8, tsel: 9, isat: 10, xl: 11, airpods: 12 };
  if (!colMap[field]) throw new Error('Field tidak valid');
  const sh = ensureSheet_(SHEET_SALES, HEAD_SALES);
  const row = _findSaleRow_(id);
  if (row < 0) throw new Error('Data tidak ditemukan');
  sh.getRange(row, colMap[field]).setValue(Number(value) || 0);
  return { ok: true };
}

function deleteSale(id) {
  ensureSetup_();
  const row = _findSaleRow_(id);
  if (row < 0) throw new Error('Data tidak ditemukan');
  const sh = ensureSheet_(SHEET_SALES, HEAD_SALES);
  sh.deleteRow(row);
  return { ok: true };
}

function getLeaderboard() {
  ensureSetup_();
  const sales = _readAllSales_();
  const members = listMembers_();
  const leaders = {};
  sales.forEach(r => {
    const m = members[r.nama];
    if (!m || m.isSuperUser) return;
    const total = r.device + r.acc + r.qoala + r.tsel + r.isat + r.xl;
    if (!leaders[r.nama]) leaders[r.nama] = { total: 0, airpods: 0 };
    leaders[r.nama].total   += total;
    leaders[r.nama].airpods += r.airpods;
  });
  return Object.keys(leaders).map(k => ({
    name: k,
    displayName: (members[k] && members[k].displayName) || k,
    photo:       (members[k] && members[k].photo) || '',
    total:   leaders[k].total,
    airpods: leaders[k].airpods
  })).sort((a, b) => b.total - a.total);
}

function getMonthlyChart(payload) {
  ensureSetup_();
  const filterUser = (payload || {}).filterUser || '';
  const sales = _readAllSales_();
  const monthNames = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];
  const monthly = {};
  sales.forEach(r => {
    if (filterUser && r.nama !== filterUser) return;
    const d = new Date(r.timestamp);
    if (isNaN(d.getTime())) return;
    const key = monthNames[d.getMonth()] + ' ' + d.getFullYear();
    const total = r.device + r.acc + r.qoala + r.tsel + r.isat + r.xl;
    monthly[key] = (monthly[key] || 0) + total;
  });
  const labels = Object.keys(monthly);
  return { labels: labels, data: labels.map(k => monthly[k]) };
}

// ===== SETTINGS (theme / background) =====
function getSettings() {
  ensureSetup_();
  const sh = ensureSheet_(SHEET_SETTINGS, HEAD_SETTINGS);
  if (sh.getLastRow() < 2) return {};
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const out = {};
  data.forEach(r => { if (r[0]) out[r[0]] = r[1]; });
  return out;
}

function setSetting(payload) {
  ensureSetup_();
  const { key, value } = payload;
  if (!key) throw new Error('Key kosong');
  const sh = ensureSheet_(SHEET_SETTINGS, HEAD_SETTINGS);
  const last = sh.getLastRow();
  if (last >= 2) {
    const keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (keys[i][0] === key) {
        sh.getRange(i + 2, 2).setValue(value);
        return { ok: true };
      }
    }
  }
  sh.appendRow([key, value]);
  return { ok: true };
}

function deleteSetting(key) {
  ensureSetup_();
  const sh = ensureSheet_(SHEET_SETTINGS, HEAD_SETTINGS);
  const last = sh.getLastRow();
  if (last < 2) return { ok: true };
  const keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) { sh.deleteRow(i + 2); break; }
  }
  return { ok: true };
}

// ===== ADMIN EXPORT =====
function getExportRows(payload) {
  ensureSetup_();
  const filterMember = (payload || {}).filterMember || '';
  const filterMonth  = (payload || {}).filterMonth  || '';
  let sales = _readAllSales_();
  if (filterMember) sales = sales.filter(x => x.nama === filterMember);
  if (filterMonth) {
    sales = sales.filter(x => {
      const d = new Date(x.timestamp);
      const ym = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      return ym === filterMonth;
    });
  }
  sales.sort((a, b) => a.timestamp - b.timestamp);
  const members = listMembers_();
  return sales.map(r => ({
    Tanggal: r.tanggal,
    Waktu: r.waktu,
    'Nama Staff': (members[r.nama] && members[r.nama].displayName) || r.nama,
    Device: r.device,
    Accessories: r.acc,
    Qoala: r.qoala,
    Telkomsel: r.tsel,
    Indosat: r.isat,
    XL: r.xl,
    'Airpods (Qty)': r.airpods,
    Total: r.device + r.acc + r.qoala + r.tsel + r.isat + r.xl
  }));
}
