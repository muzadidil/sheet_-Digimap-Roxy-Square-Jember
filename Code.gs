// === DIGIMAP ROXY - REST API ===
const SS = SpreadsheetApp.getActiveSpreadsheet();
const HEADER_PENJUALAN = ['Waktu','Tanggal','Nama Staff','Device','Accessories','Qoala','Telkomsel','Indosat','XL','Airpods (qty)'];

function doPost(e) {
  let req = {};
  if (e.parameter && e.parameter.action) {
    req.action = e.parameter.action;
    try { Object.assign(req, JSON.parse(e.parameter.payload || '{}')); } catch(err) {}
  } else {
    try { req = JSON.parse(e.postData.contents); } catch(err) {
      return _json({ success: false, message: 'Invalid request' });
    }
  }

  try {
    switch (req.action) {
      case 'getMembers':         return _json(getMembers());
      case 'validateLogin':      return _json(validateLogin(req));
      case 'submitData':         return _json(submitData(req));
      case 'getHistory':         return _json(getHistory(req.limit));
      case 'updateTransaction':  return _json(updateTransaction(req.row, req.col, req.value));
      case 'getLeaderboard':     return _json(getLeaderboard());
      case 'getMonthlyChartData':return _json(getMonthlyChartData());
      default: return _json({ success: false, message: 'Unknown action' });
    }
  } catch (err) {
    return _json({ success: false, message: err.message });
  }
}

function doGet() {
  return _json({ status: 'ok', service: 'Digimap Roxy Square Jember API' });
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _cleanNum(val) {
  if (!val) return 0;
  return Number(val.toString().replace(/\./g, '').replace(/,/g, '')) || 0;
}

function _ensurePenjualanSheet() {
  let sheet = SS.getSheetByName('Data_Penjualan');
  if (!sheet) {
    sheet = SS.insertSheet('Data_Penjualan');
    sheet.appendRow(HEADER_PENJUALAN);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_PENJUALAN);
  }
  return sheet;
}

function getMembers() {
  const sheet = SS.getSheetByName('NAMA');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getValues().map(r => r[0]).filter(v => v !== '');
}

function validateLogin(p) {
  const nama = (p.nama || '').toString().trim();
  const pin  = (p.pin  || '').toString().trim();
  if (!nama || !pin) return { success: false, message: 'Nama dan PIN wajib diisi.' };
  if (!/^\d{4}$/.test(pin)) return { success: false, message: 'PIN harus 4 digit angka.' };

  const sheet = SS.getSheetByName('NAMA');
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Data karyawan kosong.' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (const row of data) {
    if (row[0].toString().trim() === nama && row[1].toString().trim() === pin) {
      return { success: true, message: 'Login berhasil', nama: row[0].toString().trim() };
    }
  }
  return { success: false, message: 'Nama atau PIN salah.' };
}

function submitData(p) {
  if (!p.nama) throw new Error('Session tidak valid.');
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = _ensurePenjualanSheet();
    const now = new Date();
    sheet.appendRow([
      Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss'),
      Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy'),
      p.nama,
      _cleanNum(p.device),
      _cleanNum(p.accessories),
      _cleanNum(p.qoala),
      _cleanNum(p.telkomsel),
      _cleanNum(p.indosat),
      _cleanNum(p.xl),
      Number(p.airpods) || 0
    ]);
    return { success: true, message: 'Data berhasil disimpan!' };
  } finally {
    lock.releaseLock();
  }
}

function getHistory(limit) {
  const sheet = SS.getSheetByName('Data_Penjualan');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lim = Math.min(parseInt(limit, 10) || 20, 100);
  const startRow = Math.max(2, lastRow - lim + 1);
  const numRows  = lastRow - startRow + 1;

  const data = sheet.getRange(startRow, 1, numRows, 10).getDisplayValues();
  const result = [];
  for (let i = data.length - 1; i >= 0; i--) {
    if (!data[i][2]) continue;
    result.push({
      row: startRow + i,
      waktu: data[i][0], tanggal: data[i][1], staff: data[i][2],
      device: _cleanNum(data[i][3]), acc: _cleanNum(data[i][4]),
      qoala: _cleanNum(data[i][5]), tsel: _cleanNum(data[i][6]),
      isat: _cleanNum(data[i][7]), xl: _cleanNum(data[i][8]),
      airpods: _cleanNum(data[i][9])
    });
  }
  return result;
}

function updateTransaction(row, col, value) {
  const sheet = SS.getSheetByName('Data_Penjualan');
  if (!sheet) throw new Error('Sheet tidak ditemukan.');
  const lastRow = sheet.getLastRow();
  if (row < 2 || row > lastRow) throw new Error('Baris tidak valid.');

  const colMap = { device:4, acc:5, qoala:6, tsel:7, isat:8, xl:9, airpods:10 };
  const colIdx = colMap[col];
  if (!colIdx) throw new Error('Kolom tidak valid.');

  sheet.getRange(row, colIdx).setValue(Number(value) || 0);
  return { success: true, message: 'Data berhasil dikoreksi!' };
}

function getLeaderboard() {
  const sheet = SS.getSheetByName('Data_Penjualan');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 3, sheet.getLastRow() - 1, 8).getValues();
  const leaders = {};
  data.forEach(r => {
    const staff = r[0];
    if (!staff) return;
    const total = (Number(r[1])||0)+(Number(r[2])||0)+(Number(r[3])||0)+(Number(r[4])||0)+(Number(r[5])||0)+(Number(r[6])||0);
    const airpods = Number(r[7]) || 0;
    if (!leaders[staff]) leaders[staff] = { total: 0, airpods: 0 };
    leaders[staff].total += total;
    leaders[staff].airpods += airpods;
  });

  return Object.keys(leaders)
    .map(k => ({ name: k, total: leaders[k].total, airpods: leaders[k].airpods }))
    .sort((a, b) => b.total - a.total);
}

function getMonthlyChartData() {
  const sheet = SS.getSheetByName('Data_Penjualan');
  if (!sheet || sheet.getLastRow() < 2) return { labels: [], data: [] };

  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 8).getValues();
  const monthly = {};
  const monthNames = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Ags','Sep','Okt','Nov','Des'];

  data.forEach(r => {
    const dateStr = r[0];
    if (!dateStr) return;
    const parts = dateStr.toString().split('/');
    if (parts.length >= 3) {
      const mIdx = parseInt(parts[1], 10) - 1;
      const key = monthNames[mIdx] + ' ' + parts[2];
      const total = (Number(r[2])||0)+(Number(r[3])||0)+(Number(r[4])||0)+(Number(r[5])||0)+(Number(r[6])||0)+(Number(r[7])||0);
      monthly[key] = (monthly[key] || 0) + total;
    }
  });

  const labels = Object.keys(monthly);
  return { labels, data: labels.map(k => monthly[k]) };
}