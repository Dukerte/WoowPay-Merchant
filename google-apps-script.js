// ═══════════════════════════════════════════════════════════════
// WOOW PAY — Мерчант бүртгэлийн систем
// Google Apps Script — v2
//
// ТОХИРГОО ХИЙХ ЗААВАР:
//  1. Энэ файлын кодыг Google Apps Script editor-т буулгана
//  2. SETUP_RUN_ONCE() функцийг нэг удаа гараар ажиллуулна (▶ Run)
//  3. Зөвшөөрлүүдийг баталгаажуулна (Gmail, Drive, Sheets)
//  4. Web app-аа Deploy → New deployment → Web app
//     • Execute as: Me
//     • Who has access: Anyone
//  5. URL-ийг index.html-ийн GOOGLE_SHEET_URL-д хуулна
// ═══════════════════════════════════════════════════════════════

// ── ТОХИРГОО ─────────────────────────────────────────────────
const CONFIG = {
  SHEET_NAME:      'Merchants',
  DRIVE_FOLDER:    'Woow Pay - Merchant Files',
  NOTIFY_EMAILS:   ['enkhdulguun.amarbayasgalan@gmail.com', 'Munkhjinwoow@gmail.com'],
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/1KMUKsLAcEW42KG5pYzfiTMuPsM6cErYEDFCr-N-O_f4'
};

const HEADERS = [
  'Огноо', 'Лавлах дугаар', 'Статус',
  'Ургийн овог', 'Овог', 'Нэр', 'РД', 'Утас', 'Имэйл', 'Facebook',
  'Оршин суух хаяг', 'Бренд нэр', 'Худалдааны төрөл', 'Мерчант хаяг',
  'Салбарын тоо', 'Дундаж борлуулалт', 'Банкны нэр', 'Данс',
  'Урилгын код', 'Хаанаас мэдсэн', 'Нэмэлт тайлбар', 'Цахим хаяг',
  'Иргэний үнэмлэх 🔗', 'Улсын бүртгэл 🔗', 'Түрээсийн гэрээ 🔗', 'Тусгай зөвшөөрөл 🔗'
];

// ── SHEET ────────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.appendRow(HEADERS);

    const hdr = sheet.getRange(1, 1, 1, HEADERS.length);
    hdr.setFontWeight('bold')
       .setBackground('#29BDE0')
       .setFontColor('#ffffff')
       .setFontSize(11)
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');

    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 38);

    // Wider columns for key fields
    sheet.setColumnWidth(1, 165);  // Огноо
    sheet.setColumnWidth(2, 115);  // Лавлах дугаар
    sheet.setColumnWidth(3, 120);  // Статус
    sheet.setColumnWidth(9, 190);  // Имэйл
    sheet.setColumnWidth(12, 180); // Бренд нэр
    sheet.setColumnWidth(13, 210); // Худалдааны төрөл
  }
  return sheet;
}

function resetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (existing) ss.deleteSheet(existing);
  return getOrCreateSheet();
}

// ── DRIVE ────────────────────────────────────────────────────
function getOrCreateFolder(path) {
  const parts = path.split('/');
  let folder = DriveApp.getRootFolder();
  for (const part of parts) {
    const iter = folder.getFoldersByName(part);
    folder = iter.hasNext() ? iter.next() : folder.createFolder(part);
  }
  return folder;
}

function saveFilesToDrive(files, refNumber) {
  if (!files || !Object.keys(files).length) return {};

  const folder = getOrCreateFolder(CONFIG.DRIVE_FOLDER + '/' + refNumber);
  const labels = {
    id_zurag:       'Иргэний_үнэмлэх',
    uls_burtgel:    'Улсын_бүртгэл',
    turees_geree:   'Түрээсийн_гэрээ',
    tus_zovshoorul: 'Тусгай_зөвшөөрөл'
  };
  const links = {};

  for (const [key, file] of Object.entries(files)) {
    if (!file || !file.data || !file.name) continue;
    try {
      const bytes   = Utilities.base64Decode(file.data);
      const ext     = file.name.split('.').pop() || '';
      const name    = (labels[key] || key) + (ext ? '.' + ext : '');
      const blob    = Utilities.newBlob(bytes, file.type || 'application/octet-stream', name);
      const driveFile = folder.createFile(blob);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      links[key] = driveFile.getUrl();
    } catch (e) {
      Logger.log('File error [' + key + ']: ' + e.message);
    }
  }
  return links;
}

// ── EMAIL ────────────────────────────────────────────────────
function emailRow(label, value) {
  return '<tr style="border-bottom:1px solid #F0F7FA;">'
    + '<td style="padding:8px 14px;color:#5A7A8A;font-size:13px;width:38%;background:#FAFCFE;">' + label + '</td>'
    + '<td style="padding:8px 14px;color:#0F1C3F;font-size:13px;font-weight:500;">' + (value || '—') + '</td>'
    + '</tr>';
}

function sendNotificationEmail(data, fileLinks) {
  const name    = [data.urg_ovog, data.ovog, data.ner].filter(Boolean).join(' ') || '—';
  const subject = '🦉 Woow Pay — Шинэ мерчант: ' + (data.brand || name);

  // File rows
  const fileOrder = ['id_zurag', 'uls_burtgel', 'turees_geree', 'tus_zovshoorul'];
  const fileLabels = {
    id_zurag:       'Иргэний үнэмлэх',
    uls_burtgel:    'Улсын бүртгэл',
    turees_geree:   'Түрээсийн гэрээ',
    tus_zovshoorul: 'Тусгай зөвшөөрөл'
  };
  const fileRows = fileOrder
    .filter(k => fileLinks && fileLinks[k])
    .map(k => '<tr style="border-bottom:1px solid #F0F7FA;">'
      + '<td style="padding:8px 14px;color:#5A7A8A;font-size:13px;width:38%;background:#FAFCFE;">' + fileLabels[k] + '</td>'
      + '<td style="padding:8px 14px;"><a href="' + fileLinks[k] + '" style="color:#29BDE0;font-weight:700;text-decoration:none;">📎 Файл харах →</a></td>'
      + '</tr>')
    .join('');

  const htmlBody = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#eef2f7;font-family:Arial,Helvetica,sans-serif;">'
    + '<div style="max-width:620px;margin:28px auto;background:#fff;border-radius:20px;overflow:hidden;box-shadow:0 6px 24px rgba(0,0,0,.12);">'

    // Header
    + '<div style="background:linear-gradient(135deg,#060C24 0%,#0A1840 100%);padding:32px;text-align:center;">'
    + '<div style="font-size:32px;margin-bottom:10px;">🦉</div>'
    + '<h1 style="color:#fff;margin:0;font-size:22px;font-weight:700;letter-spacing:-.02em;">Woow Pay</h1>'
    + '<p style="color:#A8E9F7;margin:8px 0 0;font-size:14px;">Шинэ мерчант бүртгэлийн хүсэлт</p>'
    + '</div>'

    // Ref number banner
    + '<div style="background:#EBF8FC;border-bottom:2px solid #C4E8F5;padding:14px 32px;display:flex;justify-content:space-between;align-items:center;">'
    + '<span style="font-size:13px;color:#1A5068;font-weight:600;">Лавлах дугаар</span>'
    + '<span style="font-weight:800;color:#29BDE0;font-size:18px;letter-spacing:.1em;">' + (data.ref_number || '—') + '</span>'
    + '</div>'

    + '<div style="padding:28px 32px;">'

    // Personal
    + '<h3 style="margin:0 0 10px;font-size:11px;font-weight:700;color:#29BDE0;text-transform:uppercase;letter-spacing:.08em;">👤 Хувийн мэдээлэл</h3>'
    + '<table style="width:100%;border-collapse:collapse;border:1px solid #E0EFF5;border-radius:12px;overflow:hidden;margin-bottom:22px;">'
    + emailRow('Нэр', name)
    + emailRow('РД', data.rd)
    + emailRow('Утас', data.utas)
    + emailRow('Имэйл', data.email)
    + emailRow('Facebook', data.facebook)
    + emailRow('Хаяг', data.hayag)
    + '</table>'

    // Business
    + '<h3 style="margin:0 0 10px;font-size:11px;font-weight:700;color:#29BDE0;text-transform:uppercase;letter-spacing:.08em;">🏪 Дэлгүүрийн мэдээлэл</h3>'
    + '<table style="width:100%;border-collapse:collapse;border:1px solid #E0EFF5;border-radius:12px;overflow:hidden;margin-bottom:22px;">'
    + emailRow('Бренд', data.brand)
    + emailRow('Худалдааны төрөл', data.trade_type)
    + emailRow('Мерчант хаяг', data.merchant_hayag)
    + emailRow('Цахим хаяг', data.merchant_social)
    + emailRow('Салбарын тоо', data.branch_count)
    + emailRow('Дундаж борлуулалт', data.daily_sales_range)
    + emailRow('Банк', data.bank_name)
    + emailRow('Дансны дугаар', data.dans)
    + '</table>'

    // Other
    + '<h3 style="margin:0 0 10px;font-size:11px;font-weight:700;color:#29BDE0;text-transform:uppercase;letter-spacing:.08em;">📋 Бусад</h3>'
    + '<table style="width:100%;border-collapse:collapse;border:1px solid #E0EFF5;border-radius:12px;overflow:hidden;margin-bottom:22px;">'
    + emailRow('Урилгын код', data.invite_code)
    + emailRow('Хаанаас мэдсэн', data.source_type)
    + emailRow('Тайлбар', data.notes)
    + '</table>'

    // Files
    + (fileRows
        ? '<h3 style="margin:0 0 10px;font-size:11px;font-weight:700;color:#29BDE0;text-transform:uppercase;letter-spacing:.08em;">📎 Хавсаргасан файлууд</h3>'
          + '<table style="width:100%;border-collapse:collapse;border:1px solid #E0EFF5;border-radius:12px;overflow:hidden;margin-bottom:22px;">'
          + fileRows + '</table>'
        : '<p style="color:#aaa;font-size:13px;margin-bottom:22px;">Файл хавсаргаагүй</p>')

    // CTA
    + '<div style="text-align:center;margin-top:8px;">'
    + '<a href="' + CONFIG.SPREADSHEET_URL + '" style="display:inline-block;background:linear-gradient(135deg,#29BDE0,#5CD4EE);color:#fff;padding:14px 32px;border-radius:12px;text-decoration:none;font-weight:700;font-size:15px;">📊 Google Sheets харах</a>'
    + '</div>'
    + '</div>'

    // Footer
    + '<div style="background:#F5FBFE;border-top:1px solid #C4E8F5;padding:16px 32px;text-align:center;">'
    + '<p style="color:#aaa;font-size:12px;margin:0;">© 2026 Woow Pay · Автоматаар илгээгдсэн</p>'
    + '</div>'
    + '</div></body></html>';

  // Plain text fallback
  const plain = [
    'ШИНЭ МЕРЧАНТ БҮРТГЭЛ',
    'Лавлах: ' + (data.ref_number || '—'),
    'Огноо:  ' + new Date().toLocaleString('mn-MN'),
    '',
    '── ХУВИЙН ──',
    'Нэр:   ' + name,
    'РД:    ' + (data.rd || '—'),
    'Утас:  ' + (data.utas || '—'),
    'Имэйл: ' + (data.email || '—'),
    'Хаяг:  ' + (data.hayag || '—'),
    '',
    '── ДЭЛГҮҮР ──',
    'Бренд:        ' + (data.brand || '—'),
    'Худалдааны:   ' + (data.trade_type || '—'),
    'Мерчант хаяг: ' + (data.merchant_hayag || '—'),
    'Банк:         ' + (data.bank_name || '—'),
    'Данс:         ' + (data.dans || '—'),
    '',
    '── ФАЙЛУУД ──',
    ...fileOrder.filter(k => fileLinks && fileLinks[k]).map(k => fileLabels[k] + ': ' + fileLinks[k]),
    '',
    'Sheets: ' + CONFIG.SPREADSHEET_URL
  ].join('\n');

  MailApp.sendEmail({
    to:       CONFIG.NOTIFY_EMAILS.join(', '),
    subject:  subject,
    body:     plain,
    htmlBody: htmlBody
  });
}

// ── APPEND ROW ───────────────────────────────────────────────
function appendToSheet(data, fileLinks) {
  const sheet = getOrCreateSheet();
  sheet.appendRow([
    data.timestamp || new Date().toISOString(),
    data.ref_number || '',
    data.status || 'Шинэ хүсэлт',
    data.urg_ovog || '', data.ovog || '', data.ner || '',
    data.rd || '', data.utas || '', data.email || '', data.facebook || '',
    data.hayag || '', data.brand || '', data.trade_type || '', data.merchant_hayag || '',
    data.branch_count || '', data.daily_sales_range || '',
    data.bank_name || '', data.dans || '',
    data.invite_code || '', data.source_type || '', data.notes || '', data.merchant_social || '',
    (fileLinks && fileLinks.id_zurag)       || '',
    (fileLinks && fileLinks.uls_burtgel)    || '',
    (fileLinks && fileLinks.turees_geree)   || '',
    (fileLinks && fileLinks.tus_zovshoorul) || ''
  ]);

  // Highlight the Status cell yellow for new entries
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 3).setBackground('#FFF3CD');
}

// ── MAIN HANDLER (text data only) ────────────────────────────
function handleSubmission(payload) {
  const files = payload.files || {};
  delete payload.files;

  // Files are now sent separately via add_file action.
  // Handle legacy submissions that still bundle files in the payload.
  const fileLinks = (files && Object.keys(files).length)
    ? saveFilesToDrive(files, payload.ref_number || 'unknown')
    : {};

  appendToSheet(payload, fileLinks);

  try {
    sendNotificationEmail(payload, fileLinks);
  } catch (e) {
    Logger.log('Email error: ' + e.message);
  }
}

// ── FILE UPLOAD HANDLER ───────────────────────────────────────
// Each file is sent as a separate POST with action:'add_file'.
const FILE_COLUMNS = {
  id_zurag:       23,   // Column W — Иргэний үнэмлэх
  uls_burtgel:    24,   // Column X — Улсын бүртгэл
  turees_geree:   25,   // Column Y — Түрээсийн гэрээ
  tus_zovshoorul: 26    // Column Z — Тусгай зөвшөөрөл
};

function handleFileUpload(payload) {
  const { ref_number, field, file } = payload;
  if (!ref_number || !field || !file || !file.data || !file.name) {
    Logger.log('handleFileUpload: missing fields');
    return;
  }
  const col = FILE_COLUMNS[field];
  if (!col) { Logger.log('Unknown field: ' + field); return; }

  const labels = {
    id_zurag:       'Иргэний_үнэмлэх',
    uls_burtgel:    'Улсын_бүртгэл',
    turees_geree:   'Түрээсийн_гэрээ',
    tus_zovshoorul: 'Тусгай_зөвшөөрөл'
  };

  try {
    const folder = getOrCreateFolder(CONFIG.DRIVE_FOLDER + '/' + ref_number);
    const ext    = file.name.split('.').pop() || '';
    const name   = (labels[field] || field) + (ext ? '.' + ext : '');
    const bytes  = Utilities.base64Decode(file.data);
    const blob   = Utilities.newBlob(bytes, file.type || 'application/octet-stream', name);
    const f      = folder.createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = f.getUrl();
    Logger.log('File saved: ' + name + ' → ' + url);

    // Find the sheet row by ref_number and write the Drive link
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(ref_number)) {
        sheet.getRange(i + 1, col).setValue(url);
        Logger.log('Updated row ' + (i + 1) + ', col ' + col);
        return;
      }
    }
    Logger.log('Row not found for ref: ' + ref_number);
  } catch (e) {
    Logger.log('handleFileUpload error: ' + e.message);
  }
}

// ── WEB APP ENTRY POINTS ─────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === 'add_file') {
      handleFileUpload(payload);
    } else {
      handleSubmission(payload);
    }
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.payload) {
      handleSubmission(JSON.parse(decodeURIComponent(e.parameter.payload)));
    }
    return ContentService
      .createTextOutput('ok')
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    Logger.log('doGet error: ' + err.message);
    return ContentService
      .createTextOutput('error: ' + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ══════════════════════════════════════════════════════════════
//  ▶ SETUP — Paste this script, then RUN THIS FUNCTION ONCE.
//    It will:  1) Reset the sheet with clean headers
//              2) Authorize Gmail + Drive
//              3) Send a test email to confirm everything works
// ══════════════════════════════════════════════════════════════
function SETUP_RUN_ONCE() {
  // 1. Reset sheet
  resetSheet();
  Logger.log('✅ Sheet reset with clean headers');

  // 2. Test Drive folder
  try {
    const folder = getOrCreateFolder(CONFIG.DRIVE_FOLDER);
    Logger.log('✅ Drive folder ready: ' + folder.getName());
  } catch (e) {
    Logger.log('❌ Drive error: ' + e.message);
    throw e;
  }

  // 3. Send test email (this also authorizes MailApp)
  try {
    MailApp.sendEmail({
      to:      CONFIG.NOTIFY_EMAILS.join(', '),
      subject: '✅ Woow Pay — Тохиргоо амжилттай боллоо',
      body:    'Gmail, Google Sheets болон Drive файл хадгалах систем идэвхжлээ.\n\nWoow Pay Merchant System',
      htmlBody: '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#eef2f7;font-family:Arial,sans-serif;">'
        + '<div style="max-width:420px;margin:32px auto;background:#060C24;border-radius:20px;padding:36px;text-align:center;">'
        + '<div style="font-size:44px;margin-bottom:14px;">✅</div>'
        + '<h2 style="color:#A8E9F7;margin:0 0 10px;font-size:20px;">Тохиргоо амжилттай!</h2>'
        + '<p style="color:#7BC8DA;margin:0;font-size:14px;line-height:1.6;">Woow Pay мерчант бүртгэлийн систем<br>бүрэн бэлэн боллоо.</p>'
        + '</div></body></html>'
    });
    Logger.log('✅ Test email sent to: ' + CONFIG.NOTIFY_EMAILS.join(', '));
  } catch (e) {
    Logger.log('❌ Email error: ' + e.message);
    throw e;
  }

  Logger.log('🎉 Setup complete! Now deploy as Web App.');
}

// ── TEST — simulate a real submission ────────────────────────
function TEST_SUBMISSION() {
  handleSubmission({
    timestamp: new Date().toISOString(),
    ref_number: 'WP-TEST01',
    status: 'Тест',
    urg_ovog: 'Тест', ovog: 'Овог', ner: 'Нэр',
    rd: 'АА12345678', utas: '99001122',
    email: CONFIG.NOTIFY_EMAILS[0],
    facebook: 'test.merchant', hayag: 'Улаанбаатар, Сүхбаатар дүүрэг',
    brand: 'Тест Дэлгүүр', trade_type: '🛒 Хүнсний дэлгүүр',
    merchant_hayag: 'Санаа плаза, 2 давхар', branch_count: '1',
    daily_sales_range: '300,000₮ – 700,000₮', bank_name: 'Хаан банк',
    dans: '5012345678', invite_code: '116',
    source_type: 'Найз / танил', notes: 'Туршилтын бүртгэл.',
    files: {}
  });
  Logger.log('✅ Test done — check Sheets + email inbox!');
}
