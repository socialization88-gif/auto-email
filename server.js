import express from 'express';
import cors from 'cors';
import nodemailer from 'nodemailer';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
const port = Number(process.env.PORT || 3001);

app.use(cors());
app.use(express.json({ limit: '2mb' }));
app.use(express.static('.'));

function pad(value) {
  return String(value).padStart(2, '0');
}

function formatDateTime(date = new Date()) {
  let hours = date.getHours();
  const minutes = pad(date.getMinutes());
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12 || 12;
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(hours)}:${minutes} ${ampm}`;
}

function escapeHtml(value = '') {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function normalizeLookupKey(value = '') {
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function extractGoogleSheetId(value = '') {
  const text = String(value || '').trim();
  if (!text) return '';

  const urlMatch = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/i);
  if (urlMatch?.[1]) return urlMatch[1];

  const directIdMatch = text.match(/^[a-zA-Z0-9-_]{20,}$/);
  return directIdMatch?.[0] || '';
}

function parseCsv(text = '') {
  const rows = [];
  let currentRow = [];
  let currentCell = '';
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (inQuotes) {
      if (char === '"') {
        if (next === '"') {
          currentCell += '"';
          index += 1;
        } else {
          inQuotes = false;
        }
      } else {
        currentCell += char;
      }
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      continue;
    }

    if (char === ',') {
      currentRow.push(currentCell);
      currentCell = '';
      continue;
    }

    if (char === '\n') {
      currentRow.push(currentCell);
      rows.push(currentRow);
      currentRow = [];
      currentCell = '';
      continue;
    }

    if (char !== '\r') {
      currentCell += char;
    }
  }

  currentRow.push(currentCell);
  rows.push(currentRow);

  return rows.filter((row) => row.some((cell) => String(cell || '').trim() !== ''));
}

function pickCellByHeaders(headerMap, row, candidates) {
  for (const candidate of candidates) {
    const index = headerMap.get(normalizeLookupKey(candidate));
    if (Number.isInteger(index) && index >= 0) {
      return String(row[index] ?? '').trim();
    }
  }
  return '';
}

async function fetchSheetRows(sheetInfo = {}) {
  const sheetId = extractGoogleSheetId(sheetInfo.id || sheetInfo.link || '');
  if (!sheetId) {
    throw new Error('Sheet ID is missing or invalid.');
  }

  const sheetName = String(sheetInfo.name || '').trim();
  const csvUrl = new URL(`https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq`);
  csvUrl.searchParams.set('tqx', 'out:csv');
  if (sheetName) {
    csvUrl.searchParams.set('sheet', sheetName);
  }

  const response = await fetch(csvUrl, {
    headers: {
      'User-Agent': 'Mozilla/5.0'
    }
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch sheet data (${response.status}).`);
  }

  const csvText = await response.text();
  const rows = parseCsv(csvText);
  if (rows.length < 2) {
    throw new Error('Selected sheet returned no data rows.');
  }

  const headers = rows[0];
  const headerMap = new Map(headers.map((header, index) => [normalizeLookupKey(header), index]));

  return rows.slice(1).map((row) => {
    const rawName = pickCellByHeaders(headerMap, row, ['name', 'full name', 'username', 'student name', 'member name']);
    const rawEmail = pickCellByHeaders(headerMap, row, ['email', 'e-mail', 'mail', 'email address']);
    const rawId = pickCellByHeaders(headerMap, row, ['id', 'user id', 'student id', 'member id']);
    const dateRange = pickCellByHeaders(headerMap, row, ['date range', 'daterange', 'date', 'period', 'week', 'range']);
    const taskName = pickCellByHeaders(headerMap, row, ['task name', 'taskname', 'task', 'activity', 'work']);
    const points = pickCellByHeaders(headerMap, row, ['points', 'point', 'score']);
    const remarks = pickCellByHeaders(headerMap, row, ['remarks', 'remark', 'note']);

    return {
      id: rawId,
      name: rawName,
      email: rawEmail,
      dateRange,
      taskName,
      points: Number(String(points || 0).replace(/[^0-9.-]/g, '')) || 0,
      remarks,
      raw: Object.fromEntries(headers.map((header, index) => [header, String(row[index] ?? '').trim()]))
    };
  }).filter((row) => row.name || row.email || row.dateRange || row.taskName || row.points);
}

const defaultUsers = [
  { id: '001', name: 'Zannatul Adan', email: 'adanzannat@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '002', name: 'Chotan Chowdhury', email: 'chowdhury.chotan24@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '003', name: 'Ukyaa Marma', email: 'ukyaamarma.cu@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '004', name: 'Abid Rohman Zenith', email: 'abidrohmanzenith@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '005', name: 'Keosajai Marma', email: 'keosazaimarma@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '006', name: 'Md. Rafikul Islam\nHriday', email: 'hriday1833@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '007', name: 'Otingmang Chak', email: 'utingmong99@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '008', name: 'Joshaimong Marma', email: 'joshaimongmarma149@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '009', name: 'Bendikar Bawm', email: 'bawmbendi@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '010', name: 'Yayoung Murong', email: 'yayoungmro@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '011', name: 'M. Najmol Islam', email: 'najmolislam1557@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '012', name: 'Khiamong Chak', email: 'kmdhrubo24@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '013', name: 'Hlathoaicha Chak', email: 'hlathoaichachak878@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '014', name: 'Prachurja Chakma', email: 'prachurjachakmap@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '015', name: 'Kyachaimong Chak', email: 'kyachaimongchak4@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '016', name: 'Baidaram Tripura', email: 'ramtripura175@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '017', name: 'Orbish Chakma', email: 'orbishchakma835@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '018', name: 'Rakib Hasan', email: 'rakibhasanraka188@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '019', name: 'Subrata Chakma', email: 'subratatuzim178@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '020', name: 'Forhadul Islam', email: 'forhadulpc@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '021', name: 'Shuvo Baidya', email: 'm.shuvobaidya710@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '022', name: 'Rangchang Mro', email: 'chongmro248@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '023', name: 'Reasad Ali', email: 'mdreasadali1107@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '024', name: 'Kaingwy Mro', email: 'kaingwym@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '025', name: 'Kaingpre Mro', email: 'praymrokaing@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '026', name: 'Md. Haidar Ali', email: 'haidarali45466@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '027', name: 'Arif Shahriar', email: 'mdarifshahriarbappy@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '028', name: 'Jebedai Bawm', email: 'bawmjebedai@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '029', name: 'Rumonjoy Tanchangya', email: 'rumonjoy2025@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '030', name: 'Soshamoy\nTongchongya', email: 'jovantcg04@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '031', name: 'M. Mojahid Sheikh', email: 'sk.mujahid8875@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '032', name: 'Koshai Mro', email: '287kusai@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '033', name: 'Uthoaingyo Chak', email: 'uthoaingyochak16@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '034', name: 'Mangrum Mro', email: 'mangrummro6@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '035', name: 'Denwy Murong', email: 'denwymro@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '036', name: 'Thonging Khumi', email: 'thongingkhumi@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '037', name: 'Manok Mro', email: 'mankoh4142@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '038', name: 'Md. Abul Hossain\nSurjo', email: 'mdabulhossainsurjo23@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '039', name: 'Prue Mong U Marma', email: 'promong302@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '040', name: 'Hlathuiching Marma', email: '', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '041', name: 'Tonu Ray Tripura', email: 'tonuraytripura14@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '042', name: 'Waingjow Marma', email: 'wainggyomarma@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '043', name: 'Julinmoy Tripura', email: 'julinmoy@gmail.com', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '044', name: 'Lalneihsang Bawm', email: '', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' },
  { id: '045', name: 'Rubaiya Sultana', email: '', statusType: 'status-pending', statusText: '⏱ Pending', time: 'Not sent' }
];

function createDefaultCcRecipients() {
  return [
    { id: '01', name: 'Zannatul Adan', email: 'adanzannat@gmail.com', checked: false, dailyCount: 0 },
    { id: '02', name: 'Chotan Chowdhury', email: 'chowdhury.chotan24@gmail.com', checked: false, dailyCount: 0 },
    { id: '03', name: 'Ukyaa Marma', email: 'ukyaamarma.cu@gmail.com', checked: false, dailyCount: 0 }
  ];
}

const defaultState = {
  selectedSheet: 'MAY 2026',
  selectedWeeks: [1, 2],
  weekType: '5week',
  users: defaultUsers,
  sheets: [
    { name: 'May 2026', id: '--', link: '#', checked: true },
    { name: 'June 2026', id: '--', link: '#', checked: false },
    { name: 'July 2026', id: '--', link: '#', checked: false },
    { name: 'August 2026', id: '--', link: '#', checked: false },
    { name: 'September 2026', id: '--', link: '#', checked: false },
    { name: 'October 2026', id: '--', link: '#', checked: false },
    { name: 'November 2026', id: '--', link: '#', checked: false },
    { name: 'December 2026', id: '--', link: '#', checked: false }
  ],
  ccRecipients: createDefaultCcRecipients(),
  logs: [
    { time: '2026-06-02 03:55 PM', message: 'System initialized' },
    { time: '2026-06-02 03:56 PM', message: 'Seed data loaded' },
    { time: '2026-06-02 03:57 PM', message: 'Nodemailer server ready' }
  ],
  lastUpdated: formatDateTime()
};

function buildReportHtml(payload) {
  const {
    reportTitle = 'PERFORMANCE EMAIL SYSTEM',
    period = 'Progress Report',
    name = 'Team Member',
    totalLabel = 'Total Points',
    totalPoints = 0,
    reducePoint = 0,
    remarks = 'Keep it up!',
    rows = [],
    note = 'This report was created automatically.'
  } = payload;

  const rowsHtml = rows.length
    ? rows.map((row) => `
        <tr>
          <td style="border:1px solid #333; padding:10px;">${escapeHtml(row.dateRange || '')}</td>
          <td style="border:1px solid #333; padding:10px;">${escapeHtml(row.taskName || '')}</td>
          <td style="border:1px solid #333; padding:10px; text-align:center;">${escapeHtml(row.points ?? 0)}</td>
        </tr>
      `).join('')
    : `
        <tr>
          <td style="border:1px solid #333; padding:10px;" colspan="3">No detailed rows were provided.</td>
        </tr>
      `;

  return `
  <div style="font-family: Arial, sans-serif; max-width: 850px; margin: auto; border: 1px solid #ccc; border-radius: 12px; overflow: hidden;">
    <div style="background-color: #006D6F; color: white; padding: 25px; text-align: center;">
      <h2 style="margin:0;">${escapeHtml(reportTitle)}</h2>
      <p style="margin-top:10px;"><strong>Period:</strong> ${escapeHtml(period)}</p>
    </div>
    <div style="padding: 20px;">
      <p>Assalamu Alaikum <strong>${escapeHtml(name)}</strong>,<br><br></p>
      <p>This report shows your recent activities and the points you earned.</p>
      <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
        <thead>
          <tr style="background-color:#D3D3D3;">
            <th style="border:1px solid #333; padding:10px; width: 25%;">DATE RANGE</th>
            <th style="border:1px solid #333; padding:10px; width: 55%;">TASK NAME</th>
            <th style="border:1px solid #333; padding:10px; width: 20%;">POINTS</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml}
          <tr style="background-color: #ffffff; color: #d32f2f; font-weight: bold;">
            <td style="border: 1px solid #000000; border-right: none;"></td>
            <td style="padding: 10px; border: 1px solid #000000; border-left: none; border-right: none; text-transform: uppercase; letter-spacing: 1px;">
              Self-Development Reduce Point
            </td>
            <td style="padding: 10px; border: 1px solid #000000; text-align: center; font-size: 16px;">
              - ${escapeHtml(reducePoint)}
            </td>
          </tr>
          <tr style="background-color: #004d4e; color: white; font-weight: bold;">
            <td colspan="2" style="border:1px solid #333; padding:12px; text-align: right;">${escapeHtml(totalLabel)}</td>
            <td style="border:1px solid #333; padding:12px; text-align: center;">${escapeHtml(totalPoints)}</td>
          </tr>
        </tbody>
      </table>
      <div style="margin-top: 20px; border-left: 4px solid #006D6F; background:#cfe2e6; padding: 15px;">
        <strong>Remarks:</strong><br>
        <p style="font-style: italic; margin-top: 5px;">"${escapeHtml(remarks)}"</p>
      </div>
      <div style="margin-top: 30px;">
        <p style="margin: 0; font-weight: bold; color: #333;">Best Regards,<br></p>
        <p style="margin: 5px 0 0 0; font-weight: bold; color: #006D6F; font-size: 18px;">Socialization Team</p>
        <p style="margin: 0; color: #666; font-size: 14px;">Quantum Foundation</p>
      </div>
      <div style="margin-top: 25px; padding: 15px; border: 1px dashed #ccc; border-radius: 8px; background-color: #fafafa;">
        <p style="margin: 0; font-size: 13px; color: #555;">
          <span style="color: #d9534f; font-weight: bold; text-transform: uppercase;">SYSTEM NOTE:</span>
          ${escapeHtml(note)}
        </p>
      </div>
    </div>
  </div>`;
}

function createTransporter() {
  const host = process.env.SMTP_HOST;
  const user = process.env.SMTP_USER;
  const pass = process.env.SMTP_PASS;

  if (!host || !user || !pass) {
    throw new Error('Missing SMTP config. Set SMTP_HOST, SMTP_PORT, SMTP_USER, and SMTP_PASS.');
  }

  return nodemailer.createTransport({
    host,
    port: Number(process.env.SMTP_PORT || 587),
    secure: String(process.env.SMTP_SECURE || 'false').toLowerCase() === 'true',
    auth: {
      user,
      pass
    }
  });
}

function classifyEmailFailure(error) {
  const raw = String(error?.message || error || 'Unknown email failure');
  const code = String(error?.code || '').trim();
  const normalized = `${code} ${raw}`.toUpperCase();

  if (/MISSING SMTP CONFIG/.test(normalized)) {
    return {
      reason: 'SMTP configuration is missing.',
      fixSuggestion: 'Set SMTP_HOST, SMTP_PORT, SMTP_USER, and SMTP_PASS in .env, then restart the server.'
    };
  }

  if (/MISSING NAME OR EMAIL/.test(normalized) || /MISSING NAME\/EMAIL/.test(normalized)) {
    return {
      reason: 'A recipient is missing a name or email.',
      fixSuggestion: 'Fill in both name and email for every selected row before sending.'
    };
  }

  if (/ENOTFOUND/.test(normalized)) {
    return {
      reason: 'SMTP host could not be found.',
      fixSuggestion: 'Check SMTP_HOST spelling and make sure the server hostname is reachable.'
    };
  }

  if (/ECONNREFUSED/.test(normalized)) {
    return {
      reason: 'Connection to the SMTP server was refused.',
      fixSuggestion: 'Verify SMTP_HOST and SMTP_PORT, and make sure the mail server allows connections from this machine.'
    };
  }

  if (/ETIMEDOUT/.test(normalized) || /TIMEOUT/.test(normalized)) {
    return {
      reason: 'SMTP connection timed out.',
      fixSuggestion: 'Check internet connectivity, firewall rules, and whether the SMTP server is responding on the selected port.'
    };
  }

  if (/EAUTH/.test(normalized) || /AUTH/.test(normalized) || /535/.test(normalized) || /534/.test(normalized)) {
    return {
      reason: 'SMTP authentication failed.',
      fixSuggestion: 'Use the correct SMTP username/password or Gmail App Password, then restart the server.'
    };
  }

  if (/454/.test(normalized) || /452/.test(normalized) || /DAILY.*LIMIT/.test(normalized) || /QUOTA/.test(normalized) || /RATE LIMIT/.test(normalized) || /TEMPORARILY DISABLED/.test(normalized)) {
    return {
      reason: 'Daily email limit or sending quota has been reached.',
      fixSuggestion: 'Wait for the quota reset, reduce the number of recipients, or use an SMTP account with a higher sending limit.'
    };
  }

  if (/550|551|552|553|554/.test(normalized)) {
    return {
      reason: 'The SMTP server rejected the recipient or message.',
      fixSuggestion: 'Confirm the recipient email is valid and check whether the sender/domain is allowed by the mail server.'
    };
  }

  if (/INVALID|ENVELOPE/.test(normalized)) {
    return {
      reason: 'One of the email addresses or headers is invalid.',
      fixSuggestion: 'Check sender, recipient, and CC email values for typos and empty strings.'
    };
  }

  return {
    reason: raw,
    fixSuggestion: 'Check the server console, SMTP settings, and recipient email values for the exact cause.'
  };
}

app.get('/', (req, res) => {
  res.sendFile('preview (10).html', { root: '.' });
});

app.get('/health', (req, res) => {
  res.json({ ok: true, status: 'ready', time: formatDateTime() });
});

app.get('/summary', (req, res) => {
  const users = defaultState.users;
  const sheets = defaultState.sheets;
  const ccRecipients = defaultState.ccRecipients;
  const logs = defaultState.logs;
  const sentUsers = users.filter((user) => String(user.statusText || '').includes('Sent')).length;
  const pendingUsers = users.length - sentUsers;
  const selectedSheets = sheets.filter((sheet) => sheet.checked).length;
  const selectedCc = ccRecipients.filter((cc) => cc.checked).length;

  const html = `<!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Performance Email Summary</title>
    <style>
      body{margin:0;font-family:Arial,sans-serif;background:#eef4fb;color:#102033}
      .wrap{max-width:1200px;margin:0 auto;padding:24px}
      .hero{background:#fff;border:1px solid #c8d7ea;border-radius:20px;padding:20px;box-shadow:0 14px 34px rgba(47,115,216,.12)}
      h1{margin:0 0 8px;color:#1f5fae}
      .grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-top:18px}
      .card{background:#fff;border:1px solid #dbe6f3;border-radius:18px;padding:16px;box-shadow:0 8px 18px rgba(47,115,216,.08)}
      .label{font-size:12px;color:#5f7188;font-weight:700;text-transform:uppercase;letter-spacing:.3px}
      .value{font-size:28px;font-weight:900;color:#1f5fae;margin-top:8px}
      .section{margin-top:18px;background:#fff;border:1px solid #dbe6f3;border-radius:18px;overflow:hidden}
      .section h2{margin:0;padding:16px 18px;background:#dbeafe;color:#1f5fae;font-size:16px}
      table{width:100%;border-collapse:collapse}
      th,td{padding:12px 14px;border-bottom:1px solid #e4edf8;text-align:left;font-size:13px}
      th{background:#f4f8fd;color:#1f5fae;text-transform:uppercase;letter-spacing:.3px}
      .topbar{display:flex;justify-content:space-between;align-items:center;gap:12px}
      .btn{display:inline-block;background:#2f73d8;color:#fff;text-decoration:none;padding:10px 14px;border-radius:10px;font-weight:900}
      @media(max-width:900px){.grid{grid-template-columns:repeat(2,minmax(0,1fr))}.topbar{flex-direction:column;align-items:flex-start}}
      @media(max-width:600px){.grid{grid-template-columns:1fr}}
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="hero">
        <div class="topbar">
          <div>
            <h1>Data Summary</h1>
            <div>Last updated: ${escapeHtml(defaultState.lastUpdated || formatDateTime())}</div>
          </div>
          <a class="btn" href="/">Back to App</a>
        </div>
        <div class="grid">
          <div class="card"><div class="label">Users</div><div class="value">${users.length}</div></div>
          <div class="card"><div class="label">Sent</div><div class="value">${sentUsers}</div></div>
          <div class="card"><div class="label">Pending</div><div class="value">${pendingUsers}</div></div>
          <div class="card"><div class="label">Selected Sheets</div><div class="value">${selectedSheets}</div></div>
          <div class="card"><div class="label">CC Receivers</div><div class="value">${ccRecipients.length}</div></div>
          <div class="card"><div class="label">Selected CC</div><div class="value">${selectedCc}</div></div>
          <div class="card"><div class="label">Logs</div><div class="value">${logs.length}</div></div>
          <div class="card"><div class="label">Week Mode</div><div class="value">${escapeHtml(appState.weekType || '5week')}</div></div>
        </div>
      </div>

      <div class="section">
        <h2>Recent Logs</h2>
        <table>
          <thead><tr><th>Time</th><th>Message</th></tr></thead>
          <tbody>
            ${logs.length ? logs.map((log) => `<tr><td>${escapeHtml(log.time || '')}</td><td>${escapeHtml(log.message || '')}</td></tr>`).join('') : '<tr><td colspan="2">No logs available.</td></tr>'}
          </tbody>
        </table>
      </div>
    </div>
  </body>
  </html>`;

  res.send(html);
});

app.post('/api/send-progress-emails', async (req, res) => {
  const {
    recipients = [],
    ccEmails = [],
    subjectPrefix = 'Monthly Progress Report',
    reportTitle = 'PERFORMANCE EMAIL SYSTEM',
    period = 'Progress Report',
    sheetName = '',
    sheetId = '',
    sheetLink = '',
    totalLabel = 'Total Points',
    totalPoints = 0,
    reducePoint = 0,
    remarks = 'Keep it up!',
    rows = [],
    note = 'This report was created automatically.',
    fromName = process.env.FROM_NAME || 'Performance Email System',
    fromEmail = process.env.FROM_EMAIL || process.env.SMTP_USER
  } = req.body || {};

  if (!Array.isArray(recipients) || recipients.length === 0) {
    return res.status(400).json({ ok: false, error: 'No recipients were provided.' });
  }

  let transporter;
  try {
    transporter = createTransporter();
  } catch (error) {
    const diagnostic = classifyEmailFailure(error);
    return res.status(500).json({
      ok: false,
      error: error.message,
      ...diagnostic
    });
  }

  let sheetRows = [];
  try {
    sheetRows = await fetchSheetRows({
      name: sheetName || period,
      id: sheetId,
      link: sheetLink
    });
  } catch (error) {
    const rawMessage = String(error?.message || error || 'Unable to fetch sheet data.');
    const normalized = rawMessage.toLowerCase();
    const fixSuggestion = /returned no data rows|no usable|missing/.test(normalized)
      ? 'Make sure the selected Google Sheet has columns for name, date range, task name, and points, then add at least one data row.'
      : 'Check that the sheet link/id is correct, the tab name matches, and the Google Sheet is accessible or published for export.';

    return res.status(400).json({
      ok: false,
      error: rawMessage,
      reason: 'Unable to fetch sheet data.',
      fixSuggestion
    });
  }

  if (!sheetRows.length) {
    return res.status(400).json({
      ok: false,
      error: 'Selected sheet has no usable rows.',
      reason: 'Selected sheet is empty.',
      fixSuggestion: 'Add rows with name, date range, task name, and points before sending email.'
    });
  }

  const results = [];
  const validCc = Array.isArray(ccEmails) ? ccEmails.filter(Boolean).map(String) : [];
  const normalizeName = (value = '') => normalizeLookupKey(value);

  for (const recipient of recipients) {
    const name = String(recipient?.name || '').trim();
    const email = String(recipient?.email || '').trim();

    if (!name || !email) {
      const diagnostic = classifyEmailFailure({ message: 'Missing name or email.' });
      results.push({
        name,
        email,
        success: false,
        error: 'Missing name or email.',
        errorCode: 'INVALID_RECIPIENT',
        ...diagnostic
      });
      continue;
    }

    const matchedRows = sheetRows.filter((row) => normalizeName(row.name) === normalizeName(name));
    const effectiveRows = matchedRows
      .map((row) => ({
        dateRange: row.dateRange || '',
        taskName: row.taskName || '',
        points: Number(row.points || 0)
      }))
      .filter((row) => row.dateRange || row.taskName || row.points !== 0);

    if (!effectiveRows.length) {
      results.push({
        name,
        email,
        success: false,
        error: 'No matching sheet rows found for this user.',
        errorCode: 'NO_SHEET_MATCH',
        reason: 'The selected sheet does not contain report rows for this user.',
        fixSuggestion: 'Add a row in the selected sheet with the exact same name, plus date range, task name, and points.'
      });
      continue;
    }

    const resolvedEmail = email || matchedRows.find((row) => String(row.email || '').trim())?.email || '';
    if (!resolvedEmail) {
      results.push({
        name,
        email,
        success: false,
        error: 'Recipient email is missing.',
        errorCode: 'MISSING_RECIPIENT_EMAIL',
        reason: 'Email address could not be resolved for this user.',
        fixSuggestion: 'Add an email to the user row or include an email column in the selected sheet.'
      });
      continue;
    }

    const perRecipientTotal = effectiveRows.reduce((sum, row) => sum + Number(row.points || 0), 0);
    const recipientRemarks = matchedRows.find((row) => String(row.remarks || '').trim())?.remarks || remarks;
    const sentAt = formatDateTime();
    const htmlBody = buildReportHtml({
      reportTitle,
      period,
      name,
      totalLabel,
      totalPoints: perRecipientTotal || totalPoints,
      reducePoint,
      remarks: recipientRemarks,
      rows: effectiveRows,
      note
    });

    try {
      const info = await transporter.sendMail({
        from: `${fromName} <${fromEmail}>`,
        to: resolvedEmail,
        cc: validCc.length ? validCc.join(',') : undefined,
        subject: `${subjectPrefix} | ${name}`,
        html: htmlBody
      });

      results.push({
        name,
        email: resolvedEmail,
        success: true,
        sentAt,
        messageId: info.messageId
      });
    } catch (error) {
      const diagnostic = classifyEmailFailure(error);
      results.push({
        name,
        email: resolvedEmail,
        success: false,
        error: error.message,
        errorCode: error.code || '',
        ...diagnostic
      });
    }
  }

  const sentCount = results.filter((item) => item.success).length;
  res.json({
    ok: true,
    sentCount,
    failedCount: results.length - sentCount,
    sentAt: formatDateTime(),
    results
  });
});

app.listen(port, () => {
  console.log(`Nodemailer server running on http://localhost:${port}`);
});
