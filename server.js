import fs from 'fs/promises';
import path from 'path';
import express from 'express';
import cors from 'cors';
import nodemailer from 'nodemailer';
import dotenv from 'dotenv';
import { google } from 'googleapis';

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
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function normalizeLookupKey(value = '') {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function cleanText(value = '') {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function quoteSheetName(value = '') {
  return `'${String(value ?? '').replace(/'/g, "''")}'`;
}

function toNumber(value = '') {
  const num = Number(String(value ?? '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(num) ? num : 0;
}

function getTodayKey(date = new Date()) {
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function extractGoogleSheetId(value = '') {
  const text = String(value || '').trim();
  if (!text) return '';

  const urlMatch = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/i);
  if (urlMatch?.[1]) return urlMatch[1];

  const directIdMatch = text.match(/^[a-zA-Z0-9-_]{20,}$/);
  return directIdMatch?.[0] || '';
}

async function readGoogleCredentialsFromFile(filePath) {
  const raw = await fs.readFile(filePath, 'utf8');
  return JSON.parse(raw);
}

async function loadGoogleSheetsCredentials() {
  const inlineJson = String(process.env.GOOGLE_SERVICE_ACCOUNT_JSON || '').trim();
  if (inlineJson) {
    return JSON.parse(inlineJson);
  }

  const candidatePaths = [
    process.env.GOOGLE_APPLICATION_CREDENTIALS,
    process.env.GOOGLE_SERVICE_ACCOUNT_KEYFILE,
    path.join(process.cwd(), 'storied-courier-497606-n6-212c3f734a4e.json')
  ].filter(Boolean);

  for (const candidatePath of candidatePaths) {
    try {
      return await readGoogleCredentialsFromFile(candidatePath);
    } catch (error) {
      // Try the next credential source.
    }
  }

  throw new Error(
    'Google Sheets credentials are missing. Set GOOGLE_APPLICATION_CREDENTIALS or GOOGLE_SERVICE_ACCOUNT_JSON.'
  );
}

let googleSheetsClientPromise = null;

async function getGoogleSheetsClient() {
  if (!googleSheetsClientPromise) {
    googleSheetsClientPromise = (async () => {
      const credentials = await loadGoogleSheetsCredentials();

      if (!credentials?.client_email || !credentials?.private_key) {
        throw new Error(
          'Google Sheets credentials are invalid. The service account JSON must include client_email and private_key.'
        );
      }

      const auth = new google.auth.JWT({
        email: credentials.client_email,
        key: credentials.private_key,
        scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
      });

      await auth.authorize();
      return google.sheets({ version: 'v4', auth });
    })().catch((error) => {
      googleSheetsClientPromise = null;
      throw error;
    });
  }

  return googleSheetsClientPromise;
}

function normalizeSheetTitle(value = '') {
  return normalizeLookupKey(value);
}

async function fetchSheetRowsViaGoogleApi(sheetInfo = {}) {
  const sheetId = extractGoogleSheetId(sheetInfo.id || sheetInfo.link || '');

  if (!sheetId) {
    throw new Error('Sheet ID is missing or invalid.');
  }

  const requestedSheetName = String(sheetInfo.name || '').trim();
  const sheets = await getGoogleSheetsClient();
  const metadata = await sheets.spreadsheets.get({
    spreadsheetId: sheetId,
    includeGridData: false
  });

  const availableTabs = Array.isArray(metadata.data.sheets)
    ? metadata.data.sheets
        .map((sheet) => sheet?.properties?.title)
        .filter(Boolean)
    : [];

  if (!availableTabs.length) {
    throw new Error('The spreadsheet does not contain any readable tabs.');
  }

  const resolvedSheetName = requestedSheetName
    ? availableTabs.find((title) => title === requestedSheetName) ||
      availableTabs.find((title) => normalizeSheetTitle(title) === normalizeSheetTitle(requestedSheetName)) ||
      availableTabs.find((title) => normalizeSheetTitle(title).includes(normalizeSheetTitle(requestedSheetName))) ||
      availableTabs[0]
    : availableTabs[0];

  const range = quoteSheetName(resolvedSheetName);
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: sheetId,
    range
  });

  const rows = Array.isArray(response.data.values) ? response.data.values : [];

  if (rows.length < 2) {
    throw new Error(
      `Selected sheet tab "${resolvedSheetName}" returned no data rows.`
    );
  }

  return rows;
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

  return rows.filter((row) =>
    row.some((cell) => String(cell || '').trim() !== '')
  );
}

function looksLikeSalarySheet(rows) {
  return rows.some((row) =>
    row.some((cell) =>
      String(cell || '').toLowerCase().includes('salary points sheet')
    )
  );
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

function looksLikeTaskText(value = '') {
  const text = cleanText(value).toLowerCase();

  if (!text) return false;

  return (
    text.includes('report prepared') ||
    text.includes('running') ||
    text.includes('automation') ||
    text.includes('update') ||
    text.includes('messenger') ||
    text.includes('email reporting') ||
    text.includes('task') ||
    /^\d+\./.test(text)
  );
}

function looksLikeDateRange(value = '') {
  const text = cleanText(value);

  if (!text) return false;

  if (looksLikeTaskText(text)) return false;

  const lower = text.toLowerCase();

  if (
    lower.includes('salary points sheet') ||
    lower.includes('total points') ||
    lower.includes('subtotal') ||
    lower.includes('self-development') ||
    lower.includes('lwp') ||
    lower.includes('absence')
  ) {
    return false;
  }

  const monthNames =
    '(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)';

  const patterns = [
    new RegExp(`\\b\\d{1,2}\\s*-\\s*\\d{1,2}\\s+${monthNames}\\b`, 'i'),
    new RegExp(`\\b\\d{1,2}\\s+${monthNames}\\s*-\\s*\\d{1,2}\\s+${monthNames}\\b`, 'i'),
    new RegExp(`\\b\\d{1,2}\\s+${monthNames}\\s*-\\s*\\d{1,2}\\s+${monthNames},?\\s*\\d{4}\\b`, 'i'),
    new RegExp(`\\b${monthNames}\\s+\\d{1,2}\\s*-\\s*${monthNames}\\s+\\d{1,2}\\b`, 'i')
  ];

  return patterns.some((pattern) => pattern.test(text));
}

function isPossibleUserStartRow(row = []) {
  const id = cleanText(row[0]);
  const name = cleanText(row[1]);

  if (!name) return false;

  const lowerName = name.toLowerCase();

  if (
    lowerName.includes('name') ||
    lowerName.includes('salary') ||
    lowerName.includes('subtotal') ||
    lowerName.includes('total') ||
    lowerName.includes('date') ||
    lowerName.includes('task') ||
    lowerName.includes('point')
  ) {
    return false;
  }

  // In your sheet a user start row normally has either serial/id in A or a real name in B.
  return Boolean(id || name);
}

function rowMatchesRecipient(row = [], recipient = {}) {
  const rowId = normalizeLookupKey(row[0] || '');
  const rowName = normalizeLookupKey(row[1] || '');

  const recipientId = normalizeLookupKey(
    recipient?.id || recipient?.employeeId || ''
  );

  const recipientName = normalizeLookupKey(recipient?.name || '');

  if (recipientId && rowId && recipientId === rowId) return true;
  if (recipientName && rowName && recipientName === rowName) return true;

  return false;
}

function findUserStartIndex(rows = [], recipient = {}) {
  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i] || [];

    if (!isPossibleUserStartRow(row)) continue;

    if (rowMatchesRecipient(row, recipient)) {
      return i;
    }
  }

  return -1;
}

function extractUserBlock(rows = [], startIndex = -1) {
  if (startIndex < 0) return [];

  const block = [];

  for (let i = startIndex; i < rows.length; i += 1) {
    const row = rows[i] || [];

    if (i > startIndex && isPossibleUserStartRow(row)) {
      break;
    }

    block.push(row);
  }

  return block;
}

function findDateRangeInColumn(rows = [], colIndex, fallback = '') {
  for (const row of rows) {
    const value = cleanText(row?.[colIndex]);

    if (looksLikeDateRange(value)) {
      return value;
    }
  }

  return fallback;
}

function getWeekDateRange({
  fullRows,
  userBlock,
  colIndex,
  fallback
}) {
  const fromUserBlock = findDateRangeInColumn(userBlock, colIndex, '');

  if (fromUserBlock) return fromUserBlock;

  const fromFullSheet = findDateRangeInColumn(fullRows, colIndex, '');

  if (fromFullSheet) return fromFullSheet;

  return fallback;
}

function isActualTaskEntry(value = '') {
  const text = cleanText(value);
  if (!text) return false;

  const lower = text.toLowerCase();

  if (
    lower.includes('subtotal') ||
    lower.includes('total points') ||
    lower.includes('self-development reduce point') ||
    lower.includes('self-development remarks') ||
    lower.includes('remarks') ||
    lower.includes('salary points sheet') ||
    lower.includes('week 5') ||
    lower.includes('lwp')
  ) {
    return false;
  }

  return looksLikeTaskText(text) || /^\d+\./.test(text);
}

function detectSalarySheetWeekCount(rows = []) {
  const week5HasTaskData = rows.some((row) => isActualTaskEntry(row?.[10]));

  return week5HasTaskData ? 5 : 4;
}

function findReducePointFromRows(rows = []) {
  for (const row of rows) {
    const labelIndex = row.findIndex((cell) =>
      /self-development reduce point/i.test(cleanText(cell))
    );

    if (labelIndex < 0) continue;

    const tail = row.slice(labelIndex + 1);
    for (let index = tail.length - 1; index >= 0; index -= 1) {
      if (cleanText(tail[index]) !== '') {
        return toNumber(tail[index]);
      }
    }

    return 0;
  }

  return 0;
}

/**
 * Specific-user parser for your Salary Points Sheet.
 *
 * Important:
 * - It does NOT parse all users together.
 * - It first isolates target user block by id/name.
 * - Then it extracts tasks only from that user block.
 * - Date range is found by scanning same week column:
 *   first in user block, then in full sheet.
 */
function parseSalarySheetForRecipient(rows, recipient, maxWeekCount = 5) {
  let period = 'Progress Report';

  for (const row of rows.slice(0, 10)) {
    for (const cell of row) {
      const text = cleanText(cell);

      if (/salary points sheet/i.test(text)) {
        period = text.replace(/salary points sheet/gi, '').trim() || period;
      }
    }
  }

  const userStartIndex = findUserStartIndex(rows, recipient);

  if (userStartIndex < 0) {
    return null;
  }

  const userStartRow = rows[userStartIndex] || [];
  const userBlock = extractUserBlock(rows, userStartIndex);

  const id = cleanText(userStartRow[0]);
  const name = cleanText(userStartRow[1]);

  const weekPairs = [
    { week: 1, taskCol: 2, pointCol: 3 },
    { week: 2, taskCol: 4, pointCol: 5 },
    { week: 3, taskCol: 6, pointCol: 7 },
    { week: 4, taskCol: 8, pointCol: 9 },
    { week: 5, taskCol: 10, pointCol: 11 }
  ];

  const blocks = [];

  const remarks =
    cleanText(userStartRow[10]) ||
    userBlock.map((row) => cleanText(row[10])).find(Boolean) ||
    '';

  const specialDate =
    cleanText(userStartRow[11]) ||
    userBlock.map((row) => cleanText(row[11])).find(Boolean) ||
    '';

  const rowTotal =
    toNumber(userStartRow[12]) ||
    userBlock.map((row) => toNumber(row[12])).find((value) => value > 0) ||
    0;

  const activeWeekPairs = weekPairs.slice(0, Math.max(1, Math.min(5, Number(maxWeekCount) || 5)));

  for (const pair of activeWeekPairs) {
    const dateRange = getWeekDateRange({
      fullRows: rows,
      userBlock,
      colIndex: pair.taskCol,
      fallback: `Week ${pair.week}`
    });

    const tasks = [];
    const seen = new Set();

    for (const row of userBlock) {
      const taskName = cleanText(row[pair.taskCol]);
      const points = toNumber(row[pair.pointCol]);

      if (!taskName) continue;

      const lowerTask = taskName.toLowerCase();

      if (
        lowerTask.includes('total points') ||
        lowerTask.includes('self-development reduce') ||
        lowerTask.includes('subtotal') ||
        lowerTask.includes('salary points sheet')
      ) {
        continue;
      }

      if (looksLikeDateRange(taskName)) {
        continue;
      }

      const uniqueKey = normalizeLookupKey(`${taskName}-${points}`);

      if (seen.has(uniqueKey)) continue;

      seen.add(uniqueKey);

      tasks.push({
        taskName,
        points
      });
    }

    if (!tasks.length) continue;

    const subtotal = tasks.reduce(
      (sum, task) => sum + Number(task.points || 0),
      0
    );

    blocks.push({
      week: pair.week,
      dateRange,
      tasks,
      subtotal
    });
  }

  if (!blocks.length) {
    return null;
  }

  return {
    id,
    employeeId: id,
    name,
    email: '',
    period,
    blocks,
    remarks,
    specialDate,
    reducePoint: findReducePointFromRows(userBlock),
    totalPoints:
      rowTotal ||
      blocks.reduce((sum, block) => sum + Number(block.subtotal || 0), 0),
    raw: {}
  };
}

async function fetchSheetRawRows(sheetInfo = {}) {
  const sheetId = extractGoogleSheetId(sheetInfo.id || sheetInfo.link || '');

  if (!sheetId) {
    throw new Error('Sheet ID is missing or invalid.');
  }

  const sheetName = String(sheetInfo.name || '').trim();

  const fetchErrors = [];

  try {
    const csvUrl = new URL(
      `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq`
    );

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

    return rows;
  } catch (error) {
    fetchErrors.push(error);
  }

  try {
    return await fetchSheetRowsViaGoogleApi(sheetInfo);
  } catch (error) {
    fetchErrors.push(error);
  }

  const combinedMessage = fetchErrors
    .map((error) => String(error?.message || error || '').trim())
    .filter(Boolean)
    .join(' | ');

  throw new Error(
    combinedMessage ||
      'Unable to fetch sheet data. Check that the sheet link/id is correct, the tab name matches, and the Google Sheet is accessible or published for export.'
  );
}

function parseFlatRowsForRecipient(rows, recipient) {
  const headers = rows[0];
  const headerMap = new Map(
    headers.map((header, index) => [normalizeLookupKey(header), index])
  );

  const recipientId = normalizeLookupKey(
    recipient?.id || recipient?.employeeId || ''
  );

  const recipientName = normalizeLookupKey(recipient?.name || '');
  const recipientEmail = normalizeLookupKey(recipient?.email || '');

  const matchedRows = rows.slice(1).map((row) => {
    const rawName = pickCellByHeaders(headerMap, row, [
      'name',
      'full name',
      'username',
      'student name',
      'member name'
    ]);

    const rawEmail = pickCellByHeaders(headerMap, row, [
      'email',
      'e-mail',
      'mail',
      'email address'
    ]);

    const rawId = pickCellByHeaders(headerMap, row, [
      'id',
      'user id',
      'student id',
      'member id',
      'employee id'
    ]);

    const dateRange = pickCellByHeaders(headerMap, row, [
      'date range',
      'daterange',
      'date',
      'period',
      'week',
      'range'
    ]);

    const taskName = pickCellByHeaders(headerMap, row, [
      'task name',
      'taskname',
      'task',
      'activity',
      'work'
    ]);

    const points = pickCellByHeaders(headerMap, row, [
      'points',
      'point',
      'score'
    ]);

    const remarks = pickCellByHeaders(headerMap, row, [
      'remarks',
      'remark',
      'note'
    ]);

    return {
      id: rawId,
      employeeId: rawId,
      name: rawName,
      email: rawEmail,
      dateRange,
      taskName,
      points: toNumber(points),
      remarks
    };
  }).filter((row) => {
    const rowId = normalizeLookupKey(row.id || row.employeeId || '');
    const rowName = normalizeLookupKey(row.name || '');
    const rowEmail = normalizeLookupKey(row.email || '');

    if (recipientId && rowId && recipientId === rowId) return true;
    if (recipientEmail && rowEmail && recipientEmail === rowEmail) return true;
    if (recipientName && rowName && recipientName === rowName) return true;

    return false;
  });

  if (!matchedRows.length) return null;

  const blocks = matchedRows
    .filter((row) => row.taskName || row.points)
    .map((row, index) => ({
      week: index + 1,
      dateRange: row.dateRange || `Row ${index + 1}`,
      tasks: [
        {
          taskName: row.taskName,
          points: row.points
        }
      ],
      subtotal: row.points
    }));

  if (!blocks.length) return null;

  return {
    id: matchedRows[0].id || recipient.id || '',
    employeeId: matchedRows[0].employeeId || recipient.employeeId || recipient.id || '',
    name: matchedRows[0].name || recipient.name || '',
    email: matchedRows[0].email || recipient.email || '',
    period: matchedRows[0].dateRange || 'Progress Report',
    blocks,
    remarks: matchedRows.find((row) => row.remarks)?.remarks || '',
    specialDate: '',
    totalPoints: blocks.reduce((sum, block) => sum + Number(block.subtotal || 0), 0),
    reducePoint: 0,
    raw: {}
  };
}

function parseReportForRecipient(rows, recipient, options = {}) {
  if (looksLikeSalarySheet(rows)) {
    return parseSalarySheetForRecipient(rows, recipient, options.maxWeekCount || 5);
  }

  return parseFlatRowsForRecipient(rows, recipient);
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
  ccLastResetDate: getTodayKey(),
  users: defaultUsers,
  sheets: [
    { name: 'May 2026', id: '--', link: '#', checked: true, weekType: '5week' },
    { name: 'June 2026', id: '--', link: '#', checked: false, weekType: '5week' },
    { name: 'July 2026', id: '--', link: '#', checked: false, weekType: '5week' },
    { name: 'August 2026', id: '--', link: '#', checked: false, weekType: '5week' },
    { name: 'September 2026', id: '--', link: '#', checked: false, weekType: '5week' },
    { name: 'October 2026', id: '--', link: '#', checked: false, weekType: '5week' },
    { name: 'November 2026', id: '--', link: '#', checked: false, weekType: '5week' },
    { name: 'December 2026', id: '--', link: '#', checked: false, weekType: '5week' }
  ],
  ccRecipients: createDefaultCcRecipients(),
  logs: [
    { time: '2026-06-02 03:55 PM', message: 'System initialized' },
    { time: '2026-06-02 03:56 PM', message: 'Seed data loaded' },
    { time: '2026-06-02 03:57 PM', message: 'Nodemailer server ready' }
  ],
  lastUpdated: formatDateTime()
};

const dataDir = path.join(process.cwd(), 'data');
const stateFilePath = path.join(dataDir, 'app-state.json');
let currentState = cloneState(defaultState);

function cloneState(value) {
  return JSON.parse(JSON.stringify(value));
}

function normalizeState(value = {}) {
  const incoming = value && typeof value === 'object' ? value : {};
  const users = Array.isArray(incoming.users) ? incoming.users : defaultState.users;
  const sheets = Array.isArray(incoming.sheets) ? incoming.sheets : defaultState.sheets;
  const ccRecipients = Array.isArray(incoming.ccRecipients) ? incoming.ccRecipients : defaultState.ccRecipients;
  const logs = Array.isArray(incoming.logs) ? incoming.logs : defaultState.logs;

  return {
    selectedSheet: String(incoming.selectedSheet || defaultState.selectedSheet),
    selectedWeeks: Array.isArray(incoming.selectedWeeks)
      ? incoming.selectedWeeks.map(Number).filter((week) => Number.isInteger(week) && week >= 1 && week <= 5)
      : cloneState(defaultState.selectedWeeks),
    weekType: String(incoming.weekType || defaultState.weekType) === '4week' ? '4week' : '5week',
    ccLastResetDate: String(incoming.ccLastResetDate || defaultState.ccLastResetDate || getTodayKey()),
    users: cloneState(users),
    sheets: cloneState(sheets.map((sheet) => ({
      ...sheet,
      weekType: String(sheet?.weekType || '5week') === '4week' ? '4week' : '5week'
    }))),
    ccRecipients: cloneState(ccRecipients),
    logs: cloneState(logs),
    lastUpdated: String(incoming.lastUpdated || defaultState.lastUpdated || formatDateTime())
  };
}

async function ensureStateFile() {
  await fs.mkdir(dataDir, { recursive: true });

  try {
    await fs.access(stateFilePath);
  } catch {
    await fs.writeFile(stateFilePath, `${JSON.stringify(defaultState, null, 2)}\n`, 'utf8');
  }
}

async function loadStateFromDisk() {
  await ensureStateFile();

  try {
    const raw = await fs.readFile(stateFilePath, 'utf8');
    return normalizeState(raw ? JSON.parse(raw) : defaultState);
  } catch {
    return normalizeState(defaultState);
  }
}

async function saveStateToDisk(nextState) {
  currentState = normalizeState(nextState);
  await fs.mkdir(dataDir, { recursive: true });
  await fs.writeFile(stateFilePath, `${JSON.stringify(currentState, null, 2)}\n`, 'utf8');
  return currentState;
}

currentState = await loadStateFromDisk();

function buildGroupedTableHtml(blocks = [], totalPoints = 0, totalLabel = 'Total', reducePoint = 0) {
  const bodyRows = [];

  for (const block of blocks) {
    const tasks = Array.isArray(block.tasks) ? block.tasks : [];

    if (!tasks.length) continue;

    const rowSpan = tasks.length + 1;

    tasks.forEach((task, index) => {
      bodyRows.push(`
        <tr>
          ${
            index === 0
              ? `<td rowspan="${rowSpan}" style="padding:10px; border:1px solid #333; text-align:center; font-weight:bold; vertical-align:middle; width:18%; background:#fcfcfc;">
                  ${escapeHtml(block.dateRange)}
                </td>`
              : ''
          }
          <td style="padding:10px; border:1px solid #333; width:67%;">
            ${escapeHtml(task.taskName)}
          </td>
          <td style="padding:10px; border:1px solid #333; text-align:center; width:15%;">
            ${escapeHtml(task.points)}
          </td>
        </tr>
      `);
    });

    bodyRows.push(`
      <tr style="background-color:#F5F5DC; color:#000; font-weight:bold;">
        <td style="padding:8px; border:1px solid #333; text-align:left;">
          Subtotal
        </td>
        <td style="padding:8px; border:1px solid #333; text-align:center;">
          ${escapeHtml(block.subtotal)}
        </td>
      </tr>
    `);
  }

  if (!bodyRows.length) {
    bodyRows.push(`
      <tr>
        <td colspan="3" style="padding:12px; border:1px solid #333; text-align:center;">
          No task data found.
        </td>
      </tr>
    `);
  }

  return `
    <table style="width:100%; border-collapse:collapse; margin-top:15px;">
      <thead>
        <tr style="background-color:#f1f1f1;">
          <th style="border:1px solid #333; padding:10px; width:18%;">DATE</th>
          <th style="border:1px solid #333; padding:10px; width:67%; text-align:left;">TASKS</th>
          <th style="border:1px solid #333; padding:10px; width:15%;">POINTS</th>
        </tr>
      </thead>
      <tbody>
        ${bodyRows.join('')}
        <tr style="background-color:#ffffff; color:#d32f2f; font-weight:bold;">
          <td colspan="2" style="border:1px solid #333; padding:10px; text-align:right;">
            Self-Development Reduce Point
          </td>
          <td style="border:1px solid #333; padding:10px; text-align:center;">
            - ${escapeHtml(reducePoint)}
          </td>
        </tr>
        <tr style="background-color:#004d4e; color:white; font-weight:bold;">
          <td colspan="2" style="border:1px solid #333; padding:12px; text-align:right;">
            ${escapeHtml(totalLabel)}
          </td>
          <td style="border:1px solid #333; padding:12px; text-align:center;">
            ${escapeHtml(totalPoints)}
          </td>
        </tr>
      </tbody>
    </table>
  `;
}

function buildReportHtml(payload) {
  const {
    reportTitle = 'MONTHLY PERFORMANCE SUMMARY',
    period = 'Progress Report',
    name = 'Team Member',
    totalLabel = 'Total',
    totalPoints = 0,
    reducePoint = 0,
    remarks = 'Excellent effort! Keep it up.',
    blocks = [],
    note = 'This report was created automatically. If you have any questions, please contact the administrator.'
  } = payload;

  const tableHtml = buildGroupedTableHtml(blocks, totalPoints, totalLabel, reducePoint);

  return `
  <div style="font-family: Arial, sans-serif; max-width:850px; margin:auto; border:1px solid #ccc; border-radius:8px; overflow:hidden;">
    <div style="background-color:#006D6F; color:white; padding:25px; text-align:center;">
      <h2 style="margin:0; letter-spacing:.5px;">${escapeHtml(reportTitle)}</h2>
      <p style="margin-top:10px;"><strong>Period:</strong> ${escapeHtml(period)}</p>
    </div>

    <div style="padding:20px;">
      <p>Assalamu Alaikum <strong>${escapeHtml(name)}</strong>,<br><br></p>

      <p>
        This report shows your recent activities and the points you earned.
        Each entry helps track your <b>Performance</b> toward your goals.
      </p>

      ${tableHtml}

      <div style="margin-top:20px; border-left:4px solid #006D6F; background:#cfe2e6; padding:15px;">
        <strong>Self-Development Remarks:</strong><br>
        <p style="font-style:italic; margin-top:5px;">"${escapeHtml(remarks)}"</p>
      </div>

      <div style="margin-top:35px;">
        <p style="margin:0; font-weight:bold; color:#333;">Best Regards,<br></p>
        <p style="margin:8px 0 0 0; font-weight:bold; color:#006D6F; font-size:18px;">
          Socialization Team
        </p>
        <p style="margin:0; color:#666; font-size:14px;">Quantum Foundation</p>
      </div>

      <div style="margin-top:25px; padding:15px; border:1px dashed #ccc; border-radius:8px; background-color:#fafafa;">
        <p style="margin:0; font-size:13px; color:#555;">
          <span style="color:#d9534f; font-weight:bold; text-transform:uppercase;">SYSTEM NOTE:</span>
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
    throw new Error(
      'Missing SMTP config. Set SMTP_HOST, SMTP_PORT, SMTP_USER, and SMTP_PASS.'
    );
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
      fixSuggestion:
        'Set SMTP_HOST, SMTP_PORT, SMTP_USER, and SMTP_PASS in .env, then restart the server.'
    };
  }

  if (
    /MISSING NAME OR EMAIL/.test(normalized) ||
    /MISSING NAME\/EMAIL/.test(normalized)
  ) {
    return {
      reason: 'A recipient is missing a name or email.',
      fixSuggestion:
        'Fill in both name and email for every selected row before sending.'
    };
  }

  if (/ENOTFOUND/.test(normalized)) {
    return {
      reason: 'SMTP host could not be found.',
      fixSuggestion:
        'Check SMTP_HOST spelling and make sure the server hostname is reachable.'
    };
  }

  if (/ECONNREFUSED/.test(normalized)) {
    return {
      reason: 'Connection to the SMTP server was refused.',
      fixSuggestion:
        'Verify SMTP_HOST and SMTP_PORT, and make sure the mail server allows connections from this machine.'
    };
  }

  if (/ETIMEDOUT/.test(normalized) || /TIMEOUT/.test(normalized)) {
    return {
      reason: 'SMTP connection timed out.',
      fixSuggestion:
        'Check internet connectivity, firewall rules, and whether the SMTP server is responding on the selected port.'
    };
  }

  if (
    /EAUTH/.test(normalized) ||
    /AUTH/.test(normalized) ||
    /535/.test(normalized) ||
    /534/.test(normalized)
  ) {
    return {
      reason: 'SMTP authentication failed.',
      fixSuggestion:
        'Use the correct SMTP username/password or Gmail App Password, then restart the server.'
    };
  }

  if (
    /454/.test(normalized) ||
    /452/.test(normalized) ||
    /DAILY.*LIMIT/.test(normalized) ||
    /QUOTA/.test(normalized) ||
    /RATE LIMIT/.test(normalized) ||
    /TEMPORARILY DISABLED/.test(normalized)
  ) {
    return {
      reason: 'Daily email limit or sending quota has been reached.',
      fixSuggestion:
        'Wait for the quota reset, reduce the number of recipients, or use an SMTP account with a higher sending limit.'
    };
  }

  if (/550|551|552|553|554/.test(normalized)) {
    return {
      reason: 'The SMTP server rejected the recipient or message.',
      fixSuggestion:
        'Confirm the recipient email is valid and check whether the sender/domain is allowed by the mail server.'
    };
  }

  if (/INVALID|ENVELOPE/.test(normalized)) {
    return {
      reason: 'One of the email addresses or headers is invalid.',
      fixSuggestion:
        'Check sender, recipient, and CC email values for typos and empty strings.'
    };
  }

  return {
    reason: raw,
    fixSuggestion:
      'Check the server console, SMTP settings, and recipient email values for the exact cause.'
  };
}

app.get('/', (req, res) => {
  res.sendFile('preview (10).html', { root: '.' });
});

app.get('/health', (req, res) => {
  res.json({
    ok: true,
    status: 'ready',
    time: formatDateTime()
  });
});

app.get('/summary', (req, res) => {
  const users = currentState.users || defaultState.users;
  const sheets = currentState.sheets || defaultState.sheets;
  const ccRecipients = currentState.ccRecipients || defaultState.ccRecipients;
  const logs = currentState.logs || defaultState.logs;

  const sentUsers = users.filter((user) =>
    String(user.statusText || '').includes('Sent')
  ).length;

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
            <div>Last updated: ${escapeHtml(
              currentState.lastUpdated || formatDateTime()
            )}</div>
          </div>
          <a class="btn" href="/">Back to App</a>
        </div>
        <div class="grid">
          <div class="card"><div class="label">Users</div><div class="value">${
            users.length
          }</div></div>
          <div class="card"><div class="label">Sent</div><div class="value">${sentUsers}</div></div>
          <div class="card"><div class="label">Pending</div><div class="value">${pendingUsers}</div></div>
          <div class="card"><div class="label">Selected Sheets</div><div class="value">${selectedSheets}</div></div>
          <div class="card"><div class="label">CC Receivers</div><div class="value">${
            ccRecipients.length
          }</div></div>
          <div class="card"><div class="label">Selected CC</div><div class="value">${selectedCc}</div></div>
          <div class="card"><div class="label">Logs</div><div class="value">${
            logs.length
          }</div></div>
          <div class="card"><div class="label">Week Mode</div><div class="value">${escapeHtml(
            currentState.weekType || '5week'
          )}</div></div>
        </div>
      </div>

      <div class="section">
        <h2>Recent Logs</h2>
        <table>
          <thead><tr><th>Time</th><th>Message</th></tr></thead>
          <tbody>
            ${
              logs.length
                ? logs
                    .map(
                      (log) =>
                        `<tr><td>${escapeHtml(
                          log.time || ''
                        )}</td><td>${escapeHtml(log.message || '')}</td></tr>`
                    )
                    .join('')
                : '<tr><td colspan="2">No logs available.</td></tr>'
            }
          </tbody>
        </table>
      </div>
    </div>
  </body>
  </html>`;

  res.send(html);
});

app.get('/api/state', (req, res) => {
  res.json({
    ok: true,
    state: currentState
  });
});

app.put('/api/state', async (req, res) => {
  const incoming = req.body && typeof req.body === 'object' ? req.body : {};

  try {
    const saved = await saveStateToDisk({
      ...currentState,
      ...incoming,
      lastUpdated: formatDateTime()
    });

    res.json({
      ok: true,
      state: saved
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message || 'Failed to save application state.'
    });
  }
});

app.post('/api/send-progress-emails', async (req, res) => {
  const {
    recipients = [],
    ccEmails = [],
    selectedWeeks = [],
    subjectPrefix = 'Monthly Progress Report',
    reportTitle = 'MONTHLY PERFORMANCE SUMMARY',
    period = 'Progress Report',
    sheetName = '',
    sheetId = '',
    sheetLink = '',
    sheetWeekType = '5week',
    weekType = '5week',
    totalLabel = 'Total',
    totalPoints = 0,
    remarks = 'Excellent effort! Keep it up.',
    note = 'This report was created automatically. If you have any questions, please contact the administrator.',
    fromName = process.env.FROM_NAME || 'Performance Email System',
    fromEmail = process.env.FROM_EMAIL || process.env.SMTP_USER
  } = req.body || {};

  if (!Array.isArray(recipients) || recipients.length === 0) {
    return res.status(400).json({
      ok: false,
      error: 'No recipients were provided.'
    });
  }

  const requestedWeeks = Array.isArray(selectedWeeks)
    ? selectedWeeks
        .map(Number)
        .filter((week) => Number.isInteger(week) && week >= 1 && week <= 5)
    : [];

  const fallbackWeekCount = String(weekType || '').toLowerCase() === '4week' ? 4 : 5;
  const resolvedWeeks = requestedWeeks.length
    ? requestedWeeks
    : Array.from({ length: fallbackWeekCount }, (_, index) => index + 1);

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

  let rawRows = [];

  try {
    rawRows = await fetchSheetRawRows({
      name: sheetName || period,
      id: sheetId,
      link: sheetLink
    });
  } catch (error) {
    const rawMessage = String(
      error?.message || error || 'Unable to fetch sheet data.'
    );

    const normalized = rawMessage.toLowerCase();

    const fixSuggestion =
      /returned no data rows|no usable|missing|salary sheet/.test(normalized)
        ? 'Make sure the selected Google Sheet tab has valid salary/performance data and the selected user name exists in column B.'
        : /credentials|auth|invalid jwt|permission|forbidden|unauthorized/.test(
            normalized
          )
          ? 'If this sheet is private, provide valid Google Sheets service account credentials or share the sheet with the service account, then restart the server.'
          : 'Check that the sheet link/id is correct, the tab name matches, and the Google Sheet is accessible or published for export.';

    return res.status(400).json({
      ok: false,
      error: rawMessage,
      reason: 'Unable to fetch sheet data.',
      fixSuggestion
    });
  }

  const actualSheetWeekCount = detectSalarySheetWeekCount(rawRows);

  const results = [];
  const validCc = Array.isArray(ccEmails)
    ? ccEmails.filter(Boolean).map(String)
    : [];

  const firstMatchedUser = recipients
    .map((recipient) => parseReportForRecipient(rawRows, recipient, { maxWeekCount: resolvedWeeks.length }))
    .find((item) => item && Array.isArray(item.blocks) && item.blocks.length > 0);

  if (!firstMatchedUser) {
    return res.status(400).json({
      ok: false,
      error: 'No matching sheet data found for the selected users.',
      reason: 'The selected sheet does not contain report blocks for the selected names.',
      fixSuggestion: 'Make sure the selected sheet has rows for the chosen users and that the names or IDs match exactly.'
    });
  }

  for (const recipient of recipients) {
    const name = String(recipient?.name || '').trim();
    const email = String(recipient?.email || '').trim();

    if (!name || !email) {
      const diagnostic = classifyEmailFailure({
        message: 'Missing name or email.'
      });

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

    const matchedUser = parseReportForRecipient(rawRows, recipient, { maxWeekCount: actualSheetWeekCount });

    if (
      !matchedUser ||
      !Array.isArray(matchedUser.blocks) ||
      matchedUser.blocks.length === 0
    ) {
      results.push({
        name,
        email,
        success: false,
        error: 'No matching sheet data found for this user.',
        errorCode: 'NO_SHEET_MATCH',
        reason: 'The selected sheet does not contain report blocks for this user.',
        fixSuggestion:
          'Make sure the selected user name/id matches the row in the salary sheet.'
      });

      continue;
    }

    const filteredBlocks = resolvedWeeks.length
      ? matchedUser.blocks.filter((block) =>
          resolvedWeeks.includes(Number(block.week))
        )
      : matchedUser.blocks;

    if (!filteredBlocks.length) {
      results.push({
        name,
        email,
        success: false,
        error: 'No matching week data found for this user.',
        errorCode: 'NO_WEEK_MATCH',
        reason:
          'The selected user exists, but the requested week data is not available.',
        fixSuggestion: 'Choose a week that exists in the selected sheet.'
      });

      continue;
    }

    const resolvedEmail = email || matchedUser.email || '';

    if (!resolvedEmail) {
      results.push({
        name,
        email,
        success: false,
        error: 'Recipient email is missing.',
        errorCode: 'MISSING_RECIPIENT_EMAIL',
        reason: 'Email address could not be resolved for this user.',
        fixSuggestion: 'Add email to the selected user.'
      });

      continue;
    }

    const filteredTotalPoints = filteredBlocks.reduce(
      (sum, block) => sum + Number(block.subtotal || 0),
      0
    );

    const reportTotalPoints = resolvedWeeks.length
      ? filteredTotalPoints
      : matchedUser.totalPoints || totalPoints;

    const sentAt = formatDateTime();

    const htmlBody = buildReportHtml({
      reportTitle,
      period: matchedUser.period || period,
      name,
      totalLabel,
      totalPoints: reportTotalPoints,
      remarks: matchedUser.remarks || remarks,
      blocks: filteredBlocks,
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
        messageId: info.messageId,
        rowCount: filteredBlocks.reduce(
          (sum, block) =>
            sum + (Array.isArray(block.tasks) ? block.tasks.length : 0),
          0
        ),
        blockCount: filteredBlocks.length,
        selectedWeeks: resolvedWeeks,
        totalPoints: reportTotalPoints,
        debugDateRanges: filteredBlocks.map((block) => ({
          week: block.week,
          dateRange: block.dateRange
        }))
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
