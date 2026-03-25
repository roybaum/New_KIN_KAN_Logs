const TOKEN_PROPERTY_KEY = 'MASTER_WEBHOOK_TOKEN';
const DEFAULT_SHEET_NAME = 'Inventory';
const DEFAULT_KEY_FIELDS = ['Category', 'Number'];

type IngestCellValue = string | number | boolean | Date | null;
type IngestRow = Record<string, IngestCellValue>;

interface WebhookPayload {
  token?: string;
  rows?: IngestRow[];
  sheetName?: string;
  writeMode?: string;
  keyFields?: string[];
  batchNumber?: number | null;
  batchCount?: number | null;
  runId?: string | null;
  isFirstBatch?: boolean;
}

interface IngestResult {
  processedRows: number;
  skippedRows: number;
}

function doGet(): GoogleAppsScript.Content.TextOutput {
  return jsonResponse({ ok: true, message: 'Master Inventory webhook is running.' });
}

function doPost(e: GoogleAppsScript.Events.DoPost): GoogleAppsScript.Content.TextOutput {
  try {
    const payload = parsePayload(e);
    authorizeRequest(payload);

    const rows: IngestRow[] = Array.isArray(payload.rows) ? payload.rows : [];
    if (rows.length === 0) {
      return jsonResponse({ ok: true, processedRows: 0, skippedRows: 0, message: 'No rows in payload.' });
    }

    const sheetName = String(payload.sheetName || DEFAULT_SHEET_NAME);
    const writeMode = String(payload.writeMode || 'upsert').toLowerCase();
    const keyFields = Array.isArray(payload.keyFields) && payload.keyFields.length > 0
      ? payload.keyFields
      : DEFAULT_KEY_FIELDS;

    if (writeMode !== 'upsert' && writeMode !== 'replace') {
      throw new Error('Unsupported writeMode. Use upsert or replace.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    const result: IngestResult = writeMode === 'replace'
      ? applyReplace(sheet, rows, payload)
      : applyUpsert(sheet, rows, keyFields);

    return jsonResponse({
      ok: true,
      sheetName,
      writeMode,
      batchNumber: payload.batchNumber ?? null,
      batchCount: payload.batchCount ?? null,
      runId: payload.runId ?? null,
      processedRows: result.processedRows,
      skippedRows: result.skippedRows
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: (err as Error).message });
  }
}

function setWebhookToken(token: string): string {
  if (!token || String(token).trim() === '') {
    throw new Error('Token cannot be empty.');
  }
  PropertiesService.getScriptProperties().setProperty(TOKEN_PROPERTY_KEY, String(token).trim());
  return 'Webhook token saved.';
}

function parsePayload(e: GoogleAppsScript.Events.DoPost): WebhookPayload {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing JSON request body.');
  }
  try {
    return JSON.parse(e.postData.contents) as WebhookPayload;
  } catch {
    throw new Error('Invalid JSON body.');
  }
}

function authorizeRequest(payload: WebhookPayload): void {
  const expectedToken = PropertiesService.getScriptProperties().getProperty(TOKEN_PROPERTY_KEY);
  if (!expectedToken) {
    throw new Error('Webhook token is not configured in Script Properties.');
  }
  const actualToken = payload && payload.token ? String(payload.token) : '';
  if (actualToken !== expectedToken) {
    throw new Error('Unauthorized request token.');
  }
}

function applyReplace(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: IngestRow[],
  payload: WebhookPayload
): IngestResult {
  const incomingHeaders = collectIngestHeaders(rows);
  const values = ingestRowsToValues(rows, incomingHeaders);
  const isFirstBatch = payload.isFirstBatch === true;

  if (isFirstBatch) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, incomingHeaders.length).setValues([incomingHeaders]);
    if (values.length > 0) {
      sheet.getRange(2, 1, values.length, incomingHeaders.length).setValues(values);
    }
    sheet.setFrozenRows(1);
    return { processedRows: rows.length, skippedRows: 0 };
  }

  let headers = getIngestSheetHeaders(sheet);
  if (headers.length === 0) {
    headers = incomingHeaders;
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  const alignedValues = ingestRowsToValues(rows, headers);
  const startRow = Math.max(2, sheet.getLastRow() + 1);
  sheet.getRange(startRow, 1, alignedValues.length, headers.length).setValues(alignedValues);

  return { processedRows: rows.length, skippedRows: 0 };
}

function applyUpsert(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: IngestRow[],
  keyFields: string[]
): IngestResult {
  const existingData = sheet.getDataRange().getValues() as IngestCellValue[][];
  const existingHeaders = existingData.length > 0 ? existingData[0].map(String) : [];
  const incomingHeaders = collectIngestHeaders(rows);
  const headers = unionIngestHeaders(existingHeaders, incomingHeaders);

  if (headers.length === 0) {
    throw new Error('No headers available for upsert operation.');
  }

  const recordsByKey: Record<string, IngestRow> = {};
  const orderedKeys: string[] = [];

  const startRow = existingData.length > 0 ? 1 : 0;
  for (let i = startRow; i < existingData.length; i++) {
    const record = ingestRowToObject(existingData[i], existingHeaders);
    const key = buildIngestKey(record, keyFields);
    if (!key) continue;
    if (!recordsByKey[key]) orderedKeys.push(key);
    recordsByKey[key] = record;
  }

  let skippedRows = 0;
  for (const incoming of rows) {
    const key = buildIngestKey(incoming, keyFields);
    if (!key) { skippedRows++; continue; }
    if (!recordsByKey[key]) { orderedKeys.push(key); recordsByKey[key] = {}; }
    recordsByKey[key] = { ...recordsByKey[key], ...incoming };
  }

  const outputValues: IngestCellValue[][] = [headers];
  for (const key of orderedKeys) {
    outputValues.push(ingestObjectToRow(recordsByKey[key], headers));
  }

  sheet.clearContents();
  sheet.getRange(1, 1, outputValues.length, headers.length).setValues(outputValues);
  sheet.setFrozenRows(1);

  return { processedRows: rows.length - skippedRows, skippedRows };
}

function collectIngestHeaders(rows: IngestRow[]): string[] {
  const seen: Record<string, boolean> = {};
  const headers: string[] = [];
  for (const row of rows) {
    for (const key of Object.keys(row || {})) {
      if (!seen[key]) { seen[key] = true; headers.push(key); }
    }
  }
  return headers;
}

function getIngestSheetHeaders(sheet: GoogleAppsScript.Spreadsheet.Sheet): string[] {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
  return (sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] as IngestCellValue[])
    .map(String)
    .filter(h => h.trim() !== '');
}

function unionIngestHeaders(existingHeaders: string[], incomingHeaders: string[]): string[] {
  const seen: Record<string, boolean> = {};
  const headers: string[] = [];
  for (const h of [...existingHeaders, ...incomingHeaders]) {
    if (h && !seen[h]) { seen[h] = true; headers.push(h); }
  }
  return headers;
}

function buildIngestKey(record: IngestRow, keyFields: string[]): string {
  const parts = keyFields.map(field =>
    record && record[field] != null ? String(record[field]).trim() : ''
  );
  const key = parts.join('|');
  return key.replace(/\|/g, '') === '' ? '' : key;
}

function ingestRowToObject(values: IngestCellValue[], headers: string[]): IngestRow {
  const obj: IngestRow = {};
  for (let i = 0; i < headers.length; i++) {
    obj[headers[i]] = values[i];
  }
  return obj;
}

function ingestObjectToRow(obj: IngestRow, headers: string[]): IngestCellValue[] {
  return headers.map(h => (obj[h] !== undefined ? obj[h] : ''));
}

function ingestRowsToValues(rows: IngestRow[], headers: string[]): IngestCellValue[][] {
  return rows.map(row => headers.map(h => (row[h] !== undefined ? row[h] : '')));
}

function jsonResponse(payload: object): GoogleAppsScript.Content.TextOutput {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
