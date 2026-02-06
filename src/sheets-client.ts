import { google, type Auth } from 'googleapis';
import type { sheets_v4 } from 'googleapis';

type SheetsAPI = sheets_v4.Sheets;

export function createSheetsClient(auth: Auth.OAuth2Client) {
  const sheets = google.sheets({ version: 'v4', auth });

  return {
    createSpreadsheet: (title: string) => createSpreadsheet(sheets, title),
    getSpreadsheet: (spreadsheetId: string) => getSpreadsheet(sheets, spreadsheetId),
    listSheets: (spreadsheetId: string) => listSheets(sheets, spreadsheetId),
    addSheet: (spreadsheetId: string, title: string) => addSheet(sheets, spreadsheetId, title),
    deleteSheet: (spreadsheetId: string, sheetId: number) => deleteSheet(sheets, spreadsheetId, sheetId),
    readRange: (spreadsheetId: string, range: string) => readRange(sheets, spreadsheetId, range),
    writeRange: (spreadsheetId: string, range: string, values: string[][]) =>
      writeRange(sheets, spreadsheetId, range, values),
    appendRows: (spreadsheetId: string, range: string, rows: string[][]) =>
      appendRows(sheets, spreadsheetId, range, rows),
    clearRange: (spreadsheetId: string, range: string) => clearRange(sheets, spreadsheetId, range),
    formatCells: (spreadsheetId: string, sheetId: number, range: CellRange, format: CellFormat) =>
      formatCells(sheets, spreadsheetId, sheetId, range, format),
  };
}

// -- Types --

export interface SpreadsheetInfo {
  spreadsheetId: string;
  title: string;
  url: string;
  sheets: SheetInfo[];
}

export interface SheetInfo {
  sheetId: number;
  title: string;
  index: number;
  rowCount: number;
  columnCount: number;
}

export interface CellRange {
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
}

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  foregroundColor?: { red?: number; green?: number; blue?: number };
  backgroundColor?: { red?: number; green?: number; blue?: number };
  horizontalAlignment?: 'LEFT' | 'CENTER' | 'RIGHT';
}

// -- Spreadsheet operations --

async function createSpreadsheet(sheets: SheetsAPI, title: string): Promise<SpreadsheetInfo> {
  const res = await sheets.spreadsheets.create({
    requestBody: { properties: { title } },
  });
  return mapSpreadsheet(res.data);
}

async function getSpreadsheet(sheets: SheetsAPI, spreadsheetId: string): Promise<SpreadsheetInfo> {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return mapSpreadsheet(res.data);
}

// -- Sheet (tab) operations --

async function listSheets(sheets: SheetsAPI, spreadsheetId: string): Promise<SheetInfo[]> {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return (res.data.sheets ?? []).map(mapSheet);
}

async function addSheet(sheets: SheetsAPI, spreadsheetId: string, title: string): Promise<SheetInfo> {
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ addSheet: { properties: { title } } }],
    },
  });
  const reply = res.data.replies?.[0]?.addSheet;
  if (!reply?.properties) {
    throw new Error('Failed to add sheet — no response from API');
  }
  return mapSheet({ properties: reply.properties });
}

async function deleteSheet(sheets: SheetsAPI, spreadsheetId: string, sheetId: number): Promise<void> {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ deleteSheet: { sheetId } }],
    },
  });
}

// -- Data operations --

async function readRange(sheets: SheetsAPI, spreadsheetId: string, range: string): Promise<string[][]> {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  return (res.data.values as string[][] | undefined) ?? [];
}

async function writeRange(
  sheets: SheetsAPI,
  spreadsheetId: string,
  range: string,
  values: string[][],
): Promise<{ updatedCells: number }> {
  const res = await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values },
  });
  return { updatedCells: res.data.updatedCells ?? 0 };
}

async function appendRows(
  sheets: SheetsAPI,
  spreadsheetId: string,
  range: string,
  rows: string[][],
): Promise<{ updatedCells: number }> {
  const res = await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: rows },
  });
  return { updatedCells: res.data.updates?.updatedCells ?? 0 };
}

async function clearRange(sheets: SheetsAPI, spreadsheetId: string, range: string): Promise<void> {
  await sheets.spreadsheets.values.clear({
    spreadsheetId,
    range,
    requestBody: {},
  });
}

// -- Formatting --

async function formatCells(
  sheets: SheetsAPI,
  spreadsheetId: string,
  sheetId: number,
  range: CellRange,
  format: CellFormat,
): Promise<void> {
  const textFormat: sheets_v4.Schema$TextFormat = {};
  if (format.bold !== undefined) textFormat.bold = format.bold;
  if (format.italic !== undefined) textFormat.italic = format.italic;
  if (format.fontSize !== undefined) textFormat.fontSize = format.fontSize;
  if (format.foregroundColor) {
    textFormat.foregroundColorStyle = {
      rgbColor: format.foregroundColor,
    };
  }

  const cellFormat: sheets_v4.Schema$CellFormat = { textFormat };
  if (format.backgroundColor) {
    cellFormat.backgroundColorStyle = {
      rgbColor: format.backgroundColor,
    };
  }
  if (format.horizontalAlignment) {
    cellFormat.horizontalAlignment = format.horizontalAlignment;
  }

  const fields: string[] = [];
  if (format.bold !== undefined) fields.push('userEnteredFormat.textFormat.bold');
  if (format.italic !== undefined) fields.push('userEnteredFormat.textFormat.italic');
  if (format.fontSize !== undefined) fields.push('userEnteredFormat.textFormat.fontSize');
  if (format.foregroundColor) fields.push('userEnteredFormat.textFormat.foregroundColorStyle');
  if (format.backgroundColor) fields.push('userEnteredFormat.backgroundColorStyle');
  if (format.horizontalAlignment) fields.push('userEnteredFormat.horizontalAlignment');

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: { sheetId, ...range },
            cell: { userEnteredFormat: cellFormat },
            fields: fields.join(','),
          },
        },
      ],
    },
  });
}

// -- Mappers --

function mapSpreadsheet(data: sheets_v4.Schema$Spreadsheet): SpreadsheetInfo {
  return {
    spreadsheetId: data.spreadsheetId!,
    title: data.properties?.title ?? '',
    url: data.spreadsheetUrl ?? '',
    sheets: (data.sheets ?? []).map(mapSheet),
  };
}

function mapSheet(sheet: sheets_v4.Schema$Sheet): SheetInfo {
  const props = sheet.properties!;
  return {
    sheetId: props.sheetId ?? 0,
    title: props.title ?? '',
    index: props.index ?? 0,
    rowCount: props.gridProperties?.rowCount ?? 0,
    columnCount: props.gridProperties?.columnCount ?? 0,
  };
}
