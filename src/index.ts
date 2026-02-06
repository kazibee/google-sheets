import { createAuthClient } from './auth.js';
import { createSheetsClient, type CellRange, type CellFormat } from './sheets-client.js';

let client: ReturnType<typeof createSheetsClient> | null = null;

function getClient() {
  if (!client) {
    const auth = createAuthClient();
    client = createSheetsClient(auth);
  }
  return client;
}

// -- Spreadsheet operations --

export async function createSpreadsheet(title: string) {
  return getClient().createSpreadsheet(title);
}

export async function getSpreadsheet(spreadsheetId: string) {
  return getClient().getSpreadsheet(spreadsheetId);
}

// -- Sheet (tab) operations --

export async function listSheets(spreadsheetId: string) {
  return getClient().listSheets(spreadsheetId);
}

export async function addSheet(spreadsheetId: string, title: string) {
  return getClient().addSheet(spreadsheetId, title);
}

export async function deleteSheet(spreadsheetId: string, sheetId: number) {
  return getClient().deleteSheet(spreadsheetId, sheetId);
}

// -- Data operations --

export async function readRange(spreadsheetId: string, range: string) {
  return getClient().readRange(spreadsheetId, range);
}

export async function writeRange(spreadsheetId: string, range: string, values: string[][]) {
  return getClient().writeRange(spreadsheetId, range, values);
}

export async function appendRows(spreadsheetId: string, range: string, rows: string[][]) {
  return getClient().appendRows(spreadsheetId, range, rows);
}

export async function clearRange(spreadsheetId: string, range: string) {
  return getClient().clearRange(spreadsheetId, range);
}

// -- Formatting --

export async function formatCells(
  spreadsheetId: string,
  sheetId: number,
  range: CellRange,
  format: CellFormat,
) {
  return getClient().formatCells(spreadsheetId, sheetId, range, format);
}
