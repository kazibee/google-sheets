import { createAuthClient, type Env } from './auth';
import { createSheetsClient } from './sheets-client';

export type { Env } from './auth';
export type { SpreadsheetInfo, SheetInfo, CellRange, CellFormat } from './sheets-client';

export default function main(env: Env) {
  const auth = createAuthClient(env);
  return createSheetsClient(auth);
}
