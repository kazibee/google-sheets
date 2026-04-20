import { sheets as createSheets, type sheets_v4 } from '@googleapis/sheets';
import type { OAuth2Client } from 'google-auth-library';

type SheetsAPI = sheets_v4.Sheets;

/**
 * Create a Google Sheets API client bound to the given OAuth2 credentials.
 * Returns an object exposing all direct methods and the batch builder factory.
 */
export function createSheetsClient(auth: OAuth2Client) {
  const sheets = createSheets({ version: 'v4', auth });

  const client = {
    /** Create a new spreadsheet with the given title. */
    createSpreadsheet: (title: string) => createSpreadsheet(sheets, title),
    /** Alias for createSpreadsheet. */
    create: (title: string) => createSpreadsheet(sheets, title),
    /** Retrieve metadata for an existing spreadsheet. */
    getSpreadsheet: (spreadsheetId: string) => getSpreadsheet(sheets, spreadsheetId),
    /** List all sheets (tabs) within a spreadsheet. */
    listSheets: (spreadsheetId: string) => listSheets(sheets, spreadsheetId),
    /** Add a new sheet (tab) with the given title. */
    addSheet: (spreadsheetId: string, title: string) => addSheet(sheets, spreadsheetId, title),
    /** Delete a sheet (tab) by its numeric sheetId. */
    deleteSheet: (spreadsheetId: string, sheetId: number) => deleteSheet(sheets, spreadsheetId, sheetId),
    /** Read cell values from an A1-notation range. */
    readRange: (spreadsheetId: string, range: string) => readRange(sheets, spreadsheetId, range),
    /** Read multiple ranges in a single API call. */
    readRanges: (spreadsheetId: string, ranges: string[]) => readRanges(sheets, spreadsheetId, ranges),
    /** Write cell values to an A1-notation range, replacing existing data. */
    writeRange: (spreadsheetId: string, range: string, values: string[][]) =>
      writeRange(sheets, spreadsheetId, range, values),
    /** Alias for writeRange. */
    updateRange: (spreadsheetId: string, range: string, values: string[][]) =>
      writeRange(sheets, spreadsheetId, range, values),
    /** Append rows after existing data in the given range. */
    appendRows: (spreadsheetId: string, range: string, rows: string[][]) =>
      appendRows(sheets, spreadsheetId, range, rows),
    /** Add a named range to a spreadsheet. */
    addNamedRange: (spreadsheetId: string, name: string, sheetId: number, range: CellRange) =>
      addNamedRange(sheets, spreadsheetId, name, sheetId, range),
    /** Delete a named range by its ID. */
    deleteNamedRange: (spreadsheetId: string, namedRangeId: string) =>
      deleteNamedRange(sheets, spreadsheetId, namedRangeId),
    /** List all named ranges in a spreadsheet. */
    listNamedRanges: (spreadsheetId: string) => listNamedRanges(sheets, spreadsheetId),
    /** Create a batch builder to queue multiple write/format operations for a single send. */
    batch: (spreadsheetId: string) => createBatchBuilder(sheets, spreadsheetId),
  };

  return client;
}

// -- Types --

/** Basic spreadsheet metadata returned from create/get operations. */
export interface SpreadsheetInfo {
  /** Unique identifier for the spreadsheet. */
  spreadsheetId: string;
  /** Display title of the spreadsheet. */
  title: string;
  /** Web URL to open the spreadsheet in a browser. */
  url: string;
  /** All sheets (tabs) within the spreadsheet. */
  sheets: SheetInfo[];
}

/** Metadata for a single sheet (tab) within a spreadsheet. */
export interface SheetInfo {
  /** Numeric ID of the sheet (used in batchUpdate requests). */
  sheetId: number;
  /** Display name of the sheet tab. */
  title: string;
  /** Zero-based position of the sheet tab. */
  index: number;
  /** Number of rows in the sheet grid. */
  rowCount: number;
  /** Number of columns in the sheet grid. */
  columnCount: number;
}

/** Zero-based cell range used by batchUpdate requests. */
export interface CellRange {
  /** First row (inclusive, zero-based). */
  startRowIndex: number;
  /** Last row (exclusive, zero-based). */
  endRowIndex: number;
  /** First column (inclusive, zero-based). */
  startColumnIndex: number;
  /** Last column (exclusive, zero-based). */
  endColumnIndex: number;
}

/** RGB color specification with values in the 0-1 range. */
export interface RGBColor {
  red?: number;
  green?: number;
  blue?: number;
}

/** Number format type applied to cell values. */
export type NumberFormatType = 'TEXT' | 'NUMBER' | 'PERCENT' | 'CURRENCY' | 'DATE' | 'TIME' | 'SCIENTIFIC';

/** Number format specification for cell values. */
export interface NumberFormat {
  /** The type of number format. */
  type: NumberFormatType;
  /** Optional ICU pattern string (e.g. "#,##0.00" for numbers, "yyyy-mm-dd" for dates). */
  pattern?: string;
}

/** Line style for a single border edge. */
export type BorderStyleType = 'SOLID' | 'DASHED' | 'DOTTED' | 'SOLID_MEDIUM' | 'SOLID_THICK' | 'DOUBLE';

/** Defines the style and color for a single border edge. */
export interface BorderStyle {
  /** The line style of the border. */
  style: BorderStyleType;
  /** Optional color for the border line. */
  color?: RGBColor;
}

/** Border configuration for all four edges of a cell range. */
export interface Borders {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
}

/** Cell formatting options applied via repeatCell or updateBorders requests. */
export interface CellFormat {
  /** Whether the text is bold. */
  bold?: boolean;
  /** Whether the text is italic. */
  italic?: boolean;
  /** Font size in points. */
  fontSize?: number;
  /** Text (foreground) color. */
  foregroundColor?: RGBColor;
  /** Cell background color. */
  backgroundColor?: RGBColor;
  /** Horizontal text alignment within the cell. */
  horizontalAlignment?: 'LEFT' | 'CENTER' | 'RIGHT';
  /** Number format applied to cell values. */
  numberFormat?: NumberFormat;
  /** Border styles for the cell range. Processed as a separate updateBorders request. */
  borders?: Borders;
}

/** Sort order direction. */
export type SortOrder = 'ASCENDING' | 'DESCENDING';

/** Specifies a single column sort within a sortRange request. */
export interface SortSpec {
  /** Zero-based column index to sort by. */
  dimensionIndex: number;
  /** Sort direction. */
  sortOrder: SortOrder;
}

/** Supported chart types for addChart. */
export type ChartType = 'BAR' | 'LINE' | 'PIE' | 'COLUMN' | 'AREA' | 'SCATTER';

/** Overlay position for a chart anchored to a cell. */
export interface ChartOverlayPosition {
  /** The cell the chart is anchored to. */
  anchorCell: { sheetId: number; rowIndex: number; columnIndex: number };
  /** Offset in pixels from the anchor cell. */
  offsetXPixels?: number;
  /** Offset in pixels from the anchor cell. */
  offsetYPixels?: number;
  /** Width of the chart in pixels. */
  widthPixels?: number;
  /** Height of the chart in pixels. */
  heightPixels?: number;
}

/** Data source range for a chart, specified as a CellRange with sheetId. */
export interface ChartSourceRange {
  sheetId: number;
  range: CellRange;
}

/** Configuration for creating an embedded chart. */
export interface ChartConfig {
  /** The type of chart to create. */
  chartType: ChartType;
  /** Optional chart title. */
  title?: string;
  /** Data source ranges for the chart. */
  sourceRanges: ChartSourceRange[];
  /** Position overlay for the chart. */
  position: ChartOverlayPosition;
}

/** Condition type for boolean-based conditional formatting rules. */
export type ConditionType =
  | 'NUMBER_GREATER' | 'NUMBER_GREATER_THAN_EQ'
  | 'NUMBER_LESS' | 'NUMBER_LESS_THAN_EQ'
  | 'NUMBER_EQ' | 'NUMBER_NOT_EQ'
  | 'NUMBER_BETWEEN' | 'NUMBER_NOT_BETWEEN'
  | 'TEXT_CONTAINS' | 'TEXT_NOT_CONTAINS'
  | 'TEXT_STARTS_WITH' | 'TEXT_ENDS_WITH'
  | 'TEXT_EQ' | 'TEXT_IS_EMAIL' | 'TEXT_IS_URL'
  | 'DATE_EQ' | 'DATE_BEFORE' | 'DATE_AFTER'
  | 'DATE_ON_OR_BEFORE' | 'DATE_ON_OR_AFTER'
  | 'DATE_BETWEEN' | 'DATE_NOT_BETWEEN'
  | 'DATE_IS_VALID'
  | 'ONE_OF_RANGE' | 'ONE_OF_LIST'
  | 'BLANK' | 'NOT_BLANK'
  | 'CUSTOM_FORMULA'
  | 'BOOLEAN';

/** A boolean condition with comparison values. */
export interface BooleanCondition {
  /** The type of condition to evaluate. */
  type: ConditionType;
  /** Values used by the condition (e.g. the threshold number, the formula string). */
  values?: Array<{ userEnteredValue?: string; relativeDate?: string }>;
}

/** Boolean-based conditional format rule: condition + format to apply. */
export interface BooleanRule {
  /** The condition that triggers the formatting. */
  condition: BooleanCondition;
  /** The format to apply when the condition is met. */
  format: CellFormat;
}

/** Interpolation point for gradient rules. */
export interface InterpolationPoint {
  /** The color at this point. */
  color: RGBColor;
  /** The type of value this point represents. */
  type: 'MIN' | 'MAX' | 'NUMBER' | 'PERCENT' | 'PERCENTILE';
  /** The value for this point (required when type is NUMBER, PERCENT, or PERCENTILE). */
  value?: string;
}

/** Gradient-based conditional format rule. */
export interface GradientRule {
  /** The starting point of the gradient. */
  minpoint: InterpolationPoint;
  /** The optional midpoint of the gradient. */
  midpoint?: InterpolationPoint;
  /** The ending point of the gradient. */
  maxpoint: InterpolationPoint;
}

/** A conditional formatting rule applied to one or more ranges. */
export interface ConditionalFormatRule {
  /** Cell ranges to apply this rule to. */
  ranges: CellRange[];
  /** Boolean-based rule (mutually exclusive with gradientRule). */
  booleanRule?: BooleanRule;
  /** Gradient-based rule (mutually exclusive with booleanRule). */
  gradientRule?: GradientRule;
}

/** Merge type for mergeCells request. */
export type MergeType = 'MERGE_ALL' | 'MERGE_COLUMNS' | 'MERGE_ROWS';

// -- Task #4: Data management types --

/** Validation rule type for data validation. */
export type DataValidationType =
  | 'ONE_OF_LIST'
  | 'ONE_OF_RANGE'
  | 'NUMBER_BETWEEN'
  | 'NUMBER_GREATER'
  | 'NUMBER_LESS'
  | 'BOOLEAN'
  | 'CUSTOM_FORMULA';

/** Data validation rule applied to a cell range. */
export interface DataValidationRule {
  /** The type of validation to apply. */
  type: DataValidationType;
  /** List of allowed values (for ONE_OF_LIST). */
  values?: string[];
  /** Formula string (for CUSTOM_FORMULA or ONE_OF_RANGE). */
  formula?: string;
  /** Minimum value (for NUMBER_BETWEEN). */
  min?: number;
  /** Maximum value (for NUMBER_BETWEEN). */
  max?: number;
  /** If true, reject invalid input; if false, show a warning only. Defaults to true. */
  strict?: boolean;
  /** If true, show a dropdown arrow in the cell (for ONE_OF_LIST). */
  showCustomUi?: boolean;
}

/** Summarize function used by pivot table value aggregation. */
export type PivotSummarizeFunction = 'SUM' | 'COUNT' | 'AVERAGE' | 'MAX' | 'MIN' | 'CUSTOM';

/** A pivot table row or column grouping. */
export interface PivotGroup {
  /** Zero-based column offset in the source data to group by. */
  sourceColumnOffset: number;
  /** Sort order for this group. */
  sortOrder?: SortOrder;
}

/** A pivot table value (aggregated column). */
export interface PivotValue {
  /** Zero-based column offset in the source data to aggregate. */
  sourceColumnOffset: number;
  /** The aggregation function. */
  summarizeFunction: PivotSummarizeFunction;
}

/** Configuration for creating a pivot table. */
export interface PivotTableConfig {
  /** Source data range for the pivot table. */
  sourceRange: { sheetId: number; range: CellRange };
  /** Row groupings. */
  rows: PivotGroup[];
  /** Column groupings. */
  columns: PivotGroup[];
  /** Aggregated values. */
  values: PivotValue[];
  /** Cell where the pivot table is anchored. */
  anchorCell: { sheetId: number; rowIndex: number; columnIndex: number };
}

/** Configuration for find and replace operations. */
export interface FindReplaceConfig {
  /** The string to find. */
  find: string;
  /** The replacement string. */
  replacement: string;
  /** Whether the search is case-sensitive. */
  matchCase?: boolean;
  /** Whether to match the entire cell content. */
  matchEntireCell?: boolean;
  /** Whether to treat the find string as a regex. */
  searchByRegex?: boolean;
  /** Whether to search all sheets instead of just the specified one. */
  allSheets?: boolean;
}

/** Named range within a spreadsheet. */
export interface NamedRange {
  /** The unique ID of the named range. */
  namedRangeId: string;
  /** The name of the range. */
  name: string;
  /** The sheet ID containing the range. */
  sheetId: number;
  /** The cell range. */
  range: CellRange;
}

/** Result from a batch read of multiple ranges. */
export interface RangeValues {
  /** The A1-notation range that was read. */
  range: string;
  /** The values in the range. */
  values: string[][];
}

// -- Task #5: Structure and layout types --

/** Configuration for dimension grouping operations. */
export type DimensionType = 'ROWS' | 'COLUMNS';

// -- Task #6: Collaboration and rich content types --

/** Location type for developer metadata. */
export type MetadataLocationType = 'SHEET' | 'ROW' | 'COLUMN';

/** A note to set on a specific cell. */
export interface CellNote {
  /** Zero-based row index. */
  row: number;
  /** Zero-based column index. */
  col: number;
  /** The note text. */
  note: string;
}

// -- Spreadsheet operations --

/** Create a new Google Sheets spreadsheet with the given title. */
async function createSpreadsheet(sheets: SheetsAPI, title: string): Promise<SpreadsheetInfo> {
  const res = await sheets.spreadsheets.create({
    requestBody: { properties: { title } },
  });
  return mapSpreadsheet(res.data);
}

/** Retrieve metadata (title, URL, sheets) for an existing spreadsheet. */
async function getSpreadsheet(sheets: SheetsAPI, spreadsheetId: string): Promise<SpreadsheetInfo> {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return mapSpreadsheet(res.data);
}

// -- Sheet (tab) operations --

/** List all sheets (tabs) in a spreadsheet, returning metadata for each. */
async function listSheets(sheets: SheetsAPI, spreadsheetId: string): Promise<SheetInfo[]> {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return (res.data.sheets ?? []).map(mapSheet);
}

/** Add a new sheet (tab) with the given title and return its metadata. */
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

/** Delete a sheet (tab) by its numeric sheetId. */
async function deleteSheet(sheets: SheetsAPI, spreadsheetId: string, sheetId: number): Promise<void> {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ deleteSheet: { sheetId } }],
    },
  });
}

// -- Data operations --

/** Read cell values from an A1-notation range, returning a 2D string array. */
async function readRange(sheets: SheetsAPI, spreadsheetId: string, range: string): Promise<string[][]> {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  return (res.data.values as string[][] | undefined) ?? [];
}

/** Write cell values to an A1-notation range, replacing any existing data. */
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

/** Append rows after the last row of existing data in the specified range. */
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

/** Read multiple ranges in a single API call using batchGet. */
async function readRanges(
  sheets: SheetsAPI,
  spreadsheetId: string,
  ranges: string[],
): Promise<RangeValues[]> {
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId,
    ranges,
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  return (res.data.valueRanges ?? []).map((vr) => ({
    range: vr.range ?? '',
    values: (vr.values as string[][] | undefined) ?? [],
  }));
}

// -- Named range operations --

/** Add a named range to a spreadsheet. */
async function addNamedRange(
  sheets: SheetsAPI,
  spreadsheetId: string,
  name: string,
  sheetId: number,
  range: CellRange,
): Promise<NamedRange> {
  const res = await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          addNamedRange: {
            namedRange: {
              name,
              range: { sheetId, ...range },
            },
          },
        },
      ],
    },
  });
  const reply = res.data.replies?.[0]?.addNamedRange?.namedRange;
  if (!reply) {
    throw new Error('Failed to add named range — no response from API');
  }
  return {
    namedRangeId: reply.namedRangeId ?? '',
    name: reply.name ?? '',
    sheetId: reply.range?.sheetId ?? 0,
    range: {
      startRowIndex: reply.range?.startRowIndex ?? 0,
      endRowIndex: reply.range?.endRowIndex ?? 0,
      startColumnIndex: reply.range?.startColumnIndex ?? 0,
      endColumnIndex: reply.range?.endColumnIndex ?? 0,
    },
  };
}

/** Delete a named range by its ID. */
async function deleteNamedRange(
  sheets: SheetsAPI,
  spreadsheetId: string,
  namedRangeId: string,
): Promise<void> {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{ deleteNamedRange: { namedRangeId } }],
    },
  });
}

/** List all named ranges in a spreadsheet. */
async function listNamedRanges(
  sheets: SheetsAPI,
  spreadsheetId: string,
): Promise<NamedRange[]> {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  return (res.data.namedRanges ?? []).map((nr) => ({
    namedRangeId: nr.namedRangeId ?? '',
    name: nr.name ?? '',
    sheetId: nr.range?.sheetId ?? 0,
    range: {
      startRowIndex: nr.range?.startRowIndex ?? 0,
      endRowIndex: nr.range?.endRowIndex ?? 0,
      startColumnIndex: nr.range?.startColumnIndex ?? 0,
      endColumnIndex: nr.range?.endColumnIndex ?? 0,
    },
  }));
}

// -- Batch builder --

/** A queued cell-formatting operation within a batch (sheetId + range + format). */
export interface FormatOperation {
  /** Numeric ID of the target sheet. */
  sheetId: number;
  /** Zero-based cell range to format. */
  range: CellRange;
  /** Formatting options to apply. */
  format: CellFormat;
}

/** Aggregated result returned after a batch send completes. */
export interface BatchResult {
  /** Total number of cells written via batchUpdate values. */
  writtenCells: number;
  /** Number of ranges cleared via batchClear. */
  clearedRanges: number;
  /** Number of format operations executed. */
  formatOperations: number;
}

/** Builds a repeatCell request from a FormatOperation (excludes borders, handled separately). */
function buildFormatRequest(op: FormatOperation): sheets_v4.Schema$Request {
  const textFormat: sheets_v4.Schema$TextFormat = {};
  if (op.format.bold !== undefined) textFormat.bold = op.format.bold;
  if (op.format.italic !== undefined) textFormat.italic = op.format.italic;
  if (op.format.fontSize !== undefined) textFormat.fontSize = op.format.fontSize;
  if (op.format.foregroundColor) {
    textFormat.foregroundColorStyle = { rgbColor: op.format.foregroundColor };
  }

  const cellFormat: sheets_v4.Schema$CellFormat = { textFormat };
  if (op.format.backgroundColor) {
    cellFormat.backgroundColorStyle = { rgbColor: op.format.backgroundColor };
  }
  if (op.format.horizontalAlignment) {
    cellFormat.horizontalAlignment = op.format.horizontalAlignment;
  }
  if (op.format.numberFormat) {
    cellFormat.numberFormat = {
      type: op.format.numberFormat.type,
      pattern: op.format.numberFormat.pattern ?? '',
    };
  }

  const fields: string[] = [];
  if (op.format.bold !== undefined) fields.push('userEnteredFormat.textFormat.bold');
  if (op.format.italic !== undefined) fields.push('userEnteredFormat.textFormat.italic');
  if (op.format.fontSize !== undefined) fields.push('userEnteredFormat.textFormat.fontSize');
  if (op.format.foregroundColor) fields.push('userEnteredFormat.textFormat.foregroundColorStyle');
  if (op.format.backgroundColor) fields.push('userEnteredFormat.backgroundColorStyle');
  if (op.format.horizontalAlignment) fields.push('userEnteredFormat.horizontalAlignment');
  if (op.format.numberFormat) fields.push('userEnteredFormat.numberFormat');

  return {
    repeatCell: {
      range: { sheetId: op.sheetId, ...op.range },
      cell: { userEnteredFormat: cellFormat },
      fields: fields.join(','),
    },
  };
}

/** Converts a BorderStyle to the Sheets API border object. */
function mapBorderStyle(b: BorderStyle): sheets_v4.Schema$Border {
  return {
    style: b.style,
    ...(b.color ? { colorStyle: { rgbColor: b.color } } : {}),
  };
}

/** Builds an updateBorders request for a FormatOperation that has borders. */
function buildBorderRequest(op: FormatOperation): sheets_v4.Schema$Request | null {
  if (!op.format.borders) return null;
  const { borders } = op.format;
  const borderReq: sheets_v4.Schema$UpdateBordersRequest = {
    range: { sheetId: op.sheetId, ...op.range },
  };
  if (borders.top) borderReq.top = mapBorderStyle(borders.top);
  if (borders.bottom) borderReq.bottom = mapBorderStyle(borders.bottom);
  if (borders.left) borderReq.left = mapBorderStyle(borders.left);
  if (borders.right) borderReq.right = mapBorderStyle(borders.right);
  return { updateBorders: borderReq };
}

/** Chainable builder for collecting multiple Sheets API operations into a single batch send. */
export interface BatchBuilder {
  /** Queue a range write operation (values.batchUpdate). */
  writeRange(range: string, values: string[][]): BatchBuilder;
  /** Queue a range clear operation (values.batchClear). */
  clearRange(range: string): BatchBuilder;
  /** Queue cell formatting (repeatCell + optional updateBorders). */
  formatCells(sheetId: number, range: CellRange, format: CellFormat): BatchBuilder;
  /** Queue a sort operation on a range. */
  sortRange(sheetId: number, range: CellRange, sortSpecs: SortSpec[]): BatchBuilder;
  /** Queue an embedded chart creation. */
  addChart(sheetId: number, chartConfig: ChartConfig): BatchBuilder;
  /** Queue a conditional formatting rule. */
  addConditionalFormat(sheetId: number, rule: ConditionalFormatRule): BatchBuilder;
  /** Queue a merge cells operation. */
  mergeCells(sheetId: number, range: CellRange, mergeType?: MergeType): BatchBuilder;
  /** Queue an unmerge cells operation. */
  unmergeCells(sheetId: number, range: CellRange): BatchBuilder;
  /** Queue a basic filter on a range. */
  setBasicFilter(sheetId: number, range: CellRange): BatchBuilder;
  /** Queue removal of the basic filter from a sheet. */
  removeBasicFilter(sheetId: number): BatchBuilder;
  /** Queue a column or row resize operation. */
  resizeDimension(sheetId: number, dimension: 'ROWS' | 'COLUMNS', startIndex: number, endIndex: number, pixelSize: number): BatchBuilder;

  // -- Task #4: Data management --

  /** Queue a data validation rule on a range. */
  setDataValidation(sheetId: number, range: CellRange, rule: DataValidationRule): BatchBuilder;
  /** Queue removal of data validation from a range. */
  removeDataValidation(sheetId: number, range: CellRange): BatchBuilder;
  /** Queue a pivot table creation. */
  addPivotTable(sheetId: number, config: PivotTableConfig): BatchBuilder;
  /** Queue a find and replace operation. */
  findReplace(sheetId: number, config: FindReplaceConfig): BatchBuilder;

  // -- Task #5: Structure and layout --

  /** Queue freezing rows and/or columns on a sheet. */
  freezeDimensions(sheetId: number, frozenRowCount?: number, frozenColumnCount?: number): BatchBuilder;
  /** Queue a dimension group (outline) creation. */
  groupDimension(sheetId: number, dimension: DimensionType, startIndex: number, endIndex: number): BatchBuilder;
  /** Queue deletion of a dimension group. */
  ungroupDimension(sheetId: number, dimension: DimensionType, startIndex: number, endIndex: number): BatchBuilder;
  /** Queue collapsing or expanding a dimension group. */
  collapseGroup(sheetId: number, dimension: DimensionType, startIndex: number, endIndex: number, collapsed: boolean): BatchBuilder;
  /** Queue hiding a sheet. */
  hideSheet(sheetId: number): BatchBuilder;
  /** Queue unhiding a sheet. */
  unhideSheet(sheetId: number): BatchBuilder;
  /** Queue duplication of a sheet. */
  duplicateSheet(sheetId: number, newName?: string, insertAtIndex?: number): BatchBuilder;
  /** Queue auto-resizing rows or columns to fit content. */
  autoResizeDimension(sheetId: number, dimension: DimensionType, startIndex: number, endIndex: number): BatchBuilder;
  /** Queue moving rows or columns to a new position. */
  moveDimension(sheetId: number, dimension: DimensionType, sourceStartIndex: number, sourceEndIndex: number, destinationIndex: number): BatchBuilder;
  /** Queue setting the tab color of a sheet. */
  setTabColor(sheetId: number, color: RGBColor): BatchBuilder;

  // -- Task #6: Collaboration and rich content --

  /** Queue a protected range. */
  protectRange(sheetId: number, range: CellRange, description?: string, warningOnly?: boolean): BatchBuilder;
  /** Queue removal of a protected range by its ID. */
  removeProtection(protectedRangeId: number): BatchBuilder;
  /** Queue setting a note on a single cell. */
  setNote(sheetId: number, row: number, col: number, note: string): BatchBuilder;
  /** Queue clearing a note from a single cell. */
  clearNote(sheetId: number, row: number, col: number): BatchBuilder;
  /** Queue setting notes on multiple cells. */
  setNotes(sheetId: number, notes: CellNote[]): BatchBuilder;
  /** Queue setting a hyperlink on a cell using =HYPERLINK formula. */
  setHyperlink(sheetId: number, row: number, col: number, url: string, label?: string): BatchBuilder;
  /** Queue inserting an image into a cell using =IMAGE formula. */
  insertImage(sheetId: number, row: number, col: number, imageUrl: string): BatchBuilder;
  /** Queue adding developer metadata to a sheet, row, or column. */
  addDeveloperMetadata(sheetId: number, key: string, value: string, location: MetadataLocationType, locationIndex?: number): BatchBuilder;
  /** Queue deletion of developer metadata by its ID. */
  deleteDeveloperMetadata(metadataId: number): BatchBuilder;

  /** Execute all queued operations and return aggregated results. */
  send(): Promise<BatchResult>;
}

/**
 * Create a chainable batch builder that collects writes, clears, format,
 * and structural operations into a single `send()` call.
 */
function createBatchBuilder(sheets: SheetsAPI, spreadsheetId: string): BatchBuilder {
  const writes: { range: string; values: string[][] }[] = [];
  const clears: string[] = [];
  const formats: FormatOperation[] = [];
  const requests: sheets_v4.Schema$Request[] = [];

  const builder: BatchBuilder = {
    writeRange(range: string, values: string[][]) {
      writes.push({ range, values });
      return builder;
    },
    clearRange(range: string) {
      clears.push(range);
      return builder;
    },
    formatCells(sheetId: number, range: CellRange, format: CellFormat) {
      formats.push({ sheetId, range, format });
      return builder;
    },
    sortRange(sheetId: number, range: CellRange, sortSpecs: SortSpec[]) {
      requests.push({
        sortRange: {
          range: { sheetId, ...range },
          sortSpecs: sortSpecs.map((s) => ({
            dimensionIndex: s.dimensionIndex,
            sortOrder: s.sortOrder,
          })),
        },
      });
      return builder;
    },
    addChart(_sheetId: number, chartConfig: ChartConfig) {
      const sourceSheetsRanges = chartConfig.sourceRanges.map((sr) => ({
        sheetId: sr.sheetId,
        ...sr.range,
      }));

      const spec: sheets_v4.Schema$ChartSpec = {
        title: chartConfig.title,
        basicChart: {
          chartType: chartConfig.chartType,
          domains: [
            {
              domain: {
                sourceRange: { sources: sourceSheetsRanges },
              },
            },
          ],
          series: [
            {
              series: {
                sourceRange: { sources: sourceSheetsRanges },
              },
            },
          ],
        },
      };

      const position: sheets_v4.Schema$EmbeddedObjectPosition = {
        overlayPosition: {
          anchorCell: chartConfig.position.anchorCell,
          offsetXPixels: chartConfig.position.offsetXPixels ?? 0,
          offsetYPixels: chartConfig.position.offsetYPixels ?? 0,
          widthPixels: chartConfig.position.widthPixels ?? 600,
          heightPixels: chartConfig.position.heightPixels ?? 371,
        },
      };

      requests.push({
        addChart: {
          chart: { spec, position },
        },
      });
      return builder;
    },
    addConditionalFormat(sheetId: number, rule: ConditionalFormatRule) {
      const ranges = rule.ranges.map((r) => ({ sheetId, ...r }));

      const apiRule: sheets_v4.Schema$ConditionalFormatRule = { ranges };

      if (rule.booleanRule) {
        const fmt = rule.booleanRule.format;
        const cellFmt: sheets_v4.Schema$CellFormat = {};
        if (fmt.bold !== undefined || fmt.italic !== undefined || fmt.fontSize !== undefined || fmt.foregroundColor) {
          const tf: sheets_v4.Schema$TextFormat = {};
          if (fmt.bold !== undefined) tf.bold = fmt.bold;
          if (fmt.italic !== undefined) tf.italic = fmt.italic;
          if (fmt.fontSize !== undefined) tf.fontSize = fmt.fontSize;
          if (fmt.foregroundColor) tf.foregroundColorStyle = { rgbColor: fmt.foregroundColor };
          cellFmt.textFormat = tf;
        }
        if (fmt.backgroundColor) cellFmt.backgroundColorStyle = { rgbColor: fmt.backgroundColor };

        apiRule.booleanRule = {
          condition: {
            type: rule.booleanRule.condition.type,
            values: rule.booleanRule.condition.values?.map((v) => ({
              userEnteredValue: v.userEnteredValue,
              relativeDate: v.relativeDate,
            })),
          },
          format: cellFmt,
        };
      }

      if (rule.gradientRule) {
        const mapPoint = (p: InterpolationPoint): sheets_v4.Schema$InterpolationPoint => ({
          colorStyle: { rgbColor: p.color },
          type: p.type,
          value: p.value,
        });
        apiRule.gradientRule = {
          minpoint: mapPoint(rule.gradientRule.minpoint),
          maxpoint: mapPoint(rule.gradientRule.maxpoint),
          ...(rule.gradientRule.midpoint ? { midpoint: mapPoint(rule.gradientRule.midpoint) } : {}),
        };
      }

      requests.push({
        addConditionalFormatRule: { rule: apiRule, index: 0 },
      });
      return builder;
    },
    mergeCells(sheetId: number, range: CellRange, mergeType?: MergeType) {
      requests.push({
        mergeCells: {
          range: { sheetId, ...range },
          mergeType: mergeType ?? 'MERGE_ALL',
        },
      });
      return builder;
    },
    unmergeCells(sheetId: number, range: CellRange) {
      requests.push({
        unmergeCells: {
          range: { sheetId, ...range },
        },
      });
      return builder;
    },
    setBasicFilter(sheetId: number, range: CellRange) {
      requests.push({
        setBasicFilter: {
          filter: {
            range: { sheetId, ...range },
          },
        },
      });
      return builder;
    },
    removeBasicFilter(sheetId: number) {
      requests.push({
        clearBasicFilter: { sheetId },
      });
      return builder;
    },
    resizeDimension(sheetId: number, dimension: 'ROWS' | 'COLUMNS', startIndex: number, endIndex: number, pixelSize: number) {
      requests.push({
        updateDimensionProperties: {
          range: {
            sheetId,
            dimension,
            startIndex,
            endIndex,
          },
          properties: { pixelSize },
          fields: 'pixelSize',
        },
      });
      return builder;
    },

    // -- Task #4: Data management implementations --

    setDataValidation(sheetId: number, range: CellRange, rule: DataValidationRule) {
      const condition: sheets_v4.Schema$BooleanCondition = { type: rule.type, values: [] };

      switch (rule.type) {
        case 'ONE_OF_LIST':
          condition.values = (rule.values ?? []).map((v) => ({ userEnteredValue: v }));
          break;
        case 'ONE_OF_RANGE':
        case 'CUSTOM_FORMULA':
          if (rule.formula) condition.values = [{ userEnteredValue: rule.formula }];
          break;
        case 'NUMBER_BETWEEN':
          if (rule.min !== undefined) condition.values!.push({ userEnteredValue: String(rule.min) });
          if (rule.max !== undefined) condition.values!.push({ userEnteredValue: String(rule.max) });
          break;
        case 'NUMBER_GREATER':
        case 'NUMBER_LESS':
          if (rule.min !== undefined) condition.values = [{ userEnteredValue: String(rule.min) }];
          if (rule.max !== undefined) condition.values = [{ userEnteredValue: String(rule.max) }];
          break;
        case 'BOOLEAN':
          break;
      }

      requests.push({
        setDataValidation: {
          range: { sheetId, ...range },
          rule: {
            condition,
            strict: rule.strict ?? true,
            showCustomUi: rule.showCustomUi ?? false,
          },
        },
      });
      return builder;
    },
    removeDataValidation(sheetId: number, range: CellRange) {
      requests.push({
        setDataValidation: {
          range: { sheetId, ...range },
        },
      });
      return builder;
    },
    addPivotTable(_sheetId: number, config: PivotTableConfig) {
      const pivotTable: sheets_v4.Schema$PivotTable = {
        source: {
          sheetId: config.sourceRange.sheetId,
          ...config.sourceRange.range,
        },
        rows: config.rows.map((r) => ({
          sourceColumnOffset: r.sourceColumnOffset,
          sortOrder: r.sortOrder ?? 'ASCENDING',
          showTotals: true,
        })),
        columns: config.columns.map((c) => ({
          sourceColumnOffset: c.sourceColumnOffset,
          sortOrder: c.sortOrder ?? 'ASCENDING',
          showTotals: true,
        })),
        values: config.values.map((v) => ({
          sourceColumnOffset: v.sourceColumnOffset,
          summarizeFunction: v.summarizeFunction,
        })),
      };

      requests.push({
        updateCells: {
          rows: [
            {
              values: [{ pivotTable }],
            },
          ],
          start: {
            sheetId: config.anchorCell.sheetId,
            rowIndex: config.anchorCell.rowIndex,
            columnIndex: config.anchorCell.columnIndex,
          },
          fields: 'pivotTable',
        },
      });
      return builder;
    },
    findReplace(sheetId: number, config: FindReplaceConfig) {
      requests.push({
        findReplace: {
          find: config.find,
          replacement: config.replacement,
          matchCase: config.matchCase ?? false,
          matchEntireCell: config.matchEntireCell ?? false,
          searchByRegex: config.searchByRegex ?? false,
          allSheets: config.allSheets ?? false,
          ...(config.allSheets ? {} : { sheetId }),
        },
      });
      return builder;
    },

    // -- Task #5: Structure and layout implementations --

    freezeDimensions(sheetId: number, frozenRowCount?: number, frozenColumnCount?: number) {
      const gridProperties: Record<string, number> = {};
      const fieldParts: string[] = [];
      if (frozenRowCount !== undefined) {
        gridProperties.frozenRowCount = frozenRowCount;
        fieldParts.push('gridProperties.frozenRowCount');
      }
      if (frozenColumnCount !== undefined) {
        gridProperties.frozenColumnCount = frozenColumnCount;
        fieldParts.push('gridProperties.frozenColumnCount');
      }
      if (fieldParts.length > 0) {
        requests.push({
          updateSheetProperties: {
            properties: { sheetId, gridProperties },
            fields: fieldParts.join(','),
          },
        });
      }
      return builder;
    },
    groupDimension(sheetId: number, dimension: DimensionType, startIndex: number, endIndex: number) {
      requests.push({
        addDimensionGroup: {
          range: { sheetId, dimension, startIndex, endIndex },
        },
      });
      return builder;
    },
    ungroupDimension(sheetId: number, dimension: DimensionType, startIndex: number, endIndex: number) {
      requests.push({
        deleteDimensionGroup: {
          range: { sheetId, dimension, startIndex, endIndex },
        },
      });
      return builder;
    },
    collapseGroup(sheetId: number, dimension: DimensionType, startIndex: number, endIndex: number, collapsed: boolean) {
      requests.push({
        updateDimensionGroup: {
          dimensionGroup: {
            range: { sheetId, dimension, startIndex, endIndex },
            collapsed,
            depth: 1,
          },
          fields: 'collapsed',
        },
      });
      return builder;
    },
    hideSheet(sheetId: number) {
      requests.push({
        updateSheetProperties: {
          properties: { sheetId, hidden: true },
          fields: 'hidden',
        },
      });
      return builder;
    },
    unhideSheet(sheetId: number) {
      requests.push({
        updateSheetProperties: {
          properties: { sheetId, hidden: false },
          fields: 'hidden',
        },
      });
      return builder;
    },
    duplicateSheet(sheetId: number, newName?: string, insertAtIndex?: number) {
      requests.push({
        duplicateSheet: {
          sourceSheetId: sheetId,
          ...(newName ? { newSheetName: newName } : {}),
          ...(insertAtIndex !== undefined ? { insertSheetIndex: insertAtIndex } : {}),
        },
      });
      return builder;
    },
    autoResizeDimension(sheetId: number, dimension: DimensionType, startIndex: number, endIndex: number) {
      requests.push({
        autoResizeDimensions: {
          dimensions: { sheetId, dimension, startIndex, endIndex },
        },
      });
      return builder;
    },
    moveDimension(sheetId: number, dimension: DimensionType, sourceStartIndex: number, sourceEndIndex: number, destinationIndex: number) {
      requests.push({
        moveDimension: {
          source: { sheetId, dimension, startIndex: sourceStartIndex, endIndex: sourceEndIndex },
          destinationIndex,
        },
      });
      return builder;
    },
    setTabColor(sheetId: number, color: RGBColor) {
      requests.push({
        updateSheetProperties: {
          properties: { sheetId, tabColorStyle: { rgbColor: color } },
          fields: 'tabColorStyle',
        },
      });
      return builder;
    },

    // -- Task #6: Collaboration and rich content implementations --

    protectRange(sheetId: number, range: CellRange, description?: string, warningOnly?: boolean) {
      requests.push({
        addProtectedRange: {
          protectedRange: {
            range: { sheetId, ...range },
            description: description ?? '',
            warningOnly: warningOnly ?? false,
          },
        },
      });
      return builder;
    },
    removeProtection(protectedRangeId: number) {
      requests.push({
        deleteProtectedRange: { protectedRangeId },
      });
      return builder;
    },
    setNote(sheetId: number, row: number, col: number, note: string) {
      requests.push({
        updateCells: {
          rows: [{ values: [{ note }] }],
          start: { sheetId, rowIndex: row, columnIndex: col },
          fields: 'note',
        },
      });
      return builder;
    },
    clearNote(sheetId: number, row: number, col: number) {
      requests.push({
        updateCells: {
          rows: [{ values: [{ note: '' }] }],
          start: { sheetId, rowIndex: row, columnIndex: col },
          fields: 'note',
        },
      });
      return builder;
    },
    setNotes(sheetId: number, notes: CellNote[]) {
      for (const n of notes) {
        requests.push({
          updateCells: {
            rows: [{ values: [{ note: n.note }] }],
            start: { sheetId, rowIndex: n.row, columnIndex: n.col },
            fields: 'note',
          },
        });
      }
      return builder;
    },
    setHyperlink(sheetId: number, row: number, col: number, url: string, label?: string) {
      const formula = label
        ? `=HYPERLINK("${url}","${label}")`
        : `=HYPERLINK("${url}")`;
      requests.push({
        updateCells: {
          rows: [{ values: [{ userEnteredValue: { formulaValue: formula } }] }],
          start: { sheetId, rowIndex: row, columnIndex: col },
          fields: 'userEnteredValue.formulaValue',
        },
      });
      return builder;
    },
    insertImage(sheetId: number, row: number, col: number, imageUrl: string) {
      requests.push({
        updateCells: {
          rows: [{ values: [{ userEnteredValue: { formulaValue: `=IMAGE("${imageUrl}")` } }] }],
          start: { sheetId, rowIndex: row, columnIndex: col },
          fields: 'userEnteredValue.formulaValue',
        },
      });
      return builder;
    },
    addDeveloperMetadata(sheetId: number, key: string, value: string, location: MetadataLocationType, locationIndex?: number) {
      const metadataLocation: sheets_v4.Schema$DeveloperMetadataLocation = {};
      if (location === 'SHEET') {
        metadataLocation.sheetId = sheetId;
      } else {
        metadataLocation.dimensionRange = {
          sheetId,
          dimension: location === 'ROW' ? 'ROWS' : 'COLUMNS',
          startIndex: locationIndex ?? 0,
          endIndex: (locationIndex ?? 0) + 1,
        };
      }
      requests.push({
        createDeveloperMetadata: {
          developerMetadata: {
            metadataKey: key,
            metadataValue: value,
            location: metadataLocation,
            visibility: 'PROJECT',
          },
        },
      });
      return builder;
    },
    deleteDeveloperMetadata(metadataId: number) {
      requests.push({
        deleteDeveloperMetadata: {
          dataFilter: { developerMetadataLookup: { metadataId } },
        },
      });
      return builder;
    },

    async send() {
      const result: BatchResult = { writtenCells: 0, clearedRanges: 0, formatOperations: 0 };

      const promises: Promise<void>[] = [];

      if (writes.length > 0) {
        promises.push(
          sheets.spreadsheets.values.batchUpdate({
            spreadsheetId,
            requestBody: {
              valueInputOption: 'USER_ENTERED',
              data: writes.map((w) => ({ range: w.range, values: w.values })),
            },
          }).then((res) => {
            result.writtenCells = res.data.totalUpdatedCells ?? 0;
          }),
        );
      }

      if (clears.length > 0) {
        promises.push(
          sheets.spreadsheets.values.batchClear({
            spreadsheetId,
            requestBody: { ranges: clears },
          }).then(() => {
            result.clearedRanges = clears.length;
          }),
        );
      }

      // Collect all batchUpdate requests: format + border + other operations
      const allRequests: sheets_v4.Schema$Request[] = [...requests];
      for (const op of formats) {
        allRequests.push(buildFormatRequest(op));
        const borderReq = buildBorderRequest(op);
        if (borderReq) allRequests.push(borderReq);
      }

      if (allRequests.length > 0) {
        promises.push(
          sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests: allRequests },
          }).then(() => {
            result.formatOperations = formats.length;
          }),
        );
      }

      await Promise.all(promises);
      return result;
    },
  };

  return builder;
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
