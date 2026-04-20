import { createAuthClient, type Env } from './auth';
import { createSheetsClient } from './sheets-client';

export type { Env } from './auth';
export type {
  SpreadsheetInfo,
  SheetInfo,
  CellRange,
  RGBColor,
  NumberFormat,
  NumberFormatType,
  BorderStyle,
  BorderStyleType,
  Borders,
  CellFormat,
  SortOrder,
  SortSpec,
  ChartType,
  ChartOverlayPosition,
  ChartSourceRange,
  ChartConfig,
  ConditionType,
  BooleanCondition,
  BooleanRule,
  InterpolationPoint,
  GradientRule,
  ConditionalFormatRule,
  MergeType,
  // Task #4: Data management types
  DataValidationType,
  DataValidationRule,
  PivotSummarizeFunction,
  PivotGroup,
  PivotValue,
  PivotTableConfig,
  FindReplaceConfig,
  NamedRange,
  RangeValues,
  // Task #5: Structure and layout types
  DimensionType,
  // Task #6: Collaboration and rich content types
  MetadataLocationType,
  CellNote,
  FormatOperation,
  BatchResult,
  BatchBuilder,
} from './sheets-client';

/**
 * Tool entry point. Receives scoped secrets from the runtime, creates an
 * authenticated Google Sheets client, and returns the full API surface.
 */
export default function main(env: Env) {
  const auth = createAuthClient(env);
  return createSheetsClient(auth);
}
