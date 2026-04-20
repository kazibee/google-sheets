/** Scoped secrets provided by the tool runtime for Google OAuth2 authentication. */
export interface Env {
	/** Google OAuth2 client ID. */
	CLIENT_ID: string;
	/** Google OAuth2 client secret. */
	CLIENT_SECRET: string;
	/** Long-lived refresh token obtained during `kazibee google-sheets login`. */
	REFRESH_TOKEN: string;
}
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
export type NumberFormatType = "TEXT" | "NUMBER" | "PERCENT" | "CURRENCY" | "DATE" | "TIME" | "SCIENTIFIC";
/** Number format specification for cell values. */
export interface NumberFormat {
	/** The type of number format. */
	type: NumberFormatType;
	/** Optional ICU pattern string (e.g. "#,##0.00" for numbers, "yyyy-mm-dd" for dates). */
	pattern?: string;
}
/** Line style for a single border edge. */
export type BorderStyleType = "SOLID" | "DASHED" | "DOTTED" | "SOLID_MEDIUM" | "SOLID_THICK" | "DOUBLE";
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
	horizontalAlignment?: "LEFT" | "CENTER" | "RIGHT";
	/** Number format applied to cell values. */
	numberFormat?: NumberFormat;
	/** Border styles for the cell range. Processed as a separate updateBorders request. */
	borders?: Borders;
}
/** Sort order direction. */
export type SortOrder = "ASCENDING" | "DESCENDING";
/** Specifies a single column sort within a sortRange request. */
export interface SortSpec {
	/** Zero-based column index to sort by. */
	dimensionIndex: number;
	/** Sort direction. */
	sortOrder: SortOrder;
}
/** Supported chart types for addChart. */
export type ChartType = "BAR" | "LINE" | "PIE" | "COLUMN" | "AREA" | "SCATTER";
/** Overlay position for a chart anchored to a cell. */
export interface ChartOverlayPosition {
	/** The cell the chart is anchored to. */
	anchorCell: {
		sheetId: number;
		rowIndex: number;
		columnIndex: number;
	};
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
export type ConditionType = "NUMBER_GREATER" | "NUMBER_GREATER_THAN_EQ" | "NUMBER_LESS" | "NUMBER_LESS_THAN_EQ" | "NUMBER_EQ" | "NUMBER_NOT_EQ" | "NUMBER_BETWEEN" | "NUMBER_NOT_BETWEEN" | "TEXT_CONTAINS" | "TEXT_NOT_CONTAINS" | "TEXT_STARTS_WITH" | "TEXT_ENDS_WITH" | "TEXT_EQ" | "TEXT_IS_EMAIL" | "TEXT_IS_URL" | "DATE_EQ" | "DATE_BEFORE" | "DATE_AFTER" | "DATE_ON_OR_BEFORE" | "DATE_ON_OR_AFTER" | "DATE_BETWEEN" | "DATE_NOT_BETWEEN" | "DATE_IS_VALID" | "ONE_OF_RANGE" | "ONE_OF_LIST" | "BLANK" | "NOT_BLANK" | "CUSTOM_FORMULA" | "BOOLEAN";
/** A boolean condition with comparison values. */
export interface BooleanCondition {
	/** The type of condition to evaluate. */
	type: ConditionType;
	/** Values used by the condition (e.g. the threshold number, the formula string). */
	values?: Array<{
		userEnteredValue?: string;
		relativeDate?: string;
	}>;
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
	type: "MIN" | "MAX" | "NUMBER" | "PERCENT" | "PERCENTILE";
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
export type MergeType = "MERGE_ALL" | "MERGE_COLUMNS" | "MERGE_ROWS";
/** Validation rule type for data validation. */
export type DataValidationType = "ONE_OF_LIST" | "ONE_OF_RANGE" | "NUMBER_BETWEEN" | "NUMBER_GREATER" | "NUMBER_LESS" | "BOOLEAN" | "CUSTOM_FORMULA";
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
export type PivotSummarizeFunction = "SUM" | "COUNT" | "AVERAGE" | "MAX" | "MIN" | "CUSTOM";
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
	sourceRange: {
		sheetId: number;
		range: CellRange;
	};
	/** Row groupings. */
	rows: PivotGroup[];
	/** Column groupings. */
	columns: PivotGroup[];
	/** Aggregated values. */
	values: PivotValue[];
	/** Cell where the pivot table is anchored. */
	anchorCell: {
		sheetId: number;
		rowIndex: number;
		columnIndex: number;
	};
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
/** Configuration for dimension grouping operations. */
export type DimensionType = "ROWS" | "COLUMNS";
/** Location type for developer metadata. */
export type MetadataLocationType = "SHEET" | "ROW" | "COLUMN";
/** A note to set on a specific cell. */
export interface CellNote {
	/** Zero-based row index. */
	row: number;
	/** Zero-based column index. */
	col: number;
	/** The note text. */
	note: string;
}
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
	resizeDimension(sheetId: number, dimension: "ROWS" | "COLUMNS", startIndex: number, endIndex: number, pixelSize: number): BatchBuilder;
	/** Queue a data validation rule on a range. */
	setDataValidation(sheetId: number, range: CellRange, rule: DataValidationRule): BatchBuilder;
	/** Queue removal of data validation from a range. */
	removeDataValidation(sheetId: number, range: CellRange): BatchBuilder;
	/** Queue a pivot table creation. */
	addPivotTable(sheetId: number, config: PivotTableConfig): BatchBuilder;
	/** Queue a find and replace operation. */
	findReplace(sheetId: number, config: FindReplaceConfig): BatchBuilder;
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
 * Tool entry point. Receives scoped secrets from the runtime, creates an
 * authenticated Google Sheets client, and returns the full API surface.
 */
declare function main(env: Env): {
	createSpreadsheet: (title: string) => Promise<SpreadsheetInfo>;
	getSpreadsheet: (spreadsheetId: string) => Promise<SpreadsheetInfo>;
	listSheets: (spreadsheetId: string) => Promise<SheetInfo[]>;
	addSheet: (spreadsheetId: string, title: string) => Promise<SheetInfo>;
	deleteSheet: (spreadsheetId: string, sheetId: number) => Promise<void>;
	readRange: (spreadsheetId: string, range: string) => Promise<string[][]>;
	readRanges: (spreadsheetId: string, ranges: string[]) => Promise<RangeValues[]>;
	appendRows: (spreadsheetId: string, range: string, rows: string[][]) => Promise<{
		updatedCells: number;
	}>;
	addNamedRange: (spreadsheetId: string, name: string, sheetId: number, range: CellRange) => Promise<NamedRange>;
	deleteNamedRange: (spreadsheetId: string, namedRangeId: string) => Promise<void>;
	listNamedRanges: (spreadsheetId: string) => Promise<NamedRange[]>;
	batch: (spreadsheetId: string) => Promise<BatchBuilder>;
};

export {
	main as default,
};

export {};
