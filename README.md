# @kazibee/google-sheets

Google Sheets tool for kazibee. Read, write, format, chart, validate, and manage spreadsheets from the sandbox.

## Install

```bash
kazibee install google-sheets github:kazibee/google-sheets
```

Install globally with `-g`:

```bash
kazibee install -g google-sheets github:kazibee/google-sheets
```

Or pin to a specific commit:

```bash
kazibee install google-sheets github:kazibee/google-sheets#COMMIT_SHA
```

## Login

```bash
kazibee google-sheets login
```

Opens your browser to authorize with Google. Credentials are stored automatically.

## API

### Spreadsheets

- `createSpreadsheet(title)` -- create a new spreadsheet
- `getSpreadsheet(spreadsheetId)` -- get spreadsheet metadata

### Sheets (tabs)

- `listSheets(spreadsheetId)` -- list all sheets/tabs
- `addSheet(spreadsheetId, title)` -- add a new sheet/tab
- `deleteSheet(spreadsheetId, sheetId)` -- delete a sheet/tab

### Data (reads)

- `readRange(spreadsheetId, range)` -- read cells (returns 2D array)
- `readRanges(spreadsheetId, ranges)` -- read multiple ranges in one call
- `appendRows(spreadsheetId, range, rows)` -- append rows after existing data

### Named Ranges

- `addNamedRange(spreadsheetId, name, sheetId, range)` -- create a named range
- `deleteNamedRange(spreadsheetId, namedRangeId)` -- delete a named range
- `listNamedRanges(spreadsheetId)` -- list all named ranges

### Batch Builder (writes, clears, formatting, and all mutations)

All write, clear, format, and structural operations go through the batch builder. This prevents hitting Google's write-rate quota by combining operations into minimal API calls.

- `batch(spreadsheetId)` -- start a batch builder, then chain:

**Core operations:**
  - `.writeRange(range, values)` -- queue a cell write
  - `.clearRange(range)` -- queue a range clear
  - `.formatCells(sheetId, range, format)` -- queue formatting (bold, colors, borders, number formats)

**Sorting and filtering:**
  - `.sortRange(sheetId, range, sortSpecs)` -- sort a range by one or more columns
  - `.setBasicFilter(sheetId, range)` -- add a filter view
  - `.removeBasicFilter(sheetId)` -- remove the filter

**Charts:**
  - `.addChart(sheetId, chartConfig)` -- create an embedded chart (BAR, LINE, PIE, COLUMN, AREA, SCATTER)

**Conditional formatting:**
  - `.addConditionalFormat(sheetId, rule)` -- add boolean or gradient rules

**Merging:**
  - `.mergeCells(sheetId, range, mergeType?)` -- merge cells (MERGE_ALL, MERGE_COLUMNS, MERGE_ROWS)
  - `.unmergeCells(sheetId, range)` -- unmerge cells

**Data management:**
  - `.setDataValidation(sheetId, range, rule)` -- add dropdown, number, or custom validation
  - `.removeDataValidation(sheetId, range)` -- remove validation
  - `.addPivotTable(sheetId, config)` -- create a pivot table
  - `.findReplace(sheetId, config)` -- find and replace text

**Structure and layout:**
  - `.resizeDimension(sheetId, dim, start, end, px)` -- set row/column size
  - `.autoResizeDimension(sheetId, dim, start, end)` -- auto-fit to content
  - `.freezeDimensions(sheetId, rows?, cols?)` -- freeze rows/columns
  - `.groupDimension(sheetId, dim, start, end)` -- create a row/column group
  - `.ungroupDimension(sheetId, dim, start, end)` -- remove a group
  - `.collapseGroup(sheetId, dim, start, end, collapsed)` -- collapse or expand
  - `.hideSheet(sheetId)` / `.unhideSheet(sheetId)` -- toggle sheet visibility
  - `.duplicateSheet(sheetId, newName?, index?)` -- copy a sheet
  - `.moveDimension(sheetId, dim, srcStart, srcEnd, dest)` -- move rows/columns
  - `.setTabColor(sheetId, color)` -- set a sheet tab color

**Collaboration and rich content:**
  - `.protectRange(sheetId, range, description?, warningOnly?)` -- protect a range
  - `.removeProtection(protectedRangeId)` -- remove protection
  - `.setNote(sheetId, row, col, note)` -- set a cell note
  - `.clearNote(sheetId, row, col)` -- clear a cell note
  - `.setNotes(sheetId, notes)` -- set notes on multiple cells
  - `.setHyperlink(sheetId, row, col, url, label?)` -- add a hyperlink
  - `.insertImage(sheetId, row, col, imageUrl)` -- insert an image via formula
  - `.addDeveloperMetadata(sheetId, key, value, location, index?)` -- add metadata
  - `.deleteDeveloperMetadata(metadataId)` -- remove metadata

**Execute:**
  - `.send()` -- execute all queued operations in one batch

## Usage

```javascript
// Read data (direct -- reads are not rate-limited)
const rows = await tools["google-sheets"].readRange("SPREADSHEET_ID", "Sheet1!A1:C10");

// Read multiple ranges at once
const results = await tools["google-sheets"].readRanges("SPREADSHEET_ID", [
  "Sheet1!A1:C10",
  "Sheet2!A1:B5"
]);

// Append rows (direct -- single atomic append)
await tools["google-sheets"].appendRows("SPREADSHEET_ID", "Sheet1!A:C", [
  ["Bob", "bob@example.com", "User"],
]);

// Write + format in one batch (all mutations go here)
const result = await tools["google-sheets"].batch("SPREADSHEET_ID")
  .writeRange("Sheet1!A1:C1", [["Name", "Email", "Role"]])
  .writeRange("Sheet1!A2:C3", [
    ["Alice", "alice@example.com", "Admin"],
    ["Bob", "bob@example.com", "User"],
  ])
  .formatCells(0, { startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 3 },
    { bold: true, horizontalAlignment: "CENTER" })
  .clearRange("Sheet1!D1:Z1000")
  .send();

// Sort, chart, and conditional formatting in one batch
await tools["google-sheets"].batch("SPREADSHEET_ID")
  .sortRange(0,
    { startRowIndex: 1, endRowIndex: 100, startColumnIndex: 0, endColumnIndex: 3 },
    [{ dimensionIndex: 0, sortOrder: "ASCENDING" }]
  )
  .addChart(0, {
    chartType: "BAR",
    title: "Sales",
    sourceRanges: [{ sheetId: 0, range: { startRowIndex: 0, endRowIndex: 10, startColumnIndex: 0, endColumnIndex: 2 } }],
    position: { anchorCell: { sheetId: 0, rowIndex: 0, columnIndex: 4 } }
  })
  .addConditionalFormat(0, {
    ranges: [{ startRowIndex: 1, endRowIndex: 50, startColumnIndex: 2, endColumnIndex: 3 }],
    booleanRule: {
      condition: { type: "NUMBER_GREATER", values: [{ userEnteredValue: "100" }] },
      format: { backgroundColor: { red: 1, green: 0.8, blue: 0.8 } }
    }
  })
  .send();

// Data validation dropdown
await tools["google-sheets"].batch("SPREADSHEET_ID")
  .setDataValidation(0,
    { startRowIndex: 1, endRowIndex: 100, startColumnIndex: 1, endColumnIndex: 2 },
    { type: "ONE_OF_LIST", values: ["Active", "Inactive", "Pending"], showCustomUi: true }
  )
  .send();

// Freeze header row and protect it
await tools["google-sheets"].batch("SPREADSHEET_ID")
  .freezeDimensions(0, 1)
  .protectRange(0,
    { startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 10 },
    "Header row", true
  )
  .send();

// Create a new spreadsheet
const sheet = await tools["google-sheets"].createSpreadsheet("Q1 Report");
```
