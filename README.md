# @workerbee/google-sheets

Google Sheets tool for kazibee. Read, write, and manage spreadsheets from the sandbox.

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

- `createSpreadsheet(title)` — create a new spreadsheet
- `getSpreadsheet(spreadsheetId)` — get spreadsheet metadata

### Sheets (tabs)

- `listSheets(spreadsheetId)` — list all sheets/tabs
- `addSheet(spreadsheetId, title)` — add a new sheet/tab
- `deleteSheet(spreadsheetId, sheetId)` — delete a sheet/tab

### Data

- `readRange(spreadsheetId, range)` — read cells (returns 2D array)
- `writeRange(spreadsheetId, range, values)` — write cells
- `appendRows(spreadsheetId, range, rows)` — append rows after existing data
- `clearRange(spreadsheetId, range)` — clear cell contents

### Formatting

- `formatCells(spreadsheetId, sheetId, range, format)` — bold, color, alignment, etc.

## Usage

```javascript
// Read data
const rows = await tools["google-sheets"].readRange("SPREADSHEET_ID", "Sheet1!A1:C10");

// Write data
await tools["google-sheets"].writeRange("SPREADSHEET_ID", "Sheet1!A1", [
  ["Name", "Email", "Role"],
  ["Alice", "alice@example.com", "Admin"],
]);

// Append rows
await tools["google-sheets"].appendRows("SPREADSHEET_ID", "Sheet1!A:C", [
  ["Bob", "bob@example.com", "User"],
]);

// Create a new spreadsheet
const sheet = await tools["google-sheets"].createSpreadsheet("Q1 Report");
```
