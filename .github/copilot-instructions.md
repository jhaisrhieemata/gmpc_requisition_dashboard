# Copilot Instructions for this Repo

Purpose: Help AI coding agents be productive quickly in this Google Apps Script project.

- **Big picture**: This is a Google Apps Script web app. Server-side code lives in [Code.js](Code.js) (GS/JavaScript) and the client is a single-page UI in [Index.html](Index.html). Data is stored in a Google Spreadsheet referenced by the `SHEET_FILE_ID` constant in `Code.js`. The project manifest is [appsscript.json](appsscript.json).

- **Primary flows**:
  - UI calls server functions with `google.script.run` (see `Index.html` — `refreshAll()` and action handlers).
  - Server functions read spreadsheet sheets, infer headers (flexible header matching) and return JSON shapes consumed by the UI. Key server functions: `getDashboardData()`, `getPendingGroupedByBranch()`, `getPendingBreakdown()`, `approvePendingRow()`, `deletePendingRow()`, `editPendingRow()`, `getBatchPdfUrls()`.

- **Important patterns & conventions (do not change lightly)**:
  - Header detection is flexible: `findIdx(headers, alts)` + `headerNames` mapping in `getDashboardData()` — code expects header row at row 1 and uses case-insensitive includes. Follow the same approach when adding new columns.
  - The sheet named `ADD STOCKS` is treated as an inventory sheet (special parsing path). Other sheets are treated as request forms.
  - Request type is inferred from sheet name: `detectRequestType(sheetName)` returns `SPECIAL` if the sheet name contains `special` (case-insensitive); otherwise `OFFICE`. Respect this when changing sheet naming or logic.
  - Status is required to consider a row: server code skips rows with empty status. Pending rows are those whose `status.toLowerCase() === 'pending'`.
  - PDF selection: `getBatchPdfUrls(rows)` prefers one PDF per branch and prefers `SPECIAL` over `OFFICE` when both exist for a branch.

- **Integration points & permissions**:
  - The app uses advanced services defined in `appsscript.json` (Sheets and Gmail). It also uses `DriveApp` and `Utilities` for PDF creation and sharing.
  - Manifest config: `webapp.executeAs` is `USER_DEPLOYING` and `webapp.access` is `ANYONE_ANONYMOUS` — deployments run as the deployer and allow public access. `exceptionLogging` is `STACKDRIVER` (Cloud Logging).

- **Developer workflows (discoverable from repo)**:
  - To run/debug: open the project in the Apps Script editor (or use `clasp` if you have it configured). Use `Logger.log()` output and Cloud Logging to inspect runtime logs.
  - To change the data source, update `SHEET_FILE_ID` in `Code.js`.
  - There are no automated tests or build steps in repo files; changes are validated by deploying and exercising the UI in a browser.

- **Quick code examples (preserve style and APIs)**:
  - Client → server call: `google.script.run.withSuccessHandler(renderGrouped).getPendingGroupedByBranch();` (see `Index.html`).
  - Update a status cell: `approvePendingRow(sheetName, rowNumber)` finds the `status` header and sets `Approved` (see `Code.js`).

- **When you edit server code**:
  - Keep Spreadsheet ranges and header-index logic consistent — many functions rely on header index lookup instead of hard-coded column numbers.
  - Keep `_createPdfFileForRow()` behavior: HTML → blob → Drive file, then attempt to set sharing to `ANYONE_WITH_LINK`. Avoid removing Drive sharing unless you also update the client PDF flow.

- **What the AI should not assume**:
  - Do not assume specific column positions — use header name matching. Do not assume authentication model beyond what's in `appsscript.json`.

If anything is missing or you want me to expand examples (e.g., include sample `clasp` commands or annotate specific functions), tell me which area to expand.
