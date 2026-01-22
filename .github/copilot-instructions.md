# Copilot Instructions for this Repo

**Purpose**: Help AI coding agents become productive quickly with this Google Apps Script web app that manages inventory and requisitions across branches.

## Big Picture

- **Architecture**: Single-file server ([Code.js](../Code.js)) + single-page client ([Index.html](../Index.html))
- **Data source**: Google Sheet (ID: `SHEET_FILE_ID` in Code.js line 4)
- **Communication**: Client calls server via `google.script.run.withSuccessHandler()` (10s auto-refresh loop)
- **Deployment**: Apps Script editor or clasp (no build step); manifest in [appsscript.json](../appsscript.json)

## Data Model

- **Sheets**: One per requisition form (e.g., "Office Supplies", "Special Request Items") + special "ADD STOCKS" inventory sheet
- **Row 1 Headers**: Detected via `headerNames` object + `findIdx()` case-insensitive matching (NOT column positions)
- **Pending Rows**: Classified by sheet name (contains "special" → SPECIAL, else OFFICE); status must be "Pending"
- **Row Object**: `{ sheetName, rowNumber, itemId, description, qty, unit, branch, email, status }`

## Key Server Functions (start here)

- `getDashboardData()` → `{ pendingOffice, pendingSpecial, statusCounts, addStocks }` (full read of all sheets)
- `getPendingGroupedByBranch()` → groups pending by branch for UI display
- `getPendingBreakdown()` → count summary for badge updates
- `approvePendingRow(sheetName, rowNumber)`, `deletePendingRow()`, `editPendingRow(sheetName, rowNumber, vals)` — mutate rows via header-index lookup (do NOT hard-code column numbers)
- `getBatchPdfUrls(rows)` and `_createPdfFileForRow(row)` — render rows as HTML → PDF blob → Drive file with ANYONE_WITH_LINK share
- `batchApprove(rows)` — approve multiple rows at once

## Patterns & Conventions (do not change lightly)

**Header Detection**

```javascript
const headerNames = {
  status: ["status", "approval"],
  qty: ["qty", "quantity"],
};
const headers = data[0].map((h) => ("" + h).toLowerCase());
const statusIdx = findIdx(headers, headerNames.status); // -1 if not found
const qtyIdx = findIdx(headers, headerNames.qty);
```

- Headers are matched case-insensitively via `includes()` against `headerNames` aliases
- Always validate result (`>= 0`) before using as column index
- `findIdx()` scans **all aliases** in order and returns first match

**Inventory Sheet**

- Sheet named exactly `"ADD STOCKS"` (case-sensitive) is parsed separately at Code.js:52–70
- Columns: `item_id`, `description`, `unit`, `total running stocks`, `status`
- Data rendered in dashboard pie chart + inventory table; no action buttons (display-only)

**Request Type & PDF Preference**

- `detectRequestType(sheetName)` → "SPECIAL" if name contains "special" (case-insensitive), else "OFFICE"
- `getBatchPdfUrls()` selects **one PDF per branch**: if both OFFICE and SPECIAL exist, **prefers SPECIAL**

**Rows with empty status are ignored**; only `status.toLowerCase() === "pending"` triggers display in UI

**Client Modal Workflow** (Index.html:366–410)

1. User clicks "View" on branch → `openModal(branchGroup)` → stores rows in `currentRows`
2. Batch actions (Approve, Print PDF, Email) operate on `currentRows`
3. After action, client calls `refreshAll()` to update UI

**Branch Email Lookup** (Index.html:254–291)

- Hardcoded `branchEmailMap` in client for email composition
- `getEmailForBranch(branch)` — case-insensitive fallback to "ADMIN" if branch not found
- Add new branches to this map for email integration

## Integrations & Permissions

- **Sheets API** (v4) + **Gmail API** (v1): Declared in [appsscript.json](../appsscript.json) `enabledAdvancedServices`
- **DriveApp** + **Utilities**: PDF creation/sharing. `file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)`
- **Webapp Config**: `executeAs: USER_DEPLOYING`, `access: ANYONE_ANONYMOUS`, `exceptionLogging: STACKDRIVER`
- **Optional Gmail**: If adding email notifications, extend manifest scopes and look for "email" header via `findIdx()` + `GmailApp.sendEmail(to, subject, body)`

## Developer Workflows

**Edit & Debug**

- Use [Google Apps Script Editor](https://script.google.com) or `clasp` if configured
- No build step; changes deploy immediately to web app URL
- **Logging**: `Logger.log()` appears in Execution Logs; errors in Cloud Logging (Stackdriver)
- **Data Source**: Update `SHEET_FILE_ID` (Code.js:4) and redeploy to switch sheets

**Adding a New Sheet**

1. Create sheet in target Google Sheet with row 1 headers
2. If sheet name contains "special" (case-insensitive) → classified as SPECIAL; else OFFICE
3. Headers automatically detected via `headerNames.*` aliases (no column number hard-coding needed)

**Adding a Header Alias**

- Extend `headerNames` object (Code.js:37–45, line 189–191) with new aliases
- Example: `status: ["status", "approval", "approval_status"]`
- `findIdx()` will scan all aliases case-insensitively; first match wins

**Validation**

- No automated tests. Deploy to web app and exercise UI to validate changes.
- Verify Sheet access permissions (shared with script execution user)

## Examples & Common Tasks

**Header Lookup (Core Pattern)**

```javascript
const idx = findIdx(headers, headerNames.qty);  // ✓ Use this
if (idx >= 0) s.getRange(rowNumber, idx + 1).setValue(...);  // idx+1 for GAS 1-indexing
```

**Server Function Call from Client**

```javascript
google.script.run.withSuccessHandler(renderGrouped).getPendingGroupedByBranch();
```

**Approve a Row**

```javascript
approvePendingRow("Office Supplies", 5);
```

**Fetch PDF URLs**

```javascript
getBatchPdfUrls(rows); // Returns { branch → { url, type, itemId, description } }
```

**Detect Request Type**

```javascript
detectRequestType("Special Request Items"); // → "SPECIAL"
detectRequestType("Office Supplies"); // → "OFFICE"
```

## Where to Look for Related Code

- **Code.js** — core logic, `headerNames` definitions, PDF generation, sheet mutations
- **Index.html** — client dashboard UI, `google.script.run` calls, branch email map, modal workflow
- **appsscript.json** — webapp access config, advanced service declarations (Sheets, Gmail)
