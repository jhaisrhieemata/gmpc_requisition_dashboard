# Copilot Instructions for this Repo

Purpose: help AI coding agents become productive quickly with this Google Apps Script web app.

- Big picture
  - Single-file server (Code.js) + single-page client (Index.html). Data is a Google Sheet referenced by SHEET_FILE_ID in Code.js. Manifest: appsscript.json.
  - UI → server via google.script.run (see Index.html: refreshAll(), action handlers).

- Key server functions (start here)
  - getDashboardData() — reads all sheets, infers header indexes using headerNames + findIdx(headers, alts), returns { pendingOffice, pendingSpecial, statusCounts, addStocks }.
  - getPendingGroupedByBranch(), getPendingBreakdown() — lightweight transforms of dashboard data.
  - approvePendingRow(sheetName, rowNumber), deletePendingRow(...), editPendingRow(...) — mutate sheet rows using header-index lookup (do NOT hard-code column numbers).
  - getBatchPdfUrls(rows) and _createPdfFileForRow(row) — create PDFs (HTML → blob → Drive) and return file URLs; Drive sharing is attempted with ANYONE_WITH_LINK.
  - batchApprove(rows) — convenience to approve many rows.

- Patterns & conventions (do not change lightly)
  - Header detection: use findIdx(headers, altList) and the headerNames object. Headers are expected on row 1 and matched case-insensitively via includes().
  - Inventory sheet: sheet named exactly "ADD STOCKS" is parsed as inventory (different columns).
  - Request type inference: detectRequestType(sheetName) → "SPECIAL" if sheet name contains "special" (case-insensitive), otherwise "OFFICE". PDF selection prefers SPECIAL over OFFICE per branch.
  - Rows with empty status are ignored; pending rows are status.toLowerCase() === "pending".

- Integrations & permissions
  - Uses DriveApp and Utilities for PDF creation/sharing. appsscript.json sets webapp.executeAs USER_DEPLOYING and webapp.access ANYONE_ANONYMOUS; exceptionLogging = STACKDRIVER.
  - Optional Gmail notifications: if sendApprovalEmail or GmailApp is added, ensure manifest scopes permit Gmail and document where email header detection occurs. If present, look for an "email" header via findIdx and use GmailApp.sendEmail(to, subject, body).

- Developer workflows
  - Edit & debug in the Apps Script online editor or use clasp if configured. Use Logger.log() and Cloud Logging (Stackdriver) for runtime inspection.
  - To change data source, update SHEET_FILE_ID in Code.js.
  - No build step or automated tests in repo — validate changes by deploying and exercising UI.

- Examples / idioms to copy
  - Header lookup: const i = findIdx(headers, headerNames.qty);
  - Client call: google.script.run.withSuccessHandler(renderGrouped).getPendingGroupedByBranch();
  - Approve: approvePendingRow("Form Name", 5);

- Where to look for related code
  - Code.js — core logic, header maps, PDF creation, actions.
  - Index.html — client UI and google.script.run calls.
  - appsscript.json — webapp and API/permission declarations.

If anything here is unclear or you want expansion (e.g., add example clasp commands, document specific headers used by a particular sheet, or include recommended manifest scopes when enabling Gmail), tell me which area to expand.
