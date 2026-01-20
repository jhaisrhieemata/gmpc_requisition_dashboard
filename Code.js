// ===============================
// CONFIG
// ===============================
const SHEET_FILE_ID = "1MVY1ucbqCTRQkoEEMaQc6tEI6u62psbup6iL023xGsI";
// ===============================
// WEB APP ENTRY
// ===============================
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Requisitions / Stocks Dashboard")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ===============================
// HELPERS
// ===============================
function detectRequestType(sheetName) {
  return sheetName.toLowerCase().includes("special") ? "SPECIAL" : "OFFICE";
}

function findIdx(headers, alts) {
  for (let i = 0; i < headers.length; i++)
    for (let a of alts) if (headers[i].includes(a)) return i;
  return -1;
}

// ===============================
// DASHBOARD DATA
// ===============================
function getDashboardData() {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const sheets = ss.getSheets();

  const headerNames = {
    itemId: ["item_id", "item id", "item"],
    description: ["description", "desc"],
    qty: ["qty", "quantity"],
    unit: ["unit", "uom"],
    branch: ["branch", "office", "location"],
    email: ["email", "e-mail", "email address", "contact email", "contact"],
    status: ["status", "approval"],
    pdfUrl: ["pdf", "pdf url", "pdf_url", "pdf link", "pdf link", "drive link"],
  };

  const pendingOffice = [];
  const pendingSpecial = [];
  const statusCounts = {};
  const addStocks = [];

  // ===== INVENTORY =====
  const stockSheet = ss.getSheetByName("ADD STOCKS");
  if (stockSheet) {
    const v = stockSheet.getDataRange().getValues();
    const h = v[0].map((x) => ("" + x).toLowerCase());
    const idx = {
      itemId: h.indexOf("item_id"),
      description: h.indexOf("description"),
      unit: h.indexOf("unit"),
      total: h.indexOf("total running stocks"),
      status: h.indexOf("status"),
    };

    for (let i = 1; i < v.length; i++) {
      addStocks.push({
        itemId: v[i][idx.itemId],
        description: v[i][idx.description],
        unit: v[i][idx.unit],
        totalRunningStocks: Number(v[i][idx.total] || 0),
        status: v[i][idx.status],
      });
    }
  }

  // ===== REQUESTS =====
  sheets.forEach((sheet) => {
    if (sheet.getName() === "ADD STOCKS") return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map((h) => ("" + h).toLowerCase());
    const idx = {
      itemId: findIdx(headers, headerNames.itemId),
      description: findIdx(headers, headerNames.description),
      qty: findIdx(headers, headerNames.qty),
      unit: findIdx(headers, headerNames.unit),
      branch: findIdx(headers, headerNames.branch),
      email: findIdx(headers, headerNames.email),
      status: findIdx(headers, headerNames.status),
    };

    for (let r = 1; r < data.length; r++) {
      const status = (data[r][idx.status] || "").toString();
      if (!status) continue;

      statusCounts[status] = (statusCounts[status] || 0) + 1;

      if (status.toLowerCase() === "pending") {
        const row = {
          sheetName: sheet.getName(),
          rowNumber: r + 1,
          itemId: data[r][idx.itemId] || "",
          description: data[r][idx.description] || "",
          qty: Number(data[r][idx.qty] || 0),
          unit: data[r][idx.unit] || "",
          branch: data[r][idx.branch] || "",
          email: data[r][idx.email] || "",
          status,
        };

        detectRequestType(sheet.getName()) === "SPECIAL"
          ? pendingSpecial.push(row)
          : pendingOffice.push(row);
      }
    }
  });

  return { pendingOffice, pendingSpecial, statusCounts, addStocks };
}

// ===============================
// GROUP BY BRANCH
// ===============================
function getPendingGroupedByBranch() {
  const d = getDashboardData();

  function group(rows) {
    const m = {};
    rows.forEach((r) => {
      if (!m[r.branch]) m[r.branch] = [];
      m[r.branch].push(r);
    });
    return Object.keys(m).map((b) => ({
      branch: b,
      total: m[b].length,
      rows: m[b],
    }));
  }

  return {
    office: group(d.pendingOffice),
    special: group(d.pendingSpecial),
  };
}

// ===============================
// ACTIONS
// ===============================
function approvePendingRow(sheetName, rowNumber) {
  const s = SpreadsheetApp.openById(SHEET_FILE_ID).getSheetByName(sheetName);
  const h = s
    .getRange(1, 1, 1, s.getLastColumn())
    .getValues()[0]
    .map((x) => ("" + x).toLowerCase());
  const idx = h.findIndex((x) => x.includes("status"));
  s.getRange(rowNumber, idx + 1).setValue("Approved");
  return true;
}

function deletePendingRow(sheetName, rowNumber) {
  SpreadsheetApp.openById(SHEET_FILE_ID)
    .getSheetByName(sheetName)
    .deleteRow(rowNumber);
  return true;
}

function editPendingRow(sheetName, rowNumber, vals) {
  const s = SpreadsheetApp.openById(SHEET_FILE_ID).getSheetByName(sheetName);
  const h = s
    .getRange(1, 1, 1, s.getLastColumn())
    .getValues()[0]
    .map((x) => ("" + x).toLowerCase());
  Object.keys(vals).forEach((k) => {
    const i = h.findIndex((x) => x.includes(k.toLowerCase()));
    if (i > -1) s.getRange(rowNumber, i + 1).setValue(vals[k]);
  });
  return true;
}

// ===============================
// BATCH PDF PRINT
// ===============================
function getBatchPdfUrls(rows) {
  // Fetch one PDF URL per branch. Prefer SPECIAL over OFFICE when both exist.
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const selectedPerBranch = {};

  const headerNames = {
    pdfUrl: ["pdf", "pdf url", "pdf_url", "pdf link", "drive link"],
    branch: ["branch", "office", "location"],
  };

  rows.forEach((r) => {
    try {
      const sh = ss.getSheetByName(r.sheetName);
      if (!sh) return;

      const headers = sh
        .getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map((x) => ("" + x).toLowerCase());

      const pdfUrlIdx = findIdx(headers, headerNames.pdfUrl);
      const branchIdx = findIdx(headers, headerNames.branch);

      const data = sh
        .getRange(r.rowNumber, 1, 1, sh.getLastColumn())
        .getValues()[0];

      const pdfUrl =
        pdfUrlIdx >= 0 ? (data[pdfUrlIdx] || "").toString().trim() : "";
      const branch =
        branchIdx >= 0
          ? (data[branchIdx] || "").toString().trim()
          : r.branch || "Unknown";

      if (!pdfUrl) return;

      const type = detectRequestType(r.sheetName); // 'SPECIAL' or 'OFFICE'

      // If we haven't selected a pdf for this branch yet, pick this one.
      if (!selectedPerBranch[branch]) {
        selectedPerBranch[branch] = {
          url: pdfUrl,
          type: type,
          itemId: r.itemId,
          description: r.description,
        };
        return;
      }

      // If we already have an OFFICE and this one is SPECIAL, prefer SPECIAL.
      const existing = selectedPerBranch[branch];
      if (existing.type !== "SPECIAL" && type === "SPECIAL") {
        selectedPerBranch[branch] = {
          url: pdfUrl,
          type: type,
          itemId: r.itemId,
          description: r.description,
        };
      }
    } catch (e) {
      Logger.log(
        "PDF fetch failed for %s#%s: %s",
        r.sheetName,
        r.rowNumber,
        e.toString()
      );
    }
  });

  return selectedPerBranch;
}

function _createPdfFileForRow(row, ssOpt) {
  const ss = ssOpt || SpreadsheetApp.openById(SHEET_FILE_ID);
  const sh = ss.getSheetByName(row.sheetName);
  if (!sh) return null;
  const headers = sh
    .getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map((x) => "" + x);
  const data = sh
    .getRange(row.rowNumber, 1, 1, sh.getLastColumn())
    .getValues()[0];

  let html =
    '<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;padding:18px}table{border-collapse:collapse}td{padding:6px;vertical-align:top}</style></head><body>';
  html += `<h2>Request — ${row.sheetName} #${row.rowNumber}</h2><table>`;
  for (let i = 0; i < headers.length; i++) {
    html += `<tr><td><strong>${headers[i] || ""}</strong></td><td>${
      data[i] || ""
    }</td></tr>`;
  }
  html += "</table></body></html>";

  const blob = Utilities.newBlob(html, "text/html")
    .getAs("application/pdf")
    .setName(`request-${row.sheetName}-${row.rowNumber}.pdf`);
  const file = DriveApp.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  return { file, url: file.getUrl() };
}

// ===============================
// BATCH
// ===============================
function batchApprove(rows) {
  if (!rows || !Array.isArray(rows)) {
    Logger.log("batchApprove: Invalid rows parameter");
    return false;
  }
  rows.forEach((r) => approvePendingRow(r.sheetName, r.rowNumber));
  return true;
}

function getPendingBreakdown() {
  const d = getDashboardData();
  return { office: d.pendingOffice.length, special: d.pendingSpecial.length };
}
