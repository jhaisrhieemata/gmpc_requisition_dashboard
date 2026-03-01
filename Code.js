// ===============================
// CONFIG
// ===============================
const SHEET_FILE_ID = "1MVY1ucbqCTRQkoEEMaQc6tEI6u62psbup6iL023xGsI";

// ===============================
// WEB APP ENTRY
// ===============================
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("GMPC Requisitions Dashboard")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ===============================
// ROUTER
// ===============================
function doPost(e) {
  const action = e.parameter.action;
  const data = JSON.parse(e.parameter.data || "{}");
  
  let result = { success: false, message: "" };
  
  try {
    switch(action) {
      // Suppliers
      case "getSuppliers": result = getSuppliers(); break;
      case "addSupplier": result = addSupplier(data); break;
      case "updateSupplier": result = updateSupplier(data); break;
      case "deleteSupplier": result = deleteSupplier(data.id); break;
      
      // Supplier Items
      case "getSupplierItems": result = getSupplierItems(); break;
      case "addSupplierItem": result = addSupplierItem(data); break;
      case "updateSupplierItem": result = updateSupplierItem(data); break;
      case "deleteSupplierItem": result = deleteSupplierItem(data.id); break;
      
      // Users
      case "getUsers": result = getUsers(); break;
      case "addUser": result = addUser(data); break;
      case "updateUser": result = updateUser(data); break;
      case "deleteUser": result = deleteUser(data.id); break;
      case "resetPassword": result = resetPassword(data.id); break;
      case "getPasswordResetRequests": result = getPasswordResetRequests(); break;
      case "handlePasswordReset": result = handlePasswordReset(data); break;
      
      // Branches
      case "getBranches": result = getBranches(); break;
      case "addBranch": result = addBranch(data); break;
      case "updateBranch": result = updateBranch(data); break;
      case "deleteBranch": result = deleteBranch(data.id); break;
      
      // Inventory
      case "getInventory": result = getInventory(); break;
      case "updateInventoryItem": result = updateInventoryItem(data); break;
      case "addInventoryItem": result = addInventoryItem(data); break;
      case "deleteInventoryItem": result = deleteInventoryItem(data.id); break;
      
      // Office Requests
      case "getOfficeRequests": result = getOfficeRequests(); break;
      case "updateOfficeRequest": result = updateOfficeRequest(data); break;
      case "deleteOfficeRequest": result = deleteOfficeRequest(data.id); break;
      
      // Special Requests
      case "getSpecialRequests": result = getSpecialRequests(); break;
      case "updateSpecialRequest": result = updateSpecialRequest(data); break;
      case "deleteSpecialRequest": result = deleteSpecialRequest(data.id); break;
      
      // Reports
      case "generateReport": result = generateReport(data); break;
      case "getLowStocksReport": result = getLowStocksReport(); break;
      case "getBranchDashboard": result = getBranchDashboard(); break;
      case "getBranchReport": result = getBranchReport(data); break;
      
      // Dashboard
      case "getDashboardData": result = getDashboardData(); break;
      case "getPendingGroupedByBranch": result = getPendingGroupedByBranch(); break;
      case "getPendingBreakdown": result = getPendingBreakdown(); break;
      
      // Batch Actions
      case "batchAction": result = batchAction(data.rows, data.action); break;
      case "rowAction": result = rowAction(data.sheetName, data.rowNumber, data.action); break;
      case "editPendingRow": result = editPendingRow(data.sheetName, data.rowNumber, data.values); break;
      case "deletePendingRow": result = deletePendingRow(data.sheetName, data.rowNumber); break;
      case "getBatchPdfUrls": result = getBatchPdfUrls(data.rows); break;
      
      default:
        result = { success: false, message: "Unknown action: " + action };
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
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

function getSheetByNameSafe(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  return ss.getSheetByName(sheetName);
}

function generateId() {
  return Utilities.getUuid();
}

function getCurrentTimestamp() {
  return new Date().toISOString();
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
    unit: ["unit"],
    branch: ["branch", "office", "location"],
    email: ["email", "e-mail", "email address", "contact email", "contact"],
    status: ["status", "approval"],
    pdfUrl: ["pdf", "pdf url", "pdf_url", "pdf link", "pdf link", "drive link"],
    uprice: ["uprice", "unit price", "price"],
    amount: ["amount", "total amount"],
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
      supplier: h.indexOf("supplier"),
      classification: h.indexOf("classification"),
    };

    for (let i = 1; i < v.length; i++) {
      addStocks.push({
        id: i + 1,
        itemId: v[i][idx.itemId],
        description: v[i][idx.description],
        unit: v[i][idx.unit],
        totalRunningStocks: Number(v[i][idx.total] || 0),
        status: v[i][idx.status],
        supplier: v[i][idx.supplier] || "",
        classification: v[i][idx.classification] || "",
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
      uprice: findIdx(headers, headerNames.uprice),
      amount: findIdx(headers, headerNames.amount),
    };

    for (let r = 1; r < data.length; r++) {
      const status = (data[r][idx.status] || "").toString();
      if (!status) continue;

      statusCounts[status] = (statusCounts[status] || 0) + 1;

      if (status.toLowerCase() === "pending") {
        const uprice = Number(data[r][idx.uprice] || 0);
        const qty = Number(data[r][idx.qty] || 0);
        const amount = uprice * qty;
        const row = {
          id: r + 1,
          sheetName: sheet.getName(),
          rowNumber: r + 1,
          itemId: data[r][idx.itemId] || "",
          description: data[r][idx.description] || "",
          qty: qty,
          unit: data[r][idx.unit] || "",
          branch: data[r][idx.branch] || "",
          email: data[r][idx.email] || "",
          status,
          uprice: uprice,
          amount: amount,
        };

        detectRequestType(sheet.getName()) === "SPECIAL"
          ? pendingSpecial.push(row)
          : pendingOffice.push(row);
      }
    }
  });

  return { 
    success: true, 
    pendingOffice, 
    pendingSpecial, 
    statusCounts, 
    addStocks,
    lowStockCount: addStocks.filter(s => s.totalRunningStocks <= 10).length
  };
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
    success: true,
    office: group(d.pendingOffice),
    special: group(d.pendingSpecial),
  };
}

function getPendingBreakdown() {
  const d = getDashboardData();
  return { success: true, office: d.pendingOffice.length, special: d.pendingSpecial.length };
}

// ===============================
// SUPPLIERS
// ===============================
function getSuppliers() {
  const sheet = getSheetByNameSafe("SUPPLIERS");
  if (!sheet) return { success: false, message: "SUPPLIERS sheet not found", data: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  
  const headers = data[0].map(h => ("" + h).toLowerCase());
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i + 1,
      supplier_id: data[i][headers.indexOf("supplier_id")] || "",
      supplier_name: data[i][headers.indexOf("supplier_name")] || "",
      contact_person: data[i][headers.indexOf("contact_person")] || "",
      email: data[i][headers.indexOf("email")] || "",
      phone: data[i][headers.indexOf("phone")] || "",
      address: data[i][headers.indexOf("address")] || "",
      classification: data[i][headers.indexOf("classification")] || "",
      status: data[i][headers.indexOf("status")] || "Active"
    });
  }
  
  return { success: true, data: result };
}

function addSupplier(data) {
  const sheet = getSheetByNameSafe("SUPPLIERS");
  if (!sheet) return { success: false, message: "SUPPLIERS sheet not found" };
  
  const row = [
    data.supplier_id || generateId(),
    data.supplier_name,
    data.contact_person,
    data.email,
    data.phone,
    data.address,
    data.classification || "General",
    data.status || "Active",
    getCurrentTimestamp()
  ];
  
  sheet.appendRow(row);
  return { success: true, message: "Supplier added successfully" };
}

function updateSupplier(data) {
  const sheet = getSheetByNameSafe("SUPPLIERS");
  if (!sheet) return { success: false, message: "SUPPLIERS sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(data.id);
  
  sheet.getRange(rowNum, headers.indexOf("supplier_name") + 1).setValue(data.supplier_name);
  sheet.getRange(rowNum, headers.indexOf("contact_person") + 1).setValue(data.contact_person);
  sheet.getRange(rowNum, headers.indexOf("email") + 1).setValue(data.email);
  sheet.getRange(rowNum, headers.indexOf("phone") + 1).setValue(data.phone);
  sheet.getRange(rowNum, headers.indexOf("address") + 1).setValue(data.address);
  sheet.getRange(rowNum, headers.indexOf("classification") + 1).setValue(data.classification);
  sheet.getRange(rowNum, headers.indexOf("status") + 1).setValue(data.status);
  
  return { success: true, message: "Supplier updated successfully" };
}

function deleteSupplier(id) {
  const sheet = getSheetByNameSafe("SUPPLIERS");
  if (!sheet) return { success: false, message: "SUPPLIERS sheet not found" };
  
  sheet.deleteRow(parseInt(id));
  return { success: true, message: "Supplier deleted successfully" };
}

// ===============================
// SUPPLIER ITEMS
// ===============================
function getSupplierItems() {
  const sheet = getSheetByNameSafe("SUPPLIER ITEMS");
  if (!sheet) return { success: false, message: "SUPPLIER ITEMS sheet not found", data: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  
  const headers = data[0].map(h => ("" + h).toLowerCase());
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i + 1,
      item_id: data[i][headers.indexOf("item_id")] || "",
      supplier_id: data[i][headers.indexOf("supplier_id")] || "",
      description: data[i][headers.indexOf("description")] || "",
      unit: data[i][headers.indexOf("unit")] || "",
      unit_price: data[i][headers.indexOf("unit_price")] || 0,
      min_stock_level: data[i][headers.indexOf("min_stock_level")] || 0,
      current_stock: data[i][headers.indexOf("current_stock")] || 0
    });
  }
  
  return { success: true, data: result };
}

function addSupplierItem(data) {
  const sheet = getSheetByNameSafe("SUPPLIER ITEMS");
  if (!sheet) return { success: false, message: "SUPPLIER ITEMS sheet not found" };
  
  const row = [
    data.item_id || generateId(),
    data.supplier_id,
    data.description,
    data.unit,
    data.unit_price || 0,
    data.min_stock_level || 0,
    data.current_stock || 0
  ];
  
  sheet.appendRow(row);
  return { success: true, message: "Supplier item added successfully" };
}

function updateSupplierItem(data) {
  const sheet = getSheetByNameSafe("SUPPLIER ITEMS");
  if (!sheet) return { success: false, message: "SUPPLIER ITEMS sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(data.id);
  
  sheet.getRange(rowNum, headers.indexOf("supplier_id") + 1).setValue(data.supplier_id);
  sheet.getRange(rowNum, headers.indexOf("description") + 1).setValue(data.description);
  sheet.getRange(rowNum, headers.indexOf("unit") + 1).setValue(data.unit);
  sheet.getRange(rowNum, headers.indexOf("unit_price") + 1).setValue(data.unit_price);
  sheet.getRange(rowNum, headers.indexOf("min_stock_level") + 1).setValue(data.min_stock_level);
  sheet.getRange(rowNum, headers.indexOf("current_stock") + 1).setValue(data.current_stock);
  
  return { success: true, message: "Supplier item updated successfully" };
}

function deleteSupplierItem(id) {
  const sheet = getSheetByNameSafe("SUPPLIER ITEMS");
  if (!sheet) return { success: false, message: "SUPPLIER ITEMS sheet not found" };
  
  sheet.deleteRow(parseInt(id));
  return { success: true, message: "Supplier item deleted successfully" };
}

// ===============================
// USERS
// ===============================
function getUsers() {
  const sheet = getSheetByNameSafe("USERS");
  if (!sheet) return { success: false, message: "USERS sheet not found", data: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  
  const headers = data[0].map(h => ("" + h).toLowerCase());
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i + 1,
      user_id: data[i][headers.indexOf("user_id")] || "",
      username: data[i][headers.indexOf("username")] || "",
      email: data[i][headers.indexOf("email")] || "",
      role: data[i][headers.indexOf("role")] || "",
      branch_access: data[i][headers.indexOf("branch_access")] || "",
      status: data[i][headers.indexOf("status")] || "Active",
      created_at: data[i][headers.indexOf("created_at")] || ""
    });
  }
  
  return { success: true, data: result };
}

function addUser(data) {
  const sheet = getSheetByNameSafe("USERS");
  if (!sheet) return { success: false, message: "USERS sheet not found" };
  
  const row = [
    data.user_id || generateId(),
    data.username,
    data.email,
    data.role || "Branch",
    data.branch_access || "",
    data.password_hash || "changeme123",
    data.status || "Active",
    getCurrentTimestamp()
  ];
  
  sheet.appendRow(row);
  return { success: true, message: "User added successfully" };
}

function updateUser(data) {
  const sheet = getSheetByNameSafe("USERS");
  if (!sheet) return { success: false, message: "USERS sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(data.id);
  
  sheet.getRange(rowNum, headers.indexOf("username") + 1).setValue(data.username);
  sheet.getRange(rowNum, headers.indexOf("email") + 1).setValue(data.email);
  sheet.getRange(rowNum, headers.indexOf("role") + 1).setValue(data.role);
  sheet.getRange(rowNum, headers.indexOf("branch_access") + 1).setValue(data.branch_access);
  sheet.getRange(rowNum, headers.indexOf("status") + 1).setValue(data.status);
  
  return { success: true, message: "User updated successfully" };
}

function deleteUser(id) {
  const sheet = getSheetByNameSafe("USERS");
  if (!sheet) return { success: false, message: "USERS sheet not found" };
  
  sheet.deleteRow(parseInt(id));
  return { success: true, message: "User deleted successfully" };
}

function resetPassword(id) {
  const sheet = getSheetByNameSafe("USERS");
  if (!sheet) return { success: false, message: "USERS sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(id);
  
  sheet.getRange(rowNum, headers.indexOf("password_hash") + 1).setValue("changeme123");
  
  return { success: true, message: "Password reset successfully" };
}

function getPasswordResetRequests() {
  const sheet = getSheetByNameSafe("PASSWORD RESET");
  if (!sheet) return { success: false, message: "PASSWORD RESET sheet not found", data: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  
  const headers = data[0].map(h => ("" + h).toLowerCase());
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i + 1,
      request_id: data[i][headers.indexOf("request_id")] || "",
      user_id: data[i][headers.indexOf("user_id")] || "",
      token: data[i][headers.indexOf("token")] || "",
      request_date: data[i][headers.indexOf("request_date")] || "",
      status: data[i][headers.indexOf("status")] || "Pending"
    });
  }
  
  return { success: true, data: result };
}

function handlePasswordReset(data) {
  const sheet = getSheetByNameSafe("PASSWORD RESET");
  if (!sheet) return { success: false, message: "PASSWORD RESET sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(data.id);
  
  sheet.getRange(rowNum, headers.indexOf("status") + 1).setValue(data.status);
  
  if (data.status === "Approved") {
    const usersSheet = getSheetByNameSafe("USERS");
    const usersData = usersSheet.getDataRange().getValues();
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][usersData[0].map(h => ("" + h).toLowerCase()).indexOf("user_id")] == data.user_id) {
        usersSheet.getRange(i + 1, usersData[0].map(h => ("" + h).toLowerCase()).indexOf("password_hash") + 1).setValue("changeme123");
        break;
      }
    }
  }
  
  return { success: true, message: "Password reset request handled" };
}

// ===============================
// BRANCHES
// ===============================
function getBranches() {
  const sheet = getSheetByNameSafe("BRANCHES");
  if (!sheet) return { success: false, message: "BRANCHES sheet not found", data: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  
  const headers = data[0].map(h => ("" + h).toLowerCase());
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i + 1,
      branch_id: data[i][headers.indexOf("branch_id")] || "",
      branch_name: data[i][headers.indexOf("branch_name")] || "",
      email: data[i][headers.indexOf("email")] || "",
      location: data[i][headers.indexOf("location")] || "",
      classification: data[i][headers.indexOf("classification")] || "",
      status: data[i][headers.indexOf("status")] || "Active",
      created_at: data[i][headers.indexOf("created_at")] || ""
    });
  }
  
  return { success: true, data: result };
}

function addBranch(data) {
  const sheet = getSheetByNameSafe("BRANCHES");
  if (!sheet) return { success: false, message: "BRANCHES sheet not found" };
  
  const row = [
    data.branch_id || generateId(),
    data.branch_name,
    data.email,
    data.location,
    data.classification || "Regular",
    data.status || "Active",
    getCurrentTimestamp()
  ];
  
  sheet.appendRow(row);
  return { success: true, message: "Branch added successfully" };
}

function updateBranch(data) {
  const sheet = getSheetByNameSafe("BRANCHES");
  if (!sheet) return { success: false, message: "BRANCHES sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(data.id);
  
  sheet.getRange(rowNum, headers.indexOf("branch_name") + 1).setValue(data.branch_name);
  sheet.getRange(rowNum, headers.indexOf("email") + 1).setValue(data.email);
  sheet.getRange(rowNum, headers.indexOf("location") + 1).setValue(data.location);
  sheet.getRange(rowNum, headers.indexOf("classification") + 1).setValue(data.classification);
  sheet.getRange(rowNum, headers.indexOf("status") + 1).setValue(data.status);
  
  return { success: true, message: "Branch updated successfully" };
}

function deleteBranch(id) {
  const sheet = getSheetByNameSafe("BRANCHES");
  if (!sheet) return { success: false, message: "BRANCHES sheet not found" };
  
  sheet.deleteRow(parseInt(id));
  return { success: true, message: "Branch deleted successfully" };
}

// ===============================
// INVENTORY
// ===============================
function getInventory() {
  const sheet = getSheetByNameSafe("ADD STOCKS");
  if (!sheet) return { success: false, message: "ADD STOCKS sheet not found", data: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  
  const headers = data[0].map(h => ("" + h).toLowerCase());
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i + 1,
      item_id: data[i][headers.indexOf("item_id")] || "",
      description: data[i][headers.indexOf("description")] || "",
      unit: data[i][headers.indexOf("unit")] || "",
      total_running_stocks: Number(data[i][headers.indexOf("total running stocks")] || 0),
      status: data[i][headers.indexOf("status")] || "",
      supplier: data[i][headers.indexOf("supplier")] || "",
      classification: data[i][headers.indexOf("classification")] || ""
    });
  }
  
  return { success: true, data: result };
}

function addInventoryItem(data) {
  const sheet = getSheetByNameSafe("ADD STOCKS");
  if (!sheet) return { success: false, message: "ADD STOCKS sheet not found" };
  
  const row = [
    data.item_id || generateId(),
    data.description,
    data.unit,
    data.total_running_stocks || 0,
    data.status || "Active",
    data.supplier || "",
    data.classification || ""
  ];
  
  sheet.appendRow(row);
  return { success: true, message: "Inventory item added successfully" };
}

function updateInventoryItem(data) {
  const sheet = getSheetByNameSafe("ADD STOCKS");
  if (!sheet) return { success: false, message: "ADD STOCKS sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(data.id);
  
  if (headers.indexOf("description") > -1)
    sheet.getRange(rowNum, headers.indexOf("description") + 1).setValue(data.description);
  if (headers.indexOf("unit") > -1)
    sheet.getRange(rowNum, headers.indexOf("unit") + 1).setValue(data.unit);
  if (headers.indexOf("total running stocks") > -1)
    sheet.getRange(rowNum, headers.indexOf("total running stocks") + 1).setValue(data.total_running_stocks);
  if (headers.indexOf("status") > -1)
    sheet.getRange(rowNum, headers.indexOf("status") + 1).setValue(data.status);
  if (headers.indexOf("supplier") > -1)
    sheet.getRange(rowNum, headers.indexOf("supplier") + 1).setValue(data.supplier);
  if (headers.indexOf("classification") > -1)
    sheet.getRange(rowNum, headers.indexOf("classification") + 1).setValue(data.classification);
  
  return { success: true, message: "Inventory item updated successfully" };
}

function deleteInventoryItem(id) {
  const sheet = getSheetByNameSafe("ADD STOCKS");
  if (!sheet) return { success: false, message: "ADD STOCKS sheet not found" };
  
  sheet.deleteRow(parseInt(id));
  return { success: true, message: "Inventory item deleted successfully" };
}

// ===============================
// OFFICE REQUESTS
// ===============================
function getOfficeRequests() {
  const sheet = getSheetByNameSafe("OFFICE REQUESTS");
  if (!sheet) return { success: false, message: "OFFICE REQUESTS sheet not found", data: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  
  const headers = data[0].map(h => ("" + h).toLowerCase());
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i + 1,
      sheetName: "OFFICE REQUESTS",
      rowNumber: i + 1,
      date: data[i][headers.indexOf("date")] || "",
      branch: data[i][headers.indexOf("branch")] || "",
      description: data[i][headers.indexOf("description")] || "",
      qty: Number(data[i][headers.indexOf("qty")] || 0),
      unit: data[i][headers.indexOf("unit")] || "",
      uprice: Number(data[i][headers.indexOf("uprice")] || 0),
      amount: Number(data[i][headers.indexOf("amount")] || 0),
      status: data[i][headers.indexOf("status")] || "",
      requested_by: data[i][headers.indexOf("requested_by")] || ""
    });
  }
  
  return { success: true, data: result };
}

function updateOfficeRequest(data) {
  const sheet = getSheetByNameSafe("OFFICE REQUESTS");
  if (!sheet) return { success: false, message: "OFFICE REQUESTS sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(data.id);
  
  if (data.status && headers.indexOf("status") > -1)
    sheet.getRange(rowNum, headers.indexOf("status") + 1).setValue(data.status);
  
  return { success: true, message: "Office request updated successfully" };
}

function deleteOfficeRequest(id) {
  const sheet = getSheetByNameSafe("OFFICE REQUESTS");
  if (!sheet) return { success: false, message: "OFFICE REQUESTS sheet not found" };
  
  sheet.deleteRow(parseInt(id));
  return { success: true, message: "Office request deleted successfully" };
}

// ===============================
// SPECIAL REQUESTS
// ===============================
function getSpecialRequests() {
  const sheet = getSheetByNameSafe("SPECIAL REQUESTS");
  if (!sheet) return { success: false, message: "SPECIAL REQUESTS sheet not found", data: [] };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, data: [] };
  
  const headers = data[0].map(h => ("" + h).toLowerCase());
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    result.push({
      id: i + 1,
      sheetName: "SPECIAL REQUESTS",
      rowNumber: i + 1,
      date: data[i][headers.indexOf("date")] || "",
      branch: data[i][headers.indexOf("branch")] || "",
      description: data[i][headers.indexOf("description")] || "",
      qty: Number(data[i][headers.indexOf("qty")] || 0),
      unit: data[i][headers.indexOf("unit")] || "",
      uprice: Number(data[i][headers.indexOf("uprice")] || 0),
      amount: Number(data[i][headers.indexOf("amount")] || 0),
      status: data[i][headers.indexOf("status")] || "",
      requested_by: data[i][headers.indexOf("requested_by")] || "",
      approved_by_purchasing: data[i][headers.indexOf("approved_by_purchasing")] || "",
      approved_by_accounting: data[i][headers.indexOf("approved_by_accounting")] || "",
      approved_by_admin: data[i][headers.indexOf("approved_by_admin")] || ""
    });
  }
  
  return { success: true, data: result };
}

function updateSpecialRequest(data) {
  const sheet = getSheetByNameSafe("SPECIAL REQUESTS");
  if (!sheet) return { success: false, message: "SPECIAL REQUESTS sheet not found" };
  
  const headers = sheet.getDataRange().getValues()[0].map(h => ("" + h).toLowerCase());
  const rowNum = parseInt(data.id);
  
  if (data.status && headers.indexOf("status") > -1)
    sheet.getRange(rowNum, headers.indexOf("status") + 1).setValue(data.status);
  if (data.approved_by_purchasing && headers.indexOf("approved_by_purchasing") > -1)
    sheet.getRange(rowNum, headers.indexOf("approved_by_purchasing") + 1).setValue(data.approved_by_purchasing);
  if (data.approved_by_accounting && headers.indexOf("approved_by_accounting") > -1)
    sheet.getRange(rowNum, headers.indexOf("approved_by_accounting") + 1).setValue(data.approved_by_accounting);
  if (data.approved_by_admin && headers.indexOf("approved_by_admin") > -1)
    sheet.getRange(rowNum, headers.indexOf("approved_by_admin") + 1).setValue(data.approved_by_admin);
  
  return { success: true, message: "Special request updated successfully" };
}

function deleteSpecialRequest(id) {
  const sheet = getSheetByNameSafe("SPECIAL REQUESTS");
  if (!sheet) return { success: false, message: "SPECIAL REQUESTS sheet not found" };
  
  sheet.deleteRow(parseInt(id));
  return { success: true, message: "Special request deleted successfully" };
}

// ===============================
// REPORTS
// ===============================
function generateReport(data) {
  const ss = SpreadsheetApp.openById(SHEET_FILE_ID);
  const { dateFrom, dateTo, branch, requestType } = data;
  
  const results = [];
  const sheets = ss.getSheets();
  
  const headerNames = {
    date: ["date", "date created", "request date"],
    branch: ["branch", "office", "location"],
    status: ["status", "approval"],
    qty: ["qty", "quantity"],
    amount: ["amount", "total amount"],
  };
  
  sheets.forEach((sheet) => {
    const sheetName = sheet.getName().toLowerCase();
    if (sheetName === "add stocks" || sheetName === "suppliers" || 
        sheetName === "supplier items" || sheetName === "users" ||
        sheetName === "branches" || sheetName === "password reset" ||
        sheetName === "low stocks log" || sheetName === "reports log") return;
    
    // Filter by request type
    if (requestType && requestType !== "All") {
      if (requestType === "Office" && !sheetName.includes("office")) return;
      if (requestType === "Special" && !sheetName.includes("special")) return;
    }
    
    const sheetData = sheet.getDataRange().getValues();
    if (sheetData.length < 2) return;
    
    const headers = sheetData[0].map(h => ("" + h).toLowerCase());
    const idx = {
      date: findIdx(headers, headerNames.date),
      branch: findIdx(headers, headerNames.branch),
      status: findIdx(headers, headerNames.status),
      qty: findIdx(headers, headerNames.qty),
      amount: findIdx(headers, headerNames.amount),
    };
    
    for (let r = 1; r < sheetData.length; r++) {
      const rowDate = sheetData[r][idx.date];
      const rowBranch = sheetData[r][idx.branch] || "";
      const rowStatus = sheetData[r][idx.status] || "";
      
      // Date filter
      if (dateFrom && dateTo && rowDate) {
        const rowDateObj = new Date(rowDate);
        const fromObj = new Date(dateFrom);
        const toObj = new Date(dateTo);
        if (rowDateObj < fromObj || rowDateObj > toObj) continue;
      }
      
      // Branch filter
      if (branch && branch !== "All" && rowBranch.toLowerCase() !== branch.toLowerCase()) continue;
      
      results.push({
        sheetName: sheet.getName(),
        rowNumber: r + 1,
        date: rowDate,
        branch: rowBranch,
        description: sheetData[r][headers.indexOf("description")] || "",
        qty: sheetData[r][idx.qty] || 0,
        amount: sheetData[r][idx.amount] || 0,
        status: rowStatus
      });
    }
  });
  
  return { success: true, data: results };
}

function getLowStocksReport() {
  const inv = getInventory();
  if (!inv.success) return inv;
  
  const lowStocks = inv.data.filter(item => {
    return item.total_running_stocks <= 10;
  });
  
  return { success: true, data: lowStocks };
}

function getBranchDashboard() {
  const branches = getBranches();
  const officeReq = getOfficeRequests();
  const specialReq = getSpecialRequests();
  
  const branchData = {};
  
  // Initialize branch data
  (branches.data || []).forEach(b => {
    branchData[b.branch_name] = {
      branch: b.branch_name,
      officePending: 0,
      officeApproved: 0,
      officeRejected: 0,
      specialPending: 0,
      specialApproved: 0,
      specialRejected: 0,
      calendar: []
    };
  });
  
  // Count office requests by branch
  (officeReq.data || []).forEach(req => {
    const b = req.branch;
    if (!branchData[b]) {
      branchData[b] = {
        branch: b,
        officePending: 0,
        officeApproved: 0,
        officeRejected: 0,
        specialPending: 0,
        specialApproved: 0,
        specialRejected: 0,
        calendar: []
      };
    }
    
    const status = (req.status || "").toLowerCase();
    if (status === "pending") branchData[b].officePending++;
    else if (status === "approved") branchData[b].officeApproved++;
    else if (status === "cancel" || status === "rejected") branchData[b].officeRejected++;
    
    if (req.date) {
      branchData[b].calendar.push({
        date: req.date,
        type: "Office",
        status: req.status,
        description: req.description
      });
    }
  });
  
  // Count special requests by branch
  (specialReq.data || []).forEach(req => {
    const b = req.branch;
    if (!branchData[b]) {
      branchData[b] = {
        branch: b,
        officePending: 0,
        officeApproved: 0,
        officeRejected: 0,
        specialPending: 0,
        specialApproved: 0,
        specialRejected: 0,
        calendar: []
      };
    }
    
    const status = (req.status || "").toLowerCase();
    if (status === "pending") branchData[b].specialPending++;
    else if (status === "approved" || status === "approved by accounting" || status === "to purchased") branchData[b].specialApproved++;
    else if (status === "cancel" || status === "rejected") branchData[b].specialRejected++;
    
    if (req.date) {
      branchData[b].calendar.push({
        date: req.date,
        type: "Special",
        status: req.status,
        description: req.description
      });
    }
  });
  
  return { success: true, data: Object.values(branchData) };
}

function getBranchReport(data) {
  const { branch, period, year, month } = data;
  
  const officeReq = getOfficeRequests();
  const specialReq = getSpecialRequests();
  
  const results = {
    branch: branch,
    officeRequests: [],
    specialRequests: [],
    summary: {
      officeTotal: 0,
      officePending: 0,
      officeApproved: 0,
      specialTotal: 0,
      specialPending: 0,
      specialApproved: 0
    }
  };
  
  const currentYear = year || new Date().getFullYear();
  const currentMonth = month ? parseInt(month) : null;
  
  const filterByPeriod = (dateStr) => {
    if (!dateStr) return true;
    const d = new Date(dateStr);
    if (period === "yearly") {
      return d.getFullYear() === currentYear;
    } else if (period === "monthly") {
      return d.getFullYear() === currentYear && (currentMonth ? d.getMonth() + 1 === currentMonth : true);
    }
    return true;
  };
  
  const filterByBranch = (b) => {
    return !branch || branch === "All" || b.toLowerCase() === branch.toLowerCase();
  };
  
  (officeReq.data || []).forEach(req => {
    if (filterByBranch(req.branch) && filterByPeriod(req.date)) {
      results.officeRequests.push(req);
      results.summary.officeTotal++;
      const status = (req.status || "").toLowerCase();
      if (status === "pending") results.summary.officePending++;
      else if (status === "approved") results.summary.officeApproved++;
    }
  });
  
  (specialReq.data || []).forEach(req => {
    if (filterByBranch(req.branch) && filterByPeriod(req.date)) {
      results.specialRequests.push(req);
      results.summary.specialTotal++;
      const status = (req.status || "").toLowerCase();
      if (status === "pending") results.summary.specialPending++;
      else if (status === "approved" || status === "approved by accounting") results.summary.specialApproved++;
    }
  });
  
  return { success: true, data: results };
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
  return { success: true };
}

function deletePendingRow(sheetName, rowNumber) {
  SpreadsheetApp.openById(SHEET_FILE_ID)
    .getSheetByName(sheetName)
    .deleteRow(rowNumber);
  return { success: true };
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
  return { success: true };
}

// ===============================
// BATCH PDF PRINT
// ===============================
function getBatchPdfUrls(rows) {
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

      const type = detectRequestType(r.sheetName);

      if (!selectedPerBranch[branch]) {
        selectedPerBranch[branch] = {
          url: pdfUrl,
          type: type,
          itemId: r.itemId,
          description: r.description,
        };
        return;
      }

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
        e.toString(),
      );
    }
  });

  return { success: true, data: selectedPerBranch };
}

// ===============================
// BATCH
// ===============================
function batchAction(rows, action) {
  if (!rows || !Array.isArray(rows)) {
    return { success: false, message: "Invalid rows parameter" };
  }
  rows.forEach((r) => rowAction(r.sheetName, r.rowNumber, action));
  return { success: true, message: "Batch action completed" };
}

function rowAction(sheetName, rowNumber, action) {
  const s = SpreadsheetApp.openById(SHEET_FILE_ID).getSheetByName(sheetName);
  const h = s
    .getRange(1, 1, 1, s.getLastColumn())
    .getValues()[0]
    .map((x) => ("" + x).toLowerCase());
  const idx = h.findIndex((x) => x.includes("status"));
  s.getRange(rowNumber, idx + 1).setValue(action);
  return { success: true };
}
