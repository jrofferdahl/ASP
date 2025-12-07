// Code.gs — WOP (Work Order Portal) Apps Script Web App API
// This script provides a Web App API for the Work Order Portal
// DO NOT modify the main ASP application files

// Configuration - Update this with your Apps Script Web App URL after deployment
const APPS_SCRIPT_URL = 'YOUR_WEB_APP_URL_HERE';

// Spreadsheet ID - Update with your Google Sheets ID
const SPREADSHEET_ID = '1hBLUJSVeNMjYo5ieYZRlhPUAFdqB7eEPkP5v0HzRlKM';

// Sheet names
const SHEET_WORK_ORDERS = 'WorkOrders';
const SHEET_ENGINEERS = 'Engineers';

/**
 * Get spreadsheet and sheet
 */
function getSheet_(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // Create sheet if it doesn't exist
    sheet = ss.insertSheet(sheetName);
    
    // Add headers based on sheet name
    if (sheetName === SHEET_WORK_ORDERS) {
      sheet.appendRow(['WorkOrderID', 'WorkOrderText', 'EngineerID', 'EngineerName', 'ScheduledDate', 'CreatedUTC', 'LastUpdate']);
    } else if (sheetName === SHEET_ENGINEERS) {
      sheet.appendRow(['EngineerID', 'EngineerName', 'EngineerEmail', 'Active']);
    }
  }
  
  return sheet;
}

/**
 * Read all rows from a sheet as objects
 */
function readSheet_(sheetName) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return [];
  
  const headers = data[0];
  const rows = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(cell => cell === '' || cell === null)) continue;
    
    const obj = {};
    headers.forEach((header, index) => {
      if (header) obj[header] = row[index];
    });
    rows.push(obj);
  }
  
  return rows;
}

/**
 * Find row index by key-value pair
 */
function findRowIndex_(sheet, keyColumn, keyValue) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIndex = headers.indexOf(keyColumn);
  
  if (keyIndex === -1) return -1;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIndex]) === String(keyValue)) {
      return i + 1; // Sheet rows are 1-indexed
    }
  }
  
  return -1;
}

/**
 * Update row by key
 */
function updateRowByKey_(sheetName, keyColumn, keyValue, updates) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowIndex = findRowIndex_(sheet, keyColumn, keyValue);
  
  if (rowIndex === -1) {
    throw new Error(`Row not found: ${keyColumn}=${keyValue}`);
  }
  
  Object.keys(updates).forEach(key => {
    const colIndex = headers.indexOf(key);
    if (colIndex !== -1) {
      sheet.getRange(rowIndex, colIndex + 1).setValue(updates[key]);
    }
  });
}

/**
 * Append row to sheet
 */
function appendRow_(sheetName, obj) {
  const sheet = getSheet_(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const row = headers.map(header => obj[header] || '');
  sheet.appendRow(row);
}

/**
 * Ensure required columns exist
 */
function ensureColumns_(sheetName, requiredColumns) {
  const sheet = getSheet_(sheetName);
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  
  requiredColumns.forEach(col => {
    if (!headers.includes(col)) {
      // Add missing column
      const newColIndex = headers.length + 1;
      sheet.getRange(1, newColIndex).setValue(col);
      headers.push(col);
    }
  });
}

/**
 * Web App doGet handler
 */
function doGet(e) {
  const op = e.parameter.op;
  
  try {
    if (op === 'getEngineers') {
      const engineers = readSheet_(SHEET_ENGINEERS);
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        engineers: engineers
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: 'Unknown operation'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Web App doPost handler
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    
    if (action === 'assignEngineer') {
      return assignEngineerPost_(payload);
    } else if (action === 'scheduleWorkOrder') {
      return scheduleWorkOrderPost_(payload);
    } else if (action === 'createRow') {
      return createRowPost_(payload);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: 'Unknown action'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Assign engineer to work order
 */
function assignEngineerPost_(payload) {
  const { workOrderId, engineerId, engineerName } = payload;
  
  if (!workOrderId || !engineerId) {
    throw new Error('workOrderId and engineerId are required');
  }
  
  // Ensure columns exist
  ensureColumns_(SHEET_WORK_ORDERS, ['WorkOrderID', 'EngineerID', 'EngineerName', 'LastUpdate']);
  
  try {
    // Update existing row
    updateRowByKey_(SHEET_WORK_ORDERS, 'WorkOrderID', workOrderId, {
      EngineerID: engineerId,
      EngineerName: engineerName || '',
      LastUpdate: new Date()
    });
  } catch (error) {
    // Row doesn't exist, create it
    appendRow_(SHEET_WORK_ORDERS, {
      WorkOrderID: workOrderId,
      EngineerID: engineerId,
      EngineerName: engineerName || '',
      CreatedUTC: new Date(),
      LastUpdate: new Date()
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    message: 'Engineer assigned successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Schedule work order
 */
function scheduleWorkOrderPost_(payload) {
  const { workOrderId, scheduleDate } = payload;
  
  if (!workOrderId || !scheduleDate) {
    throw new Error('workOrderId and scheduleDate are required');
  }
  
  // Ensure columns exist
  ensureColumns_(SHEET_WORK_ORDERS, ['WorkOrderID', 'ScheduledDate', 'LastUpdate']);
  
  try {
    // Update existing row
    updateRowByKey_(SHEET_WORK_ORDERS, 'WorkOrderID', workOrderId, {
      ScheduledDate: scheduleDate,
      LastUpdate: new Date()
    });
  } catch (error) {
    // Row doesn't exist, create it
    appendRow_(SHEET_WORK_ORDERS, {
      WorkOrderID: workOrderId,
      ScheduledDate: scheduleDate,
      CreatedUTC: new Date(),
      LastUpdate: new Date()
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    message: 'Work order scheduled successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Create a new row in WorkOrders sheet
 */
function createRowPost_(payload) {
  // Ensure columns exist
  ensureColumns_(SHEET_WORK_ORDERS, ['WorkOrderID', 'WorkOrderText', 'EngineerID', 'EngineerName', 'ScheduledDate', 'CreatedUTC', 'LastUpdate']);
  
  const row = {
    WorkOrderID: payload.WorkOrderID || '',
    WorkOrderText: payload.WorkOrderText || '',
    EngineerID: payload.EngineerID || '',
    EngineerName: payload.EngineerName || '',
    ScheduledDate: payload.ScheduledDate || '',
    CreatedUTC: new Date(),
    LastUpdate: new Date()
  };
  
  appendRow_(SHEET_WORK_ORDERS, row);
  
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    message: 'Row created successfully'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Get engineers (for Google Apps Script calls)
 */
function getEngineers() {
  return readSheet_(SHEET_ENGINEERS);
}

/**
 * Get work orders (for Google Apps Script calls)
 */
function getWorkOrders() {
  return readSheet_(SHEET_WORK_ORDERS);
}

/**
 * Create work order (for Google Apps Script calls)
 */
function createWorkOrder(payload) {
  ensureColumns_(SHEET_WORK_ORDERS, ['WorkOrderID', 'WorkOrderText', 'CreatedUTC', 'LastUpdate']);
  
  const workOrderId = payload.WorkOrderID || generateWorkOrderId();
  const workOrderText = payload.WorkOrderText || ' FROM WOP';
  
  appendRow_(SHEET_WORK_ORDERS, {
    WorkOrderID: workOrderId,
    WorkOrderText: workOrderText,
    CreatedUTC: new Date(),
    LastUpdate: new Date()
  });
  
  return {
    ok: true,
    WorkOrderID: workOrderId,
    message: 'Work order created successfully'
  };
}

/**
 * Generate Work Order ID in format WO-YYYYMMDD-HHMMSS (UTC)
 */
function generateWorkOrderId() {
  const now = new Date();
  const year = now.getUTCFullYear();
  const month = String(now.getUTCMonth() + 1).padStart(2, '0');
  const day = String(now.getUTCDate()).padStart(2, '0');
  const hours = String(now.getUTCHours()).padStart(2, '0');
  const minutes = String(now.getUTCMinutes()).padStart(2, '0');
  const seconds = String(now.getUTCSeconds()).padStart(2, '0');
  
  return `WO-${year}${month}${day}-${hours}${minutes}${seconds}`;
}

/**
 * Assign engineer (for Google Apps Script calls)
 */
function assignEngineer(payload) {
  return JSON.parse(assignEngineerPost_(payload).getContent());
}

/**
 * Schedule work order (for Google Apps Script calls)
 */
function scheduleWorkOrder(payload) {
  return JSON.parse(scheduleWorkOrderPost_(payload).getContent());
}
