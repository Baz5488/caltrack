/**
 * Google Apps Script backend for Personal Calorie Tracker
 * Stores Food Database, Daily Food Logs, and User Targets in a bound Google Sheet.
 */

const SHEETS = {
  LOGS: 'FoodLogs',
  DB: 'FoodDatabase',
  SETTINGS: 'Settings'
};

/**
 * Serves the HTML frontend when deployed as a Web App.
 * Make sure your HTML file in Apps Script is named 'index.html' (or update below).
 */
function doGet(e) {
  // If ?action=... is passed, act as a JSON GET endpoint
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter.action, e.parameter);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Personal Calorie Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Handles JSON POST requests if calling the Web App via fetch()
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const result = handleApiRequest(payload.action, payload.data || {});
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Unified router for both google.script.run and HTTP fetch()
 */
function handleApiRequest(action, data) {
  setupSheetsIfNeeded();

  switch (action) {
    case 'getAllData':
      return getAllData();
    case 'addLog':
      return addLogEntry(data);
    case 'deleteLog':
      return deleteLogEntry(data.id);
    case 'saveFood':
      return saveFoodItem(data);
    case 'deleteFood':
      return deleteFoodItem(data.id);
    case 'saveTargets':
      return saveTargets(data.calories, data.protein);
    case 'importBackup':
      return importBackupData(data);
    default:
      throw new Error('Unknown action: ' + action);
  }
}

/**
 * Initializes required tabs and headers in the Google Sheet if they do not exist.
 */
function setupSheetsIfNeeded() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Daily Food Logs
  let logSheet = ss.getSheetByName(SHEETS.LOGS);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEETS.LOGS);
    logSheet.appendRow([
      'id', 'date', 'meal', 'foodId', 'foodName', 'brand',
      'servings', 'calories', 'protein', 'carbs', 'fat', 'createdAt'
    ]);
    logSheet.setFrozenRows(1);
  }

  // 2. Food Database
  let dbSheet = ss.getSheetByName(SHEETS.DB);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(SHEETS.DB);
    dbSheet.appendRow([
      'id', 'name', 'brand', 'category', 'servingAmount',
      'unit', 'calories', 'protein', 'carbs', 'fat', 'barcode'
    ]);
    dbSheet.setFrozenRows(1);
  }

  // 3. Settings / Targets
  let settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEETS.SETTINGS);
    settingsSheet.appendRow(['key', 'value']);
    settingsSheet.appendRow(['calorieTarget', 1500]);
    settingsSheet.appendRow(['proteinTarget', 100]);
    settingsSheet.setFrozenRows(1);
  }
}

/**
 * Fetches all logs, food database items, and targets in one call.
 */
function getAllData() {
  setupSheetsIfNeeded();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return {
    logs: sheetToObjects(ss.getSheetByName(SHEETS.LOGS)),
    foods: sheetToObjects(ss.getSheetByName(SHEETS.DB)),
    targets: getTargets()
  };
}

/**
 * Logs a food entry for a specific date and meal.
 */
function addLogEntry(entry) {
  setupSheetsIfNeeded();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGS);
  const id = entry.id || Utilities.getUuid();
  const createdAt = new Date().toISOString();

  const row = [
    id,
    entry.date,
    entry.meal,
    entry.foodId || '',
    entry.foodName,
    entry.brand || '',
    Number(entry.servings) || 1,
    Number(entry.calories) || 0,
    Number(entry.protein) || 0,
    Number(entry.carbs) || 0,
    Number(entry.fat) || 0,
    createdAt
  ];

  sheet.appendRow(row);
  return { ...entry, id, createdAt };
}

/**
 * Deletes a logged food entry by ID.
 */
function deleteLogEntry(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGS);
  return deleteRowById(sheet, id);
}

/**
 * Adds a new food or updates an existing one in the Food Database.
 */
function saveFoodItem(food) {
  setupSheetsIfNeeded();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.DB);
  const id = food.id || Utilities.getUuid();

  const rowValues = [
    id,
    food.name,
    food.brand || '',
    food.category || 'Other',
    Number(food.servingAmount) || 100,
    food.unit || 'g',
    Number(food.calories) || 0,
    Number(food.protein) || 0,
    Number(food.carbs) || 0,
    Number(food.fat) || 0,
    food.barcode || ''
  ];

  const rowIndex = findRowIndexById(sheet, id);
  if (rowIndex > 1) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { ...food, id };
}

/**
 * Deletes a food item from the database by ID.
 */
function deleteFoodItem(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.DB);
  return deleteRowById(sheet, id);
}

/**
 * Reads daily calorie and protein targets.
 */
function getTargets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  const targets = { calories: 1500, protein: 100 };

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'calorieTarget') targets.calories = Number(data[i][1]) || 1500;
    if (data[i][0] === 'proteinTarget') targets.protein = Number(data[i][1]) || 100;
  }
  return targets;
}

/**
 * Updates daily calorie and protein targets.
 */
function saveTargets(calories, protein) {
  setupSheetsIfNeeded();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'calorieTarget') {
      sheet.getRange(i + 1, 2).setValue(Number(calories) || 1500);
    }
    if (data[i][0] === 'proteinTarget') {
      sheet.getRange(i + 1, 2).setValue(Number(protein) || 100);
    }
  }
  return { calories: Number(calories), protein: Number(protein) };
}

/**
 * Restores data from your existing app's Backup JSON structure.
 */
function importBackupData(backup) {
  setupSheetsIfNeeded();
  if (backup.targets) {
    saveTargets(backup.targets.calories, backup.targets.protein);
  }
  if (Array.isArray(backup.foods)) {
    backup.foods.forEach(food => saveFoodItem(food));
  }
  if (Array.isArray(backup.logs)) {
    backup.logs.forEach(log => addLogEntry(log));
  }
  return getAllData();
}

/* ================= HELPER FUNCTIONS ================= */

function sheetToObjects(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, idx) => {
      let val = row[idx];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[header] = val;
    });
    return obj;
  });
}

function findRowIndexById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function deleteRowById(sheet, id) {
  const rowIndex = findRowIndexById(sheet, id);
  if (rowIndex > 1) {
    sheet.deleteRow(rowIndex);
    return true;
  }
  return false;
}

