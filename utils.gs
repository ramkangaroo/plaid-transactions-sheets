//Active Spreadsheet
const ss = SpreadsheetApp.getActiveSpreadsheet();

//Sheet names
const SECRETS_SHEET_NAME = "secrets";
const ACCOUNTS_SHEET_NAME = "accounts";
const BALANCES_SHEET_NAME = "balances";
const TRANSACTIONS_SHEET_NAME = "transactions";
const CATEGORIES_SHEET_NAME = "categories"

//Sheets
const secretsSheet = ss.getSheetByName(SECRETS_SHEET_NAME);
const accSheet = ss.getSheetByName(ACCOUNTS_SHEET_NAME);
const balSheet = ss.getSheetByName(BALANCES_SHEET_NAME);
const transSheet = ss.getSheetByName(TRANSACTIONS_SHEET_NAME);
const catSheet = ss.getSheetByName(CATEGORIES_SHEET_NAME);

//Key and Secret data for Plaid development environment
const url = secretsSheet.getRange("B3").getValue();
const client_id = secretsSheet.getRange("B4").getValue();
const secret = secretsSheet.getRange("B5").getValue();

//Formated dates for reference in sheet and values returned
const formatDate = date => Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
const formatDateToday = formatDate(new Date());
const formatLastRunDate = formatDate(new Date(secretsSheet.getRange("B12").getValue()));

//Spreadsheet UI menu to run functions on demand
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Bank Sync')
    .addItem('Refresh Accounts', 'bankSync')
    .addItem('Update Accounts', 'processTokens')
    .addToUi();
}

//API call function
function makeRequest(url, params, errorMsg) {
  try {
    // Make the POST request
    const response = UrlFetchApp.fetch(url, params);
    secretsSheet.getRange(errorMsg).setValue(null);

    return response.getContentText();

  } catch (error) {
    Logger.log(`Error: ${error}`);
    secretsSheet.getRange(errorMsg).setValue("Error");
    return null
  }
}

//Duplicate row removal - needs work
function removeDups(sheetName) {
  try {
    const sheet = ss.getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const columnIndices = (sheetName === BALANCES_SHEET_NAME) ? [1,8] : [1];
 
    sheet.getRange(2, 1, lastRow, lastCol).removeDuplicates(columnIndices);

  } catch (error) {
    Logger.log(`Error removing duplicates from ${sheetName}: ${error}`);
  }
}
