/**
 * Download transactions from Plaid /transactions/get to google sheets
 * 
 * Author: Edrick Larkin
 * Project Page: https://github.com/edricklarkin/investment-transactions/code.gs
 * 
 **/

/**
 * Makes API call using makeRequest() and appends results to the account and transaction tabs.
 * Plaid /transactions/get paginates. @offset is required offset from prior run.
 * Returns remaining number of transactions to download in future call
 */
function downloadTransactions(cursor, access_token, bank, errorMsg) {
  
  try {
  // Prepare the account request body
  const accBody = {
	"client_id": client_id,
	"secret": secret,
	"access_token": access_token
  }

  // Prepare the transaction request body
  const transBody = {
    "client_id": client_id,
    "secret": secret,
    "access_token": access_token,
    "cursor": cursor,
    "count": 500,
    "options": {
      "include_personal_finance_category": true
    }
  };

  // Condense the above into a single object
  const accParams = {
    "contentType": "application/json",
    "payload": JSON.stringify(accBody),
  }
  const transParams = {
    "contentType": "application/json",
    "payload": JSON.stringify(transBody),
  };

  // Make the POST request
  const accResult = JSON.parse(makeRequest(`${url}/accounts/get`, accParams, errorMsg));
  const transResult = JSON.parse(makeRequest(`${url}/transactions/sync`, transParams, errorMsg));

  //Number of transactions in the result object
  const total_count = transResult.added.length;

  Logger.log(`${total_count} available transactions downloaded from ${bank}.`);

  // Create array of account results
  const account_array = accResult.accounts.map(account => [
    account.account_id,
    account.official_name,
    account.balances.current,
    account.balances.available,
    account.balances.limit,
    account.type,
    account.subtype,
    formatDateToday
  ]);

  // Find last rows to start appending data
  const acctLastRow = accSheet.getLastRow() + 1;
  const balLastRow = balSheet.getLastRow() + 1;

  // Append data to sheets
  accSheet.getRange(acctLastRow, 1, account_array.length, account_array[0].length).setValues(account_array);
  balSheet.getRange(balLastRow, 1, account_array.length, account_array[0].length).setValues(account_array);

  // Check if the request has transactions
  if (total_count < 1) {
    return 0;
  }

  // Create array of transaction results to append to spreadsheet
  const trans_array = transResult.added.map(added => [
    added.transaction_id,
    added.account_id,
    added.personal_finance_category.detailed,
    added.amount,
    added.authorized_date,
    added.name,
    added.payment_channel,
    getPersonalFinanceCategory(added, catSheet)
  ]);

  //Find last rows to start appending data
  const transLastRow = transSheet.getLastRow() + 1;

  //Append data to sheets
  transSheet.getRange(transLastRow, 1, trans_array.length, trans_array[0].length).setValues(trans_array);

  //Append next cursor to secrets
  secretsSheet.createTextFinder(bank).findNext().offset(0,3).setValue(transResult.next_cursor);
  
  } catch (error) {
    Logger.log(`Error: ${error}`);
    //throw new Error(error);
  }
}

/**
 * Top level function that calls all other functions
 * Loops through each instituion's access token on secrets tab
 * Calls downloadTransactions() until no additional transactions are available from Plaid
 * Removeds duplate transactions with removeDups()
 **/
function bankSync() {

  //Get access token
  var token_cell = SpreadsheetApp.getActiveSpreadsheet().getRange("secrets!B8");

  //Loop through each institution's client token
  let z = 0

  while (token_cell.offset(z,0).isBlank() == false){
    //Get cursor
    var cursor = token_cell.offset(z,2).getValue();

    //Get bank name
    var bank = token_cell.offset(z,-1).getValue();

    //Get error message cell a1notation
    var errorMsg = token_cell.offset(z,1).getA1Notation();
    
    downloadTransactions(cursor, token_cell.offset(z,0).getValue(), bank, errorMsg);
    z = z + 1;
  }

  //Sort Transactions by date
  transSheet.getRange(2, 1, transSheet.getLastRow(), transSheet.getLastColumn()).sort({column: 5, ascending: false});
  accSheet.getRange(2, 1, accSheet.getLastRow(),  accSheet.getLastColumn()).sort({column: 8, ascending: false});
  balSheet.getRange(2, 1, balSheet.getLastRow(),  balSheet.getLastColumn()).sort({column: 8, ascending: false});

  //Stamp last successful script run
  secretsSheet.getRange("B12").setValue(new Date());

  //Remove any duplicate records after all API calls
  removeDups(ACCOUNTS_SHEET_NAME);
  removeDups(TRANSACTIONS_SHEET_NAME);
  removeDups(BALANCES_SHEET_NAME);
}

// Function to get Personal Finance Category using VLOOKUP-like logic
function getPersonalFinanceCategory(added, categoriesSheet) {
  if (added.personal_finance_category && added.personal_finance_category.detailed) {
    const category = added.personal_finance_category.detailed;
    const categoriesRange = categoriesSheet.getRange("B1:E105").getValues();

    for (let i = 0; i < categoriesRange.length; i++) {
      if (categoriesRange[i][0] == category) {
        return categoriesRange[i][2];
      }
    }
  }
  return "";
}
