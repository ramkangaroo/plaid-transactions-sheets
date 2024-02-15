/**
 * Create a template for the Plaid investment transations download project
 * 
 * Author: Edrick Larkin
 * Project Page: https://github.com/edricklarkin/investment-transactions/
 * 
 */

function createTemplate() {

  //create a new sheets file  
  var ss = SpreadsheetApp.create("Plaid Finance Transactions");

  //add sheets to the file
  var acct = ss.insertSheet("accounts", 0);
  var bal = ss.insertSheet("balances", 1);
  var trans = ss.insertSheet("transactions", 2);
  var secrets = ss.insertSheet("secrets", 3);
  var cat = ss.insertSheet("categories", 4);

  //create headers for account sheet
  let account_header = [["account_id","official_name","current_balance","available_balance","limit","type","subtype","date"]];
  acct.getRange(1,1,1,8).setValues(account_header);

  //create headers for balances
  let balance_headers = [["account_id","official_name","current_balance","available_balance","limit","type","subtype","date"]];
  bal.getRange(1,1,1,8).setValues(balance_headers);

  //create headers for transactions sheet
  let trans_headers = [["account_id","official_name","current_balance","available_balance","limit","type","subtype","date","override_name"]];
  trans.getRange(1,1,1,9).setValues(trans_headers);

  //setup secrets tab
  secrets.getRange("A1").setValue("Important: This tab is where you store the secrets for each institution");

  //create headers for transactions sheet
  let cat_headers = [["PRIMARY","DETAILED","DESCRIPTION","categoryDetail","categoryGroup"]];
  cat.getRange(1,1,1,5).setValues(cat_headers);

  let generic_secrets = [["url", "https://sandbox.plaid.com"], ["client_id", "{Enter client_id from the Plaid dashboard}"], ["secret","{Enter secret from Plaid dashboard}"]];

  let inst_secrets = [["instituion_name", "access_token"], ["{Name for first institution}", "{Enter access_token from Plaid quickstart}"],["{Name for second institution}", "{Enter access_token from Plaid quickstart}"],["Enter any number of insitutions and access tokens", ""]];

  secrets.getRange("A3:B5").setValues(generic_secrets);
  secrets.getRange("A7:B10").setValues(inst_secrets);  


}
