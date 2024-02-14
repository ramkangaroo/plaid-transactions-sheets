function getLinkToken(access_token,errorMsg) {
    const url = secretsSheet.getRange("B3").getValue();
    const client_id = secretsSheet.getRange("B4").getValue();
    const secret = secretsSheet.getRange("B5").getValue();

    // Prepare the request body
    const updateBody = {
        "client_id": client_id,
        "secret": secret,
        "client_name": "Z&A Budget",
        "country_codes": ["CA"],
        "language": "en",
        "user": {
            "client_user_id": "1234"
        },
        "access_token": access_token
    };

    // Condense the above into a single object
    const updateParams = {
        "contentType": "application/json",
        "payload": JSON.stringify(updateBody),
    };

    // Make the POST request
    const updateResult = JSON.parse(makeRequest(`${url}/link/token/create`, updateParams,errorMsg));

    return updateResult ? updateResult.link_token : null;
}

function processTokens() {

  // Find error cells and grab link token
  var error_cell_text = secretsSheet.createTextFinder('Error').findNext();
  if (error_cell_text){
    var access_token = error_cell_text.offset(0,-1).getValue();
    var errorCell = secretsSheet.createTextFinder(access_token).findNext().offset(0,1).getA1Notation();
    var linkToken = getLinkToken(access_token,errorCell);

    //Set link token to cell next to cursor. This is to process the banksync function for a specific access token
    secretsSheet.createTextFinder(access_token).findNext().offset(0,3).setValue('Update');

    if (linkToken) {

      var template = HtmlService.createTemplateFromFile('link.html');
      template.linkToken = linkToken;
      var dialog = template.evaluate()
        .setWidth(375)
        .setHeight(600);
      SpreadsheetApp.getUi().showModelessDialog(dialog,'Update Link');

    }
  } else {
      return singleBankSync();
  }
}

function singleBankSync(){

  var updateFinder = secretsSheet.createTextFinder('Update').findNext();

  if(updateFinder){
  var singleAccessToken = updateFinder.offset(0,-3).getValue();
  var singleCursor = updateFinder.offset(0,-1).getValue();
  var singleBank = updateFinder.offset(0,-4).getValue();
  var singleErrorMsg = updateFinder.offset(0,-2).getA1Notation();

  downloadTransactions(singleCursor, singleAccessToken, singleBank, singleErrorMsg);

  secretsSheet.createTextFinder('Update').findNext().setValue(null);
  }
}
}
