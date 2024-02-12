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
    var access_token = secretsSheet.createTextFinder('Error').findNext().offset(0,-1).getValue();
    var errorCell = secretsSheet.createTextFinder(access_token).findNext().offset(0,1).getA1Notation();
    var linkToken = getLinkToken(access_token,errorCell);

    if (linkToken) {
      var template = HtmlService.createTemplateFromFile('link.html');
      template.linkToken = linkToken;
      var dialog = template.evaluate()
        .setWidth(375)
        .setHeight(600);
      SpreadsheetApp.getUi().showModelessDialog(dialog,'Update Link');
    }
}
