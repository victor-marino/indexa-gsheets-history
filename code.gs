// SETTINGS //
var token = "API_TOKEN"; // Your Indexa Capital API token
var accountNumber = "ACCOUNT_NUMBER"; // Your Indexa Capital account number
var filename = "FILE NAME"; // Desired name for the Google Sheets file. Will be created if it doesn't exist.
var newestOnTop = false; // Set to true if you want the table sorted from newest to oldest
//////////////

const tableHeaders = [["Date", "Price", "Titles", "Cost Amount", "Amount"]];

function run() {
  let file = getActiveSpreadsheet(filename);
  let positions = getAccountPositions(token, accountNumber);
  updateSheet(positions, file);
}

function getActiveSpreadsheet(filename) {
  // Find a Google Sheets document with the specified title
  let files = DriveApp.searchFiles('title = "' + filename + '" and mimeType = "' + MimeType.GOOGLE_SHEETS + '"');
  // Open it if it exists, otherwise create it.
  if (files.hasNext()) {
    Logger.log("Existing file " + filename + " found");
    return SpreadsheetApp.open(files.next());
  } else {
    Logger.log("File " + filename + " not found. Creating it.");
    return SpreadsheetApp.create(filename);
  }  
}


function getAccountPositions(token, accountNumber) {
  // Prepare API request
  let host = "https://api.indexacapital.com";
  let url = host + "/accounts/" + accountNumber + "/portfolio";
  let options = {
    'method': 'get',
    'headers': {'X-AUTH-TOKEN': token, 'Accept' : '*/*'},
    'muteHttpExceptions': true
  };
  // Send query
  let response = UrlFetchApp.fetch(url, options);
  // Parse response
  let data = JSON.parse(response.getContentText());
  // Return data from all positions
  return data.instrument_accounts[0].positions;
}


function extractPositionValues(positionData) {
  let position = {};
  // Get the identifier name (ISIN, DGS or EPSV)
  let identifierName = positionData.instrument.identifier_name;
  position.instrumentIdentifierName = positionData.instrument.identifier_name;
  
  // Get the correct identifier variable name for the instrument and any subfunds (in case of pension plans)
  let identifierVariable = "";
  let subfundIdentifierVariable = "";

  if (identifierName == "ISIN") {
    identifierVariable = "isin_code";
  }
  if (identifierName == "DGS") {
    identifierVariable = "dgs_code";
    subfundIdentifierVariable = "dgs_fund_code";
  }
  if (identifierName == "EPSV") {
    identifierVariable = "epsv_plan_code";
    subfundIdentifierVariable = "epsv_fund_code";
  }

  // Populate the value of the instrument identifier
  let identifier = positionData.instrument[identifierVariable];
  let subfundIdentifier = "";
  if (subfundIdentifierVariable != "" && positionData.instrument[subfundIdentifierVariable] != "") {
    subfundIdentifier = positionData.instrument[subfundIdentifierVariable];
    position.instrumentIdentifier = identifier + " - " + subfundIdentifier;
  } else {
    position.instrumentIdentifier = identifier;
  }
  
  // Populate the other variables related to the instrument
  position.instrumentName = positionData.instrument.name;
  position.instrumentAssetClass = positionData.instrument.asset_class;
  position.instrumentManagementCompanyDescription = positionData.instrument.management_company_description;
  
  // Populate the variables related to the current position status
  position.price = positionData.price;
  position.titles = positionData.titles;
  position.date = positionData.date;
  position.amount = positionData.amount;
  position.costAmount = positionData.cost_amount;

  // Return the object containing all the variables
  return position;
}

function sheetIsEmpty(sheet) {
  // Check if selected sheet is empty
  return sheet.getLastRow() == 0 && sheet.getLastColumn() == 0;
}

function addCurrentValues(sheet, position) {
  // Add current position values to the table
  let lastRow = sheet.getRange("A1:A").getValues().filter(String).length;
  let newRange;
  if (newestOnTop) {
    Logger.log("Inserting new row at the top");
    newRange = sheet.getRange("A2:E2");
    if (!newRange.isBlank()) {
      // If the table is not empty, we shift down the existing data to add the new entry on top.
      newRange.insertCells(SpreadsheetApp.Dimension.ROWS);
    }
  } else {
    Logger.log("Adding new row at the bottom");
    let nextRow = lastRow + 1;
    newRange = sheet.getRange("A" + nextRow + ":E" + nextRow);
  }
  newRange.setValues(
    [
      [position.date, position.price, position.titles, position.costAmount, position.amount]
    ]
  );
}

function addInstrumentInfo(sheet, position) {
  sheet.getRange('G1:H4').setValues(
    [
      ['Name', position.instrumentName],
      [position.instrumentIdentifierName, position.instrumentIdentifier],
      ["Asset class", position.instrumentAssetClass],
      ["Management Company", position.instrumentManagementCompanyDescription]
    ]
  );
}

function formatSheet(sheet) {
    sheet.setColumnWidths(1, 5, 110);
    sheet.setColumnWidths(7, 1, 200);
    sheet.setColumnWidths(8, 1, 300);
    sheet.getRange('A1:E1').setFontWeight('bold');
    sheet.getRange('G1:G4').setFontWeight('bold');
    sheet.getRange('G1:H4').setBorder(true, true, true, true, false, true, null, SpreadsheetApp.BorderStyle.SOLID);
}

function updateSheet(positions, file) {
  // Loop through each position
  for (let i = 0; i < positions.length; i++) {
    position = extractPositionValues(positions[i]);
    // Check if there's already a sheet for the current instrument
    sheet = file.getSheetByName(position.instrumentIdentifier);
    if (!sheet) {
      // Otherwise, create it.
      Logger.log("Sheet " + position.instrumentIdentifier + " not found. Inserting it.");
      file.insertSheet(position.instrumentIdentifier);
      sheet = file.getSheetByName(position.instrumentIdentifier);
    }
    
    Logger.log("Selected sheet: " + sheet.getName());

    if (sheetIsEmpty(sheet)) {
      // If the sheet is empty add the table headers and the instrument information
      Logger.log("Sheet is empty. Adding headers and instrument information.");
      sheet.getRange('A1:E1').setValues(tableHeaders);
      addInstrumentInfo(sheet, position);
      // Add the first entry with the current position status
      Logger.log("Adding current position status.");
      addCurrentValues(sheet, position);
      
      // Apply sheet formatting
      formatSheet(sheet);
      continue;
    }
    
    Logger.log("Sheet is not empty. Checking contents...");

    if (JSON.stringify(sheet.getRange('A1:E1').getValues()) != JSON.stringify(tableHeaders)) {
      // If the sheet is not empty and the headers don't match our format, we assume there's something else on this sheet and stop execution.
      Logger.log("Headers don't look right, stopping script. Is this sheet being used for something else?");
      Logger.log("Expected " + tableHeaders + ", found " + sheet.getRange('A1:E1').getValues());
      continue;  
    }

    Logger.log("Headers look correct. Checking if there's new data to add.");

    // Headers look right, so we check the date column to see if today's data is already in place.
    let sheetIsUpToDate = sheet.getRange("A1:A").createTextFinder(position.date).matchEntireCell(true).findNext();
    if (sheetIsUpToDate) {
      Logger.log("Sheet already contains latest data from " + position.date + ". Nothing to add.");
      continue;
    }

    // Today's data is not in the table, so we add it.
    Logger.log("Adding today's values.");
    addCurrentValues(sheet, position);

    // Finally, we update the instrument information, just in case it has changed.
    Logger.log("Updating instrument information.");
    addInstrumentInfo(sheet, position);
    
    // Apply sheet formatting
    formatSheet(sheet);
  }
}
