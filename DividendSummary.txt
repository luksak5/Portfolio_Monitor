function generateDividendSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Input"); // Adjust sheet name if needed
  var abcSheet = ss.getSheetByName("Dividend Amount"); // Adjust sheet name if needed

  if (!inputSheet || !abcSheet) {
    logger.log("Error: One or both sheets are missing.");
    return;
  }

  var inputData = inputSheet.getDataRange().getValues();
  var abcData = abcSheet.getDataRange().getValues();

  if (inputData.length < 2 || abcData.length < 2) {
    logger.log("Error: One or both sheets have no data.");
    return;
  }

  // Extract headers
  var inputHeaders = inputData[0];
  var abcHeaders = abcData[0];

  // Find column indexes in Input Sheet
  var clientIndex = inputHeaders.indexOf("Client_Name");
  var emailIndex = inputHeaders.indexOf("Email_Id");
  var assetTypeIndex = inputHeaders.indexOf("Asset_Type");
  var securityIndex = inputHeaders.indexOf("Security_Name");
  var tickerIndex = inputHeaders.indexOf("Ticker");
  var transactionIndex = inputHeaders.indexOf("Transaction_Type");
  var unitsIndex = inputHeaders.indexOf("Units");
  var dateIndex = inputHeaders.indexOf("Date");
  var currencyIndex = inputHeaders.indexOf("Currency");
  var reportingCurrencyIndex = inputHeaders.indexOf("Reporting_Currency");

  // Find column indexes in ABC Sheet
  var abcTickerIndex = abcHeaders.indexOf("Ticker");
  var dividendAmountIndex = abcHeaders.indexOf("Dividend Amount");
  var exDividendDateIndex = abcHeaders.indexOf("Ex-Dividend Date");

  if ([clientIndex, emailIndex, assetTypeIndex, securityIndex, tickerIndex, transactionIndex, unitsIndex, dateIndex, currencyIndex, reportingCurrencyIndex].includes(-1) ||
      [abcTickerIndex, dividendAmountIndex, exDividendDateIndex].includes(-1)) {
      logger.log("Error: Missing columns in sheets.");
    return;
  }

  // Prepare Dividend Summary Data
  var dividendSummary = [["Client_Name", "Email_Id", "Asset_Type", "Transaction_Type", "Security_Name", "Ticker", "Cash_Flow", "Date", "Currency", "Reporting_Currency", "Exchange Rate"]];

  // Iterate over each row in ABC Sheet (Multiple Ex-Dividend Dates for each Ticker)
  for (var i = 1; i < abcData.length; i++) {
    var ticker = abcData[i][abcTickerIndex];
    var dividendAmount = abcData[i][dividendAmountIndex];
    var exDividendDate = new Date(abcData[i][exDividendDateIndex]); // Convert to Date object

    if (!ticker || !dividendAmount || isNaN(exDividendDate)) continue; // Skip invalid rows

    // Track cumulative units per client before each Ex-Dividend Date
    var clientHoldings = {}; // Stores cumulative units for each client per ticker before each ex-dividend date

    for (var j = 1; j < inputData.length; j++) {
      var txnTicker = inputData[j][tickerIndex];
      var txnType = inputData[j][transactionIndex].toLowerCase();
      var txnUnits = inputData[j][unitsIndex];
      var txnDate = new Date(inputData[j][dateIndex]); // Convert to Date object
      var currency = inputData[j][currencyIndex];
      var reportingCurrency = inputData[j][reportingCurrencyIndex];

      if (txnTicker === ticker && txnDate < exDividendDate) { // Only consider transactions before Ex-Dividend Date
        var clientKey = inputData[j][clientIndex] + "|" + inputData[j][emailIndex]; // Unique client identifier

        if (!clientHoldings[clientKey]) {
          clientHoldings[clientKey] = {
            clientName: inputData[j][clientIndex],
            email: inputData[j][emailIndex],
            assetType: inputData[j][assetTypeIndex],
            securityName: inputData[j][securityIndex],
            cumulativeUnits: 0,
            currency: currency,
            reportingCurrency: reportingCurrency
          };
        }

        // Buy transactions add units, Sell transactions remove units
        if (txnType === "buy") {
          clientHoldings[clientKey].cumulativeUnits += txnUnits;
        } else if (txnType === "sell") {
          clientHoldings[clientKey].cumulativeUnits -= txnUnits;
        }
      }
    }

    // Process each client's holdings for the given ex-dividend date
    for (var clientKey in clientHoldings) {
      var clientData = clientHoldings[clientKey];

      if (clientData.cumulativeUnits > 0) { // Ensure a valid holding before Ex-Dividend Date
        // Get exchange rate
        var exchangeRate = 1;
        if (clientData.currency !== clientData.reportingCurrency) {
          exchangeRate = getHistoricalExchangeRate(clientData.currency, clientData.reportingCurrency, exDividendDate, 5); // Retries up to 5 days
        }

        var cashFlow = clientData.cumulativeUnits * dividendAmount * exchangeRate;

        dividendSummary.push([
          clientData.clientName,
          clientData.email,
          clientData.assetType,
          "dividend payout",
          clientData.securityName,
          ticker,
          cashFlow,
          exDividendDate, // Use Ex-Dividend Date in Dividend Summary
          clientData.currency,
          clientData.reportingCurrency,
          exchangeRate
        ]);
      }
    }
  }

  // Write results to a new sheet
  var summarySheet = ss.getSheetByName("Dividend Summary") || ss.insertSheet("Dividend Summary");
  summarySheet.clear(); // Clear previous data
  summarySheet.getRange(1, 1, dividendSummary.length, dividendSummary[0].length).setValues(dividendSummary);
  
  // **Apply Calibri font to entire sheet**
  summarySheet.getRange(1, 1, dividendSummary.length, dividendSummary[0].length).setFontFamily("Calibri");

  // **Make headers bold**
  summarySheet.getRange(1, 1, 1, dividendSummary[0].length).setFontWeight("bold");

  // **Auto-resize columns for better readability**
  summarySheet.autoResizeColumns(1, dividendSummary[0].length);

 
  Logger.log("✅ Dividend Summary has been generated successfully!");
}

/**
 * Fetches historical exchange rate from Google Finance.
 * If the exchange rate is missing, it will check previous days (up to retryDays attempts).
 */

function getHistoricalExchangeRate(fromCurrency, toCurrency, date, retryDays) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ✅ Check if the temporary sheet already exists
  var tempSheet = ss.getSheetByName("Exchange Rate Temp");

  if (!tempSheet) {
    tempSheet = ss.insertSheet("Exchange Rate Temp"); // Create if not exists
  } else {
    tempSheet.clear(); // If exists, clear old data
  }

  // Extract year, month, day from Date object
  var year = date.getFullYear();
  var month = date.getMonth() + 1; // JS months are 0-based, so add 1
  var day = date.getDate();

  // Set the end date as one day after the ex-dividend date
  var endDate = new Date(date);
  endDate.setDate(date.getDate() + 1);
  var endYear = endDate.getFullYear();
  var endMonth = endDate.getMonth() + 1; // Adjust for 0-based month
  var endDay = endDate.getDate();

  // Construct the formula with proper DATE arguments
  var formula = `=GOOGLEFINANCE("CURRENCY:${fromCurrency}${toCurrency}", "price", DATE(${year}, ${month}, ${day}), DATE(${endYear}, ${endMonth}, ${endDay}))`;
  tempSheet.getRange("A1").setFormula(formula);

  SpreadsheetApp.flush(); // Ensure formula calculates
  //Utilities.sleep(2000); // Allow data to load properly

  // Fetch exchange rate from the sheet
  var exchangeRate = tempSheet.getRange("B2").getValue(); 

  // If exchange rate is missing and retries are available, check previous day
  if ((!exchangeRate || exchangeRate === "") && retryDays > 0) {
    var previousDate = new Date(date);
    previousDate.setDate(previousDate.getDate() - 1); // Go back one day

    Logger.log(`Exchange rate missing for ${year}-${month}-${day}. Trying previous day: ${previousDate.toISOString().split("T")[0]}`);
    
    return getHistoricalExchangeRate(fromCurrency, toCurrency, previousDate, retryDays - 1);
  }

  // **Remove the temporary sheet**
  ss.deleteSheet(tempSheet);
  // Logger.log("🗑️ Temporary sheet 'Exchange Rate Temp' deleted.");

  return exchangeRate || 1; // Default to 1 if no data is available
}

