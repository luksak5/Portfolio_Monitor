function portfolioSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Input");
  if (!sourceSheet) {
    Logger.log("Error: Input sheet not found.");
    return;
  }
  const data = sourceSheet.getDataRange().getValues();

  // Read Dividend Summary sheet data
  const dividendSheet = ss.getSheetByName("Dividend Summary");
  if (!dividendSheet) {
    Logger.log("Error: Dividend Summary sheet not found.");
    return;
  }
  const dividendData = dividendSheet.getDataRange().getValues();
  const dividendHeaders = dividendData[0];

  // Create or clear the "Portfolio Summary" output sheet
  let outputSheet = ss.getSheetByName("Portfolio Summary");
  if (!outputSheet) {
    outputSheet = ss.insertSheet("Portfolio Summary");
  } else {
    outputSheet.clear();
  }

  const headers = [
    "Client Name", "Client ID", "Ticker", "Security Name", "Asset Class",
    "Net Purchase Value (Reporting Currency)", "Realized Gain/Loss (Reporting Currency)",
    "Current Value (Reporting Currency)", "Weightage (%)", "Dividend Paid (Reporting Currency)",
    "XIRR (Reporting Currency)", "Total Units", "Average Purchase Value"
  ];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headersRow = data[0];
  const clientNameIndex = headersRow.indexOf("Client_Name");
  const emailIdIndex = headersRow.indexOf("Email_Id");
  const tickerIndex = headersRow.indexOf("Ticker");
  const securityNameIndex = headersRow.indexOf("Security_Name");
  const assetClassIndex = headersRow.indexOf("Asset_Type");
  const transactionTypeIndex = headersRow.indexOf("Transaction_Type");
  const unitsIndex = headersRow.indexOf("Units");
  const priceIndex = headersRow.indexOf("Price");
  const currencyIndex = headersRow.indexOf("Currency");
  const reportingCurrencyIndex = headersRow.indexOf("Reporting_Currency");
  const exchangeRateIndex = headersRow.indexOf("Exchange Rate");
  const dateIndex = headersRow.indexOf("Date");

  if (
    clientNameIndex < 0 || emailIdIndex < 0 || tickerIndex < 0 ||
    securityNameIndex < 0 || assetClassIndex < 0 || transactionTypeIndex < 0 ||
    unitsIndex < 0 || priceIndex < 0 || currencyIndex < 0 || 
    reportingCurrencyIndex < 0 || exchangeRateIndex < 0 || dateIndex < 0
  ) {
    Logger.log("Error: One or more required columns are missing in the input data.");
    return;
  }

  const clientData = {};

  // Process Buy/Sell transactions from Input sheet
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const clientName = row[clientNameIndex];
    const emailId = row[emailIdIndex];
    const ticker = row[tickerIndex];
    const securityName = row[securityNameIndex];
    const assetClass = row[assetClassIndex];
    const transactionType = row[transactionTypeIndex].toLowerCase();
    const units = parseFloat(row[unitsIndex]);
    const price = parseFloat(row[priceIndex]);
    const currency = row[currencyIndex];
    const reportingCurrency = row[reportingCurrencyIndex];
    const exchangeRate = parseFloat(row[exchangeRateIndex]);
    const transactionDate = new Date(row[dateIndex]);

    if (!clientName || !emailId || !ticker || isNaN(exchangeRate) || isNaN(transactionDate.getTime())) {
      continue;
    }

    const clientKey = `${clientName}_${emailId}_${ticker}`;
    if (!clientData[clientKey]) {
      clientData[clientKey] = {
        clientName,
        emailId,
        ticker,
        securityName,
        assetClass,
        netPurchaseValueInReportingCurrency: 0,
        realizedGainLossInReportingCurrency: 0,
        currentValueInReportingCurrency: 0,
        dividendPaidInReportingCurrency: 0,
        cashFlows: [],
        fifoQueue: [],
        remainingUnits: 0,
        totalUnits: 0,
        averagePurchaseValue: 0,
        baseCurrency: currency,
        reportingCurrency: reportingCurrency
      };
    }

    const clientEntry = clientData[clientKey];

    if (transactionType === "buy") {
      clientEntry.fifoQueue.push({ units, price, exchangeRate });
      clientEntry.remainingUnits += units;
      clientEntry.totalUnits += units;
      clientEntry.cashFlows.push([-units * price * exchangeRate, transactionDate]);
    } else if (transactionType === "sell") {
      let remainingUnitsToSell = units;
      let realizedGain = 0;

      while (remainingUnitsToSell > 0 && clientEntry.fifoQueue.length > 0) {
        const firstBuy = clientEntry.fifoQueue[0];
        const sellUnits = Math.min(firstBuy.units, remainingUnitsToSell);
        const buyPrice = firstBuy.price * firstBuy.exchangeRate;
        const sellPrice = price * exchangeRate;
        realizedGain += sellUnits * (sellPrice - buyPrice);

        firstBuy.units -= sellUnits;
        remainingUnitsToSell -= sellUnits;

        if (firstBuy.units === 0) {
          clientEntry.fifoQueue.shift();
        }
      }

      clientEntry.realizedGainLossInReportingCurrency += realizedGain;
      clientEntry.remainingUnits -= units;
      clientEntry.totalUnits -= units;
      clientEntry.cashFlows.push([units * price * exchangeRate, transactionDate]);
    }
    let netPurchaseValue = 0;
    let totalUnitsForAverage = 0;
    clientEntry.fifoQueue.forEach(fifoEntry => {
      if (fifoEntry.units > 0) {
        netPurchaseValue += fifoEntry.units * fifoEntry.price * fifoEntry.exchangeRate;
        totalUnitsForAverage += fifoEntry.units;
      }
    });
    clientEntry.netPurchaseValueInReportingCurrency = netPurchaseValue;
    clientEntry.averagePurchaseValue = totalUnitsForAverage > 0
      ? (netPurchaseValue / totalUnitsForAverage)
      : 0;
  
  }

  // Process Dividends from Dividend Summary sheet
  const divClientNameIndex = dividendHeaders.indexOf("Client_Name");
  const divEmailIdIndex = dividendHeaders.indexOf("Email_Id");
  const divTickerIndex = dividendHeaders.indexOf("Ticker");
  const divDateIndex = dividendHeaders.indexOf("Date");
  const divCashFlowIndex = dividendHeaders.indexOf("Cash_Flow");

  if (divClientNameIndex < 0 || divEmailIdIndex < 0 || divTickerIndex < 0 || divDateIndex < 0 || divCashFlowIndex < 0) {
    Logger.log("Error: Required columns in Dividend Summary sheet are missing.");
    return;
  }

  for (let i = 1; i < dividendData.length; i++) {
    const row = dividendData[i];
    const clientName = row[divClientNameIndex];
    const emailId = row[divEmailIdIndex];
    const ticker = row[divTickerIndex];
    const dividendDate = new Date(row[divDateIndex]);
    const cashFlow = parseFloat(row[divCashFlowIndex]);

    if (!clientName || !emailId || !ticker || isNaN(dividendDate.getTime()) || isNaN(cashFlow)) {
      continue;
    }

    const clientKey = `${clientName}_${emailId}_${ticker}`;
    const clientEntry = clientData[clientKey];

    if (clientEntry) {
      clientEntry.dividendPaidInReportingCurrency += cashFlow;
      clientEntry.cashFlows.push([cashFlow, dividendDate]);
    } 
  
  }

  // Calculate current value and weightage
  const clientPortfolioValues = {};
  for (const clientKey in clientData) {
    const client = clientData[clientKey];
    const currentPrice = getLivePrice(client.ticker);
    const liveExchangeRate = getLiveExchangeRate(client.baseCurrency, client.reportingCurrency);

    if (!isNaN(currentPrice) && !isNaN(liveExchangeRate) && client.remainingUnits > 0) {
      const currentValue = client.remainingUnits * currentPrice * liveExchangeRate;
      client.currentValueInReportingCurrency = currentValue;
      client.cashFlows.push([currentValue, new Date()]);
    }

    const individualClientKey = `${client.clientName}_${client.emailId}`;
    if (!clientPortfolioValues[individualClientKey]) {
      clientPortfolioValues[individualClientKey] = 0;
    }
    clientPortfolioValues[individualClientKey] += client.currentValueInReportingCurrency || 0;
  }

  const portfolioData = [];
  for (const clientKey in clientData) {
    const client = clientData[clientKey];
    const xirr = calculateXIRR(client.cashFlows);

    const individualClientKey = `${client.clientName}_${client.emailId}`;
    const totalClientPortfolioValue = clientPortfolioValues[individualClientKey];
    const weightage = totalClientPortfolioValue
      ? (client.currentValueInReportingCurrency / totalClientPortfolioValue) * 100
      : 0;

    portfolioData.push([
      client.clientName,
      client.emailId,
      client.ticker,
      client.securityName,
      client.assetClass,
      client.netPurchaseValueInReportingCurrency.toFixed(2),
      client.realizedGainLossInReportingCurrency.toFixed(2),
      client.currentValueInReportingCurrency.toFixed(2),
      weightage.toFixed(2) + "%",
      client.dividendPaidInReportingCurrency.toFixed(2),
      xirr,
      client.totalUnits,
      client.averagePurchaseValue.toFixed(2)
    ]);
  }
  outputSheet.getRange(2, 1, portfolioData.length, portfolioData[0].length).setValues(portfolioData);
  outputSheet.getRange(1, 1, outputSheet.getLastRow(), outputSheet.getLastColumn()).setFontFamily("Calibri");

  // Generate Overall Portfolio Summary
  let overallSheet = ss.getSheetByName("Overall Portfolio Summary");
  if (!overallSheet) {
    overallSheet = ss.insertSheet("Overall Portfolio Summary");
  } else {
    overallSheet.clear();
  }

  const assetTypesSet = new Set();
  Object.values(clientData).forEach(client => {
    if (client.assetClass) {
      assetTypesSet.add(client.assetClass);
    }
  });
  const dynamicAssetTypes = Array.from(assetTypesSet);

  const overallHeaders = [
    "Client Name", "Email ID",
    ...dynamicAssetTypes.map(assetType => `${assetType} (%)`),
    "Portfolio Current Value(Reporting Currency)",
    "Total Net Purchase Value (Reporting Currency)",
    "Total Realized Gain/Loss (Reporting Currency)",
    "Total Dividend Payout (Reporting Currency)",
    "Overall Portfolio Gain/Loss (Reporting Currency)",
    "Portfolio XIRR (Reporting Currency)"
  ];
  overallSheet.getRange(1, 1, 1, overallHeaders.length).setValues([overallHeaders]);

  const overallClientData = {};
  for (const clientKey in clientData) {
    const client = clientData[clientKey];
    const overallKey = `${client.clientName}_${client.emailId}`;
    if (!overallClientData[overallKey]) {
      overallClientData[overallKey] = {
        clientName: client.clientName,
        emailId: client.emailId,
        assetTypeValues: {},
        portfolioCurrentValue: 0,
        totalNetPurchaseValue: 0,
        totalRealizedGainLoss: 0,
        totalDividendPayout: 0,
        cashFlows: []
      };
    }

    const overallEntry = overallClientData[overallKey];
    overallEntry.totalNetPurchaseValue += client.netPurchaseValueInReportingCurrency;
    overallEntry.totalRealizedGainLoss += client.realizedGainLossInReportingCurrency;
    overallEntry.totalDividendPayout += client.dividendPaidInReportingCurrency;
    overallEntry.portfolioCurrentValue += client.currentValueInReportingCurrency;

    // Dynamically handle asset type values
    overallEntry.assetTypeValues[client.assetClass] =
      (overallEntry.assetTypeValues[client.assetClass] || 0) + client.currentValueInReportingCurrency;

    overallEntry.cashFlows = overallEntry.cashFlows.concat(client.cashFlows);
  }

  const overallOutput = [];
  for (const key in overallClientData) {
    const entry = overallClientData[key];
    const overallGainLoss =
      entry.portfolioCurrentValue - entry.totalNetPurchaseValue +
      entry.totalRealizedGainLoss + entry.totalDividendPayout;
    const overallXIRR = calculateXIRR(entry.cashFlows);
    const totalPortfolioValue = Object.values(entry.assetTypeValues).reduce((a, b) => a + b, 0);

    // Calculate percentages for all dynamic asset types
    const assetTypePercentages = dynamicAssetTypes.map(assetType => {
      return totalPortfolioValue
        ? ((entry.assetTypeValues[assetType] || 0) / totalPortfolioValue) * 100
        : 0;
    });

    overallOutput.push([
      entry.clientName,
      entry.emailId,
      ...assetTypePercentages.map(percent => percent.toFixed(2) + "%"),
      entry.portfolioCurrentValue.toFixed(2),
      entry.totalNetPurchaseValue.toFixed(2),
      entry.totalRealizedGainLoss.toFixed(2),
      entry.totalDividendPayout.toFixed(2),
      overallGainLoss.toFixed(2),
      overallXIRR
    ]);
  }

  overallSheet.getRange(2, 1, overallOutput.length, overallOutput[0].length).setValues(overallOutput);
  overallSheet.getRange(1, 1, overallSheet.getLastRow(), overallSheet.getLastColumn()).setFontFamily("Calibri");
  SpreadsheetApp.flush();
}


// Function to calculate XIRR using Newton's method
function calculateXIRR(cashFlows) {
  const maxIterations = 100;
  const tolerance = 1e-6;
  let x0 = 0.1; // Initial guess for XIRR (10%)
  let iteration = 0;

  // Adjust the initial guess dynamically based on cash flow trends
  const totalPositive = cashFlows.reduce((sum, [amount]) => amount > 0 ? sum + amount : sum, 0);
  const totalNegative = cashFlows.reduce((sum, [amount]) => amount < 0 ? sum + amount : sum, 0);
  if (Math.abs(totalNegative) > totalPositive) {
    x0 = -0.9; // Set an initial guess to a more negative value
  } else if (totalPositive > Math.abs(totalNegative)) {
    x0 = 0.9; // Set an initial guess to a more positive value
  }

  while (iteration < maxIterations) {
    let fValue = 0;
    let fPrimeValue = 0;

    cashFlows.forEach(([amount, date]) => {
      const t = (date - cashFlows[0][1]) / (1000 * 60 * 60 * 24 * 365); // Time difference in years
      const discountFactor = Math.pow(1 + x0, t);

      fValue += amount / discountFactor;
      fPrimeValue -= (amount * t) / (discountFactor * (1 + x0));
    });

    if (Math.abs(fPrimeValue) < tolerance) {
      break; // Avoid division by near-zero derivative
    }

    const newX0 = x0 - fValue / fPrimeValue;

    if (Math.abs(newX0 - x0) < tolerance) {
      return (newX0 * 100).toFixed(2) + "%"; // Return XIRR as a percentage
    }

    x0 = newX0;
    iteration++;
  }

  // Retry with a broader range of initial guesses if it fails
  if (iteration >= maxIterations) {
    const retryGuesses = [-0.99, -0.5, 0.1, 0.5, 0.99];
    for (const guess of retryGuesses) {
      x0 = guess;
      iteration = 0;
      while (iteration < maxIterations) {
        let fValue = 0;
        let fPrimeValue = 0;

        cashFlows.forEach(([amount, date]) => {
          const t = (date - cashFlows[0][1]) / (1000 * 60 * 60 * 24 * 365);
          const discountFactor = Math.pow(1 + x0, t);

          fValue += amount / discountFactor;
          fPrimeValue -= (amount * t) / (discountFactor * (1 + x0));
        });

        if (Math.abs(fPrimeValue) < tolerance) {
          break;
        }

        const newX0 = x0 - fValue / fPrimeValue;

        if (Math.abs(newX0 - x0) < tolerance) {
          return (newX0 * 100).toFixed(2) + "%";
        }

        x0 = newX0;
        iteration++;
      }
    }
  }

  return "XIRR calculation failed"; // If XIRR cannot be calculated within max iterations and retries
}


// Function to get live price for a ticker with retry mechanism and delay
function getLivePrice(ticker) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName("Input");
    if (!sourceSheet) {
      throw new Error("Input sheet not found.");
    }

    const data = sourceSheet.getDataRange().getValues();
    const headersRow = data[0];
    const tickerIndex = headersRow.indexOf("Ticker");
    const livePriceIndex = headersRow.indexOf("Live_Price");

    if (tickerIndex === -1 || livePriceIndex === -1) {
      throw new Error("Ticker or Live_Price column not found in the Input sheet.");
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][tickerIndex] === ticker) {
        const livePrice = parseFloat(data[i][livePriceIndex]);
        if (!isNaN(livePrice)) {
          return livePrice; // Return the live price if found and valid
        } else {
          throw new Error(`Invalid or missing price for ticker: ${ticker}`);
        }
      }
    }

    throw new Error(`Ticker ${ticker} not found in the Input sheet.`);
  } catch (error) {
    Logger.log(`Error fetching live price for ${ticker}: ${error.message}`);
    return NaN; // Return NaN if any error occurs
  }
}

// Function to get live exchange rate between two currencies
function getLiveExchangeRate(baseCurrency, targetCurrency) {
  try {
    // If base and target currencies are the same, return exchange rate as 1
    if (baseCurrency === targetCurrency) {
      //Logger.log(`Base currency and target currency are the same: ${baseCurrency}. Returning exchange rate as 1.`);
      return 1;
    }

    // Fetch the live exchange rate using GOOGLEFINANCE
    const tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    tempSheet.getRange("A1").setFormula(`=GOOGLEFINANCE("${baseCurrency}${targetCurrency}")`);
    SpreadsheetApp.flush();
    const exchangeRate = tempSheet.getRange("A1").getValue();
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tempSheet);

    if (!exchangeRate || isNaN(exchangeRate)) {
      throw new Error(`Exchange rate not found for ${baseCurrency} to ${targetCurrency}`);
    }

    return exchangeRate;
  } catch (error) {
    Logger.log(`Error fetching exchange rate for ${baseCurrency} to ${targetCurrency}: ${error.message}`);
    return NaN;
  }
}
