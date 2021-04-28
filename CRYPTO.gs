/**
 * This is a copy of: https://github.com/rushashell/crypto-sheets/blob/main/CRYPTO.gs
 * Forked from: https://github.com/kbouchard/crypto-sheets
 * First version was created by kbouchard. Historical functionality with wallet values added by ELiT3
 * Requirements: 
 *   - Live price: nothing (requested on demand from coinstats API)
 *   - Historical price: FREE Cryptocompare account with API key: https://min-api.cryptocompare.com/
*/
const defaultWalletSheetName = "Wallet value";

/**
 * @OnlyCurrentDoc
 */
function onOpen() 
{
  var ui = SpreadsheetApp.getUi();

  var app = SpreadsheetApp.getActiveSpreadsheet();
  var walletValueSheet = app.getSheetByName(sheetName);

  var newMenu = ui.createMenu('CRYPTO')
      .addItem('💲 Refresh prices', 'cryptoRefresh')
      .addItem('🔃 Fetch API data', 'cryptoFetchData');

  if (walletValueSheet)
  {
    newMenu.addItem('🗓️ Add wallet date columns', 'addWalletDateColumns')
  }
      
  newMenu.addSeparator()
    .addItem('🌐 Configure exchange', 'showSelectExchange')
    .addItem('🔑 Configure CryptoCompare API key', 'showSetApiKey')
    .addSeparator()
    .addItem('❓ How to auto-refresh rates', 'ShowRefreshInfo')
    .addToUi();
}

/**
 * @OnlyCurrentDoc
 */
function ShowRefreshInfo() 
{
  var ui = SpreadsheetApp.getUi()
  ui.alert(
    "How to refresh rates",
    'Coming soon...',
    ui.ButtonSet.OK
  )
}

/**
 * @OnlyCurrentDoc
 */
function showSelectExchange() 
{
  var ui = SpreadsheetApp.getUi();
  var userProperties = PropertiesService.getUserProperties();
  var userExchange = userProperties.getProperty("CRYPTO_EXCHANGE")

  var result = ui.prompt(
    'Exchange setting',
    `Set the exchange you want to use for prices (e.g "binance", "kraken", ...)\nIt is currently set to "${userExchange}"`,
    ui.ButtonSet.OK_CANCEL
  );
  
  var button = result.getSelectedButton();
  var user_input = result.getResponseText().replace(/\s+/g, '');
  if (button == ui.Button.OK) {
    if (user_input) {
      userProperties.setProperty("CRYPTO_EXCHANGE", user_input);
      ui.alert(
        'Exchange successfully saved',
        'If it does not work right away, please try to refresh manually.',
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * @OnlyCurrentDoc
 */
function showSetApiKey() 
{
  var ui = SpreadsheetApp.getUi();
  var userProperties = PropertiesService.getUserProperties();
  var currentKey = userProperties.getProperty("CRYPTO_CRYPTOCOMPARE_API");

  var keyText = currentKey != null ? "It is currently set" : "No current key set.";

  var result = ui.prompt(
    'CryptoCompare API',
    `Define the CryptoCompare API key.\n${keyText}`,
    ui.ButtonSet.OK_CANCEL
  );
  
  var button = result.getSelectedButton();
  var user_input = result.getResponseText().replace(/\s+/g, '');
  if (button == ui.Button.OK) {
    if (user_input) {
      userProperties.setProperty("CRYPTO_CRYPTOCOMPARE_API", user_input);
      ui.alert(
        'CryptoCompare API key successfully saved',
        'If it does not work right away, please check the log.',
        ui.ButtonSet.OK
      );
    }
  }
}

/*
  refresh()
  TRIGGER
    Borrowed form:
    https://tanaikech.github.io/2019/10/28/automatic-recalculation-of-custom-function-on-spreadsheet-part-2/
*/
function cryptoRefresh() 
{
  const customFunctions = ["CRYPTO_PRICE", "CRYPTO_PRICE_HISTORICAL"]; // Please set the function names here.

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var temp = Utilities.getUuid();
  var loadingStr = 'Refreshing coin...';

  customFunctions.forEach(function(funcName) {
    ss.createTextFinder("=" + funcName)
      .matchFormulaText(true)
      .replaceAllWith(loadingStr);
    ss.createTextFinder(loadingStr)
      .matchFormulaText(true)
      .replaceAllWith("=" + funcName);
  });
}

function cryptoDeleteData() 
{
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('COIN_DATA');
  scriptProperties.deleteProperty('CRYPTO_COIN_DATA');
  scriptProperties.deleteProperty('CRYPTO_CRYPTOCOMPARE_API');
  scriptProperties.deleteProperty('CRYPTO_CRYPTOCOMPARE_CACHE');
}

/*
  fetchAPIData()
  TRIGGER
*/
function cryptoFetchData() 
{
  var userProperties = PropertiesService.getUserProperties();
  var exchange = userProperties.getProperty("CRYPTO_EXCHANGE") || "binance";
  var url=`https://api.coinstats.app/public/v1/coins?skip=0&limit=1000&exchange=${exchange}&currency=USDT`;
  var response = UrlFetchApp.fetch(url); // get feed
  var jsonData = JSON.parse(response.getContentText());

  // To make the retrieving easier, key them by their ticker
  const data = {};
  Object.values(jsonData.coins).forEach((coin) => {
    data[coin.symbol] = {
      price: coin.price,
      priceBtc: coin.priceBtc,
    };
  });

  console.log('cryptoFetchData().CRYPTO_EXCHANGE', exchange);
  console.log('cryptoFetchData().CRYPTO_COIN_DATA', data);
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('CRYPTO_COIN_DATA', JSON.stringify(data));
}

function cryptoGetCachedHistoricalPrice(input, currency, date)
{
  var scriptProperties = PropertiesService.getScriptProperties();
  const cachedDataJson = scriptProperties.getProperty('CRYPTO_CRYPTOCOMPARE_CACHE');
  var cachedData = JSON.parse(cachedDataJson) || {};
  if (!cachedData) return null;

  var dateObject = new Date(date);
  var timestamp = dateObject.getTime() / 1000;
  var cachedInput = cachedData[input];
  var cachedCurrency = cachedInput != null ? cachedInput[currency] : null;
  var cachedPrice = cachedCurrency != null ? cachedCurrency[timestamp] : null;
  if (cachedPrice != null)
  {
     // Found in cache. Return
     console.log(`Found price for ${input}, timestamp ${date} in cache.`);
     return cachedPrice;
  }

  console.log(`Price was not found in cache: ${cachedInput}, ${cachedCurrency}. ${cachedPrice}.`);
  return null;
}

function cryptoUpdateCachedHistoricalPrice(input, currency, date, price)
{
  var scriptProperties = PropertiesService.getScriptProperties();
  const cachedDataJson = scriptProperties.getProperty('CRYPTO_CRYPTOCOMPARE_CACHE');
  var cachedData = JSON.parse(cachedDataJson) || {};
  if (!cachedData) return null;

  var dateObject = new Date(date);
  var timestamp = dateObject.getTime() / 1000;
  var cachedInput = cachedData[input];
  var cachedCurrency = cachedInput ? cachedInput[currency] : null;

  if (!cachedInput) cachedData[input] = {};
  if (!cachedCurrency) cachedData[input][currency] = {};
  cachedData[input][currency][timestamp] = price;
  
  console.log(`Writing cache: ${input}, ${currency}, ${timestamp}...`)
  // I'm not sure if below line is in trouble when function is called async. If so, there should be a lock mechanism (not ideal).
  scriptProperties.setProperty('CRYPTO_CRYPTOCOMPARE_CACHE', JSON.stringify(cachedData));
}

function cryptoFetchDataHistorical(input, currency, date)
{
  var userProperties = PropertiesService.getUserProperties();
  const exchange = userProperties.getProperty("CRYPTO2_EXCHANGE") || "CCCAGG";
  const currentKey = userProperties.getProperty("CRYPTO_CRYPTOCOMPARE_API");
  
  if (currentKey == null)
  {
    throw new Error("No CryptoCompare API key is set.");
  }

  var dateObject = new Date(date);
  var timestamp = dateObject.getTime() / 1000;
  var url=`https://min-api.cryptocompare.com/data/v2/histoday?fsym=${input}&tsym=${currency}&e=${exchange}&toTs=${timestamp}&limit=1&api_key=${currentKey}`
  var response = UrlFetchApp.fetch(url); 
  var jsonData = JSON.parse(response.getContentText());
  if (jsonData == null)
  {
    throw new Error(`Invalid response from API: ${response}`);
  }

  if (jsonData.Data == null || jsonData.Data.Data == null || jsonData.Data.Data.length == 0)
  {
    return "-1";
  }

  var price = jsonData.Data.Data[1].open;  
  return price;
}

/**
 * Get a ticker price from a Range of tickers historical.
 * @param {ticker(s)} range.
 * @param {date} the historical date.
 * @param {convertCurrency} only USD, EUR, USDT, BTC and ETH are supported for now.
 * @return The value of the ticker for the given currency at the given date.
 * @customfunction
 */
function CRYPTO_PRICE_HISTORICAL(input = "BTC", date = "2017-01-01", convertCurrency = "USDT")
{
  if (input.map) 
  {
    return input.map((ticker) => CRYPTO_PRICE_HISTORICAL(ticker, date, convertCurrency));
  }
  else 
  {
    if (!["USD", "EUR", "USDT", "BTC", "ETH"].includes(convertCurrency)) throw new Error("The currency param must be USD, EUR, USDT, BTC or ETH, Default to USDT.");

    var dateObject = new Date(date);
    if (dateObject == null)
    {
      throw new Error("Invalid date given.");
    }

    // check if is cached
    var price = cryptoGetCachedHistoricalPrice(input, convertCurrency, date);
    if (price != null)
    {
      return price;
    }

    // Get price from API, save it to cache and return.
    var fetchedPrice = cryptoFetchDataHistorical(input, convertCurrency, date);
    cryptoUpdateCachedHistoricalPrice(input, convertCurrency, date, fetchedPrice);
    return fetchedPrice;
  }
}

/**
 * Get a ticker price from a Range of tickers.
 * @param {ticker(s)} range.
 * @param {currency} only USDT and BTC are supported for now.
 * @return The value of the ticker for the given currency.
 * @customfunction
 */
function CRYPTO_PRICE(input = "BTC", currency = "USDT") 
{   
  if (input.map) 
  {
    return input.map((ticker) => CRYPTO_PRICE(ticker, currency));
  }
  else 
  {
    var scriptProperties = PropertiesService.getScriptProperties();
    const coinDataJSON = scriptProperties.getProperty('CRYPTO_COIN_DATA');
    let coinData = JSON.parse(coinDataJSON);

    // DEBUG mode
    if (!coinData) coinData = {"BTC": {"price": "1337", priceBtc: "1"}};
    
    console.log('currency', currency);
    if (!["USDT", "BTC"].includes(currency)) throw new Error("The currency param must be USDT or BTC, Default to USDT.");

    const priceProperty = currency === "USDT" ? "price" : "priceBtc";

    const price = coinData[input] && coinData[input][priceProperty] ? coinData[input][priceProperty] : 'n/a';

    return price;
  }
}

function addWalletDateColumns(sheetName = "")
{
  if (sheetName == "")
  {
    sheetName = defaultWalletSheetName;
  }

  var app = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = app.getSheetByName(sheetName);
  if (!sheet)
  {
    throw new error("Could not find sheet '" + sheetName + "'");
  }

  // Define positions
  const columnInsertBefore = 8;
  const columnsNeeded = 3;
  const newColumnStart = columnInsertBefore+1;
  const previousDataColumnStart = 9 + columnsNeeded;
  const newDateColumnIndex = newColumnStart+1;
  const newDateRowIndex = 4;

  // Add new columns needed for new data
  sheet.insertColumns(columnInsertBefore, columnsNeeded);

  // We make sure we take the same width as previous
  for(var x=newColumnStart; x < (newColumnStart+columnsNeeded); x++)
  {
    var newWidth = sheet.getColumnWidth(x+columnsNeeded);
    sheet.setColumnWidth(x, newWidth);
  }  

  // Get the last data from the sheet, take new position into account.
  var previousData = sheet.getRange(1, previousDataColumnStart, 1000, columnsNeeded);

  // Paste previous data to sheet
  previousData.copyTo(sheet.getRange(1, newColumnStart)); //, newColumnEnd, 1, 1000));

  // Change the date of the new columns
  var today = new Date();
  var day = String(today.getDate()).padStart(2, '0');
  var month = String(today.getMonth() + 1).padStart(2, '0'); 
  var year = today.getFullYear();
  var newDateValue = `="${year}-${month}-${day}"`;
  sheet.getRange(newDateRowIndex, newDateColumnIndex).setFormula(newDateValue);
}