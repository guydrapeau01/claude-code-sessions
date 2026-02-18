/***** GLOBAL CONFIG *****/
var ALPHAVANTAGE_API_KEY = 'X7L7XAQ31K35QFG3'; // <-- Your Alpha Vantage key
var FMP_API_KEY = '040fQZIXNugGMLHUmPSKQtWzRQWcEVdc'; // Financial Modeling Prep key
var USE_FMP_BACKUP = true; // Set to true to enable FMP fallback when AV hits rate limit
var DEFAULT_GROWTH = 0.05; // 5% default growth for Graham simplified formula

/***** SILENT MODE (for Messenger triggers) *****/
var SILENT_MODE = false;

function safeAlert_(message) {
  if (SILENT_MODE) {
    Logger.log('Silent mode alert: ' + message);
  } else {
    try {
      SpreadsheetApp.getUi().alert(message);
    } catch (e) {
      Logger.log('Alert (UI unavailable): ' + message);
    }
  }
}

/***** DCF SHEET CELLS (Row 2) *****/
var TICKER_CELL = 'A2';
var DISCOUNT_CELL = 'B2';
var YEARS_CELL = 'D2';
var TERM_GROWTH_CELL = 'G2';
var GROWTH_CELL = 'C2';
var FCF_CELL = 'E2';
var SHARES_CELL = 'F2';
var PRICE_CELL = 'H2';
var IV_CELL = 'I2';
var MOS_CELL = 'J2';
var STATUS_CELL = 'Q2';
var TS_CELL = 'S2';

// EPS & PE cells (info only)
var EPS_CELL = 'K2';
var PE_CELL = 'L2';

// Default growth used ONLY when a proper growth cannot be computed
var DEFAULT_GROWTH = 0.05; // 5% CAGR

/***** HELPERS (shared by DCF + Agent) *****/
function sleepMs_(ms){ Utilities.sleep(ms); }
function toNum_(v){ var n = Number(v); return (typeof n === 'number' && isFinite(n)) ? n : null; }
function isPos_(n){ return typeof n === 'number' && isFinite(n) && n > 0; }
function stripExchangePrefix_(s){
  s = String(s || '').trim();
  if (!s) return s;
  var parts = s.split(':');            // remove "EXCH:TICKER" prefixes
  if (parts.length === 2) return parts[1].trim();
  return s;
}

/***** Throttle, caching (CacheService), and backoff *****/
function avThrottle_(){ 
  sleepMs_(15000); // 15 seconds between calls = 4 calls/minute (safer than 5/min)
}
function cacheGet_(key){
  try { var cache = CacheService.getDocumentCache(); var val = cache.get(key); return val || null; }
  catch(e){ return null; }
}
function cacheSet_(key, body){
  try { var cache = CacheService.getDocumentCache(); cache.put(key, body); } catch(e){}
}
function httpGetAV_(url){
  avThrottle_();
  var options = { 
    muteHttpExceptions: true, 
    followRedirects: true,
    timeout: 30 // 30 second timeout (default is 60 but we want it explicit)
  };
  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  var js = null;
  try { js = JSON.parse(body); } catch(e){}
  if (js && (js['Note'] || js['Information'])) { // AV throttling/information messages
    return { status: 429, body: body, json: js, throttle: true };
  }
  return { status: code, body: body, json: js };
}
function httpGetAVCached_(cacheKey, url){
  var cached = cacheGet_(cacheKey);
  if (cached){
    try { return { status: 200, body: cached, json: JSON.parse(cached) }; } catch(e){}
  }
  var r = httpGetAV_(url);
  if (r.status === 200 && r.body) cacheSet_(cacheKey, r.body);
  return r;
}
function httpGetAVWithRetryCached_(cacheKey, url, maxRetries){
  var attempt = 0, delay = 2000;
  while (attempt <= (maxRetries || 2)){ // Default to 2 retries max
    var r = httpGetAVCached_(cacheKey, url);
    
    // SUCCESS - return immediately
    if (r.status === 200 && r.json && !r.throttle) {
      Logger.log('Success for ' + cacheKey);
      return r;
    }
    
    // RATE LIMIT - DO NOT RETRY, just return the error
    if (r.throttle || r.status === 429) {
      Logger.log('Rate limit hit for ' + cacheKey + ' - NOT retrying (would waste quota)');
      return r; // Return error immediately, don't retry
    }
    
    // SERVER ERROR (500s) - Retry makes sense here
    if (r.status >= 500 && r.status < 600) {
      Logger.log('Server error for ' + cacheKey + ' (HTTP ' + r.status + ') - retry attempt ' + attempt);
      sleepMs_(delay);
      delay = Math.min(delay * 2, 10 * 1000); // Max 10 second delay
      attempt++;
      continue;
    }
    
    // OTHER ERROR (400s, network, etc) - Return immediately, don't retry
    Logger.log('Error for ' + cacheKey + ' (HTTP ' + r.status + ') - NOT retrying');
    return r;
  }
  
  Logger.log('Retries exhausted for ' + cacheKey);
  return { status: 503, body: 'Retries exhausted', json: null };
}

/***** Alpha Vantage fetchers *****/
// (Optional) Symbol search pre-check — DISABLED to save a call and reduce noise
function avSymbolExists_(symbol){
  var url = 'https://www.alphavantage.co/query?function=SYMBOL_SEARCH&keywords='
            + encodeURIComponent(symbol) + '&apikey=' + ALPHAVANTAGE_API_KEY;
  var r = httpGetAVWithRetryCached_('AV_SYM_' + symbol, url, 1);
  if (r.status !== 200 || !r.json || !Array.isArray(r.json.bestMatches)) return false;
  return r.json.bestMatches.some(function(m){
    return String(m['1. symbol']||'').toUpperCase() === String(symbol||'').toUpperCase();
  });
}

// Price: GLOBAL_QUOTE → fallback TIME_SERIES_DAILY (last close)
function avFetchPrice_(symbol){
  var urlQ = 'https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol='
             + encodeURIComponent(symbol) + '&apikey=' + ALPHAVANTAGE_API_KEY;
  var r1 = httpGetAVWithRetryCached_('AV_PQ_' + symbol, urlQ, 1);
  if (r1.status === 200 && r1.json && r1.json['Global Quote']){
    var p = toNum_(r1.json['Global Quote']['05. price']);
    if (isPos_(p)) return p;
  }
  var urlD = 'https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol='
             + encodeURIComponent(symbol) + '&apikey=' + ALPHAVANTAGE_API_KEY;
  var r2 = httpGetAVWithRetryCached_('AV_PD_' + symbol, urlD, 1);
  if (r2.status === 200 && r2.json && r2.json['Time Series (Daily)']){
    var ts = r2.json['Time Series (Daily)'];
    var dates = Object.keys(ts).sort().reverse();
    if (dates.length){
      var last = ts[dates[0]];
      var close = toNum_(last['4. close']);
      if (isPos_(close)) return close;
    }
  }
  return null;
}

// Overview: SharesOutstanding / MarketCap / EPS / PEs
function avFetchOverview_(symbol){
  var url = 'https://www.alphavantage.co/query?function=OVERVIEW&symbol='
            + encodeURIComponent(symbol) + '&apikey=' + ALPHAVANTAGE_API_KEY;
  var r = httpGetAVWithRetryCached_('AV_OV_' + symbol, url, 1);
  if (r.status !== 200 || !r.json) return null;
  var so = toNum_(r.json['SharesOutstanding']);
  var mkt = toNum_(r.json['MarketCapitalization']);
  var eps = toNum_(r.json['EPS']);              // annual EPS
  var epsT = toNum_(r.json['DilutedEPSTTM']);   // TTM EPS
  var pe   = toNum_(r.json['PERatio']);         // PE (may be missing)
  var peTr = toNum_(r.json['TrailingPE']);      // trailing PE (sometimes present)
  var peF  = toNum_(r.json['ForwardPE']);       // forward PE (sometimes present)
  var bvps = toNum_(r.json['BookValue']);       // Book Value Per Share
  return {
    shares: isPos_(so) ? so : null,
    marketCap: isPos_(mkt) ? mkt : null,
    eps: (eps !== null ? eps : (epsT !== null ? epsT : null)),
    epsTTM: (epsT !== null ? epsT : null),
    pe: (pe !== null ? pe : (peTr !== null ? peTr : null)),
    peForward: (peF !== null ? peF : null),
    bookValuePerShare: (bvps !== null && isFinite(bvps)) ? bvps : null
  };
}

// Current FCF: annual (OCF - |CapEx|) preferred → TTM fallback → OCF proxy last resort
function avFetchFCF_(symbol){
  var url = 'https://www.alphavantage.co/query?function=CASH_FLOW&symbol='
            + encodeURIComponent(symbol) + '&apikey=' + ALPHAVANTAGE_API_KEY;
  var r = httpGetAVWithRetryCached_('AV_CF_' + symbol, url, 1);
  if (r.status !== 200 || !r.json) return null;
  var annual = r.json.annualReports;
  var quarterly = r.json.quarterlyReports;

  // Annual first (newest)
  if (Array.isArray(annual) && annual.length){
    var a = annual[0];
    var ocf = toNum_(a.netCashProvidedByOperatingActivities)
          || toNum_(a.operatingCashflow)
          || toNum_(a.operatingCashFlow);
    var capex = toNum_(a.capitalExpenditures)
             || toNum_(a.investmentsInPropertyPlantAndEquipment);
    if (ocf !== null && capex !== null) return ocf - Math.abs(capex);
    var fcf = toNum_(a.freeCashFlow);
    if (fcf !== null) return fcf;
    if (ocf !== null) return ocf; // proxy when CapEx not exposed
  }

  // Quarterly TTM
  if (Array.isArray(quarterly) && quarterly.length >= 4){
    var ocfSum = 0, capexSum = 0, haveOcf = true, haveCapex = true;
    for (var i=0; i<4; i++){
      var q = quarterly[i];
      var ocfq = toNum_(q.netCashProvidedByOperatingActivities)
              || toNum_(q.operatingCashflow)
              || toNum_(q.operatingCashFlow);
      var capexq = toNum_(q.capitalExpenditures)
                || toNum_(q.investmentsInPropertyPlantAndEquipment);
      if (ocfq === null) haveOcf = false; else ocfSum += ocfq;
      if (capexq === null) haveCapex = false; else capexSum += Math.abs(capexq);
    }
    if (haveOcf && haveCapex) return ocfSum - capexSum;
    if (haveOcf && !haveCapex) return ocfSum;
    var fcfSum = 0, ok = true;
    for (var j=0; j<4; j++){
      var fcfq = toNum_(quarterly[j].freeCashFlow);
      if (fcfq === null){ ok = false; break; }
      fcfSum += fcfq;
    }
    if (ok) return fcfSum;
  }
  return null;
}

/**
 * Growth from FCF: Annual CAGR (up to 5 most-recent points) → fallback TTM YoY.
 */
function avFetchGrowthFromFCF_(symbol){
  var url = 'https://www.alphavantage.co/query?function=CASH_FLOW&symbol='
            + encodeURIComponent(symbol) + '&apikey=' + ALPHAVANTAGE_API_KEY;
  var r = httpGetAVWithRetryCached_('AV_CF_' + symbol, url, 1);
  if (r.status !== 200 || !r.json) return null;

  var annual = r.json.annualReports;
  if (Array.isArray(annual) && annual.length >= 2){
    var series = [];
    for (var i=0; i<annual.length && series.length<5; i++){
      var a = annual[i];
      var ocf = toNum_(a.netCashProvidedByOperatingActivities)
             || toNum_(a.operatingCashflow)
             || toNum_(a.operatingCashFlow);
      var capex = toNum_(a.capitalExpenditures)
               || toNum_(a.investmentsInPropertyPlantAndEquipment);
      var fcf = null;
      if (ocf !== null && capex !== null) fcf = ocf - Math.abs(capex);
      else {
        fcf = toNum_(a.freeCashFlow);
        if (fcf === null && ocf !== null) fcf = ocf; // last-resort proxy
      }
      if (fcf !== null && isFinite(fcf)) series.push(fcf);
    }
    if (series.length >= 2){
      var endVal = series[0], startVal = series[series.length - 1];
      if (isPos_(endVal) && isPos_(startVal)){
        var years = series.length - 1;
        return Math.pow(endVal / startVal, 1/years) - 1;
      }
    }
  }

  // Fallback: TTM YoY (sum last 4 vs previous 4)
  var quarterly = r.json.quarterlyReports;
  if (Array.isArray(quarterly) && quarterly.length >= 8){
    var ttmCurr = 0, ttmPrev = 0;
    var okCurr = true, okPrev = true;
    for (var k=0; k<4; k++){
      var q = quarterly[k];
      var ocfq = toNum_(q.netCashProvidedByOperatingActivities)
              || toNum_(q.operatingCashflow)
              || toNum_(q.operatingCashFlow);
      var capexq = toNum_(q.capitalExpenditures)
                || toNum_(q.investmentsInPropertyPlantAndEquipment);
      var fcfq = null;
      if (ocfq !== null && capexq !== null) fcfq = ocfq - Math.abs(capexq);
      else {
        fcfq = toNum_(q.freeCashFlow);
        if (fcfq === null && ocfq !== null) fcfq = ocfq;
      }
      if (fcfq === null){ okCurr = false; break; }
      ttmCurr += fcfq;
    }
    for (var m=4; m<8; m++){
      var q2 = quarterly[m];
      var ocfq2 = toNum_(q2.netCashProvidedByOperatingActivities)
               || toNum_(q2.operatingCashflow)
               || toNum_(q2.operatingCashFlow);
      var capexq2 = toNum_(q2.capitalExpenditures)
                 || toNum_(q2.investmentsInPropertyPlantAndEquipment);
      var fcfq2 = null;
      if (ocfq2 !== null && capexq2 !== null) fcfq2 = ocfq2 - Math.abs(capexq2);
      else {
        fcfq2 = toNum_(q2.freeCashFlow);
        if (fcfq2 === null && ocfq2 !== null) fcfq2 = ocfq2;
      }
      if (fcfq2 === null){ okPrev = false; break; }
      ttmPrev += fcfq2;
    }
    if (okCurr && okPrev && isPos_(ttmPrev)) return (ttmCurr / ttmPrev) - 1;
  }
  return null;
}

/***** FINANCIAL MODELING PREP (FMP) BACKUP API *****/

/**
 * FMP: Fetch current stock price
 */
function fmpFetchPrice_(symbol) {
  if (!USE_FMP_BACKUP) {
    Logger.log('FMP backup is disabled (USE_FMP_BACKUP = false)');
    return null;
  }
  
  if (!FMP_API_KEY || FMP_API_KEY === 'YOUR_FMP_KEY_HERE') {
    Logger.log('FMP API key not configured - skipping FMP backup');
    return null;
  }
  
  var url = 'https://financialmodelingprep.com/api/v3/quote/' + 
            encodeURIComponent(symbol) + '?apikey=' + FMP_API_KEY;
  
  try {
    Logger.log('Fetching FMP price for ' + symbol + '...');
    var options = {
      muteHttpExceptions: true,
      timeout: 10 // 10 second timeout
    };
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    Logger.log('FMP price response code: ' + code);
    
    if (code !== 200) {
      Logger.log('FMP price failed - HTTP ' + code + ': ' + response.getContentText().substring(0, 100));
      return null;
    }
    
    var data = JSON.parse(response.getContentText());
    
    if (Array.isArray(data) && data.length > 0) {
      var price = toNum_(data[0].price);
      if (isPos_(price)) {
        Logger.log('FMP price SUCCESS for ' + symbol + ': ' + price);
        return price;
      }
    }
    
    Logger.log('FMP price: no valid data in response');
  } catch (e) {
    Logger.log('FMP price fetch error: ' + e.toString());
  }
  
  return null;
}

/**
 * FMP: Fetch company profile (shares outstanding)
 */
function fmpFetchShares_(symbol) {
  if (!USE_FMP_BACKUP || !FMP_API_KEY || FMP_API_KEY === 'YOUR_FMP_KEY_HERE') {
    Logger.log('FMP backup disabled or no API key');
    return null;
  }
  
  var url = 'https://financialmodelingprep.com/api/v3/profile/' + 
            encodeURIComponent(symbol) + '?apikey=' + FMP_API_KEY;
  
  try {
    Logger.log('Fetching FMP shares...');
    var options = {
      muteHttpExceptions: true,
      timeout: 10
    };
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    
    if (code !== 200) {
      Logger.log('FMP shares failed with code: ' + code);
      return null;
    }
    
    var data = JSON.parse(response.getContentText());
    
    if (Array.isArray(data) && data.length > 0) {
      var shares = toNum_(data[0].sharesOutstanding);
      if (isPos_(shares)) {
        Logger.log('FMP shares for ' + symbol + ': ' + shares);
        return shares;
      }
    }
  } catch (e) {
    Logger.log('FMP shares fetch error: ' + e.toString());
  }
  
  return null;
}

/**
 * FMP: Fetch free cash flow from financial statements
 * NOTE: Cash flow endpoint requires FMP Premium plan!
 */
function fmpFetchFCF_(symbol) {
  if (!USE_FMP_BACKUP || !FMP_API_KEY || FMP_API_KEY === 'YOUR_FMP_KEY_HERE') {
    Logger.log('FMP backup disabled or no API key');
    return null;
  }
  
  // Cash flow statement endpoint is PREMIUM ONLY on FMP free tier
  Logger.log('FMP FCF: Cash flow endpoint requires premium plan - skipping');
  return null;
  
  /* PREMIUM ONLY - Commented out
  var url = 'https://financialmodelingprep.com/api/v3/cash-flow-statement/' + 
            encodeURIComponent(symbol) + '?limit=1&apikey=' + FMP_API_KEY;
  
  try {
    Logger.log('Fetching FMP FCF...');
    var options = {
      muteHttpExceptions: true,
      timeout: 10
    };
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    
    if (code !== 200) {
      Logger.log('FMP FCF failed with code: ' + code);
      return null;
    }
    
    var data = JSON.parse(response.getContentText());
    
    if (Array.isArray(data) && data.length > 0) {
      var annual = data[0];
      var fcf = toNum_(annual.freeCashFlow);
      
      if (fcf !== null && isFinite(fcf)) {
        Logger.log('FMP FCF for ' + symbol + ': ' + fcf);
        return fcf;
      }
      
      // Fallback: calculate from OCF - CapEx
      var ocf = toNum_(annual.operatingCashFlow);
      var capex = toNum_(annual.capitalExpenditure);
      
      if (ocf !== null && capex !== null) {
        fcf = ocf - Math.abs(capex);
        Logger.log('FMP FCF (calculated) for ' + symbol + ': ' + fcf);
        return fcf;
      }
    }
  } catch (e) {
    Logger.log('FMP FCF fetch error: ' + e.toString());
  }
  */
  
  return null;
}

/**
 * FMP: Fetch growth rate from historical cash flows
 * NOTE: Cash flow endpoint requires FMP Premium plan!
 */
function fmpFetchGrowth_(symbol) {
  if (!USE_FMP_BACKUP || !FMP_API_KEY || FMP_API_KEY === 'YOUR_FMP_KEY_HERE') {
    Logger.log('FMP backup disabled or no API key');
    return null;
  }
  
  // Cash flow statement endpoint is PREMIUM ONLY on FMP free tier
  Logger.log('FMP Growth: Cash flow endpoint requires premium plan - skipping');
  return null;
  
  /* PREMIUM ONLY - Commented out
  var url = 'https://financialmodelingprep.com/api/v3/cash-flow-statement/' + 
            encodeURIComponent(symbol) + '?limit=5&apikey=' + FMP_API_KEY;
  
  try {
    Logger.log('Fetching FMP growth...');
    var options = {
      muteHttpExceptions: true,
      timeout: 10
    };
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    
    if (code !== 200) {
      Logger.log('FMP growth failed with code: ' + code);
      return null;
    }
    
    var data = JSON.parse(response.getContentText());
    
    if (Array.isArray(data) && data.length >= 2) {
      var fcfs = [];
      for (var i = 0; i < data.length; i++) {
        var fcf = toNum_(data[i].freeCashFlow);
        if (fcf !== null && isFinite(fcf)) {
          fcfs.push(fcf);
        }
      }
      
      if (fcfs.length >= 2) {
        // Simple CAGR calculation
        var n = fcfs.length - 1;
        var endValue = fcfs[0]; // Most recent
        var startValue = fcfs[fcfs.length - 1]; // Oldest
        
        if (isPos_(startValue) && isPos_(endValue)) {
          var growth = Math.pow(endValue / startValue, 1 / n) - 1;
          Logger.log('FMP growth for ' + symbol + ': ' + (growth * 100).toFixed(1) + '%');
          return growth;
        }
      }
    }
  } catch (e) {
    Logger.log('FMP growth fetch error: ' + e.toString());
  }
  */
  
  return null;
}

/***** DCF MAIN (Alpha Vantage only) *****/
function calculateIntrinsicValueAlphaVantage(targetSheet){
  try {
    // Use provided sheet or try to get active sheet (for manual runs)
    var sheet = targetSheet || SpreadsheetApp.getActiveSheet();
    if (!sheet) {
      Logger.log('Error: No sheet available');
      return;
    }
    var rawTicker = String(sheet.getRange(TICKER_CELL).getValue()).trim();
    if (!rawTicker) { 
      if (!SILENT_MODE) safeAlert_('Enter a ticker in ' + TICKER_CELL); 
      return; 
    }
    var symbol = stripExchangePrefix_(rawTicker); // keep ".TO" etc.
    var discountRate = toNum_(sheet.getRange(DISCOUNT_CELL).getValue());
    var years = Number(sheet.getRange(YEARS_CELL).getValue());
    var terminalGrowthRate = toNum_(sheet.getRange(TERM_GROWTH_CELL).getValue());
    if (!isPos_(discountRate) || !(years > 0) || terminalGrowthRate === null) {
      if (!SILENT_MODE) safeAlert_('Check B2 (discount), D2 (years), G2 (terminal growth).');
      return;
    }
    if (discountRate <= terminalGrowthRate){
      if (!SILENT_MODE) safeAlert_('Discount rate must be greater than terminal growth.');
      return;
    }
    var status = [];

    // (Intentionally skipping SYMBOL_SEARCH pre-check)

    // Price - Try AV first, fallback to FMP
    var price = avFetchPrice_(symbol);
    if (price) {
      status.push('Price OK (AV)');
    } else if (USE_FMP_BACKUP) {
      Logger.log('AV price failed, trying FMP backup...');
      price = fmpFetchPrice_(symbol);
      if (price) {
        status.push('Price OK (FMP backup)');
      } else {
        status.push('Price missing (AV & FMP)');
      }
    } else {
      status.push('Price missing (AV)');
    }

    // Overview → Shares / MarketCap / EPS / PE - Try AV first, fallback to FMP
    var overview = avFetchOverview_(symbol);
    var shares = (overview && overview.shares) ? overview.shares : null;
    
    if (isPos_(shares)) {
      status.push('Shares OK (AV OVERVIEW)');
      sheet.getRange(SHARES_CELL).setValue(shares);
    } else if (USE_FMP_BACKUP) {
      Logger.log('AV shares failed, trying FMP backup...');
      shares = fmpFetchShares_(symbol);
      if (isPos_(shares)) {
        status.push('Shares OK (FMP backup)');
        sheet.getRange(SHARES_CELL).setValue(shares);
      } else {
        // Try user input or calculate from market cap
        var userShares = toNum_(sheet.getRange(SHARES_CELL).getValue());
        if (isPos_(userShares)) {
          shares = userShares; 
          status.push('Shares from sheet F2');
        } else if (overview && isPos_(overview.marketCap) && isPos_(price)) {
          shares = overview.marketCap / price;
          sheet.getRange(SHARES_CELL).setValue(shares);
          status.push('Shares computed from MarketCap/Price (OVERVIEW)');
        } else {
          status.push('Shares missing (AV, FMP & sheet)');
        }
      }
    } else {
      // FMP backup disabled, try fallbacks
      var userShares = toNum_(sheet.getRange(SHARES_CELL).getValue());
      if (isPos_(userShares)) {
        shares = userShares; 
        status.push('Shares from sheet F2');
      } else if (overview && isPos_(overview.marketCap) && isPos_(price)) {
        shares = overview.marketCap / price;
        sheet.getRange(SHARES_CELL).setValue(shares);
        status.push('Shares computed from MarketCap/Price (OVERVIEW)');
      } else {
        status.push('Shares missing (AV & sheet)');
      }
    }

    // EPS & PE (info only)
    var epsOut = (overview && overview.epsTTM) ? overview.epsTTM
               : (overview && overview.eps) ? overview.eps : null;
    var peOut = (overview && overview.pe) ? overview.pe : null;
    if ((peOut === null || !isFinite(peOut)) && isPos_(price) && overview && overview.epsTTM && overview.epsTTM !== 0) {
      peOut = price / overview.epsTTM;
      status.push('PE computed (Price / EPS_TTM)');
    }
    sheet.getRange(EPS_CELL).setValue((epsOut !== null && isFinite(epsOut)) ? epsOut : '');
    sheet.getRange(PE_CELL).setValue((peOut !== null && isFinite(peOut)) ? peOut : '');
    
    // Store for Graham analysis (avoids stale sheet reads between ticker changes)
    var bvpsOut = (overview && overview.bookValuePerShare !== null) ? overview.bookValuePerShare : null;
    var props = PropertiesService.getScriptProperties();
    props.setProperty('LAST_TICKER', String(symbol));
    if (epsOut !== null && isFinite(epsOut))  props.setProperty('LAST_EPS',  String(epsOut));  else props.deleteProperty('LAST_EPS');
    if (peOut  !== null && isFinite(peOut))   props.setProperty('LAST_PE',   String(peOut));   else props.deleteProperty('LAST_PE');
    if (bvpsOut !== null && isFinite(bvpsOut)) props.setProperty('LAST_BVPS', String(bvpsOut)); else props.deleteProperty('LAST_BVPS');
    
    status.push((epsOut !== null) ? 'EPS OK (OVERVIEW)' : 'EPS missing (OVERVIEW)');
    status.push((peOut !== null) ? 'PE OK (OVERVIEW/Computed)' : 'PE missing (OVERVIEW)');
    status.push((bvpsOut !== null) ? 'BookValue OK (OVERVIEW)' : 'BookValue missing (OVERVIEW)');

    // FCF (current) - Try AV first, fallback to FMP
    var fcf = avFetchFCF_(symbol);
    if (fcf !== null && isFinite(fcf)) {
      status.push('FCF OK (AV annual/TTM)');
    } else if (USE_FMP_BACKUP) {
      Logger.log('AV FCF failed, trying FMP backup...');
      fcf = fmpFetchFCF_(symbol);
      if (fcf !== null && isFinite(fcf)) {
        status.push('FCF OK (FMP backup)');
      } else {
        status.push('FCF missing (AV & FMP); please enter E2 manually');
      }
    } else {
      status.push('FCF missing (AV); please enter E2 manually');
    }

    // Growth from FCF - Try AV first, fallback to FMP
    var growth = avFetchGrowthFromFCF_(symbol);
    if (growth !== null && isFinite(growth)) {
      status.push('Growth OK (AV FCF CAGR)');
    } else if (USE_FMP_BACKUP) {
      Logger.log('AV growth failed, trying FMP backup...');
      growth = fmpFetchGrowth_(symbol);
      if (growth !== null && isFinite(growth)) {
        status.push('Growth OK (FMP backup)');
      } else {
        growth = DEFAULT_GROWTH;
        status.push('Growth DEFAULT ' + DEFAULT_GROWTH + ' (AV & FMP unavailable)');
      }
    } else {
      growth = DEFAULT_GROWTH;
      status.push('Growth DEFAULT ' + DEFAULT_GROWTH + ' (FCF growth unavailable)');
    }

    // Write basics
    sheet.getRange(GROWTH_CELL).setValue(growth);
    if (fcf !== null && isFinite(fcf)) sheet.getRange(FCF_CELL).setValue(fcf);
    if (isPos_(price)) sheet.getRange(PRICE_CELL).setValue(price);

    // Must have shares & FCF for IV/MOS
    if (!isPos_(shares) || !(fcf !== null && isFinite(fcf))) {
      var msg = [];
      if (!isPos_(shares)) msg.push('Shares (F2)');
      if (!(fcf !== null && isFinite(fcf))) msg.push('FCF (E2)');
      if (!SILENT_MODE) safeAlert_('Missing: ' + msg.join(', ') + '. Enter values and rerun.');
      sheet.getRange(STATUS_CELL).setValue('Status: ' + status.join('\n'));
      sheet.getRange(TS_CELL).setValue('Timestamp: ' + new Date().toISOString());
      return;
    }

    // DCF calculation
    var intrinsicValue = 0;
    var projectedFCF = fcf;
    for (var i=1; i<=years; i++){
      projectedFCF *= (1 + growth);
      intrinsicValue += projectedFCF / Math.pow(1 + discountRate, i);
    }
    var terminalValue = projectedFCF * (1 + terminalGrowthRate) / (discountRate - terminalGrowthRate);
    intrinsicValue += terminalValue / Math.pow(1 + discountRate, years);
    var ivPerShare = intrinsicValue / shares;
    sheet.getRange(IV_CELL).setValue(ivPerShare);

    // Margin of Safety (%)
    var mos = (isPos_(ivPerShare) && isPos_(price)) ? ((ivPerShare - price) / ivPerShare) * 100 : null;
    sheet.getRange(MOS_CELL).setValue(mos !== null ? mos : '');

    // Status & timestamp
    sheet.getRange(STATUS_CELL).setValue('Status: ' + status.join('\n'));
    sheet.getRange(TS_CELL).setValue('Timestamp: ' + new Date().toISOString());
  } catch (err) {
    safeAlert_(err.message);
    Logger.log('Error: ' + err.message);
  }
}

/***** Maintenance: Clear legacy AV_* properties (to fix quota errors) *****/
function clearLegacyAvProperties_(){
  try {
    var sp = PropertiesService.getScriptProperties();
    var props = sp.getProperties();
    var deleted = 0;
    Object.keys(props).forEach(function(k){
      if (k.startsWith('AV_') || k === 'AV_CACHE_INDEX') { sp.deleteProperty(k); deleted++; }
    });
    SpreadsheetApp.getActiveSheet().getRange(STATUS_CELL)
      .setValue('Status: cleared ' + deleted + ' legacy AV_* properties');
  } catch(e){
    SpreadsheetApp.getActiveSheet().getRange(STATUS_CELL)
      .setValue('Status: error clearing legacy properties → ' + e.message);
  }
}

/***********************************************************************
 * INVESTING AGENT (Alpha Vantage + Google Sheets)
 ***********************************************************************/
function ensureSheet_(name, tabColorHex){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh){
    sh = ss.insertSheet(name);
    if (tabColorHex){ try{ sh.setTabColor(tabColorHex); }catch(e){} }
  }
  return sh;
}

/** Initialize Agent tabs: Config, Watchlist, Signals, Trade_Log */
function InitializeAgentSheets(){
  var ss = SpreadsheetApp.getActive();
  // Config
  var cfg = ensureSheet_('Agent_Config', '#6AA84F');
  var cfgHeaders = ['MACRO','EMAIL','LOG_HEARTBEAT'];
  if (cfg.getLastRow()<1 || String(cfg.getRange(1,1).getValue()).toUpperCase()!=='MACRO'){
    cfg.clear();
    cfg.getRange(1,1,1,cfgHeaders.length).setValues([cfgHeaders]);
    cfg.getRange(2,1).setValue('BASE'); // BASE / SHALLOW_RECESSION / HARD_BEAR / UPSIDE
    cfg.getRange(2,2).setValue(Session.getActiveUser().getEmail());
    cfg.getRange(2,3).setValue(true);
    cfg.setFrozenRows(1);
    cfg.autoResizeColumns(1,cfgHeaders.length);
  }
  // Watchlist
  var wl = ensureSheet_('Agent_Watchlist', '#3D85C6');
  var wlHeaders = [
    'Ticker','Exchange','PositionWeight%','MaxWeight%','ATH','LastTrimPrice',
    'AddTrigger%','TrimTrigger%','StopLoss%','Notes'
  ];
  if (wl.getLastRow()<1 || String(wl.getRange(1,1).getValue()).toUpperCase()!=='TICKER'){
    wl.clear();
    wl.getRange(1,1,1,wlHeaders.length).setValues([wlHeaders]);
    var starter = [
      ['VEQT','TSX',75,100,'','', 8, 0, 0,'Core anchor (never sell; rebalance only)'],
      ['NVDA','NASDAQ',5, 7,'','',10, 20, 0,'Trim +20% vs ATH; add only on macro dips'],
      ['VRT','NYSE',4, 5,'','',10, 30, 0,'Add -20% pullback if fundamentals OK'],
      ['CLS','TSX',4, 5,'','',10, 25, 0,'Trim parabolic after guidance raises'],
      ['PHYS','NYSE',5, 8,'','', 0, 0, 0,'Hedge; add when inflation >3% or recession risk rises'],
      ['RIVN','NASDAQ',2, 3,'','',15, 25, 20,'Add only after R2 validation; stop-loss if needed'],
      ['IONQ','NYSE',2, 3,'','',15, 25, 20,'Add on real revenue wins; trim on dilution spikes'],
      ['QNC.V','TSXV',2.5,2.5,'','',20, 25, 30,'Quantum eMotion; add on certified adoption/enterprise deals']
    ];
    wl.getRange(2,1,starter.length,starter[0].length).setValues(starter);
    wl.setFrozenRows(1);
    wl.autoResizeColumns(1,wlHeaders.length);
  }
  // Signals
  var sig = ensureSheet_('Agent_Signals', '#E69138');
  var sigHeaders = ['Timestamp','Ticker','Symbol','Price','Macro','Action','Size%','Reason'];
  if (sig.getLastRow()<1 || String(sig.getRange(1,1).getValue()).toUpperCase()!=='TIMESTAMP'){
    sig.clear();
    sig.getRange(1,1,1,sigHeaders.length).setValues([sigHeaders]);
    sig.setFrozenRows(1);
    sig.autoResizeColumns(1,sigHeaders.length);
  }
  // Trade log
  var tl = ensureSheet_('Trade_Log', '#C27BA0');
  var tlHeaders = ['Timestamp','Ticker','Action','QtyOrPct','Price','Macro','Note'];
  if (tl.getLastRow()<1 || String(tl.getRange(1,1).getValue()).toUpperCase()!=='TIMESTAMP'){
    tl.clear();
    tl.getRange(1,1,1,tlHeaders.length).setValues([tlHeaders]);
    tl.setFrozenRows(1);
    tl.autoResizeColumns(1,tlHeaders.length);
  }
  // Timezone: America/Toronto
  try{
    var tz = ss.getSpreadsheetTimeZone();
    if (!/America\/(Toronto|New_York)/.test(tz)) ss.setSpreadsheetTimeZone('America/Toronto');
  }catch(e){}
  safeAlert_('Agent tabs ready:\n- Agent_Config\n- Agent_Watchlist\n- Agent_Signals\n- Trade_Log');
}

/** Initialize Holdings tab with headers + sample rows */
function InitializeHoldingsSheet(){
  var ss = SpreadsheetApp.getActive();
  var hd = ss.getSheetByName('Holdings');
  if (!hd){
    hd = ss.insertSheet('Holdings');
    try{ hd.setTabColor('#8E7CC3'); }catch(e){}
  }
  var hasHeaders = hd.getLastRow()>=1 && String(hd.getRange(1,1).getValue()).toUpperCase()==='TICKER';
  if (!hasHeaders){
    hd.clear();
    hd.getRange(1,1,1,3).setValues([['Ticker','Shares','Notes']]);
    var starter = [
      ['VEQT', 1250, 'Core'],
      ['NVDA', 45, ''],
      ['VRT', 120, ''],
      ['CLS', 300, ''],
      ['PHYS', 600, ''],
      ['RIVN', 80, 'Spec'],
      ['IONQ', 100, 'Spec'],
      ['QNC.V',12000,'Spec – Quantum eMotion']
    ];
    hd.getRange(2,1,starter.length,starter[0].length).setValues(starter);
    hd.setFrozenRows(1);
    hd.autoResizeColumns(1,3);
  }
  safeAlert_('Holdings sheet initialized. Edit your actual shares.');
}

/** Map tickers to AlphaVantage symbol where needed (.TO/.V) */
function mapTickerAlpha_(ticker){
  ticker = String(ticker || '').trim();
  if (!ticker) return null;
  if (/\.[A-Z]+$/.test(ticker)) return ticker;  // keep explicit suffix
  if (ticker === 'VEQT') return 'VEQT.TO';
  if (ticker === 'CLS')  return 'CLS.TO';
  if (ticker === 'PHYS') return 'PHYS';
  if (ticker === 'NVDA') return 'NVDA';
  if (ticker === 'VRT')  return 'VRT';
  if (ticker === 'IONQ') return 'IONQ';
  if (ticker === 'RIVN') return 'RIVN';
  if (ticker === 'QNC')  return 'QNC.V';
  return ticker;
}

/** Update PositionWeight% in Agent_Watchlist from Holdings (Shares × Price) */
function UpdatePositionWeightsFromHoldings(){
  var ss = SpreadsheetApp.getActive();
  var wl = ss.getSheetByName('Agent_Watchlist');
  if (!wl){ InitializeAgentSheets(); wl = ss.getSheetByName('Agent_Watchlist'); }
  if (!wl) throw new Error('Agent_Watchlist could not be created.');
  var hd = ss.getSheetByName('Holdings');
  if (!hd){ InitializeHoldingsSheet(); hd = ss.getSheetByName('Holdings'); }
  if (!hd) throw new Error('Holdings could not be created.');

  var lastH = hd.getLastRow();
  if (lastH < 2){ safeAlert_('Holdings has no data rows. Fill your shares.'); return; }
  var hRows = hd.getRange(2,1,lastH-1,2).getValues(); // Ticker | Shares
  var sharesMap = {};
  hRows.forEach(function(r){
    var t = String(r[0] || '').trim();
    var s = Number(r[1] || 0);
    if (t && s > 0) sharesMap[t.toUpperCase()] = s;
  });

  var lastW = wl.getLastRow();
  if (lastW < 2){ safeAlert_('Agent_Watchlist has no data rows.'); return; }
  var headers = wl.getRange(1,1,1,10).getValues()[0];
  var idx = {}; headers.forEach(function(h,i){ idx[String(h).trim()] = i; });
  var wData = wl.getRange(2,1,lastW-1,10).getValues();

  var rowsInfo = [];
  var totalValue = 0;
  for (var r=0; r<wData.length; r++){
    var row = wData[r];
    var ticker = String(row[idx['Ticker']] || '').trim();
    if (!ticker) continue;
    var tUpper = ticker.toUpperCase();
    var shares = sharesMap[tUpper] || 0;
    var symbol = mapTickerAlpha_(ticker);
    var price = avFetchPrice_(symbol) || 0;
    var value = (shares > 0 && price > 0) ? (shares * price) : 0;
    rowsInfo.push({ rowIndex: 2 + r, ticker: ticker, symbol: symbol, price: price, shares: shares, value: value });
    totalValue += value;
  }
  if (totalValue <= 0){ safeAlert_('Total portfolio value = 0. Check shares & quotes.'); return; }

  rowsInfo.forEach(function(info){
    var weightPct = (info.value > 0) ? (info.value / totalValue * 100) : 0;
    wl.getRange(info.rowIndex, idx['PositionWeight%']+1).setValue(weightPct);
  });
  safeAlert_('Position weights updated from Holdings. Total ≈ ' + totalValue.toFixed(2));
}

/** Agent: signals (ADD / TRIM / CUT / ALERT) */
function getAgentConfig_(){
  var ss = SpreadsheetApp.getActive();
  var cfg = ss.getSheetByName('Agent_Config');
  if (!cfg) throw new Error('Run InitializeAgentSheets() first');
  var macro = String(cfg.getRange('A2').getValue()).trim().toUpperCase() || 'BASE';
  var email = String(cfg.getRange('B2').getValue()).trim();
  var heartbeat = Boolean(cfg.getRange('C2').getValue());
  return { macro: macro, email: email, heartbeat: heartbeat };
}
function suggestTrimSize_(ticker){
  ticker = String(ticker || '').toUpperCase();
  if (ticker === 'NVDA') return 10;
  if (ticker === 'VRT' || ticker === 'CLS') return 10;
  if (ticker === 'RIVN' || ticker === 'IONQ' || ticker === 'QNC' || ticker === 'QNC.V') return 25;
  return 10;
}
function suggestAddSize_(ticker, macro){
  macro = String(macro || '').toUpperCase();
  if (macro === 'HARD_BEAR') return 0;
  if (macro === 'BASE') return 5;
  if (macro === 'SHALLOW_RECESSION') return 7;
  if (macro === 'UPSIDE') return 3;
  return 0;
}
function sendEmailAlerts_(rows, email){
  if (!email || !rows || !rows.length) return;
  try{
    var lines = rows.map(function(r){
      var ts = Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      return '• ' + r[1] + ' @ ' + Number(r[3]).toFixed(2) +
             ' → ' + r[5] + ' ' + r[6] + '% \n' + r[7] + ' (' + r[4] + ', ' + ts + ')';
    }).join('\n');
    var subj = 'Agent Alerts [' + rows.length + '] — ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    MailApp.sendEmail(email, subj, lines);
  }catch(e){ Logger.log('sendEmailAlerts_ error: ' + e.message); }
}
function pushHeartbeatRow_(rowsOut, ticker, symbol, price, macro){
  rowsOut.push([ new Date(), ticker, symbol, price, macro, 'CHECK', 0, 'No signal fired' ]);
}
function RunAgentNow(){
  UpdatePositionWeightsFromHoldings();
  var ss = SpreadsheetApp.getActive();
  var cfg = getAgentConfig_();
  var wl = ss.getSheetByName('Agent_Watchlist');
  var sig = ss.getSheetByName('Agent_Signals');
  var lastRow = wl.getLastRow();
  if (lastRow < 2) throw new Error('Agent_Watchlist has no data rows.');
  var headers = wl.getRange(1,1,1,10).getValues()[0];
  var idx = {}; headers.forEach(function(h,i){ idx[String(h).trim()] = i; });
  var data = wl.getRange(2,1,lastRow-1,10).getValues();
  var outRows = [];
  var specTickers = {'RIVN':true,'IONQ':true,'QNC':true,'QNC.V':true};
  var specBucket = 0;

  for (var r=0; r<data.length; r++){
    var row = data[r];
    var ticker = String(row[idx['Ticker']] || '').trim();
    if (!ticker) continue;
    var symbol = mapTickerAlpha_(ticker);
    var positionWeight = Number(row[idx['PositionWeight%']] || 0);
    var maxWeight = Number(row[idx['MaxWeight%']] || 0);
    var ath = Number(row[idx['ATH']] || 0);
    var lastTrimPrice = Number(row[idx['LastTrimPrice']] || 0);
    var addTriggerPct = Number(row[idx['AddTrigger%']] || 0);
    var trimTriggerPct = Number(row[idx['TrimTrigger%']] || 0);
    var stopLossPct = Number(row[idx['StopLoss%']] || 0);
    var tUpper = ticker.toUpperCase();
    if (specTickers[tUpper]) specBucket += positionWeight;

    var price = avFetchPrice_(symbol);
    if (!price || !isPos_(price)) {
      if (cfg.heartbeat) pushHeartbeatRow_(outRows, ticker, symbol, 0, cfg.macro);
      continue;
    }

    // Update ATH
    var sheetRow = 2 + r;
    if (!ath || price > ath){
      ath = price;
      wl.getRange(sheetRow, idx['ATH']+1).setValue(ath);
    }

    var signals = [];
    if (tUpper !== 'VEQT'){
      if (trimTriggerPct > 0 && price >= ath * (1 + trimTriggerPct/100)){
        signals.push({ action:'TRIM', sizePct:suggestTrimSize_(ticker), reason:'+' + trimTriggerPct + '% vs ATH' });
      }
      if (addTriggerPct > 0 && price <= ath * (1 - addTriggerPct/100)){
        var addSize = suggestAddSize_(ticker, cfg.macro);
        if (addSize > 0) signals.push({ action:'ADD', sizePct:addSize, reason:'-' + addTriggerPct + '% from ATH (macro ' + cfg.macro + ')' });
      }
      if (stopLossPct > 0 && lastTrimPrice && price <= lastTrimPrice * (1 - stopLossPct/100)){
        var cutSize = (specTickers[tUpper] ? 50 : 25);
        signals.push({ action:'CUT', sizePct:cutSize, reason:'Stop-loss ' + stopLossPct + '% from last trim' });
      }
      if (maxWeight > 0 && positionWeight > maxWeight){
        signals.push({ action:'TRIM', sizePct:10, reason:'Position ' + positionWeight + '% > max ' + maxWeight + '%' });
      }
    }

    // Macro override
    if (cfg.macro === 'HARD_BEAR'){
      signals = signals.filter(function(s){ return s.action !== 'ADD'; });
      if (specTickers[tUpper]){
        signals.push({ action:'CUT', sizePct:50, reason:'HARD_BEAR regime: cut speculative exposure' });
      }
    }

    signals.forEach(function(s){
      outRows.push([ new Date(), ticker, symbol, price, cfg.macro, s.action, s.sizePct, s.reason ]);
    });
    if (cfg.heartbeat && signals.length === 0){
      pushHeartbeatRow_(outRows, ticker, symbol, price, cfg.macro);
    }
  }

  if (specBucket > 5){
    outRows.push([
      new Date(), 'SPEC_BUCKET', '-', 0, cfg.macro, 'ALERT', 0,
      'Spec bucket (RIVN+IONQ+QNC.V) = ' + specBucket + '% > 5% cap'
    ]);
  }

  if (outRows.length){
    sig.getRange(sig.getLastRow()+1,1,outRows.length,outRows[0].length).setValues(outRows);
    sendEmailAlerts_(outRows, cfg.email);
  }
}

/** Trade logging + Macro helpers */
function RecordTradeAction(ticker, action, qtyOrPct, note){
  var ss = SpreadsheetApp.getActive();
  var cfg = getAgentConfig_();
  var wl = ss.getSheetByName('Agent_Watchlist');
  var tl = ss.getSheetByName('Trade_Log');
  if (!wl || !tl) throw new Error('Run InitializeAgentSheets() first');
  var lastRow = wl.getLastRow();
  var headers = wl.getRange(1,1,1,10).getValues()[0];
  var idx = {}; headers.forEach(function(h,i){ idx[String(h).trim()] = i; });
  var data = wl.getRange(2,1,lastRow-1,10).getValues();
  var rowIndex = null;
  for (var r=0; r<data.length; r++){
    if (String(data[r][idx['Ticker']]).trim().toUpperCase() === String(ticker).trim().toUpperCase()){
      rowIndex = 2 + r; break;
    }
  }
  var symbol = mapTickerAlpha_(ticker);
  var price = avFetchPrice_(symbol);
  tl.getRange(tl.getLastRow()+1,1,1,7).setValues([[
    new Date(), ticker, action, qtyOrPct, price, cfg.macro, note || ''
  ]]);
  if (rowIndex !== null && String(action).toUpperCase()==='TRIM' && price && isPos_(price)){
    wl.getRange(rowIndex, idx['LastTrimPrice']+1).setValue(price);
  }
}
function SetMacro_BASE(){ SpreadsheetApp.getActive().getSheetByName('Agent_Config').getRange('A2').setValue('BASE'); }
function SetMacro_SHALLOW_RECESSION(){ SpreadsheetApp.getActive().getSheetByName('Agent_Config').getRange('A2').setValue('SHALLOW_RECESSION'); }
function SetMacro_HARD_BEAR(){ SpreadsheetApp.getActive().getSheetByName('Agent_Config').getRange('A2').setValue('HARD_BEAR'); }
function SetMacro_UPSIDE(){ SpreadsheetApp.getActive().getSheetByName('Agent_Config').getRange('A2').setValue('UPSIDE'); }

/** Debug (optional) */
function RunAgentDebug(){
  var ss = SpreadsheetApp.getActive();
  var cfg = getAgentConfig_();
  var wl = ss.getSheetByName('Agent_Watchlist');
  if (!wl) throw new Error('Run InitializeAgentSheets() first');
  var dbg = ss.getSheetByName('Agent_Debug') || ss.insertSheet('Agent_Debug');
  dbg.clear();
  dbg.appendRow(['Timestamp','Ticker','Symbol','Price','ATH','AddTrigger%','TrimTrigger%','StopLoss%','Would_ADD?','Would_TRIM?','Would_CUT?','Macro']);
  var lastRow = wl.getLastRow();
  var headers = wl.getRange(1,1,1,10).getValues()[0];
  var idx = {}; headers.forEach(function(h,i){ idx[String(h).trim()] = i; });
  var data = wl.getRange(2,1,lastRow-1,10).getValues();
  for (var r=0; r<data.length; r++){
    var row = data[r];
    var ticker = String(row[idx['Ticker']] || '').trim(); if (!ticker) continue;
    var symbol = mapTickerAlpha_(ticker);
    var price = avFetchPrice_(symbol);
    var ath = Number(row[idx['ATH']] || 0);
    var addP = Number(row[idx['AddTrigger%']] || 0);
    var trimP = Number(row[idx['TrimTrigger%']] || 0);
    var slP = Number(row[idx['StopLoss%']] || 0);
    var lastTrim = Number(row[idx['LastTrimPrice']] || 0);
    var wouldAdd = (addP > 0 && ath && price && price <= ath * (1 - addP/100) && cfg.macro !== 'HARD_BEAR') ? 'YES' : 'no';
    var wouldTrim = (trimP > 0 && ath && price && price >= ath * (1 + trimP/100)) ? 'YES' : 'no';
    var wouldCut = (slP > 0 && lastTrim && price && price <= lastTrim * (1 - slP/100)) ? 'YES' : 'no';
    dbg.appendRow([new Date(), ticker, symbol, price, ath, addP, trimP, slP, wouldAdd, wouldTrim, wouldCut, cfg.macro]);
  }
  dbg.setFrozenRows(1);
  dbg.autoResizeColumns(1, 12);
  safeAlert_('Debug snapshot written to Agent_Debug');
}

/***** SINGLE onOpen(): both menus *****/
function onOpen(){
  var ui = SpreadsheetApp.getUi();

  // --- DCF menu ---
  ui.createMenu('DCF Tool (Alpha Vantage)')
    .addItem('Run DCF (Row 2)', 'calculateIntrinsicValueAlphaVantage')
    .addSeparator()
    .addItem('Clear Legacy Properties', 'clearLegacyAvProperties_')
    .addToUi();

  // --- Investing Agent menu ---
  ui.createMenu('Investing Agent')
    .addItem('Initialize Agent Sheets', 'InitializeAgentSheets')
    .addItem('Initialize Holdings Sheet', 'InitializeHoldingsSheet')
    .addSeparator()
    .addItem('Update Position Weights', 'UpdatePositionWeightsFromHoldings')
    .addItem('Run Agent Now', 'RunAgentNow')
    .addItem('Install Agent Timers', 'InstallAgentTimers')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Set Macro')
        .addItem('BASE', 'SetMacro_BASE')
        .addItem('SHALLOW_RECESSION', 'SetMacro_SHALLOW_RECESSION')
        .addItem('HARD_BEAR', 'SetMacro_HARD_BEAR')
        .addItem('UPSIDE', 'SetMacro_UPSIDE')
    )
    .addSeparator()
    .addItem('Run Agent Debug', 'RunAgentDebug')
    .addItem('Record Trade: TRIM NVDA 10%', 'RecordTradeAction_NVDA_Trim10')
    .addToUi();

    // Add this to your existing onOpen() function
  ui.createMenu('💬 Messenger Remote')
    .addItem('✅ Install Messenger Trigger', 'InstallMessengerTrigger')
    .addItem('❌ Uninstall Trigger', 'UninstallMessengerTrigger')
    .addSeparator()
    .addItem('🧪 Test Connection', 'TestMessengerConnection')
    .addItem('🔄 Manual Check Messages', 'ManualCheckMessages')
    .addItem('🧪 Test Command Parsing', 'TestCommandsParsing')
    .addSeparator()
    .addItem('📖 Setup Guide', 'ShowMessengerSetupGuide')
    .addToUi();
}

/** Convenience menu action */
function RecordTradeAction_NVDA_Trim10(){
  RecordTradeAction('NVDA','TRIM','10%','Menu quick action');
}

/** Quick diagnostic: verify key & endpoints without the whole DCF */
function PingAlphaVantageQuick(){
  var symbol = 'NVDA';
  var url1 = 'https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol='
             + encodeURIComponent(symbol) + '&apikey=' + ALPHAVANTAGE_API_KEY;
  var url2 = 'https://www.alphavantage.co/query?function=CASH_FLOW&symbol='
             + encodeURIComponent(symbol) + '&apikey=' + ALPHAVANTAGE_API_KEY;

  var r1 = httpGetAV_(url1);
  var r2 = httpGetAV_(url2);

  Logger.log('GLOBAL_QUOTE status: ' + r1.status + ' body: ' + (r1.body ? r1.body.slice(0,300) : ''));
  Logger.log('CASH_FLOW    status: ' + r2.status + ' body: ' + (r2.body ? r2.body.slice(0,300) : ''));

  SpreadsheetApp.getUi().alert(
    'GLOBAL_QUOTE: ' + r1.status + '\nCASH_FLOW: ' + r2.status +
    '\n\nCheck View → Logs for the first 300 chars of each response.'
  );
}
/**
 * MESSENGER REMOTE CONTROL FOR DCF & PORTFOLIO AGENT
 * 
 * Control your investment analysis via Facebook Messenger!
 * 
 * Commands: DCF AAPL, AGENT, STATUS, PRICE NVDA, HELP
 * 
 * Setup:
 * 1. Create Facebook page
 * 2. Get Page Access Token from Facebook Developer Console
 * 3. Get your Facebook User ID
 * 4. Configure MESSENGER_CONFIG below
 * 5. Run InstallMessengerTrigger()
 * 6. Message your page!
 */

// ===== CONFIGURATION =====

var MESSENGER_CONFIG = {
  enabled: true,
  pageAccessToken: 'EAANqc3eZBXwABQlZCAzAHlkGBhHZCQuZBYKP2PasnYvzhFgSI8HO4A2sirQqhBYDK11VxPQr5lvqxqMQmlDzzMwgYbWVkeIm6t8cnrB7Wp2jqgIpTWkZAWeXHpPYOan7hea8nk21oTC9ZA64C97AZAkZAvN0UK08CiW1bfVLw6XQRKP72wNx1vKCG7qf8PDn13fgXzydAdlqjQZDZD', // PASTE YOUR PAGE ACCESS TOKEN HERE
  yourFacebookId: '33834830566132027',     // PASTE YOUR FACEBOOK USER ID HERE
  checkIntervalMinutes: 1,                 // How often to check for messages (1-30)
  sendTypingIndicator: true                // Show "typing..." while processing
};

// ===== COMMAND PARSER =====

/**
 * Parse incoming command text
 * Returns: { command, params, valid }
 */
function parseCommand_(text) {
  var cleaned = String(text || '').trim().toUpperCase();
  var parts = cleaned.split(/\s+/);
  var command = parts[0];
  var params = parts.slice(1);
  
  // Command aliases
  var commandMap = {
    'DCF': 'DCF',
    'ANALYZE': 'DCF',
    'VALUE': 'DCF',
    'VALUATION': 'DCF',
    'AGENT': 'AGENT',
    'RUN': 'AGENT',
    'RUNAGENT': 'AGENT',
    'STATUS': 'STATUS',
    'PORTFOLIO': 'STATUS',
    'POSITIONS': 'STATUS',
    'PRICE': 'PRICE',
    'QUOTE': 'PRICE',
    'DEBUG': 'DEBUG',
    'SHOW': 'DEBUG',
    'HELP': 'HELP',
    'COMMANDS': 'HELP'
  };
  
  var normalizedCommand = commandMap[command];
  
  if (!normalizedCommand) {
    return { 
      valid: false, 
      error: '❓ Unknown command: "' + command + '"\n\nSend HELP to see available commands.'
    };
  }
  
  return {
    valid: true,
    command: normalizedCommand,
    params: params,
    rawText: text
  };
}

// ===== COMMAND EXECUTORS =====

/**
 * Execute parsed command
 * Returns: { success, message, data }
 */
function executeCommand_(parsed) {
  try {
    switch (parsed.command) {
      case 'DCF':
        return executeDCFCommand_(parsed.params);
      
      case 'AGENT':
        return executeAgentCommand_(parsed.params);
      
      case 'STATUS':
        return executeStatusCommand_(parsed.params);
      
      case 'PRICE':
        return executePriceCommand_(parsed.params);
      
      case 'DEBUG':
        return executeDebugCommand_(parsed.params);
      
      case 'HELP':
        return executeHelpCommand_();
      
      default:
        return { 
          success: false, 
          message: '❌ Command not implemented: ' + parsed.command 
        };
    }
  } catch (e) {
    Logger.log('Command execution error: ' + e.toString());
    return { 
      success: false, 
      message: '❌ Error: ' + e.toString() 
    };
  }
}

/**
 * DCF Command: Analyze a stock
 * Usage: DCF AAPL
 */
function executeDCFCommand_(params) {
  if (params.length === 0) {
    return { 
      success: false, 
      message: '📊 DCF Analysis\n\nUsage: DCF <TICKER>\n\nExamples:\n• DCF AAPL\n• DCF NVDA\n• DCF MSFT' 
    };
  }
  
  var ticker = params[0];
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('DCF') || ss.getSheetByName('Sheet1');
  
  if (!sheet) {
    return { 
      success: false, 
      message: '❌ DCF sheet not found. Please set up DCF sheet first.' 
    };
  }
  
  // Simple approach: Just set the ticker and call the SAME function the menu uses
  Logger.log('Setting ticker to: ' + ticker);
  sheet.getRange(TICKER_CELL).setValue(ticker);
  
  // Small delay for sheet to update
  Utilities.sleep(500);
  
  // Enable silent mode to suppress alerts
  SILENT_MODE = true;
  
  try {
    // Call the EXACT same function the menu calls - this works!
    Logger.log('Running calculateIntrinsicValueAlphaVantage...');
    calculateIntrinsicValueAlphaVantage(sheet);
    Logger.log('Calculate function returned');
  } catch (e) {
    SILENT_MODE = false;
    Logger.log('DCF execution error: ' + e.toString());
    return {
      success: false,
      message: '❌ Analysis failed: ' + e.toString()
    };
  } finally {
    SILENT_MODE = false;
  }
  
  // The function runs synchronously, so by this point the sheet SHOULD be populated
  // But API calls might still be running in background, so let's give it time
  Logger.log('Waiting 30 seconds for API calls to complete...');
  Utilities.sleep(30000); // 30 seconds - generous wait
  
  // Now read the results from the sheet
  Logger.log('Reading results from sheet...');
  var price = sheet.getRange(PRICE_CELL).getValue();
  var iv = sheet.getRange(IV_CELL).getValue();
  var mos = sheet.getRange(MOS_CELL).getValue();
  var status = sheet.getRange(STATUS_CELL).getValue();
  var fcf = sheet.getRange(FCF_CELL).getValue();
  var shares = sheet.getRange(SHARES_CELL).getValue();
  var growth = sheet.getRange(GROWTH_CELL).getValue();
  
  Logger.log('Results: Price=' + price + ', IV=' + iv + ', MOS=' + mos + ', FCF=' + fcf);
  
  // Check status for rate limit or error messages
  var statusStr = String(status || '');
  var isRateLimited = statusStr.indexOf('Note') >= 0 || 
                      statusStr.indexOf('call frequency') >= 0 ||
                      statusStr.indexOf('rate limit') >= 0;
  
  // Check if we got minimal data (at least price)
  if (!price && !iv && !fcf) {
    var errorMsg = '❌ No data received for ' + ticker + '\n\n';
    
    if (isRateLimited) {
      errorMsg += '⏱️ RATE LIMIT HIT\n\n';
      errorMsg += 'Alpha Vantage free tier:\n';
      errorMsg += '• 5 API calls per minute\n';
      errorMsg += '• Each DCF uses 3-4 calls\n\n';
      errorMsg += '⏰ Wait 2-3 minutes, then try again.';
    } else if (statusStr) {
      // Include status info if available
      errorMsg += 'Status from sheet:\n' + statusStr.substring(0, 200);
    } else {
      errorMsg += 'Possible causes:\n';
      errorMsg += '• Invalid ticker\n';
      errorMsg += '• API timeout\n';
      errorMsg += '• Network issue\n\n';
      errorMsg += 'Try running DCF manually in the sheet to see the error.';
    }
    
    return {
      success: false,
      message: errorMsg
    };
  }
  
  // Format the response message
  var message = '📊 DCF Analysis: ' + ticker + '\n';
  message += '═══════════════════\n';
  message += '💰 Price: $' + (price ? price.toFixed(2) : 'N/A') + '\n';
  message += '💎 Intrinsic Value: $' + (iv ? iv.toFixed(2) : 'N/A') + '\n';
  message += '📈 Margin of Safety: ' + (mos ? mos.toFixed(1) + '%' : 'N/A') + '\n';
  
  // Add FCF and Shares if available
  if (fcf) {
    message += '💵 FCF: $' + (fcf / 1e9).toFixed(2) + 'B\n';
  }
  if (shares) {
    message += '📊 Shares: ' + (shares / 1e9).toFixed(2) + 'B\n';
  }
  if (growth) {
    message += '📈 Growth: ' + (growth * 100).toFixed(1) + '%\n';
  }
  
  // ── Graham Number Section ──
  // Read EPS, PE, BVPS from ScriptProperties (stored during DCF calculation)
  var eps = null, pe = null, bvps = null;
  try {
    var props2 = PropertiesService.getScriptProperties();
    var storedTicker = props2.getProperty('LAST_TICKER');
    // Normalize both tickers for comparison
    var normalizedStored = String(storedTicker || '').toUpperCase().trim();
    var normalizedCurrent = String(ticker || '').toUpperCase().trim();
    
    if (normalizedStored === normalizedCurrent) {
      var rawEps = props2.getProperty('LAST_EPS');
      var rawPe = props2.getProperty('LAST_PE');
      var rawBvps = props2.getProperty('LAST_BVPS');
      eps = (rawEps && !isNaN(Number(rawEps))) ? Number(rawEps) : null;
      pe = (rawPe && !isNaN(Number(rawPe))) ? Number(rawPe) : null;
      bvps = (rawBvps && !isNaN(Number(rawBvps))) ? Number(rawBvps) : null;
    }
  } catch(e) {}
  
  if (eps || bvps) {
    message += '\n📐 Graham Analysis\n';
    var grahamNumber = null;
    if (eps && eps > 0 && bvps && bvps > 0) {
      grahamNumber = Math.sqrt(22.5 * eps * bvps);
      var grahamMOS = ((grahamNumber - price) / grahamNumber) * 100;
      message += '📖 Graham #: $' + grahamNumber.toFixed(2) + '\n';
      message += '📉 vs Price: ' + grahamMOS.toFixed(1) + '%\n';
    } else if (eps && eps > 0) {
      var growthVal = growth || DEFAULT_GROWTH;
      var growthPct = growthVal * 100;
      var grahamSimple = eps * (8.5 + 2 * growthPct);
      message += '📖 Graham IV: $' + grahamSimple.toFixed(2) + ' (simplified)\n';
      var grahamMOS2 = ((grahamSimple - price) / grahamSimple) * 100;
      message += '📉 vs Price: ' + grahamMOS2.toFixed(1) + '%\n';
    } else {
      message += '⚠️ Need EPS for Graham #\n';
    }
    
    // Graham Criteria
    message += '\n✅ Graham Criteria\n';
    var criteria = 0;
    var peVal = pe || (price && eps && eps > 0 ? price / eps : null);
    if (peVal) {
      var pePass = peVal < 15;
      message += (pePass ? '✅' : '❌') + ' P/E: ' + peVal.toFixed(1) + ' (max 15)\n';
      if (pePass) criteria++;
    } else {
      message += '❓ P/E: N/A\n';
    }
    
    var pbVal = (price && bvps && bvps > 0) ? price / bvps : null;
    if (pbVal) {
      var pbPass = pbVal < 1.5;
      message += (pbPass ? '✅' : '❌') + ' P/B: ' + pbVal.toFixed(2) + ' (max 1.5)\n';
      if (pbPass) criteria++;
    } else {
      message += '❓ P/B: N/A\n';
    }
    
    if (peVal && pbVal) {
      var combined = peVal * pbVal;
      var combPass = combined < 22.5;
      message += (combPass ? '✅' : '❌') + ' PE×PB: ' + combined.toFixed(1) + ' (max 22.5)\n';
      if (combPass) criteria++;
    } else {
      message += '❓ PE×PB: N/A\n';
    }
    
    // 4th criterion: Price below Graham intrinsic value (grahamNumber or grahamSimple)
    var grahamIV = grahamNumber || (eps && eps > 0 ? eps * (8.5 + 2 * ((growth || DEFAULT_GROWTH) * 100)) : null);
    if (grahamIV && price) {
      var grahamMOSval = ((grahamIV - price) / grahamIV) * 100;
      var ivPass = price < grahamIV; // Price must be below Graham IV
      message += (ivPass ? '✅' : '❌') + ' Graham MOS: ' + grahamMOSval.toFixed(1) + '% (>0% = undervalued)\n';
      if (ivPass) criteria++;
    } else {
      message += '❓ Graham MOS: N/A\n';
    }
    
    message += '\n🎯 Score: ' + criteria + '/4 criteria\n';
  }
  
  // Add concise status if available
  if (statusStr && !isRateLimited) {
    var statusLines = statusStr.split('\n').filter(function(line) {
      return line && !line.startsWith('Status:');
    });
    if (statusLines.length > 0) {
      // Just show first 2 status lines
      message += '\n📋 ' + statusLines.slice(0, 2).join('\n   ');
    }
  }
  
  // Add recommendation
  message += '\n\n';
  if (!iv || !mos) {
    message += '⚠️ INCOMPLETE DATA\n';
    if (!fcf) message += 'Missing: FCF\n';
    if (!shares) message += 'Missing: Shares\n';
    message += 'Check your DCF sheet for details.';
  } else if (mos > 20) {
    message += '✅ UNDERVALUED\nConsider buying opportunity';
  } else if (mos > 0) {
    message += '➡️ FAIRLY VALUED\nWait for better entry';
  } else if (mos < -10) {
    message += '⚠️ OVERVALUED\nConsider taking profits';
  } else {
    message += '➡️ FAIRLY VALUED';
  }
  
  return {
    success: true,
    message: message,
    data: { ticker: ticker, price: price, iv: iv, mos: mos, fcf: fcf, shares: shares }
  };
}

/**
 * Agent Command: Run portfolio agent
 * Usage: AGENT
 */
function executeAgentCommand_(params) {
  var ss = SpreadsheetApp.getActive();
  var sig = ss.getSheetByName('Agent_Signals');
  
  if (!sig) {
    return { 
      success: false, 
      message: '❌ Agent not initialized.\n\nRun "Initialize Agent Sheets" from the menu first.' 
    };
  }
  
  // Enable silent mode
  SILENT_MODE = true;
  
  // Run agent
  try {
    RunAgentNow();
    Utilities.sleep(3000); // Wait for completion
  } catch (e) {
    SILENT_MODE = false;
    return {
      success: false,
      message: '❌ Agent run failed: ' + e.toString()
    };
  } finally {
    SILENT_MODE = false;
  }
  
  // Get latest signals
  var lastRow = sig.getLastRow();
  if (lastRow < 2) {
    return {
      success: true,
      message: '✅ Agent Run Complete\n\n📊 No new signals generated.\n\nYour portfolio looks good!'
    };
  }
  
  // Get last 10 signals
  var numSignals = Math.min(10, lastRow - 1);
  var signals = sig.getRange(lastRow - numSignals + 1, 1, numSignals, 8).getValues();
  
  // Count by action type
  var counts = { ADD: 0, TRIM: 0, CUT: 0, ALERT: 0 };
  for (var i = 0; i < signals.length; i++) {
    var action = String(signals[i][5] || '').toUpperCase();
    if (counts.hasOwnProperty(action)) counts[action]++;
  }
  
  var message = '🤖 Agent Run Complete\n';
  message += '═══════════════════\n';
  message += '📊 Total Signals: ' + numSignals + '\n\n';
  
  // Summary counts
  if (counts.ADD > 0) message += '✅ BUY/ADD: ' + counts.ADD + '\n';
  if (counts.TRIM > 0) message += '✂️ TRIM/REDUCE: ' + counts.TRIM + '\n';
  if (counts.CUT > 0) message += '❌ CUT/SELL: ' + counts.CUT + '\n';
  if (counts.ALERT > 0) message += '⚠️ ALERTS: ' + counts.ALERT + '\n';
  
  // Show top 5 signals
  message += '\n📋 Latest Signals:\n';
  for (var j = 0; j < Math.min(5, signals.length); j++) {
    var s = signals[j];
    var ticker = s[1];
    var action = s[5];
    var sizePct = s[6];
    var reason = s[7];
    
    message += '\n• ' + ticker + ': ' + action;
    if (sizePct) message += ' ' + sizePct + '%';
    message += '\n  ' + reason;
  }
  
  return {
    success: true,
    message: message,
    data: { signalCount: numSignals, counts: counts }
  };
}

/**
 * Status Command: Get portfolio overview
 * Usage: STATUS
 */
function executeStatusCommand_(params) {
  var ss = SpreadsheetApp.getActive();
  var wl = ss.getSheetByName('Agent_Watchlist');
  
  if (!wl) {
    return { 
      success: false, 
      message: '❌ Agent not initialized.\n\nRun "Initialize Agent Sheets" from the menu first.' 
    };
  }
  
  var cfg = getAgentConfig_();
  var lastRow = wl.getLastRow();
  var headers = wl.getRange(1,1,1,10).getValues()[0];
  var idx = {};
  headers.forEach(function(h,i){ idx[String(h).trim()] = i; });
  
  var data = wl.getRange(2,1,lastRow-1,10).getValues();
  
  var message = '📊 Portfolio Status\n';
  message += '═══════════════════\n';
  message += '📈 Macro Regime: ' + cfg.macro + '\n';
  message += '📁 Total Positions: ' + data.length + '\n\n';
  
  // Collect and sort positions by weight
  var positions = [];
  for (var i = 0; i < data.length; i++) {
    var ticker = String(data[i][idx['Ticker']] || '').trim();
    var weight = Number(data[i][idx['PositionWeight%']] || 0);
    if (ticker) {
      positions.push({ ticker: ticker, weight: weight });
    }
  }
  
  positions.sort(function(a, b) { return b.weight - a.weight; });
  
  // Show top holdings
  message += '🏆 Top Holdings:\n';
  for (var j = 0; j < Math.min(5, positions.length); j++) {
    message += (j+1) + '. ' + positions[j].ticker + ': ' + positions[j].weight.toFixed(1) + '%\n';
  }
  
  // Calculate concentration
  var top3Weight = 0;
  for (var k = 0; k < Math.min(3, positions.length); k++) {
    top3Weight += positions[k].weight;
  }
  
  message += '\n📊 Top 3 Concentration: ' + top3Weight.toFixed(1) + '%';
  
  return {
    success: true,
    message: message,
    data: { 
      macro: cfg.macro, 
      positions: positions.length,
      top3Concentration: top3Weight
    }
  };
}

/**
 * Price Command: Get current price
 * Usage: PRICE NVDA
 */
function executePriceCommand_(params) {
  if (params.length === 0) {
    return { 
      success: false, 
      message: '💰 Stock Price\n\nUsage: PRICE <TICKER>\n\nExamples:\n• PRICE NVDA\n• PRICE AAPL' 
    };
  }
  
  var ticker = params[0];
  var symbol = mapTickerAlpha_(ticker);
  
  try {
    var price = avFetchPrice_(symbol);
    
    if (!price || !isPos_(price)) {
      return { 
        success: false, 
        message: '❌ Could not fetch price for ' + ticker + '\n\nTicker may be invalid or API limit reached.' 
      };
    }
    
    var message = '💰 ' + ticker + ': $' + price.toFixed(2);
    
    return {
      success: true,
      message: message,
      data: { ticker: ticker, price: price }
    };
  } catch (e) {
    return {
      success: false,
      message: '❌ Error fetching price: ' + e.toString()
    };
  }
}

/**
 * Help Command: List available commands
 */
function executeHelpCommand_() {
  var message = '📱 Available Commands\n';
  message += '═══════════════════\n\n';
  
  message += '📊 DCF <TICKER>\n';
  message += 'Run DCF valuation analysis\n';
  message += 'Example: DCF AAPL\n';
  message += '⏱️ Takes ~30-40 seconds\n\n';
  
  message += '🤖 AGENT\n';
  message += 'Run portfolio rebalancing agent\n';
  message += 'Shows buy/sell signals\n\n';
  
  message += '📈 STATUS\n';
  message += 'Get portfolio overview\n';
  message += 'Shows top holdings & macro regime\n\n';
  
  message += '💰 PRICE <TICKER>\n';
  message += 'Get current stock price\n';
  message += 'Example: PRICE NVDA\n';
  message += '⚡ Fast (1 API call)\n\n';
  
  message += '🔍 DEBUG\n';
  message += 'Show what\'s in DCF sheet\n';
  message += 'Useful for troubleshooting\n\n';
  
  message += '❓ HELP\n';
  message += 'Show this message\n\n';
  
  message += '⚠️ IMPORTANT:\n';
  message += 'Wait 2 minutes between DCF commands!\n';
  message += 'Free API limit: 5 calls/minute';
  
  return {
    success: true,
    message: message
  };
}

/**
 * Debug Command: Show current DCF sheet values
 */
function executeDebugCommand_(params) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('DCF') || ss.getSheetByName('Sheet1');
  
  if (!sheet) {
    return { 
      success: false, 
      message: '❌ DCF sheet not found' 
    };
  }
  
  var ticker = sheet.getRange(TICKER_CELL).getValue();
  var discount = sheet.getRange(DISCOUNT_CELL).getValue();
  var growth = sheet.getRange(GROWTH_CELL).getValue();
  var fcf = sheet.getRange(FCF_CELL).getValue();
  var shares = sheet.getRange(SHARES_CELL).getValue();
  var price = sheet.getRange(PRICE_CELL).getValue();
  var iv = sheet.getRange(IV_CELL).getValue();
  var mos = sheet.getRange(MOS_CELL).getValue();
  var status = sheet.getRange(STATUS_CELL).getValue();
  var timestamp = sheet.getRange(TS_CELL).getValue();
  
  var message = '🔍 DCF Sheet Debug\n';
  message += '═══════════════════\n';
  message += 'Ticker: ' + (ticker || 'EMPTY') + '\n';
  message += 'Discount: ' + (discount || 'EMPTY') + '\n';
  message += 'Growth: ' + (growth || 'EMPTY') + '\n';
  message += 'FCF: ' + (fcf || 'EMPTY') + '\n';
  message += 'Shares: ' + (shares || 'EMPTY') + '\n';
  message += 'Price: ' + (price || 'EMPTY') + '\n';
  message += 'IV: ' + (iv || 'EMPTY') + '\n';
  message += 'MOS: ' + (mos || 'EMPTY') + '\n\n';
  
  if (status) {
    var statusStr = String(status).substring(0, 300);
    message += 'Status:\n' + statusStr;
  }
  
  if (timestamp) {
    message += '\n\nTimestamp:\n' + timestamp;
  }
  
  return {
    success: true,
    message: message
  };
}

// ===== MESSENGER API FUNCTIONS =====

/**
 * Check Facebook Messenger for new messages
 * This runs on a timer (every 1 minute by default)
 */
function CheckMessengerMessages() {
  if (!MESSENGER_CONFIG.enabled) {
    Logger.log('Messenger is disabled in config');
    return;
  }
  
  // Use lock to prevent concurrent executions
  var lock = LockService.getScriptLock();
  try {
    // Try to get lock, wait up to 10 seconds
    lock.tryLock(10000);
  } catch (e) {
    Logger.log('Could not acquire lock, another instance running');
    return;
  }
  
  try {
    // Fixed URL - simpler field syntax
    var url = 'https://graph.facebook.com/v18.0/me/conversations?access_token=' + 
              MESSENGER_CONFIG.pageAccessToken;
    
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var code = response.getResponseCode();
    
    if (code !== 200) {
      Logger.log('Messenger API error: HTTP ' + code);
      return;
    }
    
    var data = JSON.parse(response.getContentText());
    
    if (!data.data || data.data.length === 0) {
      Logger.log('No conversations found');
      return;
    }
    
    // Get the first conversation ID
    var conversationId = data.data[0].id;
    
    // Get messages from that conversation
    var msgUrl = 'https://graph.facebook.com/v18.0/' + conversationId + 
                 '?fields=messages&access_token=' + MESSENGER_CONFIG.pageAccessToken;
    
    var msgResponse = UrlFetchApp.fetch(msgUrl, { muteHttpExceptions: true });
    var msgData = JSON.parse(msgResponse.getContentText());
    
    if (!msgData.messages || !msgData.messages.data) {
      Logger.log('No messages in conversation');
      return;
    }
    
    var messages = msgData.messages.data;
    
    // Get last processed message ID
    var props = PropertiesService.getScriptProperties();
    var lastProcessed = props.getProperty('lastMessengerId');
    
    Logger.log('Checking messages. Last processed: ' + lastProcessed);
    
    // Process newest messages first
    for (var i = 0; i < messages.length; i++) {
      var msg = messages[i];
      
      // Skip if already processed
      if (msg.id === lastProcessed) {
        Logger.log('Reached last processed message');
        break;
      }
      
      // Get message details
      var detailUrl = 'https://graph.facebook.com/v18.0/' + msg.id + 
                      '?fields=message,from&access_token=' + MESSENGER_CONFIG.pageAccessToken;
      
      var detailResponse = UrlFetchApp.fetch(detailUrl, { muteHttpExceptions: true });
      var detail = JSON.parse(detailResponse.getContentText());
      
      // Only process messages from you (not from the page itself)
      if (!detail.from || detail.from.id !== MESSENGER_CONFIG.yourFacebookId) {
        Logger.log('Skipping message not from user: ' + (detail.from ? detail.from.id : 'unknown'));
        continue;
      }
      
      Logger.log('Processing message: ' + detail.message);
      
      // CRITICAL: Mark as processed IMMEDIATELY before doing anything else
      props.setProperty('lastMessengerId', msg.id);
      Logger.log('Marked message as processed: ' + msg.id);
      
      // Show typing indicator
      if (MESSENGER_CONFIG.sendTypingIndicator) {
        sendMessengerTyping_(true);
      }
      
      // Parse and execute command
      var commandText = detail.message || '';
      var parsed = parseCommand_(commandText);
      
      var result;
      if (!parsed.valid) {
        result = { message: parsed.error };
      } else {
        result = executeCommand_(parsed);
      }
      
      // Stop typing indicator
      if (MESSENGER_CONFIG.sendTypingIndicator) {
        sendMessengerTyping_(false);
      }
      
      // Send reply
      sendMessengerReply_(result.message);
      
      Logger.log('Message processed and reply sent');
      
      // Only process one message per run to avoid timeouts
      break;
    }
    
  } catch (e) {
    Logger.log('CheckMessengerMessages error: ' + e.toString());
  } finally {
    // Always release the lock
    lock.releaseLock();
  }
}

/**
 * Send reply via Messenger
 */
function sendMessengerReply_(message) {
  try {
    var url = 'https://graph.facebook.com/v18.0/me/messages?' +
              'access_token=' + MESSENGER_CONFIG.pageAccessToken;
    
    var payload = {
      recipient: { id: MESSENGER_CONFIG.yourFacebookId },
      message: { text: message }
    };
    
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    
    if (code !== 200) {
      Logger.log('Send message error: HTTP ' + code);
      Logger.log('Response: ' + response.getContentText());
    } else {
      Logger.log('Message sent successfully');
    }
    
  } catch (e) {
    Logger.log('sendMessengerReply error: ' + e.toString());
  }
}

/**
 * Send typing indicator
 */
function sendMessengerTyping_(isTyping) {
  try {
    var url = 'https://graph.facebook.com/v18.0/me/messages?' +
              'access_token=' + MESSENGER_CONFIG.pageAccessToken;
    
    var payload = {
      recipient: { id: MESSENGER_CONFIG.yourFacebookId },
      sender_action: isTyping ? 'typing_on' : 'typing_off'
    };
    
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    UrlFetchApp.fetch(url, options);
    
  } catch (e) {
    Logger.log('sendMessengerTyping error: ' + e.toString());
  }
}

// ===== INSTALLATION & TESTING =====

/**
 * Install trigger to check Messenger every minute
 */
function InstallMessengerTrigger() {
  // Delete any existing triggers for this function
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'CheckMessengerMessages') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new trigger
  ScriptApp.newTrigger('CheckMessengerMessages')
    .timeBased()
    .everyMinutes(MESSENGER_CONFIG.checkIntervalMinutes)
    .create();
  
  Logger.log('Messenger trigger installed');
  
  SpreadsheetApp.getUi().alert(
    '✅ Messenger Trigger Installed!\n\n' +
    '⏱️ Checking every ' + MESSENGER_CONFIG.checkIntervalMinutes + ' minute(s)\n\n' +
    '📱 Message your Facebook page to test:\n' +
    '• DCF AAPL\n' +
    '• AGENT\n' +
    '• STATUS\n' +
    '• HELP\n\n' +
    'Replies within ' + MESSENGER_CONFIG.checkIntervalMinutes + ' minute(s)!'
  );
}

/**
 * Remove trigger
 */
function UninstallMessengerTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'CheckMessengerMessages') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  
  Logger.log('Removed ' + removed + ' trigger(s)');
  
  SpreadsheetApp.getUi().alert(
    '✅ Messenger Trigger Uninstalled\n\n' +
    'Removed ' + removed + ' trigger(s).'
  );
}

/**
 * Test Messenger connection
 */
function TestMessengerConnection() {
  if (!MESSENGER_CONFIG.enabled) {
    SpreadsheetApp.getUi().alert('❌ Messenger is disabled in MESSENGER_CONFIG');
    return;
  }
  
  if (!MESSENGER_CONFIG.pageAccessToken || MESSENGER_CONFIG.pageAccessToken.startsWith('EAAx')) {
    SpreadsheetApp.getUi().alert('❌ Please configure pageAccessToken in MESSENGER_CONFIG');
    return;
  }
  
  if (!MESSENGER_CONFIG.yourFacebookId || MESSENGER_CONFIG.yourFacebookId === '1234567890123456') {
    SpreadsheetApp.getUi().alert('❌ Please configure yourFacebookId in MESSENGER_CONFIG');
    return;
  }
  
  // Send test message
  var testMessage = '🎉 Connection Successful!\n\n' +
                   '✅ Your Messenger remote control is working!\n\n' +
                   '📱 Try these commands:\n' +
                   '• DCF AAPL\n' +
                   '• AGENT\n' +
                   '• STATUS\n' +
                   '• PRICE NVDA\n' +
                   '• HELP\n\n' +
                   'Powered by Google Apps Script 🚀';
  
  sendMessengerReply_(testMessage);
  
  SpreadsheetApp.getUi().alert(
    '✅ Test Message Sent!\n\n' +
    'Check your Messenger for the test message.\n\n' +
    'If you received it, your setup is complete!'
  );
}

/**
 * Test command parsing locally (for debugging)
 */
function TestCommandsParsing() {
  var testCommands = [
    'DCF AAPL',
    'AGENT',
    'STATUS',
    'PRICE NVDA',
    'HELP',
    'INVALID COMMAND'
  ];
  
  Logger.log('=== Testing Command Parsing ===');
  
  for (var i = 0; i < testCommands.length; i++) {
    var cmd = testCommands[i];
    var parsed = parseCommand_(cmd);
    
    Logger.log('\nCommand: ' + cmd);
    Logger.log('Valid: ' + parsed.valid);
    Logger.log('Parsed: ' + JSON.stringify(parsed));
  }
  
  SpreadsheetApp.getUi().alert('Command parsing test complete!\n\nCheck View > Executions for results.');
}

/**
 * Manual check for messages (for testing)
 */
function ManualCheckMessages() {
  CheckMessengerMessages();
  SpreadsheetApp.getUi().alert('Manual check complete!\n\nCheck View > Executions for logs.');
}

// ===== MENU INTEGRATION =====

/**
 * Show setup guide
 */
function ShowMessengerSetupGuide() {
  var html = '<h2>📱 Messenger Remote Control Setup</h2>';
  html += '<h3>Step 1: Create Facebook Page</h3>';
  html += '<p>Go to <a href="https://www.facebook.com/pages/create" target="_blank">facebook.com/pages/create</a></p>';
  html += '<p>Create a page (name: "My Investment Bot", category: Finance)</p>';
  
  html += '<h3>Step 2: Create Facebook App</h3>';
  html += '<p>Go to <a href="https://developers.facebook.com" target="_blank">developers.facebook.com</a></p>';
  html += '<p>Create App → Business type → Add Messenger product</p>';
  
  html += '<h3>Step 3: Get Credentials</h3>';
  html += '<p><strong>Page Access Token:</strong> In Messenger settings, under Access Tokens</p>';
  html += '<p><strong>Your Facebook ID:</strong> Use <a href="https://findmyfbid.com" target="_blank">findmyfbid.com</a></p>';
  
  html += '<h3>Step 4: Configure Script</h3>';
  html += '<p>Update MESSENGER_CONFIG at top of script with your credentials</p>';
  
  html += '<h3>Step 5: Install & Test</h3>';
  html += '<p>1. Click "Install Messenger Trigger"</p>';
  html += '<p>2. Click "Test Connection"</p>';
  html += '<p>3. Message your page: "HELP"</p>';
  
  html += '<h3>Commands</h3>';
  html += '<ul>';
  html += '<li><strong>DCF AAPL</strong> - Analyze stock</li>';
  html += '<li><strong>AGENT</strong> - Run portfolio agent</li>';
  html += '<li><strong>STATUS</strong> - Portfolio overview</li>';
  html += '<li><strong>PRICE NVDA</strong> - Get price</li>';
  html += '<li><strong>HELP</strong> - List commands</li>';
  html += '</ul>';
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📱 Messenger Setup Guide');
}
function TestTokenDirectly() {
  Logger.log('Testing with token: ' + MESSENGER_CONFIG.pageAccessToken.substring(0, 20) + '...');
  Logger.log('Testing with ID: ' + MESSENGER_CONFIG.yourFacebookId);
  
  var url = 'https://graph.facebook.com/v18.0/me/messages?access_token=' + 
            MESSENGER_CONFIG.pageAccessToken;
  
  var payload = {
    recipient: { id: MESSENGER_CONFIG.yourFacebookId },
    message: { text: '🧪 Direct test message!' }
  };
  
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  var text = response.getContentText();
  
  Logger.log('Response code: ' + code);
  Logger.log('Response: ' + text);
  
  if (code === 200) {
    SpreadsheetApp.getUi().alert('✅ SUCCESS! Check Messenger now!');
  } else {
    SpreadsheetApp.getUi().alert('❌ Error Code: ' + code + '\n\nCheck View > Executions for details');
  }
}
function GetMyPSID() {
  var url = 'https://graph.facebook.com/v18.0/me/conversations?access_token=' + 
            MESSENGER_CONFIG.pageAccessToken;
  
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var data = JSON.parse(response.getContentText());
  
  Logger.log('Response:');
  Logger.log(JSON.stringify(data, null, 2));
  
  SpreadsheetApp.getUi().alert('Check View > Executions for your PSID in the logs');
}
function GetParticipants() {
  var conversationId = 't_25721871087507278'; // From your response
  
  var url = 'https://graph.facebook.com/v18.0/' + conversationId + 
            '?fields=participants&access_token=' + MESSENGER_CONFIG.pageAccessToken;
  
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  var data = JSON.parse(response.getContentText());
  
  Logger.log('Participants:');
  Logger.log(JSON.stringify(data, null, 2));
  
  if (data.participants && data.participants.data) {
    data.participants.data.forEach(function(p) {
      Logger.log('Participant ID: ' + p.id + ' - Name: ' + p.name);
    });
  }
  
  SpreadsheetApp.getUi().alert('Check View > Executions for participant IDs');
}