# Apps Script Investment Bot

A comprehensive Google Sheets investment analysis tool with Facebook Messenger integration.

## Features

### 📊 DCF Analysis
- Fetches real-time stock data from Alpha Vantage API
- Calculates Discounted Cash Flow intrinsic value
- Computes Margin of Safety (MOS)
- Fallback to Financial Modeling Prep API

### 📐 Graham Valuation
- Benjamin Graham's conservative value investing formula
- Graham Number: √(22.5 × EPS × Book Value)
- Simplified formula for stocks without book value
- 4-criteria scoring system:
  - P/E < 15
  - P/B < 1.5
  - PE × PB < 22.5
  - Graham MOS > 0% (price below intrinsic value)

### 🤖 Portfolio Agent
- Automated rebalancing signals
- Buy/sell/hold recommendations
- Position sizing based on conviction
- Configurable risk scenarios (BASE, RECESSION, BEAR, UPSIDE)

### 💬 Messenger Integration
- Remote command execution
- Real-time analysis via Facebook Messenger
- Status updates and alerts
- No need to open Google Sheets

## Installation

### 1. Create Google Sheets Structure

Create a new Google Sheet with these tabs:

#### DCF Sheet
| A2 | B2 | C2 | D2 | E2 | F2 | G2 | H2 | I2 | J2 | K2 | L2 | M2 |
|----|----|----|----|----|----|----|----|----|----|----|----|----|
| Ticker | Discount Rate | Growth | Years | FCF | Shares | Terminal Growth | Price | IV | MOS | EPS | P/E | (reserved) |

Example values:
- B2: 0.10 (10% discount rate)
- C2: 0.05 (5% growth)
- D2: 10 (years)
- G2: 0.03 (3% terminal growth)

#### Agent Sheets
- `Agent_Signals` - Buy/sell signals
- `Agent_Holdings` - Current positions
- `Agent_Config` - Risk scenarios and allocation

### 2. Set Up Apps Script

1. Open **Extensions** → **Apps Script**
2. Delete default `Code.gs` content
3. Copy entire contents of `Code.gs` from this folder
4. **Save** (Ctrl+S)

### 3. Configure API Keys

Edit these lines at the top of `Code.gs`:

```javascript
var ALPHAVANTAGE_API_KEY = 'YOUR_KEY_HERE';  // Get from https://www.alphavantage.co/support/#api-key
var FMP_API_KEY = 'YOUR_KEY_HERE';           // Optional: https://financialmodelingprep.com
var USE_FMP_BACKUP = true;                    // Enable FMP fallback
```

### 4. (Optional) Set Up Messenger Bot

1. Create a Facebook Page
2. Create a Facebook App with Messenger integration
3. Get Page Access Token
4. Configure webhook

Edit Messenger config in `Code.gs`:

```javascript
var MESSENGER_CONFIG = {
  enabled: true,
  pageAccessToken: 'YOUR_PAGE_TOKEN',
  yourFacebookId: 'YOUR_PSID'
};
```

5. Set up trigger:
   - **Triggers** (clock icon in left sidebar)
   - **Add Trigger**
   - Function: `CheckMessengerMessages`
   - Event: Time-driven, Minutes timer, Every 1 minute

## Usage

### From Google Sheets

1. Enter ticker in cell A2 (e.g., `AAPL`)
2. Click **DCF Menu** → **Calculate Intrinsic Value**
3. Results appear in columns H-J:
   - H2: Current Price
   - I2: Intrinsic Value
   - J2: Margin of Safety %
   - K2: EPS
   - L2: P/E Ratio

### From Messenger

Send commands to your Facebook Page:

```
DCF AAPL
```

Response:
```
📊 AAPL Analysis
═══════════════════
📐 DCF Valuation
💰 Price:    $227.63
💎 IV (DCF): $121.81
📈 MOS:       -86.9%

📐 Graham Analysis
📖 Graham #: $23.21
📉 vs Price: -880.8%

✅ Graham Criteria
❌ P/E: 35.8 (max 15)
❌ P/B: 60.38 (max 1.5)
❌ PE×PB: 2161.6 (max 22.5)
❌ Graham MOS: -880.8% (>0% = undervalued)

🎯 Score: 0/4 criteria

💵 Fundamentals
FCF: $95.00B
Growth: 6.8%
EPS: $6.35
Book Value: $3.77

🔴 OVERVALUED
Consider taking profits
```

## API Rate Limits

### Alpha Vantage (Free Tier)
- **5 calls per minute**
- **500 calls per day**
- Each DCF uses 3-4 calls (Price, Overview, Cash Flow)
- Smart retry logic avoids wasting quota

### Financial Modeling Prep (Free Tier)
- **250 calls per day**
- Used as fallback for price and shares data
- Cash flow data requires premium ($14/month)

### Combined Capacity
- ~200-300 price checks per day
- ~30-50 full DCF analyses per day

## Troubleshooting

### "No data received"
- **Rate limited**: Wait 2-3 minutes between runs
- **Invalid ticker**: Check ticker symbol is correct
- **API timeout**: Try again in a moment

### "EPS/PE/BVPS showing N/A"
- Some stocks don't have this data in Alpha Vantage
- Use `OVERVIEW <TICKER>` command to see raw API response
- Graham analysis will use simplified formula without book value

### Messenger not responding
- Check trigger is set up (Apps Script → Triggers)
- Verify Messenger config has correct tokens
- Check execution logs for errors

## Development Notes

### Code Structure
- Lines 1-50: Configuration and global variables
- Lines 51-400: API helper functions (Alpha Vantage, FMP)
- Lines 401-800: DCF calculation logic
- Lines 801-1200: Agent/portfolio functions
- Lines 1201-1500: Command execution (DCF, AGENT, STATUS, etc.)
- Lines 1501-2000: Messenger integration
- Lines 2001+: Utility functions

### Key Functions
- `calculateIntrinsicValueAlphaVantage()` - Main DCF calculation
- `executeDCFCommand_()` - Handles Messenger DCF requests
- `avFetchPrice_()`, `avFetchOverview_()`, `avFetchFCF_()` - API calls
- `CheckMessengerMessages()` - Polls for new Messenger commands

### Caching
- API responses cached in `CacheService` (6 hours)
- Reduces API calls for repeated ticker lookups
- Cache keys: `AV_PQ_<TICKER>`, `AV_OV_<TICKER>`, `AV_CF_<TICKER>`

### Error Handling
- Silent mode for trigger-based execution
- Retry logic with exponential backoff
- Rate limit detection and friendly error messages

## Version History

See parent `CHANGELOG.md` for detailed version history.

### Current Version: 1.0.0 (Feb 2026)
- ✅ DCF analysis with Alpha Vantage
- ✅ Graham Number valuation
- ✅ FMP backup API
- ✅ Messenger bot integration
- ✅ Portfolio agent
- ✅ Smart retry logic

## Contributing

This is a personal project but feel free to fork and modify!

## License

MIT License
