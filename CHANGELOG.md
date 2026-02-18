# Changelog

All notable changes to the Claude Code Sessions projects.

## [Apps Script Bot 1.0.0] - 2026-02-17

### Added
- **DCF Analysis**
  - Real-time stock price fetching from Alpha Vantage
  - Free Cash Flow calculation
  - Intrinsic value calculation using DCF model
  - Margin of Safety computation
  
- **Graham Valuation**
  - Graham Number formula: √(22.5 × EPS × Book Value)
  - Simplified Graham formula for stocks without book value
  - Four-criteria scoring system (P/E, P/B, PE×PB, Graham MOS)
  - Conservative value investing analysis
  
- **API Integration**
  - Alpha Vantage primary API
  - Financial Modeling Prep backup API
  - Smart retry logic with exponential backoff
  - Rate limit detection and handling
  - Response caching (6 hour TTL)
  
- **Facebook Messenger Bot**
  - Remote command execution
  - Commands: DCF, AGENT, STATUS, PRICE, DEBUG, OVERVIEW, HELP
  - Time-triggered message polling
  - Duplicate message prevention with script locks
  
- **Portfolio Agent**
  - Automated rebalancing signals
  - Position sizing recommendations
  - Multiple risk scenarios (BASE, RECESSION, BEAR, UPSIDE)
  
### Fixed
- **Feb 17 Morning**: Fixed rate limit handling - removed wasteful retries on 429 responses
- **Feb 17 Afternoon**: Fixed duplicate Messenger responses - script lock + early message marking
- **Feb 17 Evening**: Fixed stale data between ticker changes - EPS/PE/BVPS now stored in ScriptProperties
- **Feb 17 Night**: Fixed Graham MOS showing DCF MOS - now correctly calculates Graham-based MOS

### Technical Details
- **Throttling**: 1.2 second delay between API calls (~50 req/hour)
- **Retry Logic**: Exponential backoff on server errors (500s), no retry on rate limits (429)
- **Cache Strategy**: Document cache with ticker-specific keys
- **Execution Time**: ~30 seconds per DCF analysis from Messenger
- **API Quota**: 3-4 calls per DCF analysis

### Known Limitations
- Graham criteria very strict (by design) - most growth stocks fail
- FMP free tier doesn't include cash flow data (requires premium)
- 30-second response time in Messenger due to API calls
- Alpha Vantage free tier: 5 calls/min, 500 calls/day

## Future Plans
- [ ] Add more valuation methods (P/E multiples, PEG ratio)
- [ ] Historical performance tracking
- [ ] Email notifications for agent signals
- [ ] Google Drive integration for reports
- [ ] Multi-currency support
- [ ] Dividend discount model (DDM)
