# Claude Code Sessions

Code and projects developed during Claude AI coding sessions.

## Projects

### 📊 Apps Script Investment Bot
**Location:** `/apps-script-bot/`

A Google Apps Script bot that provides:
- **DCF (Discounted Cash Flow) Analysis** - Calculate intrinsic value of stocks
- **Graham Number Valuation** - Conservative value investing analysis based on Benjamin Graham's principles
- **Portfolio Agent** - Automated rebalancing signals
- **Facebook Messenger Integration** - Remote commands via Messenger

**Features:**
- Alpha Vantage API integration with FMP fallback
- Smart retry logic and rate limit handling
- Dual API backup system (250+ price checks/day)
- Real-time stock analysis via Messenger
- Graham criteria scoring (P/E, P/B, PE×PB, MOS)

**Commands:**
- `DCF <TICKER>` - Full valuation analysis with Graham scoring
- `AGENT` - Run portfolio rebalancing
- `STATUS` - Portfolio overview
- `PRICE <TICKER>` - Quick price check
- `DEBUG` - Show current sheet values
- `OVERVIEW <TICKER>` - Show raw API data
- `HELP` - List all commands

### 🐍 Python Projects
**Location:** `/python-projects/`

*Coming soon*

## Setup

### Apps Script Bot

1. Create a new Google Apps Script project
2. Copy contents of `apps-script-bot/Code.gs`
3. Set up your API keys in the config section:
   - Alpha Vantage API key
   - Financial Modeling Prep API key (optional)
   - Facebook Messenger credentials (optional)
4. Create required sheets: `DCF`, `Agent_Signals`, `Agent_Holdings`, `Agent_Config`
5. Set up time-triggered functions for Messenger polling

See `apps-script-bot/README.md` for detailed setup instructions.

## Version History

See `CHANGELOG.md` for detailed version history.

## License

MIT License - Feel free to use and modify!

## Author

Developed with assistance from Claude (Anthropic)
