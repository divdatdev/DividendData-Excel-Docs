# Dividend Data - Excel Add-in

<p align="center">
  <img src="https://raw.githubusercontent.com/divdatdev/Dividend-Data-Spreadsheet-Docs/main/logo/LogoNew.png" alt="Dividend Data Logo" width="120"/>
</p>

<p align="center">
  <strong>Access 30+ years of financial data directly in Microsoft Excel.</strong>
</p>

<p align="center">
  <a href="https://www.dividenddata.com">Website</a> •
  <a href="https://help.dividenddata.com">Help Center</a> •
  <a href="https://www.youtube.com/@DividendData">YouTube</a>
</p>

---

## Why Dividend Data?

**Stop copying and pasting financial data.** The Dividend Data Excel Add-in brings institutional-quality financial data directly into your spreadsheets with simple formulas.

### Built for Dividend Investors
- **30+ years of dividend history** for thousands of stocks
- **Dividend growth rates** (1, 3, 5, 10-year CAGRs)
- **Payout ratios** (EPS and Free Cash Flow based)
- **Ex-dividend dates, payment dates, frequencies**

### Complete Fundamental Data
- **Financial statements** — Income, Balance Sheet, Cash Flow (annual & quarterly)
- **100+ financial metrics** — Revenue, EPS, Free Cash Flow, and more
- **50+ financial ratios** — P/E, ROE, Debt/Equity, Current Ratio, etc.
- **Growth rates** — Revenue growth, EPS growth, dividend growth

### Real-Time & Historical Quotes
- **Live stock prices** with 50 & 200-day moving averages
- **Historical price data** going back decades
- **Batch quotes** for multiple tickers at once

### Advanced Analytics
- **Analyst estimates** — EPS and revenue forecasts
- **Price targets** — Consensus, high, low, median
- **Revenue segments** — Product and geographic breakdowns
- **ETF data** — Holdings, expense ratios, sector allocations

### Use Cases
- **Dividend tracking spreadsheets** — Monitor your portfolio's income
- **Stock screeners** — Build custom screens with real data
- **Valuation models** — DCF, DDM, and comparable analysis
- **Research dashboards** — Track metrics over time
- **Watchlists** — Live prices and key metrics in one view

---

## Function Reference

All functions use the `DIVIDENDDATA.` namespace in Excel.

### Quick Reference

| Function | Description |
|----------|-------------|
| `DIVIDENDDATA.DIVIDENDS` | Dividend data (payouts, yields, history, growth) |
| `DIVIDENDDATA.DIVIDENDS_BATCH` | Batch dividend data for multiple stocks |
| `DIVIDENDDATA.STATEMENT` | Financial statements (income, balance, cash flow) |
| `DIVIDENDDATA.METRICS` | Individual financial metrics (revenue, EPS, etc.) |
| `DIVIDENDDATA.RATIOS` | Financial ratios (P/E, current ratio, ROE, etc.) |
| `DIVIDENDDATA.GROWTH` | Growth metrics (revenue growth, EPS growth, etc.) |
| `DIVIDENDDATA.QUOTE` | Stock quotes (price, volume, history) |
| `DIVIDENDDATA.QUOTE_BATCH` | Batch quotes for multiple stocks |
| `DIVIDENDDATA.PROFILE` | Company profiles (market cap, sector, CEO, etc.) |
| `DIVIDENDDATA.FUND` | ETF/Fund data (holdings, expense ratio) |
| `DIVIDENDDATA.ESTIMATES` | Analyst estimates (EPS, revenue forecasts) |
| `DIVIDENDDATA.SEGMENTS` | Revenue segments (product, geographic) |
| `DIVIDENDDATA.KPIS` | Key performance indicators |
| `DIVIDENDDATA.COMMODITIES` | Commodities data (oil, gold, etc.) |
| `DIVIDENDDATA.CRYPTO` | Cryptocurrency data |
| `DIVIDENDDATA.PRICE_TARGET` | Analyst price targets |

---

## DIVIDENDDATA.DIVIDENDS

Retrieves dividend-related data for a stock including payouts, yields, history, and growth rates.

### Syntax
```excel
=DIVIDENDDATA.DIVIDENDS(symbol, [metric], [showHeaders])
```

### Parameters
| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `symbol` | Yes | - | Stock ticker (e.g., "MSFT") |
| `metric` | No | "fwd_payout" | The dividend metric to retrieve |
| `showHeaders` | No | FALSE | Include headers in table results |

### Available Metrics
| Metric | Description | Example Output |
|--------|-------------|----------------|
| `"fwd_payout"` | Forward annual dividend | 3.32 |
| `"ttm_payout"` | Trailing 12-month payout | 3.00 |
| `"fwd_yield"` | Forward dividend yield | 0.0082 (0.82%) |
| `"ttm_yield"` | TTM dividend yield | 0.0075 (0.75%) |
| `"next_dividend"` | Next declared dividend amount | 0.83 |
| `"next_dividend_full"` | Full next dividend details | Table with amount, dates |
| `"frequency"` | Payment frequency | "Quarterly" |
| `"ex_div_date"` | Most recent ex-dividend date | 2025-02-18 |
| `"payment_date"` | Most recent payment date | 2025-03-11 |
| `"declaration_date"` | Most recent declaration date | 2025-01-15 |
| `"history"` | Full dividend history | Table |
| `"summary"` | Dividend summary | Table |
| `"growth"` | Dividend growth rates | Table |
| `"1y_cagr"` | 1-year CAGR | 0.10 (10%) |
| `"3y_cagr"` | 3-year CAGR | 0.08 (8%) |
| `"5y_cagr"` | 5-year CAGR | 0.09 (9%) |
| `"10y_cagr"` | 10-year CAGR | 0.11 (11%) |
| `"payout_ratio"` | Payout ratio (EPS-based) | 0.28 (28%) |
| `"fcf_payout_ratio"` | Free cash flow payout ratio | 0.35 (35%) |

### Examples

**Get Microsoft's forward dividend yield:**
```excel
=DIVIDENDDATA.DIVIDENDS("MSFT", "fwd_yield")
```
→ Returns: `0.0082` (0.82%)

**Get Coca-Cola's 5-year dividend growth rate:**
```excel
=DIVIDENDDATA.DIVIDENDS("KO", "5y_cagr")
```
→ Returns: `0.035` (3.5% annual growth)

**Get Johnson & Johnson's next ex-dividend date:**
```excel
=DIVIDENDDATA.DIVIDENDS("JNJ", "ex_div_date")
```
→ Returns: `2025-02-24`

**Get Realty Income's payment frequency:**
```excel
=DIVIDENDDATA.DIVIDENDS("O", "frequency")
```
→ Returns: `Monthly`

**Get Apple's full dividend history with headers:**
```excel
=DIVIDENDDATA.DIVIDENDS("AAPL", "history", TRUE)
```
→ Returns table:
| Declaration Date | Record Date | Payment Date | Dividend | Adj Dividend | Growth Rate |
|-----------------|-------------|--------------|----------|--------------|-------------|
| 2025-01-30 | 2025-02-10 | 2025-02-13 | 0.25 | 0.25 | 0.04 |
| 2024-10-31 | 2024-11-11 | 2024-11-14 | 0.25 | 0.25 | 0.00 |
| ... | ... | ... | ... | ... | ... |

**Get Procter & Gamble's dividend summary:**
```excel
=DIVIDENDDATA.DIVIDENDS("PG", "summary")
```
→ Returns table with: Annual Dividend (FWD), Yield (FWD), Annual Dividend (TTM), Yield (TTM), Frequency, Next Dividend, Declaration Date, Ex-Div Date, Payment Date, 1Y/3Y/5Y/10Y CAGRs, Payout Ratios

---

## DIVIDENDDATA.DIVIDENDS_BATCH

Retrieves dividend data for multiple stocks at once. Efficient for building dividend watchlists and portfolio trackers.

### Syntax
```excel
=DIVIDENDDATA.DIVIDENDS_BATCH(symbols, [metric], [showHeaders])
```

### Parameters
| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `symbols` | Yes | - | Comma-separated tickers (e.g., "MSFT,AAPL,KO") |
| `metric` | No | "fwd_payout" | Metric(s) to retrieve (comma-separated) |
| `showHeaders` | No | TRUE | Include headers and symbol column |

### Available Metrics
- `"all"` — All dividend metrics
- `"history"` — Combined dividend history for all symbols
- `"fwd_payout"` — Forward annual payouts
- `"adjdividend"` — Most recent adjusted dividend
- `"dividend"` — Most recent declared dividend
- `"record_date"` / `"recorddate"` — Record dates
- `"declaration_date"` / `"declarationdate"` — Declaration dates
- `"payment_date"` / `"paymentdate"` — Payment dates
- `"yield"` — Dividend yields
- `"frequency"` — Payment frequencies

### Examples

**Get forward yields for a dividend portfolio:**
```excel
=DIVIDENDDATA.DIVIDENDS_BATCH("JNJ,PG,KO,PEP,MCD", "fwd_payout,yield", TRUE)
```
→ Returns:
| Symbol | Fwd Payout | Yield |
|--------|-----------|-------|
| JNJ | 5.08 | 0.032 |
| PG | 4.03 | 0.024 |
| KO | 2.00 | 0.031 |
| PEP | 5.42 | 0.030 |
| MCD | 7.08 | 0.023 |

**Get all dividend data for REIT portfolio:**
```excel
=DIVIDENDDATA.DIVIDENDS_BATCH("O,MAIN,STAG,NNN", "all")
```
→ Returns complete dividend data including yields, payouts, frequencies, dates, and growth rates

**Compare ex-dividend dates:**
```excel
=DIVIDENDDATA.DIVIDENDS_BATCH("AAPL,MSFT,GOOGL,AMZN,META", "ex_div_date,payment_date,dividend")
```
→ Returns table with upcoming dividend dates and amounts

---

## DIVIDENDDATA.STATEMENT

Retrieves full financial statements (income statement, balance sheet, cash flow statement).

### Syntax
```excel
=DIVIDENDDATA.STATEMENT(symbol, metric, [showHeaders], [period], [year])
```

### Parameters
| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `symbol` | Yes | - | Stock ticker |
| `metric` | Yes | - | `"income"`, `"balance"`, or `"cash_flow"` |
| `showHeaders` | No | FALSE | Include header row |
| `period` | No | "" | `"FY"`, `"Q1"`-`"Q4"`, `"annual"`, `"quarter"`, or `"ttm"` |
| `year` | No | "" | Specific year (e.g., 2024) |

### Examples

**Get Apple's 2024 annual income statement:**
```excel
=DIVIDENDDATA.STATEMENT("AAPL", "income", TRUE, "annual", 2024)
```
→ Returns full income statement with Revenue, Gross Profit, Operating Income, Net Income, EPS, etc.

**Get Microsoft's TTM cash flow statement:**
```excel
=DIVIDENDDATA.STATEMENT("MSFT", "cash_flow", TRUE, "ttm")
```
→ Returns: Operating Cash Flow, CapEx, Free Cash Flow, Dividends Paid, Share Repurchases, etc.

**Get Amazon's Q4 2024 balance sheet:**
```excel
=DIVIDENDDATA.STATEMENT("AMZN", "balance", TRUE, "Q4", 2024)
```
→ Returns: Total Assets, Total Liabilities, Cash, Debt, Shareholders' Equity, etc.

---

## DIVIDENDDATA.METRICS

Retrieves a specific metric from financial statements. Perfect for building custom financial models.

### Syntax
```excel
=DIVIDENDDATA.METRICS(symbol, metric, [showHeaders], [period], [year])
```

### Parameters
| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `symbol` | Yes | - | Stock ticker |
| `metric` | Yes | - | Financial metric (see list below) |
| `showHeaders` | No | FALSE | TRUE returns history table; FALSE returns latest value |
| `period` | No | "" | `"annual"`, `"quarter"`, or `"ttm"` |
| `year` | No | "" | Specific year filter |

### Popular Metrics

**Income Statement:**
- `Revenue`, `GrossProfit`, `OperatingIncome`, `NetIncome`
- `Eps`, `EpsDiluted`, `Ebitda`, `Ebit`
- `CostOfRevenue`, `OperatingExpenses`, `InterestExpense`

**Balance Sheet:**
- `TotalAssets`, `TotalLiabilities`, `TotalStockholdersEquity`
- `CashAndCashEquivalents`, `TotalDebt`, `NetDebt`
- `Inventory`, `AccountsReceivables`, `AccountPayables`

**Cash Flow:**
- `OperatingCashFlow`, `FreeCashFlow`, `CapitalExpenditure`
- `CommonDividendsPaid`, `CommonStockRepurchased`
- `NetChangeInCash`

### Examples

**Get NVIDIA's latest annual revenue:**
```excel
=DIVIDENDDATA.METRICS("NVDA", "Revenue")
```
→ Returns: `60922000000` ($60.9B)

**Get Apple's free cash flow history:**
```excel
=DIVIDENDDATA.METRICS("AAPL", "FreeCashFlow", TRUE, "annual")
```
→ Returns table:
| Date | Period | FreeCashFlow |
|------|--------|--------------|
| 2024-09-28 | FY | 108807000000 |
| 2023-09-30 | FY | 99584000000 |
| 2022-09-24 | FY | 111443000000 |
| ... | ... | ... |

**Get Microsoft's quarterly EPS for 2024:**
```excel
=DIVIDENDDATA.METRICS("MSFT", "EpsDiluted", TRUE, "quarter", 2024)
```
→ Returns quarterly diluted EPS values

**Get Johnson & Johnson's total debt:**
```excel
=DIVIDENDDATA.METRICS("JNJ", "TotalDebt")
```
→ Returns: `29850000000` ($29.85B)

**Get Coca-Cola's dividends paid:**
```excel
=DIVIDENDDATA.METRICS("KO", "CommonDividendsPaid", TRUE, "annual")
```
→ Returns annual dividend payment history

---

## DIVIDENDDATA.RATIOS

Retrieves financial ratios and valuation metrics.

### Syntax
```excel
=DIVIDENDDATA.RATIOS(symbol, metric, [showHeaders], [period], [year])
```

### Popular Metrics

**Valuation:**
- `PeRatio`, `PriceToSalesRatio`, `PbRatio`, `PfcfRatio`
- `EnterpriseValueOverEBITDA`, `EvToFreeCashFlow`
- `EarningsYield`, `FreeCashFlowYield`

**Profitability:**
- `GrossProfitMargin`, `OperatingProfitMargin`, `NetProfitMargin`
- `Roe` (Return on Equity), `Roic` (Return on Invested Capital)
- `ReturnOnAssets`

**Liquidity:**
- `CurrentRatio`, `QuickRatio`, `CashRatio`

**Leverage:**
- `DebtToEquity`, `DebtToAssets`, `InterestCoverage`
- `NetDebtToEBITDA`

**Dividends:**
- `DividendYield`, `PayoutRatio`, `DividendPayoutRatio`

### Examples

**Get Apple's P/E ratio:**
```excel
=DIVIDENDDATA.RATIOS("AAPL", "PeRatio")
```
→ Returns: `28.5`

**Get Microsoft's return on equity:**
```excel
=DIVIDENDDATA.RATIOS("MSFT", "Roe")
```
→ Returns: `0.35` (35%)

**Compare current ratios over time:**
```excel
=DIVIDENDDATA.RATIOS("WMT", "CurrentRatio", TRUE, "annual")
```
→ Returns table:
| Date | CurrentRatio |
|------|--------------|
| 2024-01-31 | 0.83 |
| 2023-01-31 | 0.82 |
| 2022-01-31 | 0.93 |

**Get AT&T's debt to equity:**
```excel
=DIVIDENDDATA.RATIOS("T", "DebtToEquity")
```
→ Returns: `1.42`

**Get Verizon's dividend payout ratio:**
```excel
=DIVIDENDDATA.RATIOS("VZ", "PayoutRatio")
```
→ Returns: `0.58` (58%)

---

## DIVIDENDDATA.GROWTH

Retrieves growth metrics for financial figures.

### Syntax
```excel
=DIVIDENDDATA.GROWTH(symbol, metric, [showHeaders], [period], [year])
```

### Popular Metrics
- `RevenueGrowth`, `GrossProfitGrowth`, `OperatingIncomeGrowth`, `NetIncomeGrowth`
- `EpsGrowth`, `EpsDilutedGrowth`, `DividendsPerShareGrowth`
- `FreeCashFlowGrowth`, `OperatingCashFlowGrowth`
- `BookValuePerShareGrowth`, `AssetGrowth`, `DebtGrowth`
- `FiveYRevenueGrowthPerShare`, `TenYRevenueGrowthPerShare`
- `FiveYDividendPerShareGrowthPerShare`

### Examples

**Get NVIDIA's revenue growth:**
```excel
=DIVIDENDDATA.GROWTH("NVDA", "RevenueGrowth")
```
→ Returns: `1.26` (126% YoY growth)

**Get Costco's EPS growth history:**
```excel
=DIVIDENDDATA.GROWTH("COST", "EpsGrowth", TRUE, "annual")
```
→ Returns table with annual EPS growth rates

**Get Home Depot's 5-year revenue growth per share:**
```excel
=DIVIDENDDATA.GROWTH("HD", "FiveYRevenueGrowthPerShare")
```
→ Returns: `0.08` (8% CAGR)

**Get Dividend Aristocrat's dividend growth:**
```excel
=DIVIDENDDATA.GROWTH("PG", "DividendsPerShareGrowth", TRUE, "annual")
```
→ Returns annual dividend growth history

---

## DIVIDENDDATA.QUOTE

Retrieves stock quote data including current price and historical prices.

### Syntax
```excel
=DIVIDENDDATA.QUOTE(symbol, [metric], [fromDate], [toDate], [showHeaders])
```

### Parameters
| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `symbol` | Yes | - | Stock ticker |
| `metric` | No | "price" | Quote metric |
| `fromDate` | No | - | Start date for history ("YYYY-MM-DD") |
| `toDate` | No | - | End date for history |
| `showHeaders` | No | TRUE | Include headers for tables |

### Available Metrics
| Metric | Description |
|--------|-------------|
| `"price"` | Current stock price |
| `"change"` | Price change (dollars) |
| `"volume"` | Trading volume |
| `"priceAvg50"` | 50-day moving average |
| `"priceAvg200"` | 200-day moving average |
| `"yearHigh"` | 52-week high |
| `"yearLow"` | 52-week low |
| `"full"` | All quote data (table) |
| `"history"` | Historical prices (table) |

### Examples

**Get current Apple stock price:**
```excel
=DIVIDENDDATA.QUOTE("AAPL", "price")
```
→ Returns: `185.92`

**Get Microsoft's 52-week high:**
```excel
=DIVIDENDDATA.QUOTE("MSFT", "yearHigh")
```
→ Returns: `468.35`

**Check if stock is above 200-day MA:**
```excel
=DIVIDENDDATA.QUOTE("SPY", "priceAvg200")
```
→ Returns: `542.18`

**Get Tesla's 2024 price history:**
```excel
=DIVIDENDDATA.QUOTE("TSLA", "history", "2024-01-01", "2024-12-31", TRUE)
```
→ Returns table with Date, Open, High, Low, Close, Volume

**Get full quote data:**
```excel
=DIVIDENDDATA.QUOTE("GOOGL", "full", , , TRUE)
```
→ Returns all quote metrics in a table

---

## DIVIDENDDATA.QUOTE_BATCH

Retrieves quotes for multiple stocks at once. Perfect for watchlists.

### Syntax
```excel
=DIVIDENDDATA.QUOTE_BATCH(symbols, [metrics], [showHeaders])
```

### Examples

**Build a simple watchlist:**
```excel
=DIVIDENDDATA.QUOTE_BATCH("AAPL,MSFT,GOOGL,AMZN,META", "price,change", TRUE)
```
→ Returns:
| Symbol | Price | Change |
|--------|-------|--------|
| AAPL | 185.92 | 2.34 |
| MSFT | 415.50 | -1.23 |
| GOOGL | 175.20 | 0.85 |
| AMZN | 198.45 | 3.12 |
| META | 585.30 | -2.45 |

**Get moving averages for dividend stocks:**
```excel
=DIVIDENDDATA.QUOTE_BATCH("JNJ,PG,KO,PEP", "price,priceAvg50,priceAvg200")
```
→ Returns prices with 50 and 200-day moving averages

---

## DIVIDENDDATA.PROFILE

Retrieves company profile and overview information.

### Syntax
```excel
=DIVIDENDDATA.PROFILE(symbol, [metric], [showHeaders])
```

### Available Metrics
- `"marketcap"`, `"price"`, `"beta"`, `"volume"`, `"averagevolume"`
- `"companyname"`, `"description"`, `"ceo"`, `"sector"`, `"industry"`
- `"exchange"`, `"country"`, `"website"`, `"fulltimeemployees"`
- `"ipodate"`, `"isetf"`, `"isfund"`, `"isadr"`
- `"full"` — Complete profile (table)

### Examples

**Get Apple's market cap:**
```excel
=DIVIDENDDATA.PROFILE("AAPL", "marketcap")
```
→ Returns: `2850000000000` ($2.85T)

**Get a company's sector:**
```excel
=DIVIDENDDATA.PROFILE("XOM", "sector")
```
→ Returns: `Energy`

**Get employee count:**
```excel
=DIVIDENDDATA.PROFILE("WMT", "fulltimeemployees")
```
→ Returns: `2100000`

**Get full company profile:**
```excel
=DIVIDENDDATA.PROFILE("MSFT", "full", TRUE)
```
→ Returns complete profile with all metrics

---

## DIVIDENDDATA.FUND

Retrieves ETF and mutual fund data.

### Syntax
```excel
=DIVIDENDDATA.FUND(symbol, [metric], [showHeaders])
```

### Available Metrics
- `"expenseratio"` — Expense ratio
- `"assetsundermanagement"` — Total AUM
- `"holdings"` — Top holdings (table)
- `"holdingscount"` — Number of holdings
- `"nav"` — Net asset value
- `"sectorslist"` — Sector breakdown (table)
- `"countryweighting"` — Country allocation (table)
- `"etfcompany"`, `"description"`, `"inceptiondate"`

### Examples

**Get SCHD's expense ratio:**
```excel
=DIVIDENDDATA.FUND("SCHD", "expenseratio")
```
→ Returns: `0.0006` (0.06%)

**Get VTI's top holdings:**
```excel
=DIVIDENDDATA.FUND("VTI", "holdings", TRUE)
```
→ Returns table with top holdings, weights, and share counts

**Get SPY's assets under management:**
```excel
=DIVIDENDDATA.FUND("SPY", "assetsundermanagement")
```
→ Returns: `550000000000` ($550B)

**Get QQQ's sector breakdown:**
```excel
=DIVIDENDDATA.FUND("QQQ", "sectorslist", TRUE)
```
→ Returns table with sector names and weights

---

## DIVIDENDDATA.ESTIMATES

Retrieves analyst estimates for earnings and revenue.

### Syntax
```excel
=DIVIDENDDATA.ESTIMATES(symbol, [metric], [period], [showHeaders])
```

### Available Metrics
| Metric | Description |
|--------|-------------|
| `"summary"` | Full estimates table |
| `"eps"` | EPS estimates history |
| `"revenue"` | Revenue estimates history |
| `"next_year_eps"` | Next year's estimated EPS (single value) |
| `"next_year_revenue"` | Next year's estimated revenue |
| `"next_year_eps_growth"` | Next year's EPS growth % |
| `"next_year_revenue_growth"` | Next year's revenue growth % |

### Examples

**Get NVIDIA's estimated EPS for next year:**
```excel
=DIVIDENDDATA.ESTIMATES("NVDA", "next_year_eps")
```
→ Returns: `4.25`

**Get Apple's estimated revenue growth:**
```excel
=DIVIDENDDATA.ESTIMATES("AAPL", "next_year_revenue_growth")
```
→ Returns: `0.08` (8% expected growth)

**Get full estimates summary:**
```excel
=DIVIDENDDATA.ESTIMATES("MSFT", "summary")
```
→ Returns table with EPS estimates, revenue estimates, growth rates, and analyst counts

---

## DIVIDENDDATA.PRICE_TARGET

Retrieves analyst price targets.

### Syntax
```excel
=DIVIDENDDATA.PRICE_TARGET(symbol, [metric], [showHeaders])
```

### Available Metrics
- `"summary"` — Price target summary table
- `"consensus"` — Consensus price target
- `"high"` — Highest analyst target
- `"low"` — Lowest analyst target
- `"median"` — Median target
- `"news"` — Recent price target news (table)

### Examples

**Get Tesla's consensus price target:**
```excel
=DIVIDENDDATA.PRICE_TARGET("TSLA", "consensus")
```
→ Returns: `275.00`

**Get the range of analyst targets:**
```excel
=DIVIDENDDATA.PRICE_TARGET("AAPL", "high")
```
→ Returns: `250.00`

```excel
=DIVIDENDDATA.PRICE_TARGET("AAPL", "low")
```
→ Returns: `180.00`

**Get full price target summary:**
```excel
=DIVIDENDDATA.PRICE_TARGET("AMZN", "summary", TRUE)
```
→ Returns table with consensus, high, low, median, and number of analysts

---

## DIVIDENDDATA.SEGMENTS

Retrieves revenue segmentation by product or geography.

### Syntax
```excel
=DIVIDENDDATA.SEGMENTS(symbol, metric, [period], [year], [showHeaders])
```

### Parameters
| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| `symbol` | Yes | - | Stock ticker |
| `metric` | Yes | - | `"products"` or `"geographic"` |
| `period` | No | "annual" | `"annual"` or `"quarter"` |
| `year` | No | "" | Specific year |
| `showHeaders` | No | FALSE | Include headers |

### Examples

**Get Apple's product revenue breakdown:**
```excel
=DIVIDENDDATA.SEGMENTS("AAPL", "products", "annual", 2024, TRUE)
```
→ Returns:
| Segment | Revenue |
|---------|---------|
| iPhone | 200000000000 |
| Services | 85000000000 |
| Mac | 29000000000 |
| iPad | 28000000000 |
| Wearables | 40000000000 |

**Get Microsoft's geographic revenue:**
```excel
=DIVIDENDDATA.SEGMENTS("MSFT", "geographic", "annual", 2024, TRUE)
```
→ Returns revenue by region (Americas, EMEA, Asia Pacific)

---

## DIVIDENDDATA.KPIS

Retrieves key performance indicators from Fiscal.ai — company-specific operational metrics not found in standard financial statements.

### Syntax
```excel
=DIVIDENDDATA.KPIS(ticker, [period], [year], [showHeaders])
```

### Examples

**Get Microsoft's KPIs (Azure revenue, Office 365 subscribers, etc.):**
```excel
=DIVIDENDDATA.KPIS("MSFT", "annual", 2024, TRUE)
```
→ Returns operational KPIs specific to Microsoft

**Get Netflix's subscriber metrics:**
```excel
=DIVIDENDDATA.KPIS("NFLX", "quarterly", , TRUE)
```
→ Returns subscriber counts, ARPU, and other streaming metrics

---

## DIVIDENDDATA.COMMODITIES

Retrieves commodities data.

### Syntax
```excel
=DIVIDENDDATA.COMMODITIES([symbol], metric, [fromDate], [toDate], [showHeaders])
```

### Popular Symbols
| Symbol | Commodity |
|--------|-----------|
| `"CLUSD"` | Crude Oil (WTI) |
| `"BZUSD"` | Brent Crude Oil |
| `"GCUSD"` | Gold |
| `"SIUSD"` | Silver |
| `"NGUSD"` | Natural Gas |
| `"HGUSD"` | Copper |

### Examples

**Get current gold price:**
```excel
=DIVIDENDDATA.COMMODITIES("GCUSD", "price")
```
→ Returns: `2045.30`

**Get crude oil price history:**
```excel
=DIVIDENDDATA.COMMODITIES("CLUSD", "history", "2024-01-01", "2024-12-31", TRUE)
```
→ Returns daily oil price history

**List all available commodities:**
```excel
=DIVIDENDDATA.COMMODITIES(, "list", , , TRUE)
```
→ Returns list of all commodity symbols and names

---

## DIVIDENDDATA.CRYPTO

Retrieves cryptocurrency data.

### Syntax
```excel
=DIVIDENDDATA.CRYPTO([symbol], metric, [fromDate], [toDate], [showHeaders])
```

### Popular Symbols
| Symbol | Cryptocurrency |
|--------|----------------|
| `"BTCUSD"` | Bitcoin |
| `"ETHUSD"` | Ethereum |
| `"SOLUSD"` | Solana |
| `"ADAUSD"` | Cardano |
| `"DOGEUSD"` | Dogecoin |

### Examples

**Get current Bitcoin price:**
```excel
=DIVIDENDDATA.CRYPTO("BTCUSD", "price")
```
→ Returns: `97500.00`

**Get Ethereum price history for 2024:**
```excel
=DIVIDENDDATA.CRYPTO("ETHUSD", "history", "2024-01-01", "2024-12-31", TRUE)
```
→ Returns daily ETH price history

**List all available cryptocurrencies:**
```excel
=DIVIDENDDATA.CRYPTO(, "list", , , TRUE)
```
→ Returns list of all crypto symbols

---

## Excel vs Google Sheets

If you're coming from the Google Sheets add-in, note these naming differences:

| Google Sheets | Excel |
|---------------|-------|
| `=DIVIDENDDATA("MSFT", ...)` | `=DIVIDENDDATA.DIVIDENDS("MSFT", ...)` |
| `=DIVIDENDDATA_BATCH(...)` | `=DIVIDENDDATA.DIVIDENDS_BATCH(...)` |
| `=DIVIDENDDATA_STATEMENT(...)` | `=DIVIDENDDATA.STATEMENT(...)` |
| `=DIVIDENDDATA_METRICS(...)` | `=DIVIDENDDATA.METRICS(...)` |
| `=DIVIDENDDATA_RATIOS(...)` | `=DIVIDENDDATA.RATIOS(...)` |
| `=DIVIDENDDATA_GROWTH(...)` | `=DIVIDENDDATA.GROWTH(...)` |
| `=DIVIDENDDATA_QUOTE(...)` | `=DIVIDENDDATA.QUOTE(...)` |
| `=DIVIDENDDATA_QUOTE_BATCH(...)` | `=DIVIDENDDATA.QUOTE_BATCH(...)` |
| `=DIVIDENDDATA_PROFILE(...)` | `=DIVIDENDDATA.PROFILE(...)` |
| `=DIVIDENDDATA_FUND(...)` | `=DIVIDENDDATA.FUND(...)` |
| `=DIVIDENDDATA_SEGMENTS(...)` | `=DIVIDENDDATA.SEGMENTS(...)` |
| `=DIVIDENDDATA_KPIS(...)` | `=DIVIDENDDATA.KPIS(...)` |
| `=DIVIDENDDATA_COMMODITIES(...)` | `=DIVIDENDDATA.COMMODITIES(...)` |
| `=DIVIDENDDATA_CRYPTO(...)` | `=DIVIDENDDATA.CRYPTO(...)` |
| `=DIVIDENDDATA_PRICE_TARGET(...)` | `=DIVIDENDDATA.PRICE_TARGET(...)` |
| `=DIVIDENDDATA_ESTIMATES(...)` | `=DIVIDENDDATA.ESTIMATES(...)` |

**Key difference:** Excel uses the `DIVIDENDDATA.` namespace prefix with shorter function names.

---

## Support

- **Help Center:** [help.dividenddata.com](https://help.dividenddata.com)
- **Email:** support@dividenddata.com
- **YouTube Tutorials:** [@DividendData](https://www.youtube.com/@DividendData)

---

<p align="center">
  <strong>Built for dividend investors, by dividend investors.</strong>
</p>
