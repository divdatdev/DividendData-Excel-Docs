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

### Quick Reference

All functions use the `DIVIDENDDATA.` namespace in Excel.

| Function | Description |
|----------|-------------|
| [`DIVIDENDDATA.DIVIDENDS`](#1-dividenddatadividends) | Dividend data (payouts, yields, history, growth) |
| [`DIVIDENDDATA.DIVIDENDS_BATCH`](#2-dividenddatadividends_batch) | Batch dividend data for multiple stocks |
| [`DIVIDENDDATA.STATEMENT`](#3-dividenddatastatement) | Financial statements (income, balance, cash flow) |
| [`DIVIDENDDATA.METRICS`](#4-dividenddatametrics) | Individual financial metrics (revenue, EPS, etc.) |
| [`DIVIDENDDATA.RATIOS`](#5-dividenddataratios) | Financial ratios (P/E, current ratio, ROE, etc.) |
| [`DIVIDENDDATA.GROWTH`](#6-dividenddatagrowth) | Growth metrics (revenue growth, EPS growth, etc.) |
| [`DIVIDENDDATA.QUOTE`](#7-dividenddataquote) | Stock quotes (price, volume, history) |
| [`DIVIDENDDATA.QUOTE_BATCH`](#8-dividenddataquote_batch) | Batch quotes for multiple stocks |
| [`DIVIDENDDATA.PROFILE`](#9-dividenddataprofile) | Company profiles (market cap, sector, CEO, etc.) |
| [`DIVIDENDDATA.FUND`](#10-dividenddatafund) | ETF/Fund data (holdings, expense ratio) |
| [`DIVIDENDDATA.ESTIMATES`](#11-dividenddataestimates) | Analyst estimates (EPS, revenue forecasts) |
| [`DIVIDENDDATA.PRICE_TARGET`](#12-dividenddataprice_target) | Analyst price targets |
| [`DIVIDENDDATA.SEGMENTS`](#13-dividenddatasegments) | Revenue segments (product, geographic) |
| [`DIVIDENDDATA.KPIS`](#14-dividenddatakpis) | Key performance indicators |
| [`DIVIDENDDATA.COMMODITIES`](#15-dividenddatacommodities) | Commodities data (oil, gold, etc.) |
| [`DIVIDENDDATA.CRYPTO`](#16-dividenddatacrypto) | Cryptocurrency data |

---

## Installation

Getting started is easy. You can install the add-in directly from the Microsoft AppSource marketplace or search for it inside Excel.

### Option 1: One-Click Install (Recommended)
The fastest way to install is via our direct marketplace listing.

1.  **Click this link:** [**Dividend Data on Microsoft AppSource**](https://marketplace.microsoft.com/en-us/product/WA200010010)
2.  Click the blue **"Get it now"** button.
3.  Sign in with your Microsoft account if prompted.
4.  Click **"Open in Excel"** to launch a new spreadsheet with the add-in pre-loaded.

### Option 2: Install Inside Excel
You can also find the add-in without leaving your spreadsheet.

**For Excel Desktop (Windows/Mac):**
1.  Open Excel and go to the **Home** tab on the top ribbon.
2.  Click the **Add-ins** button (sometimes labeled as "Get Add-ins" or a grid icon).
3.  Click **"More Add-ins"** if a menu appears.
4.  In the store window, search for **"Dividend Data"**.
5.  Locate our logo and click **Add**.

**For Excel on the Web:**
1.  Go to the **Home** tab.
2.  Click the **Add-ins** icon on the far right of the ribbon.
3.  Search for **"Dividend Data"** and click **Add**.

### ✅ Verification
Once installed, a **"Dividend Data"** tab may appear in your ribbon, or you will receive a notification that the add-in is loaded. To test it immediately, type this formula into any empty cell:

### 🚀 Next Steps
Open the Sidebar: Click the Dividend Data tab in the ribbon (or the add-in icon) and select "Open Sidebar" to sign in and access quick tools.


---
## Function Reference


## 1) DIVIDENDDATA.DIVIDENDS

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

## 2) DIVIDENDDATA.DIVIDENDS_BATCH

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

## 3) DIVIDENDDATA.STATEMENT

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

## 4) DIVIDENDDATA.METRICS

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

* `"Revenue"` (Revenue)
* `"CostOfRevenue"` (Cost of Revenue)
* `"GrossProfit"` (Gross Profit)
* `"ResearchAndDevelopmentExpenses"` (Research and Development Expenses)
* `"GeneralAndAdministrativeExpenses"` (General and Administrative Expenses)
* `"SellingAndMarketingExpenses"` (Selling and Marketing Expenses)
* `"SellingGeneralAndAdministrativeExpenses"` (Selling General and Administrative Expenses)
* `"OtherExpenses"` (Other Expenses)
* `"OperatingExpenses"` (Operating Expenses)
* `"CostAndExpenses"` (Cost and Expenses)
* `"NetInterestIncome"` (Net Interest Income)
* `"InterestIncome"` (Interest Income)
* `"InterestExpense"` (Interest Expense)
* `"DepreciationAndAmortization"` (Depreciation and Amortization)
* `"Ebitda"` (Ebitda)
* `"Ebit"` (Ebit)
* `"NonOperatingIncomeExcludingInterest"` (Non Operating Income Excluding Interest)
* `"OperatingIncome"` (Operating Income)
* `"TotalOtherIncomeExpensesNet"` (Total Other Income Expenses Net)
* `"IncomeBeforeTax"` (Income Before Tax)
* `"IncomeTaxExpense"` (Income Tax Expense)
* `"NetIncomeFromContinuingOperations"` (Net Income From Continuing Operations)
* `"NetIncomeFromDiscontinuedOperations"` (Net Income From Discontinued Operations)
* `"OtherAdjustmentsToNetIncome"` (Other Adjustments To Net Income)
* `"NetIncome"` (Net Income)
* `"NetIncomeDeductions"` (Net Income Deductions)
* `"BottomLineNetIncome"` (Bottom Line Net Income)
* `"Eps"` (Eps)
* `"EpsDiluted"` (Eps Diluted)
* `"WeightedAverageShsOut"` (Weighted Average Shs Out)
* `"WeightedAverageShsOutDil"` (Weighted Average Shs Out Dil)
* `"CashAndCashEquivalents"` (Cash and Cash Equivalents)
* `"ShortTermInvestments"` (Short Term Investments)
* `"CashAndShortTermInvestments"` (Cash and Short Term Investments)
* `"NetReceivables"` (Net Receivables)
* `"AccountsReceivables"` (Accounts Receivables)
* `"OtherReceivables"` (Other Receivables)
* `"Inventory"` (Inventory)
* `"Prepaids"` (Prepaids)
* `"OtherCurrentAssets"` (Other Current Assets)
* `"TotalCurrentAssets"` (Total Current Assets)
* `"PropertyPlantEquipmentNet"` (Property Plant Equipment Net)
* `"Goodwill"` (Goodwill)
* `"IntangibleAssets"` (Intangible Assets)
* `"GoodwillAndIntangibleAssets"` (Goodwill and Intangible Assets)
* `"LongTermInvestments"` (Long Term Investments)
* `"TaxAssets"` (Tax Assets)
* `"OtherNonCurrentAssets"` (Other Non Current Assets)
* `"TotalNonCurrentAssets"` (Total Non Current Assets)
* `"OtherAssets"` (Other Assets)
* `"TotalAssets"` (Total Assets)
* `"TotalPayables"` (Total Payables)
* `"AccountPayables"` (Account Payables)
* `"OtherPayables"` (Other Payables)
* `"AccruedExpenses"` (Accrued Expenses)
* `"ShortTermDebt"` (Short Term Debt)
* `"CapitalLeaseObligationsCurrent"` (Capital Lease Obligations Current)
* `"TaxPayables"` (Tax Payables)
* `"DeferredRevenue"` (Deferred Revenue)
* `"OtherCurrentLiabilities"` (Other Current Liabilities)
* `"TotalCurrentLiabilities"` (Total Current Liabilities)
* `"LongTermDebt"` (Long Term Debt)
* `"CapitalLeaseObligationsNonCurrent"` (Capital Lease Obligations Non Current)
* `"DeferredRevenueNonCurrent"` (Deferred Revenue Non Current)
* `"DeferredTaxLiabilitiesNonCurrent"` (Deferred Tax Liabilities Non Current)
* `"OtherNonCurrentLiabilities"` (Other Non Current Liabilities)
* `"TotalNonCurrentLiabilities"` (Total Non Current Liabilities)
* `"OtherLiabilities"` (Other Liabilities)
* `"CapitalLeaseObligations"` (Capital Lease Obligations)
* `"TotalLiabilities"` (Total Liabilities)
* `"TreasuryStock"` (Treasury Stock)
* `"PreferredStock"` (Preferred Stock)
* `"CommonStock"` (Common Stock)
* `"RetainedEarnings"` (Retained Earnings)
* `"AdditionalPaidInCapital"` (Additional Paid In Capital)
* `"AccumulatedOtherComprehensiveIncomeLoss"` (Accumulated Other Comprehensive Income Loss)
* `"OtherTotalStockholdersEquity"` (Other Total Stockholders Equity)
* `"TotalStockholdersEquity"` (Total Stockholders Equity)
* `"TotalEquity"` (Total Equity)
* `"MinorityInterest"` (Minority Interest)
* `"TotalLiabilitiesAndTotalEquity"` (Total Liabilities and Total Equity)
* `"TotalInvestments"` (Total Investments)
* `"TotalDebt"` (Total Debt)
* `"NetDebt"` (Net Debt)
* `"NetIncome"` (Net Income)
* `"DepreciationAndAmortization"` (Depreciation and Amortization)
* `"DeferredIncomeTax"` (Deferred Income Tax)
* `"StockBasedCompensation"` (Stock Based Compensation)
* `"ChangeInWorkingCapital"` (Change In Working Capital)
* `"AccountsReceivables"` (Accounts Receivables)
* `"Inventory"` (Inventory)
* `"AccountsPayables"` (Accounts Payables)
* `"OtherWorkingCapital"` (Other Working Capital)
* `"OtherNonCashItems"` (Other Non Cash Items)
* `"NetCashProvidedByOperatingActivities"` (Net Cash Provided By Operating Activities)
* `"InvestmentsInPropertyPlantAndEquipment"` (Investments In Property Plant and Equipment)
* `"AcquisitionsNet"` (Acquisitions Net)
* `"PurchasesOfInvestments"` (Purchases of Investments)
* `"SalesMaturitiesOfInvestments"` (Sales Maturities of Investments)
* `"OtherInvestingActivities"` (Other Investing Activities)
* `"NetCashProvidedByInvestingActivities"` (Net Cash Provided By Investing Activities)
* `"NetDebtIssuance"` (Net Debt Issuance)
* `"LongTermNetDebtIssuance"` (Long Term Net Debt Issuance)
* `"ShortTermNetDebtIssuance"` (Short Term Net Debt Issuance)
* `"NetStockIssuance"` (Net Stock Issuance)
* `"NetCommonStockIssuance"` (Net Common Stock Issuance)
* `"CommonStockIssuance"` (Common Stock Issuance)
* `"CommonStockRepurchased"` (Common Stock Repurchased)
* `"NetPreferredStockIssuance"` (Net Preferred Stock Issuance)
* `"NetDividendsPaid"` (Net Dividends Paid)
* `"CommonDividendsPaid"` (Common Dividends Paid)
* `"PreferredDividendsPaid"` (Preferred Dividends Paid)
* `"OtherFinancingActivities"` (Other Financing Activities)
* `"NetCashProvidedByFinancingActivities"` (Net Cash Provided By Financing Activities)
* `"EffectOfForexChangesOnCash"` (Effect of Forex Changes On Cash)
* `"NetChangeInCash"` (Net Change In Cash)
* `"CashAtEndOfPeriod"` (Cash At End of Period)
* `"CashAtBeginningOfPeriod"` (Cash At Beginning of Period)
* `"OperatingCashFlow"` (Operating Cash Flow)
* `"CapitalExpenditure"` (Capital Expenditure)
* `"FreeCashFlow"` (Free Cash Flow)
* `"IncomeTaxesPaid"` (Income Taxes Paid)
* `"InterestPaid"` (Interest Paid)

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

## 5) DIVIDENDDATA.RATIOS

Retrieves financial ratios and valuation metrics.

### Syntax
```excel
=DIVIDENDDATA.RATIOS(symbol, metric, [showHeaders], [period], [year])
```

### Popular Metrics

_All available metrics:_
* `"RevenuePerShare"` (Revenue Per Share)
* `"NetIncomePerShare"` (Net Income Per Share)
* `"OperatingCashFlowPerShare"` (Operating Cash Flow Per Share)
* `"FreeCashFlowPerShare"` (Free Cash Flow Per Share)
* `"CashPerShare"` (Cash Per Share)
* `"BookValuePerShare"` (Book Value Per Share)
* `"TangibleBookValuePerShare"` (Tangible Book Value Per Share)
* `"ShareholdersEquityPerShare"` (Shareholders Equity Per Share)
* `"InterestDebtPerShare"` (Interest Debt Per Share)
* `"MarketCap"` (Market Cap)
* `"EnterpriseValue"` (Enterprise Value)
* `"PeRatio"` (Pe Ratio)
* `"PriceToSalesRatio"` (Price To Sales Ratio)
* `"Pocfratio"` (Price to Operating Cash Flow Ratio)
* `"PfcfRatio"` (Price to Free Cash Flow Ratio)
* `"PbRatio"` (Price Book Value Ratio)
* `"PtbRatio"` (Price to Book Ratio)
* `"EvToSales"` (Enterprise Value to Sales)
* `"EnterpriseValueOverEBITDA"` (Enterprise Value Over EBITDA)
* `"EvToOperatingCashFlow"` (Enterprise Value To Operating Cash Flow)
* `"EvToFreeCashFlow"` (Enterprise Value To Free Cash Flow)
* `"EarningsYield"` (Earnings Yield)
* `"FreeCashFlowYield"` (Free Cash Flow Yield)
* `"DebtToEquity"` (Debt To Equity)
* `"DebtToAssets"` (Debt To Assets)
* `"NetDebtToEBITDA"` (Net Debt To EBITDA)
* `"CurrentRatio"` (Current Ratio)
* `"InterestCoverage"` (Interest Coverage)
* `"IncomeQuality"` (Income Quality)
* `"DividendYield"` (Dividend Yield)
* `"PayoutRatio"` (Payout Ratio)
* `"SalesGeneralAndAdministrativeToRevenue"` (Sales General And Administrative To Revenue)
* `"ResearchAndDevelopmentToRevenue"` (Research And Development To Revenue)
* `"IntangiblesToTotalAssets"` (Intangibles To Total Assets)
* `"CapexToOperatingCashFlow"` (Capex To Operating Cash Flow)
* `"CapexToRevenue"` (Capex To Revenue)
* `"CapexToDepreciation"` (Capex To Depreciation)
* `"StockBasedCompensationToRevenue"` (Stock Based Compensation To Revenue)
* `"GrahamNumber"` (Graham Number)
* `"Roic"` (Return on Investing Capital)
* `"ReturnOnTangibleAssets"` (Return On Tangible Assets)
* `"GrahamNetNet"` (Graham Net Net)
* `"WorkingCapital"` (Working Capital)
* `"TangibleAssetValue"` (Tangible Asset Value)
* `"NetCurrentAssetValue"` (Net Current Asset Value)
* `"InvestedCapital"` (Invested Capital)
* `"AverageReceivables"` (Average Receivables)
* `"AveragePayables"` (Average Payables)
* `"AverageInventory"` (Average Inventory)
* `"DaysSalesOutstanding"` (Days Sales Outstanding)
* `"DaysPayablesOutstanding"` (Days Payables Outstanding)
* `"DaysOfInventoryOnHand"` (Days Of Inventory On Hand)
* `"ReceivablesTurnover"` (Receivables Turnover)
* `"PayablesTurnover"` (Payables Turnover)
* `"InventoryTurnover"` (Inventory Turnover)
* `"Roe"` (Return on Equity)
* `"CapexPerShare"` (Capex Per Share)
* `"CurrentRatio"` (Current Ratio)
* `"QuickRatio"` (Quick Ratio)
* `"CashRatio"` (Cash Ratio)
* `"DaysOfSalesOutstanding"` (Days Of Sales Outstanding)
* `"DaysOfInventoryOutstanding"` (Days Of Inventory Outstanding)
* `"OperatingCycle"` (Operating Cycle)
* `"DaysOfPayablesOutstanding"` (Days Of Payables Outstanding)
* `"CashConversionCycle"` (Cash Conversion Cycle)
* `"GrossProfitMargin"` (Gross Profit Margin)
* `"OperatingProfitMargin"` (Operating Profit Margin)
* `"PretaxProfitMargin"` (Pretax Profit Margin)
* `"NetProfitMargin"` (Net Profit Margin)
* `"EffectiveTaxRate"` (Effective Tax Rate)
* `"ReturnOnAssets"` (Return On Assets)
* `"ReturnOnEquity"` (Return On Equity)
* `"ReturnOnCapitalEmployed"` (Return On Capital Employed)
* `"NetIncomePerEBT"` (Net Income Per EBT)
* `"EbtPerEbit"` (Ebt Per Ebit)
* `"EbitPerRevenue"` (Ebit Per Revenue)
* `"DebtRatio"` (Debt Ratio)
* `"DebtEquityRatio"` (Debt Equity Ratio)
* `"LongTermDebtToCapitalization"` (Long Term Debt To Capitalization)
* `"TotalDebtToCapitalization"` (Total Debt To Capitalization)
* `"InterestCoverage"` (Interest Coverage)
* `"CashFlowCoverageRatios"` (Cash Flow Coverage Ratios)
* `"ShortTermCoverageRatios"` (Short Term Coverage Ratios)
* `"CapitalExpenditureCoverageRatio"` (Capital Expenditure Coverage Ratio)
* `"DividendPaidAndCapexCoverageRatio"` (Dividend Paid And Capex Coverage Ratio)
* `"DividendPayoutRatio"` (Dividend Payout Ratio)
* `"PriceBookValueRatio"` (Price Book Value Ratio)
* `"PriceToBookRatio"` (Price To Book Ratio)
* `"PriceToSalesRatio"` (Price To Sales Ratio)
* `"PriceEarningsRatio"` (Price Earnings Ratio)
* `"PriceToFreeCashFlowsRatio"` (Price To Free Cash Flows Ratio)
* `"PriceToOperatingCashFlowsRatio"` (Price To Operating Cash Flows Ratio)
* `"PriceCashFlowRatio"` (Price Cash Flow Ratio)
* `"PriceEarningsToGrowthRatio"` (Price Earnings To Growth Ratio)
* `"PriceSalesRatio"` (Price Sales Ratio)
* `"DividendYield"` (Dividend Yield)
* `"EnterpriseValueMultiple"` (Enterprise Value Multiple)
* `"PriceFairValue"` (Price Fair Value)

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

## 6) DIVIDENDDATA.GROWTH

Retrieves growth metrics for financial figures.

### Syntax
```excel
=DIVIDENDDATA.GROWTH(symbol, metric, [showHeaders], [period], [year])
```

### Popular Metrics
_All available metrics:_
* `"RevenueGrowth"` (Revenue Growth)
* `"GrossProfitGrowth"` (Gross Profit Growth)
* `"Ebitgrowth"` (Ebitgrowth)
* `"OperatingIncomeGrowth"` (Operating Income Growth)
* `"NetIncomeGrowth"` (Net Income Growth)
* `"Epsgrowth"` (GAAP EPS growth)
* `"EpsdilutedGrowth"` (EPS diluted Growth)
* `"WeightedAverageSharesGrowth"` (Weighted Average Shares Growth)
* `"WeightedAverageSharesDilutedGrowth"` (Weighted Average Shares Diluted Growth)
* `"DividendsPerShareGrowth"` (Dividends Per Share Growth)
* `"OperatingCashFlowGrowth"` (Operating Cash Flow Growth)
* `"ReceivablesGrowth"` (Receivables Growth)
* `"InventoryGrowth"` (Inventory Growth)
* `"AssetGrowth"` (Asset Growth)
* `"BookValuePerShareGrowth"` (Book Value Per Share Growth)
* `"DebtGrowth"` (Debt Growth)
* `"RdexpenseGrowth"` (Rdexpense Growth)
* `"SgaexpensesGrowth"` (Sgaexpenses Growth)
* `"FreeCashFlowGrowth"` (Free Cash Flow Growth)
* `"TenYRevenueGrowthPerShare"` (Ten Y Revenue Growth Per Share)
* `"FiveYRevenueGrowthPerShare"` (Five Y Revenue Growth Per Share)
* `"ThreeYRevenueGrowthPerShare"` (Three Y Revenue Growth Per Share)
* `"TenYOperatingCFGrowthPerShare"` (Ten Y Operating Cash Flow Growth Per Share)
* `"FiveYOperatingCFGrowthPerShare"` (Five Y Operating Cash Flow Growth Per Share)
* `"ThreeYOperatingCFGrowthPerShare"` (Three Y Operating Cash Flow Growth Per Share)
* `"TenYNetIncomeGrowthPerShare"` (Ten Y Net Income Growth Per Share)
* `"FiveYNetIncomeGrowthPerShare"` (Five Y Net Income Growth Per Share)
* `"ThreeYNetIncomeGrowthPerShare"` (Three Y Net Income Growth Per Share)
* `"TenYShareholdersEquityGrowthPerShare"` (Ten Y Shareholders Equity Growth Per Share)
* `"FiveYShareholdersEquityGrowthPerShare"` (Five Y Shareholders Equity Growth Per Share)
* `"ThreeYShareholdersEquityGrowthPerShare"` (Three Y Shareholders Equity Growth Per Share)
* `"TenYDividendPerShareGrowthPerShare"` (Ten Y Dividend Per Share Growth Per Share)
* `"FiveYDividendPerShareGrowthPerShare"` (Five Y Dividend Per Share Growth Per Share)
* `"ThreeYDividendPerShareGrowthPerShare"` (Three Y Dividend Per Share Growth Per Share)
* `"EbitdaGrowth"` (Ebitda Growth)
* `"GrowthCapitalExpenditure"` (Growth Capital Expenditure)
* `"TenYBottomLineNetIncomeGrowthPerShare"` (Ten Y Bottom Line Net Income Growth Per Share)
* `"FiveYBottomLineNetIncomeGrowthPerShare"` (Five Y Bottom Line Net Income Growth Per Share)
* `"ThreeYBottomLineNetIncomeGrowthPerShare"` (Three Y Bottom Line Net Income Growth Per Share)

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

## 7) DIVIDENDDATA.QUOTE

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

## 8) DIVIDENDDATA.QUOTE_BATCH

Retrieves quotes for multiple stocks at once. Perfect for watchlists.

### Syntax
```excel
=DIVIDENDDATA.QUOTE_BATCH(symbols, [metrics], [showHeaders])
```

### Parameters

**symbols**: Comma-separated tickers (e.g., `"AAPL,MSFT"`). Required.

**metrics**: Comma-separated metrics or `"all"`: `"price"`, `"change"`, `"volume"` (default: `"all"`).

**showHeaders**: Include headers and symbol column (defaults based on metrics).


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

## 9) DIVIDENDDATA.PROFILE

Retrieves company profile and overview information.

### Syntax
```excel
=DIVIDENDDATA.PROFILE(symbol, [metric], [showHeaders])
```
### Description
Retrieves company profile information. Useful for overview details like market cap, sector, or description.

### Parameters

**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: Specific profile metric or `"full"`.

_All available metrics:_
* `"symbol"` (Symbol)
* `"price"` (Price)
* `"marketcap"` (Market Cap)
* `"beta"` (Beta)
* `"lastdividend"` (Last Dividend)
* `"range"` (Range)
* `"change"` (Change)
* `"changepercentage"` (Change Percentage)
* `"volume"` (Volume)
* `"averagevolume"` (Average Volume)
* `"companyname"` (Company Name)
* `"currency"` (Currency)
* `"cik"` (Cik)
* `"isin"` (Isin)
* `"cusip"` (Cusip)
* `"exchangefullname"` (Exchange Full Name)
* `"exchange"` (Exchange)
* `"industry"` (Industry)
* `"website"` (Website)
* `"description"` (Description)
* `"ceo"` (Ceo)
* `"sector"` (Sector)
* `"country"` (Country)
* `"fulltimeemployees"` (Full Time Employees)
* `"phone"` (Phone)
* `"address"` (Address)
* `"city"` (City)
* `"state"` (State)
* `"zip"` (Zip)
* `"image"` (Image)
* `"ipodate"` (Ipo Date)
* `"defaultimage"` (Default Image)
* `"isetf"` (Is Etf)
* `"isactivelytrading"` (Is Actively Trading)
* `"isadr"` (Is ADR?)
* `"isfund"` (Is Fund?)
* `"full"` (Return Full Profile)

**showHeaders**: Include headers for `"full"` (default: `false`).

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

## 10) DIVIDENDDATA.FUND

Retrieves ETF and mutual fund data.

### Syntax
```excel
=DIVIDENDDATA.FUND(symbol, [metric], [showHeaders])
```
### Description
Retrieves data for ETFs or mutual funds. Useful for fund analysis, including holdings, expense ratios, or sector exposures.

### Parameters

**symbol**: Fund ticker symbol (e.g., `"SPY"`). Required.

**metric**: Fund metrics like `"holdings"`, `"expenseRatio"`, etc.

_All available metrics:_
* `"etfcompany"` (Etf Company)
* `"expenseratio"` (Expense Ratio)
* `"assetsundermanagement"` (Assets Under Management)
* `"holdings"` (Holdings)
* `"nav"` (Net Asset Value)
* `"navcurrency"` (NAV Currency)
* `"holdingscount"` (Holdings Count)
* `"countryweighting"` (Country Weighting)
* `"sectorslist"` (Sectors List)
* `"description"` (Description)
* `"assetclass"` (Asset Class)
* `"domicile"` (Domicile)
* `"website"` (Website)
* `"avgvolume"` (Avg Volume)
* `"inceptiondate"` (Inception Date)
* `"updatedat"` (Updated At)

**showHeaders**: Include headers for tables (default: `false`).

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

## 11) DIVIDENDDATA.ESTIMATES

### Description
Retrieves analyst estimates for earnings and revenue. Returns a summary table of future estimates, specific metric history (estimates vs actuals), or single next-year values for quick watchlist building. Useful for forecasting future growth, valuation models, and gauging analyst expectations.

### Parameters
**symbol**: Stock ticker symbol (e.g., `"MSFT"`). Required.

**metric**: The estimate metric to retrieve (default: `"summary"`).

_Available metrics:_
* `"summary"` (Returns a full table including: Date, Avg EPS Estimate, EPS Growth, High/Low EPS Estimates, Avg Revenue Estimate, Revenue Growth, High/Low Revenue Estimates, and Analyst Counts).
* `"eps"` (Returns: Date, Avg EPS Estimate, EPS Growth).
* `"revenue"` (Returns: Date, Avg Revenue Estimate, Revenue Growth).
* `"next_year_eps"` (Returns a single value for the nearest future year's estimated EPS).
* `"next_year_revenue"` (Returns a single value for the nearest future year's estimated Revenue).
* `"next_year_eps_growth"` (Returns a single value for next year's estimated EPS growth %).
* `"next_year_revenue_growth"` (Returns a single value for next year's estimated Revenue growth %).

**period**: Period: `"annual"` or `"quarter"` (default: `"annual"`).

**showHeaders**: Include headers for tables (default: `true`).

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

## 12) DIVIDENDDATA.PRICE_TARGET

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

## 13) DIVIDENDDATA.SEGMENTS

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

## 14) DIVIDENDDATA.KPIS

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

## 15) DIVIDENDDATA.COMMODITIES

Retrieves commodities data.

### Syntax
```excel
=DIVIDENDDATA.COMMODITIES([symbol], metric, [fromDate], [toDate], [showHeaders])
```

### Description
Retrieves commodities data, such as prices or history. Useful for tracking commodity markets like oil or gold.

### Parameters

**symbol**: Commodity symbol (e.g., `"CLUSD"`). Required except for `"list"`.

_Full list of symbols:_
* `"ZQUSD"` (30 Day Fed Fund Futures)
* `"ALIUSD"` (Aluminum Futures)
* `"ZBUSD"` (30 Year U.S. Treasury Bond Futures)
* `"ZOUSX"` (Oat Futures)
* `"ESUSD"` (E-Mini S&P 500)
* `"ZMUSD"` (Soybean Meal Futures)
* `"GCUSD"` (Gold Futures)
* `"ZLUSX"` (Soybean Oil Futures)
* `"KEUSX"` (Wheat Futures)
* `"ZFUSD"` (Five-Year US Treasury Note)
* `"SILUSD"` (Micro Silver Futures)
* `"ZCUSX"` (Corn Futures)
* `"HEUSX"` (Lean Hogs Futures)
* `"PLUSD"` (Platinum)
* `"HGUSD"` (Copper)
* `"MGCUSD"` (Micro Gold Futures)
* `"SBUSX"` (Sugar)
* `"SIUSD"` (Silver Futures)
* `"CTUSX"` (Cotton)
* `"DXUSD"` (US Dollar)
* `"ZSUSX"` (Soybean Futures)
* `"LBUSD"` (Lumber Futures)
* `"LEUSX"` (Live Cattle Futures)
* `"NGUSD"` (Natural Gas)
* `"CLUSD"` (Crude Oil)
* `"OJUSX"` (Orange Juice)
* `"KCUSX"` (Coffee)
* `"PAUSD"` (Palladium)
* `"GFUSX"` (Feeder Cattle Futures)
* `"ZTUSD"` (2-Year T-Note Futures)
* `"ZRUSD"` (Rough Rice Futures)
* `"CCUSD"` (Cocoa)
* `"NQUSD"` (Nasdaq 100)
* `"ZNUSD"` (10-Year T-Note Futures)
* `"RTYUSD"` (Micro E-mini Russell 2000 Index Futures)
* `"BZUSD"` (Brent Crude Oil)
* `"DCUSD"` (Class III Milk Futures)
* `"YMUSD"` (Mini Dow Jones Industrial Average Index)
* `"RBUSD"` (Gasoline RBOB)
* `"HOUSD"` (Heating Oil)

**metric**: Metric: `"list"`, `"price"`, `"fullquote"`, `"history"`. Required.

**fromDate**: Start date for history.

**toDate**: End date for history.

**showHeaders**: Include headers for tables.

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

## 16) DIVIDENDDATA.CRYPTO

Retrieves cryptocurrency data.

### Description
Retrieves cryptocurrency data, such as prices or history. Useful for tracking crypto markets like Bitcoin or Ethereum.

### Parameters
**symbol**: Cryptocurrency symbol (e.g., "BTCUSD"). Required except for "list".

_Full list of symbols (examples, may vary):_
* `"BTCUSD"` (Bitcoin)
* `"ETHUSD"` (Ethereum)
* `"BNBUSD"` (Binance Coin)
* `"XRPUSD"` (Ripple)
* `"ADAUSD"` (Cardano)
* `"SOLUSD"` (Solana)
* `"DOGEUSD"` (Dogecoin)
* `"LINKUSD"` (Chainlink)

**metric**: Metric: `"list"`, `"price"`, `"fullquote"`, `"history"`. Required.

**fromDate**: Start date for history (`"YYYY-MM-DD"`).

**toDate**: End date for history (`"YYYY-MM-DD"`).

**showHeaders**: Include headers for tables (default: `true`).

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

Function names are consistent across both platforms. The only difference is the separator: Excel uses `.` (dot) and Google Sheets uses `_` (underscore). This is a platform requirement — not a Dividend Data choice.

**The rule: same function name, swap the separator.**

| Function | Excel | Google Sheets |
|----------|-------|--------------|
| Dividends | `=DIVIDENDDATA.DIVIDENDS(...)` | `=DIVIDENDDATA_DIVIDENDS(...)` |
| Dividends Batch | `=DIVIDENDDATA.DIVIDENDS_BATCH(...)` | `=DIVIDENDDATA_DIVIDENDS_BATCH(...)` |
| Statement | `=DIVIDENDDATA.STATEMENT(...)` | `=DIVIDENDDATA_STATEMENT(...)` |
| Metrics | `=DIVIDENDDATA.METRICS(...)` | `=DIVIDENDDATA_METRICS(...)` |
| Ratios | `=DIVIDENDDATA.RATIOS(...)` | `=DIVIDENDDATA_RATIOS(...)` |
| Growth | `=DIVIDENDDATA.GROWTH(...)` | `=DIVIDENDDATA_GROWTH(...)` |
| Quote | `=DIVIDENDDATA.QUOTE(...)` | `=DIVIDENDDATA_QUOTE(...)` |
| Quote Batch | `=DIVIDENDDATA.QUOTE_BATCH(...)` | `=DIVIDENDDATA_QUOTE_BATCH(...)` |
| Profile | `=DIVIDENDDATA.PROFILE(...)` | `=DIVIDENDDATA_PROFILE(...)` |
| Fund | `=DIVIDENDDATA.FUND(...)` | `=DIVIDENDDATA_FUND(...)` |
| Estimates | `=DIVIDENDDATA.ESTIMATES(...)` | `=DIVIDENDDATA_ESTIMATES(...)` |
| Segments | `=DIVIDENDDATA.SEGMENTS(...)` | `=DIVIDENDDATA_SEGMENTS(...)` |
| KPIs | `=DIVIDENDDATA.KPIS(...)` | `=DIVIDENDDATA_KPIS(...)` |
| Commodities | `=DIVIDENDDATA.COMMODITIES(...)` | `=DIVIDENDDATA_COMMODITIES(...)` |
| Crypto | `=DIVIDENDDATA.CRYPTO(...)` | `=DIVIDENDDATA_CRYPTO(...)` |
| Price Target | `=DIVIDENDDATA.PRICE_TARGET(...)` | `=DIVIDENDDATA_PRICE_TARGET(...)` |

Parameters, metrics, and outputs are identical across both platforms.

Also available for Google Sheets! See the [Google Sheets Add-in Documentation](https://github.com/divdatdev/Dividend-Data-Spreadsheet-Docs).

---

## Support

- **Help Center:** [help.dividenddata.com](https://help.dividenddata.com)
- **Email:** support@dividenddata.com
- **YouTube Tutorials:** [@DividendData](https://www.youtube.com/@DividendData)

---

<p align="center">
  <strong>Built for dividend investors, by dividend investors.</strong>
</p>
