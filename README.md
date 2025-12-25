# Finanial_Modeling_in_Excel
This project is a comprehensive real estate acquisition and investment analysis built in Excel

## Project Objectives
- Estimate property purchase and exit values using cap rate valuation
- Forecast rental income, operating expenses, and NOI
- Analyze investment profitability and risk
- Perform scenario, goal-seek, and sensitivity analysis
- Evaluate returns using **TVM and benchmark comparisons

## Tools & Techniques
- Microsoft Excel (Advanced Financial Modeling)
- Named Ranges for scalable formulas
- Excel Financial Functions:
  - `NPV`, `XNPV`
  - `IRR`, `XIRR`
  - `FV`, `PV`
- Lookup & Logic Functions:
  - `HLOOKUP`
  - `IF`
- Scenario Manager
- Goal Seek
- One-Variable & Two-Variable Data Tables
- Conditional Formatting

## Analysis Summary

### 1 Acquisition & Valuation
- Defined acquisition assumptions including target cap rate (6.00%)
- Calculated **purchase price** using the income capitalization formula:
  
  > Property Value = Net Operating Income ÷ Cap Rate

- Structured Year 0 acquisition cash flow as a negative investment
- Created named ranges to improve model clarity and accuracy

### 2 Operating Income Forecast
- Built a 10-year pro forma income statement
- Forecasted:
  - Rental income
  - Operating expenses
  - Capital expenditures
  - Net operating income (NOI)
  - Net income
- Applied rent growth (5%) and cost growth (2%)
- Converted monthly rent assumptions into annual net potential rent

### 3 Dynamic Sale & Exit Modeling
- Modeled a dynamic exit strategy using:
  - Holding period assumption
  - Exit cap rate
- Used `HLOOKUP()` to retrieve NOI for the selected sale year
- Calculated sale price using the cap rate formula
- Dynamically displayed sale proceeds in the correct year using `IF()`

### 4 Scenario Analysis
- Built multiple scenarios using Scenario Manager:
  - Expected Rent scenario
  - High Rent scenario
- Compared Year 10 NOI across scenarios
- Quantified differences to support investment decision-making

### 5 Goal Seek Optimization
- Used Goal Seek to determine:
  - Required starting rent to reach target NOI
  - Required rent growth rate to achieve income targets
  - Exit cap rate needed to reach a $100M sale price
- Demonstrated sensitivity of outcomes to key assumptions

### 6 Sensitivity Analysis (Data Tables)
- Performed one-variable sensitivity analysis on rent growth %
- Built two-variable data tables analyzing:
  - Starting rent
  - Rent growth %
- Visualized results using conditional formatting

### 7 Investment Performance Metrics
- Calculated:
  - Total Investment
  - Total Return
  - Return on Investment (ROI): 857.61%
- Benchmarked performance against a 20.00% required return

### 8 Time Value of Money (TVM) Analysis
- Calculated:
  - Future Value (FV)
  - Present Value (PV)
  - Net Present Value (NPV)
  - XNPV using actual date-based cash flows
- Confirmed the investment exceeds benchmark requirements

### 9 Internal Rate of Return (IRR)
- Computed:
  - IRR
  - XIRR (date-accurate)
- Final annualized return (XIRR): 29.75%
- Validated results by equating benchmark rate to IRR
