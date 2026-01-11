# Excel VBA Portfolio Analyzer

This Excel workbook automates portfolio return and risk analysis using historical stock price data. The tool allows users to import prices and calculate asset returns with one click, while Stage 4 portfolio-level statistics are computed using **built-in Excel functions** rather than VBA. This project demonstrates automation, financial reasoning, and practical Excel & VBA skills.

---

## Project Goal

Build an Excel VBA tool to streamline portfolio analysis by converting historical prices into daily returns and computing portfolio-level statistics such as expected return and volatility.

---

## Project Stages

### Stage 1: Workbook Setup (Foundation)

**Tasks:**
- Create sheets: `Portfolio_Input`, `Price_Data`, `Returns`, `Summary`
- In `Portfolio_Input`, add columns:
  - `Ticker`
  - `Portfolio Weight` (ensure weights sum to 1, manually or via validation)

**Deliverable:** Clean workbook layout with labeled sheets

---

### Stage 2: Import Price Data (Automation)

**Tasks:**
- Write a VBA macro to:
  - Open a file dialog
  - Import a CSV containing historical prices
  - Load the data into `Price_Data` sheet
  - Clear old data before importing new data

**VBA Concepts Used:**
- File dialog
- Worksheet automation
- Dynamic ranges

**Deliverable:** Button labeled **“Import Price Data”** that updates the sheet

---

### Stage 3: Calculate Asset Returns

**Tasks:**
- Write a VBA macro to:
  - Loop through rows and columns of price data
  - Calculate daily returns using:  
    ```
    (Price_t / Price_{t-1}) - 1
    ```
  - Output results to `Returns` sheet

**VBA Concepts Used:**
- Nested For loops
- Cell referencing
- Basic financial math

**Deliverable:** Automated population of the `Returns` sheet

---

### Stage 4: Portfolio Calculations

**Tasks:**
- Calculate using **Excel built-in functions** (not VBA):
  - Average return per asset (`AVERAGE`)
  - Portfolio expected return (weighted average using `SUMPRODUCT`)
  - Portfolio volatility (`STDEV.P`)

**Deliverable:** Updated statistics displayed in the `Summary` sheet

---

### Stage 5: User Interface & Control

**Tasks:**
- Add Form Control buttons:
  - “Import Price Data”
  - “Calculate Returns”
- Assign macros to buttons
- Optional: format outputs for readability

**Deliverable:** One-click workflow from data import to results  

> **Note:** There is no “Update Portfolio” button in this version.

### Stage 7: Documentation

**Tasks:**
- Write this README
- Save the Excel file as macro-enabled (`.xlsm`)
- Upload to GitHub

**Deliverable:** Complete project ready for resume and interviews

---

## File Structure

- **Portfolio_Input:** Enter tickers and portfolio weights  
- **Price_Data:** Imported historical prices  
- **Returns:** Daily asset returns  
- **Summary:** Portfolio-level statistics (average returns, weighted return, volatility)  
- **Macros:** VBA code for importing data and calculating returns  

---

## Technologies & Skills Demonstrated

- Microsoft Excel
- VBA macros for automation  
- Excel built-in functions for calculations (`AVERAGE`, `SUMPRODUCT`, `STDEV.P`)  
- Portfolio analytics and financial reasoning  
- Clean, modular spreadsheet design  

---

**Outcome:**  
A functional, professional Excel portfolio analysis tool demonstrating automation, time series analysis, and practical Excel/VBA skills suitable for co-op or internship portfolios.
