# Gonzaga Student Investment Fund Portfolio Tracker

This project is an Excel VBA portfolio tracker created for a student investment fund project.  
It is designed to help an investment club track transactions, holdings, portfolio value, allocation, warnings, rebalancing needs, and historical portfolio snapshots.

## Project Overview

The workbook combines Excel formulas, structured tables, dashboard visuals, and VBA automation.  
The main goal of the project is to make portfolio tracking easier, more organized, and more automated than a basic spreadsheet.

Users can add Buy and Sell transactions, update market prices, refresh the portfolio, review portfolio performance, check allocation problems, and save portfolio snapshots over time.

## Main Features

- Automated transaction input form
- Buy and Sell transaction tracking
- Input validation before adding transactions
- Holdings calculation by ticker
- Current market value tracking
- Average cost and cost basis logic
- Unrealized gain and loss calculation
- Realized gain and loss logic for sell transactions
- Portfolio summary report
- Dashboard with key portfolio metrics
- Allocation and rebalancing check
- Portfolio warning system
- Historical snapshot tracking
- VBA log for macro activity
- Clean workbook structure with technical support sheets hidden

## Workbook Sheets

### Dashboard

The Dashboard is the main page of the project.  
It shows the portfolio overview, key financial metrics, portfolio warnings, allocation checks, and rebalancing information.

Main parts include:

- Portfolio overview
- Portfolio health
- Allocation check
- Rebalancing table
- Portfolio warnings
- Refresh button
- Save snapshot button

### Transactions

The Transactions sheet stores all portfolio transactions.  
It includes a structured transaction table and an input form for adding new transactions.

The transaction form allows the user to enter:

- Date
- Ticker
- Asset Type
- Action
- Quantity
- Price per Unit
- Broker Fee
- Account
- Notes

The sheet also calculates transaction totals, net cash flow, average cost before sale, and realized gain or loss.

### Holdings

The Holdings sheet calculates the current portfolio positions.  
It summarizes each asset and shows values such as:

- Total bought
- Total sold
- Net quantity
- Average cost
- Current price
- Market value
- Cost basis
- Unrealized gain or loss
- Portfolio weight

### Prices

The Prices sheet stores current market prices for each asset.  
These prices are used to calculate current market value and unrealized performance.

### Summary

The Summary sheet provides portfolio-level financial metrics.  
It summarizes the overall value, cost basis, gains and losses, number of assets, largest holding, and other key statistics.

### History

The History sheet saves portfolio snapshots over time.  
This allows the user to track how the portfolio changes after updates.

### VBA_Log

The VBA_Log sheet records macro activity.  
It helps show when important actions were completed, such as refreshing the portfolio, saving snapshots, or adding transactions.

### Lists

The Lists sheet is a technical support sheet used for dropdowns and validation lists.  
It is hidden because regular users do not need to edit it.

## VBA Automation

The workbook uses VBA macros to automate important tasks.

Main VBA functions include:

- Adding a new transaction
- Clearing the transaction input form
- Validating user input
- Checking available quantity before selling
- Refreshing the portfolio
- Updating the summary report
- Building portfolio warnings
- Saving historical snapshots
- Writing activity to the VBA log

## How to Use the Workbook

1. Open the Excel macro-enabled workbook.
2. Enable macros.
3. Go to the Transactions sheet.
4. Enter a new transaction in the input form.
5. Click the Add Transaction button.
6. Update prices on the Prices sheet if needed.
7. Go to the Dashboard.
8. Click the Refresh Portfolio button.
9. Review the updated portfolio metrics, warnings, and rebalancing results.
10. Click Save Snapshot if you want to save the current portfolio status to the History sheet.

## Financial Logic

The project uses financial logic to calculate portfolio performance.  
It tracks both buying and selling activity and separates current holdings from historical transactions.

The workbook calculates:

- Market value
- Cost basis
- Net quantity
- Average cost
- Unrealized gain or loss
- Realized gain or loss
- Portfolio weight
- Allocation gap
- Rebalancing recommendation

## Project Purpose

The purpose of this project is to demonstrate how Excel and VBA can be used to build a practical automated portfolio tracker.  
The project shows skills in spreadsheet design, financial analysis, data organization, automation, dashboard creation, and VBA programming.

## Tools Used

- Microsoft Excel
- Excel formulas
- Excel structured tables
- VBA macros
- Dashboard visuals
- GitHub for project storage and presentation

## File Included

- `Gonzaga_Investment_Fund_Portfolio_Tracker.xlsm`

This is the main macro-enabled Excel workbook for the project.

## Notes

This project was created for educational purposes.  
The data inside the workbook is sample portfolio data used to demonstrate how the tracker works.
