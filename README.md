# Gonzaga Student Investment Fund Portfolio Tracker

## Project idea

A basic spreadsheet can track numbers, but it is easy to make mistakes when adding transactions or updating the portfolio manually.  
This project uses Excel formulas, tables, dashboard sections, and VBA code to make the process more organized.

The workbook can track buy and sell transactions, calculate holdings, show portfolio performance, check allocation, give warnings, and save snapshots of the portfolio over time.

## Main functionality

The project includes:

- transaction input form
- buy and sell transaction tracking
- input validation before adding a transaction
- automatic calculations for transaction totals
- net cash flow calculation
- average cost logic
- realized gain or loss for sell transactions
- holdings calculation by ticker
- current market value calculation
- cost basis calculation
- unrealized gain or loss calculation
- portfolio weight calculation
- dashboard refresh
- portfolio summary update
- allocation and rebalancing check
- portfolio warning messages
- history snapshot saving
- VBA activity logging

## Workbook sheets

### Dashboard

The Dashboard is the main page of the workbook.  
It shows the main portfolio information in one place.

It includes:

- portfolio overview
- key portfolio numbers
- portfolio health section
- allocation check
- rebalancing table
- warning messages
- Refresh Portfolio button
- Save Snapshot button

The Dashboard is meant to be the main page that a user looks at after entering transactions and updating prices.

### Transactions

The Transactions sheet stores all buy and sell transactions.

It has two main parts:

- the transaction table
- the input form

The input form is used to enter a new transaction. After the user clicks the Add Transaction button, VBA checks the input and adds the transaction to the table.

The Transactions sheet also includes calculated columns for:

- total amount
- net cash flow
- average cost before sale
- realized gain or loss

### Holdings

The Holdings sheet calculates the current position for each asset.

It shows:

- ticker
- asset type
- total bought
- total sold
- net quantity
- average cost
- current price
- market value
- cost basis
- unrealized gain or loss
- portfolio weight

This sheet connects the transaction history with the current prices.

### Prices

The Prices sheet stores the current prices for each asset.  
These prices are used to calculate current market value and unrealized gain or loss.

### Summary

The Summary sheet shows portfolio-level results.

It includes totals such as:

- current portfolio value
- total cost basis
- total unrealized gain or loss
- number of assets
- largest holding
- portfolio performance information

Some values on this sheet are updated by VBA when the portfolio is refreshed.

### History

The History sheet saves portfolio snapshots.

When the user clicks Save Snapshot, the macro saves the current portfolio status.  
This makes it possible to track portfolio changes over time.

### VBA_Log

The VBA_Log sheet records macro activity.

For example, it can record when:

- the portfolio was refreshed
- a transaction was added
- a snapshot was saved

This helps show that the VBA automation is working.

### Lists

The Lists sheet is a hidden technical sheet.  
It is used for dropdown lists and validation options.

## VBA code

The workbook uses several VBA modules.  
I also included the VBA code in the `vba-code` folder so it can be viewed directly on GitHub.

### modDashboard.bas

This module controls the main dashboard refresh process.

It includes code that:

- refreshes the portfolio
- recalculates the workbook
- updates the dashboard timestamp
- updates portfolio warnings
- connects the refresh button to the refresh macro

This is one of the main automation parts of the project because it updates the workbook after changes are made.

### modTransactions.bas

This module controls the transaction form logic.

It includes code that:

- clears the input form
- validates required fields
- checks that quantity and price are valid
- checks that broker fee is not negative
- checks available quantity before a sell transaction
- adds a new transaction to the transaction table
- writes formulas into the new row
- refreshes the portfolio after a transaction is added

This module is important because it helps prevent bad transaction data from being added.

### Sheet2_Transactions.bas

This is the worksheet code for the Transactions sheet.

It responds to changes on the Transactions sheet.  
For example, it helps autofill some form information when a ticker is selected.

This makes the input form easier to use.

### modSummary.bas

This module updates the Summary sheet.

It calculates and updates portfolio-level information such as:

- portfolio value
- cost basis
- unrealized gain or loss
- number of assets
- largest holding
- top performer
- worst performer

This module helps connect the detailed worksheet calculations to the summary report.

### modHistory.bas

This module saves portfolio snapshots.

It adds a new row to the History sheet with current portfolio information.  
This allows the user to keep a record of portfolio changes over time.

### modLogging.bas

This module writes actions to the VBA_Log sheet.

It is used to record macro activity, such as refreshing the portfolio, adding transactions, and saving snapshots.

## How to use the workbook

1. Open the `.xlsm` Excel file.
2. Enable macros.
3. Go to the Transactions sheet.
4. Enter the transaction date, ticker, action, quantity, price, broker fee, account, and notes.
5. Click Add Transaction.
6. Update asset prices on the Prices sheet if needed.
7. Go to the Dashboard.
8. Click Refresh Portfolio.
9. Check the portfolio value, holdings, warnings, and rebalancing information.
10. Click Save Snapshot if you want to save the current portfolio status to the History sheet.

## Files in this repository

- `Gonzaga_Investment_Fund_Portfolio_Tracker.xlsm`  
  Main Excel macro-enabled workbook.

- `vba-code/modDashboard.bas`  
  Dashboard refresh and warning logic.

- `vba-code/modTransactions.bas`  
  Transaction form, validation, and add transaction logic.

- `vba-code/Sheet2_Transactions.bas`  
  Worksheet event code for the Transactions sheet.

- `vba-code/modSummary.bas`  
  Summary sheet update logic.

- `vba-code/modHistory.bas`  
  History snapshot logic.

- `vba-code/modLogging.bas`  
  VBA log writing logic.
