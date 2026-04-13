# Invoice Automation System

## Overview

This project automates invoice generation from raw CSV input data and produces professionally formatted Excel invoice files for each invoice ID.

It is designed as a portfolio-grade freelance automation project that demonstrates:

- data cleaning
- validation
- Excel invoice generation
- logging
- automated output creation
- reusable project structure

---

## Project Structure

invoice_automation_system/
├── input/
├── output/
├── log/
├── README.md
└── invoice_generator.py

---

## Features

- Reads invoice data from a CSV input file
- Includes CSV and Excel loading logic in the codebase for future extension
- Standardizes column names automatically
- Validates required columns before processing
- Removes completely empty rows
- Cleans and trims key text fields
- Validates invoice dates
- Validates numeric fields such as quantity, unit price, and tax rate
- Rejects invalid records such as:
  - missing required values in key fields
  - invalid dates
  - non-numeric quantity, price, or tax values
  - quantity less than or equal to zero
  - negative unit prices
  - tax rates outside the range of 0 to 1
- Calculates:
  - line subtotal
  - line tax
  - line total
  - due date
- Creates one formatted Excel invoice per invoice ID
- Generates a run summary workbook
- Writes logs to `log/run.log`

## Input File

Place the input file in the `input/` folder as:

`invoice_data.csv`

Required columns:

- invoice_id
- client_name
- client_email
- invoice_date
- item_description
- quantity
- unit_price
- tax_rate

---

## Example Input Data

```csv
invoice_id,client_name,client_email,invoice_date,item_description,quantity,unit_price,tax_rate
INV-1001,Acme Pty Ltd,accounts@acme.com,2026-04-01,Excel Dashboard Development,1,850,0.10
INV-1001,Acme Pty Ltd,accounts@acme.com,2026-04-01,Monthly KPI Report Setup,2,220,0.10
INV-1002,Blue Horizon Co,finance@bluehorizon.com,2026-04-03,Invoice Automation Script,1,680,0.10
INV-1002,Blue Horizon Co,finance@bluehorizon.com,2026-04-03,Data Cleanup Service,3,95,0.10
INV-1003,Northpeak Studio,admin@northpeak.com,2026-04-05,Custom Reporting Workflow,1,1200,0.10

```
---

## Output

After running the script, the `output/` folder will contain:

- one Excel invoice per invoice ID
- one invoice run summary workbook

Example output files:

- `invoice_INV-1001_Acme_Pty_Ltd.xlsx`
- `invoice_INV-1002_Blue_Horizon_Co.xlsx`
- `invoice_INV-1003_Northpeak_Studio.xlsx`
- `invoice_run_summary.xlsx`

## Logging

All execution logs are stored in:

`log/run.log`

The log captures:

- file loading
- cleaning steps
- validation progress
- invoice generation progress
- summary report creation
- errors and exception details

## How to Run

Open a terminal in the project folder and run:

```bash
python invoice_generator.py
```

## What This Project Demonstrates

This project demonstrates skills relevant to freelance Excel and Python automation work:

- transforming raw business data into structured deliverables
- validating and cleaning real-world input data
- generating client-ready Excel invoice files
- creating summary reporting outputs
- organizing a project with clear folder structure
- adding traceability and debugging support through logging

## Possible Future Upgrades

- direct Excel input selection in the main workflow
- PDF invoice export
- automatic email sending
- company logo insertion
- configurable branding
- multi-currency support
- GUI for non-technical users