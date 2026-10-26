# CBO's NIPA Data
Each year, the Congressional Budget Office publishes projections of what federal revenues and outlays would be in the current fiscal year and over the next 10 years if laws governing taxes and spending generally remained unchanged. Those baseline budget projections reflect the standard structure for budgetary accounting and are consistent with OMB’s presentation of the budget. CBO translates those baseline projections of revenues and outlays into the NIPA framework, with its categories of current receipts and expenditures, to prepare the agency’s forecast of the economy.

This repository transforms CBO NIPA spreadsheets into CSV outputs to make this data more machine-readable.

Source: `NIPA-Tables-2026.xlsx`, read directly from the CBO publication page: https://www.cbo.gov/publication/62876

## How to run

Requires Python 3.10 or later.

1. Install dependencies:

   ```powershell
   python -m pip install pandas "openpyxl>=3.1"
   ```

2. Run the converter from the repository folder:

   ```powershell
   python convert_nipa_to_csv.py [source] [output.csv]
   ```

   - `source` (optional): a URL or a local `.xlsx` path. If omitted, the workbook is downloaded from the CBO publication page.
   - `output.csv` (optional): output file path. If omitted, the CSV files are written to the repository folder.

   Examples:

   ```powershell
   # Download from the CBO publication page
   python convert_nipa_to_csv.py

   # Use a local copy of the workbook
   python convert_nipa_to_csv.py .\NIPA-Tables-2026.xlsx
   ```

## Files

- `NIPA-Tables-2026_flat.csv`: one row per series and year. Key: `series_id` + `year`. `section` identifies the top-level block (e.g. receipts, outlays, deficit).
- `NIPA-Tables-2026_footnotes.csv`: row-specific footnotes. `applies_to` lists the `series_id`s each note applies to (directly or via a section header). Join with `_flat.csv` through `footnote_ids`.

## General notes

- For more information about the national income and product accounts (NIPAs) and an explanation of how the federal budget framework differs from that of the NIPAs, see Congressional Budget Office, A Primer on CBO's Projections of Federal Receipts and Expenditures in the National Income and Product Accounts (October 2026), www.cbo.gov/publication/62876.

## Table 1. CBO's Projections of Federal Receipts and Expenditures in the National Income and Product Accounts

- **Data source:** Congressional Budget Office; Bureau of Economic Analysis. All amounts are for fiscal years. NIPAs = national income and product accounts.

## Table 2. Differences Between Revenues and Outlays in CBO's Baseline Budget Projections and Projected Receipts and Expenditures in the Federal Sector of the National Income and Product Accounts

- **Data source:** Congressional Budget Office. All amounts are for fiscal years. Coverage differences arise when transactions are reported in either the budget or the NIPAs but not both. Netting differences occur when transactions are recorded as outlays in the budget but as receipts in the NIPAs, or as revenues in the budget but as expenditures in the NIPAs. Timing differences arise when certain transactions that are recorded in the budget when receipts are collected or payments are made are recorded in the NIPAs at a different time to better align them with the timing of related production or accrued income. HI = Medicare's Hospital Insurance; NIPAs = national income and product accounts; OASDI = Social Security's Old-Age, Survivors, and Disability Insurance.

## Acknowledgments
CBO's NIPA translation code was developed by **Aaron Feinstein**.

CBO staff used generative artificial intelligence tools to help write or review the code and documentation in this repository. Staff reviewed all materials here and CBO is responsible for the final work. Results and analytic appropriateness were verified only for the analyses identified in this repository.

## Contact
Questions may be directed to CBO's Office of Communications at communications@cbo.gov. CBO will respond to inquiries as its workload permits.