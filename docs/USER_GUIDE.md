# User Guide

Excel Comparison Automation is a web-based application you run locally. It guides you through a six-step wizard: upload two Excel files, select worksheets, configure matching keys, configure comparison columns, run the comparison, and download a detailed Excel report.

This guide is written for non-technical Excel users. No programming knowledge is required.

## Before You Start

- You need two Excel files (`.xlsx` format only).
- Each file should have data organized in a table with a header row.
- The files should share at least one column that uniquely identifies each record (for example, an ID number, employee code, or order number).
- Maximum file size: 50 MB per file.

---

## Step 1: Upload File A and File B

1. Launch the application (see [Installation](INSTALLATION.md) if you have not set it up yet).
2. The app opens on **Step 1: Upload Excel Files**.
3. Under **File A (old/existing data)**, click **Browse files** and select your first Excel file.
4. Under **File B (new/current data)**, click **Browse files** and select your second Excel file.
5. When both files are uploaded, you will see a success message showing both file names.

> **Tip:** Think of File A as the "old" or "original" data and File B as the "new" or "updated" data. The report will show what is missing from File B, what is new in File B, and what changed between them.

6. Click **Next: Select Worksheets**.

---

## Step 2: Select Worksheets

Each Excel file can contain multiple worksheets (tabs). You must choose one worksheet from each file to compare.

1. From the **File A worksheet** dropdown, select the worksheet in File A that contains the data you want to compare.
2. From the **File B worksheet** dropdown, select the worksheet in File B that contains the corresponding data.
3. You will see a success message confirming both worksheets were loaded.

> **Note:** You can compare only one worksheet per file in a single run. If your data is spread across multiple worksheets, run the comparison separately for each worksheet.

4. Click **Next: Matching Keys**.

---

## Step 3: Configure Matching Keys

Matching keys tell the application how to identify the same record across both files. For example, if both files have an `Employee ID` column, you pair `Employee ID` in File A with `Employee ID` in File B.

1. **Suggested key pairs:** The app automatically suggests column pairs that have identical names in both files. Check the boxes next to any suggestions you want to use.
2. **Manual key mapping:** You can add additional pairs using the dropdowns. Select a column from File A and the matching column from File B.
3. If your files use different column names for the same field (for example, `Site ID` in File A and `Site Code` in File B), manually pair them.
4. **Composite keys:** If no single column uniquely identifies a record, add multiple key pairs. For example, pair `Department` ↔ `Department` and `Employee ID` ↔ `Employee ID`. The application treats the combination as the record identifier.
5. Click **+ Add key pair** to add more pairs, or click the **×** button next to a pair to remove it.
6. When at least one valid key pair is configured, you will see a confirmation message.

7. Click **Next: Comparison Columns**.

---

## Step 4: Configure Comparison Columns

Comparison columns are the fields whose values you want to check for changes. Key columns are excluded automatically because they are used only for matching.

1. **Suggested comparison pairs:** Exact-name column matches (excluding key columns) are suggested automatically. Check the boxes next to the columns you want to compare.
2. **Manual comparison mapping:** Select a column from File A and the corresponding column from File B. Choose a **Type** for each pair:
   - **Text** (default): Compares values as text. Leading zeros, spaces, and capitalization are handled according to your settings (whitespace is normalized and comparison is case-insensitive by default).
   - **Numeric**: Compares values as numbers with a small tolerance (for example, `100.000000001` and `100.0` are treated as equal). Use this for amounts, quantities, scores, etc.
   - **Date**: Compares values as dates. Use this for date columns.
3. Add or remove pairs using **+ Add comparison pair** and the **×** button.
4. If you do not select any comparison columns, matched records will be reported as "Matched, not compared" rather than checked for value changes.

5. Click **Next: Run Comparison**.

---

## Step 5: Run Comparison

1. Review your configuration. You will see the file names and a reminder if no comparison columns were selected.
2. Click **Run comparison**.
3. The application processes the files and shows a dashboard with summary metrics:
   - **Records with no changes:** Keys matched and all compared values are equal.
   - **Records with changes:** Keys matched but at least one compared value differs.
   - **Individual values changed:** Total count of changed cells across all changed records.
   - **Records matched, not compared:** Records matched by key but no comparison columns were configured.
   - **Records missing from File B:** Records in File A with no matching key in File B.
   - **New records in File B:** Records in File B with no matching key in File A.
   - **Duplicate identification records:** Records with duplicate key values (reported for review, not matched).
   - **Records with missing identification values:** Records where a key column is empty (reported for review).
4. Below the dashboard, five tabs show detailed results:
   - **Unchanged Records**
   - **Records With Changes**
   - **Missing From File B**
   - **New Records in File B**
   - **Records Needing Attention**
5. Click **Next: Download Report** when you are ready to save the results.

---

## Step 6: Download the Excel Report

1. On **Step 6: Download Report**, click **Download Excel Report**.
2. The file `comparison_report.xlsx` is downloaded to your computer.
3. Open it in Excel to review the six sheets.
4. When you are done, tick the confirmation box and click **Finish & Clear Data** to remove the uploaded files, worksheets, mappings, results, and report data from this session. The report stays available until you confirm.

See [REPORT_GUIDE.md](REPORT_GUIDE.md) for a detailed explanation of each sheet.

---

## Clearing Your Data

You can remove all uploaded and generated data from the current session at any time:

- **Finish & Clear Data** (Step 6) — after downloading the report. Tick the confirmation box first. The report and download button remain available until you confirm.
- **Start Over / Clear Data** (sidebar) — available on every step, to discard an unfinished comparison. Tick the confirmation box first.

Both actions return you to Step 1 with empty uploaders.

Notes:

- Clearing removes the application's references to the data. It is not a secure memory wipe, and a copy of the generated report may remain in the browser or in Streamlit's internal media store until the session ends.
- The application cannot detect when a browser download has finished, so cleanup is only ever triggered by you, never automatically.
- If you simply close the browser tab without clearing, Streamlit discards the session after a short timeout (default 120 seconds).

---

## Understanding Your Results

### What "Matched" Means

Two records are matched when their key column values are identical (after normalization). If comparison columns are configured and all compared values are equal, the record is "unchanged." If any compared value differs, the record is "changed."

### Missing vs. New

- **Missing from File B:** The record exists in File A but no record with the same key exists in File B.
- **New in File B:** The record exists in File B but no record with the same key exists in File A.

### Data Quality Issues (Records Needing Attention)

- **Duplicate key:** The same key value appears more than once in one or both files. These records are not automatically matched because the application cannot determine which occurrence corresponds to which. Review them manually.
- **Missing key value:** A record has an empty value in a key column. These records are excluded from matching and reported separately.

### Excel Row References

Every record-level sheet includes an "Excel Row" column that tells you the actual physical row number in the original source file, including any leading blank rows before the header. Use this to locate the record quickly in your source file.

---

## Tips for Best Results

- **Use consistent column names** when possible. Exact-name matches are suggested automatically and reduce setup time.
- **Clean your data before comparing.** Remove completely blank rows and trailing empty rows from your worksheets.
- **Use Numeric type for amounts.** Text comparison treats `100` and `100.00` as different values.
- **Keep a backup of your source files.** The application never modifies your original Excel files, but it is good practice to keep backups.
- **Run one worksheet at a time.** If your workbook has multiple sheets, repeat the process for each sheet you need to compare.