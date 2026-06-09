# Duplicate Invoice Detector

A Python script for Accounts Payable teams that automatically detects potentially duplicated invoices in Oracle ERP exports (including cases where duplicates are hidden behind a different supplier name or a different legal entity).

---

## The Problem It Solves

Standard ERP duplicate checks only flag invoices that are identical across all fields. This script goes further by running two separate fuzzy-matching checks:

| Check | What it ignores | What it catches |
|---|---|---|
| **Check 1 – Ignore Legal Entity** | Legal Entity | Same invoice posted twice on different legal entities |
| **Check 2 – Ignore Supplier** | Supplier Name | Same invoice posted twice under different vendor names |

Invoice numbers are compared using **fuzzy matching** (via `thefuzz`), so minor typos or formatting differences (e.g. `INV-001` vs `INV001`) don't cause false negatives.

---

## Tech Stack

| Purpose | Library |
|---|---|
| Data processing | pandas |
| Fuzzy string matching | thefuzz |
| File discovery | `glob`, `os` |
| Excel | `openpyxl` |

---

## How It Works

```
Oracle ERP exports (*.xlsx) in input folder
              │
     Load & combine all files
              │
   Filter out cancelled invoices (Amount = 0)
              
        ┌─────┴─────┐
   Check 1        Check 2
  ignore LE    ignore Supplier
        └─────┬─────┘
              
  Fuzzy-match invoice numbers within each group
  (threshold: 95% similarity)
              │
  Compare against history file -> keep only NEW findings
              │
  Save new results to .xlsx + update history
```

### Fuzzy Matching Logic

Before comparison, invoice numbers are **normalized**: converted to uppercase and stripped of all special characters (`INV-2024/001` -> `INV2024001`). Then `fuzz.ratio()` computes similarity — pairs scoring ≥ 95% are flagged as potential duplicates. The threshold can be adjusted in the `is_similar()` function.

### History Deduplication

Each run saves its findings to `potential_duplicates_history.xlsx`. On subsequent runs, only **newly discovered** duplicates are written to the output files. Previously flagged pairs are skipped. This makes the script safe to run repeatedly (e.g. daily/weekly) without generating noise.

---

## Input File

Oracle ERP export has following columns:

| Column | Description |
|---|---|
| Invoice Number | Invoice identifier |
| Invoice Date | Date of the invoice |
| Invoice Amount | Amount (rows with 0.0 are treated as cancelled and excluded) |
| Supplier Name | Vendor |
| Legal Entity | Legal entity the invoice was posted to |
| Business Unit | Business unit |

Multiple export files are automatically merged before processing.

---

## Output Files

### `potential_duplicates_no_LE.xlsx`
New duplicate pairs found by **Check 1** (same invoice on different legal entities).

### `potential_duplicates_no_supplier.xlsx`
New duplicate pairs found by **Check 2** (same invoice under different supplier names).

Both files share the same structure:

| Invoice Number 1 | Invoice Number 2 | Supplier 1 | Supplier 2 | Amount | Business Unit | Found Date |
|---|---|---|---|---|---|---|

---

## Possible Improvements

- Send results by email automatically
- Add a confidence score column to the output based on the fuzzy match ratio
- Schedule as a recurring task

---

## License

MIT
