# Reference Guide

Use this file only for durable, reusable **general rules**. Put narrow, example-driven, or document-specific learnings in `cases.md` instead.

## Bank Statement Analysis Worksheet: Cell Mapping

### 1. Account Information (Header Rows)

These fields are extracted once per borrower/account and written to the active worksheet.

| JSON Field                 | Excel Cell | Worksheet Label                      | Notes                                  |
| :------------------------- | :--------- | :----------------------------------- | :------------------------------------- |
| `account_number`           | C6         | Account Number                       | Text value (last 4 digits preferred)   |
| `account_holder`           | C7         | Account Holder                       | Text value                             |
| `account_holder_address`   | C8         | Account Holder Address               | Text value                             |
| `account_type`             | C9         | Account Type                         | Text value (e.g., Business Checking)   |
| `expiration_date`          | C10        | Expiration Date                      | Text value (MM/DD/YYYY format)         |

### 2. Monthly Data (Row Mapping)

Each calendar month corresponds to a fixed row in the worksheet:

| Month     | Excel Row |
| :-------- | :-------- |
| January   | 14        |
| February  | 15        |
| March     | 16        |
| April     | 17        |
| May       | 18        |
| June      | 19        |
| July      | 20        |
| August    | 21        |
| September | 22        |
| October   | 23        |
| November  | 24        |
| December  | 25        |

### 3. Monthly Data (Column Mapping by Year)

The columns are determined by the `year` and the specific data point being written.

| JSON Field               | Year | Excel Column | Worksheet Label                 | Notes                                    |
| :----------------------- | :--- | :----------- | :------------------------------ | :--------------------------------------- |
| `total_deposits`         | 2024 | B            | Total Deposits (2024)           | Numeric float (e.g., 15000.00)           |
| `total_deposits`         | 2025 | C            | Total Deposits (2025)           | Numeric float                            |
| `total_deposits`         | 2026 | D            | Total Deposits (2026)           | Numeric float                            |
| `total_non_considered`   | 2024 | E            | Total Non-Considered (2024)     | Numeric float (Always 0.00 for MVP)      |
| `total_non_considered`   | 2025 | F            | Total Non-Considered (2025)     | Numeric float (Always 0.00 for MVP)      |
| `total_non_considered`   | 2026 | G            | Total Non-Considered (2026)     | Numeric float (Always 0.00 for MVP)      |
| `nsf_count`              | Any  | H            | NSF Count                       | Integer. Applies to the row regardless of year. |

---

## Business Logic & Edge Cases

### 1. Statement Period Assignment (The "Majority Days" Rule)

Bank statements frequently span non-calendar months (e.g., 04/12/2025 - 05/11/2025). 
- **Rule:** Assign the statement to the calendar month that contains the highest number of days within that statement period.
- **Example Calculation:**
  - Period: April 12 to May 11
  - April days: April 12 through April 30 = 19 days
  - May days: May 1 through May 11 = 11 days
  - Result: 19 > 11, so this statement belongs to **April**.
- **Tie-breaker:** If the days are exactly equal (e.g., 15 days in each month), flag the statement for manual review. Do not guess.

### 2. Total Deposits Definition

- Extract the **Gross Total Additions / Deposits** as explicitly stated on the bank statement summary page.
- Do not manually sum individual line items unless the summary page is missing or illegible.
- Do not deduct any withdrawals, fees, or negative balances.

### 3. Non-Considered Deposits (MVP Constraint)

- In the current minimum viable product (MVP) phase, the exclusion of specific deposits (e.g., transfers, loan proceeds, refunds) is **not implemented**.
- The `total_non_considered` field must always be `0.00`.
- The `non_considered_details` array must always be empty `[]`.
- Do not write any values to columns E, F, or G other than `0`.

### 4. NSF (Non-Sufficient Funds) Count

- Count any fee explicitly labeled as:
  - "NSF Fee"
  - "Non-Sufficient Funds"
  - "Overdraft Fee"
  - "Returned Item Fee"
  - "Returned Check Fee"
- Do not count standard monthly maintenance fees, wire fees, or ATM fees.
- If no such fees exist, the count is `0`.

### 5. Preserving Formulas

- **Row 26** contains sum formulas for totals. **DO NOT** overwrite or write data to row 26.
- The Python script must load the workbook without the `data_only=True` flag to ensure formulas remain intact when saving.