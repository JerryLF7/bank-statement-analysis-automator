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

### 3. Non-Considered Deposits (Ineligible Deposits)

Only ongoing business revenue should be considered as qualifying income. You must identify and exclude non-revenue deposits. The sum of these goes into `total_non_considered`, and the itemized list goes into `non_considered_details`.

**Rules for Exclusion:**
1. **Internal/Inter-account Transfers:** Funds moving between the borrower's own accounts. 
   - *Keywords:* "Transfer from", "Online Banking Transfer", "Zelle from [Borrower Name]".
2. **Loan Proceeds & Cash Advances:** Borrowed funds are liabilities, not income.
   - *Keywords:* "SBA Treas", "Loan Disbursement", "Amex Advance", "Kabbage", "Credit Card Advance".
3. **Refunds, Returns & Reversals:** Money sent back from vendors or reversed payments.
   - *Keywords:* "Refund", "Return Item", "Reversal", "Chargeback".
4. **W-2 Wages & Payroll:** Personal salary income from other employment must be excluded (it is calculated separately).
   - *Keywords:* "Payroll", "Paystub", "Salary", "Direct Deposit [Employer Name]".
5. **One-time / Unusual Deposits:** Non-recurring personal or unusual lump sums.
   - *Keywords:* "IRS TREAS 310" (Tax refunds), Escrow payouts, Insurance claims.

**Handling Details:**
- Every excluded transaction must be recorded in the `non_considered_details` JSON array with its `date`, `amount`, `description`, and the `reason` (referencing one of the rules above).
- Also report excluded transactions in the final chat summary so the user can see what was excluded and why without opening intermediate artifacts.
- Keep worksheet filling behavior unchanged; the explanation layer belongs in the chat response, not in Excel comments.

### 3A. Required Chat Summary

After generating the Excel output, include a concise, readable summary in the chat response with exactly these sections:

- **Excluded deposits**
- **Needs review**

Rules:
- Use bullets for readability.
- For each excluded deposit, include date or month context, amount, short description, and reason.
- For each review item, describe what is uncertain and what the user should verify.
- If a section has no items, write `None`.
- Do not hide important uncertainty only in internal reasoning.

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