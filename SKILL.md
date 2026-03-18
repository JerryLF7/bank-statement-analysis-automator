---
name: bank-statement-analysis-automator
description: Extracts monthly total deposits and NSF counts from bank statements and writes them into a 12-month analysis Excel template.
---

# Bank Statement Analysis Automator

Use this skill when the user wants to extract monthly deposit totals and NSF (Non-Sufficient Funds) counts from bank statement PDFs/images and enter them into a "12 Months Bank Statement Analysis" Excel worksheet that already contains formulas.

When this skill is active, you should:

1. Read the bank statements from the provided PDF or image.
2. Extract the required account-level information (account number, holder, etc.).
3. Determine the correct calendar month for each statement based on the majority day count rule.
4. Extract the total deposits and NSF count for each statement period.
5. Build a clean JSON object for the worksheet input.
6. Ask for confirmation when any extracted value is unclear or confidence is low.
7. Write the confirmed values into the Excel worksheet without overwriting formulas.
8. Return the filled Excel file to the user.

Read `extraction_prompt.md` for the exact extraction prompt and output schema.
Read `reference.md` for worksheet mapping rules and business logic.
Use `template.xlsx` as the Excel worksheet template.
Use `scripts/write_excel.py` to write confirmed JSON values into the worksheet.

## Examples
- User uploads a 12-month bank statement PDF and a blank analysis worksheet and asks to fill the worksheet automatically.
- User provides a scanned bank statement package and wants the monthly deposit totals transferred into the Excel template.
- User uploads cross-month statements (e.g., mid-month to mid-month) and needs them assigned to the correct calendar months in the worksheet.

## Guidelines
- Extract values directly using the model's document understanding capability; do not rely on Python OCR logic in `write_excel.py`.
- Use structured JSON as the handoff format between extraction and Excel writing.
- Extract at least these fields: `account_number`, `account_holder`, `account_holder_address`, `account_type`, `expiration_date`, and `monthly_data` (year, month, total_deposits, nsf_count).
- **Important MVP Rule:** Do not attempt to calculate or extract excluded (non-considered) deposits in this version. Always set `total_non_considered` to `0.00` and `non_considered_details` to `[]`.
- **Cross-Month Rule:** If a statement spans two months, assign it to the month with the most days in that period.
- If a field is missing, unclear, or not explicitly shown, set it to `null` and flag it for review.
- Keep numeric fields numeric in JSON and Excel output.
- Preserve all formulas, formatting, and labels in the Excel file.