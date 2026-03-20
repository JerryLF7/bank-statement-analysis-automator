# Bank Statement Chat Summary Template

Use this template after the Excel workbook is ready and sent to the user.

Keep the summary in the chat body. Do not hide it in Excel comments, JSON-only output, or internal notes.

## Purpose

The summary gives the underwriter a quick review view of:
- which deposits were excluded
- why they were excluded
- which items still need human review

## Output Template

Use this exact section structure:

```text
Done — I filled the bank statement worksheet and attached the Excel file.

Excluded deposits
- <date or month> | <amount> | <short description> | Reason: <reason>
- <date or month> | <amount> | <short description> | Reason: <reason>
- None

Needs review
- <date or month> | <short description> | Review: <what is uncertain and what the user should check>
- <date or month> | <short description> | Review: <what is uncertain and what the user should check>
- None
```

## Formatting Rules

- Use plain, readable bullets.
- Keep one item per bullet.
- Do not use markdown tables.
- Keep each bullet concise but specific.
- If there are no excluded deposits, write `- None` under `Excluded deposits`.
- If there are no review items, write `- None` under `Needs review`.
- Do not omit a section just because it is empty.

## Excluded Deposit Bullet Format

Preferred format:

```text
- 2025-02-14 | $2,500.00 | ONLINE TRANSFER FROM SAVINGS | Reason: Internal transfer between borrower accounts
```

Minimum required fields:
- date or month context
- amount
- short description
- exclusion reason

## Needs Review Bullet Format

Preferred format:

```text
- 2025-03 | MOBILE DEPOSIT | Review: Source of funds is unclear from statement text; please confirm whether this is operating revenue or transfer
```

Minimum required fields:
- date or month context
- short description
- what is uncertain
- what the user should verify

## Style Guidance

- Be factual, not defensive.
- Do not dump raw JSON.
- Do not include chain-of-thought or internal reasoning.
- Prefer business-readable wording over technical jargon.
- If many excluded deposits share the same reason, keep bullets separate anyway for auditability.

## Example

```text
Done — I filled the bank statement worksheet and attached the Excel file.

Excluded deposits
- 2025-01-09 | $15,000.00 | SBA TREAS LOAN DISBURSEMENT | Reason: Loan proceeds are liabilities, not qualifying operating income
- 2025-01-22 | $3,200.00 | ONLINE TRANSFER FROM SAVINGS | Reason: Internal transfer between borrower accounts

Needs review
- 2025-02 | MOBILE DEPOSIT | Review: The source is not clear from the statement description; please confirm whether it is business income or a transfer
- 2025-03 | Statement period assignment | Review: The statement spans two months with a near-even split; please verify the month assignment is correct
```
