# AGENTS.md

## Purpose

This repository is a Google Apps Script monitoring application backed by Google Sheets and Google Drive.

Primary responsibilities:

- Internal monitoring UI for staff
- External shared read-only views
- Printable monitoring views
- Member master lookup and updates
- Monitoring record storage and retrieval
- Share-link creation, QR generation, and access logging
- AI-generated summaries and care advice

## Architecture

### Runtime model

- Backend: Google Apps Script (`V8`)
- Primary server entrypoints:
  - `doGet(e)` in `コード.js`
  - `doPost(e)` in `コード.js`
- Frontends:
  - `member.html`: internal staff UI
  - `share.html`: external share UI
  - `print.html`: print-friendly shared view

### Data model

The application is spreadsheet-driven. Main sheets:

- `Monitoring`: monitoring records
- `ほのぼのID`: member master / center / staff / QR references
- `ExternalShares`: issued external share links
- `ExternalShareAccessLog`: share access audit log
- `MonitoringReports`: monthly monitoring report storage

Drive is used for:

- record attachments
- generated QR images
- generated PDF reports

OpenAI is used for:

- monitoring summaries
- care advice generation

### Current code organization

`コード.js` is the main application script and currently contains multiple domains:

- routing
- record CRUD
- member lookup and updates
- dashboard aggregation
- share-link management
- Drive helpers
- QR helpers
- OpenAI calls
- report generation

`src/get/billingGet.gs` is separate billing-related retrieval logic.

### Architectural constraints

- Treat the spreadsheet schema as part of the application contract.
- Avoid adding more unrelated responsibilities to the same function or region of `コード.js`.
- Prefer extracting new logic into clearly named helpers or new `.gs` files instead of increasing duplication.
- Do not introduce silent behavior changes to sheet headers, column meanings, or share-token behavior without updating documentation and tests.

## Coding Rules

### General

- Preserve the current Apps Script deployment model unless the task explicitly changes it.
- Prefer small, local changes over broad rewrites.
- Use ASCII by default in new identifiers and comments. Japanese user-facing strings are fine where appropriate.
- Keep functions focused on one responsibility.
- Reuse existing helpers before adding new ones.
- Remove dead code only when you are confident it is unused.

### File and module rules

- Do not add more duplicate helper names. This repository already contains duplicate definitions; avoid making that worse.
- If adding substantial backend logic, prefer a new `.gs` file grouped by domain:
  - `records`
  - `members`
  - `shares`
  - `reports`
  - `drive`
  - `openai`
- If changing HTML files, keep structure, styles, and behavior readable. Do not add another large block of mixed logic if an existing helper can be extended.

### Spreadsheet and data handling

- Never assume column order without checking headers first, unless the codepath is explicitly index-based and documented.
- Normalize member IDs consistently before comparisons.
- Keep read/write behavior backward-compatible with existing sheet data.
- When changing a sheet schema or header name:
  - update all dependent readers/writers
  - document the change in `README.md` or task notes
  - add or update validation/test coverage

### External shares and privacy

- Treat external-share behavior as security-sensitive.
- Do not weaken password, masking, attachment-filtering, expiry, or revocation behavior without explicit justification.
- Default to least-privilege data exposure in shared views.

### OpenAI integration

- Keep prompts deterministic and bounded.
- Avoid unnecessary token growth from repeated raw record dumps.
- Never hardcode secrets in source files.
- Prefer Script Properties for API keys and environment-specific IDs.

### Logging

- Keep logs useful and minimal.
- Remove temporary debug logging before finishing unless it is intentionally retained for operational support.
- Do not log sensitive secrets, passwords, or raw tokens unnecessarily.

## Commit Style

- Use small, reviewable commits.
- Keep each commit focused on one logical change.
- Prefer conventional-style prefixes:
  - `fix:`
  - `feat:`
  - `refactor:`
  - `docs:`
  - `test:`
  - `chore:`
- Write commit subjects in imperative mood.
- Keep the subject concise and specific.

Good examples:

- `fix: normalize share token handling in external view`
- `refactor: extract member lookup helpers`
- `docs: add repository guidance for Apps Script agents`
- `test: cover share masking and expiry rules`

Avoid:

- mixed-purpose commits
- vague messages like `update code`
- formatting-only commits bundled with behavior changes unless necessary

## Testing Rules

### Minimum expectation

Every change should be validated at the narrowest useful level before completion.

### For backend logic changes

Validate at least one of:

- existing Apps Script test/helper function
- deterministic manual execution path
- sheet-driven scenario using representative data

At minimum, verify:

- no syntax errors
- affected entrypoints still return expected shapes
- existing sheet/header assumptions still hold

### For share-flow changes

Manually verify, when relevant:

- valid token works
- revoked token fails
- expired share behavior is correct
- password-protected share requires password
- masked text and attachment filtering still apply

### For member/record changes

Manually verify, when relevant:

- member lookup still resolves normalized IDs
- record save/load still works for existing rows
- dashboard counts still match sheet data
- center/staff updates still write to intended columns

### For UI changes

Check both:

- internal member UI (`member.html`)
- external share UI (`share.html`) when affected

Verify desktop and mobile layout for any substantial visual change.

### Regression awareness

This codebase has duplicated functions and implicit spreadsheet contracts. When modifying shared helpers:

- search for all call sites
- check for later duplicate definitions with the same name
- confirm which definition actually wins at runtime

### If testing is incomplete

- state exactly what was and was not verified
- call out residual risk clearly

