# Test results

## First public release — 2026-10-11

The release preserves the tested checker logic. The published worksheet source was extracted from the executed document module, including Excel's identifier casing and equivalent numeric-literal notation. Packaging creates a separate one-sheet legacy XLS donor, removes document metadata and inherited names, and sanitises the Mac Forms library username. The original development files and previous backups remain local.

The exact final release passed **35 read-only inspection checks**. Its sole executable worksheet module matches src/vba/BudgetCheckWorksheet.vba exactly. The donor contains no prior findings, formulas, external workbook links, names, cell comments or personal document metadata. Worksheet bytes are unchanged by library-path sanitisation.

Opened the sanitised release in desktop Excel for Mac with macros enabled. Copied the entire checker worksheet into a new workbook containing a fictional 21-section detail budget and both summary sheets. Closed the release donor before running checks. All checker assertions completed:

- Correct baseline with account codes on cost and fringe rows, direct I totals and no K/L helpers: **0 findings**.
- Positive base line titled Järjestäjä with zero fringes: role title and fee/invoice/exemption guidance appear.
- Labelled 135% fringe: excessive-fringe error detected.
- I subtotal omitting three zero-value H lines: exactly **1** subtotal-gap review, with the formula and omitted rows; no duplicate coverage review on directly linked summaries.
- A different incomplete summary link: separate summary coverage review retained.
- Positive costs on the omitted group: H/I and H/summary reviews detected, together with a differing fringe-rate review.
- Missing rate on a standard template-style IF calculation: active-zero-line review detected.
- Rate stored as text: numeric-text review detected.
- Copied calculation using another group's inputs: unexpected-reference review detected.
- Hardcoded H total: error detected.
- Missing section marker with short account codes still on cost rows: missing-marker review detected; section comparison paused.
- Changed summary grand total: error detected.
- Physically inserted an account-coded cost row and extended its I subtotal: **0 findings**; deleting it and restoring the subtotal also produced **0 findings**.
- Baseline and restored detail inputs/formulas were compared before and after checking and were unchanged.

Evidence: tests/release_check.applescript. The test constructs its budget from scratch and uses no third-party template or private budget. Its saved disposable fixture was read back with **0 findings** and closed. The final script's cleanup refers to the saved filename because Save As changes Excel's workbook name; that cleanup was verified separately and the complete script compiled successfully.

## Earlier development verification

Native Excel checks in the local development project covered the supplied Finnish template, K/L helper tracing, direct I layouts with K/L physically deleted, inserted H subtotals, unsupported calculations, missing and duplicate section markers, role context and a merged report status area. The full injected-error fixture for the final checker produced **14 findings**, retaining all six additional checks and original formula/fringe errors, with one redundant summary coverage finding removed. Templates, fixtures and historical backup workbooks are excluded from this repository and download.

## English-template verification — 2026-10-11

Tested the unchanged v1.0.0 release in desktop Excel for Mac using disposable copies of the supplied English template. Copied the whole checker worksheet into the template and closed the donor before running it. The blank template and a fictional budget populated across all 21 sections each produced **0 findings**, including the laboratory and equipment sections.

Injected cases correctly detected hardcoded H and I totals, holiday compensation and social charges above 100%, a missing rate, a numeric rate stored as text, a copied formula referencing another line, differing fringe rates, and a changed summary grand total. Zero fringes on a paid line produced review guidance with the English role title and the salary/fee distinction. Short account codes in A on cost lines did not cause false section warnings.

Both a shortened K helper range and a shortened direct I range reported omitted zero-value salary lines. A separate custom fixture with K/L physically deleted, section totals calculated in I and summary links adjusted to those totals produced **0 findings**. Inserting an account-coded travel row and subsequently removing it each produced **0 findings**. Incorrect ranges, a paid travel line omitted from its subtotal, a wrong laboratory summary reference and a missing section number were detected. Restoring the custom fixture returned **0 findings**.

Detail inputs and formulas were compared before and after checker runs and remained unchanged. The supplied English template and published release retained their original file checksums. No checker code update was needed; this verification updates the documentation only. The third-party template and disposable workbooks are not included in the repository or release.

## Limits

Tested on desktop Excel for Mac with the supplied Finnish and English templates. Windows compatibility is unverified. Custom formulas, named/dynamic references and ambiguous layout changes need manual review. The checker is a review aid and does not repair budgets. Missing-input and ordinary copied-reference checks target the supplied IF formula pattern; formula coverage is not a complete arithmetic proof. No private production budget was opened for these tests.
