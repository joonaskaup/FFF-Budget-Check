# FFF Budget Check

Check formulas, subtotals and fringe inputs in compatible FFF / SES Excel budgets. The checker produces a linked review report and leaves budget inputs and formulas unchanged.

**[Download FFF Budget Check (.xls)](https://github.com/joonaskaup/FFF-Budget-Check/releases/latest/download/FFF_Budget_Check.xls)**

[Complete download with licence and source (.zip)](https://github.com/joonaskaup/FFF-Budget-Check/releases/download/v1.0.0/FFF_Budget_Check_v1.0.0.zip) · [All releases](https://github.com/joonaskaup/FFF-Budget-Check/releases) · [Report a problem](https://github.com/joonaskaup/FFF-Budget-Check/issues)

Requires desktop **Microsoft Excel with VBA macros**. Tested on Excel for Mac with the supplied Finnish and English templates. Windows compatibility has not been verified. Excel online, Google Sheets and LibreOffice are not supported by this release.

The download contains one independent **Budget Check** worksheet. Its VBA lives in the worksheet, so copying the whole tab also copies the checker. Replace older checker tabs with this release.

## Start here

1. Open the checker and your budget in desktop Excel. Enable this checker's macros. If updating, save a backup of your budget and rename its old checker tab before copying the new one.
2. Right-click the **Budget Check** tab, choose **Move or Copy**, select your budget workbook, and tick **Create a copy**. Copy the whole worksheet, not its cells. Unprotect the workbook structure if Excel prevents adding a tab.
3. On the copied tab, leave B5 blank for automatic selection of **Erittely** or **DetailedBudget**, or enter your budget sheet's name.
4. Double-click the coloured row 3 to run. Click a cell link in column B of the report to visit the source cell.
5. Save the budget as `.xls` or `.xlsm` to preserve VBA. Run again after changing the budget.

The `.xls` release supports copying into the supplied legacy `.xls` template as well as modern workbooks. Excel requires a legacy-format donor to be reopened before copying into a legacy workbook; the distributed release is already saved in that format.

Checks inspect the selected sheet from row 6 through its used range. H5 and I5 must contain the template's total headers.

- Any populated H/I cell without a formula is an error, including literal zero and text.
- Simple constant formulas such as `=500` are errors too.
- Excel errors in H/I are errors; dependent error cells are also reported.
- Negative calculated totals are review warnings.
- Blank H cells with inputs or fringe labels are review warnings when I does not already calculate a total. Blank I cells with a total label are review warnings. Other intentionally blank cells are ignored.
- Fringe rows are identified by an exact explicit label in B or C, such as Finnish `lomakorvaus` / `palkansivukulut` or English `holiday compensation` / `social charges`. Other recognized English labels include `holiday pay`, `social costs` and `fringe`. Surrounding whitespace and a trailing colon are ignored. A word such as `social` inside an equipment description or person’s name is not a fringe label. The percentage input is column D in the inspected templates.
- Negative fringes, fringes above 100%, text fringe inputs and fringe formula errors are errors. B6 sets an adjustable high-fringe review threshold (default **60%**). This is a review preference, not a statement about permitted employment rates.

Additional checks:

- **Account numbers in A:** cost, fringe and subtotal rows may carry account codes, including codes such as `03`. A code counts as a section number only on a recognized header: a title in B, with C–I empty and no formulas, and no fringe/total label. Salary-to-fringe matching also stops at recognized headers, rather than every populated A cell.
- **Section numbering:** expect each FFF category number 01–21 exactly once on a recognized header in column A, wherever its header currently falls. Numeric values and numeric text such as `03` are accepted; account codes on cost/fringe/subtotal rows do not count. Missing numbers are listed; duplicate numbers identify the repeated header cell and its first occurrence. Incomplete or duplicate numbering pauses section reconciliation, subtotal coverage and within-section fringe comparisons, so costs are not silently grouped into the preceding category. Individual formula, input and fringe checks, and the direct grand-total comparison, still run. Restore missing numbers on retained headers; an intentionally deleted entire category requests manual reconciliation.
- **Missing inputs:** a standard line with at least one nonzero D/F/G input totals zero because D quantity or G rate is missing/zero. Template descriptions alone do not activate a line, and F=0 is valid.
- **Copied references:** standard H line formulas referencing another row's D/F/G inputs, and fringe formulas referencing another salary group, request review.
- **Subtotal gaps:** trace recognized H cost lines through H subtotals and I/K/L helpers into each numbered section's closing total. H subtotals are not added a second time as cost lines. Explicit references to the same worksheet are supported. The first labelled closing row after the cost lines is used, including totals without “yhteensä”; a larger production aggregate below it is not substituted. Untraced lines are reported even when they currently total zero. The finding shows the I subtotal formula, omitted H rows and their current sum. This can catch an incomplete I range before any omitted line has a cost; the numerical H comparison only differs once omitted costs affect the total. Summary links directly to that flagged I subtotal do not repeat the same coverage warning; independent incomplete summary links are still reported. If the section contains H formulas the checker cannot interpret, it requests review instead of asserting a subtotal gap.
- **Reconciliation:** compare the sum of recognized section H cost lines with section I totals and the corresponding H totals on Yhteenveto and Summary. A difference above 0.01 requests review and lists the exact H row ranges included; the user must verify the row scope before changing a total. If any H line has an unsupported calculation pattern, the checker identifies those rows and skips the section's H/I and H/summary comparisons. Summary grand totals are also compared with the detail grand total; those differences remain errors.
- **Fringe consistency:** compare holiday and social rates separately within a section's paid groups. Differing rates and blank/zero fringe inputs on positive base lines request review. Fringe findings show the base line's B-column role title, for example `Järjestäjä - lomakorvaus`; rate comparisons also name the comparison role. **Paid line without fringe** asks whether the payment is salary requiring fringes or a fee/invoice/exempt payment with zero fringes. Titles provide context, not automatic payment classification; fees, exemptions and different agreements may be valid.
- **Numbers stored as text:** numeric-looking D/F/G inputs and fringe percentages stored as strings request review.

The checker writes only its own report and recalculates the selected budget sheet. It does not repair formulas, change budget values, or change budget formatting. Findings are a dated snapshot. Red rows indicate errors; yellow rows request review.

## Limits

Adding or deleting cost rows is supported when section headers and calculation structure remain recognizable. No original section row numbers are stored. Deleting a whole numbered category requests manual reconciliation because the checker cannot distinguish an intentionally removed category from a retained header with a missing number.

Header recognition follows the supplied layout. A completely stripped cost row containing only an account code 01–21 in A and a title in B has the same structure as a section header; the checker cannot infer its intended meaning. Headers with extra description/input/total cells require manual reconciliation. Ordinary cost rows with formulas, descriptions or inputs are excluded from header recognition.

This checker follows the supplied template's A section codes (1–21), D/F/G inputs and H/I totals. K/L helpers are optional: local references through them are traced when present. Custom layouts may need adaptation. It traces ordinary local A1 references and ranges; named ranges, INDIRECT, dynamic references, complex cross-sheet formulas and unusual calculation patterns need manual review. Recognized cost lines use their own D quantity and G rate inputs, or an explicit fringe label with percentage and salary references. Other H calculation patterns cannot establish an independent section sum. Coverage follows referenced cells, not every formula's arithmetic meaning, so this is not a complete formula audit. Even a recognized H calculation does not prove that it belongs in the selected I subtotal; this is why H-based numerical comparisons request review.

Missing-input and copied-reference checks target template-style IF calculations and labelled fringe formulas. Other H calculations may be reconciled without inferring missing-input or copied-reference errors. Missing-input detection requires an entered nonzero input; a name with all numeric inputs empty is not sufficient because the template contains prefilled descriptions. Fringe checks recognize existing labelled fringe rows; a deleted entire fringe block is not automatically inferred. Peer rates are compared within sections and warnings may reflect legitimate agreements. Formula errors in section summaries are reported; other workbook tabs are outside scope.

External-link auditing and calculation-mode warnings are intentionally omitted.

Tested in desktop Excel for Mac on the supplied Finnish and English templates. The English template passed with fictional costs across all 21 sections, including laboratory and equipment. Tests also covered a custom English layout with K/L deleted and totals calculated in I, inserted and removed cost rows, and account codes in A. Deliberate formula, input, fringe, subtotal-range and summary errors were detected, including omitted zero-value lines. These checks required no changes to the v1.0.0 checker. Windows compatibility has not been verified.

## Source and development

The complete worksheet VBA is in [src/vba/BudgetCheckWorksheet.vba](src/vba/BudgetCheckWorksheet.vba). It is a worksheet document module, not an ordinary imported VBA module. Editing source files does not update a workbook already copied into a budget.

See [the development workflow](docs/DEVELOPMENT.md), [test results](docs/TEST_REPORT.md), and [release inspection](docs/RELEASE_AUDIT.md). Git tracks readable source and documentation; the workbook is supplied through GitHub Releases. Templates, personal budgets, backups and generated test workbooks are excluded.

If an older checker reports `We can't do that to a merged cell`, replace it with the release worksheet. The current code clears the complete A11 merge area safely.

## Contribute

Report your Excel version, platform, template language, steps to reproduce and exact warning or error. Use fictional names and amounts in examples. See [CONTRIBUTING.md](CONTRIBUTING.md).

This is an independent tool for compatible FFF / SES layouts. The third-party budget templates and Microsoft Excel are separate products and are not included in the download.

## Licence

Copyright © 2026 Joonas Kauppinen. The tool, release workbook and original repository code are licensed under **GPL-3.0-only**. See [LICENSE](LICENSE) and [COPYRIGHT.txt](COPYRIGHT.txt). The complete download includes the workbook, licence and corresponding editable source.
