# Development

Git tracks readable VBA; Excel runs the workbook. The complete executed source is src/vba/BudgetCheckWorksheet.vba. It belongs to the **Budget Check worksheet document module**, with `Me` and the double-click event. Do not import it as an ordinary standard module. The worksheet code name may change when copied into a recipient workbook.

## Change workflow

1. Preserve the current download and work on a separate development copy on a Git branch.
2. Edit the source, then replace the complete Budget Check worksheet module in Excel's Visual Basic Editor. Compile the project.
3. Test with fictional budgets. Check zero-value omissions, account codes, missing section markers, moved rows, fee/fringe context, helper totals and direct I totals. Confirm that checking leaves detail inputs and formulas unchanged.
4. Copy the whole checker sheet into a separate workbook, close the donor and run the copied sheet. Editing the source alone does not update existing workbooks.
5. Extract the executed VBA and compare it with committed source. Excel may change identifier case or the notation of a numeric literal; the published source must match the actual release exactly.
6. Commit source, documentation and evidence, then tag a verified version.

## Packaging

Use the tested, blank, one-sheet FFF_Budget_Check.xls as the open donor. Run `src/tools/package_release.applescript` with a separate full output .xls path. It copies only the checker, clears document properties and inherited names, preserves legacy format, and closes the original without saving. It never opens a budget.

On the closed release copy, run `src/tools/scrub_library_path.py` to replace the local Mac username in the Forms library reference with a fixed-width generic placeholder. Worksheet bytes, VBA source and library GUIDs remain unchanged. Do not save the sanitised file in Excel before publishing; a later save can regenerate local metadata. Use a separate copy for testing.

Install the read-only audit dependencies with `python -m pip install -r src/tools/requirements-audit.txt`, then run:

```sh
python src/tools/audit_release.py release/FFF_Budget_Check.xls
```

This checks the clean donor, exact embedded source and metadata, and writes the inventory, inspection report and workbook checksum. It does not replace native execution.

Open that exact final release in desktop Excel with macros enabled. Run `tests/release_check.applescript` with a full disposable output .xls path. It builds a fictional 21-section budget from scratch, copies the checker, closes the donor, injects errors and verifies the restored baseline. The test assumes the release's Sheet1 code name survives copying into a new workbook, not copying into an arbitrary existing workbook. Do not run it against a production workbook.

Keep the download filename FFF_Budget_Check.xls stable. Use semantic tags (v1.0.0), update CHANGELOG and release notes, and package the workbook with its licence and tagged editable source. Upload binaries through GitHub Releases. Git ignores spreadsheets, private folders and generated test output.

The original development folder and its backups are separate from this repository. Do not copy them wholesale into a release.
