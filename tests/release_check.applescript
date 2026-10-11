-- Open the exact release FFF_Budget_Check.xls with macros enabled first.
-- Builds a fictional budget from scratch; never opens a production budget.
-- Argument: absolute output path for the disposable .xls fixture.
on requireCheck(checker, ruleName)
    tell application "Microsoft Excel"
        set rules to value of range "D13:D100" of checker as text
    end tell
    if rules does not contain ruleName then error "Missing check: " & ruleName
end requireCheck

on countCheck(checker, ruleName)
    tell application "Microsoft Excel"
        set rules to value of range "D13:D100" of checker
    end tell
    set n to 0
    repeat with ruleRow in rules
        if item 1 of ruleRow = ruleName then set n to n + 1
    end repeat
    return n
end countCheck

on run argv
    tell application "Microsoft Excel"
        set donor to workbook "FFF_Budget_Check.xls"
        if count of worksheets of donor is not 1 then error "Expected standalone release."
        copy worksheet worksheet "Budget Check" of donor
        set testWorkbookName to name of active workbook
        set w to workbook testWorkbookName
        set a to worksheet "Budget Check" of w
        make new worksheet at end of w with properties {name:"Erittely"}
        make new worksheet at end of w with properties {name:"Yhteenveto"}
        make new worksheet at end of w with properties {name:"Summary"}
        set s to worksheet "Erittely" of w
        set fi to worksheet "Yhteenveto" of w
        set en to worksheet "Summary" of w
        set value of range "H5" of s to "Line total"
        set value of range "I5" of s to "Group total"
        set grandFormula to "=SUM("
        repeat with code from 1 to 21
            set firstRow to 6 + 8 * (code - 1)
            set payRow to firstRow + 1
            set holidayRow to firstRow + 2
            set socialRow to firstRow + 3
            set futurePay to firstRow + 4
            set futureHoliday to firstRow + 5
            set futureSocial to firstRow + 6
            set totalRow to firstRow + 7
            set value of range ("A" & firstRow) of s to code
            set value of range ("B" & firstRow) of s to "Category " & code
            set value of range ("B" & payRow) of s to "Role " & code
            set value of range ("B" & futurePay) of s to "Future role " & code
            repeat with baseRow in {payRow, futurePay}
                set r to baseRow as integer
                set value of range ("A" & r & ":A" & (r + 2)) of s to "03"
                set value of range ("C" & r) of s to "nimi"
                set value of range ("C" & (r + 1)) of s to "lomakorvaus"
                set value of range ("C" & (r + 2)) of s to "palkansivukulut"
                set formula of range ("H" & r) of s to "=IF(F" & r & "=0,D" & r & "*G" & r & ",D" & r & "*F" & r & "*G" & r & ")"
                set formula of range ("H" & (r + 1)) of s to "=H" & r & "*D" & (r + 1)
                set formula of range ("H" & (r + 2)) of s to "=(H" & r & "+H" & (r + 1) & ")*D" & (r + 2)
                set formula of range ("I" & (r + 2)) of s to "=SUM(H" & r & ":H" & (r + 2) & ")"
            end repeat
            set value of range ("D" & payRow) of s to 1
            set value of range ("G" & payRow) of s to 100
            set value of range ("D" & holidayRow) of s to 0.15
            set value of range ("D" & socialRow) of s to 0.3
            set value of range ("B" & totalRow) of s to "Category " & code & " yhteensä"
            set formula of range ("I" & totalRow) of s to "=SUM(I" & payRow & ":I" & futureSocial & ")"
            repeat with summarySheet in {fi, en}
                set sh to contents of summarySheet
                set value of range ("B" & (code + 9)) of sh to code
                set formula of range ("H" & (code + 9)) of sh to "=Erittely!I" & totalRow
            end repeat
            if code > 1 then set grandFormula to grandFormula & ","
            set grandFormula to grandFormula & "I" & totalRow
        end repeat
        set value of range "B175" of s to "Grand total"
        set formula of range "I175" of s to grandFormula & ")"
        repeat with summarySheet in {fi, en}
            set sh to contents of summarySheet
            set value of range "C32" of sh to "Grand total"
            set formula of range "H32" of sh to "=Erittely!I175"
        end repeat
        close donor saving no
        activate object w
        -- A whole-sheet copy into a new workbook retains this release's Sheet1 code name.
        set beforeData to formula of range "A1:I175" of s
        run VB macro "Sheet1.RunChecks"
        set baseline to value of range "A4" of a
        if baseline does not start with "0 findings" then error "Baseline: " & baseline
        if formula of range "A1:I175" of s is not beforeData then error "Checker altered budget."

        set value of range "B7" of s to "Järjestäjä"
        set value of range "D8:D9" of s to 0
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Paid line without fringe")
        set feeReport to value of range "A13:F100" of a as text
        if feeReport does not contain "Järjestäjä - lomakorvaus" or feeReport does not contain "fee/invoice/exempt payment" then error "Role/fee context missing."
        set value of range "D8" of s to 1.35
        set value of range "D9" of s to 0.3
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Fringe outside 0-100%")
        set value of range "D8" of s to 0.15

        set formula of range "I13" of s to "=SUM(I7:I9)"
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Possible subtotal gap")
        if my countCheck(a, "Possible subtotal gap") is not 1 then error "Expected one gap."
        if my countCheck(a, "Summary link omits detail lines") is not 0 then error "Duplicate summary gap."
        if value of range "A4" of a does not start with "1 findings" then error "Unexpected zero-gap findings."
        set gapReport to value of range "F13:F100" of a as text
        if gapReport does not contain "H10 H11 H12" or gapReport does not contain "=SUM(I7:I9)" then error "Gap guidance missing."
        set formula of range "H10" of en to "=Erittely!I9"
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Summary link omits detail lines")
        set formula of range "H10" of en to "=Erittely!I13"
        set value of range "D10" of s to 1
        set value of range "G10" of s to 60
        set value of range "D11" of s to 0.2
        set value of range "D12" of s to 0.3
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "H/I comparison needs review")
        my requireCheck(a, "H/summary comparison needs review")
        my requireCheck(a, "Fringe rates differ within section")

        set formula of range "I13" of s to "=SUM(I7:I12)"
        set value of range "D10:G10" of s to ""
        set value of range "D11:D12" of s to ""
        set value of range "G7" of s to ""
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Active line totals zero")
        set number format of range "G7" of s to "@"
        set value of range "G7" of s to "100"
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Numeric input stored as text")
        set number format of range "G7" of s to "General"
        set value of range "G7" of s to 100
        set formula of range "H7" of s to "=IF(F10=0,D10*G10,D10*F10*G10)"
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Unexpected line references")
        set value of range "H7" of s to 500
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Hardcoded total")
        set formula of range "H7" of s to "=IF(F7=0,D7*G7,D7*F7*G7)"

        set value of range "A6" of s to ""
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Missing section number")
        if my countCheck(a, "H/I comparison needs review") is not 0 then error "Invalid numbering reconciliation."
        set value of range "A6" of s to 1
        set formula of range "H32" of fi to "=Erittely!I175+10"
        run VB macro "Sheet1.RunChecks"
        my requireCheck(a, "Grand total differs from detail")
        set formula of range "H32" of fi to "=Erittely!I175"

        insert into range (entire row of range "A13" of s)
        set value of range "A13" of s to "03"
        set value of range "B13" of s to "Inserted travel"
        set value of range "D13" of s to 1
        set value of range "G13" of s to 40
        set formula of range "H13" of s to "=D13*G13"
        set formula of range "I14" of s to "=SUM(I7:I12,H13)"
        run VB macro "Sheet1.RunChecks"
        if value of range "A4" of a does not start with "0 findings" then error "Inserted row regression."
        delete range (entire row of range "A13" of s)
        set formula of range "I13" of s to "=SUM(I7:I12)"
        set beforeData to formula of range "A1:I175" of s
        run VB macro "Sheet1.RunChecks"
        set finalResult to value of range "A4" of a
        if finalResult does not start with "0 findings" then error "Restored fixture: " & finalResult
        if formula of range "A1:I175" of s is not beforeData then error "Checker altered restored budget."
        save workbook as w filename (item 1 of argv) file format Excel98to2004 file format
        set AppleScript's text item delimiters to "/"
        set outputWorkbookName to last text item of (item 1 of argv)
        set AppleScript's text item delimiters to ""
        close workbook outputWorkbookName saving no
        return {baseline, finalResult, "All injected rules detected; donor closed; budget unchanged."}
    end tell
end run
