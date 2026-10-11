-- Create a separate clean donor from the known one-sheet checker.
-- Open the tested FFF_Budget_Check.xls first; no other workbook is accessed.
on run argv
    tell application "Microsoft Excel"
        set donor to workbook "FFF_Budget_Check.xls"
        if count of worksheets of donor is not 1 then error "Expected only Budget Check in the donor."
        copy worksheet worksheet "Budget Check" of donor
        set w to active workbook
        set a to worksheet "Budget Check" of w
        set value of range "B5" of a to ""
        set value of range "B6" of a to 0.6
        set value of range "A4" of a to "Ready - copy this tab into your budget to start."
        set value of range "A11" of a to ""
        if value of range "A13" of a is not "" then error "Donor must have a blank report."
        repeat 4 times
            repeat with n in name of every named item of w
                delete named item (contents of n) of w
            end repeat
        end repeat
        repeat with propertyName in {"Author", "Last Author", "Manager", "Company", "Subject", "Keywords", "Comments", "Category"}
            set value of document property (contents of propertyName) of w to ""
        end repeat
        set value of document property "Title" of w to "FFF Budget Check"
        repeat with propertyName in name of every custom document property of w
            delete custom document property (contents of propertyName) of w
        end repeat
        set remove personal information of w to true
        if count of worksheets of w is not 1 or count of named items of w is not 0 then error "Unexpected sheets or names."
        set links to get link sources w
        if links is not "" and links is not {} then error "External links remain."
        close donor saving no
        save workbook as w filename (item 1 of argv) file format Excel98to2004 file format
        close w saving no
        return "Clean standalone checker saved."
    end tell
end run
