Option Explicit
Private reportRow As Long
Private errorsFound As Long
Private reviewsFound As Long
Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    If Not Intersect(target, Me.Range("A3:F3")) Is Nothing Then
        Cancel = True
        RunChecks
    End If
End Sub
Public Sub RunChecks()
    Dim s As Worksheet, candidate As Worksheet, hits As Long
    Dim r As Long, col As Long, c As Range, v As Variant
    Dim label As String, fringe As Boolean, limit As Double, requested As String
    Dim oldEvents As Boolean, oldScreen As Boolean
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Me.Range("A13:F" & Me.Rows.Count).Clear
    Me.Range("A11").MergeArea.ClearContents
    reportRow = 13
    errorsFound = 0: reviewsFound = 0
    Me.Range("A4").Value = "Checking..."
    requested = Trim$(CStr(Me.Range("B5").Value))
    If requested <> "" Then
        Set s = Me.Parent.Worksheets(requested)
    Else
        For Each candidate In Me.Parent.Worksheets
            If candidate.Name = "Erittely" Or candidate.Name = "DetailedBudget" Then
                Set s = candidate
                hits = hits + 1
            End If
        Next candidate
        If hits <> 1 Then Err.Raise 1001, , "Enter budget sheet name in B5. Auto selection requires one Erittely or DetailedBudget."
    End If
    If s Is Me Then Err.Raise 1002, , "Choose a budget sheet, not this checker."
    If Not IsNumeric(Me.Range("B6").Value) Then Err.Raise 1003, , "B6 must contain a percentage, for example 60%."
    limit = CDbl(Me.Range("B6").Value)
    If limit <= 0 Or limit > 1 Then Err.Raise 1004, , "B6 must be above 0% and at most 100%."
    If SafeText(s.Range("H5")) = "" Or SafeText(s.Range("I5")) = "" Then Err.Raise 1005, , "Expected total headers in H5 and I5. Check template layout."
    s.Calculate
    For r = 6 To s.UsedRange.Row + s.UsedRange.Rows.Count - 1
        label = ReportLineLabel(s, r)
        fringe = (FringeKind(s, r) > 0)
        For col = 8 To 9
            Set c = s.Cells(r, col)
            v = c.Value2
            If IsError(v) Then
                AddIssue c, "ERROR", label, "Excel error", "Review this formula and its source inputs."
            ElseIf c.HasFormula Then
                If IsNumeric(Mid$(c.Formula, 2)) Then AddIssue c, "ERROR", label, "Fixed formula", "Restore a formula linked to inputs; =500 is fixed."
                If IsNumeric(v) And Len(CStr(v)) > 0 Then
                    If CDbl(v) < 0 Then AddIssue c, "REVIEW", label, "Negative total", "Confirm this credit or negative cost is intentional."
                End If
            ElseIf Len(CStr(v)) > 0 Then
                AddIssue c, "ERROR", label, "Hardcoded total", "Restore intended formula from a clean template or comparable row."
            ElseIf col = 8 And Not s.Cells(r, 9).HasFormula And (fringe Or Application.CountA(s.Range("D" & r & ":G" & r)) > 0) Then
                AddIssue c, "REVIEW", label, "Missing line total", "Inputs or fringe label exist, but H is blank. Check for a deleted formula."
            ElseIf col = 9 And IsTotalLabel(SafeText(s.Cells(r, 2)) & " " & SafeText(s.Cells(r, 3))) Then
                AddIssue c, "REVIEW", label, "Missing subtotal", "Total label suggests an I formula may be missing."
            End If
        Next col
        If fringe Then
            Set c = s.Cells(r, 4)
            v = c.Value2
            If IsError(v) Then
                AddIssue c, "ERROR", label, "Fringe error", "Review the percentage formula."
            ElseIf Len(CStr(v)) > 0 Then
                If Not IsNumeric(v) Then
                    AddIssue c, "ERROR", label, "Text fringe", "Enter a numeric percentage, for example 35%."
                ElseIf CDbl(v) < 0 Or CDbl(v) > 1 Then
                    AddIssue c, "ERROR", label, "Fringe outside 0-100%", "Check typing errors, such as 135% instead of 35%."
                ElseIf CDbl(v) > limit Then
                    AddIssue c, "REVIEW", label, "High fringe", "Above review threshold in B6; confirm intended rate."
                End If
            End If
        End If
    Next r
    CheckExtended s
    Me.Range("A4").Value = (reportRow - 13) & " findings | " & s.Name & " | " & Format$(Now, "yyyy-mm-dd hh:nn")
    Me.Range("A11").Value = errorsFound & " errors | " & reviewsFound & " to review. Results are a snapshot."
    If reportRow = 13 Then Me.Range("A13").Value = "No findings in checks performed."
    Me.Range("A13:F" & reportRow).WrapText = True
    Me.Range("A13:F" & reportRow).Rows.AutoFit
    GoTo CleanUp
Failed:
    Me.Range("A4").Value = "CHECK INCOMPLETE: " & Err.Description
    Me.Range("A11").Value = "Results incomplete. Resolve error and run again."
CleanUp:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
End Sub
Private Function FringeKind(ByVal s As Worksheet, ByVal r As Long) As Long
    Dim col As Long, label As String
    'Match the label cell, never a substring of an item or person's name.
    For col = 2 To 3
        label = LCase$(Trim$(Replace(SafeText(s.Cells(r, col)), ChrW(160), " ")))
        Do While InStr(label, "  ") > 0
            label = Replace(label, "  ", " ")
        Loop
        If Right$(label, 1) = ":" Then label = Trim$(Left$(label, Len(label) - 1))
        Select Case label
            Case "lomakorvaus", "holiday pay", "holiday compensation"
                FringeKind = 1: Exit Function
            Case "palkansivukulut", "palkan sivukulut", "sosiaalikulut", "fringe", "fringes", "social costs", "social charges", "social security", "social security costs"
                FringeKind = 2: Exit Function
        End Select
    Next col
End Function
Private Function ReportLineLabel(ByVal s As Worksheet, ByVal r As Long) As String
    Dim label As String, roleTitle As String, base As Long
    label = Trim$(SafeText(s.Cells(r, 2)) & " " & SafeText(s.Cells(r, 3)))
    If FringeKind(s, r) > 0 Then
        base = PayRow(s, r)
        If base > 0 Then
            roleTitle = Trim$(SafeText(s.Cells(base, 2)))
            If roleTitle = "" Then roleTitle = "Budget line H" & base
            label = roleTitle & " - " & label
        End If
    End If
    ReportLineLabel = label
End Function
Private Function HRole(ByVal c As Range) As Long
    '0 empty, 1 recognized cost line, 2 local total, 3 uncertain formula/value.
    Dim refs As Collection, ref As Variant, atom As Range
    Dim inputs As Boolean, totals As Boolean, qty As Boolean, rate As Boolean
    Dim fringe As Long, base As Long, f As String
    If c.Column <> 8 Then Exit Function
    If Not c.HasFormula Then
        If Len(SafeText(c)) > 0 Then HRole = 3
        Exit Function
    End If
    HRole = 3
    f = LocalFormula(c)
    If InStr(f, "!") > 0 Or InStr(f, "[") > 0 Then Exit Function
    'Named/dynamic references cannot establish a trustworthy independent sum.
    If Not SupportedHFormula(f) Then Exit Function
    Set refs = References(c)
    If refs.Count = 0 Then Exit Function
    fringe = FringeKind(c.Parent, c.Row)
    If fringe > 0 Then base = PayRow(c.Parent, c.Row)
    For Each ref In refs
        For Each atom In ref.Cells
            If atom.Column >= 8 And atom.Column <= 12 Then
                totals = True
                If fringe > 0 Then
                    If atom.Column <> 8 Or base = 0 Then Exit Function
                    If atom.Row < base Or atom.Row >= c.Row Then Exit Function
                End If
            ElseIf atom.Column = 4 Or atom.Column = 6 Or atom.Column = 7 Then
                inputs = True
                If atom.Row <> c.Row Then Exit Function
                If atom.Column = 4 Then qty = True
                If atom.Column = 7 Then rate = True
            Else
                Exit Function
            End If
        Next atom
    Next ref
    If fringe > 0 And qty And totals Then
        HRole = 1
    ElseIf inputs And qty And rate And Not totals Then
        HRole = 1
    ElseIf totals And Not inputs Then
        HRole = 2
    End If
End Function
Private Function SupportedHFormula(ByVal f As String) As Boolean
    Dim i As Long, j As Long, token As String, address As String, quoted As Boolean
    i = 1
    Do While i <= Len(f)
        If Mid$(f, i, 1) = Chr$(34) Then quoted = Not quoted
        If Not quoted And (Letter(Mid$(f, i, 1)) Or Mid$(f, i, 1) = "$") Then
            j = i: address = ReadAddress(f, j)
            If address <> "" Then
                i = j
            Else
                j = i
                Do While Letter(Mid$(f, j, 1))
                    j = j + 1
                Loop
                If j = i Then Exit Function
                token = Mid$(f, i, j - i)
                Select Case token
                    Case "IF", "IFERROR", "ROUND", "ROUNDUP", "ROUNDDOWN", "SUM", "SUBTOTAL", "MAX", "MIN", "ABS", "TRUE", "FALSE"
                    Case Else
                        Exit Function
                End Select
                i = j
            End If
        Else
            i = i + 1
        End If
    Loop
    SupportedHFormula = Not quoted
End Function
Private Function IsLineTotal(ByVal c As Range) As Boolean
    Dim role As Long
    role = HRole(c)
    IsLineTotal = (role = 1 Or role = 3)
End Function
Private Function LocalFormula(ByVal c As Range) As String
    Dim f As String, sheetName As String
    f = UCase$(CStr(c.Formula)): sheetName = UCase$(c.Parent.Name)
    f = Replace(f, "'" & Replace(sheetName, "'", "''") & "'!", "")
    f = Replace(f, sheetName & "!", "")
    LocalFormula = f
End Function
Private Function HRowList(ByVal s As Worksheet, ByVal first As Long, ByVal lastLine As Long, ByVal wanted As Long) As String
    Dim r As Long, start As Long, result As String
    For r = first To lastLine + 1
        If r <= lastLine Then
            If HRole(s.Cells(r, 8)) = wanted Then
                If start = 0 Then start = r
            ElseIf start > 0 Then
                AppendHRows result, start, r - 1: start = 0
            End If
        ElseIf start > 0 Then
            AppendHRows result, start, r - 1
        End If
    Next r
    HRowList = result
End Function
Private Sub AppendHRows(ByRef result As String, ByVal first As Long, ByVal last As Long)
    If result <> "" Then result = result & ", "
    result = result & "H" & first
    If last <> first Then result = result & ":H" & last
End Sub
Private Function SafeText(ByVal c As Range) As String
    If IsError(c.Value2) Then
        SafeText = c.Text
    Else
        SafeText = CStr(c.Value2)
    End If
End Function
Private Sub AddIssue(ByVal c As Range, ByVal severity As String, ByVal label As String, ByVal rule As String, ByVal guidance As String)
    If severity = "ERROR" Then errorsFound = errorsFound + 1 Else reviewsFound = reviewsFound + 1
    Me.Cells(reportRow, 1).Value = severity
    Me.Hyperlinks.Add Anchor:=Me.Cells(reportRow, 2), address:="", SubAddress:="'" & Replace(c.Parent.Name, "'", "''") & "'!" & c.address, TextToDisplay:=c.address(False, False)
    Me.Cells(reportRow, 3).Value = Trim$(label)
    Me.Cells(reportRow, 4).Value = rule
    Me.Cells(reportRow, 5).NumberFormat = "@"
    Me.Cells(reportRow, 5).Value = SafeText(c)
    Me.Cells(reportRow, 6).Value = guidance
    If severity = "ERROR" Then
        Me.Range("A" & reportRow & ":F" & reportRow).Interior.Color = RGB(255, 225, 220)
    Else
        Me.Range("A" & reportRow & ":F" & reportRow).Interior.Color = RGB(255, 244, 210)
    End If
    reportRow = reportRow + 1
End Sub


'Additional checks use the inspected FFF D/F/G -> H -> I/K/L layout.
Private Function CellNumber(ByVal c As Range) As Double
    If Not IsError(c.Value2) Then
        If IsNumeric(c.Value2) Then CellNumber = CDbl(c.Value2)
    End If
End Function
Private Function IsTotalLabel(ByVal txt As String) As Boolean
    IsTotalLabel = InStr(1, txt, "yhteens", vbTextCompare) > 0 Or InStr(1, " " & txt & " ", " total ", vbTextCompare) > 0
End Function
Private Function Letter(ByVal x As String) As Boolean
    Letter = Len(x) = 1 And x >= "A" And x <= "Z"
End Function
Private Function Digit(ByVal x As String) As Boolean
    Digit = Len(x) = 1 And x >= "0" And x <= "9"
End Function
Private Function ReadAddress(ByVal f As String, ByRef pos As Long) As String
    Dim start As Long, letters As Long, numbers As Long, j As Long
    start = pos: j = pos
    If Mid$(f, j, 1) = "$" Then j = j + 1
    Do While Letter(Mid$(f, j, 1))
        letters = letters + 1: j = j + 1
    Loop
    If Mid$(f, j, 1) = "$" Then j = j + 1
    Do While Digit(Mid$(f, j, 1))
        numbers = numbers + 1: j = j + 1
    Loop
    If letters > 0 And letters <= 3 And numbers > 0 Then
        ReadAddress = Mid$(f, start, j - start): pos = j
    End If
End Function
Private Function References(ByVal c As Range) As Collection
    Dim result As New Collection, f As String, i As Long, j As Long
    Dim a As String, b As String, quoted As Boolean, ref As Range
    f = LocalFormula(c)
    'Cross-sheet expressions require separate template mapping; never guess.
    If InStr(f, "!") > 0 Or InStr(f, "[") > 0 Then
        Set References = result: Exit Function
    End If
    i = 1
    Do While i <= Len(f)
        If Mid$(f, i, 1) = Chr$(34) Then quoted = Not quoted
        If Not quoted And (Letter(Mid$(f, i, 1)) Or Mid$(f, i, 1) = "$") Then
            j = i: a = ReadAddress(f, j)
            If a <> "" Then
                If Mid$(f, j, 1) = ":" Then
                    j = j + 1: b = ReadAddress(f, j)
                    If b <> "" Then a = a & ":" & b
                End If
                Set ref = Nothing
                On Error Resume Next
                Set ref = c.Parent.Range(a)
                On Error GoTo 0
                If Not ref Is Nothing Then
                    If ref.CountLarge <= 10000 Then result.Add ref
                End If
                i = j
            Else
                Do While Letter(Mid$(f, i, 1)) Or Mid$(f, i, 1) = "$"
                    i = i + 1
                Loop
            End If
        Else
            i = i + 1
        End If
    Loop
    Set References = result
End Function
Private Function PayRow(ByVal s As Worksheet, ByVal r As Long) As Long
    Dim n As Long
    For n = r - 1 To 6 Step -1
        'Account codes in A on salary/fringe rows are not section boundaries.
        If IsSectionHeader(s, n) Then Exit Function
        If s.Cells(n, 8).HasFormula And FringeKind(s, n) = 0 Then
            PayRow = n: Exit Function
        End If
    Next n
End Function
Private Sub CheckInputsAndReferences(ByVal s As Worksheet, ByVal r As Long)
    Dim c As Range, inputCell As Range, ref As Variant, atom As Range, refs As Collection
    Dim label As String, active As Boolean, missing As String, formulaWrong As Boolean
    Dim base As Long, fringe As Boolean, v As Variant
    label = ReportLineLabel(s, r)
    fringe = (FringeKind(s, r) > 0)
    Set c = s.Cells(r, 8)
    If Not c.HasFormula Then Exit Sub
    'Only standard line multiplication and known fringe lines have an inferred reference pattern.
    If Not fringe And InStr(1, c.Formula, "IF(", vbTextCompare) = 0 Then Exit Sub
    If InStr(c.Formula, "!") > 0 Or InStr(c.Formula, "[") > 0 Then
        AddIssue c, "REVIEW", label, "Line references another sheet", "This line normally uses its own D/F/G inputs. Confirm the cross-sheet reference."
        Exit Sub
    End If
    Set refs = References(c)
    If fringe Then base = PayRow(s, r)
    For Each ref In refs
        For Each atom In ref.Cells
            If fringe Then
                If atom.Column = 4 And atom.Row <> r Then formulaWrong = True
                If atom.Column = 8 And base > 0 Then
                    If atom.Row < base Or atom.Row >= r Then formulaWrong = True
                End If
            Else
                If atom.Column = 4 Or atom.Column = 6 Or atom.Column = 7 Then
                    If atom.Row <> r Then formulaWrong = True
                End If
            End If
        Next atom
    Next ref
    If formulaWrong Then AddIssue c, "REVIEW", label, "Unexpected line references", "Check for a copied formula linked to another line or salary group."
    If fringe Then
        v = s.Cells(r, 4).Value2
        If Not IsError(v) Then
            If VarType(v) = vbString And IsNumeric(v) Then AddIssue s.Cells(r, 4), "REVIEW", label, "Numeric input stored as text", "Convert the fringe input to a numeric percentage."
        End If
        If base > 0 Then
            If CellNumber(s.Cells(base, 8)) > 0 Then
                If Not IsError(s.Cells(r, 4).Value2) Then
                    If Len(SafeText(s.Cells(r, 4))) = 0 Or CellNumber(s.Cells(r, 4)) = 0 Then AddIssue s.Cells(r, 4), "REVIEW", label, "Paid line without fringe", "Positive amount in H" & base & ". Confirm whether this is salary requiring fringes or a fee/invoice/exempt payment with zero fringes."
                End If
            End If
        End If
        Exit Sub
    End If
    active = CellNumber(s.Cells(r, 4)) <> 0 Or CellNumber(s.Cells(r, 6)) <> 0 Or CellNumber(s.Cells(r, 7)) <> 0
    label = LCase$(Trim$(SafeText(s.Cells(r, 3))))
    For Each inputCell In s.Range("D" & r & ":G" & r).Cells
        If inputCell.Column <> 5 Then
            v = inputCell.Value2
            If Not IsError(v) And VarType(v) = vbString And Len(Trim$(CStr(v))) > 0 Then
                If IsNumeric(v) Then AddIssue inputCell, "REVIEW", SafeText(s.Cells(r, 2)) & " " & SafeText(s.Cells(r, 3)), "Numeric input stored as text", "Convert to a number and confirm the displayed rate or quantity."
            End If
        End If
    Next inputCell
    If active And Not IsError(c.Value2) Then
        If CellNumber(c) = 0 Then
            If CellNumber(s.Cells(r, 4)) = 0 Then missing = "quantity in D"
            If CellNumber(s.Cells(r, 7)) = 0 Then
                If missing <> "" Then missing = missing & " and "
                missing = missing & "rate in G"
            End If
            If missing <> "" Then AddIssue c, "REVIEW", SafeText(s.Cells(r, 2)) & " " & SafeText(s.Cells(r, 3)), "Active line totals zero", "Missing or zero " & missing & ". Confirm this line is intentionally unpaid. F=0 is allowed by the template."
        End If
    End If
End Sub
Private Sub TraceTotal(ByVal c As Range, ByRef covered() As Boolean, ByRef visited() As Boolean, ByVal maxRow As Long, ByVal depth As Long)
    Dim refs As Collection, ref As Variant, atom As Range, index As Long
    If depth > 40 Or c.Row > maxRow Or c.Column > 12 Or c.Column < 8 Then Exit Sub
    index = (c.Row - 1) * 5 + c.Column - 7
    If visited(index) Then Exit Sub
    visited(index) = True
    If c.Column = 8 And IsLineTotal(c) Then
        covered(c.Row) = True: Exit Sub
    End If
    If Not c.HasFormula Then Exit Sub
    Set refs = References(c)
    For Each ref In refs
        For Each atom In ref.Cells
            TraceTotal atom, covered, visited, maxRow, depth + 1
        Next atom
    Next ref
End Sub
Private Sub CheckSection(ByVal s As Worksheet, ByVal first As Long, ByVal finish As Long, ByVal code As String)
    Dim r As Long, totalRow As Long, lastLine As Long, expected As Double, role As Long, uncertain As Long
    Dim covered() As Boolean, visited() As Boolean, holes As String, n As Long, omittedValue As Double
    Dim label As String, base As Long, kind As Long, referenceRate(1 To 2) As Double
    Dim referenceRow(1 To 2) As Long, differing(1 To 2) As Boolean, rate As Double
    For r = first + 1 To finish
        role = HRole(s.Cells(r, 8))
        If role = 1 Or role = 3 Then lastLine = r
        If role = 1 Then
            If Not IsError(s.Cells(r, 8).Value2) Then expected = expected + CellNumber(s.Cells(r, 8))
        ElseIf role = 3 Then
            uncertain = uncertain + 1
        End If
        label = ReportLineLabel(s, r)
        If FringeKind(s, r) > 0 Then
            base = PayRow(s, r)
            If base > 0 Then
                If CellNumber(s.Cells(base, 8)) > 0 And CellNumber(s.Cells(r, 4)) > 0 Then
                    kind = FringeKind(s, r)
                    rate = CellNumber(s.Cells(r, 4))
                    If referenceRow(kind) = 0 Then
                        referenceRate(kind) = rate: referenceRow(kind) = r
                    ElseIf Abs(rate - referenceRate(kind)) > 1E-06 And Not differing(kind) Then
                        AddIssue s.Cells(r, 4), "REVIEW", label, "Fringe rates differ within section", "Rate differs from " & ReportLineLabel(s, referenceRow(kind)) & " in D" & referenceRow(kind) & " (" & Format$(referenceRate(kind), "0.0%") & "). Different agreements, fees or exemptions may justify this."
                        differing(kind) = True
                    End If
                End If
            End If
        End If
    Next r
    If lastLine = 0 Then Exit Sub
    For r = lastLine + 1 To finish
        If Len(Trim$(SafeText(s.Cells(r, 2)))) > 0 Then
            'The first labelled closing I subtotal belongs to this section.
            'Do not skip it in favour of an aggregate production total below it.
            totalRow = r: Exit For
        End If
    Next r
    If totalRow = 0 Then
        AddIssue s.Cells(first, 1), "REVIEW", SafeText(s.Cells(first, 2)), "Section total not identified", "Cannot reconcile this section automatically; inspect its closing subtotal."
        Exit Sub
    End If
    ReDim covered(1 To finish): ReDim visited(1 To finish * 5)
    TraceTotal s.Cells(totalRow, 9), covered, visited, finish, 0
    For r = first + 1 To lastLine
        If IsLineTotal(s.Cells(r, 8)) Then
            If Not covered(r) Then
                n = n + 1
                omittedValue = omittedValue + CellNumber(s.Cells(r, 8))
                If n <= 8 Then holes = holes & "H" & r & " "
            End If
        End If
    Next r
    If n > 0 And uncertain = 0 Then AddIssue s.Cells(totalRow, 9), "REVIEW", SafeText(s.Cells(first, 2)), "Possible subtotal gap", n & " line totals are not traced into this subtotal: " & holes & ". Check " & s.Cells(totalRow, 9).address(False, False) & " formula " & s.Cells(totalRow, 9).Formula & ". Omitted H lines currently sum to " & Format$(omittedValue, "0.00") & "; even zero lines can be missed after inputs change. Complex formulas need manual review."
    If uncertain > 0 Then
        AddIssue s.Cells(totalRow, 9), "REVIEW", SafeText(s.Cells(first, 2)), "H reconciliation needs review", "Cannot establish an independent cost sum: review " & HRowList(s, first + 1, lastLine, 3) & ". I total is " & Format$(CellNumber(s.Cells(totalRow, 9)), "0.00") & ". No H/I or H/summary mismatch is asserted for this section."
    Else
        If Not IsError(s.Cells(totalRow, 9).Value2) Then
            If Abs(CellNumber(s.Cells(totalRow, 9)) - expected) > 0.01 Then AddIssue s.Cells(totalRow, 9), "REVIEW", SafeText(s.Cells(first, 2)), "H/I comparison needs review", "Recognized H estimate " & Format$(expected, "0.00") & "; I total " & Format$(CellNumber(s.Cells(totalRow, 9)), "0.00") & ". Included: " & HRowList(s, first + 1, lastLine, 1) & ". Verify this row scope and the intended costs before changing I."
        End If
        CheckSummary s, code, expected, first, lastLine, totalRow, (n > 0)
    End If
End Sub
Private Sub CheckSummary(ByVal s As Worksheet, ByVal code As String, ByVal expected As Double, ByVal first As Long, ByVal lastLine As Long, ByVal detailTotalRow As Long, ByVal detailGapReported As Boolean)
    Dim sh As Worksheet, r As Long, v As Variant, c As Range, f As String, p As Long, target As Range
    Dim covered() As Boolean, visited() As Boolean, n As Long, k As Long
    For Each sh In Me.Parent.Worksheets
        If sh.Name = "Yhteenveto" Or sh.Name = "Summary" Then
            sh.Calculate
            For r = 10 To sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
                v = sh.Cells(r, 2).Value2
                If Not IsError(v) Then
                    If Len(CStr(v)) > 0 And IsNumeric(v) Then
                        If Val(CStr(v)) = Val(code) Then
                            Set c = sh.Cells(r, 8)
                            If IsError(c.Value2) Then
                                AddIssue c, "ERROR", sh.Name & " section " & code, "Summary total error", "Review the link to the detail section total."
                            ElseIf Abs(CellNumber(c) - expected) > 0.01 Then
                                AddIssue c, "REVIEW", sh.Name & " section " & code, "H/summary comparison needs review", "Recognized H estimate is " & Format$(expected, "0.00") & ". Included: " & HRowList(s, first + 1, lastLine, 1) & ". Verify row scope and summary mapping before changing the summary."
                            End If
                            If Not c.HasFormula Then
                                AddIssue c, "ERROR", sh.Name & " section " & code, "Summary total not a formula", "Restore the link to the detail budget."
                            Else
                                f = Replace(CStr(c.Formula), "=", ""): p = InStr(f, "!")
                                If p > 0 Then
                                    If Replace(Left$(f, p - 1), "'", "") = s.Name Then
                                        Set target = Nothing
                                        On Error Resume Next
                                        Set target = s.Range(Mid$(f, p + 1))
                                        On Error GoTo 0
                                        If Not target Is Nothing Then
                                            If target.CountLarge = 1 Then
                                                ReDim covered(1 To s.UsedRange.Row + s.UsedRange.Rows.Count - 1)
                                                ReDim visited(1 To UBound(covered) * 5)
                                                TraceTotal target, covered, visited, UBound(covered), 0
                                                n = 0
                                                For k = first + 1 To lastLine
                                                    If IsLineTotal(s.Cells(k, 8)) And Not covered(k) Then n = n + 1
                                                Next k
                                                If n > 0 Then
                                                    'A direct link repeats the already reported detail omission.
                                                    If Not (detailGapReported And target.Row = detailTotalRow And target.Column = 9) Then
                                                        AddIssue c, "REVIEW", sh.Name & " section " & code, "Summary link omits detail lines", n & " H totals cannot be traced through this summary link. Check the detail total or helper referenced by " & c.Formula & "."
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        AddIssue c, "REVIEW", sh.Name & " section " & code, "Summary links another sheet", "Confirm the section link points to the selected detail budget."
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next r
        End If
    Next sh
End Sub
Private Sub CheckGrandTotal(ByVal s As Worksheet)
    Dim sh As Worksheet, r As Long, grand As Range, label As String
    For r = 6 To s.UsedRange.Row + s.UsedRange.Rows.Count - 1
        label = SafeText(s.Cells(r, 2))
        If InStr(1, label, "kokonaiskustannukset", vbTextCompare) > 0 Or InStr(1, label, "grand total", vbTextCompare) > 0 Then Set grand = s.Cells(r, 9)
    Next r
    If grand Is Nothing Then Exit Sub
    For Each sh In Me.Parent.Worksheets
        If sh.Name = "Yhteenveto" Or sh.Name = "Summary" Then
            sh.Calculate
            For r = 10 To sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1
                label = SafeText(sh.Cells(r, 3))
                If InStr(1, label, "kokonaiskustannukset", vbTextCompare) > 0 Or InStr(1, label, "grand total", vbTextCompare) > 0 Then
                    If Not IsError(sh.Cells(r, 8).Value2) And Not IsError(grand.Value2) Then
                        If Abs(CellNumber(sh.Cells(r, 8)) - CellNumber(grand)) > 0.01 Then AddIssue sh.Cells(r, 8), "ERROR", sh.Name, "Grand total differs from detail", "Detail " & grand.address(False, False) & " is " & Format$(CellNumber(grand), "0.00") & "."
                    Else
                        AddIssue sh.Cells(r, 8), "ERROR", sh.Name, "Grand total error", "Resolve errors in the detail or summary grand total."
                    End If
                End If
            Next r
        End If
    Next sh
End Sub
Private Function IsSectionHeader(ByVal s As Worksheet, ByVal r As Long) As Boolean
    Dim col As Long, c As Range
    'FFF headers have a title in B and no description, inputs or H/I totals.
    'Do not infer a section solely from a numeric account code in column A.
    If Len(Trim$(SafeText(s.Cells(r, 2)))) = 0 Then Exit Function
    If FringeKind(s, r) > 0 Or IsTotalLabel(SafeText(s.Cells(r, 2))) Then Exit Function
    For col = 3 To 9
        Set c = s.Cells(r, col)
        If c.HasFormula Or Len(Trim$(SafeText(c))) > 0 Then Exit Function
    Next col
    IsSectionHeader = True
End Function
Private Function SectionNumber(ByVal c As Range) As Long
    Dim v As Variant, number As Double
    v = c.Value2
    If IsError(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Or Not IsNumeric(v) Then Exit Function
    number = CDbl(v)
    If number >= 1 And number <= 21 Then
        If number = Fix(number) Then
            If IsSectionHeader(c.Parent, c.Row) Then SectionNumber = CLng(number)
        End If
    End If
End Function
Private Function CheckSectionNumbers(ByVal s As Worksheet, ByVal lastRow As Long) As Boolean
    Dim firstRow(1 To 21) As Long, r As Long, number As Long, missing As String
    CheckSectionNumbers = True
    For r = 6 To lastRow
        number = SectionNumber(s.Cells(r, 1))
        If number > 0 Then
            If firstRow(number) = 0 Then
                firstRow(number) = r
            Else
                CheckSectionNumbers = False
                AddIssue s.Cells(r, 1), "REVIEW", SafeText(s.Cells(r, 2)), "Duplicate section number", "Section " & Format$(number, "00") & " also appears in A" & firstRow(number) & ". Give each section its intended unique number. Section reconciliation is skipped until numbering is complete and unique; individual cell checks still run."
            End If
        End If
    Next r
    For number = 1 To 21
        If firstRow(number) = 0 Then
            If missing <> "" Then missing = missing & ", "
            missing = missing & Format$(number, "00")
        End If
    Next number
    If missing <> "" Then
        CheckSectionNumbers = False
        AddIssue s.Cells(5, 1), "REVIEW", "Section numbering", "Missing section number", "Expected FFF section number(s) absent from recognized headers in column A: " & missing & ". Account codes on cost lines do not count as section numbers. Restore the number on any retained section header; otherwise its costs may be grouped with the preceding section. If an entire section was intentionally removed or a header uses a custom layout, this layout requires manual reconciliation. Section reconciliation is skipped; individual cell checks still run."
    End If
End Function
Private Sub CheckExtended(ByVal s As Worksheet)
    Dim r As Long, first As Long, lastRow As Long, code As String, number As Long
    Dim numberingOK As Boolean
    lastRow = s.UsedRange.Row + s.UsedRange.Rows.Count - 1
    numberingOK = CheckSectionNumbers(s, lastRow)
    For r = 6 To lastRow
        CheckInputsAndReferences s, r
        If numberingOK Then
            number = SectionNumber(s.Cells(r, 1))
            If number > 0 Then
                If first > 0 Then CheckSection s, first, r - 1, code
                first = r: code = CStr(number)
            End If
        End If
    Next r
    If numberingOK And first > 0 Then CheckSection s, first, lastRow, code
    CheckGrandTotal s
End Sub
