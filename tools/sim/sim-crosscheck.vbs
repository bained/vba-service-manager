Option Explicit

' =====================================================================
'  Service_Manager - the cross check: the model answers the cases of
'  src\tests\crosscheck.txt
'
'      cscript //nologo tools\sim\run.wsf crosscheck
'
'  The workbook answers the same cases from the same file (modTests.
'  Test_CrossCheck) and tools\build.vbs compares the two files line by line. What
'  is compared is what the user would see: the rule that refuses a booking, the
'  numbers of a day, the plan of the daily grid - so a rule that drifts in one of
'  the two implementations cannot pass unnoticed.
'
'  Output: build\sim-crosscheck.txt - one line per case, a plan one line per slot
'  of the day. ASCII, because the harness compares it with the file of the
'  workbook line by line; the lines the harness skips start with # .
' =====================================================================

Sub RunModelCrosscheck()
    Dim cases, list, i, lines, result

    ModelInit

    cases = FixtureCases()
    If Len(cases) = 0 Then
        WScript.Echo "no cases in " & FixturePath()
        WScript.Quit 1
    End If

    list = Split(cases, vbCrLf)
    lines = ""
    For i = 0 To UBound(list)
        result = CaseResult(CaseKind(list(i)), CaseName(list(i)), CaseArguments(list(i)))
        If Len(result) > 0 Then
            If Len(lines) > 0 Then lines = lines & vbCrLf
            lines = lines & result
        End If
    Next

    WriteCrosscheckFile lines, UBound(list) + 1
    WScript.Quit 0
End Sub

' "rule|off-step|client_id=1;..." - the three parts of a case line.
Function CaseKind(caseLine)
    CaseKind = Split(caseLine, "|")(0)
End Function

Function CaseName(caseLine)
    Dim parts

    parts = Split(caseLine, "|")
    If UBound(parts) >= 1 Then CaseName = parts(1)
End Function

' Everything after the second "|" (the arguments of the case, "" when it has none).
Function CaseArguments(caseLine)
    Dim first, second

    first = InStr(caseLine, "|")
    If first = 0 Then Exit Function
    second = InStr(first + 1, caseLine, "|")
    If second = 0 Then Exit Function

    CaseArguments = Mid(caseLine, second + 1)
End Function

' The answer of the model for one case, in the format kind|name|value... A "plan"
' answers with one line per slot, in the order the grid draws them.
Function CaseResult(kind, caseName, args)
    Dim values, parts, planLines, plan, i, resultText, lineText

    Select Case kind
        Case "rule"
            Set values = PayloadFromArgs(args)
            CaseResult = "rule|" & caseName & "|" & RuleCode(values)

        Case "normalise"
            ' The validator fills the normalised values into the payload (ISO date,
            ' HH:MM, minutes) - the case shows what a writer would store.
            Set values = PayloadFromArgs(args)
            RuleCode values
            CaseResult = "normalise|" & caseName & "|" & PayloadText(values, "date") & _
                         "|" & PayloadText(values, "start") & _
                         "|" & PayloadText(values, "duration")

        Case "count"
            CaseResult = "count|" & caseName & "|" & CStr(AppointmentsOn(caseName))

        Case "revenue"
            CaseResult = "revenue|" & caseName & "|" & CStr(MoneyCents(RevenueOn(caseName)))

        Case "freeday"
            CaseResult = "freeday|" & caseName & "|" & CStr(FreeSlotsOn(caseName))

        Case "label"
            CaseResult = "label|" & caseName & "|" & _
                         AppointmentLabel(CLng(NumberValue(caseName)))

        Case "busy"
            CaseResult = "busy|" & caseName & "|" & BusyIntervalsOn(caseName, 0)

        Case "conflict-start"
            parts = Split(args, "|")
            CaseResult = "conflict-start|" & caseName & "|" & _
                         FirstConflictStart(caseName, TimeToMinutes(parts(0)), _
                                            CLng(NumberValue(parts(1))), 0)

        Case "plan"
            plan = DailyPlan(caseName)
            planLines = ""
            For i = 0 To UBound(plan)
                If Len(planLines) > 0 Then planLines = planLines & vbCrLf
                planLines = planLines & "plan|" & caseName & "|" & plan(i)
            Next
            CaseResult = planLines

        Case "report"
            ' The period travels in the arguments and lands in the two cells, exactly
            ' as the page does it; then the report is written row by row, the fields
            ' padded so a heading reads like the empty cells of the grid ("Days||").
            parts = Split(args, "|")
            gBook.WriteName NAME_REPORT_FROM, DateFormatText(CStr(parts(0)))
            gBook.WriteName NAME_REPORT_TO, DateFormatText(CStr(parts(1)))
            plan = Split(ReportRows(ReportFromIso(), ReportToIso()), vbCrLf)

            planLines = ""
            For i = 0 To UBound(plan)
                ' A separator line of the report is empty: VBScript's Split("") gives an
                ' empty array, so it is handled before the fields are read.
                lineText = CStr(plan(i))
                If Len(lineText) = 0 Then
                    resultText = "||"
                Else
                    parts = Split(lineText, "|")
                    resultText = CStr(parts(0)) & "|"
                    If UBound(parts) >= 1 Then resultText = resultText & CStr(parts(1))
                    resultText = resultText & "|"
                    If UBound(parts) >= 2 Then resultText = resultText & CStr(parts(2))
                End If

                If Len(planLines) > 0 Then planLines = planLines & vbCrLf
                planLines = planLines & "report|" & caseName & "|" & resultText
            Next
            CaseResult = planLines

        Case "config-rule"
            ' The rule code of a page: "" while the settings are fine, otherwise the
            ' first problem (the sentence the user reads is built from it in modText).
            Set values = ConfigFromArgs(args)
            CaseResult = "config-rule|" & caseName & "|" & ConfigRuleCode(values)

        Case "config-values"
            ' What the page stores: the canonical form (HH:MM, whole numbers, the money
            ' code up, the lists trimmed, the dates ISO). Only a page that is fine is
            ' normalised, so every case for this kind carries values the rules accept.
            Set values = ConfigFromArgs(args)
            ConfigNormalise values
            CaseResult = "config-values|" & caseName & "|" & ConfigValuesLine(values)

        Case "config-defaults"
            ' The built in values of section 16.1, in the order of the sheet.
            Set values = ConfigDefaultValues()
            CaseResult = "config-defaults|" & caseName & "|" & ConfigValuesLine(values)

        Case "config-dates"
            ' The Pick button of the non working dates: one day is added to the list
            ' (ISO, in the order of the calendar, without repeats). The list travels in
            ' the arguments, so the case does not depend on what a cell holds.
            parts = Split(args, "|")
            resultText = ""
            If UBound(parts) >= 1 Then resultText = AppendIsoDate(CStr(parts(1)), CStr(parts(0)))
            CaseResult = "config-dates|" & caseName & "|" & resultText

        Case Else
            WScript.Echo "unknown case kind: " & kind
            CaseResult = ""
    End Select
End Function

' The Config page of the fixtures, with the arguments of a case on top of it:
' "i_config_step=15" is the cell the user typed over, the fixture file holds the
' rest. The keys are the names of the cells, the same ones the fixture uses, so a
' case reads like the page it describes.
Function ConfigFromArgs(args)
    Dim values, pairs, i, item, separator

    Set values = ConfigValues()
    If Len(args) > 0 Then
        pairs = Split(args, ";")
        For i = 0 To UBound(pairs)
            item = CStr(pairs(i))
            separator = InStr(item, "=")
            If separator > 0 Then
                values(Trim(Left(item, separator - 1))) = Mid(item, separator + 1)
            End If
        Next
    End If

    Set ConfigFromArgs = values
End Function

' The eight values of a page in the order of the sheet, joined by "|" - the shape
' the workbook answers with as well (both read the page in the same order).
Function ConfigValuesLine(values)
    Dim keys, i, text

    keys = values.Keys
    text = ""
    For i = 0 To UBound(keys)
        If i > 0 Then text = text & "|"
        text = text & CStr(values(keys(i)))
    Next

    ConfigValuesLine = text
End Function

' "client_id=1;service_id=1;date=25.09.2026;start=9:5;duration=045" - the payload
' of a booking as a case writes it. The keys are the column names of
' db_appointments and an empty value is a key with nothing after the = (the VBA
' reader of the workbook splits the same way).
Function PayloadFromArgs(args)
    Dim values, pairs, i, item, separator

    Set values = CreateObject("Scripting.Dictionary")
    If Len(args) > 0 Then
        pairs = Split(args, ";")
        For i = 0 To UBound(pairs)
            item = CStr(pairs(i))
            separator = InStr(item, "=")
            If separator > 0 Then
                values(Trim(Left(item, separator - 1))) = Mid(item, separator + 1)
            End If
        Next
    End If

    Set PayloadFromArgs = values
End Function

' Money crosses the lane as cents: a Double written by CStr() would follow the
' regional settings of the machine and the two files would differ on the decimal
' separator alone (§13). Both lanes round the same way (CLng).
Function MoneyCents(value)
    MoneyCents = CLng(CDbl(value) * 100)
End Function

' build\sim-crosscheck.txt, next to the file of the workbook.
Sub WriteCrosscheckFile(lines, caseCount)
    Dim fso, buildDir, path, file

    Set fso = CreateObject("Scripting.FileSystemObject")
    buildDir = fso.BuildPath(ProjectRoot(), "build")
    If Not fso.FolderExists(buildDir) Then fso.CreateFolder buildDir

    path = fso.BuildPath(buildDir, "sim-crosscheck.txt")
    Set file = fso.CreateTextFile(path, True, False)
    file.WriteLine "# " & CStr(caseCount) & " cases of " & FixturePath()
    file.WriteLine "# answers of the script model (tools\sim), " & Now
    file.Write lines & vbCrLf
    file.Close

    WScript.Echo "cross check: " & CStr(caseCount) & " cases -> " & path
End Sub
