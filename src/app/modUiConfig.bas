Attribute VB_Name = "modUiConfig"
Option Explicit

' =====================================================================
'  Service_Manager - the Config page (Phase 6)
'
'  Eight settings drive the whole application. The page reads them, checks them
'  (one problem at a time, named), stores them in their canonical form - times as
'  09:00, whole numbers, the money code up, lists trimmed and the non-working dates
'  as ISO, because that is what the working-day rule compares - and then draws the
'  grids again, so a change of the resolution is visible at once.
'
'  The model in tools\sim carries the same rules under the same names
'  (ConfigRuleCode, ConfigNormalise), so the two lanes answer alike.
' =====================================================================

' The eight cells of the page, in the order of the sheet.
Private Function ConfigNames() As Variant
    ConfigNames = Array(NAME_CONFIG_DAY_START, NAME_CONFIG_DAY_END, NAME_CONFIG_STEP, _
                        NAME_CONFIG_DURATION, NAME_CONFIG_CURRENCY, _
                        NAME_CONFIG_NUMBER_FORMAT, NAME_CONFIG_NONWORK_DAYS, _
                        NAME_CONFIG_NONWORK_DATES)
End Function

' The values of the page, keyed by the name of the cell.
Public Function Read() As Object
    Dim values As Object
    Dim names As Variant
    Dim index As Long

    Set values = CreateObject("Scripting.Dictionary")
    names = ConfigNames()
    For index = 0 To UBound(names)
        values(CStr(names(index))) = Trim$(modNames.ReadName(CStr(names(index))))
    Next index

    Set Read = values
End Function

' Writes the values back into their cells (the page shows what was stored).
Public Sub WriteValues(ByRef values As Object)
    Dim names As Variant
    Dim index As Long
    Dim key As String

    names = ConfigNames()
    For index = 0 To UBound(names)
        key = CStr(names(index))
        modNames.WriteName key, modApp.PayloadText(values, key)
    Next index
End Sub

' The built in values (§16.1).
Public Function DefaultValues() As Object
    Dim values As Object

    Set values = CreateObject("Scripting.Dictionary")
    values(NAME_CONFIG_DAY_START) = DEFAULT_DAY_START
    values(NAME_CONFIG_DAY_END) = DEFAULT_DAY_END
    values(NAME_CONFIG_STEP) = CStr(DEFAULT_STEP_MINUTES)
    values(NAME_CONFIG_DURATION) = CStr(DEFAULT_DURATION_MINUTES)
    values(NAME_CONFIG_CURRENCY) = DEFAULT_CURRENCY
    values(NAME_CONFIG_NUMBER_FORMAT) = DEFAULT_NUMBER_FORMAT
    values(NAME_CONFIG_NONWORK_DAYS) = DEFAULT_NONWORK_WEEKDAYS
    values(NAME_CONFIG_NONWORK_DATES) = vbNullString

    Set DefaultValues = values
End Function

' =====================================================================
'  the rules of the page (§8.2.7, the same shape as a booking)
' =====================================================================

' What is wrong with the settings, as a rule code ("" = they are fine). Only the
' first problem is reported: the user fixes one thing at a time and the message
' names it. The twin of the model's ConfigRuleCode.
Public Function RuleCode(ByRef values As Object) As String
    Dim text As String
    Dim parts As Variant
    Dim index As Long
    Dim item As String
    Dim startMinutes As Long
    Dim endMinutes As Long

    ' the working day: two times, opening before closing
    text = modApp.PayloadText(values, NAME_CONFIG_DAY_START)
    If Len(text) = 0 Then
        RuleCode = RULE_REQUIRED & "|opening time"
        Exit Function
    End If
    startMinutes = modApp.TimeToMinutes(text)
    If startMinutes < 0 Then
        RuleCode = RULE_TIME & "|opening time"
        Exit Function
    End If

    text = modApp.PayloadText(values, NAME_CONFIG_DAY_END)
    If Len(text) = 0 Then
        RuleCode = RULE_REQUIRED & "|closing time"
        Exit Function
    End If
    endMinutes = modApp.TimeToMinutes(text)
    If endMinutes < 0 Then
        RuleCode = RULE_TIME & "|closing time"
        Exit Function
    End If
    If endMinutes <= startMinutes Then
        RuleCode = RULE_DAY_ORDER & "|" & modApp.MinutesToTime(startMinutes) & "|" & _
                   modApp.MinutesToTime(endMinutes)
        Exit Function
    End If

    ' the resolution of the grids and the duration of a new booking
    If Not modApp.IsDigits(modApp.PayloadText(values, NAME_CONFIG_STEP)) Then
        RuleCode = RULE_REQUIRED & "|slot step"
        Exit Function
    End If
    If CLng(Val(modApp.PayloadText(values, NAME_CONFIG_STEP))) < 1 Then
        RuleCode = RULE_REQUIRED & "|slot step"
        Exit Function
    End If
    If Not modApp.IsDigits(modApp.PayloadText(values, NAME_CONFIG_DURATION)) Then
        RuleCode = RULE_REQUIRED & "|default duration"
        Exit Function
    End If
    If CLng(Val(modApp.PayloadText(values, NAME_CONFIG_DURATION))) < 1 Then
        RuleCode = RULE_REQUIRED & "|default duration"
        Exit Function
    End If

    ' the money
    If Len(modApp.PayloadText(values, NAME_CONFIG_CURRENCY)) = 0 Then
        RuleCode = RULE_REQUIRED & "|currency"
        Exit Function
    End If
    text = modApp.PayloadText(values, NAME_CONFIG_NUMBER_FORMAT)
    If InStr(1, text, "0", vbBinaryCompare) = 0 Then
        If InStr(1, text, "#", vbBinaryCompare) = 0 Then
            RuleCode = RULE_FORMAT & "|" & text
            Exit Function
        End If
    End If

    ' the non working weekdays: numbers from 1 (Monday) to 7 (Sunday)
    text = modApp.PayloadText(values, NAME_CONFIG_NONWORK_DAYS)
    If Len(text) > 0 Then
        parts = Split(text, ",")
        For index = 0 To UBound(parts)
            item = Trim$(CStr(parts(index)))
            If Len(item) > 0 Then
                If Not modApp.IsDigits(item) Then
                    RuleCode = RULE_WEEKDAY & "|" & item
                    Exit Function
                End If
                If CLng(Val(item)) < 1 Or CLng(Val(item)) > 7 Then
                    RuleCode = RULE_WEEKDAY & "|" & item
                    Exit Function
                End If
            End If
        Next index
    End If

    ' the non working dates: real dates, ISO or European
    text = modApp.PayloadText(values, NAME_CONFIG_NONWORK_DATES)
    If Len(text) > 0 Then
        parts = Split(text, ",")
        For index = 0 To UBound(parts)
            item = Trim$(CStr(parts(index)))
            If Len(item) > 0 Then
                If Len(modPickers.DateTextToIso(item)) = 0 Then
                    RuleCode = RULE_DATES & "|" & item
                    Exit Function
                End If
            End If
        Next index
    End If
End Function

' The sentence the user reads for the first problem ("" = the page is fine).
Public Function Problems(ByRef values As Object) As String
    Dim codeText As String
    Dim parts As Variant

    codeText = RuleCode(values)
    If Len(codeText) = 0 Then Exit Function

    parts = Split(codeText, "|")
    Select Case parts(0)
        Case RULE_REQUIRED
            Problems = modText.TextFormat(modText.KEY_ERROR_REQUIRED, parts(1))
        Case RULE_TIME
            Problems = modText.TextFormat(modText.KEY_ERROR_TIME, parts(1))
        Case RULE_DAY_ORDER
            Problems = modText.TextFormat(modText.KEY_ERROR_DAY_ORDER, parts(1), parts(2))
        Case RULE_FORMAT
            Problems = modText.TextFormat(modText.KEY_ERROR_FORMAT, parts(1))
        Case RULE_WEEKDAY
            Problems = modText.TextFormat(modText.KEY_ERROR_WEEKDAY, parts(1))
        Case RULE_DATES
            Problems = modText.TextFormat(modText.KEY_ERROR_DATES, parts(1))
        Case Else
            Problems = codeText
    End Select
End Function

' The settings in the form the application stores: times as 09:00, whole numbers,
' the money code up, the lists trimmed, the non-working dates as ISO (that is what
' the working-day rule compares, whatever the cell shows).
Public Sub Normalise(ByRef values As Object)
    values(NAME_CONFIG_DAY_START) = modApp.MinutesToTime(modApp.TimeToMinutes( _
        modApp.PayloadText(values, NAME_CONFIG_DAY_START)))
    values(NAME_CONFIG_DAY_END) = modApp.MinutesToTime(modApp.TimeToMinutes( _
        modApp.PayloadText(values, NAME_CONFIG_DAY_END)))
    values(NAME_CONFIG_STEP) = CStr(CLng(Val(modApp.PayloadText(values, NAME_CONFIG_STEP))))
    values(NAME_CONFIG_DURATION) = CStr(CLng(Val(modApp.PayloadText(values, _
                                                                     NAME_CONFIG_DURATION))))
    values(NAME_CONFIG_CURRENCY) = UCase$(modApp.PayloadText(values, NAME_CONFIG_CURRENCY))
    values(NAME_CONFIG_NONWORK_DAYS) = JoinedList(modApp.PayloadText( _
                                                  values, NAME_CONFIG_NONWORK_DAYS))
    values(NAME_CONFIG_NONWORK_DATES) = JoinedIsoDates(modApp.PayloadText( _
                                                   values, NAME_CONFIG_NONWORK_DATES))
End Sub

' A list of comma separated items, trimmed, in order and without repeats
' ("7, 6, 7" -> "6,7").
Private Function JoinedList(ByVal text As String) As String
    JoinedList = AddSorted(text, vbNullString)
End Function

' The same list as ISO dates ("25.09.2026, 01.10.2026" -> "2026-09-25,2026-10-01");
' a text that is not a date is dropped. ISO sorts by itself, so one helper serves the
' weekdays and the dates.
Private Function JoinedIsoDates(ByVal text As String) As String
    JoinedIsoDates = AppendIsoDate(text, vbNullString)
End Function

' The list of the non working dates with one day more (§16.1): ISO, in the order of
' the calendar and without repeats, because a person reads the cell - and a day that
' is already there changes nothing. The Pick button of the cell calls this, and
' tools\sim\sim-model.vbs carries the same function, so the cross check pins the two
' together.
Public Function AppendIsoDate(ByVal listText As String, ByVal isoText As String) As String
    Dim parts As Variant
    Dim index As Long
    Dim result As String

    parts = Split(listText, ",")
    For index = 0 To UBound(parts)
        result = AddSorted(result, modPickers.DateTextToIso(Trim$(CStr(parts(index)))))
    Next index

    AppendIsoDate = AddSorted(result, modPickers.DateTextToIso(isoText))
End Function

' One item into a comma separated list: in order, without repeats, an empty item
' changes nothing. The items are compared as text, which is what makes an ISO date
' land in the right place without a date parser.
Private Function AddSorted(ByVal listText As String, ByVal itemText As String) As String
    Dim parts As Variant
    Dim items() As String
    Dim count As Long
    Dim index As Long

    parts = Split(listText, ",")
    For index = 0 To UBound(parts)
        count = AddOne(items, count, Trim$(CStr(parts(index))))
    Next index
    count = AddOne(items, count, Trim$(itemText))

    For index = 0 To count - 1
        If index > 0 Then AddSorted = AddSorted & ","
        AddSorted = AddSorted & items(index)
    Next index
End Function

' One item into a sorted array without repeats; the answer is the new count.
Private Function AddOne(ByRef items() As String, ByVal count As Long, ByVal itemText As String) As Long
    Dim index As Long
    Dim position As Long

    AddOne = count
    If Len(itemText) = 0 Then Exit Function

    position = count
    For index = 0 To count - 1
        If StrComp(items(index), itemText, vbBinaryCompare) = 0 Then Exit Function
        If StrComp(items(index), itemText, vbBinaryCompare) > 0 Then
            position = index
            Exit For
        End If
    Next index

    If count = 0 Then
        ReDim items(0)          ' the first item sizes the (still empty) array
    Else
        ReDim Preserve items(0 To count)
    End If
    For index = count To position + 1 Step -1
        items(index) = items(index - 1)
    Next index
    items(position) = itemText

    AddOne = count + 1
End Function

' =====================================================================
'  the buttons: config.apply and config.defaults
' =====================================================================

' Checks the page, stores what is fine in its canonical form and draws the grids
' again (the slot step is the resolution of both of them, §6.9).
Public Sub Apply(ByRef data As Object, ByRef result As Object)
    Dim values As Object
    Dim problemText As String

    Set values = Read()
    problemText = Problems(values)
    If Len(problemText) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = problemText
        Exit Sub
    End If

    Normalise values
    WriteValues values
    modDashboard.Render

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_CONFIG_APPLIED, _
        modNames.ReadName(NAME_CONFIG_DAY_START), modNames.ReadName(NAME_CONFIG_DAY_END), _
        modNames.ReadName(NAME_CONFIG_STEP))
    result(modRoutes.RESULT_OK) = True
End Sub

' Back to the built in values of §16.1 (and the grids follow).
Public Sub Defaults(ByRef data As Object, ByRef result As Object)
    Dim values As Object

    Set values = DefaultValues()
    WriteValues values
    modDashboard.Render

    result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_CONFIG_DEFAULTS)
    result(modRoutes.RESULT_OK) = True
End Sub

' =====================================================================
'  the two buttons of the non working dates
' =====================================================================

' b_config_nonwork_pick: the calendar writes one day into the list of the non
' working dates. The list is not replaced - that is what the other date fields do -
' the days pile up, because that is what a list of free days is (§16.1); the day
' lands in its place in the calendar and a day that is already there changes
' nothing (AppendIsoDate, the twin of the model). Only the cell changes: the user
' presses Apply settings, like with the pickers of Reports.
Public Sub OnPickNonwork(ByRef data As Object, ByRef result As Object)
    Dim cellName As String
    Dim picked As String
    Dim listText As String

    cellName = modApp.PayloadText(data, modRoutes.KEY_CELL)
    If Len(cellName) = 0 Then cellName = NAME_CONFIG_NONWORK_DATES

    If StrComp(cellName, NAME_CONFIG_NONWORK_DATES, vbTextCompare) <> 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_REQUIRED, NAME_CONFIG_NONWORK_DATES)
        Exit Sub
    End If

    listText = modNames.ReadName(NAME_CONFIG_NONWORK_DATES)
    picked = modApp.PayloadText(data, modRoutes.KEY_VALUE)
    If Len(picked) = 0 Then
        picked = modPickers.PickDate(modPickers.DateTextToIso(listText))
    Else
        picked = modPickers.NormaliseDateText(picked)
    End If

    If Len(picked) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CANCELLED)
        Exit Sub
    End If

    modNames.WriteName NAME_CONFIG_NONWORK_DATES, AppendIsoDate(listText, picked)
    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_CONFIG_DATE_ADDED, _
        modPickers.DateFormatText(picked))
    result(modRoutes.RESULT_OK) = True
End Sub

' b_config_nonwork_clear: the whole list goes. A list of days is easier to drop than
' to correct day by day; Apply settings stores the empty cell.
Public Sub OnClearNonwork(ByRef data As Object, ByRef result As Object)
    modNames.WriteName NAME_CONFIG_NONWORK_DATES, vbNullString

    result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_CONFIG_DATES_CLEARED)
    result(modRoutes.RESULT_OK) = True
End Sub
