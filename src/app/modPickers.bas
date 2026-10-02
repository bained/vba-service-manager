Attribute VB_Name = "modPickers"
Option Explicit

' =====================================================================
'  Service_Manager - the two pickers and their pure helpers (Phase 3)
'
'  PickDate and PickTime show the modal forms; every rule the forms use is a
'  pure function here, so a headless test never has to open a dialog (§9.1).
'  With the dialogs switched off (modUi.Interactive = False) the picker answers
'  the current value: a test drives the flow through form.pick_date /
'  form.pick_time with data("value") instead (§9.4).
' =====================================================================

' The date the user picked; the current value when the dialogs are off and ""
' when the picker was cancelled.
Public Function PickDate(ByVal initialText As String) As String
    Dim picker As frmDatePicker
    Dim picked As String

    picked = NormaliseDateText(initialText)
    If Not modUi.InteractiveUi() Then
        PickDate = picked
        Exit Function
    End If

    On Error GoTo Failed
    Set picker = New frmDatePicker
    picker.InitialiseWithDate initialText
    picker.Show vbModal
    picked = picker.PickedDateText
    Unload picker
    Set picker = Nothing

    PickDate = picked
    Exit Function

Failed:
    Err.Clear
    PickDate = NormaliseDateText(initialText)
End Function

' The time the user picked ("09:30"), the current value when the dialogs are off
' and "" when the picker was cancelled. The field decides the mode of the form
' (§9.3) and the duration is used to grey the slots that would clash.
Public Function PickTime(ByVal fieldKey As String, ByVal initialText As String, _
                         ByVal durationMinutes As Long, ByVal busyText As String) As String
    Dim picker As frmTimePicker
    Dim picked As String
    Dim windowStart As Long
    Dim windowEnd As Long

    picked = NormaliseTimeText(initialText)
    If Not modUi.InteractiveUi() Then
        PickTime = picked
        Exit Function
    End If

    ' The start of an appointment lives inside the working day of Config, so the
    ' picker offers only the hours that day has (09:00-19:00 -> 09..18, 20.09.2026).
    ' A duration is not a clock time: it keeps the whole day.
    If StrComp(fieldKey, FIELD_DURATION, vbTextCompare) = 0 Then
        windowStart = 0
        windowEnd = MINUTES_PER_DAY
    Else
        windowStart = modAppointments.DayStartMinutes()
        windowEnd = modAppointments.DayEndMinutes()
    End If

    On Error GoTo Failed
    Set picker = New frmTimePicker
    picker.InitialiseWithTime fieldKey, picked, durationMinutes, busyText, _
                              windowStart, windowEnd
    picker.Show vbModal
    picked = picker.PickedTimeText
    Unload picker
    Set picker = Nothing

    PickTime = picked
    Exit Function

Failed:
    Err.Clear
    PickTime = NormaliseTimeText(initialText)
End Function

' The first and the last hour a start may have: the working day of Config, where
' the last hour is the one the day still contains (09:00-19:00 -> 9 To 18). A window
' that makes no sense (empty, reversed) falls back to the whole day, so a broken
' Config never leaves the picker without options.
Public Function StartHourFirst(ByVal windowStartMinutes As Long, _
                               ByVal windowEndMinutes As Long) As Long
    If windowEndMinutes <= windowStartMinutes Then Exit Function
    If windowStartMinutes < 0 Then Exit Function

    StartHourFirst = windowStartMinutes \ MINUTES_PER_HOUR
    If StartHourFirst > 23 Then StartHourFirst = 0
End Function

Public Function StartHourLast(ByVal windowStartMinutes As Long, _
                              ByVal windowEndMinutes As Long) As Long
    StartHourLast = 23
    If windowEndMinutes <= windowStartMinutes Then Exit Function
    If windowEndMinutes <= 0 Then Exit Function

    StartHourLast = (windowEndMinutes - 1) \ MINUTES_PER_HOUR
    If StartHourLast > 23 Then StartHourLast = 23
    If StartHourLast < 0 Then StartHourLast = 23
End Function

' The client the user picked from the typed search. The name goes back to the
' caller: the id follows from it, exactly as it does for a name typed by hand
' (§9.5). The current value when the dialogs are off (a test drives the flow
' through form.pick_client with data("value") instead) and "" when cancelled.
Public Function PickClient(ByVal initialText As String) As String
    Dim picker As frmClientPicker
    Dim picked As String

    picked = Trim$(initialText)
    If Not modUi.InteractiveUi() Then
        PickClient = picked
        Exit Function
    End If

    On Error GoTo Failed
    Set picker = New frmClientPicker
    picker.InitialiseWithSearch picked
    picker.Show vbModal
    picked = picker.PickedClientText
    Unload picker
    Set picker = Nothing

    PickClient = picked
    Exit Function

Failed:
    Err.Clear
    PickClient = Trim$(initialText)
End Function

' The three parts of a search row ("id|name|phone"). A missing part comes back
' as "" - the form shows what there is and the tests read all three.
Public Sub SplitSearchRow(ByVal rowText As String, ByRef idText As String, _
                          ByRef nameText As String, ByRef phoneText As String)
    Dim parts As Variant

    parts = Split(rowText, "|")
    idText = Trim$(PartAt(parts, 0))
    nameText = Trim$(PartAt(parts, 1))
    phoneText = Trim$(PartAt(parts, 2))
End Sub

' What one result line reads as: the name, and the phone when the client has one.
Public Function SearchRowText(ByVal rowText As String) As String
    Dim idText As String
    Dim nameText As String
    Dim phoneText As String

    SplitSearchRow rowText, idText, nameText, phoneText
    SearchRowText = nameText
    If Len(phoneText) > 0 Then SearchRowText = nameText & "   " & phoneText
End Function

Private Function PartAt(ByRef parts As Variant, ByVal index As Long) As String
    If index < 0 Then Exit Function
    If index > UBound(parts) Then Exit Function

    PartAt = CStr(parts(index))
End Function

' =====================================================================
'  pure helpers: dates
' =====================================================================

' ISO text of any date text ("2026-09-21"), today when the text is unusable.
' Never through CDate/IsDate: those follow the regional settings, the ISO text
' does not (§13).
Public Function NormaliseDateText(ByVal text As String) As String
    Dim isoText As String

    isoText = DateTextToIso(text)
    If Len(isoText) > 0 Then
        NormaliseDateText = isoText
        Exit Function
    End If

    NormaliseDateText = modDashboard.FormatDateText(Date)
End Function

' True for the ISO text of a real date ("2026-09-25"). The parts are checked one
' by one, so no conversion function of the locale is involved.
Public Function IsIsoDate(ByVal text As String) As Boolean
    Dim trimmed As String

    trimmed = Trim$(text)
    If Len(trimmed) <> ISO_DATE_LENGTH Then Exit Function
    If Mid$(trimmed, 5, 1) <> "-" Or Mid$(trimmed, 8, 1) <> "-" Then Exit Function
    If Not modApp.IsDigits(Left$(trimmed, 4)) Then Exit Function
    If Not modApp.IsDigits(Mid$(trimmed, 6, 2)) Then Exit Function
    If Not modApp.IsDigits(Mid$(trimmed, 9, 2)) Then Exit Function

    IsIsoDate = IsRealDate(CLng(Left$(trimmed, 4)), CLng(Mid$(trimmed, 6, 2)), _
                           CLng(Mid$(trimmed, 9, 2)))
End Function

' The date the user reads on the screen: "25.09.2026" (day first, as Europe
' writes it). The data stays ISO text - only the screen is European (§9.2).
Public Function DateFormatText(ByVal isoText As String) As String
    Dim trimmed As String

    trimmed = Trim$(isoText)
    If Not IsIsoDate(trimmed) Then Exit Function

    DateFormatText = Mid$(trimmed, 9, 2) & "." & Mid$(trimmed, 6, 2) & "." & Mid$(trimmed, 1, 4)
End Function

' ISO text of a date the user typed or picked. Understood: "2026-09-25",
' "25.09.2026", "25/09/2026", "25-09-2026" and "25.9.26"; "" when it is not a
' date at all (a two digit year counts as 20xx).
Public Function DateTextToIso(ByVal text As String) As String
    Dim trimmed As String
    Dim parts As Variant
    Dim dayValue As Long
    Dim monthValue As Long
    Dim yearValue As Long

    trimmed = Trim$(text)
    If Len(trimmed) = 0 Then Exit Function

    If IsIsoDate(trimmed) Then
        DateTextToIso = trimmed
        Exit Function
    End If

    parts = Split(Replace(Replace(trimmed, "-", "."), "/", "."), ".")
    If UBound(parts) <> 2 Then Exit Function
    If Not modApp.IsDigits(Trim$(parts(0))) Then Exit Function
    If Not modApp.IsDigits(Trim$(parts(1))) Then Exit Function
    If Not modApp.IsDigits(Trim$(parts(2))) Then Exit Function

    dayValue = CLng(Val(parts(0)))
    monthValue = CLng(Val(parts(1)))
    yearValue = CLng(Val(parts(2)))
    If yearValue < 100 Then yearValue = 2000 + yearValue

    If Not IsRealDate(yearValue, monthValue, dayValue) Then Exit Function

    DateTextToIso = Format$(yearValue, "0000") & "-" & Format$(monthValue, "00") & "-" & _
                    Format$(dayValue, "00")
End Function

' The first date the picker offers: 1 January of the year before this one.
Public Function PickerFirstDate() As String
    PickerFirstDate = modDashboard.FormatDateText(DateSerial(Year(Date) - PICKER_YEARS_BACK, 1, 1))
End Function

' The last date the picker offers: 31 December, PICKER_YEARS_FORWARD years ahead.
Public Function PickerLastDate() As String
    PickerLastDate = modDashboard.FormatDateText( _
        DateSerial(Year(Date) + PICKER_YEARS_FORWARD, 12, 31))
End Function

' True when the date may be picked (ISO text compares as text).
Public Function IsInPickerRange(ByVal isoText As String) As Boolean
    Dim isoValue As String

    isoValue = DateTextToIso(isoText)
    If Len(isoValue) = 0 Then Exit Function

    IsInPickerRange = (isoValue >= PickerFirstDate()) And (isoValue <= PickerLastDate())
End Function

' True when the three numbers are a date of the calendar.
Private Function IsRealDate(ByVal yearValue As Long, ByVal monthValue As Long, _
                            ByVal dayValue As Long) As Boolean
    If yearValue < 1 Or yearValue > 9999 Then Exit Function
    If monthValue < 1 Or monthValue > 12 Then Exit Function
    If dayValue < 1 Then Exit Function
    If dayValue > DaysInMonth(yearValue, monthValue) Then Exit Function

    IsRealDate = True
End Function

' "September 2026" - the name of a month comes from modText, not from Format$,
' so it does not depend on the regional settings.
Public Function MonthTitle(ByVal yearValue As Long, ByVal monthValue As Long) As String
    MonthTitle = modText.MonthLongName(monthValue) & " " & CStr(yearValue)
End Function

' The first day of the month (grid offset), 1 = Monday.
Public Function MonthStartOffset(ByVal yearValue As Long, ByVal monthValue As Long) As Long
    MonthStartOffset = Weekday(DateSerial(yearValue, monthValue, 1), vbMonday) - 1
End Function

Public Function DaysInMonth(ByVal yearValue As Long, ByVal monthValue As Long) As Long
    DaysInMonth = Day(DateSerial(yearValue, monthValue + 1, 0))
End Function

' The ISO date of a cell of the 6 x 7 grid (index 1..42); "" for a padding cell.
Public Function GridDateText(ByVal yearValue As Long, ByVal monthValue As Long, _
                             ByVal gridIndex As Long) As String
    Dim offset As Long
    Dim dayNumber As Long

    offset = MonthStartOffset(yearValue, monthValue)
    dayNumber = gridIndex - offset

    If dayNumber < 1 Then Exit Function
    If dayNumber > DaysInMonth(yearValue, monthValue) Then Exit Function

    ' Built from the numbers, not with Format$(date, "yyyy-mm-dd"): the format
    ' codes follow the regional settings of the machine, the ISO text does not.
    GridDateText = Format$(yearValue, "0000") & "-" & Format$(monthValue, "00") & "-" & _
                   Format$(dayNumber, "00")
End Function

' The grid position (1..42) of an ISO date inside the shown month, 0 when the
' date is not in that month.
Public Function GridIndexOf(ByVal yearValue As Long, ByVal monthValue As Long, _
                            ByVal isoText As String) As Long
    Dim index As Long

    For index = 1 To PICKER_GRID_DAYS * PICKER_GRID_WEEKS
        If StrComp(GridDateText(yearValue, monthValue, index), isoText, vbTextCompare) = 0 Then
            GridIndexOf = index
            Exit Function
        End If
    Next index
End Function

' =====================================================================
'  pure helpers: times
' =====================================================================

' "09:30" of any time text. Accepts "8:30", "08:30", "830", "8", "8,30", "8.30";
' "" when the text cannot be read as a time (§9.3).
Public Function ParseTimeText(ByVal text As String) As String
    Dim trimmed As String
    Dim hoursText As String
    Dim minutesText As String
    Dim separator As Long
    Dim hours As Long
    Dim minutes As Long

    trimmed = Replace(Replace(Replace(Trim$(text), ",", ":"), ".", ":"), " ", "")
    If Len(trimmed) = 0 Then Exit Function

    separator = InStr(1, trimmed, ":", vbBinaryCompare)
    If separator > 0 Then
        hoursText = Left$(trimmed, separator - 1)
        minutesText = Mid$(trimmed, separator + 1)
    ElseIf Len(trimmed) <= 2 Then
        hoursText = trimmed
        minutesText = "0"
    Else
        hoursText = Left$(trimmed, Len(trimmed) - 2)
        minutesText = Right$(trimmed, 2)
    End If

    If Len(hoursText) = 0 Then hoursText = "0"
    If Len(minutesText) = 0 Then minutesText = "0"
    If Not IsNumeric(hoursText) Then Exit Function
    If Not IsNumeric(minutesText) Then Exit Function

    hours = CLng(Val(hoursText))
    minutes = CLng(Val(minutesText))
    If hours < 0 Or hours > 23 Then Exit Function
    If minutes < 0 Or minutes > 59 Then Exit Function

    ParseTimeText = modApp.MinutesToTime(hours * MINUTES_PER_HOUR + minutes)
End Function

' "08:30" or "" when the text is empty or unreadable.
Public Function NormaliseTimeText(ByVal text As String) As String
    NormaliseTimeText = ParseTimeText(text)
End Function

' "7" -> "07"; "" when the text is not an hour between 0 and 23 (§9.4).
Public Function HourText(ByVal text As String) As String
    Dim value As Long
    Dim trimmed As String

    trimmed = Trim$(text)
    If Len(trimmed) = 0 Then Exit Function
    If Not IsNumeric(trimmed) Then Exit Function

    value = CLng(Val(trimmed))
    If value < 0 Or value > 23 Then Exit Function

    HourText = Format$(value, "00")
End Function

' "5" -> "05"; "" when the text is not a minute on the step (0..55 by 5).
Public Function MinuteText(ByVal text As String) As String
    Dim value As Long
    Dim trimmed As String

    trimmed = Trim$(text)
    If Len(trimmed) = 0 Then Exit Function
    If Not IsNumeric(trimmed) Then Exit Function

    value = CLng(Val(trimmed))
    If value < 0 Or value > 55 Then Exit Function
    If (value Mod TIME_STEP_MINUTES) <> 0 Then Exit Function

    MinuteText = Format$(value, "00")
End Function

' True when [startMinutes, startMinutes + durationMinutes) touches one of the busy
' intervals of the day ("09:00+30,10:30+90"). A duration of 0 means "nothing is
' busy", which is the case in the duration mode of the picker.
Public Function IsSlotBusy(ByVal startMinutes As Long, ByVal durationMinutes As Long, _
                           ByVal busyText As String) As Boolean
    Dim parts As Variant
    Dim item As Variant
    Dim itemText As String
    Dim plusPos As Long
    Dim busyStart As Long
    Dim busyDuration As Long

    If durationMinutes < 1 Then Exit Function
    If Len(Trim$(busyText)) = 0 Then Exit Function

    parts = Split(busyText, ",")
    For Each item In parts
        itemText = Trim$(CStr(item))
        plusPos = InStr(1, itemText, "+", vbBinaryCompare)

        If plusPos > 0 Then
            busyStart = modApp.TimeToMinutes(Left$(itemText, plusPos - 1))
            busyDuration = CLng(Val(Mid$(itemText, plusPos + 1)))
        Else
            busyStart = modApp.TimeToMinutes(itemText)
            busyDuration = 1
        End If

        If busyStart >= 0 Then
            If busyDuration < 1 Then busyDuration = 1
            If modAppointments.Overlaps(startMinutes, durationMinutes, _
                                        busyStart, busyDuration) Then
                IsSlotBusy = True
                Exit Function
            End If
        End If
    Next item
End Function
