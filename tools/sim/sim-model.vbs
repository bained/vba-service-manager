Option Explicit

' =====================================================================
'  Service_Manager - the rules of the model
'
'  Every function names the VBA twin it mirrors, so porting a rule back into the
'  workbook stays mechanical and the two can be compared case by case (the plan
'  is a cross check of the same fixtures, see tools\sim\README.md).
'
'  Differences that are forced by the language: VBScript has no Format() (hence
'  Pad2), no types on Dim, no Optional arguments and no ByVal.
' =====================================================================

' ---- constants (the values of modConfig/modNames) ---------------------
Const NAME_CONFIG_DAY_START = "i_config_day_start"
Const NAME_CONFIG_DAY_END = "i_config_day_end"
Const NAME_CONFIG_STEP = "i_config_step"
Const NAME_CONFIG_DURATION = "i_config_duration"
Const NAME_CONFIG_NONWORK_DAYS = "i_config_nonwork_days"
Const NAME_CONFIG_NONWORK_DATES = "i_config_nonwork_dates"
Const NAME_CONFIG_CURRENCY = "i_config_currency"
Const NAME_CONFIG_NUMBER_FORMAT = "i_config_number_format"

Const DEFAULT_DAY_START = "09:00"
Const DEFAULT_DAY_END = "19:00"
Const DEFAULT_STEP_MINUTES = 5
Const DEFAULT_DURATION_MINUTES = 10
Const DEFAULT_NONWORK_WEEKDAYS = "6,7"
Const DEFAULT_CURRENCY = "EUR"
Const DEFAULT_NUMBER_FORMAT = "#,##0.00"
Const ISO_DATE_LENGTH = 10
Const MINUTES_PER_HOUR = 60
Const MAX_DURATION_MINUTES = 1440

' ---- module state -----------------------------------------------------
Dim gBook, gTables

Sub ModelInit()
    Set gBook = BuildSimBook()
    Set gTables = LoadSimTables()
End Sub

Function ModelBook()
    Set ModelBook = gBook
End Function

Function ModelTable(sheetName)
    Set ModelTable = gTables(sheetName)
End Function

' Puts the tables back to the fixtures of sim-data.vbs: a test that creates,
' edits or deletes rows gets the same starting point as the one before it (the
' VBA batch resets its own state the same way).
Sub ModelReloadTables()
    Set gTables = LoadSimTables()
End Sub

' ---- text helpers -----------------------------------------------------

' modApp.IsDigits
Function IsDigits(text)
    Dim i, ch
    IsDigits = False
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next
    IsDigits = True
End Function

' VBScript has no Val(): a lenient number out of a text ("30" -> 30, "" -> 0). The
' parts are read by hand, because CDbl("25.00") follows the regional settings and
' would read 2500 where the dot is a thousands separator (the lesson of ??13).
Function NumberValue(text)
    Dim trimmed, sign, wholeText, fractionText, parts, value

    NumberValue = 0
    trimmed = Trim(CStr(text))
    If Len(trimmed) = 0 Then Exit Function

    sign = 1
    If Left(trimmed, 1) = "-" Then
        sign = -1
        trimmed = Mid(trimmed, 2)
    End If

    If InStr(trimmed, ".") > 0 Then
        parts = Split(trimmed, ".")
        wholeText = parts(0)
        fractionText = parts(1)
    Else
        wholeText = trimmed
        fractionText = ""
    End If

    If Not IsDigits(wholeText) Then Exit Function
    If Len(fractionText) > 0 Then
        If Not IsDigits(fractionText) Then Exit Function
    End If

    value = CDbl(wholeText)
    If Len(fractionText) > 0 Then value = value + CDbl(fractionText) / (10 ^ Len(fractionText))

    NumberValue = sign * value
End Function

' Format$(value, "00") of VBA
Function Pad2(value)
    Pad2 = Right("0" & CStr(CLng(value)), 2)
End Function

' modApp.TimeToMinutes ("09:30" -> 570, anything else -> -1)
Function TimeToMinutes(text)
    Dim parts, hours, minutes

    TimeToMinutes = -1
    If Len(Trim(text)) = 0 Then Exit Function

    parts = Split(Trim(text), ":")
    If UBound(parts) <> 1 Then Exit Function
    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function

    hours = CLng(NumberValue(parts(0)))
    minutes = CLng(NumberValue(parts(1)))
    If hours < 0 Or hours > 23 Then Exit Function
    If minutes < 0 Or minutes > 59 Then Exit Function

    TimeToMinutes = hours * MINUTES_PER_HOUR + minutes
End Function

' modApp.MinutesToTime (570 -> "09:30")
Function MinutesToTime(minutes)
    If minutes < 0 Then Exit Function
    MinutesToTime = Pad2(minutes \ MINUTES_PER_HOUR) & ":" & Pad2(minutes Mod MINUTES_PER_HOUR)
End Function


' modPickers.DaysInMonth
Function DaysInMonth(yearValue, monthValue)
    Select Case monthValue
        Case 1, 3, 5, 7, 8, 10, 12
            DaysInMonth = 31
        Case 4, 6, 9, 11
            DaysInMonth = 30
        Case 2
            If (yearValue Mod 4 = 0 And yearValue Mod 100 <> 0) Or (yearValue Mod 400 = 0) Then
                DaysInMonth = 29
            Else
                DaysInMonth = 28
            End If
        Case Else
            DaysInMonth = 0
    End Select
End Function

' modPickers.IsIsoDate (+ IsRealDate)
Function IsIsoDate(text)
    Dim trimmed, yearValue, monthValue, dayValue

    IsIsoDate = False
    trimmed = Trim(text)
    If Len(trimmed) <> ISO_DATE_LENGTH Then Exit Function
    If Mid(trimmed, 5, 1) <> "-" Or Mid(trimmed, 8, 1) <> "-" Then Exit Function
    If Not IsDigits(Left(trimmed, 4)) Then Exit Function
    If Not IsDigits(Mid(trimmed, 6, 2)) Then Exit Function
    If Not IsDigits(Mid(trimmed, 9, 2)) Then Exit Function

    yearValue = CLng(Left(trimmed, 4))
    monthValue = CLng(Mid(trimmed, 6, 2))
    dayValue = CLng(Mid(trimmed, 9, 2))
    If monthValue < 1 Or monthValue > 12 Then Exit Function
    If dayValue < 1 Or dayValue > DaysInMonth(yearValue, monthValue) Then Exit Function

    IsIsoDate = True
End Function

' modPickers.DateFormatText ("2026-09-25" -> "25.09.2026")
Function DateFormatText(isoText)
    Dim trimmed
    trimmed = Trim(isoText)
    If Not IsIsoDate(trimmed) Then Exit Function
    DateFormatText = Mid(trimmed, 9, 2) & "." & Mid(trimmed, 6, 2) & "." & Mid(trimmed, 1, 4)
End Function

' modPickers.DateTextToIso ("25.09.2026" -> "2026-09-25")
Function DateTextToIso(text)
    Dim trimmed, parts, separator, dayValue, monthValue, yearValue, i

    trimmed = Trim(text)
    If Len(trimmed) = 0 Then Exit Function
    If IsIsoDate(trimmed) Then
        DateTextToIso = trimmed
        Exit Function
    End If

    separator = ""
    For i = 1 To Len(trimmed)
        If InStr(".-/", Mid(trimmed, i, 1)) > 0 Then
            separator = Mid(trimmed, i, 1)
            Exit For
        End If
    Next
    If Len(separator) = 0 Then Exit Function

    parts = Split(trimmed, separator)
    If UBound(parts) <> 2 Then Exit Function
    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Or Not IsNumeric(parts(2)) Then Exit Function

    dayValue = CLng(NumberValue(parts(0)))
    monthValue = CLng(NumberValue(parts(1)))
    yearValue = CLng(NumberValue(parts(2)))
    If yearValue < 100 Then yearValue = 2000 + yearValue
    If monthValue < 1 Or monthValue > 12 Then Exit Function
    If dayValue < 1 Or dayValue > DaysInMonth(yearValue, monthValue) Then Exit Function

    DateTextToIso = CStr(yearValue) & "-" & Pad2(monthValue) & "-" & Pad2(dayValue)
End Function

' ---- Config (modDashboard.ConfigText/ConfigNumber) --------------------

Function ConfigText(nameText, fallback)
    Dim value
    value = gBook.ReadName(nameText)
    If Len(value) = 0 Then value = fallback
    ConfigText = value
End Function

Function ConfigNumber(nameText, fallback)
    Dim value
    value = gBook.ReadName(nameText)
    If Not IsNumeric(value) Then
        ConfigNumber = fallback
        Exit Function
    End If
    ConfigNumber = CLng(NumberValue(value))
End Function

' modAppointments.DayStartMinutes / DayEndMinutes / StepMinutes
Function DayStartMinutes()
    DayStartMinutes = TimeToMinutes(ConfigText(NAME_CONFIG_DAY_START, DEFAULT_DAY_START))
    If DayStartMinutes < 0 Then DayStartMinutes = TimeToMinutes(DEFAULT_DAY_START)
End Function

Function DayEndMinutes()
    DayEndMinutes = TimeToMinutes(ConfigText(NAME_CONFIG_DAY_END, DEFAULT_DAY_END))
    If DayEndMinutes < 0 Then DayEndMinutes = TimeToMinutes(DEFAULT_DAY_END)
End Function

Function StepMinutes()
    StepMinutes = ConfigNumber(NAME_CONFIG_STEP, DEFAULT_STEP_MINUTES)
    If StepMinutes < 1 Then StepMinutes = DEFAULT_STEP_MINUTES
End Function

' modAppointments.EndTime ("09:00" + 60 -> "10:00")
Function EndTime(startText, durationMinutes)
    Dim startMinutes

    startMinutes = TimeToMinutes(startText)
    If startMinutes < 0 Then Exit Function
    If durationMinutes < 1 Then Exit Function

    EndTime = MinutesToTime(startMinutes + durationMinutes)
End Function

' modAppointments.Overlaps (intervals that only touch do not clash)
Function Overlaps(startA, durationA, startB, durationB)
    If durationA < 1 Or durationB < 1 Then Exit Function
    Overlaps = (startA < startB + durationB) And (startB < startA + durationA)
End Function

' modAppointments.IsOnStep (a step below 1 means "no rule")
Function IsOnStep(minutes, stepValue)
    If stepValue < 1 Then
        IsOnStep = True
        Exit Function
    End If

    IsOnStep = ((minutes Mod stepValue) = 0)
End Function

' modAppointments.DurationText ("45 min" / "1 h 30 min")
Function DurationText(minutes)
    Dim hours, rest

    If minutes < 1 Then Exit Function

    hours = minutes \ MINUTES_PER_HOUR
    rest = minutes Mod MINUTES_PER_HOUR

    If hours = 0 Then
        DurationText = CStr(rest) & " min"
    Else
        DurationText = CStr(hours) & " h " & CStr(rest) & " min"
    End If
End Function

' modAppointments.IsWorkingDay (??8.2.12)
Function IsWorkingDay(isoDate)
    Dim dateValue, weekdayList, dateList, parts, i

    IsWorkingDay = False
    If Not IsIsoDate(isoDate) Then Exit Function

    dateValue = DateSerial(CLng(Left(isoDate, 4)), CLng(Mid(isoDate, 6, 2)), CLng(Mid(isoDate, 9, 2)))

    ' The cell is read as it is: an empty list means no non working weekday at all
    ' (the same reading as modAppointments.IsWorkingDay).
    weekdayList = Replace(Trim(CStr(gBook.ReadName(NAME_CONFIG_NONWORK_DAYS))), " ", "")
    If Len(weekdayList) > 0 Then
        parts = Split(weekdayList, ",")
        For i = 0 To UBound(parts)
            If CStr(Weekday(dateValue, 2)) = Trim(parts(i)) Then Exit Function
        Next
    End If

    dateList = Replace(ConfigText(NAME_CONFIG_NONWORK_DATES, ""), " ", "")
    If Len(dateList) > 0 Then
        parts = Split(dateList, ",")
        For i = 0 To UBound(parts)
            If StrComp(Trim(parts(i)), isoDate, 1) = 0 Then Exit Function
        Next
    End If

    IsWorkingDay = True
End Function

' ---- tables (modData) -------------------------------------------------

' modData.ClientName / ServiceName / ServicePrice / ServiceDuration
Function ClientName(clientId)
    Dim table, i
    Set table = ModelTable(SHEET_DB_CLIENTS)
    For i = 1 To table.RowCount
        If CLng(NumberValue(table.TextAt(i, "id"))) = CLng(NumberValue(clientId)) Then
            ClientName = table.TextAt(i, "name")
            Exit Function
        End If
    Next
End Function

Function ServiceName(serviceId)
    Dim table, i
    Set table = ModelTable(SHEET_DB_SERVICES)
    For i = 1 To table.RowCount
        If CLng(NumberValue(table.TextAt(i, "id"))) = CLng(NumberValue(serviceId)) Then
            ServiceName = table.TextAt(i, "service")
            Exit Function
        End If
    Next
End Function

Function ServicePrice(serviceId)
    Dim table, i
    Set table = ModelTable(SHEET_DB_SERVICES)
    ServicePrice = 0
    For i = 1 To table.RowCount
        If CLng(NumberValue(table.TextAt(i, "id"))) = CLng(NumberValue(serviceId)) Then
            ServicePrice = NumberValue(table.TextAt(i, "price"))
            Exit Function
        End If
    Next
End Function

Function ServiceDuration(serviceId)
    Dim table, i
    Set table = ModelTable(SHEET_DB_SERVICES)
    ServiceDuration = 0
    For i = 1 To table.RowCount
        If CLng(NumberValue(table.TextAt(i, "id"))) = CLng(NumberValue(serviceId)) Then
            ServiceDuration = CLng(NumberValue(table.TextAt(i, "duration")))
            Exit Function
        End If
    Next
End Function

' modData.IsoDateText: the ISO date of a cell or of a payload value. Text is cut
' to ten characters and nothing more - the validator does NOT read a European date
' out of it (the writers normalise the text before they save: the form does it with
' modPickers.DateTextToIso, §7.5). A text that is not ISO fails IsWorkingDay, so
' "25.09.2026" is refused as a non-working day - the behaviour both lanes show in
' the cross check. Excel cells can carry a real date as well; that branch (IsDate)
' stays in the workbook, because the model only knows text.
Function IsoDateText(value)
    Dim text

    text = Trim(CStr(value))
    If Len(text) >= ISO_DATE_LENGTH Then
        IsoDateText = Left(text, ISO_DATE_LENGTH)
    Else
        IsoDateText = text
    End If
End Function

' modData.RowExists
Function RowExists(sheetName, idValue)
    Dim table, i
    Set table = ModelTable(sheetName)
    RowExists = False
    For i = 1 To table.RowCount
        If CLng(NumberValue(table.TextAt(i, "id"))) = CLng(NumberValue(idValue)) Then
            RowExists = True
            Exit Function
        End If
    Next
End Function

' ---- the appointments of a day (modData) ------------------------------

' modData.AppointmentRowsOn: "1,3" - the data rows of that date, sorted by
' start and then by id (the order the dashboard draws them in).
Function AppointmentRowsOn(isoDate)
    Dim table, i, text, starts, ids, j, swapped, swapText, swapStart, swapId
    Dim rowList, startList, idList

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    text = ""

    For i = 1 To table.RowCount
        If StrComp(table.TextAt(i, "date"), Trim(isoDate), 1) = 0 Then
            If Len(text) > 0 Then text = text & ","
            text = text & CStr(i)
        End If
    Next

    If Len(text) = 0 Then Exit Function

    rowList = Split(text, ",")
    ReDim startList(UBound(rowList))
    ReDim idList(UBound(rowList))
    For i = 0 To UBound(rowList)
        startList(i) = TimeToMinutes(table.TextAt(CLng(rowList(i)), "start"))
        idList(i) = CLng(NumberValue(table.TextAt(CLng(rowList(i)), "id")))
    Next

    ' a plain bubble sort: the fixtures are small and VBScript has no sort
    For i = 0 To UBound(rowList) - 1
        For j = 0 To UBound(rowList) - 1 - i
            swapped = False
            If startList(j) > startList(j + 1) Then
                swapped = True
            ElseIf startList(j) = startList(j + 1) And idList(j) > idList(j + 1) Then
                swapped = True
            End If
            If swapped Then
                swapText = rowList(j): rowList(j) = rowList(j + 1): rowList(j + 1) = swapText
                swapStart = startList(j): startList(j) = startList(j + 1): startList(j + 1) = swapStart
                swapId = idList(j): idList(j) = idList(j + 1): idList(j + 1) = swapId
            End If
        Next
    Next

    AppointmentRowsOn = Join(rowList, ",")
End Function

' modData.AppointmentsOn
Function AppointmentsOn(isoDate)
    Dim rows, list
    rows = AppointmentRowsOn(isoDate)
    If Len(rows) = 0 Then
        AppointmentsOn = 0
    Else
        list = Split(rows, ",")
        AppointmentsOn = UBound(list) + 1
    End If
End Function

' modData.RevenueOn
Function RevenueOn(isoDate)
    Dim table, rows, list, i, total

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    rows = AppointmentRowsOn(isoDate)
    total = 0
    If Len(rows) > 0 Then
        list = Split(rows, ",")
        For i = 0 To UBound(list)
            total = total + ServicePrice(table.TextAt(CLng(list(i)), "service_id"))
        Next
    End If

    RevenueOn = total
End Function

' modData.RevenueInPeriod (ISO from/to, both included)
Function RevenueInPeriod(fromIso, toIso)
    Dim table, i, total

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    total = 0
    For i = 1 To table.RowCount
        If StrComp(table.TextAt(i, "date"), fromIso, 1) >= 0 Then
            If StrComp(table.TextAt(i, "date"), toIso, 1) <= 0 Then
                total = total + ServicePrice(table.TextAt(i, "service_id"))
            End If
        End If
    Next

    RevenueInPeriod = total
End Function

' ---- conflicts (modAppointments) --------------------------------------

' modAppointments.ConflictCount
Function ConflictCount(isoDate, startMinutes, durationMinutes, exceptId)
    Dim table, rows, list, i, dataRow, rowStart, rowDuration, rowId, count

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    rows = AppointmentRowsOn(isoDate)
    If Len(rows) = 0 Then Exit Function

    list = Split(rows, ",")
    For i = 0 To UBound(list)
        dataRow = CLng(list(i))
        rowId = CLng(NumberValue(table.TextAt(dataRow, "id")))
        If rowId <> exceptId Then
            rowStart = TimeToMinutes(table.TextAt(dataRow, "start"))
            rowDuration = CLng(NumberValue(table.TextAt(dataRow, "duration")))
            If rowStart >= 0 Then
                If Overlaps(startMinutes, durationMinutes, rowStart, rowDuration) Then count = count + 1
            End If
        End If
    Next

    ConflictCount = count
End Function

' modAppointments.FirstConflictStart
Function FirstConflictStart(isoDate, startMinutes, durationMinutes, exceptId)
    Dim table, rows, list, i, dataRow, rowStart, rowDuration, rowId, best

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    rows = AppointmentRowsOn(isoDate)
    If Len(rows) = 0 Then Exit Function

    best = -1
    list = Split(rows, ",")
    For i = 0 To UBound(list)
        dataRow = CLng(list(i))
        rowId = CLng(NumberValue(table.TextAt(dataRow, "id")))
        If rowId <> exceptId Then
            rowStart = TimeToMinutes(table.TextAt(dataRow, "start"))
            rowDuration = CLng(NumberValue(table.TextAt(dataRow, "duration")))
            If rowStart >= 0 Then
                If Overlaps(startMinutes, durationMinutes, rowStart, rowDuration) Then
                    If best < 0 Then
                        best = rowStart
                    ElseIf rowStart < best Then
                        best = rowStart
                    End If
                End If
            End If
        End If
    Next

    If best >= 0 Then FirstConflictStart = MinutesToTime(best)
End Function

' modAppointments.BusyIntervalsOn ("09:00+30,10:30+90")
Function BusyIntervalsOn(isoDate, exceptId)
    Dim table, rows, list, i, dataRow, durationValue, text

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    rows = AppointmentRowsOn(isoDate)
    If Len(rows) = 0 Then Exit Function

    list = Split(rows, ",")
    For i = 0 To UBound(list)
        dataRow = CLng(list(i))
        If CLng(NumberValue(table.TextAt(dataRow, "id"))) <> exceptId Then
            durationValue = CLng(NumberValue(table.TextAt(dataRow, "duration")))
            If durationValue < 1 Then durationValue = 1
            If Len(text) > 0 Then text = text & ","
            text = text & table.TextAt(dataRow, "start") & "+" & CStr(durationValue)
        End If
    Next

    BusyIntervalsOn = text
End Function

' ---- validation of a booking (??8.2.7) ---------------------------------

' A stand-in for modApp.PayloadText: read a dictionary key that may be missing.
Function PayloadText(values, keyText)
    If values Is Nothing Then Exit Function
    If Not values.Exists(keyText) Then Exit Function
    PayloadText = Trim(CStr(values(keyText)))
End Function

' modAppointments.RuleCode: the rule that refuses a booking as a code with its
' parameters ("error.step|5"), "" when the booking is fine. The code is the
' identity of the rule; the sentence the user reads is built from it in VBA
' (modText + Problems) and is therefore not duplicated here. The values carry
' client_id, service_id, date, start, duration and (while editing) id.
Function RuleCode(values)
    Dim clientId, serviceId, dateText, startText, durationRaw
    Dim startMinutes, durationValue, exceptId, endMinutes, clashes

    clientId = CLng(NumberValue(PayloadText(values, "client_id")))
    If clientId < 1 Then
        RuleCode = "error.required|client"
        Exit Function
    End If
    If Not RowExists(SHEET_DB_CLIENTS, clientId) Then
        RuleCode = "error.not_found|Client|" & CStr(clientId)
        Exit Function
    End If

    serviceId = CLng(NumberValue(PayloadText(values, "service_id")))
    If serviceId < 1 Then
        RuleCode = "error.required|service"
        Exit Function
    End If
    If Not RowExists(SHEET_DB_SERVICES, serviceId) Then
        RuleCode = "error.not_found|Service|" & CStr(serviceId)
        Exit Function
    End If

    dateText = IsoDateText(PayloadText(values, "date"))
    If Len(dateText) < ISO_DATE_LENGTH Then
        RuleCode = "error.required|date"
        Exit Function
    End If
    If Not IsWorkingDay(dateText) Then
        RuleCode = "error.nonwork|" & dateText
        Exit Function
    End If
    ' The validator fills in the normalised values, exactly like the VBA twin:
    ' the caller stores what the rules accepted (ISO date, HH:MM, minutes).
    values("date") = dateText

    startText = PayloadText(values, "start")
    startMinutes = TimeToMinutes(startText)
    If startMinutes < 0 Then
        RuleCode = "error.time|start time"
        Exit Function
    End If
    values("start") = MinutesToTime(startMinutes)

    durationRaw = PayloadText(values, "duration")
    If Len(durationRaw) = 0 Then
        RuleCode = "error.required|duration"
        Exit Function
    End If
    If Not IsNumeric(durationRaw) Then
        RuleCode = "error.required|duration"
        Exit Function
    End If

    durationValue = CLng(NumberValue(durationRaw))
    If durationValue < 1 Or durationValue > MAX_DURATION_MINUTES Then
        RuleCode = "error.required|duration"
        Exit Function
    End If
    values("duration") = CStr(durationValue)

    endMinutes = startMinutes + durationValue
    If startMinutes < DayStartMinutes() Or endMinutes > DayEndMinutes() Then
        RuleCode = "error.range|" & MinutesToTime(DayStartMinutes()) & _
                       "|" & MinutesToTime(DayEndMinutes())
        Exit Function
    End If

    If Not IsOnStep(startMinutes, StepMinutes()) Then
        RuleCode = "error.step|" & CStr(StepMinutes())
        Exit Function
    End If

    exceptId = CLng(NumberValue(PayloadText(values, "id")))
    clashes = ConflictCount(dateText, startMinutes, durationValue, exceptId)
    If clashes > 0 Then
        RuleCode = "error.conflict|" & CStr(clashes) & "|" & _
                       FirstConflictStart(dateText, startMinutes, durationValue, exceptId)
        Exit Function
    End If
End Function

' modText.WeekdayShortName (1 = Monday)
Function WeekdayShortName(weekdayIndex)
    Dim names
    names = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    If weekdayIndex < 1 Or weekdayIndex > 7 Then Exit Function
    WeekdayShortName = names(weekdayIndex - 1)
End Function

' ---- what the user reads (modData) ------------------------------------

' modAppointments.AppointmentLabel ("Ivan Petrov - Haircut, Mon 21.09.2026 09:00")
Function AppointmentLabel(idValue)
    Dim table, i, dateText, startText, label, dateValue

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    For i = 1 To table.RowCount
        If CLng(NumberValue(table.TextAt(i, "id"))) = CLng(NumberValue(idValue)) Then
            dateText = table.TextAt(i, "date")
            startText = table.TextAt(i, "start")

            label = ClientName(table.TextAt(i, "client_id")) & " - " & _
                    ServiceName(table.TextAt(i, "service_id"))

            If IsIsoDate(dateText) Then
                dateValue = DateSerial(CLng(Left(dateText, 4)), CLng(Mid(dateText, 6, 2)), _
                                       CLng(Mid(dateText, 9, 2)))
                label = label & ", " & WeekdayShortName(Weekday(dateValue, 2)) & _
                        " " & DateFormatText(dateText)
            End If

            If Len(startText) > 0 Then label = label & " " & startText

            AppointmentLabel = label
            Exit Function
        End If
    Next
End Function

' modDashboard.FreeSlotsFor for the day view: the working day in slots minus the
' appointments of the day (one appointment counts as one slot, whatever it lasts
' - the same rule as the VBA twin).
Function FreeSlotsOn(isoDate)
    Dim slotsPerDay, dayStart, dayEnd, stepValue

    dayStart = DayStartMinutes()
    dayEnd = DayEndMinutes()
    stepValue = StepMinutes()
    If dayEnd <= dayStart Then Exit Function

    slotsPerDay = (dayEnd - dayStart) \ stepValue
    FreeSlotsOn = slotsPerDay - AppointmentsOn(isoDate)
    If FreeSlotsOn < 0 Then FreeSlotsOn = 0
End Function

' ---- the decision layer of the daily grid -----------------------------

' modDashboard.RenderDaily + WriteDailySlot, without Excel: one text per slot row
' - "HH:MM|free", "HH:MM|tail|id" or "HH:MM|start|id|client|service". The Excel
' renderer can dump the same texts from the same fixtures and the harness then
' compares the two (the cross check of tools\sim\README.md).
Function DailyPlan(isoDate)
    Dim table, rows, list, i, slotMinutes, step, plan(), count
    Dim slotText, ownerRow, ownerStart, rowStart, rowDuration, dataRow

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    rows = AppointmentRowsOn(isoDate)
    step = StepMinutes()
    If step < 1 Then step = DEFAULT_STEP_MINUTES
    If Len(rows) > 0 Then list = Split(rows, ",")

    count = 0
    ReDim plan(((DayEndMinutes() - DayStartMinutes()) \ step) - 1)

    For slotMinutes = DayStartMinutes() To DayEndMinutes() - step Step step
        slotText = MinutesToTime(slotMinutes)
        ownerRow = 0
        ownerStart = -1

        ' 1) the appointment that starts in this slot
        If Len(rows) > 0 Then
            For i = 0 To UBound(list)
                dataRow = CLng(list(i))
                rowStart = TimeToMinutes(table.TextAt(dataRow, "start"))
                If rowStart >= slotMinutes And rowStart < slotMinutes + step Then
                    ownerRow = dataRow
                    ownerStart = rowStart
                    Exit For
                End If
            Next
        End If

        ' 2) otherwise the first appointment that reaches into it (a continuation)
        If ownerRow = 0 Then
            If Len(rows) > 0 Then
                For i = 0 To UBound(list)
                    dataRow = CLng(list(i))
                    rowStart = TimeToMinutes(table.TextAt(dataRow, "start"))
                    rowDuration = CLng(NumberValue(table.TextAt(dataRow, "duration")))
                    If rowStart >= 0 Then
                        If Overlaps(slotMinutes, step, rowStart, rowDuration) Then
                            ownerRow = dataRow
                            ownerStart = rowStart
                            Exit For
                        End If
                    End If
                Next
            End If
        End If

        If ownerRow = 0 Then
            plan(count) = slotText & "|free"
        ElseIf ownerStart < slotMinutes Then
            plan(count) = slotText & "|tail|" & table.TextAt(ownerRow, "id")
        Else
            plan(count) = slotText & "|start|" & table.TextAt(ownerRow, "id") & "|" & _
                          ClientName(table.TextAt(ownerRow, "client_id")) & "|" & _
                          ServiceName(table.TextAt(ownerRow, "service_id"))
        End If

        count = count + 1
    Next

    DailyPlan = plan
End Function

' =====================================================================
'  reports (modReports, phase 5)
'
'  A report is a period (from..to, both included) and three groupings of the
'  appointments inside it: by day, by service and by client. The rows are what the
'  report grid of the template shows ("grid_reports_first".."grid_reports_last",
'  group|count|revenue), so the model decides the report and the renderer only
'  writes it - exactly like the daily grid (§6.9).
'
'  Money crosses the lane as cents (MoneyCents): a Double written by CStr() would
'  follow the regional settings of the machine, and the two lanes of the cross
'  check would differ on the decimal separator alone (§13).
' =====================================================================

Const NAME_REPORT_FROM = "i_report_from"
Const NAME_REPORT_TO = "i_report_to"
Const NAME_REPORT_TITLE = "o_report_title"
Const NAME_REPORT_GRID_FIRST = "grid_reports_first"
Const NAME_REPORT_GRID_LAST = "grid_reports_last"

Const REPORT_HEADING_DAYS = "Days"
Const REPORT_HEADING_SERVICES = "Services"
Const REPORT_HEADING_CLIENTS = "Clients"
Const REPORT_LINE_TOTAL = "Total"

Dim gReportLines, gReportLineCount, gReportCapped

' modData.IsoDateText for a real date value: "2026-09-21" (the workbook formats a
' date the same way, so both lanes answer with the same text).
Function DateValueIso(dateValue)
    DateValueIso = CStr(Year(dateValue)) & "-" & Pad2(Month(dateValue)) & "-" & _
                   Pad2(Day(dateValue))
End Function

Function TodayIso()
    TodayIso = DateValueIso(Date)
End Function

Function MonthFirstIso()
    MonthFirstIso = DateValueIso(DateSerial(Year(Date), Month(Date), 1))
End Function

' DateSerial(day 0 of the next month) is the last day of this one - the same trick
' the VBA twin uses.
Function MonthLastIso()
    MonthLastIso = DateValueIso(DateSerial(Year(Date), Month(Date) + 1, 0))
End Function

' modReports.ReportFrom / ReportTo: the period of the report. The cells are what
' the user sees and what the file keeps (§13); empty cells mean the current month
' (§16.1) and the renderer writes those defaults into the cells.
Function ReportFromIso()
    ReportFromIso = ReportPeriodIso(NAME_REPORT_FROM, MonthFirstIso())
End Function

Function ReportToIso()
    ReportToIso = ReportPeriodIso(NAME_REPORT_TO, MonthLastIso())
End Function

Function ReportPeriodIso(nameText, fallbackIso)
    Dim text

    ' An empty cell means the current month (§16.1); a cell that holds something
    ' else goes through as it is ("" when it is not a date), so the rule refuses it
    ' and the message names the end of the period.
    text = Trim(gBook.ReadName(nameText))
    If Len(text) = 0 Then
        ReportPeriodIso = fallbackIso
    Else
        ReportPeriodIso = DateTextToIso(text)
    End If
End Function

' modReports.PeriodRuleCode: what is wrong with the period, as a rule code
' ("" = it is fine). The twin of modAppointments.RuleCode, so the wording stays in
' modText and the two lanes compare codes.
Function PeriodRuleCode(fromIso, toIso)
    If Not IsIsoDate(fromIso) Then
        PeriodRuleCode = "error.required|period start"
        Exit Function
    End If
    If Not IsIsoDate(toIso) Then
        PeriodRuleCode = "error.required|period end"
        Exit Function
    End If
    If fromIso > toIso Then
        PeriodRuleCode = "error.period|" & DateFormatText(fromIso) & "|" & DateFormatText(toIso)
    End If
End Function

' modReports.ShiftPeriod: the same window, moved by whole months (the < and >
' buttons). DateAdd clamps a day the target month does not have (31.03 -> 28.02).
Function ShiftPeriodIso(isoText, months)
    Dim moved

    If Not IsIsoDate(isoText) Then Exit Function
    If months = 0 Then
        ShiftPeriodIso = isoText
        Exit Function
    End If

    moved = DateAdd("m", months, DateSerial(CLng(Left(isoText, 4)), _
                                            CLng(Mid(isoText, 6, 2)), _
                                            CLng(Mid(isoText, 9, 2))))
    ShiftPeriodIso = CStr(Year(moved)) & "-" & Pad2(Month(moved)) & "-" & Pad2(Day(moved))
End Function

' modData.AppointmentsInPeriod / modReports.BookingsInPeriod: how many
' appointments the period holds (ISO text, both ends included).
Function AppointmentsInPeriod(fromIso, toIso)
    Dim table, i, isoText, count

    count = 0
    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    For i = 1 To table.RowCount
        isoText = IsoDateText(table.TextAt(i, "date"))
        If StrComp(isoText, fromIso, 1) >= 0 Then
            If StrComp(isoText, toIso, 1) <= 0 Then count = count + 1
        End If
    Next

    ' CLng, because an untouched counter is Empty and CStr(Empty) is "" - the two
    ' lanes write "0" for an empty period.
    AppointmentsInPeriod = CLng(count)
End Function

Function BookingsInPeriod(fromIso, toIso)
    BookingsInPeriod = AppointmentsInPeriod(fromIso, toIso)
End Function

' The rows the report grid of the template holds (50 in the shipped template).
Function ReportGridCapacity()
    ReportGridCapacity = gBook.RangeOfName(NAME_REPORT_GRID_LAST).Row - _
                         gBook.RangeOfName(NAME_REPORT_GRID_FIRST).Row + 1
End Function

' One row of the report: the cap of §6.9 applies to the report as well - what does
' not fit is dropped and the title says so (ReportIsCapped).
Sub ReportLine(lineText)
    If gReportLineCount >= ReportGridCapacity() Then
        gReportCapped = True
        Exit Sub
    End If

    If Len(gReportLines) > 0 Then gReportLines = gReportLines & vbCrLf
    gReportLines = gReportLines & lineText
    gReportLineCount = gReportLineCount + 1
End Sub

Function ReportIsCapped()
    ReportIsCapped = gReportCapped
End Function

Function ReportLineCount()
    ReportLineCount = gReportLineCount
End Function

' modReports.ReportRows: the report itself, one line per row of the grid
' ("group|count|cents"): the days of the period, then the services, then the
' clients, then the total. Days go up, services and clients go down by revenue and
' by name when two of them are worth the same - a fixed order, because the cross
' check compares row by row.
Function ReportRows(fromIso, toIso)
    gReportLines = ""
    gReportLineCount = 0
    gReportCapped = False

    ReportLine REPORT_HEADING_DAYS
    ReportGroupRows "date", fromIso, toIso
    ReportLine ""
    ReportLine REPORT_HEADING_SERVICES
    ReportGroupRows "service_id", fromIso, toIso
    ReportLine ""
    ReportLine REPORT_HEADING_CLIENTS
    ReportGroupRows "client_id", fromIso, toIso
    ReportLine REPORT_LINE_TOTAL & "|" & CStr(BookingsInPeriod(fromIso, toIso)) & "|" & _
               CStr(MoneyCents(RevenueInPeriod(fromIso, toIso)))

    ReportRows = gReportLines
End Function

' One section of the report: the appointments of the period, grouped by the column
' asked for, as "group|count|cents" lines.
Sub ReportGroupRows(groupColumn, fromIso, toIso)
    Dim table, i, keys, keyText, isoText, isDateMode
    Dim counters, money, keyList, sortList, numberList
    Dim j, swapped, swapText, swapNumber, swapSort

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    Set counters = CreateObject("Scripting.Dictionary")
    Set money = CreateObject("Scripting.Dictionary")
    isDateMode = (groupColumn = "date")
    keys = ""

    For i = 1 To table.RowCount
        isoText = IsoDateText(table.TextAt(i, "date"))
        If StrComp(isoText, fromIso, 1) >= 0 Then
            If StrComp(isoText, toIso, 1) <= 0 Then
                If isDateMode Then
                    keyText = DateFormatText(isoText)
                ElseIf groupColumn = "service_id" Then
                    keyText = ServiceName(table.TextAt(i, groupColumn))
                Else
                    keyText = ClientName(table.TextAt(i, groupColumn))
                End If

                If Len(keyText) > 0 Then
                    If counters.Exists(keyText) Then
                        counters(keyText) = counters(keyText) + 1
                        money(keyText) = money(keyText) + _
                                         MoneyCents(ServicePrice(table.TextAt(i, "service_id")))
                    Else
                        counters(keyText) = 1
                        money(keyText) = MoneyCents(ServicePrice(table.TextAt(i, "service_id")))
                        If Len(keys) > 0 Then keys = keys & vbCrLf
                        keys = keys & keyText
                    End If
                End If
            End If
        End If
    Next

    If Len(keys) = 0 Then Exit Sub

    ' The sort keys: the ISO date of a day (so the order is chronological), the
    ' name otherwise. A plain bubble sort - a report holds a few dozen groups and
    ' VBScript has no sort.
    keyList = Split(keys, vbCrLf)
    ReDim sortList(UBound(keyList))
    ReDim numberList(UBound(keyList))
    For i = 0 To UBound(keyList)
        If isDateMode Then
            sortList(i) = DateTextToIso(keyList(i))
        Else
            sortList(i) = keyList(i)
        End If
        numberList(i) = CLng(money(keyList(i)))
    Next

    For i = 0 To UBound(keyList) - 1
        For j = 0 To UBound(keyList) - 1 - i
            swapped = False
            If isDateMode Then
                If sortList(j) > sortList(j + 1) Then swapped = True
            ElseIf numberList(j) < numberList(j + 1) Then
                swapped = True
            ElseIf numberList(j) = numberList(j + 1) Then
                If sortList(j) > sortList(j + 1) Then swapped = True
            End If

            If swapped Then
                swapText = keyList(j): keyList(j) = keyList(j + 1): keyList(j + 1) = swapText
                swapSort = sortList(j): sortList(j) = sortList(j + 1): sortList(j + 1) = swapSort
                swapNumber = numberList(j): numberList(j) = numberList(j + 1): numberList(j + 1) = swapNumber
            End If
        Next
    Next

    For i = 0 To UBound(keyList)
        ReportLine keyList(i) & "|" & CStr(counters(keyList(i))) & "|" & CStr(numberList(i))
    Next
End Sub

' =====================================================================
'  the Config page (modUiConfig, phase 6)
'
'  Eight settings drive the whole application. The rules and the canonical form
'  (times as 09:00, whole numbers, dates as ISO, lists trimmed) are the twin of
'  modUiConfig, so the two lanes answer with the same codes and the same values.
' =====================================================================

' modUiConfig.Read: the eight settings, keyed by the name of their cell.
Function ConfigValues()
    Dim values

    Set values = CreateObject("Scripting.Dictionary")
    values(NAME_CONFIG_DAY_START) = Trim(gBook.ReadName(NAME_CONFIG_DAY_START))
    values(NAME_CONFIG_DAY_END) = Trim(gBook.ReadName(NAME_CONFIG_DAY_END))
    values(NAME_CONFIG_STEP) = Trim(gBook.ReadName(NAME_CONFIG_STEP))
    values(NAME_CONFIG_DURATION) = Trim(gBook.ReadName(NAME_CONFIG_DURATION))
    values(NAME_CONFIG_CURRENCY) = Trim(gBook.ReadName(NAME_CONFIG_CURRENCY))
    values(NAME_CONFIG_NUMBER_FORMAT) = Trim(gBook.ReadName(NAME_CONFIG_NUMBER_FORMAT))
    values(NAME_CONFIG_NONWORK_DAYS) = Trim(gBook.ReadName(NAME_CONFIG_NONWORK_DAYS))
    values(NAME_CONFIG_NONWORK_DATES) = Trim(gBook.ReadName(NAME_CONFIG_NONWORK_DATES))

    Set ConfigValues = values
End Function

' modUiConfig.DefaultValues: the built in values (§16.1).
Function ConfigDefaultValues()
    Dim values

    Set values = CreateObject("Scripting.Dictionary")
    values(NAME_CONFIG_DAY_START) = DEFAULT_DAY_START
    values(NAME_CONFIG_DAY_END) = DEFAULT_DAY_END
    values(NAME_CONFIG_STEP) = CStr(DEFAULT_STEP_MINUTES)
    values(NAME_CONFIG_DURATION) = CStr(DEFAULT_DURATION_MINUTES)
    values(NAME_CONFIG_CURRENCY) = DEFAULT_CURRENCY
    values(NAME_CONFIG_NUMBER_FORMAT) = DEFAULT_NUMBER_FORMAT
    values(NAME_CONFIG_NONWORK_DAYS) = DEFAULT_NONWORK_WEEKDAYS
    values(NAME_CONFIG_NONWORK_DATES) = ""

    Set ConfigDefaultValues = values
End Function

' A list of comma separated items, trimmed, in order and without repeats
' ("7, 6, 7" -> "6,7").
Function JoinedList(text)
    JoinedList = AddSorted(text, "")
End Function

' The same list as ISO dates ("25.09.2026, 01.10.2026" -> "2026-09-25,2026-10-01"):
' the rule that reads it compares ISO texts, whatever the cell shows.
Function JoinedIsoDates(text)
    JoinedIsoDates = AppendIsoDate(text, "")
End Function

' modUiConfig.AppendIsoDate: the list of the non working dates with one day more -
' ISO, in the order of the calendar and without repeats ("nonsense" changes nothing,
' a day that is already there changes nothing either).
Function AppendIsoDate(listText, isoText)
    Dim parts, i, result

    parts = Split(CStr(listText), ",")
    result = ""
    For i = 0 To UBound(parts)
        result = AddSorted(result, DateTextToIso(Trim(parts(i))))
    Next
    AppendIsoDate = AddSorted(result, DateTextToIso(CStr(isoText)))
End Function

' One item into a comma separated list: in order, without repeats. ISO sorts by
' itself, so the weekdays and the dates use the same helper.
Function AddSorted(listText, itemText)
    Dim parts, items, count, i

    parts = Split(CStr(listText), ",")
    count = 0
    For i = 0 To UBound(parts)
        count = AddOne(items, count, Trim(parts(i)))
    Next
    count = AddOne(items, count, Trim(CStr(itemText)))

    AddSorted = ""
    For i = 0 To count - 1
        If i > 0 Then AddSorted = AddSorted & ","
        AddSorted = AddSorted & items(i)
    Next
End Function

' One item into a sorted array without repeats; the answer is the new count.
Function AddOne(items, count, itemText)
    Dim i, at

    AddOne = count
    If Len(itemText) = 0 Then Exit Function

    at = count
    For i = 0 To count - 1
        If StrComp(items(i), itemText, vbBinaryCompare) = 0 Then Exit Function
        If StrComp(items(i), itemText, vbBinaryCompare) > 0 Then
            at = i
            Exit For
        End If
    Next

    If count = 0 Then
        ReDim items(0)          ' an empty Variant only becomes an array without Preserve
    Else
        ReDim Preserve items(count)
    End If
    For i = count To at + 1 Step -1
        items(i) = items(i - 1)
    Next
    items(at) = itemText
    AddOne = count + 1
End Function

' modUiConfig.RuleCode: what is wrong with the settings, as a rule code
' ("" = they are fine). Only the first problem is reported: the user fixes one
' thing at a time and the message names it.
Function ConfigRuleCode(values)
    Dim text, parts, i, item, startMinutes, endMinutes

    ' the working day: two times, opening before closing
    text = PayloadText(values, NAME_CONFIG_DAY_START)
    If Len(text) = 0 Then
        ConfigRuleCode = "error.required|opening time"
        Exit Function
    End If
    startMinutes = TimeToMinutes(text)
    If startMinutes < 0 Then
        ConfigRuleCode = "error.time|opening time"
        Exit Function
    End If

    text = PayloadText(values, NAME_CONFIG_DAY_END)
    If Len(text) = 0 Then
        ConfigRuleCode = "error.required|closing time"
        Exit Function
    End If
    endMinutes = TimeToMinutes(text)
    If endMinutes < 0 Then
        ConfigRuleCode = "error.time|closing time"
        Exit Function
    End If
    If endMinutes <= startMinutes Then
        ConfigRuleCode = "error.day_order|" & MinutesToTime(startMinutes) & "|" & _
                         MinutesToTime(endMinutes)
        Exit Function
    End If

    ' the resolution of the grids and the duration of a new booking
    If Not IsDigits(PayloadText(values, NAME_CONFIG_STEP)) Then
        ConfigRuleCode = "error.required|slot step"
        Exit Function
    End If
    If CLng(NumberValue(PayloadText(values, NAME_CONFIG_STEP))) < 1 Then
        ConfigRuleCode = "error.required|slot step"
        Exit Function
    End If
    If Not IsDigits(PayloadText(values, NAME_CONFIG_DURATION)) Then
        ConfigRuleCode = "error.required|default duration"
        Exit Function
    End If
    If CLng(NumberValue(PayloadText(values, NAME_CONFIG_DURATION))) < 1 Then
        ConfigRuleCode = "error.required|default duration"
        Exit Function
    End If

    ' the money
    If Len(PayloadText(values, NAME_CONFIG_CURRENCY)) = 0 Then
        ConfigRuleCode = "error.required|currency"
        Exit Function
    End If
    text = PayloadText(values, NAME_CONFIG_NUMBER_FORMAT)
    If InStr(text, "0") = 0 Then
        If InStr(text, "#") = 0 Then
            ConfigRuleCode = "error.format|" & text
            Exit Function
        End If
    End If

    ' the non working weekdays: numbers from 1 (Monday) to 7 (Sunday)
    text = PayloadText(values, NAME_CONFIG_NONWORK_DAYS)
    If Len(text) > 0 Then
        parts = Split(text, ",")
        For i = 0 To UBound(parts)
            item = Trim(parts(i))
            If Len(item) > 0 Then
                If Not IsDigits(item) Then
                    ConfigRuleCode = "error.weekday|" & item
                    Exit Function
                End If
                If CLng(NumberValue(item)) < 1 Or CLng(NumberValue(item)) > 7 Then
                    ConfigRuleCode = "error.weekday|" & item
                    Exit Function
                End If
            End If
        Next
    End If

    ' the non working dates: real dates, ISO or European
    text = PayloadText(values, NAME_CONFIG_NONWORK_DATES)
    If Len(text) > 0 Then
        parts = Split(text, ",")
        For i = 0 To UBound(parts)
            item = Trim(parts(i))
            If Len(item) > 0 Then
                If Len(DateTextToIso(item)) = 0 Then
                    ConfigRuleCode = "error.dates|" & item
                    Exit Function
                End If
            End If
        Next
    End If
End Function

' modUiConfig.Normalise: the settings in the form the application stores (the
' caller writes them back into the cells, so the user sees what was stored).
Sub ConfigNormalise(values)
    values(NAME_CONFIG_DAY_START) = MinutesToTime(TimeToMinutes(PayloadText(values, NAME_CONFIG_DAY_START)))
    values(NAME_CONFIG_DAY_END) = MinutesToTime(TimeToMinutes(PayloadText(values, NAME_CONFIG_DAY_END)))
    values(NAME_CONFIG_STEP) = CStr(CLng(NumberValue(PayloadText(values, NAME_CONFIG_STEP))))
    values(NAME_CONFIG_DURATION) = CStr(CLng(NumberValue(PayloadText(values, NAME_CONFIG_DURATION))))
    values(NAME_CONFIG_CURRENCY) = UCase(PayloadText(values, NAME_CONFIG_CURRENCY))
    values(NAME_CONFIG_NONWORK_DAYS) = JoinedList(PayloadText(values, NAME_CONFIG_NONWORK_DAYS))
    values(NAME_CONFIG_NONWORK_DATES) = JoinedIsoDates(PayloadText(values, NAME_CONFIG_NONWORK_DATES))
End Sub

' =====================================================================
'  mutations (the writing half of modData + the actions of modAppointments)
'
'  The model writes into its tables itself: the flow "form -> rules -> row ->
'  grid" can be followed case by case without a workbook. What it does NOT do is
'  what only Excel can: the backup copy, the mutation rollback, the sheet
'  protection and the status cells (??6.6, ??8.2.4) - those stay in the VBA twin.
' =====================================================================

' modData.NextId: the highest id of a table plus one.
Function NextId(table)
    Dim i, highest

    For i = 1 To table.RowCount
        If NumberValue(table.TextAt(i, "id")) > highest Then
            highest = NumberValue(table.TextAt(i, "id"))
        End If
    Next

    NextId = CLng(highest) + 1
End Function

' The data row of an id (0 when there is none). The VBA twin walks the table the
' same way in AppointmentCell, UpdateAppointment and the deletes.
Function RowIndexById(table, idValue)
    Dim i

    RowIndexById = 0
    For i = 1 To table.RowCount
        If NumberValue(table.TextAt(i, "id")) = NumberValue(idValue) Then
            RowIndexById = i
            Exit Function
        End If
    Next
End Function

' modData.AppendAppointment. The column names of the db_ table ARE the keys of
' the payload (client_id, service_id, date, start, duration), so the row is built
' by walking the header - exactly what SqlMutation.Insert does in Excel.
Function AppendAppointment(values)
    Dim table, fields, i, headerName, newId

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    newId = NextId(table)

    ReDim fields(table.ColumnCount() - 1)
    For i = 1 To table.ColumnCount
        headerName = table.HeaderName(i)
        If LCase(headerName) = "id" Then
            fields(i - 1) = CStr(newId)
        Else
            fields(i - 1) = PayloadText(values, headerName)
        End If
    Next
    table.AddRow fields

    AppendAppointment = newId
End Function

' modData.UpdateAppointment: writes the payload into the row of that id and
' answers how many rows it touched (0 when the id is gone).
Function UpdateAppointment(idValue, values)
    Dim table, i, rowIndex

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    rowIndex = RowIndexById(table, idValue)
    If rowIndex = 0 Then Exit Function

    For i = 1 To table.ColumnCount
        If LCase(table.HeaderName(i)) <> "id" Then
            table.SetCell rowIndex, table.HeaderName(i), PayloadText(values, table.HeaderName(i))
        End If
    Next

    UpdateAppointment = 1
End Function

' modData.DeleteRowById: the rows it deleted (the VBA twin rolls the mutation
' back unless exactly one row matched - the caller here gets the number).
Function DeleteRowById(sheetName, idValue)
    Dim table, rowIndex

    Set table = ModelTable(sheetName)
    rowIndex = RowIndexById(table, idValue)
    If rowIndex = 0 Then Exit Function

    DeleteRowById = table.RemoveRow(rowIndex)
End Function

' modData.CountAppointmentsFor: how many appointments point at a client/service
' (the guard of the two admin sheets).
Function CountAppointmentsFor(columnName, idValue)
    Dim table, i, count

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    For i = 1 To table.RowCount
        If NumberValue(table.TextAt(i, columnName)) = NumberValue(idValue) Then count = count + 1
    Next

    CountAppointmentsFor = count
End Function

' modData.AppointmentCell / AppointmentIdAt
Function AppointmentCell(idValue, columnName)
    Dim table, rowIndex

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    rowIndex = RowIndexById(table, idValue)
    If rowIndex = 0 Then Exit Function

    AppointmentCell = table.TextAt(rowIndex, columnName)
End Function

Function AppointmentIdAt(isoDate, startText)
    Dim table, i

    Set table = ModelTable(SHEET_DB_APPOINTMENTS)
    For i = 1 To table.RowCount
        If StrComp(table.TextAt(i, "date"), Trim(isoDate), 1) = 0 Then
            If StrComp(table.TextAt(i, "start"), Trim(startText), 1) = 0 Then
                AppointmentIdAt = CLng(NumberValue(table.TextAt(i, "id")))
                Exit Function
            End If
        End If
    Next
End Function

' modData.ClientIdByName / ServiceIdByName ("first match wins", 0 when none of
' that name exists - two clients may share a name, see phase 3).
Function ClientIdByName(nameText)
    Dim table, i, wanted

    wanted = Trim(nameText)
    If Len(wanted) = 0 Then Exit Function

    Set table = ModelTable(SHEET_DB_CLIENTS)
    For i = 1 To table.RowCount
        If StrComp(table.TextAt(i, "name"), wanted, 1) = 0 Then
            ClientIdByName = CLng(NumberValue(table.TextAt(i, "id")))
            Exit Function
        End If
    Next
End Function

Function ServiceIdByName(nameText)
    Dim table, i, wanted

    wanted = Trim(nameText)
    If Len(wanted) = 0 Then Exit Function

    Set table = ModelTable(SHEET_DB_SERVICES)
    For i = 1 To table.RowCount
        If StrComp(table.TextAt(i, "service"), wanted, 1) = 0 Then
            ServiceIdByName = CLng(NumberValue(table.TextAt(i, "id")))
            Exit Function
        End If
    Next
End Function

' =====================================================================
'  the form and its actions (modUiBooking.ReadForm, the domain part of
'  modAppointments.SaveBooking / DeleteBooking)
' =====================================================================

' "45" out of the two cells of a duration field; "" while one of them is empty
' (modUiBooking.ReadForm stores "" then and the rules ask for the duration).
Function DurationFieldText(book, hoursName, minutesName)
    Dim hoursText, minutesText

    hoursText = Trim(book.ReadName(hoursName))
    minutesText = Trim(book.ReadName(minutesName))
    If Not IsDigits(hoursText) Then Exit Function
    If Not IsDigits(minutesText) Then Exit Function

    DurationFieldText = CStr(CLng(hoursText) * 60 + CLng(minutesText))
End Function

' modUiBooking.ReadForm: the payload the filled form means. The client and the
' service are looked up by the name that stands in the dropdown, the date is what
' the user typed (ISO or European), the start comes from its two cells and the
' duration in minutes from its two cells; s_edit_id makes it an edit.
Function ValuesFromForm(book)
    Dim values, durationText

    Set values = CreateObject("Scripting.Dictionary")
    ' The form normalises the date it shows (modUiBooking.FormDate): every payload
    ' leaves the form with an ISO date, so the validator never reads a European one.
    values("date") = DateTextToIso(book.ReadName("i_booking_date"))
    values("client_id") = CStr(ClientIdByName(book.ReadName("i_booking_client")))
    values("service_id") = CStr(ServiceIdByName(book.ReadName("i_booking_service")))
    values("start") = Trim(book.ReadName("i_booking_start_h")) & ":" & _
                      Trim(book.ReadName("i_booking_start_m"))

    durationText = DurationFieldText(book, "i_booking_dur_h", "i_booking_dur_m")
    values("duration") = durationText

    If Len(Trim(book.ReadName("s_edit_id"))) > 0 Then
        values("id") = Trim(book.ReadName("s_edit_id"))
    End If

    Set ValuesFromForm = values
End Function

' The domain part of modAppointments.SaveBooking: the rules decide, then the row
' is written (updated when the payload carries an id, appended otherwise). The
' answer is a dictionary: ok, id and - when a rule refused it - rule.
Function SaveBookingResult(values)
    Dim result, idValue, ruleText, newId

    Set result = CreateObject("Scripting.Dictionary")
    result("ok") = False
    result("id") = 0

    idValue = CLng(NumberValue(PayloadText(values, "id")))
    If idValue > 0 Then
        If Not RowExists(SHEET_DB_APPOINTMENTS, idValue) Then
            result("rule") = "error.not_found|Appointment|" & CStr(idValue)
            Set SaveBookingResult = result
            Exit Function
        End If
    End If

    ruleText = RuleCode(values)
    If Len(ruleText) > 0 Then
        result("rule") = ruleText
        Set SaveBookingResult = result
        Exit Function
    End If

    If idValue > 0 Then
        UpdateAppointment idValue, values
        newId = idValue
    Else
        newId = AppendAppointment(values)
    End If

    result("ok") = True
    result("id") = newId
    Set SaveBookingResult = result
End Function

' The domain part of modAppointments.DeleteBooking plus the rule of ??7.5: a form
' that was editing the deleted appointment goes back to "new appointment" (the
' Cancel/Delete button of the weekly view relies on exactly that).
Function DeleteBookingById(book, idValue)
    Dim deleted

    deleted = DeleteRowById(SHEET_DB_APPOINTMENTS, idValue)
    If deleted <> 1 Then Exit Function

    If NumberValue(book.ReadName("s_edit_id")) = NumberValue(idValue) Then
        book.WriteName "s_edit_id", ""
    End If

    DeleteBookingById = True
End Function


