Attribute VB_Name = "modAppointments"
Option Explicit

' =====================================================================
'  Service_Manager - the appointment domain (Phase 3)
'
'  The rules of §8.2 live here and nowhere else: the date must be ISO text on a
'  working day, the start must be a time on the grid step inside the working day,
'  the duration must be positive, the end must still be inside the working day,
'  and the interval [start, start + duration) must not touch another appointment
'  of the same day. Touching intervals (09:00+60 and 10:00) do NOT clash.
'
'  The pure helpers (EndTime, Overlaps, IsOnStep, ...) are used by the pickers
'  and the dashboard as well, so an appointment never has two definitions.
' =====================================================================

' ---- pure helpers -----------------------------------------------------

' "09:30" + 60 -> "10:30"; "" when the start is not a time or the duration is 0.
Public Function EndTime(ByVal startText As String, ByVal durationMinutes As Long) As String
    Dim startMinutes As Long

    startMinutes = modApp.TimeToMinutes(startText)
    If startMinutes < 0 Then Exit Function
    If durationMinutes < 1 Then Exit Function

    EndTime = modApp.MinutesToTime(startMinutes + durationMinutes)
End Function

' True when [startA, startA + durationA) and [startB, startB + durationB) share
' time. Intervals that only touch are not an overlap.
Public Function Overlaps(ByVal startA As Long, ByVal durationA As Long, _
                         ByVal startB As Long, ByVal durationB As Long) As Boolean
    If durationA < 1 Or durationB < 1 Then Exit Function

    Overlaps = (startA < startB + durationB) And (startB < startA + durationA)
End Function

' A start minute that sits on the step of Config. A step below 1 means "no rule".
Public Function IsOnStep(ByVal minutes As Long, ByVal stepMinutes As Long) As Boolean
    If stepMinutes < 1 Then
        IsOnStep = True
        Exit Function
    End If

    IsOnStep = ((minutes Mod stepMinutes) = 0)
End Function

' ---- the working day of Config ---------------------------------------

Public Function DayStartMinutes() As Long
    DayStartMinutes = modApp.TimeToMinutes( _
        modDashboard.ConfigText(NAME_CONFIG_DAY_START, DEFAULT_DAY_START))
    If DayStartMinutes < 0 Then DayStartMinutes = modApp.TimeToMinutes(DEFAULT_DAY_START)
End Function

Public Function DayEndMinutes() As Long
    DayEndMinutes = modApp.TimeToMinutes( _
        modDashboard.ConfigText(NAME_CONFIG_DAY_END, DEFAULT_DAY_END))
    If DayEndMinutes < 0 Then DayEndMinutes = modApp.TimeToMinutes(DEFAULT_DAY_END)
End Function

Public Function StepMinutes() As Long
    StepMinutes = modDashboard.ConfigNumber(NAME_CONFIG_STEP, DEFAULT_STEP_MINUTES)
    If StepMinutes < 1 Then StepMinutes = DEFAULT_STEP_MINUTES
End Function

' True when the date is a working day (§8.2.12): not in the list of weekdays
' (1 = Monday) and not in the list of dates of Config. Text that is not an ISO
' date is not a working day either.
Public Function IsWorkingDay(ByVal isoDate As String) As Boolean
    Dim dateValue As Date
    Dim weekdayList As String
    Dim dateList As String

    If Not modPickers.IsIsoDate(isoDate) Then Exit Function

    dateValue = modDashboard.DateOfText(isoDate)

        ' The cell is read as it is: an empty list means no non working weekday at
        ' all, which is a setting the user can make (Apply stores the empty cell).
        ' A fallback here would turn "I work on Saturdays" back into "6,7".
        weekdayList = Replace(Trim$(CStr(modNames.ReadName(NAME_CONFIG_NONWORK_DAYS))), " ", "")
    If Len(weekdayList) > 0 Then
        If InStr(1, "," & weekdayList & ",", "," & CStr(Weekday(dateValue, vbMonday)) & ",", _
                 vbTextCompare) > 0 Then Exit Function
    End If

    dateList = Replace(modDashboard.ConfigText(NAME_CONFIG_NONWORK_DATES, ""), " ", "")
    If Len(dateList) = 0 Then
        IsWorkingDay = True
        Exit Function
    End If

    IsWorkingDay = (InStr(1, "," & dateList & ",", "," & Trim$(isoDate) & ",", _
                          vbTextCompare) = 0)
End Function

' ---- the other appointments of a day ---------------------------------

' How many appointments of that day overlap the interval; the row that is being
' edited (exceptId) never counts itself.
Public Function ConflictCount(ByVal isoDate As String, ByVal startMinutes As Long, _
                              ByVal durationMinutes As Long, ByVal exceptId As Long) As Long
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String
    Dim rowStart As Long
    Dim rowDuration As Long
    Dim rowId As Long
    Dim count As Long

    wanted = Trim$(isoDate)
    If Len(wanted) < ISO_DATE_LENGTH Then Exit Function

    Set table = modData.AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        If StrComp(modData.DateTextOfRow(table, rowIndex), wanted, vbTextCompare) = 0 Then
            rowId = CLng(Val(CStr(modData.IdOfRow(table, rowIndex))))
            If rowId <> exceptId Then
                rowStart = modApp.TimeToMinutes(modData.StartTextOfRow(table, rowIndex))
                rowDuration = modData.DurationOfRow(table, rowIndex)
                If rowStart >= 0 Then
                    If Overlaps(startMinutes, durationMinutes, rowStart, rowDuration) Then
                        count = count + 1
                    End If
                End If
            End If
        End If
    Next rowIndex

    ConflictCount = count
End Function

' The start of the first appointment of the day that is in the way ("" when there
' is none) - used by the message and by the picker.
Public Function FirstConflictStart(ByVal isoDate As String, ByVal startMinutes As Long, _
                                   ByVal durationMinutes As Long, ByVal exceptId As Long) As String
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String
    Dim rowStart As Long
    Dim rowDuration As Long
    Dim rowId As Long
    Dim best As Long

    wanted = Trim$(isoDate)
    If Len(wanted) < ISO_DATE_LENGTH Then Exit Function

    Set table = modData.AppointmentsTable()
    best = -1
    For rowIndex = 1 To table.RowCount
        If StrComp(modData.DateTextOfRow(table, rowIndex), wanted, vbTextCompare) = 0 Then
            rowId = CLng(Val(CStr(modData.IdOfRow(table, rowIndex))))
            If rowId <> exceptId Then
                rowStart = modApp.TimeToMinutes(modData.StartTextOfRow(table, rowIndex))
                rowDuration = modData.DurationOfRow(table, rowIndex)
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
        End If
    Next rowIndex

    If best >= 0 Then FirstConflictStart = modApp.MinutesToTime(best)
End Function

' "09:00+30,10:30+90" - the intervals of the appointments of that day, without
' the row that is being edited. The time picker greys the slots that touch them,
' so the check in the picker and the check of the domain use the same intervals.
Public Function BusyIntervalsOn(ByVal isoDate As String, ByVal exceptId As Long) As String
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String
    Dim text As String
    Dim durationValue As Long

    wanted = Trim$(isoDate)
    If Len(wanted) < ISO_DATE_LENGTH Then Exit Function

    Set table = modData.AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        If StrComp(modData.DateTextOfRow(table, rowIndex), wanted, vbTextCompare) = 0 Then
            If CLng(Val(CStr(modData.IdOfRow(table, rowIndex)))) <> exceptId Then
                durationValue = modData.DurationOfRow(table, rowIndex)
                If durationValue < 1 Then durationValue = 1
                If Len(text) > 0 Then text = text & ","
                text = text & modData.StartTextOfRow(table, rowIndex) & "+" & CStr(durationValue)
            End If
        End If
    Next rowIndex

    BusyIntervalsOn = text
End Function

' =====================================================================
'  validation (§8.2.7)
' =====================================================================

' The rule that refuses a booking, as a code with its parameters ("error.step|5")
' and "" when the booking is fine. The code is the identity of the rule: Problems
' turns it into the sentence the user reads, the log and the cross check of
' tools\sim work with the code itself (the script model carries the same function
' under the same name).
' The values carry client_id, service_id, date, start, duration and - when an
' existing row is edited - id. The validator fills in the normalised values
' (ISO date, HH:MM start, numeric duration), so the writer stores exactly what
' was checked.
Public Function RuleCode(ByRef values As Object) As String
    Dim clientId As Long
    Dim serviceId As Long
    Dim dateText As String
    Dim startText As String
    Dim durationRaw As String
    Dim startMinutes As Long
    Dim durationValue As Long
    Dim exceptId As Long
    Dim endMinutes As Long
    Dim clashes As Long

    clientId = CLng(Val(modApp.PayloadText(values, COL_CLIENT_ID)))
    If clientId < 1 Then
        RuleCode = RULE_REQUIRED & "|client"
        Exit Function
    End If
    If Not modData.RowExists(SHEET_DB_CLIENTS, clientId) Then
        RuleCode = RULE_NOT_FOUND & "|Client|" & CStr(clientId)
        Exit Function
    End If

    serviceId = CLng(Val(modApp.PayloadText(values, COL_SERVICE_ID)))
    If serviceId < 1 Then
        RuleCode = RULE_REQUIRED & "|service"
        Exit Function
    End If
    If Not modData.RowExists(SHEET_DB_SERVICES, serviceId) Then
        RuleCode = RULE_NOT_FOUND & "|Service|" & CStr(serviceId)
        Exit Function
    End If

    dateText = modData.IsoDateText(modApp.PayloadText(values, COL_DATE))
    If Len(dateText) < ISO_DATE_LENGTH Then
        RuleCode = RULE_REQUIRED & "|date"
        Exit Function
    End If
    If Not IsWorkingDay(dateText) Then
        RuleCode = RULE_NONWORK & "|" & dateText
        Exit Function
    End If
    values(COL_DATE) = dateText

    startText = modApp.PayloadText(values, COL_START)
    startMinutes = modApp.TimeToMinutes(startText)
    If startMinutes < 0 Then
        RuleCode = RULE_TIME & "|start time"
        Exit Function
    End If
    values(COL_START) = modApp.MinutesToTime(startMinutes)

    durationRaw = modApp.PayloadText(values, COL_SERVICE_DURATION)
    If Len(durationRaw) = 0 Then
        RuleCode = RULE_REQUIRED & "|duration"
        Exit Function
    End If
    If Not IsNumeric(durationRaw) Then
        RuleCode = RULE_REQUIRED & "|duration"
        Exit Function
    End If

    durationValue = CLng(Val(durationRaw))
    If durationValue < 1 Or durationValue > MAX_DURATION_MINUTES Then
        RuleCode = RULE_REQUIRED & "|duration"
        Exit Function
    End If
    values(COL_SERVICE_DURATION) = durationValue

    endMinutes = startMinutes + durationValue
    If startMinutes < DayStartMinutes() Or endMinutes > DayEndMinutes() Then
        RuleCode = RULE_RANGE & "|" & modApp.MinutesToTime(DayStartMinutes()) & _
                   "|" & modApp.MinutesToTime(DayEndMinutes())
        Exit Function
    End If

    If Not IsOnStep(startMinutes, StepMinutes()) Then
        RuleCode = RULE_STEP & "|" & CStr(StepMinutes())
        Exit Function
    End If

    exceptId = modApp.PayloadNumber(values, modRoutes.KEY_ID)
    clashes = ConflictCount(dateText, startMinutes, durationValue, exceptId)
    If clashes > 0 Then
        RuleCode = RULE_CONFLICT & "|" & CStr(clashes) & "|" & _
                   FirstConflictStart(dateText, startMinutes, durationValue, exceptId)
        Exit Function
    End If
End Function

' The problems of a booking as one text for the status cell ("" = fine): the rule
' that refused it becomes the sentence the user reads, so the wording lives in
' modText only and the model of tools\sim does not have to know it. A code without
' a sentence of its own is passed through unchanged.
Public Function Problems(ByRef values As Object) As String
    Dim codeText As String
    Dim parts As Variant

    codeText = RuleCode(values)
    If Len(codeText) = 0 Then Exit Function

    parts = Split(codeText, "|")
    Select Case parts(0)
        Case RULE_REQUIRED
            Problems = modText.TextFormat(modText.KEY_ERROR_REQUIRED, parts(1))
        Case RULE_NOT_FOUND
            Problems = modText.TextFormat(modText.KEY_ERROR_NOT_FOUND, parts(1), parts(2))
        Case RULE_NONWORK
            Problems = modText.TextFormat(modText.KEY_ERROR_NONWORK, parts(1))
        Case RULE_TIME
            Problems = modText.TextFormat(modText.KEY_ERROR_TIME, parts(1))
        Case RULE_RANGE
            Problems = modText.TextFormat(modText.KEY_ERROR_RANGE, parts(1), parts(2))
        Case RULE_STEP
            Problems = modText.TextFormat(modText.KEY_ERROR_STEP, parts(1))
        Case RULE_CONFLICT
            Problems = modText.TextFormat(modText.KEY_ERROR_CONFLICT, parts(1), parts(2))
        Case Else
            Problems = codeText
    End Select
End Function

' =====================================================================
'  the actions of the domain
' =====================================================================

' "10 min" / "1 h 30 min" for the preview cells of Booking (§7.5).
Public Function DurationText(ByVal minutes As Long) As String
    Dim hours As Long
    Dim rest As Long

    If minutes < 1 Then Exit Function

    hours = minutes \ MINUTES_PER_HOUR
    rest = minutes Mod MINUTES_PER_HOUR

    If hours = 0 Then
        DurationText = modText.TextFormat(modText.KEY_TIME_MINUTES, CStr(rest))
    Else
        DurationText = modText.TextFormat(modText.KEY_TIME_HOURS_MINUTES, _
                                          CStr(hours), CStr(rest))
    End If
End Function

' "Ivan Petrov - Haircut, Tue 29.09.2026 10:20": how one appointment reads on the
' screen (the date is European, the data stays ISO); "" when it is gone.
Public Function AppointmentLabel(ByVal idValue As Long) As String
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim dateText As String
    Dim startText As String
    Dim label As String

    Set table = modData.AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        If CLng(Val(CStr(modData.IdOfRow(table, rowIndex)))) = idValue Then
            dateText = modData.DateTextOfRow(table, rowIndex)
            startText = modData.StartTextOfRow(table, rowIndex)

            label = modData.ClientName(modData.ClientIdOfRow(table, rowIndex)) & " - " & _
                    modData.ServiceName(modData.ServiceIdOfRow(table, rowIndex))

            If modPickers.IsIsoDate(dateText) Then
                label = label & ", " & modText.WeekdayShortName(Weekday( _
                        DateSerial(CLng(Left$(dateText, 4)), CLng(Mid$(dateText, 6, 2)), _
                                   CLng(Mid$(dateText, 9, 2))), vbMonday)) & _
                        " " & modPickers.DateFormatText(dateText)
            End If

            If Len(startText) > 0 Then label = label & " " & startText

            AppointmentLabel = label
            Exit Function
        End If
    Next rowIndex
End Function

' The values of an appointment row, keyed by column; Nothing when it is gone.
Public Function RowValues(ByVal idValue As Long) As Object
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim values As Object

    Set RowValues = Nothing
    Set table = modData.AppointmentsTable()

    For rowIndex = 1 To table.RowCount
        If CLng(Val(CStr(modData.IdOfRow(table, rowIndex)))) = idValue Then
            Set values = CreateObject("Scripting.Dictionary")
            values(COL_ID) = CStr(idValue)
            values(COL_CLIENT_ID) = CStr(modData.ClientIdOfRow(table, rowIndex))
            values(COL_SERVICE_ID) = CStr(modData.ServiceIdOfRow(table, rowIndex))
            values(COL_DATE) = modData.DateTextOfRow(table, rowIndex)
            values(COL_START) = modData.StartTextOfRow(table, rowIndex)
            values(COL_SERVICE_DURATION) = CStr(modData.DurationOfRow(table, rowIndex))
            Set RowValues = values
            Exit Function
        End If
    Next rowIndex
End Function

' =====================================================================
'  appointments.save / delete / edit / new
' =====================================================================

' appointments.save: insert, or update when the payload carries an id (§8.2.10).
Public Sub SaveBooking(ByRef data As Object, ByRef result As Object)
    Dim values As Object
    Dim problemText As String
    Dim idValue As Long
    Dim isUpdate As Boolean
    Dim info As Object

    Set values = modApp.PayloadFrom(data)
    idValue = modApp.PayloadNumber(values, modRoutes.KEY_ID)
    isUpdate = (idValue > 0)

    If isUpdate Then
        If Not modData.RowExists(SHEET_DB_APPOINTMENTS, idValue) Then
            result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
                modText.KEY_ERROR_NOT_FOUND, "Appointment", CStr(idValue))
            Exit Sub
        End If
    End If

    problemText = Problems(values)
    If Len(problemText) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = problemText
        Exit Sub
    End If

    If isUpdate Then
        modData.UpdateAppointment idValue, values
    Else
        idValue = modData.AppendAppointment(values)
    End If

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(idValue)
    info(COL_DATE) = CStr(values(COL_DATE))
    info(COL_START) = CStr(values(COL_START))
    info(COL_SERVICE_DURATION) = CStr(values(COL_SERVICE_DURATION))
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = BookingMessage(isUpdate, idValue, values)
    result(modRoutes.RESULT_OK) = True
End Sub

' appointments.delete: confirm, copy, delete (§6.6, §8.2.4).
Public Sub DeleteBooking(ByRef data As Object, ByRef result As Object)
    Dim idValue As Long
    Dim info As Object

    idValue = modApp.PayloadNumber(data, modRoutes.KEY_ID)
    If idValue < 1 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_BAD_ID, modApp.PayloadText(data, modRoutes.KEY_ID))
        Exit Sub
    End If

    If Not modData.RowExists(SHEET_DB_APPOINTMENTS, idValue) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_NOT_FOUND, "Appointment", CStr(idValue))
        Exit Sub
    End If

    If Not Confirmed(data, result, idValue) Then Exit Sub

    If Not modBackup.SaveCopy("appointments-delete-" & CStr(idValue)) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_BACKUP, _
                                                             modBackup.FolderPath())
        Exit Sub
    End If

    modData.DeleteRowById SHEET_DB_APPOINTMENTS, idValue

    ' A form that was editing this appointment cannot edit it any more: the row is
    ' gone, so the form goes back to "new appointment" (§7.5).
    If modUiBooking.EditingId() = idValue Then modUiBooking.ClearForm

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(idValue)
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_BOOKING_DELETED, _
                                                          CStr(idValue))
    result(modRoutes.RESULT_OK) = True
End Sub

' appointments.edit: load the row into the Booking form and show the sheet.
Public Sub EditBooking(ByRef data As Object, ByRef result As Object)
    Dim idValue As Long
    Dim values As Object
    Dim info As Object

    idValue = modApp.PayloadNumber(data, modRoutes.KEY_ID)
    If idValue < 1 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_BAD_ID, modApp.PayloadText(data, modRoutes.KEY_ID))
        Exit Sub
    End If

    Set values = RowValues(idValue)
    If values Is Nothing Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_NOT_FOUND, "Appointment", CStr(idValue))
        Exit Sub
    End If

    modUiBooking.WriteForm values
    modUi.Navigate SHEET_BOOKING          ' a deliberate move (§6.8)
    modApp.RefreshScreens

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(idValue)
    info(COL_DATE) = CStr(values(COL_DATE))
    info(COL_START) = CStr(values(COL_START))
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_BOOKING_EDIT, CStr(idValue))
    result(modRoutes.RESULT_OK) = True
End Sub

' appointments.new: a clean form for a new appointment, with the date and time
' the caller asked for. Without a date the form opens on the day the dashboard
' shows - the user is looking at that day when he clicks a free slot (§6.9).
Public Sub NewBooking(ByRef data As Object, ByRef result As Object)
    Dim dateText As String
    Dim startText As String
    Dim info As Object

    dateText = modData.IsoDateText(modApp.PayloadText(data, COL_DATE))
    If Len(dateText) < ISO_DATE_LENGTH Then dateText = modDashboard.CurrentDate()

    startText = modApp.PayloadText(data, COL_START)
    If modApp.TimeToMinutes(startText) < 0 Then
        startText = modApp.MinutesToTime(DayStartMinutes())
    End If

    modUiBooking.ResetForm dateText, startText
    modUi.Navigate SHEET_BOOKING
    modApp.RefreshScreens

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_DATE) = dateText
    info(COL_START) = startText
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_BOOKING_NEW)
    result(modRoutes.RESULT_OK) = True
End Sub

' =====================================================================
'  internals
' =====================================================================

' The confirmation of §6.6 (see modClients.Confirmed): one modal question.
Private Function Confirmed(ByRef data As Object, ByRef result As Object, _
                           ByVal idValue As Long) As Boolean
    Dim keyText As String

    keyText = "appointments.delete." & CStr(idValue)
    Confirmed = modApp.ConfirmNow(keyText)

    If Not Confirmed Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CANCELLED)
    End If
End Function

Private Function BookingMessage(ByVal isUpdate As Boolean, ByVal idValue As Long, _
                                ByRef values As Object) As String
    Dim label As String

    label = modData.ClientName(values(COL_CLIENT_ID)) & " / " & _
            modData.ServiceName(values(COL_SERVICE_ID))

    If isUpdate Then
        BookingMessage = modText.TextFormat(modText.KEY_BOOKING_UPDATED, CStr(idValue), label, _
                                            CStr(values(COL_DATE)), CStr(values(COL_START)))
    Else
        BookingMessage = modText.TextFormat(modText.KEY_BOOKING_ADDED, CStr(idValue), label, _
                                            CStr(values(COL_DATE)), CStr(values(COL_START)))
    End If
End Function
