Attribute VB_Name = "modUiBooking"
Option Explicit

' =====================================================================
'  Service_Manager - the Booking page (Phase 3, §7.5, §9.4)
'
'  Two identical time fields (hour cell + minute cell + a pick button), a date
'  with a picker, a client and a service from the lists of _Lists, and the
'  previews (client id, price, service duration, start, duration, end, mode).
'
'  app.refresh never touches the input cells - the user may be typing there - it
'  only recomputes the previews (§9.4). The fill of the duration from the service
'  happens when the service cell changes.
' =====================================================================

' =====================================================================
'  reading the form
' =====================================================================

' The values of the form, keyed by column. An unknown client or service name
' gives id 0, which the validator refuses with a clear message.
Public Function ReadForm() As Object
    Dim values As Object
    Dim editText As String

    Set values = CreateObject("Scripting.Dictionary")
    values(COL_DATE) = FormDate()
    values(COL_CLIENT_ID) = CStr(modData.ClientIdByName(modNames.ReadName(NAME_BOOKING_CLIENT)))
    values(COL_SERVICE_ID) = CStr(modData.ServiceIdByName(modNames.ReadName(NAME_BOOKING_SERVICE)))

    values(COL_START) = ReadTimeField(FIELD_START)
    If TimeFieldMinutes(FIELD_DURATION) < 0 Then
        values(COL_SERVICE_DURATION) = vbNullString
    Else
        values(COL_SERVICE_DURATION) = CStr(TimeFieldMinutes(FIELD_DURATION))
    End If

    editText = Trim$(modNames.ReadName(NAME_STATE_EDIT_ID))
    If Len(editText) > 0 Then values(modRoutes.KEY_ID) = editText

    Set ReadForm = values
End Function

' The date of the form as ISO text ("" when the cell is empty or holds something
' that is not a date). The user may type the ISO form or the European one; the
' payload, the data and the rules are always ISO text (§8.2.2).
Private Function FormDate() As String
    FormDate = modPickers.DateTextToIso(modNames.ReadName(NAME_BOOKING_DATE))
End Function

' booking.cancel: the button next to Save. On a new appointment it is "Cancel"
' and empties the form; while an appointment is being edited the same button is
' "Delete" and deletes that appointment (two step confirmation, §6.6) - the
' weekly grid has no delete column any more (17.09.2026, §6.9).
Public Sub OnCancel(ByRef data As Object, ByRef result As Object)
    Dim idValue As Long

    If data Is Nothing Then
        Err.Raise 5, "modUiBooking.OnCancel", "no payload was given."
    End If

    idValue = EditingId()
    If idValue > 0 Then
        ' the delete action works on the payload: the id of the form goes in
        data(modRoutes.KEY_ID) = CStr(idValue)
        modAppointments.DeleteBooking data, result
        Exit Sub
    End If

    ClearForm
    result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CLEARED)
    result(modRoutes.RESULT_OK) = True
End Sub

' form.pick_client: the typed search that fills the client field (§9.5). The id
' follows from the name (RenderPreviews), exactly as it does for a name typed by
' hand - one rule for both ways in, so the dropdown, the search and the keyboard
' cannot drift apart.
Public Sub OnPickClient(ByRef data As Object, ByRef result As Object)
    Dim picked As String
    Dim idValue As Long

    ' A test drives the flow without a dialog: the answer comes in the payload.
    picked = Trim$(modApp.PayloadText(data, modRoutes.KEY_VALUE))
    If Len(picked) = 0 Then
        picked = modPickers.PickClient(modNames.ReadName(NAME_BOOKING_CLIENT))
    End If

    If Len(picked) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CANCELLED)
        Exit Sub
    End If

    modNames.WriteName NAME_BOOKING_CLIENT, picked
    RenderPreviews

    idValue = modData.ClientIdByName(picked)
    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_OK_PICKED_CLIENT, _
                                                          picked, CStr(idValue))
    result(modRoutes.RESULT_OK) = True
End Sub

' "09:30" of one of the two time fields; "" while one of the cells is empty or
' holds something that is not an hour/minute (§9.4).
Public Function ReadTimeField(ByVal fieldKey As String) As String
    Dim hoursText As String
    Dim minutesText As String

    hoursText = modPickers.HourText(modNames.ReadName(HoursName(fieldKey)))
    minutesText = modPickers.MinuteText(modNames.ReadName(MinutesName(fieldKey)))
    If Len(hoursText) = 0 Or Len(minutesText) = 0 Then Exit Function

    ReadTimeField = hoursText & ":" & minutesText
End Function

' The minutes of a time field; -1 while the field is empty or unreadable.
Public Function TimeFieldMinutes(ByVal fieldKey As String) As Long
    Dim text As String

    text = ReadTimeField(fieldKey)
    If Len(text) = 0 Then
        TimeFieldMinutes = -1
        Exit Function
    End If

    TimeFieldMinutes = modApp.TimeToMinutes(text)
End Function

' =====================================================================
'  writing the form
' =====================================================================

' Writes both cells of a time field and refreshes the previews.
Public Sub WriteTimeField(ByVal fieldKey As String, ByVal hours As Long, ByVal minutes As Long)
    modNames.WriteName HoursName(fieldKey), Format$(hours, "00")
    modNames.WriteName MinutesName(fieldKey), Format$(minutes, "00")

    RenderPreviews
End Sub

' Writes a time text ("8:30", "830", "09:30") into a time field; False when the
' text cannot be read.
Public Function ApplyTimeText(ByVal fieldKey As String, ByVal textText As String) As Boolean
    Dim timeText As String
    Dim total As Long

    timeText = modPickers.NormaliseTimeText(textText)
    If Len(timeText) = 0 Then Exit Function

    total = modApp.TimeToMinutes(timeText)
    WriteTimeField fieldKey, total \ MINUTES_PER_HOUR, total Mod MINUTES_PER_HOUR

    ApplyTimeText = True
End Function

' Fills the whole form from a row (appointments.edit).
Public Sub WriteForm(ByRef values As Object)
    Dim startText As String
    Dim total As Long

    modNames.WriteName NAME_BOOKING_DATE, modPickers.DateFormatText(CStr(values(COL_DATE)))
    modNames.WriteName NAME_BOOKING_CLIENT, modData.ClientName(values(COL_CLIENT_ID))
    modNames.WriteName NAME_BOOKING_SERVICE, modData.ServiceName(values(COL_SERVICE_ID))

    startText = modPickers.NormaliseTimeText(CStr(values(COL_START)))
    If Len(startText) > 0 Then
        total = modApp.TimeToMinutes(startText)
        modNames.WriteName NAME_BOOKING_START_H, Format$(total \ MINUTES_PER_HOUR, "00")
        modNames.WriteName NAME_BOOKING_START_M, Format$(total Mod MINUTES_PER_HOUR, "00")
    Else
        ClearTimeField FIELD_START
    End If

    total = CLng(Val(CStr(values(COL_SERVICE_DURATION))))
    If total > 0 Then
        modNames.WriteName NAME_BOOKING_DUR_H, Format$(total \ MINUTES_PER_HOUR, "00")
        modNames.WriteName NAME_BOOKING_DUR_M, Format$(total Mod MINUTES_PER_HOUR, "00")
    Else
        ClearTimeField FIELD_DURATION
    End If

    modNames.WriteName NAME_STATE_EDIT_ID, CStr(values(COL_ID))
    RenderPreviews
End Sub

' A clean form for a new appointment, with the date and the start the caller
' asked for (appointments.new).
Public Sub ResetForm(ByVal dateText As String, ByVal startText As String)
    ClearForm

    modNames.WriteName NAME_BOOKING_DATE, modPickers.DateFormatText(dateText)
    If Len(startText) > 0 Then ApplyTimeText FIELD_START, startText

    RenderPreviews
End Sub

' Empties every input cell and leaves the edit mode.
Public Sub ClearForm()
    modNames.WriteName NAME_BOOKING_DATE, vbNullString
    modNames.WriteName NAME_BOOKING_CLIENT, vbNullString
    modNames.WriteName NAME_BOOKING_SERVICE, vbNullString
    ClearTimeField FIELD_START
    ClearTimeField FIELD_DURATION
    modNames.WriteName NAME_STATE_EDIT_ID, vbNullString

    RenderPreviews
End Sub

Private Sub ClearTimeField(ByVal fieldKey As String)
    modNames.WriteName HoursName(fieldKey), vbNullString
    modNames.WriteName MinutesName(fieldKey), vbNullString
End Sub

' =====================================================================
'  the previews: refresh computes these, never the input cells (§9.4)
' =====================================================================

' The entry point RenderSheet uses: only the previews, never the input cells.
Public Sub Render()
    RenderPreviews
End Sub

Public Sub RenderPreviews()
    Dim startText As String
    Dim durationValue As Long
    Dim clientId As Long
    Dim serviceId As Long
    Dim editText As String

    startText = ReadTimeField(FIELD_START)
    durationValue = TimeFieldMinutes(FIELD_DURATION)
    clientId = modData.ClientIdByName(modNames.ReadName(NAME_BOOKING_CLIENT))
    serviceId = modData.ServiceIdByName(modNames.ReadName(NAME_BOOKING_SERVICE))
    editText = Trim$(modNames.ReadName(NAME_STATE_EDIT_ID))

    If clientId > 0 Then
        modNames.WriteName NAME_BOOKING_CLIENT_ID, CStr(clientId)
    Else
        modNames.WriteName NAME_BOOKING_CLIENT_ID, vbNullString
    End If

    If serviceId > 0 Then
        modNames.WriteName NAME_BOOKING_PRICE, _
                           modDashboard.FormatMoney(modData.ServicePrice(serviceId))
        modNames.WriteName NAME_BOOKING_SERVICE_DURATION, _
                           modAppointments.DurationText(modData.ServiceDuration(serviceId))
    Else
        modNames.WriteName NAME_BOOKING_PRICE, vbNullString
        modNames.WriteName NAME_BOOKING_SERVICE_DURATION, vbNullString
    End If

    ' The previews are text ("09:30", "1 h 30 min"): in a time or number cell Excel
    ' turns 09:30 into a time and the leading zero is gone (24.09.2026, the same
    ' class of problem as the phone). The format is set here as well as in the
    ' template, so a restyled sheet cannot lose it again.
    modNames.RangeOfName(NAME_BOOKING_START).NumberFormat = "@"
    modNames.RangeOfName(NAME_BOOKING_DURATION).NumberFormat = "@"
    modNames.RangeOfName(NAME_BOOKING_END).NumberFormat = "@"

    modNames.WriteName NAME_BOOKING_START, startText
    modNames.WriteName NAME_BOOKING_DURATION, modAppointments.DurationText(durationValue)
    modNames.WriteName NAME_BOOKING_END, modAppointments.EndTime(startText, durationValue)
    modNames.WriteName NAME_BOOKING_MODE, ModeText(editText)

    ' The button next to Clear is Cancel on a new appointment and Delete while one
    ' is being edited: the weekly grid has no delete column any more (§6.9).
    If Len(editText) > 0 Then
        modNames.WriteName NAME_BTN_BOOKING_CANCEL, modText.Text(modText.KEY_BTN_DELETE)
    Else
        modNames.WriteName NAME_BTN_BOOKING_CANCEL, modText.Text(modText.KEY_BTN_CANCEL)
    End If

    ' The start picker is muted while there is nothing to check against a clash;
    ' the duration picker is always usable.
    modUi.PaintButton SHEET_BOOKING, NAME_BTN_BOOKING_START_PICK, _
                      StartPickerState(clientId, serviceId)
    modUi.PaintButton SHEET_BOOKING, NAME_BTN_BOOKING_DUR_PICK, PAINT_NORMAL
End Sub

Private Function ModeText(ByVal editText As String) As String
    If Len(editText) = 0 Then
        ModeText = modText.Text(modText.KEY_BOOKING_NEW)
    Else
        ModeText = modText.TextFormat(modText.KEY_BOOKING_EDIT, editText)
    End If
End Function

Private Function StartPickerState(ByVal clientId As Long, ByVal serviceId As Long) As String
    StartPickerState = PAINT_NORMAL

    If clientId < 1 Or serviceId < 1 Then
        StartPickerState = PAINT_MUTED
        Exit Function
    End If

    If Len(FormDate()) = 0 Then
        StartPickerState = PAINT_MUTED
    End If
End Function

' =====================================================================
'  the actions of the page
' =====================================================================

' appointments.save: the payload wins when it carries the values (a test or a
' script), otherwise the form decides.
Public Sub OnSave(ByRef data As Object, ByRef result As Object)
    Dim values As Object
    Dim fromForm As Boolean

    Set values = modApp.PayloadFrom(data)
    fromForm = Not HasBookingValues(values)
    If fromForm Then Set values = ReadForm()

    modAppointments.SaveBooking values, result

    ' A saved booking leaves a clean form for the next one.
    If result(modRoutes.RESULT_OK) = True Then
        If fromForm Then ClearForm
    End If
End Sub

' form.pick_date: opens the calendar, or takes data("value") in a test.
Public Sub OnPickDate(ByRef data As Object, ByRef result As Object)
    Dim initialText As String
    Dim picked As String
    Dim info As Object

    initialText = FormDate()
    If Len(initialText) = 0 Then initialText = modDashboard.FormatDateText(Date)

    picked = modApp.PayloadText(data, modRoutes.KEY_VALUE)
    If Len(picked) = 0 Then
        picked = modPickers.PickDate(initialText)
    Else
        picked = modPickers.NormaliseDateText(picked)
    End If

    If Len(picked) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CANCELLED)
        Exit Sub
    End If

    modNames.WriteName NAME_BOOKING_DATE, modPickers.DateFormatText(picked)
    RenderPreviews

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_DATE) = picked
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_OK_PICKED, _
                                                          modText.Text(modText.KEY_LABEL_DATE), _
                                                          modPickers.DateFormatText(picked))
    result(modRoutes.RESULT_OK) = True
End Sub

' form.pick_time: opens the time picker for one of the two time fields (§9.4).
Public Sub OnPickTime(ByRef data As Object, ByRef result As Object)
    Dim fieldKey As String
    Dim initialText As String
    Dim durationValue As Long
    Dim busyText As String
    Dim picked As String
    Dim info As Object

    fieldKey = modApp.PayloadText(data, modRoutes.KEY_FIELD)
    If StrComp(fieldKey, FIELD_DURATION, vbTextCompare) = 0 Then
        fieldKey = FIELD_DURATION
    Else
        fieldKey = FIELD_START
    End If

    initialText = ReadTimeField(fieldKey)
    durationValue = TimeFieldMinutes(FIELD_DURATION)
    busyText = modAppointments.BusyIntervalsOn(FormDate(), EditingId())

    picked = modApp.PayloadText(data, modRoutes.KEY_VALUE)
    If Len(picked) = 0 Then
        picked = modPickers.PickTime(fieldKey, initialText, durationValue, busyText)
    End If

    If Len(picked) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CANCELLED)
        Exit Sub
    End If

    If Not ApplyTimeText(fieldKey, picked) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_TIME, _
                                                              FieldLabel(fieldKey))
        Exit Sub
    End If

    Set info = CreateObject("Scripting.Dictionary")
    info(modRoutes.KEY_FIELD) = fieldKey
    info(COL_START) = ReadTimeField(fieldKey)
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_OK_PICKED, _
                                                          FieldLabel(fieldKey), _
                                                          ReadTimeField(fieldKey))
    result(modRoutes.RESULT_OK) = True
End Sub

' A manual change on the Booking sheet: the previews follow (§7.5, §9.4) and a
' chosen service fills an empty duration.
Public Sub OnChange(ByVal sheetName As String, ByRef target As Range)
    If StrComp(sheetName, SHEET_BOOKING, vbTextCompare) <> 0 Then Exit Sub
    If target Is Nothing Then Exit Sub
    If target.Cells.Count > 1 Then Exit Sub

    If modNames.IsNameAt(target.Worksheet, target, NAME_BOOKING_SERVICE) Then
        FillDurationFromService
    End If

    RenderPreviews
End Sub

' =====================================================================
'  internals
' =====================================================================

Private Function HasBookingValues(ByRef values As Object) As Boolean
    If values Is Nothing Then Exit Function

    HasBookingValues = (values.Exists(COL_START) And values.Exists(COL_CLIENT_ID))
End Function

' The id of the appointment the form is editing; 0 on a new appointment. Public
' because the delete flow (and its test) asks the form: the appointment that was
' being edited must not be edited any more after it was deleted.
Public Function EditingId() As Long
    EditingId = CLng(Val(modNames.ReadName(NAME_STATE_EDIT_ID)))
End Function

Private Function HoursName(ByVal fieldKey As String) As String
    If StrComp(fieldKey, FIELD_DURATION, vbTextCompare) = 0 Then
        HoursName = NAME_BOOKING_DUR_H
    Else
        HoursName = NAME_BOOKING_START_H
    End If
End Function

Private Function MinutesName(ByVal fieldKey As String) As String
    If StrComp(fieldKey, FIELD_DURATION, vbTextCompare) = 0 Then
        MinutesName = NAME_BOOKING_DUR_M
    Else
        MinutesName = NAME_BOOKING_START_M
    End If
End Function

Private Function FieldLabel(ByVal fieldKey As String) As String
    If StrComp(fieldKey, FIELD_DURATION, vbTextCompare) = 0 Then
        FieldLabel = modText.Text(modText.KEY_PICKER_DURATION_TITLE)
    Else
        FieldLabel = modText.Text(modText.KEY_PICKER_TIME_TITLE)
    End If
End Function

' Choosing a service fills an empty duration from db_services.duration (§7.5).
Private Sub FillDurationFromService()
    Dim serviceId As Long
    Dim durationValue As Long

    If TimeFieldMinutes(FIELD_DURATION) > 0 Then Exit Sub

    serviceId = modData.ServiceIdByName(modNames.ReadName(NAME_BOOKING_SERVICE))
    If serviceId < 1 Then Exit Sub

    durationValue = modData.ServiceDuration(serviceId)
    If durationValue < 1 Then durationValue = DEFAULT_DURATION_MINUTES

    WriteTimeField FIELD_DURATION, durationValue \ MINUTES_PER_HOUR, _
                   durationValue Mod MINUTES_PER_HOUR
End Sub
