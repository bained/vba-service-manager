Attribute VB_Name = "modDashboard"
Option Explicit

' =====================================================================
'  Service_Manager - the dashboard (header: phase 1, body: phase 4)
'
'  The header: the view (Daily / Weekly / Monthly), the period, the counters
'  and the active view button. The body draws one of the three views into the
'  block cal_first..cal_last - a list of appointments, a grid of display rows
'  x 7 days or a 6 x 7 calendar (PROJECT_CONCEPT §6.9). Nothing in the body has
'  a name, so a click is recognized by position (§6.10).
'
'  The view and the anchor date live in the state cells s_view / s_date, so the
'  workbook remembers where the user was.
' =====================================================================

Private Const DAYS_IN_WEEK As Long = 7

' =====================================================================
'  state
' =====================================================================

Public Function CurrentView() As String
    Dim viewKey As String

    viewKey = LCase$(modNames.ReadName(NAME_STATE_VIEW))
    If Len(modNames.ViewButtonName(viewKey)) = 0 Then
        viewKey = VIEW_WEEK
        modNames.WriteName NAME_STATE_VIEW, viewKey
    End If

    CurrentView = viewKey
End Function

Public Function CurrentDate() As String
    Dim raw As Variant
    Dim dateText As String

    raw = vbNullString
    On Error Resume Next
    raw = modNames.RangeOfName(NAME_STATE_DATE).Value
    On Error GoTo 0
    If IsError(raw) Then raw = vbNullString

    ' Whatever the cell holds (ISO text, a locale date, a real date, nothing):
    ' DateOfText turns it into a date and the state cell gets the ISO text back.
    dateText = FormatDateText(DateOfText(CStr(raw)))
    If StrComp(dateText, modNames.ReadName(NAME_STATE_DATE), vbBinaryCompare) <> 0 Then
        modNames.WriteName NAME_STATE_DATE, dateText
    End If

    CurrentDate = dateText
End Function

' =====================================================================
'  render
' =====================================================================

' The one entry point of the dashboard: header first, then the body. The whole
' render is one bulk write: ScreenUpdating, EnableEvents and the calculation mode
' are set once for the few hundred cells it touches, and set back even when the
' render fails (§6.9).
Public Sub Render()
    Dim viewKey As String
    Dim anchorText As String
    Dim fromText As String
    Dim toText As String

    modUi.BeginBulk
    On Error GoTo Failed

    viewKey = CurrentView()
    anchorText = CurrentDate()

    WriteViewButtons viewKey
    PeriodBounds viewKey, anchorText, fromText, toText

    modNames.WriteName NAME_VIEW, ViewText(viewKey)
    modNames.WriteName NAME_PERIOD, PeriodTitle(viewKey, anchorText)
    modNames.WriteName NAME_COUNT, CStr(modData.AppointmentsInPeriod(fromText, toText))
    modNames.WriteName NAME_REVENUE, FormatMoney(modData.RevenueInPeriod(fromText, toText))
    modNames.WriteName NAME_FREE_SLOTS, CStr(FreeSlotsFor(viewKey, anchorText))

    RenderBody viewKey, anchorText

    modUi.EndBulk
    Exit Sub

Failed:
    modUi.EndBulk
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' The body is written per view (phase 4, §6.9). Everything it draws lives in the
' block cal_first..cal_last and in the header row above cal_first; nothing there
' has a name, so a click is recognized by position (§6.10).
Private Sub RenderBody(ByVal viewKey As String, ByVal anchorText As String)
    modUi.ClearBlock NAME_CAL_FIRST, NAME_CAL_LAST
    modNames.WriteName NAME_CAL_TITLE, PeriodTitle(viewKey, anchorText)

    Select Case LCase$(viewKey)
        Case VIEW_DAY
            RenderDaily anchorText
        Case VIEW_MONTH
            RenderMonthly anchorText
        Case Else
            RenderWeekly anchorText
    End Select
End Sub

' =====================================================================
'  the geometry of the body
' =====================================================================

' The labels of the body: the row above cal_first. It belongs to the renderer
' because the labels of the weekly and the monthly grid are not the daily ones.
Public Function BodyHeaderRow() As Long
    BodyHeaderRow = BodyFirstRow() - DASH_HEADER_OFFSET
End Function

Public Function BodyFirstRow() As Long
    BodyFirstRow = modNames.RangeOfName(NAME_CAL_FIRST).Row
End Function

Public Function BodyLastRow() As Long
    BodyLastRow = modNames.RangeOfName(NAME_CAL_LAST).Row
End Function

' How many rows the block holds.
Public Function BodyCapacity() As Long
    BodyCapacity = BodyLastRow() - BodyFirstRow() + 1
End Function

' =====================================================================
'  the daily view: one row per slot of the day (§6.9)
' =====================================================================

' The resolution of Config decides how many rows the day has (5 / 15 / 30 / 60
' minutes): every slot is a row, the row of a booked slot carries the appointment
' and the time of a free slot is the button that books it.
Private Sub RenderDaily(ByVal dateText As String)
    Dim ws As Worksheet
    Dim firstRowIndex As Long
    Dim slotCount As Long
    Dim slotIndex As Long
    Dim stepMinutes As Long
    Dim dayItems As Collection

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    firstRowIndex = BodyFirstRow()
    stepMinutes = modAppointments.StepMinutes()
    slotCount = SlotRowCount()

    WriteDailyHeader ws
    Set dayItems = DayAppointments(dateText)

    For slotIndex = 1 To slotCount
        WriteDailySlot ws, firstRowIndex + slotIndex - 1, dateText, _
                       MinuteOfSlot(slotIndex), stepMinutes, dayItems
    Next slotIndex

    If SlotsAreCapped() Then
        modNames.WriteName NAME_CAL_TITLE, modText.TextFormat(modText.KEY_LIST_CAPPED, _
            PeriodTitle(VIEW_DAY, dateText), CStr(slotCount))
    End If
End Sub

Private Sub WriteDailyHeader(ByRef ws As Worksheet)
    Dim headerRowIndex As Long

    headerRowIndex = BodyHeaderRow()
    WriteHeaderCell ws, headerRowIndex, DASH_COL_START, modText.Text(modText.KEY_DASH_COL_TIME)
    WriteHeaderCell ws, headerRowIndex, DASH_COL_END, modText.Text(modText.KEY_DASH_COL_END)
    WriteHeaderCell ws, headerRowIndex, DASH_COL_CLIENT, modText.Text(modText.KEY_DASH_COL_CLIENT)
    WriteHeaderCell ws, headerRowIndex, DASH_COL_SERVICE, modText.Text(modText.KEY_DASH_COL_SERVICE)
    WriteHeaderCell ws, headerRowIndex, DASH_COL_DURATION, modText.Text(modText.KEY_DASH_COL_DURATION)
    WriteHeaderCell ws, headerRowIndex, DASH_COL_PRICE, modText.Text(modText.KEY_DASH_COL_PRICE)
    WriteHeaderCell ws, headerRowIndex, DASH_COL_EDIT, modText.Text(modText.KEY_DASH_COL_EDIT)
    WriteHeaderCell ws, headerRowIndex, DASH_COL_DELETE, modText.Text(modText.KEY_DASH_COL_DELETE)
    WriteHeaderCell ws, headerRowIndex, DASH_COL_ID, modText.Text(modText.KEY_DASH_COL_ID)
End Sub

' One appointment: the times, the client and the service, the duration and the
' price, then the two cell buttons of §6.9. The id sits in the last column of the
' row - the click of [Edit] and [Delete] reads it from there (§6.10).
' One slot of the day (§6.9). Three cases:
'  - free: the time is the button that books a new appointment in that slot,
'  - the appointment starts here: the row carries the end, the client, the
'    service, the duration, the price, [Edit], [Delete] and the id,
'  - the appointment goes on: the row is painted and the client cell says so;
'    its click edits the same appointment.
Private Sub WriteDailySlot(ByRef ws As Worksheet, ByVal sheetRow As Long, ByVal dateText As String, _
                           ByVal slotStartMinutes As Long, ByVal slotMinutes As Long, _
                           ByRef dayItems As Collection)
    Dim owner As Object
    Dim touching As Long
    Dim isTail As Boolean
    Dim slotText As String
    Dim idValue As Long
    Dim label As String
    Dim tipText As String

    slotText = modApp.MinutesToTime(slotStartMinutes)
    Set owner = SlotOwner(dayItems, slotStartMinutes, slotMinutes, touching, isTail)

    If owner Is Nothing Then
        ' a free slot: the time itself books a new appointment
        modUi.WriteTextLinkCell ws.Cells(sheetRow, DASH_COL_START), NAME_CAL_FIRST, slotText, _
            modText.TextFormat(modText.KEY_DASH_TIP_NEW, modPickers.DateFormatText(dateText), _
                               slotText)
        Exit Sub
    End If

    idValue = CLng(owner("id"))
    label = CStr(owner("description"))
    tipText = modText.TextFormat(modText.KEY_DASH_TIP_EDIT, CStr(idValue), label)

    PaintRow ws, sheetRow, COLOR_BOOKED

    If isTail Then
        ' A continuation row is painted, not clickable: it belongs to the appointment
        ' of the row above and a link here looks like a free slot nobody can book. The
        ' appointment keeps its link on its first row (the time) and on [Edit]/[Delete].
        modUi.WriteCellStyled ws.Cells(sheetRow, DASH_COL_START), slotText, _
            COLOR_BOOKED, COLOR_TEXT, True
        modUi.WriteCellStyled ws.Cells(sheetRow, DASH_COL_CLIENT), _
            modText.TextFormat(modText.KEY_DASH_TAIL, CStr(idValue)), _
            COLOR_BOOKED, COLOR_MUTED
        Exit Sub
    End If

    modUi.WriteLinkCellEx ws.Cells(sheetRow, DASH_COL_START), NAME_CAL_FIRST, slotText, _
        tipText, COLOR_BOOKED, COLOR_TEXT, True

    modUi.WriteCellStyled ws.Cells(sheetRow, DASH_COL_END), CStr(owner("end")), _
        COLOR_BOOKED, COLOR_TEXT
    modUi.WriteCellStyled ws.Cells(sheetRow, DASH_COL_CLIENT), CStr(owner("client")), _
        COLOR_BOOKED, COLOR_TEXT
    modUi.WriteCellStyled ws.Cells(sheetRow, DASH_COL_SERVICE), CStr(owner("service")), _
        COLOR_BOOKED, COLOR_TEXT
    modUi.WriteCellStyled ws.Cells(sheetRow, DASH_COL_DURATION), CStr(owner("length")), _
        COLOR_BOOKED, COLOR_TEXT
    modUi.WriteCellStyled ws.Cells(sheetRow, DASH_COL_PRICE), CStr(owner("price")), _
        COLOR_BOOKED, COLOR_TEXT
    modUi.WriteCellStyled ws.Cells(sheetRow, DASH_COL_ID), CStr(idValue), COLOR_BOOKED, COLOR_MUTED

    modUi.WriteLinkCell ws.Cells(sheetRow, DASH_COL_EDIT), NAME_CAL_FIRST, LIST_EDIT_TEXT, tipText
    modUi.WriteLinkCell ws.Cells(sheetRow, DASH_COL_DELETE), NAME_CAL_FIRST, LIST_ACTION_TEXT, _
        modText.TextFormat(modText.KEY_DASH_TIP_DELETE, CStr(idValue), label)
End Sub

' =====================================================================
'  the grids of the body: one row per slot of the working day (§6.9)
' =====================================================================

' The step of Config is the resolution the user set (5 / 15 / 30 / 60 minutes):
' the daily and the weekly grid draw one row per slot, so the resolution is what
' the user sees and clicks. The block of the template holds 151 rows - twelve
' hours of five minute slots fit; a longer day is capped and the title says so.
Public Function SlotsNeeded() As Long
    Dim spanMinutes As Long
    Dim stepMinutes As Long

    stepMinutes = modAppointments.StepMinutes()
    spanMinutes = modAppointments.DayEndMinutes() - modAppointments.DayStartMinutes()
    If spanMinutes < 1 Or stepMinutes < 1 Then Exit Function

    SlotsNeeded = (spanMinutes + stepMinutes - 1) \ stepMinutes
End Function

Public Function SlotsAreCapped() As Boolean
    SlotsAreCapped = (SlotsNeeded() > BodyCapacity())
End Function

' The rows a view draws (never more than the block holds).
Public Function SlotRowCount() As Long
    SlotRowCount = SlotsNeeded()
    If SlotRowCount > BodyCapacity() Then SlotRowCount = BodyCapacity()
End Function

' The minute the slot row starts on (slotIndex 1 = the first slot of the day).
Public Function MinuteOfSlot(ByVal slotIndex As Long) As Long
    MinuteOfSlot = modAppointments.DayStartMinutes() + _
                   (slotIndex - 1) * modAppointments.StepMinutes()
End Function

' The whole row of the body in one fill: a long appointment paints every row it
' covers and the render must not cost a call per cell.
Private Sub PaintRow(ByRef ws As Worksheet, ByVal sheetRow As Long, ByVal fillColor As Long)
    Dim target As Range

    Set target = ws.Range(ws.Cells(sheetRow, DASH_BODY_FIRST_COLUMN), _
                          ws.Cells(sheetRow, DASH_BODY_LAST_COLUMN))
    target.Interior.Color = fillColor
End Sub

Private Sub RenderWeekly(ByVal anchorText As String)
    Dim ws As Worksheet
    Dim firstRowIndex As Long
    Dim stepMinutes As Long
    Dim slotCount As Long
    Dim slotIndex As Long
    Dim dayIndex As Long
    Dim weekStart As Date
    Dim dayStartMinutes As Long
    Dim dayItems As Collection
    Dim weekItems As Collection
    Dim workingDays(0 To DAYS_IN_WEEK - 1) As Boolean

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    firstRowIndex = BodyFirstRow()
    weekStart = WeekStartOf(DateOfText(anchorText))
    dayStartMinutes = modAppointments.DayStartMinutes()
    stepMinutes = modAppointments.StepMinutes()
    slotCount = SlotRowCount()

    ' The working day of a column is read once per render, not once per cell: every
    ' read of Config is a range lookup and the grid can be 120 rows high.
    For dayIndex = 0 To DAYS_IN_WEEK - 1
        workingDays(dayIndex) = modAppointments.IsWorkingDay( _
            FormatDateText(DateAdd("d", dayIndex, weekStart)))
    Next dayIndex

    WriteWeeklyHeader ws, weekStart, workingDays

    WriteWeeklyTimes ws, firstRowIndex, slotCount

    ' The days that are not working days are painted in one go: they carry no
    ' text and no link, only a muted column.
    For dayIndex = 0 To DAYS_IN_WEEK - 1
        If Not workingDays(dayIndex) Then
            PaintQuietColumn ws, firstRowIndex, firstRowIndex + slotCount - 1, _
                             DASH_WEEK_FIRST_DAY_COL + dayIndex
        End If
    Next dayIndex

    ' The appointments of the seven days, read once.
    Set weekItems = New Collection
    For dayIndex = 0 To DAYS_IN_WEEK - 1
        weekItems.Add DayAppointments(FormatDateText(DateAdd("d", dayIndex, weekStart)))
    Next dayIndex

    For slotIndex = 1 To slotCount
        For dayIndex = 0 To DAYS_IN_WEEK - 1
            Set dayItems = weekItems.Item(dayIndex + 1)
            WriteWeeklyCell ws, firstRowIndex + slotIndex - 1, DASH_WEEK_FIRST_DAY_COL + dayIndex, _
                            workingDays(dayIndex), _
                            FormatDateText(DateAdd("d", dayIndex, weekStart)), _
                            MinuteOfSlot(slotIndex), stepMinutes, dayItems
        Next dayIndex
    Next slotIndex

    If SlotsAreCapped() Then
        modNames.WriteName NAME_CAL_TITLE, modText.TextFormat(modText.KEY_LIST_CAPPED, _
            PeriodTitle(VIEW_WEEK, anchorText), CStr(slotCount))
    End If
End Sub

' The time column of the weekly grid in one write: the labels are written as an
' array and the format once for the whole column. Writing them cell by cell meant
' ten property sets per cell - a third of the render for a column of plain text
' (§6.9).
Private Sub WriteWeeklyTimes(ByRef ws As Worksheet, ByVal firstRowIndex As Long, _
                             ByVal slotCount As Long)
    Dim target As Range
    Dim labels() As Variant
    Dim slotIndex As Long

    If slotCount < 1 Then Exit Sub

    ReDim labels(1 To slotCount, 1 To 1)
    For slotIndex = 1 To slotCount
        labels(slotIndex, 1) = modApp.MinutesToTime(MinuteOfSlot(slotIndex))
    Next slotIndex

    Set target = ws.Range(ws.Cells(firstRowIndex, DASH_WEEK_TIME_COL), _
                          ws.Cells(firstRowIndex + slotCount - 1, DASH_WEEK_TIME_COL))
    On Error Resume Next
    target.NumberFormat = "@"
    target.Value = labels
    target.Interior.Color = COLOR_GRID_HEADER
    target.Font.Color = COLOR_TEXT
    target.Font.Bold = True
    target.HorizontalAlignment = xlCenter
    On Error GoTo 0
    Err.Clear
End Sub

' A column that cannot be booked: one fill and one font for the whole column.
Private Sub PaintQuietColumn(ByRef ws As Worksheet, ByVal firstRowIndex As Long, _
                             ByVal lastRowIndex As Long, ByVal columnIndex As Long)
    Dim target As Range

    Set target = ws.Range(ws.Cells(firstRowIndex, columnIndex), ws.Cells(lastRowIndex, columnIndex))
    target.Interior.Color = COLOR_NOWORK
    target.Font.Color = COLOR_MUTED
End Sub

' The labels of the week: the time column and the seven days (with the day number
' and the "today" colour). The grid has no action column any more: an appointment
' is deleted in Booking, where the Cancel button becomes Delete (§6.9).
Private Sub WriteWeeklyHeader(ByRef ws As Worksheet, ByVal weekStart As Date, _
                              ByRef workingDays() As Boolean)
    Dim headerRowIndex As Long
    Dim dayIndex As Long
    Dim dayValue As Date
    Dim dayText As String
    Dim fillColor As Long

    headerRowIndex = BodyHeaderRow()
    WriteHeaderCell ws, headerRowIndex, DASH_WEEK_TIME_COL, modText.Text(modText.KEY_DASH_COL_TIME)

    For dayIndex = 0 To DAYS_IN_WEEK - 1
        dayValue = DateAdd("d", dayIndex, weekStart)
        dayText = FormatDateText(dayValue)
        fillColor = COLOR_GRID_HEADER
        If Not workingDays(dayIndex) Then fillColor = COLOR_NOWORK
        If StrComp(dayText, FormatDateText(Date), vbTextCompare) = 0 Then fillColor = COLOR_TODAY

        modUi.WriteCellStyled ws.Cells(headerRowIndex, DASH_WEEK_FIRST_DAY_COL + dayIndex), _
            modText.TextFormat(modText.KEY_DASH_DAY_HEADER, _
                               modText.WeekdayShortName(dayIndex + 1), CStr(Day(dayValue))), _
            fillColor, COLOR_TEXT, True, False, True
    Next dayIndex
End Sub

' One display row of one day. A booked cell opens the appointment it shows; a
' free cell of a working day books a new one at the time of the row; the columns
' that are not working days carry no link at all (§6.9).
Private Sub WriteWeeklyCell(ByRef ws As Worksheet, ByVal sheetRow As Long, _
                            ByVal columnIndex As Long, ByVal isWorkingDay As Boolean, _
                            ByVal dayText As String, ByVal slotStartMinutes As Long, _
                            ByVal slotMinutes As Long, ByRef dayItems As Collection)
    Dim owner As Object
    Dim touching As Long
    Dim isTail As Boolean
    Dim fillColor As Long
    Dim fontColor As Long
    Dim caption As String

    Set owner = SlotOwner(dayItems, slotStartMinutes, slotMinutes, touching, isTail)

    If owner Is Nothing Then
        If Not isWorkingDay Then Exit Sub        ' the column is painted quiet already
        modUi.WriteEmptyLinkCell ws.Cells(sheetRow, columnIndex), NAME_CAL_FIRST, _
            modText.TextFormat(modText.KEY_DASH_TIP_NEW, modPickers.DateFormatText(dayText), _
                               modApp.MinutesToTime(slotStartMinutes))
        Exit Sub
    End If

    fillColor = COLOR_BOOKED
    If touching > 1 Then fillColor = COLOR_CONFLICT
    fontColor = COLOR_TEXT
    If isTail Then fontColor = COLOR_MUTED

    If isTail Then
        ' The continuation of an appointment is painted, not clickable: the root row
        ' holds the appointment and its link, and a link here would look like a free
        ' slot nobody can book (§6.9).
        modUi.WriteCellStyled ws.Cells(sheetRow, columnIndex), _
            modText.TextFormat(modText.KEY_DASH_TAIL, CStr(owner("id"))), fillColor, fontColor
        Exit Sub
    End If

    caption = SlotLabel(owner, touching)

    ' The tooltip names the appointment (client, service, day, time, id), so a
    ' cell of a crowded week says what a click on it will do.
    modUi.WriteLinkCellEx ws.Cells(sheetRow, columnIndex), NAME_CAL_FIRST, caption, _
        modText.TextFormat(modText.KEY_DASH_TIP_EDIT, CStr(owner("id")), _
                           CStr(owner("description"))), fillColor, fontColor
End Sub

' The reservation a slot row belongs to: the one that starts inside the row,
' otherwise the first one that reaches into it. isTail says that the appointment
' started in an earlier row (the row shows the continuation mark); touching
' counts every reservation of the row, so a row that holds more than one is
' visible (§6.9). A slot normally holds one appointment - the domain refuses
' overlapping ones - but a hand-edited table must not hide a row.
Private Function SlotOwner(ByRef dayItems As Collection, ByVal slotStartMinutes As Long, _
                           ByVal slotMinutes As Long, ByRef touching As Long, _
                           ByRef isTail As Boolean) As Object
    Dim item As Variant
    Dim entry As Object
    Dim owner As Object
    Dim starter As Object
    Dim entryStart As Long

    Set owner = Nothing
    Set starter = Nothing
    touching = 0
    isTail = False

    For Each item In dayItems
        Set entry = item
        entryStart = CLng(entry("start_minutes"))
        If modAppointments.Overlaps(slotStartMinutes, slotMinutes, entryStart, _
                                    CLng(entry("duration"))) Then
            touching = touching + 1
            If owner Is Nothing Then Set owner = entry
            If entryStart >= slotStartMinutes And entryStart < slotStartMinutes + slotMinutes Then
                If starter Is Nothing Then Set starter = entry
            End If
        End If
    Next item

    If Not starter Is Nothing Then Set owner = starter
    If owner Is Nothing Then Exit Function

    isTail = (CLng(owner("start_minutes")) < slotStartMinutes)
    Set SlotOwner = owner
End Function

' "09:00-09:30 Ivan Petrov / Haircut (#12)" - the text of a booked cell. A row
' that holds more than one reservation says so.
Private Function SlotLabel(ByRef item As Object, ByVal touching As Long) As String
    Dim rangeText As String
    Dim whoText As String
    Dim text As String

    rangeText = modText.TextFormat(modText.KEY_DASH_RANGE, CStr(item("start")), _
                                   modAppointments.EndTime(CStr(item("start")), _
                                                           CLng(item("duration"))))
    whoText = modText.TextFormat(modText.KEY_DASH_WHO, CStr(item("client")), _
                                 CStr(item("service")))
    text = modText.TextFormat(modText.KEY_DASH_SLOT, rangeText, whoText, CStr(item("id")))

    If touching > 1 Then
        text = text & " " & modText.TextFormat(modText.KEY_DASH_MORE, CStr(touching - 1))
    End If

    SlotLabel = text
End Function

' The appointments of one day, sorted by the start: the dictionaries carry the
' id, the start (text and minutes), the duration, the client and the service, so
' the grid never asks the table twice for the same row.
Private Function DayAppointments(ByVal dateText As String) As Collection
    Dim items As Collection
    Dim rowsText As String
    Dim rowList As Variant
    Dim index As Long
    Dim table As SqlTable
    Dim dataRow As Long
    Dim entry As Object
    Dim startText As String
    Dim durationValue As Long
    Dim startMinutes As Long
    Dim serviceId As Variant

    Set items = New Collection
    rowsText = modData.AppointmentRowsOn(dateText)

    If Len(rowsText) = 0 Then
        Set DayAppointments = items
        Exit Function
    End If

    Set table = modData.AppointmentsTable()
    rowList = Split(rowsText, ",")

    For index = 0 To UBound(rowList)
        dataRow = CLng(Val(CStr(rowList(index))))
        startText = modData.StartTextOfRow(table, dataRow)
        durationValue = modData.DurationOfRow(table, dataRow)
        startMinutes = modApp.TimeToMinutes(startText)
        If startMinutes >= 0 Then
            serviceId = modData.ServiceIdOfRow(table, dataRow)
            Set entry = CreateObject("Scripting.Dictionary")
            entry("id") = CLng(Val(CStr(modData.IdOfRow(table, dataRow))))
            entry("start") = startText
            entry("start_minutes") = startMinutes
            entry("duration") = durationValue
            entry("end") = modAppointments.EndTime(startText, durationValue)
            entry("length") = modAppointments.DurationText(durationValue)
            entry("client") = modData.ClientName(modData.ClientIdOfRow(table, dataRow))
            entry("service") = modData.ServiceName(serviceId)
            entry("price") = FormatMoney(modData.ServicePrice(serviceId))
            ' how the appointment reads on the screen, for the tooltips of the grid
            entry("description") = modAppointments.AppointmentLabel(CLng(entry("id")))
            items.Add entry
        End If
    Next index

    Set DayAppointments = items
End Function

' =====================================================================
'  the monthly view: a 6 x 7 calendar grid (§6.9)
' =====================================================================

Private Sub RenderMonthly(ByVal anchorText As String)
    Dim ws As Worksheet
    Dim anchorDate As Date
    Dim yearValue As Long
    Dim monthValue As Long
    Dim firstRowIndex As Long
    Dim weekIndex As Long
    Dim dayIndex As Long
    Dim columnIndex As Long
    Dim gridIndex As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    anchorDate = DateOfText(anchorText)
    yearValue = Year(anchorDate)
    monthValue = Month(anchorDate)
    firstRowIndex = BodyFirstRow()

    WriteMonthlyHeader ws

    For weekIndex = 0 To MONTH_GRID_WEEKS - 1
        For dayIndex = 0 To DAYS_IN_WEEK - 1
            columnIndex = DASH_MONTH_FIRST_COL + dayIndex
            gridIndex = weekIndex * DAYS_IN_WEEK + dayIndex + 1
            WriteMonthCell ws, firstRowIndex + weekIndex, columnIndex, _
                           MonthDateText(yearValue, monthValue, gridIndex), _
                           yearValue, monthValue
        Next dayIndex
    Next weekIndex
End Sub

' The date of a cell of the 6 x 7 grid; DateSerial normalises the day number, so
' the cells in front of the month belong to the month before (and the ones after
' it to the month after) - they are drawn muted but stay clickable.
Private Function MonthDateText(ByVal yearValue As Long, ByVal monthValue As Long, _
                               ByVal gridIndex As Long) As String
    Dim dayNumber As Long

    dayNumber = gridIndex - modPickers.MonthStartOffset(yearValue, monthValue)
    MonthDateText = FormatDateText(DateSerial(yearValue, monthValue, dayNumber))
End Function

Private Sub WriteMonthlyHeader(ByRef ws As Worksheet)
    Dim headerRowIndex As Long
    Dim dayIndex As Long

    headerRowIndex = BodyHeaderRow()
    For dayIndex = 1 To DAYS_IN_WEEK
        WriteHeaderCell ws, headerRowIndex, DASH_MONTH_FIRST_COL + dayIndex - 1, _
                        modText.WeekdayShortName(dayIndex)
    Next dayIndex
End Sub

' One day of the calendar: "21: 1 / 25.00" (day, appointments, revenue). The day
' of today is marked; the days of the neighbouring months are muted (§6.9).
Private Sub WriteMonthCell(ByRef ws As Worksheet, ByVal sheetRow As Long, _
                           ByVal columnIndex As Long, ByVal dateText As String, _
                           ByVal yearValue As Long, ByVal monthValue As Long)
    Dim dayValue As Date
    Dim fillColor As Long
    Dim fontColor As Long
    Dim caption As String

    dayValue = DateOfText(dateText)
    fillColor = COLOR_FREE
    fontColor = COLOR_TEXT

    If StrComp(dateText, FormatDateText(Date), vbTextCompare) = 0 Then fillColor = COLOR_TODAY
    If Year(dayValue) <> yearValue Then fontColor = COLOR_MUTED
    If Month(dayValue) <> monthValue Then fontColor = COLOR_MUTED

    caption = modText.TextFormat(modText.KEY_DASH_MONTH_DAY, CStr(Day(dayValue)), _
                                 CStr(modData.AppointmentsOn(dateText)), _
                                 FormatMoney(modData.RevenueOn(dateText)))

    modUi.WriteLinkCellEx ws.Cells(sheetRow, columnIndex), NAME_CAL_FIRST, caption, _
        modText.TextFormat(modText.KEY_DASH_TIP_OPEN_DAY, modPickers.DateFormatText(dateText)), _
        fillColor, fontColor, False, False, True
End Sub

' One label of the body: grey, bold, centered.
Private Sub WriteHeaderCell(ByRef ws As Worksheet, ByVal rowIndex As Long, _
                            ByVal columnIndex As Long, ByVal text As String)
    modUi.WriteCellStyled ws.Cells(rowIndex, columnIndex), text, _
                          COLOR_GRID_HEADER, COLOR_TEXT, True, False, True
End Sub

' The active view button is green, the others are normal.
Private Sub WriteViewButtons(ByVal activeView As String)
    modUi.PaintButton SHEET_DASHBOARD, NAME_BTN_VIEW_DAILY, StateFor(activeView, VIEW_DAY)
    modUi.PaintButton SHEET_DASHBOARD, NAME_BTN_VIEW_WEEKLY, StateFor(activeView, VIEW_WEEK)
    modUi.PaintButton SHEET_DASHBOARD, NAME_BTN_VIEW_MONTHLY, StateFor(activeView, VIEW_MONTH)
End Sub

Private Function StateFor(ByVal activeView As String, ByVal viewKey As String) As String
    If StrComp(activeView, viewKey, vbTextCompare) = 0 Then
        StateFor = "active"
    Else
        StateFor = "normal"
    End If
End Function

' =====================================================================
'  actions
' =====================================================================

Public Sub SetView(ByRef data As Object, ByRef result As Object)
    Dim viewKey As String
    Dim info As Object

    viewKey = LCase$(modApp.PayloadText(data, modRoutes.KEY_VIEW))
    If Len(modNames.ViewButtonName(viewKey)) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_UNKNOWN_ACTION, "view=" & viewKey)
        Exit Sub
    End If

    modNames.WriteName NAME_STATE_VIEW, viewKey

    Set info = CreateObject("Scripting.Dictionary")
    info(modRoutes.KEY_VIEW) = viewKey
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_STATUS_VIEW, _
                                                          ViewText(viewKey), _
                                                          PeriodTitle(viewKey, CurrentDate()))
    result(modRoutes.RESULT_OK) = True
    Render
End Sub

Public Sub Move(ByRef data As Object, ByRef result As Object)
    Dim stepText As String
    Dim stepValue As Long
    Dim viewKey As String
    Dim anchorText As String
    Dim targetDate As Date
    Dim info As Object

    stepText = Trim$(modApp.PayloadText(data, modRoutes.KEY_STEP))
    If Len(stepText) = 0 Then stepText = "0"
    If Not IsNumeric(stepText) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_UNKNOWN_ACTION, "step=" & stepText)
        Exit Sub
    End If
    stepValue = CLng(Val(stepText))

    viewKey = CurrentView()
    anchorText = CurrentDate()
    targetDate = MovedDate(viewKey, anchorText, stepValue)

    modNames.WriteName NAME_STATE_DATE, FormatDateText(targetDate)

    Set info = CreateObject("Scripting.Dictionary")
    info(modRoutes.KEY_STEP) = CStr(stepValue)
    info(NAME_STATE_DATE) = FormatDateText(targetDate)
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_STATUS_VIEW, _
                                                          ViewText(viewKey), _
                                                          PeriodTitle(viewKey, FormatDateText(targetDate)))
    result(modRoutes.RESULT_OK) = True
    Render
End Sub

' dashboard.open_day: the monthly grid hands a day over to the daily view (§6.9).
Public Sub OpenDay(ByRef data As Object, ByRef result As Object)
    Dim dateText As String
    Dim info As Object

    dateText = modData.IsoDateText(modApp.PayloadText(data, modRoutes.KEY_DATE))
    If Len(dateText) < ISO_DATE_LENGTH Then dateText = FormatDateText(Date)

    modNames.WriteName NAME_STATE_VIEW, VIEW_DAY
    modNames.WriteName NAME_STATE_DATE, dateText

    Set info = CreateObject("Scripting.Dictionary")
    info(modRoutes.KEY_VIEW) = VIEW_DAY
    info(NAME_STATE_DATE) = dateText
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_STATUS_VIEW, _
                                                          ViewText(VIEW_DAY), _
                                                          PeriodTitle(VIEW_DAY, dateText))
    result(modRoutes.RESULT_OK) = True
    Render
End Sub

' =====================================================================
'  a click inside the body (§6.10)
' =====================================================================

' The action a click on a body cell means (Nothing when it means nothing). The
' cells have no name, so the position decides: the column of the cell and the
' row it sits on - see the three views below.
Public Function ClickPayload(ByRef cell As Range) As Object
    Set ClickPayload = Nothing
    If cell Is Nothing Then Exit Function

    Select Case LCase$(CurrentView())
        Case VIEW_DAY
            Set ClickPayload = DailyClickPayload(cell)
        Case VIEW_MONTH
            Set ClickPayload = MonthlyClickPayload(cell)
        Case Else
            Set ClickPayload = WeeklyClickPayload(cell)
    End Select
End Function

' The daily grid (§6.9): the [Edit] and [Delete] cells of a booked slot, the
' cell of a continuation (it names its appointment in its text) and the time of
' a free slot, which books a new appointment there.
Private Function DailyClickPayload(ByRef cell As Range) As Object
    Dim idValue As Long

    If cell.Row < BodyFirstRow() Or cell.Row > BodyLastRow() Then Exit Function

    ' the id column carries the plain id (the weekly grid carries "(#12)" inside
    ' the text of the cell)
    idValue = CLng(Val(CellText(SHEET_DASHBOARD, cell.Row, DASH_COL_ID)))
    If idValue > 0 Then
        If cell.Column = DASH_COL_DELETE Then
            Set DailyClickPayload = DeletePayload(idValue, cell)
        Else
            Set DailyClickPayload = EditPayload(idValue)
        End If
        Exit Function
    End If

    ' a continuation row of a long appointment: the cell text carries the id
    idValue = SlotId(CellText(SHEET_DASHBOARD, cell.Row, cell.Column))
    If idValue > 0 Then
        Set DailyClickPayload = EditPayload(idValue)
        Exit Function
    End If

    ' a free slot: its time cell books it
    If cell.Column = DASH_COL_START Then
        Set DailyClickPayload = NewPayload(CurrentDate(), _
                                           CellText(SHEET_DASHBOARD, cell.Row, DASH_COL_START))
    End If
End Function

' The weekly grid: a booked cell (or a continuation cell) opens the appointment
' and a free cell of a working day books one at the time of its slot. There is no
' action column: an appointment is deleted in Booking, where the Cancel button
' becomes Delete (§6.9).
Private Function WeeklyClickPayload(ByRef cell As Range) As Object
    Dim dayIndex As Long
    Dim slotIndex As Long
    Dim slotStartMinutes As Long
    Dim dayText As String
    Dim idValue As Long

    If cell.Column < DASH_WEEK_FIRST_DAY_COL Then Exit Function
    If cell.Column >= DASH_WEEK_FIRST_DAY_COL + DAYS_IN_WEEK Then Exit Function
    If cell.Row < BodyFirstRow() Or cell.Row > BodyLastRow() Then Exit Function

    slotIndex = cell.Row - BodyFirstRow() + 1
    If slotIndex > SlotRowCount() Then Exit Function

    slotStartMinutes = MinuteOfSlot(slotIndex)
    If slotStartMinutes >= modAppointments.DayEndMinutes() Then Exit Function

    dayIndex = cell.Column - DASH_WEEK_FIRST_DAY_COL
    dayText = FormatDateText(DateAdd("d", dayIndex, WeekStartOf(DateOfText(CurrentDate()))))

    idValue = SlotId(CellText(SHEET_DASHBOARD, cell.Row, cell.Column))
    If idValue > 0 Then
        Set WeeklyClickPayload = EditPayload(idValue)
    Else
        Set WeeklyClickPayload = NewPayload(dayText, modApp.MinutesToTime(slotStartMinutes))
    End If
End Function

' The monthly grid: every day cell opens that day in the daily view - including
' the muted cells of the months around it, so the calendar never dead-ends.
Private Function MonthlyClickPayload(ByRef cell As Range) As Object
    Dim anchorDate As Date
    Dim weekIndex As Long
    Dim dayIndex As Long
    Dim gridIndex As Long
    Dim dayNumber As Long

    If cell.Column < DASH_MONTH_FIRST_COL Then Exit Function
    If cell.Column >= DASH_MONTH_FIRST_COL + DAYS_IN_WEEK Then Exit Function

    weekIndex = cell.Row - BodyFirstRow()
    If weekIndex < 0 Or weekIndex > MONTH_GRID_WEEKS - 1 Then Exit Function

    dayIndex = cell.Column - DASH_MONTH_FIRST_COL
    gridIndex = weekIndex * DAYS_IN_WEEK + dayIndex + 1

    anchorDate = DateOfText(CurrentDate())
    dayNumber = gridIndex - modPickers.MonthStartOffset(Year(anchorDate), Month(anchorDate))

    Set MonthlyClickPayload = OpenDayPayload( _
        FormatDateText(DateSerial(Year(anchorDate), Month(anchorDate), dayNumber)))
End Function

' The appointment id inside a rendered text: "(#12)" of a booked cell and
' "#12" of a continuation cell both carry it, so a booked cell needs no column of
' its own (§6.9).
Private Function SlotId(ByVal text As String) As Long
    Dim hash As Long
    Dim index As Long
    Dim digits As String
    Dim character As String

    hash = InStr(1, text, "#", vbBinaryCompare)
    If hash = 0 Then Exit Function

    For index = hash + 1 To Len(text)
        character = Mid$(text, index, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit For
        digits = digits & character
    Next index

    If Len(digits) = 0 Then Exit Function
    SlotId = CLng(Val(digits))
End Function

' The text of one cell of a rendered block ("" when it is empty).
Private Function CellText(ByVal sheetName As String, ByVal rowIndex As Long, _
                          ByVal columnIndex As Long) As String
    Dim value As Variant

    value = vbNullString
    On Error Resume Next
    value = ThisWorkbook.Worksheets(sheetName).Cells(rowIndex, columnIndex).Value
    On Error GoTo 0

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    CellText = Trim$(CStr(value))
End Function

' The payloads a click builds: the action, the id and - for a destructive action
' - the cell, so the confirmation can paint it (§6.6).
Private Function EditPayload(ByVal idValue As Long) As Object
    Dim payload As Object

    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_ACTION) = modRoutes.ACTION_APPOINTMENTS_EDIT
    payload(modRoutes.KEY_ID) = CStr(idValue)

    Set EditPayload = payload
End Function

Private Function DeletePayload(ByVal idValue As Long, ByRef cell As Range) As Object
    Dim payload As Object

    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_ACTION) = modRoutes.ACTION_APPOINTMENTS_DELETE
    payload(modRoutes.KEY_ID) = CStr(idValue)
    payload(modRoutes.KEY_SHEET) = SHEET_DASHBOARD
    payload(modRoutes.KEY_CELL) = cell.Address(False, False)

    Set DeletePayload = payload
End Function

Private Function NewPayload(ByVal dateText As String, ByVal startText As String) As Object
    Dim payload As Object

    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_ACTION) = modRoutes.ACTION_APPOINTMENTS_NEW
    payload(COL_DATE) = dateText
    If Len(startText) > 0 Then payload(COL_START) = startText

    Set NewPayload = payload
End Function

Private Function OpenDayPayload(ByVal dateText As String) As Object
    Dim payload As Object

    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_ACTION) = modRoutes.ACTION_DASHBOARD_OPEN_DAY
    payload(modRoutes.KEY_DATE) = dateText

    Set OpenDayPayload = payload
End Function

' =====================================================================
'  period and format helpers
' =====================================================================

' The first and the last date of the shown period.
Public Sub PeriodBounds(ByVal viewKey As String, ByVal anchorText As String, _
                        ByRef fromText As String, ByRef toText As String)
    Dim anchorDate As Date
    Dim weekStart As Date
    Dim monthStart As Date

    anchorDate = DateOfText(anchorText)

    Select Case LCase$(viewKey)
        Case VIEW_DAY
            fromText = FormatDateText(anchorDate)
            toText = fromText
        Case VIEW_MONTH
            monthStart = DateSerial(Year(anchorDate), Month(anchorDate), 1)
            fromText = FormatDateText(monthStart)
            toText = FormatDateText(DateSerial(Year(anchorDate), Month(anchorDate) + 1, 0))
        Case Else
            weekStart = WeekStartOf(anchorDate)
            fromText = FormatDateText(weekStart)
            toText = FormatDateText(DateAdd("d", DAYS_IN_WEEK - 1, weekStart))
    End Select
End Sub

Public Function PeriodTitle(ByVal viewKey As String, ByVal anchorText As String) As String
    Dim anchorDate As Date
    Dim weekStart As Date

    anchorDate = DateOfText(anchorText)

    Select Case LCase$(viewKey)
        Case VIEW_DAY
            PeriodTitle = modText.WeekdayShortName(Weekday(anchorDate, vbMonday)) & " " & _
                          CStr(Day(anchorDate)) & " " & _
                          modText.MonthShortName(Month(anchorDate)) & " " & CStr(Year(anchorDate))
        Case VIEW_MONTH
            PeriodTitle = modText.MonthLongName(Month(anchorDate)) & " " & CStr(Year(anchorDate))
        Case Else
            weekStart = WeekStartOf(anchorDate)
            PeriodTitle = CStr(Day(weekStart)) & "-" & _
                          CStr(Day(DateAdd("d", DAYS_IN_WEEK - 1, weekStart))) & " " & _
                          modText.MonthShortName(Month(weekStart)) & " " & CStr(Year(weekStart))
    End Select
End Function

' Free slots of the shown period: the working day divided by the step, minus the
' appointments that occupy a slot.
Public Function FreeSlotsFor(ByVal viewKey As String, ByVal anchorText As String) As Long
    Dim dayStart As Long
    Dim dayEnd As Long
    Dim stepMinutes As Long
    Dim slotsPerDay As Long
    Dim daysInPeriod As Long
    Dim booked As Long
    Dim fromText As String
    Dim toText As String

    dayStart = modApp.TimeToMinutes(ConfigText(NAME_CONFIG_DAY_START, DEFAULT_DAY_START))
    dayEnd = modApp.TimeToMinutes(ConfigText(NAME_CONFIG_DAY_END, DEFAULT_DAY_END))
    stepMinutes = ConfigNumber(NAME_CONFIG_STEP, DEFAULT_STEP_MINUTES)

    If stepMinutes < 1 Then stepMinutes = DEFAULT_STEP_MINUTES
    If dayEnd <= dayStart Then
        FreeSlotsFor = 0
        Exit Function
    End If

    slotsPerDay = (dayEnd - dayStart) \ stepMinutes
    PeriodBounds viewKey, anchorText, fromText, toText
    daysInPeriod = DayCount(fromText, toText)
    booked = modData.AppointmentsInPeriod(fromText, toText)

    FreeSlotsFor = slotsPerDay * daysInPeriod - booked
    If FreeSlotsFor < 0 Then FreeSlotsFor = 0
End Function

' The date the period moves to (step 0 = today, whatever the view).
Private Function MovedDate(ByVal viewKey As String, ByVal anchorText As String, _
                           ByVal stepValue As Long) As Date
    Dim anchorDate As Date

    If stepValue = 0 Then
        MovedDate = Date
        Exit Function
    End If

    anchorDate = DateOfText(anchorText)
    Select Case LCase$(viewKey)
        Case VIEW_DAY
            MovedDate = DateAdd("d", stepValue, anchorDate)
        Case VIEW_MONTH
            MovedDate = DateAdd("m", stepValue, anchorDate)
        Case Else
            MovedDate = DateAdd("d", stepValue * DAYS_IN_WEEK, anchorDate)
    End Select
End Function

Public Function ViewText(ByVal viewKey As String) As String
    Select Case LCase$(Trim$(viewKey))
        Case VIEW_DAY
            ViewText = VIEW_TEXT_DAY
        Case VIEW_MONTH
            ViewText = VIEW_TEXT_MONTH
        Case Else
            ViewText = VIEW_TEXT_WEEK
    End Select
End Function

Public Function FormatDateText(ByVal value As Date) As String
    ' Built from the numbers: Format$(value, "yyyy-mm-dd") would follow the
    ' regional settings of the machine.
    FormatDateText = Format$(Year(value), "0000") & "-" & Format$(Month(value), "00") & "-" & _
                     Format$(Day(value), "00")
End Function

' ISO text (or a date) to a date; an empty or broken text gives today.
Public Function DateOfText(ByVal text As String) As Date
    Dim trimmed As String

    trimmed = Trim$(text)

    If Len(trimmed) >= ISO_DATE_LENGTH Then
        If IsDate(trimmed) Then
            DateOfText = CDate(trimmed)
            Exit Function
        End If
    End If

    If IsDate(trimmed) Then
        DateOfText = CDate(trimmed)
        Exit Function
    End If

    DateOfText = Date
End Function

' Monday of the week the date belongs to (ISO weeks).
Public Function WeekStartOf(ByVal value As Date) As Date
    WeekStartOf = DateAdd("d", 1 - Weekday(value, vbMonday), value)
End Function

Private Function DayCount(ByVal fromText As String, ByVal toText As String) As Long
    DayCount = DateDiff("d", DateOfText(fromText), DateOfText(toText)) + 1
    If DayCount < 1 Then DayCount = 1
End Function

' Money with the number format of the Config sheet.
Public Function FormatMoney(ByVal value As Double) As String
    Dim formatText As String

    formatText = ConfigText(NAME_CONFIG_NUMBER_FORMAT, DEFAULT_NUMBER_FORMAT)
    If Len(formatText) = 0 Then formatText = DEFAULT_NUMBER_FORMAT

    FormatMoney = Format$(value, formatText)
End Function

' A setting of the Config sheet, with a built in fallback.
Public Function ConfigText(ByVal nameText As String, ByVal fallback As String) As String
    Dim value As String

    value = modNames.ReadName(nameText)
    If Len(value) = 0 Then value = fallback

    ConfigText = value
End Function

Public Function ConfigNumber(ByVal nameText As String, ByVal fallback As Long) As Long
    Dim value As String

    value = modNames.ReadName(nameText)
    If Not IsNumeric(value) Then
        ConfigNumber = fallback
        Exit Function
    End If

    ConfigNumber = CLng(Val(value))
End Function
