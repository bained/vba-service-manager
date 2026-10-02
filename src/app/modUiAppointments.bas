Attribute VB_Name = "modUiAppointments"
Option Explicit

' =====================================================================
'  Service_Manager - the all-appointments page
'
'  The table is rebuilt from the current data model whenever it is activated,
'  refreshed explicitly, or any related client/service/appointment is changed.
'  Dynamic row actions are recognized by column and row position, as on the
'  ClientsAdmin and ServicesAdmin pages.
' =====================================================================

Private mLastRenderedRow As Long

' appointments.list: redraw the complete list and report the number of rows.
Public Sub OnList(ByRef data As Object, ByRef result As Object)
    Dim rows As Long
    Dim info As Object

    ' The explicit Refresh button is also the recovery path after a manual edit
    ' on a visible db_* sheet, so do not rely on modData's in-memory snapshot.
    modData.LoadData
    modData.RefreshLists
    rows = Render()

    Set info = CreateObject("Scripting.Dictionary")
    info("rows") = CStr(rows)
    Set result(modRoutes.RESULT_DATA) = info
    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
        modText.KEY_STATUS_APPOINTMENTS_REFRESHED, CStr(rows))
    result(modRoutes.RESULT_OK) = True
End Sub

' Draw every appointment, joined with the current client contact details and
' service name. There is no fixed row cap: the used region grows with the data.
Public Function Render() As Long
    Dim ws As Worksheet
    Dim firstCell As Range
    Dim appointmentTable As SqlTable
    Dim clientTable As SqlTable
    Dim clientValues As Object
    Dim client As Object
    Dim clientId As String
    Dim appointmentCount As Long
    Dim dataRow As Long
    Dim sheetRow As Long
    Dim firstRow As Long
    Dim clearLastRow As Long
    Dim eventsWereOn As Boolean
    Dim titleText As String

    If Not modNames.SheetExists(SHEET_APPOINTMENTS) Then Exit Function
    If Not modData.IsModelLoaded() Then modData.LoadData
    If mLastRenderedRow = 0 Then
        mLastRenderedRow = modNames.RangeOfName(NAME_GRID_APPOINTMENTS_FIRST).Row
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_APPOINTMENTS)
    Set firstCell = modNames.RangeOfName(NAME_GRID_APPOINTMENTS_FIRST)
    firstRow = firstCell.Row
    Set appointmentTable = modData.AppointmentsTable()
    Set clientTable = modData.ClientsTable()
    appointmentCount = appointmentTable.RowCount

    clearLastRow = mLastRenderedRow
    If ws.Cells(ws.Rows.Count, APPOINTMENTS_COL_ID).End(xlUp).Row > clearLastRow Then
        clearLastRow = ws.Cells(ws.Rows.Count, APPOINTMENTS_COL_ID).End(xlUp).Row
    End If
    If firstRow + appointmentCount - 1 > clearLastRow Then
        clearLastRow = firstRow + appointmentCount - 1
    End If

    modUi.BeginBulk
    On Error GoTo Failed
    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    If clearLastRow >= firstRow Then ClearRows ws, firstRow, clearLastRow

    Set clientValues = ClientLookup(clientTable)
    ' The newest appointment comes first: the rows are visited in the order of
    ' AppointmentRowsNewestFirst - date descending, then start descending, then id
    ' descending (the owner's rule, 25.09.2026). The sheet row still counts forward,
    ' so the block stays top-down.
    Dim order As Collection
    Dim slotIndex As Long

    Set order = modData.AppointmentRowsNewestFirst()
    For slotIndex = 1 To order.Count
        dataRow = order(slotIndex)
        sheetRow = firstRow + slotIndex - 1
        clientId = Trim$(CStr(modData.ClientIdOfRow(appointmentTable, dataRow)))
        If clientValues.Exists(clientId) Then
            Set client = clientValues(clientId)
        Else
            Set client = Nothing
        End If
        WriteAppointmentRow ws, sheetRow, dataRow, appointmentTable, client
    Next slotIndex

    titleText = modText.TextFormat(modText.KEY_APPOINTMENTS_TITLE, CStr(appointmentCount))
    modNames.WriteName NAME_APPOINTMENTS_TITLE, titleText

    If appointmentCount = 0 Then
        modUi.WriteCell SHEET_APPOINTMENTS, firstRow, APPOINTMENTS_COL_ID, _
                        modText.Text(modText.KEY_APPOINTMENTS_NONE)
    End If

    mLastRenderedRow = firstRow + appointmentCount - 1
    Render = appointmentCount
    GoTo Cleanup

Failed:
    Application.EnableEvents = eventsWereOn
    modUi.EndBulk
    Err.Raise Err.Number, "modUiAppointments.Render", Err.Description

Cleanup:
    Application.EnableEvents = eventsWereOn
    modUi.EndBulk
End Function

' The position and the action column define a dynamic row button. The id is read
' from the same row, never inferred from the hyperlink's anchor.
Public Function ClickPayload(ByRef cell As Range) As Object
    Dim actionText As String
    Dim firstRow As Long
    Dim idValue As Variant
    Dim payload As Object

    Set ClickPayload = Nothing
    If cell Is Nothing Then Exit Function
    If StrComp(cell.Worksheet.Name, SHEET_APPOINTMENTS, vbTextCompare) <> 0 Then Exit Function

    Select Case cell.Column
        Case APPOINTMENTS_COL_EDIT
            actionText = modRoutes.ACTION_APPOINTMENTS_EDIT
        Case APPOINTMENTS_COL_DELETE
            actionText = modRoutes.ACTION_APPOINTMENTS_DELETE
        Case Else
            Exit Function
    End Select

    firstRow = modNames.RangeOfName(NAME_GRID_APPOINTMENTS_FIRST).Row
    If cell.Row < firstRow Or cell.Row > mLastRenderedRow Then Exit Function

    idValue = ThisWorkbook.Worksheets(SHEET_APPOINTMENTS).Cells( _
        cell.Row, APPOINTMENTS_COL_ID).Value
    If Not IsNumeric(idValue) Then Exit Function
    If CLng(Val(CStr(idValue))) < 1 Then Exit Function

    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_ACTION) = actionText
    payload(modRoutes.KEY_ID) = CStr(CLng(Val(CStr(idValue))))
    payload(modRoutes.KEY_SHEET) = SHEET_APPOINTMENTS
    payload(modRoutes.KEY_CELL) = cell.Address(False, False)
    Set ClickPayload = payload
End Function

Private Function ClientLookup(ByRef table As SqlTable) As Object
    Dim lookup As Object
    Dim values As Object
    Dim rowIndex As Long
    Dim idText As String

    Set lookup = CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To table.RowCount
        idText = Trim$(modData.CellTextOf(table, rowIndex, COL_ID))
        If Len(idText) > 0 Then
            Set values = CreateObject("Scripting.Dictionary")
            values(COL_CLIENT_NAME) = modData.CellTextOf(table, rowIndex, COL_CLIENT_NAME)
            values(COL_CLIENT_PHONE) = modData.CellTextOf(table, rowIndex, COL_CLIENT_PHONE)
            values(COL_CLIENT_EMAIL) = modData.CellTextOf(table, rowIndex, COL_CLIENT_EMAIL)
            Set lookup(idText) = values
        End If
    Next rowIndex

    Set ClientLookup = lookup
End Function

Private Sub WriteAppointmentRow(ByRef ws As Worksheet, ByVal sheetRow As Long, _
                                ByVal dataRow As Long, ByRef table As SqlTable, _
                                ByRef client As Object)
    Dim idText As String
    Dim startText As String
    Dim duration As Long
    Dim dateText As String
    Dim clientName As String
    Dim phoneText As String
    Dim emailText As String
    Dim serviceName As String
    Dim endText As String

    idText = modData.CellTextOf(table, dataRow, COL_ID)
    dateText = modData.DateTextOfRow(table, dataRow)
    startText = modData.StartTextOfRow(table, dataRow)
    duration = modData.DurationOfRow(table, dataRow)
    endText = modAppointments.EndTime(startText, duration)
    serviceName = modData.ServiceName(modData.ServiceIdOfRow(table, dataRow))

    If Not client Is Nothing Then
        clientName = CStr(client(COL_CLIENT_NAME))
        phoneText = CStr(client(COL_CLIENT_PHONE))
        emailText = CStr(client(COL_CLIENT_EMAIL))
    End If

    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_ID, idText
    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_DATE, dateText
    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_START, startText
    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_END, endText
    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_DURATION, _
                    modText.TextFormat(modText.KEY_TIME_MINUTES, CStr(duration))
    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_CLIENT, clientName

    ' Date, start, end, duration and phone are text: in a time or number cell Excel
    ' turns 09:30 into a time and the leading zero is gone (25.09.2026, the same
    ' class of problem as the phone). Set before writing, so a restyle cannot lose it.
    ws.Cells(sheetRow, APPOINTMENTS_COL_DATE).NumberFormat = "@"
    ws.Cells(sheetRow, APPOINTMENTS_COL_START).NumberFormat = "@"
    ws.Cells(sheetRow, APPOINTMENTS_COL_END).NumberFormat = "@"
    ws.Cells(sheetRow, APPOINTMENTS_COL_DURATION).NumberFormat = "@"
    ws.Cells(sheetRow, APPOINTMENTS_COL_PHONE).NumberFormat = "@"
    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_PHONE, phoneText
    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_EMAIL, emailText
    modUi.WriteCell SHEET_APPOINTMENTS, sheetRow, APPOINTMENTS_COL_SERVICE, serviceName

    modUi.WriteLinkCell ws.Cells(sheetRow, APPOINTMENTS_COL_EDIT), _
        NAME_GRID_APPOINTMENTS_FIRST, LIST_EDIT_TEXT, "Edit appointment #" & idText
    modUi.WriteLinkCell ws.Cells(sheetRow, APPOINTMENTS_COL_DELETE), _
        NAME_GRID_APPOINTMENTS_FIRST, LIST_ACTION_TEXT, "Delete appointment #" & idText
End Sub

Private Sub ClearRows(ByRef ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim rows As Range
    Dim links As Object

    Set rows = ws.Range(ws.Cells(firstRow, APPOINTMENTS_COL_ID), _
                        ws.Cells(lastRow, APPOINTMENTS_COL_DELETE))
    On Error Resume Next
    Set links = rows.Hyperlinks
    If Not links Is Nothing Then
        If links.Count > 0 Then links.Delete
    End If
    Err.Clear
    On Error GoTo 0
    rows.ClearContents
End Sub