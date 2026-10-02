Attribute VB_Name = "modData"
Option Explicit

' =====================================================================
'  Service_Manager - the data layer (Phase 1)
'
'  Loads the db_* sheets into SqlTable objects and exposes the lookups the
'  shell needs (counts, names, prices, revenue, next id). Writes come in the
'  next phases and go through SqlMutation here, never from a UI module.
'
'  Contract: header row 1, data from row 2, column A is the id column;
'  "date" is ISO text YYYY-MM-DD, "start" is HH:MM text, "duration" is minutes.
' =====================================================================

Private mClients As SqlTable
Private mServices As SqlTable
Private mAppointments As SqlTable
Private mLoaded As Boolean

' =====================================================================
'  loading
' =====================================================================

Public Sub LoadData()
    Set mClients = LoadTable(SHEET_DB_CLIENTS)
    Set mServices = LoadTable(SHEET_DB_SERVICES)
    Set mAppointments = LoadTable(SHEET_DB_APPOINTMENTS)
    mLoaded = True
End Sub

Public Sub UnloadData()
    Set mClients = Nothing
    Set mServices = Nothing
    Set mAppointments = Nothing
    mLoaded = False
End Sub

Public Function IsModelLoaded() As Boolean
    IsModelLoaded = mLoaded
End Function

' Loads one DB sheet as a range table (header row 1, data from row 2).
Public Function LoadTable(ByVal sheetName As String) As SqlTable
    Dim loader As SqlTable

    If Len(Trim$(sheetName)) = 0 Then
        Err.Raise 5, "modData.LoadTable", "The sheet name cannot be empty."
    End If
    If Not modNames.SheetExists(sheetName) Then
        Err.Raise 9, "modData.LoadTable", modText.TextFormat(modText.KEY_ERROR_NO_SHEET, sheetName)
    End If

    Set loader = New SqlTable
    Set LoadTable = loader.AutoRange(sheetName)
End Function

' =====================================================================
'  tables
' =====================================================================

Public Function ClientsTable() As SqlTable
    EnsureLoaded "ClientsTable"
    Set ClientsTable = mClients
End Function

Public Function ServicesTable() As SqlTable
    EnsureLoaded "ServicesTable"
    Set ServicesTable = mServices
End Function

Public Function AppointmentsTable() As SqlTable
    EnsureLoaded "AppointmentsTable"
    Set AppointmentsTable = mAppointments
End Function

Private Sub EnsureLoaded(ByVal caller As String)
    If Not mLoaded Then
        Err.Raise 5, "modData." & caller, modText.TextFormat(modText.KEY_ERROR_NO_DATA, caller)
    End If
End Sub

' Number of data rows (without the header) of a DB sheet.
Public Function RowCountOf(ByVal sheetName As String) As Long
    RowCountOf = LoadTable(sheetName).RowCount
End Function

Public Function ClientsRowCount() As Long
    ClientsRowCount = ClientsTable().RowCount
End Function

Public Function ServicesRowCount() As Long
    ServicesRowCount = ServicesTable().RowCount
End Function

Public Function AppointmentsRowCount() As Long
    AppointmentsRowCount = AppointmentsTable().RowCount
End Function

' =====================================================================
'  ids and lookups
' =====================================================================

' Highest existing id + 1 (no id is special).
Public Function NextId(ByRef table As SqlTable) As Long
    Dim rowIndex As Long
    Dim highest As Long
    Dim value As Variant

    If Not table.HasColumn(COL_ID) Then Exit Function

    For rowIndex = 1 To table.RowCount
        value = table.GetCell(rowIndex, COL_ID)
        If IsNumeric(value) Then
            If CLng(Val(CStr(value))) > highest Then highest = CLng(Val(CStr(value)))
        End If
    Next rowIndex

    NextId = highest + 1
End Function

' An unknown id gives "" (the calendar must be able to draw a row whose
' reference was deleted), not an error.
Public Function ClientName(ByVal clientId As Variant) As String
    ClientName = LookupText(ClientsTable(), COL_ID, clientId, COL_CLIENT_NAME)
End Function

Public Function ServiceName(ByVal serviceId As Variant) As String
    ServiceName = LookupText(ServicesTable(), COL_ID, serviceId, COL_SERVICE_NAME)
End Function

Public Function ServicePrice(ByVal serviceId As Variant) As Double
    Dim rowIndex As Long
    Dim value As Variant

    rowIndex = FindRowByText(ServicesTable(), COL_ID, serviceId)
    If rowIndex = 0 Then Exit Function

    value = ServicesTable().GetCell(rowIndex, COL_SERVICE_PRICE)
    If IsNumeric(value) Then ServicePrice = CDbl(value)
End Function

Public Function ServiceDuration(ByVal serviceId As Variant) As Long
    Dim rowIndex As Long
    Dim value As Variant

    rowIndex = FindRowByText(ServicesTable(), COL_ID, serviceId)
    If rowIndex = 0 Then Exit Function

    value = ServicesTable().GetCell(rowIndex, COL_SERVICE_DURATION)
    If IsNumeric(value) Then ServiceDuration = CLng(value)
End Function

Private Function LookupText(ByRef table As SqlTable, ByVal keyColumn As String, _
                            ByVal keyValue As Variant, ByVal valueColumn As String) As String
    Dim rowIndex As Long

    rowIndex = FindRowByText(table, keyColumn, keyValue)
    If rowIndex = 0 Then Exit Function

    LookupText = TextValue(table.GetCell(rowIndex, valueColumn))
End Function

' =====================================================================
'  rows of the appointments table
' =====================================================================

Public Function IdOfRow(ByRef table As SqlTable, ByVal rowIndex As Long) As Variant
    IdOfRow = table.GetCell(rowIndex, COL_ID)
End Function

Public Function ClientIdOfRow(ByRef table As SqlTable, ByVal rowIndex As Long) As Variant
    ClientIdOfRow = table.GetCell(rowIndex, COL_CLIENT_ID)
End Function

Public Function ServiceIdOfRow(ByRef table As SqlTable, ByVal rowIndex As Long) As Variant
    ServiceIdOfRow = table.GetCell(rowIndex, COL_SERVICE_ID)
End Function

' The date of a row as ISO text, whatever the cell holds (text, date, number).
Public Function DateTextOfRow(ByRef table As SqlTable, ByVal rowIndex As Long) As String
    DateTextOfRow = IsoDateText(table.GetCell(rowIndex, COL_DATE))
End Function

' The start of a row as HH:MM text.
Public Function StartTextOfRow(ByRef table As SqlTable, ByVal rowIndex As Long) As String
    StartTextOfRow = TimeTextValue(table.GetCell(rowIndex, COL_START))
End Function

' The duration of a row in minutes (0 when the column is missing).
Public Function DurationOfRow(ByRef table As SqlTable, ByVal rowIndex As Long) As Long
    Dim value As Variant

    If Not table.HasColumn(COL_SERVICE_DURATION) Then Exit Function

    value = table.GetCell(rowIndex, COL_SERVICE_DURATION)
    DurationOfRow = DurationValue(value)
End Function

' =====================================================================
'  counters of a period (the dashboard header)
' =====================================================================

Public Function AppointmentsOn(ByVal isoDate As String) As Long
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String
    Dim count As Long

    wanted = IsoDateText(isoDate)
    If Len(wanted) = 0 Then Exit Function
    If Len(wanted) < ISO_DATE_LENGTH Then Exit Function

    Set table = AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        If StrComp(DateTextOfRow(table, rowIndex), wanted, vbTextCompare) = 0 Then
            count = count + 1
        End If
    Next rowIndex

    AppointmentsOn = count
End Function

' The data rows of one day, sorted by the start of the appointment (then by id),
' as a comma separated list of row indexes of the table. The three views of the
' dashboard draw themselves from this list (§6.9), so the order of a day is
' decided here once and never depends on the order of the rows in the sheet.
Public Function AppointmentRowsOn(ByVal isoDate As String) As String
    Dim table As SqlTable
    Dim wanted As String
    Dim rowIndex As Long
    Dim inner As Long
    Dim count As Long
    Dim rows() As Long
    Dim keys() As String
    Dim index As Long
    Dim swapRow As Long
    Dim swapKey As String
    Dim startMinutes As Long
    Dim text As String

    wanted = IsoDateText(isoDate)
    If Len(wanted) < ISO_DATE_LENGTH Then Exit Function

    Set table = AppointmentsTable()
    If table.RowCount < 1 Then Exit Function

    ReDim rows(1 To table.RowCount)
    ReDim keys(1 To table.RowCount)

    For rowIndex = 1 To table.RowCount
        If StrComp(DateTextOfRow(table, rowIndex), wanted, vbTextCompare) = 0 Then
            startMinutes = modApp.TimeToMinutes(StartTextOfRow(table, rowIndex))
            If startMinutes < 0 Then startMinutes = MINUTES_PER_DAY
            count = count + 1
            rows(count) = rowIndex
            keys(count) = Format$(startMinutes, "0000") & "|" & _
                          Format$(CLng(Val(CStr(IdOfRow(table, rowIndex)))), "000000")
        End If
    Next rowIndex

    ' A selection sort: a day holds a handful of rows and this keeps the order of
    ' the views stable without a second data structure.
    For index = 1 To count - 1
        For inner = index + 1 To count
            If keys(inner) < keys(index) Then
                swapRow = rows(index)
                rows(index) = rows(inner)
                rows(inner) = swapRow
                swapKey = keys(index)
                keys(index) = keys(inner)
                keys(inner) = swapKey
            End If
        Next inner
    Next index

    For index = 1 To count
        If Len(text) > 0 Then text = text & ","
        text = text & CStr(rows(index))
    Next index

    AppointmentRowsOn = text
End Function

Public Function AppointmentsInPeriod(ByVal fromIso As String, ByVal toIso As String) As Long
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim fromText As String
    Dim toText As String
    Dim rowDate As String
    Dim count As Long

    fromText = IsoDateText(fromIso)
    toText = IsoDateText(toIso)
    If Len(fromText) < ISO_DATE_LENGTH Or Len(toText) < ISO_DATE_LENGTH Then Exit Function

    Set table = AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        rowDate = DateTextOfRow(table, rowIndex)
        If Len(rowDate) >= ISO_DATE_LENGTH Then
            If rowDate >= fromText And rowDate <= toText Then count = count + 1
        End If
    Next rowIndex

    AppointmentsInPeriod = count
End Function

' Revenue of a period: the price of the service of every appointment.
Public Function RevenueInPeriod(ByVal fromIso As String, ByVal toIso As String) As Double
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim fromText As String
    Dim toText As String
    Dim rowDate As String
    Dim total As Double

    fromText = IsoDateText(fromIso)
    toText = IsoDateText(toIso)
    If Len(fromText) < ISO_DATE_LENGTH Or Len(toText) < ISO_DATE_LENGTH Then Exit Function

    Set table = AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        rowDate = DateTextOfRow(table, rowIndex)
        If Len(rowDate) >= ISO_DATE_LENGTH Then
            If rowDate >= fromText And rowDate <= toText Then
                total = total + ServicePrice(ServiceIdOfRow(table, rowIndex))
            End If
        End If
    Next rowIndex

    RevenueInPeriod = total
End Function

' Revenue of one day: the price of the service of every appointment of that day
' (the monthly grid shows it under the number of appointments).
Public Function RevenueOn(ByVal isoDate As String) As Double
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String
    Dim total As Double

    wanted = IsoDateText(isoDate)
    If Len(wanted) < ISO_DATE_LENGTH Then Exit Function

    Set table = AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        If StrComp(DateTextOfRow(table, rowIndex), wanted, vbTextCompare) = 0 Then
            total = total + ServicePrice(ServiceIdOfRow(table, rowIndex))
        End If
    Next rowIndex

    RevenueOn = total
End Function

' =====================================================================
'  text helpers (ISO text, whatever the cell holds)
' =====================================================================

Public Function IsoDateText(ByVal value As Variant) As String
    Dim text As String

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    If VarType(value) = vbString Then
        text = Trim$(CStr(value))
        If Len(text) >= ISO_DATE_LENGTH Then
            IsoDateText = Left$(text, ISO_DATE_LENGTH)
        Else
            IsoDateText = text
        End If
        Exit Function
    End If

    If IsDate(value) Then IsoDateText = Format$(CDate(value), "yyyy-mm-dd")
End Function

Public Function TimeTextValue(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    If VarType(value) = vbString Then
        TimeTextValue = Trim$(CStr(value))
        Exit Function
    End If

    If IsDate(value) Then TimeTextValue = Format$(CDate(value), "hh:nn")
End Function

Public Function DurationValue(ByVal value As Variant) As Long
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If Not IsNumeric(value) Then Exit Function

    DurationValue = CLng(Val(CStr(value)))
End Function

' =====================================================================
'  validation
' =====================================================================

' Problems of the three DB sheets (missing sheet, missing column), one per
' line; "" when the data is usable.
Public Function TableProblems() As String
    Dim problems As String

    problems = problems & MissingColumns(SHEET_DB_CLIENTS, _
        COL_ID & "," & COL_CLIENT_NAME & "," & COL_CLIENT_PHONE & "," & _
        COL_CLIENT_EMAIL & "," & COL_CLIENT_COMMENT)
    problems = problems & MissingColumns(SHEET_DB_SERVICES, _
        COL_ID & "," & COL_SERVICE_NAME & "," & COL_SERVICE_PRICE & "," & _
        COL_SERVICE_DURATION)
    problems = problems & MissingColumns(SHEET_DB_APPOINTMENTS, _
        COL_ID & "," & COL_CLIENT_ID & "," & COL_SERVICE_ID & "," & COL_DATE & "," & _
        COL_START & "," & COL_SERVICE_DURATION)

    TableProblems = problems
End Function

Private Function MissingColumns(ByVal sheetName As String, ByVal columnsText As String) As String
    Dim table As SqlTable
    Dim columns As Variant
    Dim item As Variant
    Dim problems As String

    If Not modNames.SheetExists(sheetName) Then
        MissingColumns = modText.TextFormat(modText.KEY_ERROR_NO_SHEET, sheetName) & vbCrLf
        Exit Function
    End If

    Set table = LoadTable(sheetName)
    columns = Split(columnsText, ",")
    For Each item In columns
        If Not table.HasColumn(Trim$(CStr(item))) Then
            problems = problems & sheetName & " has no column '" & Trim$(CStr(item)) & "'." & vbCrLf
        End If
    Next item

    MissingColumns = problems
End Function

' =====================================================================
'  writes (phase 2): only through SqlMutation, always checked
' =====================================================================

' Inserts one client. The id is max(id) + 1 unless the caller gave one.
Public Function AppendClient(ByRef values As Object) As Long
    Dim mutation As SqlMutation
    Dim newId As Long

    EnsureLoaded "AppendClient"
    RequireValues values, COL_CLIENT_NAME
    newId = WithId(values, SHEET_DB_CLIENTS)

    Set mutation = New SqlMutation
    Set mutation = mutation.AutoRange(SHEET_DB_CLIENTS)
    mutation.RequireInsertColumns COL_ID, COL_CLIENT_NAME, COL_CLIENT_PHONE, _
                                  COL_CLIENT_EMAIL, COL_CLIENT_COMMENT
    mutation.Insert values
    mutation.Commit

    LoadData
    If Not RowExists(SHEET_DB_CLIENTS, newId) Then
        Err.Raise 5, "modData.AppendClient", _
                  "the client " & CStr(newId) & " was committed but is not in the sheet."
    End If

    AppendClient = newId
End Function

' Inserts one service.
Public Function AppendService(ByRef values As Object) As Long
    Dim mutation As SqlMutation
    Dim newId As Long

    EnsureLoaded "AppendService"
    RequireValues values, COL_SERVICE_NAME
    newId = WithId(values, SHEET_DB_SERVICES)

    Set mutation = New SqlMutation
    Set mutation = mutation.AutoRange(SHEET_DB_SERVICES)
    mutation.RequireInsertColumns COL_ID, COL_SERVICE_NAME, COL_SERVICE_PRICE, _
                                  COL_SERVICE_DURATION
    mutation.Insert values
    mutation.Commit

    LoadData
    If Not RowExists(SHEET_DB_SERVICES, newId) Then
        Err.Raise 5, "modData.AppendService", _
                  "the service " & CStr(newId) & " was committed but is not in the sheet."
    End If

    AppendService = newId
End Function

' Deletes exactly one row. A missing id, or a delete that would touch more than
' one row, is rolled back instead of being committed (as in the POC).
Public Function DeleteRowById(ByVal sheetName As String, ByVal idValue As Long) As Boolean
    Dim mutation As SqlMutation
    Dim affected As Long

    EnsureLoaded "DeleteRowById"

    If idValue < 1 Then
        Err.Raise 5, "modData.DeleteRowById", _
                  modText.TextFormat(modText.KEY_ERROR_BAD_ID, CStr(idValue))
    End If
    If Not modNames.SheetExists(sheetName) Then
        Err.Raise 9, "modData.DeleteRowById", _
                  modText.TextFormat(modText.KEY_ERROR_NO_SHEET, sheetName)
    End If

    Set mutation = New SqlMutation
    Set mutation = mutation.AutoRange(sheetName)
    mutation.DeleteWhere COL_ID, "=", idValue

    affected = mutation.RowsAffected
    If affected <> 1 Then
        mutation.Rollback
        Err.Raise 5, "modData.DeleteRowById", _
                  modText.TextFormat(modText.KEY_ERROR_WRITE, _
                                     CStr(affected) & " rows matched, not 1")
    End If

    mutation.Commit
    LoadData

    DeleteRowById = True
End Function

' Inserts one appointment.
Public Function AppendAppointment(ByRef values As Object) As Long
    Dim mutation As SqlMutation
    Dim newId As Long

    EnsureLoaded "AppendAppointment"
    RequireValues values, COL_CLIENT_ID & "," & COL_SERVICE_ID & "," & COL_DATE & "," & _
                         COL_START & "," & COL_SERVICE_DURATION
    newId = WithId(values, SHEET_DB_APPOINTMENTS)

    Set mutation = New SqlMutation
    Set mutation = mutation.AutoRange(SHEET_DB_APPOINTMENTS)
    mutation.RequireInsertColumns COL_ID, COL_CLIENT_ID, COL_SERVICE_ID, COL_DATE, _
                                  COL_START, COL_SERVICE_DURATION
    mutation.Insert values
    mutation.Commit

    LoadData
    If Not RowExists(SHEET_DB_APPOINTMENTS, newId) Then
        Err.Raise 5, "modData.AppendAppointment", _
                  "the appointment " & CStr(newId) & " was committed but is not in the sheet."
    End If

    AppendAppointment = newId
End Function

' The text of one column of an appointment, looked up by its id ("" when the row
' is gone) - the dashboard, the tests and the log use it.
Public Function AppointmentCell(ByVal idValue As Long, ByVal columnName As String) As String
    Dim table As SqlTable
    Dim rowIndex As Long

    Set table = AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        If CLng(Val(CStr(IdOfRow(table, rowIndex)))) = idValue Then
            AppointmentCell = CellTextOf(table, rowIndex, columnName)
            Exit Function
        End If
    Next rowIndex
End Function

' The id of the appointment that starts at that time on that day (0 when there
' is none) - the dashboard and the tests use it.
Public Function AppointmentIdAt(ByVal isoDate As String, ByVal startText As String) As Long
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String

    wanted = Trim$(isoDate)
    If Len(wanted) < ISO_DATE_LENGTH Then Exit Function

    Set table = AppointmentsTable()
    For rowIndex = 1 To table.RowCount
        If StrComp(DateTextOfRow(table, rowIndex), wanted, vbTextCompare) = 0 Then
            If StrComp(StartTextOfRow(table, rowIndex), Trim$(startText), vbTextCompare) = 0 Then
                AppointmentIdAt = CLng(Val(CStr(IdOfRow(table, rowIndex))))
                Exit Function
            End If
        End If
    Next rowIndex
End Function

' Updates exactly one appointment by id (§8.2.10): more or fewer than one matched
' row is rolled back instead of being committed.
Public Function UpdateAppointment(ByVal idValue As Long, ByRef values As Object) As Boolean
    Dim mutation As SqlMutation
    Dim affected As Long

    EnsureLoaded "UpdateAppointment"

    If idValue < 1 Then
        Err.Raise 5, "modData.UpdateAppointment", _
                  modText.TextFormat(modText.KEY_ERROR_BAD_ID, CStr(idValue))
    End If

    Set mutation = New SqlMutation
    Set mutation = mutation.AutoRange(SHEET_DB_APPOINTMENTS)
    mutation.UpdateWhere COL_ID, "=", idValue, values

    affected = mutation.RowsAffected
    If affected <> 1 Then
        mutation.Rollback
        Err.Raise 5, "modData.UpdateAppointment", _
                  modText.TextFormat(modText.KEY_ERROR_WRITE, _
                                     CStr(affected) & " rows matched, not 1")
    End If

    mutation.Commit
    LoadData

    UpdateAppointment = True
End Function

' The stored values of one client, keyed by the column names the form works with
' (an unknown id answers an empty dictionary) - the [Edit] link of the list loads
' the form from here (§6.10).
Public Function ClientValues(ByVal idValue As Long) As Object
    Set ClientValues = RowValues(ClientsTable(), idValue, _
        COL_ID & "," & COL_CLIENT_NAME & "," & COL_CLIENT_PHONE & "," & _
        COL_CLIENT_EMAIL & "," & COL_CLIENT_COMMENT)
End Function

Public Function ServiceValues(ByVal idValue As Long) As Object
    Set ServiceValues = RowValues(ServicesTable(), idValue, _
        COL_ID & "," & COL_SERVICE_NAME & "," & COL_SERVICE_PRICE & "," & _
        COL_SERVICE_DURATION)
End Function

' The named columns of the row with this id, as text.
Private Function RowValues(ByRef table As SqlTable, ByVal idValue As Long, _
                           ByVal columnsText As String) As Object
    Dim values As Object
    Dim names As Variant
    Dim item As Variant
    Dim rowIndex As Long
    Dim nameText As String

    Set values = CreateObject("Scripting.Dictionary")
    Set RowValues = values
    If idValue < 1 Then Exit Function

    For rowIndex = 1 To table.RowCount
        If CLng(Val(CStr(IdOfRow(table, rowIndex)))) = idValue Then Exit For
    Next rowIndex
    If rowIndex > table.RowCount Then Exit Function

    names = Split(columnsText, ",")
    For Each item In names
        nameText = Trim$(CStr(item))
        If Len(nameText) > 0 Then values(nameText) = CellTextOf(table, rowIndex, nameText)
    Next item
End Function

' Updates exactly one client by id (§8.2.10): more or fewer than one matched row
' is rolled back instead of being committed.
Public Function UpdateClient(ByVal idValue As Long, ByRef values As Object) As Boolean
    UpdateClient = UpdateRowById(SHEET_DB_CLIENTS, idValue, values, "UpdateClient")
End Function

Public Function UpdateService(ByVal idValue As Long, ByRef values As Object) As Boolean
    UpdateService = UpdateRowById(SHEET_DB_SERVICES, idValue, values, "UpdateService")
End Function

Private Function UpdateRowById(ByVal sheetName As String, ByVal idValue As Long, _
                               ByRef values As Object, ByVal caller As String) As Boolean
    Dim mutation As SqlMutation
    Dim affected As Long

    EnsureLoaded caller

    If idValue < 1 Then
        Err.Raise 5, "modData." & caller, _
                  modText.TextFormat(modText.KEY_ERROR_BAD_ID, CStr(idValue))
    End If

    Set mutation = New SqlMutation
    Set mutation = mutation.AutoRange(sheetName)
    mutation.UpdateWhere COL_ID, "=", idValue, values

    affected = mutation.RowsAffected
    If affected <> 1 Then
        mutation.Rollback
        Err.Raise 5, "modData." & caller, _
                  modText.TextFormat(modText.KEY_ERROR_WRITE, _
                                     CStr(affected) & " rows matched, not 1")
    End If

    mutation.Commit
    LoadData

    UpdateRowById = True
End Function

' The rows of db_appointments, newest first: by date, then by start, then by id, all
' descending (the owner's rule, 25.09.2026). It answers the row indexes of the
' table, so a renderer visits them in that order. ISO dates and zero padded times
' sort correctly as text, so the comparison is a plain string compare.
Public Function AppointmentRowsNewestFirst() As Collection
    Dim table As SqlTable
    Dim order As Collection
    Dim keys As Collection
    Dim rowIndex As Long
    Dim nameCount As Long
    Dim place As Long
    Dim position As Long
    Dim keyText As String

    Set order = New Collection
    Set keys = New Collection
    Set table = AppointmentsTable()

    For rowIndex = 1 To table.RowCount
        keyText = AppointmentOrderKey(table, rowIndex)

        ' The place of this row: before the first key that is smaller than it.
        place = nameCount + 1
        For position = 1 To nameCount
            If keyText > CStr(keys(position)) Then
                place = position
                Exit For
            End If
        Next position

        If place > nameCount Then
            order.Add rowIndex
            keys.Add keyText
        Else
            order.Add rowIndex, , place
            keys.Add keyText, , place
        End If
        nameCount = nameCount + 1
    Next rowIndex

    Set AppointmentRowsNewestFirst = order
End Function

' "2026-09-24|14:15|3" - the sort key of one appointment row.
Private Function AppointmentOrderKey(ByRef table As SqlTable, ByVal rowIndex As Long) As String
    AppointmentOrderKey = DateTextOfRow(table, rowIndex) & "|" & _
                          StartTextOfRow(table, rowIndex) & "|" & _
                          CStr(IdOfRow(table, rowIndex))
End Function

' True when a row with this id exists in the sheet.
Public Function RowExists(ByVal sheetName As String, ByVal idValue As Variant) As Boolean
    RowExists = (FindRowByText(LoadTable(sheetName), COL_ID, idValue) > 0)
End Function

' How many appointments use this id of the given column (client_id, service_id).
Public Function CountAppointmentsFor(ByVal columnName As String, _
                                     ByVal idValue As Variant) As Long
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String
    Dim count As Long

    Set table = AppointmentsTable()
    If Not table.HasColumn(columnName) Then Exit Function

    wanted = IdText(idValue)
    If Len(wanted) = 0 Then Exit Function

    For rowIndex = 1 To table.RowCount
        If StrComp(IdText(table.GetCell(rowIndex, columnName)), wanted, vbTextCompare) = 0 Then
            count = count + 1
        End If
    Next rowIndex

    CountAppointmentsFor = count
End Function

' The text of one cell of a table, trimmed ("" when the cell is empty). The UI
' renders with it, so a NULL or a number never reaches a cell unprepared.
Public Function CellTextOf(ByRef table As SqlTable, ByVal rowIndex As Long, _
                           ByVal columnName As String) As String
    CellTextOf = TextValue(table.GetCell(rowIndex, columnName))
End Function

' =====================================================================
'  the dropdown lists on _Lists (§6.5.4)
' =====================================================================

' Fills list_clients and list_services from the tables and empties the rest of
' the ranges. The ranges are cleared, never deleted: deleting a cell would break
' its name (PROJECT_CONCEPT §13.18).
Public Sub RefreshLists()
    Dim table As SqlTable
    Dim target As Range
    Dim index As Long
    Dim written As Long
    Dim slot As Long

    If Not modNames.SheetExists(SHEET_LISTS) Then Exit Sub

    Set target = modNames.RangeOfName(NAME_LIST_CLIENTS)
    target.ClearContents
    Set table = ClientsTable()
    For index = 1 To table.RowCount
        written = written + 1
        If written <= target.Rows.Count Then
            target.Cells(written, 1).Value = TextValue(table.GetCell(index, COL_CLIENT_NAME))
        End If
    Next index

    written = 0
    Set target = modNames.RangeOfName(NAME_LIST_SERVICES)
    target.ClearContents
    Set table = ServicesTable()
    For index = 1 To table.RowCount
        written = written + 1
        If written <= target.Rows.Count Then
            target.Cells(written, 1).Value = TextValue(table.GetCell(index, COL_SERVICE_NAME))
        End If
    Next index

    ' The slots of the working day (Config), for the time fields of Booking.
    written = 0
    Set target = modNames.RangeOfName(NAME_LIST_TIMES)
    target.ClearContents
    For slot = modAppointments.DayStartMinutes() To modAppointments.DayEndMinutes() - 1 _
               Step modAppointments.StepMinutes()
        written = written + 1
        If written <= target.Rows.Count Then
            target.Cells(written, 1).Value = modApp.MinutesToTime(slot)
        End If
    Next slot
End Sub

' The id of the first client with this name; 0 when there is none. Two clients
' may share a name, so the rule is "first match wins" (documented, phase 3).
Public Function ClientIdByName(ByVal nameText As String) As Long
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String

    wanted = Trim$(nameText)
    If Len(wanted) = 0 Then Exit Function

    Set table = ClientsTable()
    For rowIndex = 1 To table.RowCount
        If StrComp(TextValue(table.GetCell(rowIndex, COL_CLIENT_NAME)), wanted, _
                   vbTextCompare) = 0 Then
            ClientIdByName = CLng(Val(IdText(table.GetCell(rowIndex, COL_ID))))
            Exit Function
        End If
    Next rowIndex
End Function

Public Function ServiceIdByName(ByVal nameText As String) As Long
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim wanted As String

    wanted = Trim$(nameText)
    If Len(wanted) = 0 Then Exit Function

    Set table = ServicesTable()
    For rowIndex = 1 To table.RowCount
        If StrComp(TextValue(table.GetCell(rowIndex, COL_SERVICE_NAME)), wanted, _
                   vbTextCompare) = 0 Then
            ServiceIdByName = CLng(Val(IdText(table.GetCell(rowIndex, COL_ID))))
            Exit Function
        End If
    Next rowIndex
End Function

' =====================================================================
'  internals
' =====================================================================

' The values of an insert must carry at least these columns.
Private Sub RequireValues(ByRef values As Object, ByVal requiredText As String)
    Dim required As Variant
    Dim item As Variant

    If values Is Nothing Then
        Err.Raise 5, "modData.RequireValues", "no values were given."
    End If

    required = Split(requiredText, ",")
    For Each item In required
        If Not values.Exists(Trim$(CStr(item))) Then
            Err.Raise 5, "modData.RequireValues", _
                      "the column '" & Trim$(CStr(item)) & "' is missing."
        End If
    Next item
End Sub

' Puts the id into the values: the given one, otherwise max(id) + 1.
Private Function WithId(ByRef values As Object, ByVal sheetName As String) As Long
    Dim newId As Long

    If values.Exists(COL_ID) Then
        If IsNumeric(values(COL_ID)) Then newId = CLng(Val(CStr(values(COL_ID))))
    End If
    If newId < 1 Then newId = NextId(LoadTable(sheetName))

    values(COL_ID) = newId
    WithId = newId
End Function

' Compares ids as normalized text, so a numeric 1 in one sheet and the text "1"
' in another are the same id.
Private Function FindRowByText(ByRef table As SqlTable, ByVal keyColumn As String, _
                               ByVal keyValue As Variant) As Long
    Dim key As String
    Dim rowIndex As Long

    If table Is Nothing Then Exit Function

    key = IdText(keyValue)
    If Len(key) = 0 Then Exit Function
    If Not table.HasColumn(keyColumn) Then Exit Function

    For rowIndex = 1 To table.RowCount
        If StrComp(IdText(table.GetCell(rowIndex, keyColumn)), key, vbTextCompare) = 0 Then
            FindRowByText = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function IdText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    If VarType(value) = vbString Then
        IdText = Trim$(CStr(value))
    ElseIf IsNumeric(value) Then
        IdText = Trim$(CStr(CDbl(value)))
    Else
        IdText = Trim$(CStr(value))
    End If
End Function

Private Function TextValue(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    TextValue = Trim$(CStr(value))
End Function
