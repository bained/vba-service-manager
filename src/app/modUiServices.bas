Attribute VB_Name = "modUiServices"
Option Explicit

' =====================================================================
'  Service_Manager - the services page (Phase 2)
'
'  The same shape as modUiClients: the form, the list (with a [Delete] button on
'  every row) and the click of a row, recognized by position (§6.10).
' =====================================================================

' services.add: the payload wins (a test or a script), otherwise the form. While the
' form is editing a row the same button updates it instead of adding a second one.
Public Sub OnAdd(ByRef data As Object, ByRef result As Object)
    Dim values As Object
    Dim idValue As Long

    If HasFormValues(data) Then
        Set values = data
    Else
        Set values = ReadForm()
    End If

    idValue = EditingId()
    If idValue > 0 Then
        If Not values.Exists(modRoutes.KEY_ID) Then values(modRoutes.KEY_ID) = CStr(idValue)
        modServices.Update values, result
    Else
        modServices.Add values, result
    End If

    If result(modRoutes.RESULT_OK) = True Then ClearForm
End Sub

' services.edit: the [Edit] link of a row loads that service into the form above.
Public Sub OnEdit(ByRef data As Object, ByRef result As Object)
    Dim idText As String
    Dim idValue As Long
    Dim values As Object

    idText = modApp.PayloadText(data, modRoutes.KEY_ID)
    idValue = modApp.PayloadNumber(data, modRoutes.KEY_ID)
    If idValue < 1 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_BAD_ID, idText)
        Exit Sub
    End If

    Set values = modData.ServiceValues(idValue)
    If values.Count = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_NOT_FOUND, _
                                                             "Service", CStr(idValue))
        Exit Sub
    End If

    modNames.WriteName NAME_SERVICE_NAME, CStr(values(COL_SERVICE_NAME))
    modNames.WriteName NAME_SERVICE_PRICE, CStr(values(COL_SERVICE_PRICE))
    modNames.WriteName NAME_SERVICE_DURATION, CStr(values(COL_SERVICE_DURATION))
    modNames.WriteName NAME_STATE_EDIT_SERVICE, CStr(idValue)

    Render

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_FORM_LOADED, _
                                                          "Service", CStr(idValue))
    result(modRoutes.RESULT_OK) = True
End Sub

' services.list: draws the page and says how many rows it wrote.
Public Sub OnList(ByRef data As Object, ByRef result As Object)
    Dim rows As Long
    Dim info As Object

    rows = Render

    Set info = CreateObject("Scripting.Dictionary")
    info("rows") = CStr(rows)
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_OK_LIST, CStr(rows))
    result(modRoutes.RESULT_OK) = True
End Sub

' =====================================================================
'  the form
' =====================================================================

Public Function ReadForm() As Object
    Dim values As Object

    Set values = CreateObject("Scripting.Dictionary")
    values(COL_SERVICE_NAME) = modNames.ReadName(NAME_SERVICE_NAME)
    values(COL_SERVICE_PRICE) = modNames.ReadName(NAME_SERVICE_PRICE)
    values(COL_SERVICE_DURATION) = modNames.ReadName(NAME_SERVICE_DURATION)

    Set ReadForm = values
End Function

Public Sub ClearForm()
    modNames.WriteName NAME_SERVICE_NAME, vbNullString
    modNames.WriteName NAME_SERVICE_PRICE, vbNullString
    modNames.WriteName NAME_SERVICE_DURATION, vbNullString
    modNames.WriteName NAME_STATE_EDIT_SERVICE, vbNullString
    RenderMode
End Sub

' The id of the service the form is editing; 0 when it is a new one.
Public Function EditingId() As Long
    EditingId = CLng(Val(modNames.ReadName(NAME_STATE_EDIT_SERVICE)))
End Function

Private Sub RenderMode()
    If EditingId() > 0 Then
        modNames.WriteName NAME_BTN_SERVICES_ADD, modText.Text(modText.KEY_SERVICES_BTN_UPDATE)
    Else
        modNames.WriteName NAME_BTN_SERVICES_ADD, modText.Text(modText.KEY_SERVICES_BTN_ADD)
    End If
End Sub

Private Function HasFormValues(ByRef data As Object) As Boolean
    Dim values As Object

    Set values = modApp.PayloadFrom(data)
    HasFormValues = values.Exists(COL_SERVICE_NAME)
End Function

' =====================================================================
'  the list
' =====================================================================

Public Function Render() As Long
    Dim firstCell As Range
    Dim lastCell As Range
    Dim capacity As Long
    Dim table As SqlTable
    Dim dataRow As Long
    Dim sheetRow As Long
    Dim count As Long
    Dim titleText As String

    If Not modNames.SheetExists(SHEET_SERVICES_ADMIN) Then Exit Function

    Set firstCell = modNames.RangeOfName(NAME_GRID_SERVICES_FIRST)
    Set lastCell = modNames.RangeOfName(NAME_GRID_SERVICES_LAST)
    capacity = lastCell.Row - firstCell.Row + 1

    modUi.ClearBlock NAME_GRID_SERVICES_FIRST, NAME_GRID_SERVICES_LAST
    ClearDeleteColumn firstCell.Row, lastCell.Row
    If Not modData.IsModelLoaded() Then modData.LoadData
    Set table = modData.ServicesTable()
    count = table.RowCount

    For dataRow = 1 To count
        If dataRow > capacity Then Exit For
        sheetRow = firstCell.Row + dataRow - 1
        WriteRow sheetRow, dataRow, table
    Next dataRow

    titleText = modText.TextFormat(modText.KEY_SERVICES_TITLE, CStr(count))
    If count > capacity Then
        titleText = modText.TextFormat(modText.KEY_LIST_CAPPED, titleText, CStr(capacity))
    End If
    If EditingId() > 0 Then
        titleText = titleText & " - " & modText.TextFormat(modText.KEY_FORM_MODE_EDIT, _
                                                           CStr(EditingId()))
    End If
    modNames.WriteName NAME_SERVICES_TITLE, titleText

    RenderMode

    If count = 0 Then
        modUi.WriteCell SHEET_SERVICES_ADMIN, firstCell.Row, SERVICES_COL_ID, _
                        modText.Text(modText.KEY_SERVICES_NONE)
    End If

    Render = count
End Function

Private Sub WriteRow(ByVal sheetRow As Long, ByVal dataRow As Long, ByRef table As SqlTable)
    Dim ws As Worksheet
    Dim priceValue As Variant
    Dim durationValue As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_SERVICES_ADMIN)

    priceValue = table.GetCell(dataRow, COL_SERVICE_PRICE)
    durationValue = table.GetCell(dataRow, COL_SERVICE_DURATION)

    modUi.WriteCell SHEET_SERVICES_ADMIN, sheetRow, SERVICES_COL_ID, _
                    modData.CellTextOf(table, dataRow, COL_ID)
    modUi.WriteCell SHEET_SERVICES_ADMIN, sheetRow, SERVICES_COL_NAME, _
                    modData.CellTextOf(table, dataRow, COL_SERVICE_NAME)

    ' Price and duration are numbers in the sheet: the domain formats them.
    modUi.WriteCell SHEET_SERVICES_ADMIN, sheetRow, SERVICES_COL_PRICE, _
                    modApp.NumberValue(priceValue, 0)
    modUi.WriteCell SHEET_SERVICES_ADMIN, sheetRow, SERVICES_COL_DURATION, _
                    modApp.NumberValue(durationValue, DEFAULT_DURATION_MINUTES)

    ' [Edit] and [Delete] in two cells: a cell carries one hyperlink (§6.10).
    modUi.WriteLinkCell ws.Cells(sheetRow, SERVICES_COL_ACTION), NAME_GRID_SERVICES_FIRST, _
                        LIST_EDIT_TEXT, "Load this service into the form"
    modUi.WriteLinkCell ws.Cells(sheetRow, SERVICES_COL_DELETE), NAME_GRID_SERVICES_FIRST, _
                        LIST_ACTION_TEXT, "Delete this service"
End Sub

' The [Delete] column is outside the named grid, so ClearBlock does not reach it.
Private Sub ClearDeleteColumn(ByVal firstRow As Long, ByVal lastRow As Long)
    Dim ws As Worksheet
    Dim columnRange As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_SERVICES_ADMIN)
    Set columnRange = ws.Range(ws.Cells(firstRow, SERVICES_COL_DELETE), _
                              ws.Cells(lastRow, SERVICES_COL_DELETE))

    columnRange.Hyperlinks.Delete
    columnRange.ClearContents
End Sub

' =====================================================================
'  a click inside the list
' =====================================================================

Public Function ClickPayload(ByRef cell As Range) As Object
    Dim payload As Object
    Dim sheetRow As Long
    Dim idValue As Variant
    Dim actionText As String

    Set ClickPayload = Nothing
    If cell Is Nothing Then Exit Function

    If cell.Column = SERVICES_COL_ACTION Then
        actionText = modRoutes.ACTION_SERVICES_EDIT
    ElseIf cell.Column = SERVICES_COL_DELETE Then
        actionText = modRoutes.ACTION_SERVICES_DELETE
    Else
        Exit Function
    End If

    sheetRow = cell.Row
    If sheetRow < modNames.RangeOfName(NAME_GRID_SERVICES_FIRST).Row Then Exit Function
    If sheetRow > modNames.RangeOfName(NAME_GRID_SERVICES_LAST).Row Then Exit Function

    idValue = ThisWorkbook.Worksheets(SHEET_SERVICES_ADMIN).Cells(sheetRow, SERVICES_COL_ID).Value
    If Len(Trim$(CStr(idValue))) = 0 Then Exit Function

    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_ACTION) = actionText
    payload(modRoutes.KEY_ID) = CStr(idValue)
    payload(modRoutes.KEY_SHEET) = SHEET_SERVICES_ADMIN
    payload(modRoutes.KEY_CELL) = cell.Address(False, False)

    Set ClickPayload = payload
End Function
