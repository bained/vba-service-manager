Attribute VB_Name = "modUiClients"
Option Explicit

' =====================================================================
'  Service_Manager - the clients page (Phase 2)
'
'  The boundary between the cells and the domain: it reads the form, draws the
'  list and turns a click inside the list into an action. The rows of the list
'  have no names - a rendered cell cannot have one - so the click is recognized
'  by position (§6.10) and the hyperlink of every row points at the grid anchor.
' =====================================================================

' clients.add: the payload wins (a test or a script), otherwise the form. While the
' form is editing a row (the state cell holds its id) the same button updates that
' row instead of adding a second one (§6.10).
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
        modClients.Update values, result
    Else
        modClients.Add values, result
    End If

    If result(modRoutes.RESULT_OK) = True Then ClearForm
End Sub

' clients.edit: the [Edit] link of a row loads that client into the form above and
' turns the button into "Update client" (§6.10).
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

    Set values = modData.ClientValues(idValue)
    If values.Count = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_NOT_FOUND, _
                                                             "Client", CStr(idValue))
        Exit Sub
    End If

    modNames.WriteName NAME_CLIENT_NAME, CStr(values(COL_CLIENT_NAME))
    modNames.WriteName NAME_CLIENT_PHONE, CStr(values(COL_CLIENT_PHONE))
    modNames.WriteName NAME_CLIENT_EMAIL, CStr(values(COL_CLIENT_EMAIL))
    modNames.WriteName NAME_CLIENT_COMMENT, CStr(values(COL_CLIENT_COMMENT))
    modNames.WriteName NAME_STATE_EDIT_CLIENT, CStr(idValue)

    Render

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_FORM_LOADED, _
                                                          "Client", CStr(idValue))
    result(modRoutes.RESULT_OK) = True
End Sub

' clients.list: draws the page and says how many rows it wrote.
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

' The values of the form, keyed by column name.
Public Function ReadForm() As Object
    Dim values As Object

    Set values = CreateObject("Scripting.Dictionary")
    values(COL_CLIENT_NAME) = modNames.ReadName(NAME_CLIENT_NAME)
    values(COL_CLIENT_PHONE) = modNames.ReadName(NAME_CLIENT_PHONE)
    values(COL_CLIENT_EMAIL) = modNames.ReadName(NAME_CLIENT_EMAIL)
    values(COL_CLIENT_COMMENT) = modNames.ReadName(NAME_CLIENT_COMMENT)

    Set ReadForm = values
End Function

Public Sub ClearForm()
    modNames.WriteName NAME_CLIENT_NAME, vbNullString
    modNames.WriteName NAME_CLIENT_PHONE, vbNullString
    modNames.WriteName NAME_CLIENT_EMAIL, vbNullString
    modNames.WriteName NAME_CLIENT_COMMENT, vbNullString
    modNames.WriteName NAME_STATE_EDIT_CLIENT, vbNullString
    RenderMode
End Sub

' The id of the client the form is editing; 0 when it is a new one.
Public Function EditingId() As Long
    EditingId = CLng(Val(modNames.ReadName(NAME_STATE_EDIT_CLIENT)))
End Function

' The caption of the button follows the mode: "Add client" for a new one, "Update
' client" while a row is being edited (§6.10).
Private Sub RenderMode()
    If EditingId() > 0 Then
        modNames.WriteName NAME_BTN_CLIENTS_ADD, modText.Text(modText.KEY_CLIENTS_BTN_UPDATE)
    Else
        modNames.WriteName NAME_BTN_CLIENTS_ADD, modText.Text(modText.KEY_CLIENTS_BTN_ADD)
    End If
End Sub

' True when the payload already carries the values of a client.
Private Function HasFormValues(ByRef data As Object) As Boolean
    Dim values As Object

    Set values = modApp.PayloadFrom(data)
    HasFormValues = values.Exists(COL_CLIENT_NAME)
End Function

' =====================================================================
'  the list
' =====================================================================

' Draws the title and every row; returns how many rows the table holds.
Public Function Render() As Long
    Dim firstCell As Range
    Dim lastCell As Range
    Dim capacity As Long
    Dim table As SqlTable
    Dim dataRow As Long
    Dim sheetRow As Long
    Dim count As Long
    Dim titleText As String

    If Not modNames.SheetExists(SHEET_CLIENTS_ADMIN) Then Exit Function

    Set firstCell = modNames.RangeOfName(NAME_GRID_CLIENTS_FIRST)
    Set lastCell = modNames.RangeOfName(NAME_GRID_CLIENTS_LAST)
    capacity = lastCell.Row - firstCell.Row + 1

    modUi.ClearBlock NAME_GRID_CLIENTS_FIRST, NAME_GRID_CLIENTS_LAST
    ClearDeleteColumn firstCell.Row, lastCell.Row
    If Not modData.IsModelLoaded() Then modData.LoadData
    Set table = modData.ClientsTable()
    count = table.RowCount

    For dataRow = 1 To count
        If dataRow > capacity Then Exit For
        sheetRow = firstCell.Row + dataRow - 1
        WriteRow sheetRow, dataRow, table
    Next dataRow

    titleText = modText.TextFormat(modText.KEY_CLIENTS_TITLE, CStr(count))
    If count > capacity Then
        titleText = modText.TextFormat(modText.KEY_LIST_CAPPED, titleText, CStr(capacity))
    End If
    If EditingId() > 0 Then
        titleText = titleText & " - " & modText.TextFormat(modText.KEY_FORM_MODE_EDIT, _
                                                           CStr(EditingId()))
    End If
    modNames.WriteName NAME_CLIENTS_TITLE, titleText

    RenderMode

    If count = 0 Then
        modUi.WriteCell SHEET_CLIENTS_ADMIN, firstCell.Row, CLIENTS_COL_ID, _
                        modText.Text(modText.KEY_CLIENTS_NONE)
    End If

    Render = count
End Function

Private Sub WriteRow(ByVal sheetRow As Long, ByVal dataRow As Long, ByRef table As SqlTable)
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS_ADMIN)

    modUi.WriteCell SHEET_CLIENTS_ADMIN, sheetRow, CLIENTS_COL_ID, _
                    modData.CellTextOf(table, dataRow, COL_ID)
    modUi.WriteCell SHEET_CLIENTS_ADMIN, sheetRow, CLIENTS_COL_NAME, _
                    modData.CellTextOf(table, dataRow, COL_CLIENT_NAME)
    ' The phone is text. A cell formatted as General turns 0888123456 into a number
    ' and the leading zero is gone (24.09.2026), so the format is set here as well as
    ' in the template: a restyled sheet cannot lose it again.
    ws.Cells(sheetRow, CLIENTS_COL_PHONE).NumberFormat = "@"
    modUi.WriteCell SHEET_CLIENTS_ADMIN, sheetRow, CLIENTS_COL_PHONE, _
                    modData.CellTextOf(table, dataRow, COL_CLIENT_PHONE)
    modUi.WriteCell SHEET_CLIENTS_ADMIN, sheetRow, CLIENTS_COL_EMAIL, _
                    modData.CellTextOf(table, dataRow, COL_CLIENT_EMAIL)
    modUi.WriteCell SHEET_CLIENTS_ADMIN, sheetRow, CLIENTS_COL_COMMENT, _
                    modData.CellTextOf(table, dataRow, COL_CLIENT_COMMENT)

    ' Two link cells per row: [Edit] loads the row into the form above, [Delete]
    ' removes it. A cell carries one hyperlink, so they cannot share a column and
    ' the second one sits next to the grid (§6.10).
    modUi.WriteLinkCell ws.Cells(sheetRow, CLIENTS_COL_ACTION), NAME_GRID_CLIENTS_FIRST, _
                        LIST_EDIT_TEXT, "Load this client into the form"
    modUi.WriteLinkCell ws.Cells(sheetRow, CLIENTS_COL_DELETE), NAME_GRID_CLIENTS_FIRST, _
                        LIST_ACTION_TEXT, "Delete this client"
End Sub

' The [Delete] column lives outside the named grid (A..F), so ClearBlock does not
' reach it: the links of the rows that are gone would stay clickable.
Private Sub ClearDeleteColumn(ByVal firstRow As Long, ByVal lastRow As Long)
    Dim ws As Worksheet
    Dim columnRange As Range

    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS_ADMIN)
    Set columnRange = ws.Range(ws.Cells(firstRow, CLIENTS_COL_DELETE), _
                              ws.Cells(lastRow, CLIENTS_COL_DELETE))

    columnRange.Hyperlinks.Delete
    columnRange.ClearContents
End Sub

' =====================================================================
'  a click inside the list
' =====================================================================

' The action a click on this cell means, or Nothing when it means nothing. The
' payload carries the action, the sheet and the cell address, so the router and
' the confirmation painting know where they are.
Public Function ClickPayload(ByRef cell As Range) As Object
    Dim payload As Object
    Dim sheetRow As Long
    Dim idValue As Variant
    Dim actionText As String

    Set ClickPayload = Nothing
    If cell Is Nothing Then Exit Function

    ' The column says what the click means: [Edit] in the action column, [Delete]
    ' in the one next to the grid (§6.10).
    If cell.Column = CLIENTS_COL_ACTION Then
        actionText = modRoutes.ACTION_CLIENTS_EDIT
    ElseIf cell.Column = CLIENTS_COL_DELETE Then
        actionText = modRoutes.ACTION_CLIENTS_DELETE
    Else
        Exit Function
    End If

    sheetRow = cell.Row
    If sheetRow < modNames.RangeOfName(NAME_GRID_CLIENTS_FIRST).Row Then Exit Function
    If sheetRow > modNames.RangeOfName(NAME_GRID_CLIENTS_LAST).Row Then Exit Function

    idValue = ThisWorkbook.Worksheets(SHEET_CLIENTS_ADMIN).Cells(sheetRow, CLIENTS_COL_ID).Value
    If Len(Trim$(CStr(idValue))) = 0 Then Exit Function

    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_ACTION) = actionText
    payload(modRoutes.KEY_ID) = CStr(idValue)
    payload(modRoutes.KEY_SHEET) = SHEET_CLIENTS_ADMIN
    payload(modRoutes.KEY_CELL) = cell.Address(False, False)

    Set ClickPayload = payload
End Function
