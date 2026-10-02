Option Explicit

' =====================================================================
'  Service_Manager - the script model: a workbook in memory
'
'  What the model gives the rules is exactly what they use from Excel:
'  named cells, cells by row/column, Value/Text, colours, hyperlinks and the
'  Application switches. It is not an Excel emulator: ranges as objects,
'  formulas, selection and everything the user only sees are out of scope
'  (see tools\sim\README.md for the boundary and the trade offs).
'
'  Written with cscript in mind: no ByVal/ByRef, no types on Dim, no Optional,
'  no Format() - VBScript has none of them (a lesson of the project).
' =====================================================================

Class SimApp
    Public EnableEvents, ScreenUpdating, Interactive, DisplayAlerts

    Private Sub Class_Initialize()
        EnableEvents = True
        ScreenUpdating = True
        Interactive = True
        DisplayAlerts = True
    End Sub
End Class

Class SimCell
    Public Row, Column
    Public Value, Text
    Public Fill, FontColor, Bold, NumberFormat
    Public Linked, Link, Tip

    Private Sub Class_Initialize()
        Reset
    End Sub

    Public Sub Reset()
        Value = Empty
        Text = ""
        Fill = -1
        FontColor = -1
        Bold = False
        NumberFormat = ""
        Linked = False
        Link = ""
        Tip = ""
    End Sub

    ' Writing a value also sets what .Text shows: the rules read .Text.
    Public Sub Write(v)
        Reset
        Value = v
        Text = CStr(v)
    End Sub

    ' A cell button: caption + hyperlink, like modUi.WriteLinkCell writes it.
    Public Sub WriteLink(caption, subAddress, tipText)
        Write caption
        Linked = True
        Link = subAddress
        Tip = tipText
    End Sub
End Class

Class SimSheet
    Public Name
    Private mCells, mNames, mRows, mColumns

    Private Sub Class_Initialize()
        Set mCells = CreateObject("Scripting.Dictionary")
        Set mNames = CreateObject("Scripting.Dictionary")
        mRows = 0
        mColumns = 0
    End Sub

    Public Function Cells(r, c)
        Dim key, cell
        key = CStr(r) & ":" & CStr(c)
        If Not mCells.Exists(key) Then
            Set cell = New SimCell
            cell.Row = r
            cell.Column = c
            Set mCells(key) = cell
        End If
        If r > mRows Then mRows = r
        If c > mColumns Then mColumns = c
        Set Cells = mCells(key)
    End Function

    Public Sub DefineName(nameText, r, c)
        Set mNames(LCase(nameText)) = Cells(r, c)
    End Sub

    Public Function Named(nameText)
        Set Named = Nothing
        If mNames.Exists(LCase(nameText)) Then Set Named = mNames(LCase(nameText))
    End Function

    Public Function HasName(nameText)
        HasName = mNames.Exists(LCase(nameText))
    End Function

    Public Function NameCount()
        NameCount = mNames.Count
    End Function

    Public Function RowCount()
        RowCount = mRows
    End Function

    Public Function CellText(r, c)
        CellText = Cells(r, c).Text
    End Function
End Class

' A db_ sheet: header on row 1, data from row 2 - the layout SqlTable reads in
' Excel. In the model the rows are plain strings from sim-data.vbs (lookups are
' by header name, like SqlTable does with its header index).
Class SimTable
    Public Name
    Private mHeaders, mRows, mHeaderIndex

    Private Sub Class_Initialize()
        Set mHeaders = CreateObject("Scripting.Dictionary")
        Set mRows = CreateObject("Scripting.Dictionary")
        Set mHeaderIndex = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub Load(tableName, headerLine, dataLines)
        Dim i, columnName, parts

        Name = tableName
        Set mHeaders = CreateObject("Scripting.Dictionary")
        Set mRows = CreateObject("Scripting.Dictionary")
        Set mHeaderIndex = CreateObject("Scripting.Dictionary")

        parts = Split(headerLine, ",")
        For i = 0 To UBound(parts)
            columnName = Trim(parts(i))
            mHeaders(i + 1) = columnName
            mHeaderIndex(LCase(columnName)) = i + 1
        Next

        For i = 0 To UBound(dataLines)
            If Len(Trim(dataLines(i))) > 0 Then mRows(i + 1) = Split(dataLines(i), ",")
        Next
    End Sub

    Public Function RowCount()
        RowCount = mRows.Count
    End Function

    Public Function ColumnCount()
        ColumnCount = mHeaders.Count
    End Function

    Public Function RowNumber(headerName)
        RowNumber = 0
        If mHeaderIndex.Exists(LCase(headerName)) Then RowNumber = mHeaderIndex(LCase(headerName))
    End Function

    Public Function HeaderName(columnIndex)
        HeaderName = ""
        If mHeaders.Exists(columnIndex) Then HeaderName = mHeaders(columnIndex)
    End Function

    ' 1 based data rows, like SqlTable (the header row is not a data row).
    Public Function ValueAt(rowIndex, headerName)
        Dim columnIndex, fields

        ValueAt = Empty
        columnIndex = RowNumber(headerName)
        If columnIndex = 0 Then Exit Function
        If Not mRows.Exists(rowIndex) Then Exit Function

        fields = mRows(rowIndex)
        If columnIndex - 1 <= UBound(fields) Then ValueAt = Trim(fields(columnIndex - 1))
    End Function

    Public Function TextAt(rowIndex, headerName)
        TextAt = CStr(ValueAt(rowIndex, headerName))
    End Function

    ' ---- what a mutation does to the rows -------------------------------
    ' Deleting a row moves the rows below it up: Excel deletes the row, it does
    ' not leave a hole. The model must do the same, otherwise every later row
    ' index would mean something else than in the workbook.

    ' Appends a row: fields() holds one text per column, in the order of the
    ' header. Answers the new row index.
    Public Function AddRow(fields)
        mRows(mRows.Count + 1) = fields
        AddRow = mRows.Count
    End Function

    Public Sub SetCell(rowIndex, headerName, value)
        Dim columnIndex, fields

        columnIndex = RowNumber(headerName)
        If columnIndex = 0 Then Exit Sub
        If Not mRows.Exists(rowIndex) Then Exit Sub

        fields = mRows(rowIndex)
        fields(columnIndex - 1) = value
        mRows(rowIndex) = fields
    End Sub

    ' Answers 1 when the row was there (the VBA twin counts the rows a mutation
    ' matched and rolls back unless exactly one did).
    Public Function RemoveRow(rowIndex)
        Dim index, list

        If Not mRows.Exists(rowIndex) Then Exit Function

        Set list = CreateObject("Scripting.Dictionary")
        For index = 1 To mRows.Count
            If index <> rowIndex Then list(list.Count + 1) = mRows(index)
        Next
        Set mRows = list

        RemoveRow = 1
    End Function
End Class

Class SimBook
    Public Application
    Private mSheets, mNames

    Private Sub Class_Initialize()
        Set Application = New SimApp
        Set mSheets = CreateObject("Scripting.Dictionary")
        Set mNames = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub AddSheet(sheetName)
        Dim sheet
        Set sheet = New SimSheet
        sheet.Name = sheetName
        Set mSheets(LCase(sheetName)) = sheet
    End Sub

    Public Function SheetExists(sheetName)
        SheetExists = mSheets.Exists(LCase(sheetName))
    End Function

    Public Function Sheet(sheetName)
        Set Sheet = mSheets(LCase(sheetName))
    End Function

    Public Sub DefineName(nameText, sheetName, r, c)
        Dim sheet
        Set sheet = mSheets(LCase(sheetName))
        sheet.DefineName nameText, r, c
        Set mNames(LCase(nameText)) = sheet.Named(nameText)
    End Sub

    Public Function RangeOfName(nameText)
        Set RangeOfName = Nothing
        If mNames.Exists(LCase(nameText)) Then Set RangeOfName = mNames(LCase(nameText))
    End Function

    Public Function ReadName(nameText)
        Dim cell
        Set cell = RangeOfName(nameText)
        If cell Is Nothing Then
            ReadName = ""
        Else
            ReadName = cell.Text
        End If
    End Function

    Public Sub WriteName(nameText, v)
        Dim cell
        Set cell = RangeOfName(nameText)
        If Not cell Is Nothing Then cell.Write v
    End Sub
End Class
