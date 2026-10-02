Attribute VB_Name = "modLog"
Option Explicit

' =====================================================================
'  Service_Manager - the action log (Phase 1)
'
'  One row per action on the Log sheet: timestamp, action, ok, message. The
'  sheet keeps the last LOG_MAX_ROWS rows, because Ctrl+Z does not undo what
'  the application writes: the log is the only trace of what happened.
'
'  This module (with modUi) writes cells that are not named: the log rows are
'  created while the application runs.
' =====================================================================

' Appends one row.
Public Sub WriteRow(ByVal action As String, ByVal ok As Boolean, ByVal message As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim eventsWereOn As Boolean

    ' The log must never break an action: the first step that can fail is inside
    ' the handler already.
    On Error GoTo Failed

    If Not modNames.SheetExists(SHEET_LOG) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    nextRow = LastRow(ws) + 1

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False
    ws.Cells(nextRow, 1).NumberFormat = "@"
    ws.Cells(nextRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = action
    ws.Cells(nextRow, 3).Value = YesNoText(ok)
    ws.Cells(nextRow, 4).Value = message
    Application.EnableEvents = eventsWereOn
    GoTo Cleanup

Failed:
    ' The log must never break an action; the previous event state comes back.
    On Error Resume Next
    Application.EnableEvents = eventsWereOn
    Err.Clear
    Exit Sub

Cleanup:
    TrimTo LOG_MAX_ROWS
End Sub

' How many rows the log holds (without the header row).
Public Function RowCount() As Long
    Dim ws As Worksheet

    If Not modNames.SheetExists(SHEET_LOG) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)

    RowCount = LastRow(ws) - FirstRow() + 1
    If RowCount < 0 Then RowCount = 0
End Function

' The newest rows as text lines, newest first (diagnostics and tests).
Public Function Tail(ByVal lineCount As Long) As Collection
    Dim lines As Collection
    Dim ws As Worksheet
    Dim newestRow As Long
    Dim rowIndex As Long
    Dim count As Long

    Set lines = New Collection
    If Not modNames.SheetExists(SHEET_LOG) Then
        Set Tail = lines
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    newestRow = LastRow(ws)
    rowIndex = newestRow

    Do While rowIndex >= FirstRow() And count < lineCount
        lines.Add ws.Cells(rowIndex, 1).Text & " | " & ws.Cells(rowIndex, 2).Text & " | " & _
                  ws.Cells(rowIndex, 3).Text & " | " & ws.Cells(rowIndex, 4).Text
        rowIndex = rowIndex - 1
        count = count + 1
    Loop

    Set Tail = lines
End Function

' Keeps the newest maxRows rows. The rows are moved up and the tail is cleared,
' the rows themselves are NOT deleted: deleting a row that a defined name points
' at turns the name into #REF! (Excel), and then every later lookup of "log_first"
' fails with 1004. That cost one evening; see PROJECT_CONCEPT §13.
Public Sub TrimTo(ByVal maxRows As Long)
    Dim ws As Worksheet
    Dim count As Long
    Dim excess As Long
    Dim rowIndex As Long
    Dim target As Long
    Dim eventsWereOn As Boolean

    If maxRows < 1 Then Exit Sub
    If Not modNames.SheetExists(SHEET_LOG) Then Exit Sub

    count = RowCount()
    excess = count - maxRows
    If excess < 1 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done

    ' Newest rows (starting after the rows that fall away) are copied upwards.
    For rowIndex = FirstRow() + excess To FirstRow() + count - 1
        target = rowIndex - excess
        ws.Cells(target, 1).NumberFormat = "@"
        ws.Cells(target, 1).Value = ws.Cells(rowIndex, 1).Value
        ws.Cells(target, 2).Value = ws.Cells(rowIndex, 2).Value
        ws.Cells(target, 3).Value = ws.Cells(rowIndex, 3).Value
        ws.Cells(target, 4).Value = ws.Cells(rowIndex, 4).Value
    Next rowIndex

    ' The rows that are left over are emptied.
    ws.Range(ws.Cells(FirstRow() + maxRows, 1), _
             ws.Cells(FirstRow() + count - 1, 4)).Clear

Done:
    On Error Resume Next
    Application.EnableEvents = eventsWereOn
    Err.Clear
End Sub

' Empties the log (the tests use it, so the report stays readable).
Public Sub Clear()
    Dim ws As Worksheet
    Dim eventsWereOn As Boolean

    If Not modNames.SheetExists(SHEET_LOG) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    ws.Range(ws.Cells(FirstRow(), 1), ws.Cells(FirstRow() + LOG_MAX_ROWS, 4)).Clear
    On Error GoTo 0
    Application.EnableEvents = eventsWereOn
End Sub

' =====================================================================
'  internals
' =====================================================================

Private Function FirstRow() As Long
    Dim cell As Range

    ' The anchor name is the clean way; when it cannot be resolved the first data
    ' row of the log is used, so a broken name can never stop the logging.
    Set cell = Nothing
    On Error Resume Next
    Set cell = modNames.RangeOfName(NAME_LOG_FIRST)
    On Error GoTo 0

    If cell Is Nothing Then
        If modNames.SheetExists(SHEET_LOG) Then
            Set cell = ThisWorkbook.Worksheets(SHEET_LOG).Cells(DB_FIRST_ROW, 1)
        End If
    End If

    If cell Is Nothing Then
        FirstRow = DB_FIRST_ROW
    Else
        FirstRow = cell.Row
    End If
End Function

Private Function LastRow(ByRef ws As Worksheet) As Long
    Dim last As Long

    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If last < FirstRow() - 1 Then last = FirstRow() - 1

    LastRow = last
End Function

Private Function YesNoText(ByVal value As Boolean) As String
    If value Then
        YesNoText = "yes"
    Else
        YesNoText = "no"
    End If
End Function
