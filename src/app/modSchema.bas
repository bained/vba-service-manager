Attribute VB_Name = "modSchema"
Option Explicit

' =====================================================================
'  Service_Manager - the version of the schema (Phase 6, §8.3)
'
'  The workbook carries the code and the data at the same time, so it outlives
'  releases: a file written by an older version of the application has to keep
'  working. Config keeps a number for that (s_schema_version):
'
'    * an older file is migrated before it is read - the columns a version added
'      are created (empty, so the defaults of Config fill the meaning in), a copy
'      of the workbook is written and the Log says what was done;
'    * a newer file is refused with a sentence, instead of writing into columns
'      that mean something else in that version.
'
'  Every step is written so that running it a second time changes nothing: that is
'  what makes an interrupted migration harmless.
' =====================================================================

' What this build expects. Raise it when a table, a column or the meaning of one
' changes, and add the step to Migrate().
Public Const CURRENT_SCHEMA_VERSION As Long = 2

' What is wrong with the file, "" when it is fine for this build. modApp.Refresh
' calls it before the tables are read, so a migration happens before the first
' read and a newer file is named instead of being damaged.
Public Function Problem() As String
    Dim version As Long

    version = FileVersion()

    If version > CURRENT_SCHEMA_VERSION Then
        Problem = modText.TextFormat(modText.KEY_ERROR_SCHEMA_NEWER, _
                                     CStr(version), CStr(CURRENT_SCHEMA_VERSION))
        Exit Function
    End If

    If version < CURRENT_SCHEMA_VERSION Then Problem = Migrate(version)
End Function

' The number the file carries. 0 means the file is older than the version cell
' itself (the templates of phase 1 kept no number).
Public Function FileVersion() As Long
    FileVersion = CLng(Val(modNames.ReadName(NAME_STATE_SCHEMA)))
End Function

' Walks the steps between the number of the file and the number of the code. A
' copy comes first (a migration that fails half way must not cost data) and the
' Log keeps the record. "" when every step worked.
Private Function Migrate(ByVal fromVersion As Long) As String
    Dim step As Long

    If Not modBackup.SaveCopy("schema-" & CStr(fromVersion)) Then
        Migrate = modText.Text(modText.KEY_ERROR_BACKUP)
        Exit Function
    End If
    modLog.WriteRow "schema.migrate", True, "The file is schema " & CStr(fromVersion) & _
                    ", this build expects " & CStr(CURRENT_SCHEMA_VERSION) & "."

    For step = fromVersion + 1 To CURRENT_SCHEMA_VERSION
        Select Case step
            Case 2
                Migrate = StepTo2()
            Case Else
                Migrate = "there is no step to schema " & CStr(step)
        End Select
        If Len(Migrate) > 0 Then Exit Function
        modNames.WriteName NAME_STATE_SCHEMA, CStr(step)
    Next step
End Function

' Schema 2 is the layout of this build: every table carries the columns the
' application writes, and date and start are text (Excel must not read "09:00" as
' a time of the day). The columns are created empty on purpose - Config holds the
' default a missing duration or price means (§16.1).
Private Function StepTo2() As String
    Dim problems As String

    problems = EnsureColumns(SHEET_DB_CLIENTS, COL_ID & "," & COL_CLIENT_NAME & "," & _
                             COL_CLIENT_PHONE & "," & COL_CLIENT_EMAIL & "," & COL_CLIENT_COMMENT)
    problems = problems & EnsureColumns(SHEET_DB_SERVICES, COL_ID & "," & COL_SERVICE_NAME & _
                             "," & COL_SERVICE_PRICE & "," & COL_SERVICE_DURATION)
    problems = problems & EnsureColumns(SHEET_DB_APPOINTMENTS, COL_ID & "," & COL_CLIENT_ID & _
                             "," & COL_SERVICE_ID & "," & COL_DATE & "," & COL_START & "," & _
                             COL_SERVICE_DURATION)

    EnsureTextColumn SHEET_DB_APPOINTMENTS, COL_DATE
    EnsureTextColumn SHEET_DB_APPOINTMENTS, COL_START

    StepTo2 = problems
End Function

' The header of a table, the way the application reads it: the columns of the list
' must be there (the rest is the business of the user). A missing column is added
' at the right end, so the data of the older file stays where it is; "" is answered
' when everything was there already.
Private Function EnsureColumns(ByVal sheetName As String, ByVal columnsText As String) As String
    Dim ws As Worksheet
    Dim expected As Variant
    Dim item As Variant
    Dim lastColumn As Long

    If Not modNames.SheetExists(sheetName) Then
        EnsureColumns = modText.TextFormat(modText.KEY_ERROR_NO_SHEET, sheetName) & vbCrLf
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(sheetName)
    lastColumn = LastHeaderColumn(ws)
    expected = Split(columnsText, ",")

    For Each item In expected
        If HeaderColumn(ws, Trim$(CStr(item))) = 0 Then
            lastColumn = lastColumn + 1
            ws.Cells(1, lastColumn).Value = Trim$(CStr(item))
            modLog.WriteRow "schema.column", True, sheetName & " got the column '" & _
                            Trim$(CStr(item)) & "'"
        End If
    Next item
End Function

' date and start hold text: without the format Excel reads "09:00" as a time and
' "2026-09-25" as a date of the machine's locale (§13).
Private Sub EnsureTextColumn(ByVal sheetName As String, ByVal columnText As String)
    Dim ws As Worksheet
    Dim columnIndex As Long

    If Not modNames.SheetExists(sheetName) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(sheetName)
    columnIndex = HeaderColumn(ws, columnText)
    If columnIndex = 0 Then Exit Sub
    ws.Columns(columnIndex).NumberFormat = "@"
End Sub

' The last column of the header row.
Private Function LastHeaderColumn(ByRef ws As Worksheet) As Long
    LastHeaderColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

' The column that carries a header text, 0 when the header has no such column.
Private Function HeaderColumn(ByRef ws As Worksheet, ByVal headerText As String) As Long
    Dim lastColumn As Long
    Dim index As Long

    lastColumn = LastHeaderColumn(ws)
    For index = 1 To lastColumn
        If StrComp(Trim$(CStr(ws.Cells(1, index).Value)), headerText, vbTextCompare) = 0 Then
            HeaderColumn = index
            Exit Function
        End If
    Next index
End Function
