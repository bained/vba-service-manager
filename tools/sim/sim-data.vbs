Option Explicit

' =====================================================================
'  Service_Manager - the fixtures of the model (data, not contract)
'
'  The contract - which sheets exist and which names point where - is NOT
'  written here: tools\make-template.vbs generates it into sim-contract.vbs,
'  which run.wsf includes next to this file. So the model cannot know a name
'  that the template does not have (and the other way round).
'
'  The fixtures are not written here either. They live in ONE file,
'  src\tests\crosscheck.txt - the Config values, the three db_ tables and the
'  cases of the cross check. The workbook reads the very same file (modTests),
'  so a difference between the two lanes is a difference of the code and never
'  of the data (see tools\sim\README.md).
' =====================================================================

Const SHEET_CONFIG = "Config"
Const SHEET_BOOKING = "Booking"
Const SHEET_DASHBOARD = "Dashboard"
Const SHEET_DB_APPOINTMENTS = "db_appointments"
Const SHEET_DB_CLIENTS = "db_clients"
Const SHEET_DB_SERVICES = "db_services"

Const FIXTURE_FILE = "src\tests\crosscheck.txt"

Dim gFixtures

' The root of the project: this script sits in <root>\tools\sim\.
Function ProjectRoot()
    Dim fso

    Set fso = CreateObject("Scripting.FileSystemObject")
    ProjectRoot = fso.GetParentFolderName(fso.GetParentFolderName( _
                  fso.GetParentFolderName(WScript.ScriptFullName)))
End Function

Function FixturePath()
    Dim fso

    Set fso = CreateObject("Scripting.FileSystemObject")
    FixturePath = fso.BuildPath(ProjectRoot(), FIXTURE_FILE)
End Function

' ---- the fixture file -------------------------------------------------
' Read once: config("name"), the header of a table ("header:db_clients"), the
' rows of a table ("rows:db_clients", CrLf joined) and the case lines ("cases",
' CrLf joined). A line that starts with # is a comment, empty lines are skipped
' and [sections] open the part they belong to.
Function Fixtures()
    Dim fso, file, line, section, config, tables, casesText, parts

    ' gFixtures is Empty until the first read (IsObject, because "Empty Is
    ' Nothing" is an error in VBScript, not a False).
    If IsObject(gFixtures) Then
        Set Fixtures = gFixtures
        Exit Function
    End If

    Set gFixtures = CreateObject("Scripting.Dictionary")
    Set config = CreateObject("Scripting.Dictionary")
    Set tables = CreateObject("Scripting.Dictionary")
    casesText = ""

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(FixturePath()) Then
        WScript.Echo "the fixture file is missing: " & FixturePath()
        WScript.Quit 2
    End If

    Set file = fso.OpenTextFile(FixturePath(), 1, False)
    section = ""

    While Not file.AtEndOfStream
        line = Trim(file.ReadLine)
        If Len(line) > 0 Then
            If Left(line, 1) = "[" Then
                section = Mid(line, 2, Len(line) - 2)
            ElseIf Left(line, 1) <> "#" Then
                If section = "config" Then
                    parts = Split(line, "=")
                    If UBound(parts) >= 1 Then
                        config(Trim(parts(0))) = Trim(Mid(line, Len(parts(0)) + 2))
                    End If
                ElseIf section = "cases" Then
                    If Len(casesText) > 0 Then casesText = casesText & vbCrLf
                    casesText = casesText & line
                ElseIf Left(section, 3) = "db_" Then
                    If Not tables.Exists("header:" & section) Then
                        tables("header:" & section) = line
                    Else
                        If Len(tables("rows:" & section)) > 0 Then
                            tables("rows:" & section) = tables("rows:" & section) & vbCrLf
                        End If
                        tables("rows:" & section) = tables("rows:" & section) & line
                    End If
                End If
            End If
        End If
    Wend
    file.Close

    Set gFixtures("config") = config
    Set gFixtures("tables") = tables
    gFixtures("cases") = casesText

    Set Fixtures = gFixtures
End Function

' The value of one Config cell of the fixture file ("" when the file is silent,
' exactly like an empty cell in Excel).
Function FixtureConfig(nameText)
    Dim config, data

    Set data = Fixtures()
    Set config = data("config")
    If config.Exists(nameText) Then FixtureConfig = config(nameText)
End Function

' The header and the data rows of one db_ table, as CSV text.
Function FixtureHeader(sheetName)
    Dim tables, data

    Set data = Fixtures()
    Set tables = data("tables")
    If tables.Exists("header:" & sheetName) Then FixtureHeader = tables("header:" & sheetName)
End Function

Function FixtureRows(sheetName)
    Dim tables, data

    Set data = Fixtures()
    Set tables = data("tables")
    If tables.Exists("rows:" & sheetName) Then FixtureRows = tables("rows:" & sheetName)
End Function

Function FixtureCases()
    Dim data, casesText

    Set data = Fixtures()
    casesText = data("cases")
    FixtureCases = casesText
End Function

Function FixtureRowArray(sheetName)
    Dim rowsText

    rowsText = FixtureRows(sheetName)
    If Len(rowsText) = 0 Then
        FixtureRowArray = Array()
    Else
        FixtureRowArray = Split(rowsText, vbCrLf)
    End If
End Function

' ---- the workbook of the model ---------------------------------------
' BuildSimBook() builds the workbook from the generated contract: every sheet of
' the template and every named cell with its real cell (row and column), so a
' test that writes "i_booking_date" writes exactly the cell Excel would write.
Function BuildSimBook()

    Dim book, sheets, names, i, parts, config, keyText, data

    Set book = New SimBook

    sheets = ContractSheets()
    For i = 0 To UBound(sheets)
        book.AddSheet sheets(i)
    Next

    names = ContractNames()
    For i = 0 To UBound(names)
        parts = Split(names(i), "|")
        book.DefineName parts(0), parts(1), CLng(parts(2)), CLng(parts(3))
    Next

    ' The Config cells the fixture file fills. The others stay empty, exactly like
    ' the cells of the template - the code then falls back to the DEFAULT_*
    ' constants (and the cross check compares the two paths on the same values).
    Set data = Fixtures()
    Set config = data("config")
    For Each keyText In config.Keys
        book.WriteName keyText, config(keyText)
    Next

    Set BuildSimBook = book
End Function

' The three tables of the model: the header names and the CSV rows of the fixture
' file, but as memory (SqlTable.OpenArray does the same trick in Excel when a
' caller already has an array instead of a range).
Function LoadSimTables()
    Dim tables, table

    Set tables = CreateObject("Scripting.Dictionary")

    Set table = New SimTable
    table.Load SHEET_DB_APPOINTMENTS, FixtureHeader(SHEET_DB_APPOINTMENTS), _
               FixtureRowArray(SHEET_DB_APPOINTMENTS)
    Set tables(SHEET_DB_APPOINTMENTS) = table

    Set table = New SimTable
    table.Load SHEET_DB_CLIENTS, FixtureHeader(SHEET_DB_CLIENTS), _
               FixtureRowArray(SHEET_DB_CLIENTS)
    Set tables(SHEET_DB_CLIENTS) = table

    Set table = New SimTable
    table.Load SHEET_DB_SERVICES, FixtureHeader(SHEET_DB_SERVICES), _
               FixtureRowArray(SHEET_DB_SERVICES)
    Set tables(SHEET_DB_SERVICES) = table

    Set LoadSimTables = tables
End Function
