Attribute VB_Name = "modReports"
Option Explicit

' =====================================================================
'  Service_Manager - the reports (Phase 5)
'
'  A report is a period (from..to, both ends included) and three groupings of the
'  appointments inside it: by day, by service and by client. The period lives in
'  the two input cells of the sheet (i_report_from / i_report_to), so the file
'  keeps it (§13); empty cells mean the current month (§16.1) and this module
'  writes those defaults into the cells, so the user always sees what the numbers
'  cover.
'
'  The model in tools\sim carries the same functions under the same names
'  (PeriodRuleCode, ShiftPeriodIso, ReportRows ...): a rule is designed and tested
'  there first, then it is ported here, and the cross check of tools\build.vbs
'  compares the two row by row (tools\sim\README.md).
' =====================================================================

' ---- the state of one report -----------------------------------------
' The rows the last ReportRows built (the writer underneath needs them) and whether
' the grid was too small for them. Module level declarations come before the first
' procedure (a rule of the VBE, §13).
Private mReportLines As String
Private mReportLineCount As Long
Private mReportCapped As Boolean

' =====================================================================
'  the route: reports.render (Show, < and >)
' =====================================================================

' The payload can carry from / to (the cells are the state, so the payload wins
' for this call) and a step of whole months for the two arrow buttons. The report
' is written into the grid below the title; the status line says how many rows it
' has.
Public Sub Render(ByRef data As Object, ByRef result As Object)
    Dim fromText As String
    Dim toText As String
    Dim fromIso As String
    Dim toIso As String
    Dim stepValue As Long
    Dim problems As String
    Dim rows As Long
    Dim info As Object

    fromText = modApp.PayloadText(data, modRoutes.KEY_FROM)
    toText = modApp.PayloadText(data, modRoutes.KEY_TO)
    If Len(fromText) > 0 Then modNames.WriteName NAME_REPORT_FROM, fromText
    If Len(toText) > 0 Then modNames.WriteName NAME_REPORT_TO, toText
    stepValue = modApp.PayloadNumber(data, modRoutes.KEY_STEP)

    fromIso = ReportFromIso()
    toIso = ReportToIso()

    If stepValue <> 0 Then
        fromIso = ShiftPeriodIso(fromIso, stepValue)
        toIso = ShiftPeriodIso(toIso, stepValue)
    End If

    ' What the user will see in the cells: the period as a European date (the
    ' default of the current month lands here as well). An unreadable cell stays
    ' empty and the rule below names it.
    modNames.WriteName NAME_REPORT_FROM, modPickers.DateFormatText(fromIso)
    modNames.WriteName NAME_REPORT_TO, modPickers.DateFormatText(toIso)

    problems = PeriodProblems(fromIso, toIso)
    If Len(problems) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = problems
        Exit Sub
    End If

    rows = WriteReport(fromIso, toIso)

    Set info = CreateObject("Scripting.Dictionary")
    info("rows") = CStr(rows)
    info(modRoutes.KEY_FROM) = fromIso
    info(modRoutes.KEY_TO) = toIso
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_REPORT_RENDERED, _
        CStr(rows), modPickers.DateFormatText(fromIso), modPickers.DateFormatText(toIso))
    result(modRoutes.RESULT_OK) = True
End Sub

' =====================================================================
'  the date pickers of the period (reports.pick_date)
' =====================================================================

' The two Pick buttons carry their cell in the payload (i_report_from /
' i_report_to), the calendar opens on the day the cell holds - or on the first or
' the last day of the current month, exactly the way the period defaults work - and
' the picked day lands in the cell as a European date. The report itself is written
' when the user presses Show, just as it is after typing a date.
Public Sub OnPickDate(ByRef data As Object, ByRef result As Object)
    Dim cellName As String
    Dim initialIso As String
    Dim picked As String
    Dim isFrom As Boolean
    Dim info As Object

    cellName = modApp.PayloadText(data, modRoutes.KEY_CELL)
    isFrom = (StrComp(cellName, NAME_REPORT_FROM, vbTextCompare) = 0)
    If Not isFrom Then
        If StrComp(cellName, NAME_REPORT_TO, vbTextCompare) <> 0 Then
            result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
                modText.KEY_ERROR_REQUIRED, modText.Text(modText.KEY_LABEL_PERIOD_START))
            Exit Sub
        End If
    End If

    initialIso = modPickers.DateTextToIso(modNames.ReadName(cellName))
    If Len(initialIso) = 0 Then
        If isFrom Then
            initialIso = MonthFirstIso()
        Else
            initialIso = MonthLastIso()
        End If
    End If

    picked = modApp.PayloadText(data, modRoutes.KEY_VALUE)
    If Len(picked) = 0 Then
        picked = modPickers.PickDate(initialIso)
    Else
        picked = modPickers.NormaliseDateText(picked)
    End If

    If Len(picked) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CANCELLED)
        Exit Sub
    End If

    modNames.WriteName cellName, modPickers.DateFormatText(picked)

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_DATE) = picked
    Set result(modRoutes.RESULT_DATA) = info

    If isFrom Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_OK_PICKED, _
            modText.Text(modText.KEY_LABEL_PERIOD_START), modPickers.DateFormatText(picked))
    Else
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_OK_PICKED, _
            modText.Text(modText.KEY_LABEL_PERIOD_END), modPickers.DateFormatText(picked))
    End If
    result(modRoutes.RESULT_OK) = True
End Sub

' =====================================================================
'  the rules of the period (the same shape as a booking, §8.2.7)
' =====================================================================

' The rule that refuses the period, as a code with its parameters
' ("error.period|01.10.2026|30.09.2026"), "" when it is fine. The twin of the
' model's PeriodRuleCode, so the two lanes compare codes and the wording stays in
' modText.
Public Function PeriodRuleCode(ByVal fromIso As String, ByVal toIso As String) As String
    If Not modPickers.IsIsoDate(fromIso) Then
        PeriodRuleCode = RULE_REQUIRED & "|period start"
        Exit Function
    End If
    If Not modPickers.IsIsoDate(toIso) Then
        PeriodRuleCode = RULE_REQUIRED & "|period end"
        Exit Function
    End If
    If fromIso > toIso Then
        PeriodRuleCode = RULE_PERIOD & "|" & modPickers.DateFormatText(fromIso) & "|" & _
                         modPickers.DateFormatText(toIso)
    End If
End Function

' The sentence the user reads for a period that is refused ("" = fine): the code
' becomes the text of modText, so the wording lives in one place.
Public Function PeriodProblems(ByVal fromIso As String, ByVal toIso As String) As String
    Dim codeText As String
    Dim parts As Variant

    codeText = PeriodRuleCode(fromIso, toIso)
    If Len(codeText) = 0 Then Exit Function

    parts = Split(codeText, "|")
    Select Case parts(0)
        Case RULE_REQUIRED
            PeriodProblems = modText.TextFormat(modText.KEY_ERROR_REQUIRED, parts(1))
        Case RULE_PERIOD
            PeriodProblems = modText.TextFormat(modText.KEY_ERROR_PERIOD, parts(1), parts(2))
        Case Else
            PeriodProblems = codeText
    End Select
End Function

' =====================================================================
'  the numbers of a period
' =====================================================================

' How many appointments the period holds: the counter the dashboard uses as well.
Public Function BookingsInPeriod(ByVal fromIso As String, ByVal toIso As String) As Long
    If Not modData.IsModelLoaded() Then modData.LoadData
    BookingsInPeriod = modData.AppointmentsInPeriod(fromIso, toIso)
End Function

' The rows the report grid of the template holds (50 in the shipped template).
Public Function ReportGridCapacity() As Long
    ReportGridCapacity = modNames.RangeOfName(NAME_GRID_REPORTS_LAST).Row - _
                         modNames.RangeOfName(NAME_GRID_REPORTS_FIRST).Row + 1
End Function

' Money as cents: a Double written by CStr() follows the regional settings, and the
' two lanes of the cross check would differ on the decimal separator alone (§13).
Private Function MoneyCents(ByVal value As Double) As Long
    MoneyCents = CLng(value * 100)
End Function

' =====================================================================
'  the report
' =====================================================================

' One row of the report: the cap of §6.9 applies here as well - what does not fit
' is dropped and the title says so (ReportIsCapped).
Private Sub ReportLine(ByVal lineText As String)
    If mReportLineCount >= ReportGridCapacity() Then
        mReportCapped = True
        Exit Sub
    End If

    If Len(mReportLines) > 0 Then mReportLines = mReportLines & vbCrLf
    mReportLines = mReportLines & lineText
    mReportLineCount = mReportLineCount + 1
End Sub

Public Function ReportIsCapped() As Boolean
    ReportIsCapped = mReportCapped
End Function

Public Function ReportLineCount() As Long
    ReportLineCount = mReportLineCount
End Function

' The report itself, one line per row of the grid ("group|count|cents"): the days
' of the period, then the services, then the clients, then the total. Days go up,
' services and clients go down by revenue and by name when two of them are worth
' the same - a fixed order, because the cross check compares row by row. The rows
' stay in the module (mReportLines), because the writer underneath needs them.
Public Function ReportRows(ByVal fromIso As String, ByVal toIso As String) As String
    If Not modData.IsModelLoaded() Then modData.LoadData

    mReportLines = vbNullString
    mReportLineCount = 0
    mReportCapped = False

    ReportLine modText.Text(modText.KEY_REPORT_DAYS)
    ReportGroupRows COL_DATE, fromIso, toIso
    ReportLine vbNullString
    ReportLine modText.Text(modText.KEY_REPORT_SERVICES)
    ReportGroupRows COL_SERVICE_ID, fromIso, toIso
    ReportLine vbNullString
    ReportLine modText.Text(modText.KEY_REPORT_CLIENTS)
    ReportGroupRows COL_CLIENT_ID, fromIso, toIso
    ReportLine modText.Text(modText.KEY_REPORT_TOTAL) & "|" & _
               CStr(BookingsInPeriod(fromIso, toIso)) & "|" & _
               CStr(MoneyCents(modData.RevenueInPeriod(fromIso, toIso)))

    ReportRows = mReportLines
End Function

' One section of the report: the appointments of the period, grouped by the column
' asked for, as "group|count|cents" lines (the twin of the model's ReportGroupRows,
' so the two lanes group and sort the same way).
Private Sub ReportGroupRows(ByVal groupColumn As String, ByVal fromIso As String, _
                            ByVal toIso As String)
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim keys As Object
    Dim counters As Object
    Dim money As Object
    Dim isDateMode As Boolean
    Dim isoText As String
    Dim keyText As String
    Dim sortList() As String
    Dim numberList() As Long
    Dim index As Long
    Dim scan As Long
    Dim swapped As Boolean
    Dim swapText As String
    Dim swapSort As String
    Dim swapNumber As Long

    Set table = modData.AppointmentsTable()
    Set keys = CreateObject("Scripting.Dictionary")
    Set counters = CreateObject("Scripting.Dictionary")
    Set money = CreateObject("Scripting.Dictionary")
    isDateMode = (groupColumn = COL_DATE)

    For rowIndex = 1 To table.RowCount
        isoText = modData.IsoDateText(modData.DateTextOfRow(table, rowIndex))
        If StrComp(isoText, fromIso, vbBinaryCompare) >= 0 Then
            If StrComp(isoText, toIso, vbBinaryCompare) <= 0 Then
                If isDateMode Then
                    keyText = modPickers.DateFormatText(isoText)
                ElseIf groupColumn = COL_SERVICE_ID Then
                    keyText = modData.ServiceName(modData.ServiceIdOfRow(table, rowIndex))
                Else
                    keyText = modData.ClientName(modData.ClientIdOfRow(table, rowIndex))
                End If

                If Len(keyText) > 0 Then
                    If counters.Exists(keyText) Then
                        counters(keyText) = counters(keyText) + 1
                        money(keyText) = money(keyText) + MoneyCents( _
                            modData.ServicePrice(modData.ServiceIdOfRow(table, rowIndex)))
                    Else
                        counters(keyText) = 1
                        money(keyText) = MoneyCents( _
                            modData.ServicePrice(modData.ServiceIdOfRow(table, rowIndex)))
                        keys.Add keys.Count, keyText
                    End If
                End If
            End If
        End If
    Next rowIndex

    If keys.Count = 0 Then Exit Sub

    ' The sort keys: the ISO date of a day (so the order is chronological), the name
    ' otherwise. A plain bubble sort - a report holds a few dozen groups.
    ReDim sortList(keys.Count - 1)
    ReDim numberList(keys.Count - 1)
    For index = 0 To keys.Count - 1
        keyText = CStr(keys(index))
        If isDateMode Then
            sortList(index) = modPickers.DateTextToIso(keyText)
        Else
            sortList(index) = keyText
        End If
        numberList(index) = CLng(money(keyText))
    Next index

    For index = 0 To keys.Count - 2
        For scan = 0 To keys.Count - 2 - index
            swapped = False
            If isDateMode Then
                If StrComp(sortList(scan), sortList(scan + 1), vbBinaryCompare) > 0 Then swapped = True
            ElseIf numberList(scan) < numberList(scan + 1) Then
                swapped = True
            ElseIf numberList(scan) = numberList(scan + 1) Then
                If StrComp(sortList(scan), sortList(scan + 1), vbBinaryCompare) > 0 Then swapped = True
            End If

            If swapped Then
                swapText = CStr(keys(scan)): keys(scan) = keys(scan + 1): keys(scan + 1) = swapText
                swapSort = sortList(scan): sortList(scan) = sortList(scan + 1): sortList(scan + 1) = swapSort
                swapNumber = numberList(scan): numberList(scan) = numberList(scan + 1): numberList(scan + 1) = swapNumber
            End If
        Next scan
    Next index

    For index = 0 To keys.Count - 1
        keyText = CStr(keys(index))
        ReportLine keyText & "|" & CStr(counters(keyText)) & "|" & CStr(numberList(index))
    Next index
End Sub

' =====================================================================
'  the sheet
' =====================================================================

' Writes the report into the grid and the title above it, and answers how many rows
' it wrote. The group stays text, the count and the revenue stay numbers: the format
' of Config decides what the user reads ("25.00") and Excel can still sum them.
Public Function WriteReport(ByVal fromIso As String, ByVal toIso As String) As Long
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lines As Variant
    Dim fields As Variant
    Dim index As Long
    Dim countText As String
    Dim centsText As String
    Dim titleText As String

    If Not modNames.SheetExists(SHEET_REPORTS) Then Exit Function

    modUi.ClearBlock NAME_GRID_REPORTS_FIRST, NAME_GRID_REPORTS_LAST
    Set ws = ThisWorkbook.Worksheets(SHEET_REPORTS)
    firstRow = modNames.RangeOfName(NAME_GRID_REPORTS_FIRST).Row

    ReportRows fromIso, toIso
    lines = Split(mReportLines, vbCrLf)

    For index = 0 To UBound(lines)
        If Len(CStr(lines(index))) > 0 Then
            fields = Split(CStr(lines(index)), "|")
            modUi.WriteCell SHEET_REPORTS, firstRow + index, REPORT_COL_GROUP, CStr(fields(0))

            countText = vbNullString
            centsText = vbNullString
            If UBound(fields) >= 1 Then countText = CStr(fields(1))
            If UBound(fields) >= 2 Then centsText = CStr(fields(2))

            If Len(countText) > 0 Then
                modUi.WriteNumberCell ws.Cells(firstRow + index, REPORT_COL_COUNT), _
                                      CLng(countText), REPORT_COUNT_FORMAT
            End If
            If Len(centsText) > 0 Then
                modUi.WriteNumberCell ws.Cells(firstRow + index, REPORT_COL_REVENUE), _
                                      CDbl(centsText) / 100, _
                                      modDashboard.ConfigText(NAME_CONFIG_NUMBER_FORMAT, _
                                                              DEFAULT_NUMBER_FORMAT)
            End If
        End If
    Next index

    titleText = modText.TextFormat(modText.KEY_REPORT_TITLE, _
                                   modPickers.DateFormatText(fromIso), _
                                   modPickers.DateFormatText(toIso))
    If ReportIsCapped() Then
        titleText = modText.TextFormat(modText.KEY_LIST_CAPPED, titleText, _
                                       CStr(ReportGridCapacity()))
    End If
    modNames.WriteName NAME_REPORT_TITLE, titleText

    WriteReport = ReportLineCount()
End Function


' =====================================================================
'  the period
' =====================================================================

' The period of the report: the cells decide, an empty pair means the current month
' (§16.1). The twin of the model's ReportFromIso / ReportToIso.
Public Function ReportFromIso() As String
    ReportFromIso = ReportPeriodIso(NAME_REPORT_FROM, MonthFirstIso())
End Function

Public Function ReportToIso() As String
    ReportToIso = ReportPeriodIso(NAME_REPORT_TO, MonthLastIso())
End Function

Private Function ReportPeriodIso(ByVal nameText As String, ByVal fallbackIso As String) As String
    Dim text As String

    ' An empty cell means the current month (§16.1); a cell that holds something
    ' else goes through as it is ("" when it is not a date), so the rule refuses it
    ' and the message names the end of the period.
    text = Trim$(modNames.ReadName(nameText))
    If Len(text) = 0 Then
        ReportPeriodIso = fallbackIso
    Else
        ReportPeriodIso = modPickers.DateTextToIso(text)
    End If
End Function

' The machine's today, and the first and the last day of its month: DateSerial with
' day 0 of the next month is the last day of this one.
Public Function MonthFirstIso() As String
    MonthFirstIso = modData.IsoDateText(DateSerial(Year(Date), Month(Date), 1))
End Function

Public Function MonthLastIso() As String
    MonthLastIso = modData.IsoDateText(DateSerial(Year(Date), Month(Date) + 1, 0))
End Function

' The same window, moved by whole months (the < and > buttons). DateAdd clamps a
' day the target month does not have (31.03 -> 28.02) - the model's twin does the
' same.
Public Function ShiftPeriodIso(ByVal isoText As String, ByVal months As Long) As String
    Dim moved As Date

    If Not modPickers.IsIsoDate(isoText) Then Exit Function
    If months = 0 Then
        ShiftPeriodIso = isoText
        Exit Function
    End If

    moved = DateAdd("m", months, DateSerial(CLng(Left$(isoText, 4)), _
                                            CLng(Mid$(isoText, 6, 2)), _
                                            CLng(Mid$(isoText, 9, 2))))
    ShiftPeriodIso = modData.IsoDateText(moved)
End Function
