Option Explicit

' =====================================================================
'  Service_Manager - creates the workbook template (Phase 0)
'
'  The template is the contract of the application: every sheet, every named
'  cell and every cell button is defined here once, and modNames.VerifyContract
'  checks afterwards that the code still agrees with it. The generated file
'  Service_Manager.xlsx is macro free; tools\build.vbs copies it, injects the
'  VBA modules and saves it as a macro workbook.
'
'      cscript //nologo tools\make-template.vbs
'
'  Run it again after a layout change: it overwrites the template and the report
'  build\make-template-report.txt lists every sheet and name it created.
' =====================================================================

Dim gFso, gRoot, gTemplatePath, gReportPath, gOut
Dim gXl, gWb, gReport, gSheetCount
Dim gContract, gSheetNames, gContractPath

Set gFso = CreateObject("Scripting.FileSystemObject")
gRoot = gFso.GetParentFolderName(gFso.GetParentFolderName(WScript.ScriptFullName))
gTemplatePath = gFso.BuildPath(gRoot, "Service_Manager.xlsx")
gReportPath = gFso.BuildPath(gFso.BuildPath(gRoot, "build"), "make-template-report.txt")
' The contract of the workbook goes next to the model under cscript (tools\sim),
' which builds its own workbook from it (see WriteContract below).
gContractPath = gFso.BuildPath(gFso.BuildPath(gFso.BuildPath(gRoot, "tools"), "sim"), "sim-contract.vbs")

If Not gFso.FolderExists(gFso.BuildPath(gRoot, "build")) Then
    gFso.CreateFolder gFso.BuildPath(gRoot, "build")
End If

Set gOut = gFso.CreateTextFile(gReportPath, True, True)
Set gReport = CreateObject("System.Collections.ArrayList")
Set gContract = CreateObject("System.Collections.ArrayList")
Set gSheetNames = CreateObject("System.Collections.ArrayList")

AddLine "Service_Manager - template"
AddLine "generated : " & CStr(Now)
AddLine "file      : " & gTemplatePath

On Error Resume Next
Set gXl = CreateObject("Excel.Application")
If Err.Number <> 0 Then AbortTemplate "cannot start Excel: " & Err.Description
gXl.DisplayAlerts = False
gXl.ScreenUpdating = False
gXl.SheetsInNewWorkbook = 1
Set gWb = gXl.Workbooks.Add
On Error GoTo 0

If gWb Is Nothing Then AbortTemplate "no workbook was created"

' The workbook starts with one sheet: it becomes the Dashboard, so the first tab
' is the start screen and no empty sheet stays behind.
gWb.Worksheets(1).Name = "Dashboard"

BuildDashboard
BuildBooking
BuildClientsAdmin
BuildServicesAdmin
BuildAppointments
BuildReports
BuildConfig
BuildLog
BuildAbout
BuildLists
BuildDataSheets
BuildDropdowns

' A phone is text, not a number: in a General cell 0888123456 becomes 888123456 and
' the leading zero is gone (24.09.2026). The generator sets the format where a phone
' lives: the input cell, and the phone column of the ClientsAdmin grid and of
' db_clients, found by their header - so the layout of the builder does not matter.
Sub FixPhoneCells()
    Dim inputCell

    On Error Resume Next
    Set inputCell = gWb.Names("i_client_phone").RefersToRange
    On Error GoTo 0
    If Not inputCell Is Nothing Then inputCell.NumberFormat = "@"

    SetTextColumn "ClientsAdmin", "Phone"
    SetTextColumn "Appointments", "phone"
    SetTextColumn "db_clients", "phone"
End Sub

Sub SetTextColumn(ByVal sheetName, ByVal headerText)
    Dim ws, headerRow, colIndex

    On Error Resume Next
    Set ws = gWb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    For headerRow = 1 To 20
        For colIndex = 1 To 30
            If StrComp(Trim(CStr(ws.Cells(headerRow, colIndex).Text)), headerText, _
                       vbTextCompare) = 0 Then
                ws.Columns(colIndex).NumberFormat = "@"
                Exit Sub
            End If
        Next
    Next
End Sub

' The workbook opens on the Dashboard; the harness stores the cursor cell.
gWb.Worksheets("Dashboard").Activate
gWb.Worksheets("Dashboard").Range("A10").Select
FixPhoneCells

' The template must stay macro free: xlOpenXMLWorkbook = 51.
If gFso.FileExists(gTemplatePath) Then gFso.DeleteFile gTemplatePath, True
gWb.SaveAs gTemplatePath, 51
If Err.Number <> 0 Then
    AbortTemplate "cannot save " & gTemplatePath & ": " & Err.Description
End If

AddLine ""
AddLine "sheets    : " & CStr(gSheetCount)
AddLine "names     : " & CStr(gContract.Count)
AddLine "contract  : " & gContractPath
AddLine "state     : finished"
gWb.Close False
gXl.Quit
WriteContract
gOut.Write JoinLines()
gOut.Close
WScript.Echo "template written: " & gTemplatePath
WScript.Echo "report: " & gReportPath
WScript.Echo "contract: " & gContractPath
WScript.Quit 0

' ---------------------------------------------------------------------
Function JoinLines()
    Dim index, text

    text = ""
    For index = 0 To gReport.Count - 1
        text = text & gReport.Item(index) & vbCrLf
    Next
    JoinLines = text
End Function

Sub AbortTemplate(ByVal message)
    gOut.WriteLine "FAILED    : " & message
    gOut.Close
    On Error Resume Next
    If Not gWb Is Nothing Then gWb.Close False
    If Not gXl Is Nothing Then gXl.Quit
    WScript.Echo "FAILED: " & message
    WScript.Quit 2
End Sub

Sub AddLine(ByVal text)
    gReport.Add text
End Sub

' ---------------------------------------------------------------------
'  the contract (tools\sim\sim-contract.vbs)
' ---------------------------------------------------------------------

' The workbook contract as VBScript: the model under cscript builds its own
' workbook from it (tools\sim\sim-data.vbs), so a name can never exist in the
' model and not in the template (or the other way round). Generated, never
' edited - rerun tools\make-template.vbs after a layout change.
Sub WriteContract()
    Dim file

    Set file = gFso.CreateTextFile(gContractPath, True, False)
    file.WriteLine "' ---------------------------------------------------------------------------"
    file.WriteLine "'  Service_Manager - the contract of the workbook (GENERATED)"
    file.WriteLine "'"
    file.WriteLine "'  Written by tools\make-template.vbs: the sheets and every workbook scoped"
    file.WriteLine "'  name with the cell it points at. tools\sim (the model under cscript) builds"
    file.WriteLine "'  its workbook from this list, so the model and the template cannot disagree"
    file.WriteLine "'  about a name. Do not edit - rerun the generator after a layout change."
    file.WriteLine "'"
    file.WriteLine "'  generated : " & CStr(Now)
    file.WriteLine "'  names     : " & CStr(gContract.Count)
    file.WriteLine "' ---------------------------------------------------------------------------"
    file.WriteLine "Option Explicit"
    file.WriteLine ""
    file.WriteLine "' The sheets of the workbook, in the order of the tabs."
    WriteStringArray file, "ContractSheets", gSheetNames
    file.WriteLine ""
    file.WriteLine "' Every named cell as ""name|sheet|row|column""."
    WriteStringArray file, "ContractNames", gContract
    file.Close
End Sub

' "Function X() / X = Array( _ / ... )" - wrapped, because a physical line in
' VBScript is limited and a logical one takes at most 24 continuations.
Sub WriteStringArray(ByRef file, ByVal functionName, ByRef items)
    Dim index, step, text, prefix, suffix, first

    first = True
    index = 0
    file.WriteLine "Function " & functionName & "()"
    Do While index < items.Count
        text = ""
        For step = 1 To 6
            If index < items.Count Then
                If Len(text) > 0 Then text = text & ", "
                text = text & """" & items.Item(index) & """"
                index = index + 1
            End If
        Next

        If index >= items.Count Then
            suffix = ")"
        Else
            suffix = ", _"
        End If

        If first Then
            prefix = "    " & functionName & " = Array("
            first = False
        Else
            prefix = "        "
        End If

        file.WriteLine prefix & text & suffix
    Loop

    file.WriteLine "End Function"
End Sub

' Gets or creates a worksheet. The sheets are created in the order the Build*
' subs run, so the tab order is the order of the code below.
Function NewSheet(ByVal sheetName, ByVal gridlines)
    Dim ws

    On Error Resume Next
    Set ws = Nothing
    Set ws = gWb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = gWb.Worksheets.Add(, gWb.Worksheets(gWb.Worksheets.Count))
        ws.Name = sheetName
    End If

    ' DisplayGridlines belongs to the sheet's window, so the sheet has to be the
    ' active one; errors are ignored, they are cosmetic.
    On Error Resume Next
    ws.Activate
    ws.DisplayGridlines = gridlines
    Err.Clear
    On Error GoTo 0

    gSheetCount = gSheetCount + 1
    If Not gSheetNames.Contains(sheetName) Then gSheetNames.Add sheetName
    AddLine "sheet     : " & sheetName
    Set NewSheet = ws
End Function

' One formatted cell.
Sub Put(ByRef ws, ByVal address, ByVal text, ByVal bold, ByVal size, ByVal italic, _
        ByVal fillColor, ByVal fontColor, ByVal numberFormat)
    Dim cell

    Set cell = ws.Range(address)
    cell.NumberFormat = numberFormat
    cell.Value = text
    cell.Font.Bold = bold
    cell.Font.Size = size
    cell.Font.Italic = italic
    If fontColor >= 0 Then cell.Font.Color = fontColor
    If fillColor >= 0 Then cell.Interior.Color = fillColor
End Sub

Sub Title(ByRef ws, ByVal address, ByVal text)
    Put ws, address, text, True, 14, False, -1, -1, "@"
End Sub

Sub Hint(ByRef ws, ByVal address, ByVal text)
    Put ws, address, text, False, 9, True, -1, 8421504, "@"
End Sub

Sub Label(ByRef ws, ByVal address, ByVal text)
    Put ws, address, text, True, 10, False, -1, -1, "@"
End Sub

Sub Section(ByRef ws, ByVal address, ByVal text)
    Put ws, address, text, True, 11, False, 15263769, -1, "@"
End Sub

' An input cell: text format, so a phone number keeps its leading zero.
Sub InputCell(ByRef ws, ByVal address)
    Put ws, address, "", False, 10, False, 15790320, -1, "@"
End Sub

' An output cell (only the code writes into it).
Sub OutputCell(ByRef ws, ByVal address)
    Put ws, address, "", False, 10, False, -1, -1, "@"
End Sub

' Declares a workbook scoped name and returns the range.
Function NameCell(ByRef ws, ByVal nameText, ByVal address)
    Dim target

    Set target = ws.Range(address)
    gWb.Names.Add nameText, "=" & ws.Name & "!" & target.Address(True, True, -4159)
    gContract.Add nameText & "|" & ws.Name & "|" & CStr(target.Row) & "|" & CStr(target.Column)
    AddLine "  name    : " & nameText & " -> " & ws.Name & "!" & target.Address(False, False)
    Set NameCell = target
End Function

' A cell that behaves like a button: a hyperlink that points to its own name.
Sub Button(ByRef ws, ByVal nameText, ByVal address, ByVal text, ByVal tip)
    Dim target

    Set target = NameCell(ws, nameText, address)
    Put ws, address, text, True, 10, False, 15259684, 0, "@"
    target.HorizontalAlignment = -4108
    target.Borders.LineStyle = 1

    On Error Resume Next
    ws.Hyperlinks.Add target, "", nameText, tip, text
    If Err.Number <> 0 Then
        AddLine "  WARNING : hyperlink for " & nameText & " failed: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

    ' Excel paints a new hyperlink blue and underlined; a button is not a link.
    target.Font.Underline = -4142         ' xlUnderlineStyleNone
    target.Font.Color = 0
    target.Font.Bold = True
    target.Interior.Color = 15259684
    target.HorizontalAlignment = -4108    ' xlCenter
End Sub

' A plain (unnamed) row of headers: the texts arrive as one comma separated list,
' starting in column A.
Sub HeaderRow(ByRef ws, ByVal rowIndex, ByVal headersText)
    Dim parts, index, address

    parts = Split(headersText, ",")
    For index = 0 To UBound(parts)
        If Len(Trim(parts(index))) > 0 Then
            address = ws.Cells(rowIndex, index + 1).Address(False, False)
            Put ws, address, Trim(parts(index)), True, 10, False, 13421772, -1, "@"
        End If
    Next
End Sub

' Input cell + name in one step.
Function InName(ByRef ws, ByVal nameText, ByVal address)
    InputCell ws, address
    Set InName = NameCell(ws, nameText, address)
End Function

' An input cell that opens with a value: the Config page of a fresh workbook shows
' the built in settings of §16.1, so it can be applied without typing anything.
Sub InNameValue(ByRef ws, ByVal nameText, ByVal address, ByVal valueText)
    Dim target

    Set target = InName(ws, nameText, address)
    target.Value = valueText
End Sub

' Output cell + name in one step.
Function OutName(ByRef ws, ByVal nameText, ByVal address)
    OutputCell ws, address
    Set OutName = NameCell(ws, nameText, address)
End Function

' A state cell (s_view, s_date, ...). Text format, so a date stays the ISO text
' "2026-09-18" and Excel does not turn it into a locale date.
Function StateName(ByRef ws, ByVal nameText, ByVal address)
    Put ws, address, "", True, 10, False, 15921906, -1, "@"
    Set StateName = NameCell(ws, nameText, address)
End Function

' A state cell that opens with a value. The template of this build carries the
' number of the schema it builds (modSchema.CURRENT_SCHEMA_VERSION; Test_Schema
' compares the two, so they cannot drift apart).
Sub StateValue(ByRef ws, ByVal nameText, ByVal address, ByVal valueText)
    Dim target

    Set target = StateName(ws, nameText, address)
    target.Value = valueText
End Sub

' ---------------------------------------------------------------------
'  the sheets
' ---------------------------------------------------------------------
' Column widths of a sheet, one comma separated list ("12,24,18").
Sub Widths(ByRef ws, ByVal numbersText)
    Dim parts, index

    parts = Split(numbersText, ",")
    For index = 0 To UBound(parts)
        ws.Columns(index + 1).ColumnWidth = CDbl(Trim(parts(index)))
    Next
End Sub

Sub BuildBooking()
    Dim ws

    Set ws = NewSheet("Booking", False)
    Widths ws, "16,14,9,9,9,14"

    Title ws, "A1", "Booking"
    Hint ws, "A2", "Pick a date, a client and a service, then set the start time and the duration."

    Label ws, "A4", "Date"
    InName ws, "i_booking_date", "B4"
    Button ws, "b_booking_date_pick", "C4", "Pick", "Pick a date"
    Label ws, "A5", "Client"
    InName ws, "i_booking_client", "B5"
    ' The typed client search of §9.5 sits next to the client field, on a cell of
    ' its own: C5 already belongs to o_booking_client_id, and a cell another name
    ' writes into cannot carry a caption (it was overwritten by the id, 19.09.2026).
    Button ws, "b_booking_client_pick", "D5", "Search", "Search the clients by any part of the name or phone"
    OutName ws, "o_booking_client_id", "C5"
    Label ws, "A6", "Service"
    InName ws, "i_booking_service", "B6"
    OutName ws, "o_booking_price", "C6"
    OutName ws, "o_booking_service_duration", "D6"

    Label ws, "A7", "Start"
    InName ws, "i_booking_start_h", "B7"
    InName ws, "i_booking_start_m", "C7"
    Button ws, "b_booking_start_pick", "D7", "Pick", "Pick the start time"
    OutName ws, "o_booking_start", "F7"

    Label ws, "A8", "Duration"
    InName ws, "i_booking_dur_h", "B8"
    InName ws, "i_booking_dur_m", "C8"
    Button ws, "b_booking_dur_pick", "D8", "Pick", "Pick the duration"
    OutName ws, "o_booking_dur", "F8"

    Label ws, "A9", "End"
    OutName ws, "o_booking_end", "F9"

    Label ws, "A10", "Mode"
    OutName ws, "o_booking_mode", "B10"
    StateName ws, "s_edit_id", "C10"

    Button ws, "b_booking_save", "A12", "Save", "Save the appointment"
    Button ws, "b_booking_clear", "B12", "Clear", "Clear the form"
    Button ws, "b_booking_cancel", "C12", "Cancel", "Leave the edit mode"
    Button ws, "b_home_booking", "D12", "Dashboard", "Back to the Dashboard"

    Label ws, "A14", "Status:"
    OutName ws, "o_status_booking", "B14"

    ' The dropdowns of the client and the service are added by BuildDropdowns:
    ' they need the names of _Lists, and that sheet is built after this one.
End Sub

' The dropdowns of Booking, in a second pass. A validation whose formula names a
' range that does not exist yet makes Excel refuse the formula with "Unknown
' runtime error" and the cell silently stays a plain cell - the WARNING of the
' first version of this generator (BuildBooking ran before BuildLists, so both
' dropdowns of Booking were missing).
Sub BuildDropdowns()
    AddLine "pass      : dropdowns"
    Dropdown gWb.Worksheets("Booking"), "B5", "list_clients"
    Dropdown gWb.Worksheets("Booking"), "B6", "list_services"
End Sub

' A cell with a dropdown bound to a named range on _Lists. The formula is read
' back: a dropdown that did not arrive stops the generator instead of leaving a
' cell that cannot be filled from the list (§6.5.4).
Sub Dropdown(ByRef ws, ByVal address, ByVal listName)
    Dim target, validation, formulaText

    Set target = Nothing
    Set validation = Nothing
    formulaText = ""

    Set target = ws.Range(address)

    On Error Resume Next
    Set validation = target.Validation
    validation.Delete
    Err.Clear
    validation.Add 3, 1, 1, "=" & listName       ' xlValidateList, xlValidAlertStop
    If Err.Number <> 0 Then
        On Error GoTo 0
        AbortTemplate "the dropdown of " & ws.Name & "!" & address & " failed (" & _
                      listName & "): " & CStr(Err.Number) & " " & Err.Description
    End If
    validation.IgnoreBlank = True
    validation.InCellDropdown = True
    validation.ErrorTitle = "Pick from the list"
    validation.ErrorMessage = "Choose one of the values of the list."
    formulaText = CStr(validation.Formula1)
    If Err.Number <> 0 Then
        On Error GoTo 0
        AbortTemplate "the dropdown of " & ws.Name & "!" & address & _
                      " cannot be read back: " & CStr(Err.Number) & " " & Err.Description
    End If
    On Error GoTo 0

    If StrComp(formulaText, "=" & listName, vbTextCompare) <> 0 Then
        AbortTemplate "the dropdown of " & ws.Name & "!" & address & " points at '" & _
                      formulaText & "' instead of '=" & listName & "'"
    End If

    AddLine "dropdown  : " & ws.Name & "!" & address & " -> " & listName
End Sub

Sub BuildClientsAdmin()
    Dim ws

    Set ws = NewSheet("ClientsAdmin", False)
    Widths ws, "12,26,18,26,30,14"

    Title ws, "A1", "Clients"
    Hint ws, "A2", "Fill the form and press Add client; the list below shows the whole table."
    Section ws, "A4", "New client"

    Label ws, "A5", "Name"
    InName ws, "i_client_name", "B5"
    Label ws, "A6", "Phone"
    InName ws, "i_client_phone", "B6"
    Label ws, "A7", "Email"
    InName ws, "i_client_email", "B7"
    Label ws, "A8", "Comment"
    InName ws, "i_client_comment", "B8"

    Button ws, "b_clients_add", "D5", "Add client", "Store the new client"
    Button ws, "b_clients_clear", "D7", "Clear form", "Empty the form"
    Button ws, "b_home_clients", "D9", "Dashboard", "Back to the Dashboard"

    Label ws, "A10", "Status:"
    OutName ws, "o_status_clients", "B10"

    OutName ws, "o_clients_title", "A12"
    HeaderRow ws, 13, "id,name,phone,email,comment,actions"
    NameCell ws, "grid_clients_first", "A14"
    NameCell ws, "grid_clients_last", "F63"
End Sub

Sub BuildDashboard()
    Dim ws

    Set ws = NewSheet("Dashboard", False)
    Widths ws, "16,12,12,20,13,15,12,12,14"

    Title ws, "A1", "Service Manager"
    Hint ws, "A2", "Click the boxed cells. The header belongs to the app; type only in the cells with a background."

    Button ws, "b_view_daily", "A3", "Daily", "Show one day"
    Button ws, "b_view_weekly", "B3", "Weekly", "Show one week"
    Button ws, "b_view_monthly", "C3", "Monthly", "Show one month"
    Button ws, "b_prev", "E3", "<", "Previous period"
    Button ws, "b_today", "F3", "Today", "Go back to today"
    Button ws, "b_next", "G3", ">", "Next period"

    Label ws, "A4", "View:"
    OutName ws, "o_view", "B4"
    Label ws, "C4", "Period:"
    OutName ws, "o_period", "D4"
    Label ws, "F4", "Appointments:"
    OutName ws, "o_count", "G4"
    Label ws, "H4", "Revenue:"
    OutName ws, "o_revenue", "I4"

    Label ws, "A5", "Free slots:"
    OutName ws, "o_free_slots", "B5"
    Label ws, "C5", "Status:"
    OutName ws, "o_status_dashboard", "D5"

    Button ws, "b_new_appointment", "A6", "New appointment", "Book a new appointment"
    Button ws, "b_go_clients", "B6", "Clients", "Open the clients page"
    Button ws, "b_go_services", "C6", "Services", "Open the services page"
    Button ws, "b_go_appointments", "D6", "Appointments", "Open all appointments"
    Button ws, "b_go_reports", "E6", "Reports", "Open the reports page"
    Button ws, "b_go_config", "F6", "Config", "Open the settings"
    Button ws, "b_go_about", "G6", "About", "Version and diagnostics"

    OutName ws, "cal_title", "A8"
    HeaderRow ws, 9, "Time,End,Client,Service,Duration,Price"
    NameCell ws, "cal_first", "A10"
    ' The body holds 151 rows: the whole working day at the finest resolution
    ' (12 h / 5 min = 144 slots) must fit, because the step of Config is the
    ' resolution of the daily and the weekly grid (§6.9).
    NameCell ws, "cal_last", "I160"

    Label ws, "K3", "state (do not edit)"
    Label ws, "K4", "view"
    StateName ws, "s_view", "L4"
    Label ws, "K5", "date"
    StateName ws, "s_date", "L5"

    Button ws, "b_refresh", "K7", "Refresh", "Run app.refresh"
End Sub

Sub BuildServicesAdmin()
    Dim ws

    Set ws = NewSheet("ServicesAdmin", False)
    Widths ws, "14,28,14,16,14"

    Title ws, "A1", "Services"
    Hint ws, "A2", "The duration is what a booking of this service suggests (minutes)."
    Section ws, "A4", "New service"

    Label ws, "A5", "Service"
    InName ws, "i_service_name", "B5"
    Label ws, "A6", "Price"
    InName ws, "i_service_price", "B6"
    Label ws, "A7", "Duration (min)"
    InName ws, "i_service_duration", "B7"

    Button ws, "b_services_add", "D5", "Add service", "Store the new service"
    Button ws, "b_services_clear", "D7", "Clear form", "Empty the form"
    Button ws, "b_home_services", "D9", "Dashboard", "Back to the Dashboard"

    Label ws, "A9", "Status:"
    OutName ws, "o_status_services", "B9"

    OutName ws, "o_services_title", "A11"
    HeaderRow ws, 12, "id,service,price,duration,actions"
    NameCell ws, "grid_services_first", "A13"
    NameCell ws, "grid_services_last", "E62"
End Sub

Sub BuildAppointments()
    Dim ws

    Set ws = NewSheet("Appointments", False)
    Widths ws, "10,14,10,10,13,24,18,28,24,14,14"

    Title ws, "A1", "Appointments"
    Hint ws, "A2", "All appointments with client contact details. The list refreshes after changes."
    Button ws, "b_appointments_refresh", "A3", "Refresh", "Refresh the appointment list"
    Button ws, "b_home_appointments", "C3", "Dashboard", "Back to the Dashboard"
    Label ws, "A4", "Status:"
    OutName ws, "o_status_appointments", "B4"

    OutName ws, "o_appointments_title", "A6"
    HeaderRow ws, 7, "id,date,start,end,duration,client,phone,email,service,edit,delete"
    NameCell ws, "grid_appointments_first", "A8"
End Sub

Sub BuildReports()
    Dim ws

    Set ws = NewSheet("Reports", False)
    Widths ws, "14,18,18,18,18,14"

    Title ws, "A1", "Reports"
    Hint ws, "A2", "Pick a period and press Show. The report is written below."

    Label ws, "A4", "From"
    InName ws, "i_report_from", "B4"
    Label ws, "A5", "To"
    InName ws, "i_report_to", "B5"

    Button ws, "b_report_render", "D4", "Show", "Render the report"
    Button ws, "b_report_prev", "D5", "<", "Previous period"
    Button ws, "b_report_next", "D6", ">", "Next period"
    Button ws, "b_home_reports", "D7", "Dashboard", "Back to the Dashboard"
    Button ws, "b_report_from_pick", "C4", "Pick", "Pick the start of the period"
    Button ws, "b_report_to_pick", "C5", "Pick", "Pick the end of the period"

    Label ws, "A7", "Status:"
    OutName ws, "o_status_reports", "B7"

    OutName ws, "o_report_title", "A9"
    HeaderRow ws, 10, "group,count,revenue"
    NameCell ws, "grid_reports_first", "A11"
    NameCell ws, "grid_reports_last", "C60"
End Sub

Sub BuildConfig()
    Dim ws

    Set ws = NewSheet("Config", False)
    Widths ws, "24,18,14,14,14,14,18,14"

    Title ws, "A1", "Config"
    Hint ws, "A2", "Settings of the business. Press Apply settings after a change."
    Section ws, "A4", "Working hours and bookings"

    Label ws, "A5", "Day start"
    InNameValue ws, "i_config_day_start", "B5", "09:00"
    Label ws, "A6", "Day end"
    InNameValue ws, "i_config_day_end", "B6", "19:00"
    Label ws, "A7", "Slot step (min)"
    InNameValue ws, "i_config_step", "B7", "5"
    Label ws, "A8", "Default duration (min)"
    InNameValue ws, "i_config_duration", "B8", "10"
    Label ws, "A9", "Currency"
    InNameValue ws, "i_config_currency", "B9", "EUR"
    Label ws, "A10", "Number format"
    InNameValue ws, "i_config_number_format", "B10", "#,##0.00"
    Label ws, "A11", "Non working weekdays (1 = Mon)"
    InNameValue ws, "i_config_nonwork_days", "B11", "6,7"
    Label ws, "A12", "Non working dates"
    InName ws, "i_config_nonwork_dates", "B12"

    Button ws, "b_config_apply", "D5", "Apply settings", "Check and store the settings"
    Button ws, "b_config_defaults", "D7", "Restore defaults", "Back to the built in values"
    Button ws, "b_config_nonwork_pick", "D12", "Pick a day", "Add a day to the non working dates"
    Button ws, "b_config_nonwork_clear", "E12", "Clear dates", "Remove every non working date"
    Button ws, "b_lock_data", "D9", "Lock data sheets", "Lock db_clients, db_services, db_appointments"
    Button ws, "b_unlock_data", "D11", "Unlock data sheets", "Unlock the data sheets for manual work"
    Button ws, "b_home_config", "D13", "Dashboard", "Back to the Dashboard"

    Label ws, "A14", "Status:"
    OutName ws, "o_status_config", "B14"

    Label ws, "G5", "state (do not edit)"
    Label ws, "G6", "schema"
    StateValue ws, "s_schema_version", "H6", "2"
    Label ws, "G7", "period from"
    StateName ws, "s_period_from", "H7"
    Label ws, "G8", "period to"
    StateName ws, "s_period_to", "H8"
End Sub

Sub BuildLog()
    Dim ws

    Set ws = NewSheet("Log", False)
    Widths ws, "22,22,6,70"

    HeaderRow ws, 1, "timestamp,action,ok,message"
    NameCell ws, "log_first", "A2"
End Sub

Sub BuildAbout()
    Dim ws

    Set ws = NewSheet("About", False)
    Widths ws, "20,44,14"

    Title ws, "A1", "About"
    Label ws, "A3", "Version"
    OutName ws, "o_about_version", "B3"
    Label ws, "A4", "File"
    OutName ws, "o_about_file", "B4"
    Label ws, "A5", "Sheets"
    OutName ws, "o_about_sheets", "B5"
    Hint ws, "A7", "Diagnostics: press the button below, or run modApp.AppRouter with ""app.diagnostics""."

    Button ws, "b_home_about", "A9", "Dashboard", "Back to the Dashboard"
End Sub

Sub BuildLists()
    Dim ws

    Set ws = NewSheet("_Lists", False)
    Widths ws, "32,32,12"

    HeaderRow ws, 1, "clients,services,times"
    NameCell ws, "list_clients", "A2:A201"
    NameCell ws, "list_services", "B2:B201"
    NameCell ws, "list_times", "C2:C201"

    ws.Visible = 0            ' xlSheetHidden: the dropdown lists are technical
End Sub

' One data row: the values arrive as one comma separated list (no value may
' contain a comma), empty values are skipped.
Sub DataRow(ByRef ws, ByVal rowIndex, ByVal valuesText)
    Dim parts, index

    parts = Split(valuesText, ",")
    For index = 0 To UBound(parts)
        If Len(Trim(parts(index))) > 0 Then
            ws.Cells(rowIndex, index + 1).Value = Trim(parts(index))
        End If
    Next
End Sub

Sub BuildDataSheets()
    Dim ws

    ' ---- db_clients: id, name, phone, email, comment -------------------
    Set ws = NewSheet("db_clients", True)
    HeaderRow ws, 1, "id,name,phone,email,comment"
    Widths ws, "8,26,18,26,32"
    ws.Columns(2).NumberFormat = "@"
    ws.Columns(3).NumberFormat = "@"
    ws.Columns(4).NumberFormat = "@"
    ws.Columns(5).NumberFormat = "@"
    DataRow ws, 2, "1,Ivan Petrov,0888123456,ivan@example.com,Prefers afternoon"
    DataRow ws, 3, "2,Maria Ivanova,0888987654,maria@example.com,"
    DataRow ws, 4, "3,Elena Georgieva,0899123456,elena@example.com,"

    ' ---- db_services: id, service, price, duration ---------------------
    Set ws = NewSheet("db_services", True)
    HeaderRow ws, 1, "id,service,price,duration"
    Widths ws, "8,28,14,14"
    ws.Columns(2).NumberFormat = "@"
    DataRow ws, 2, "1,Haircut,25,30"
    DataRow ws, 3, "2,Hair coloring,60,90"
    DataRow ws, 4, "3,Manicure,30,45"

    ' ---- db_appointments: id, client_id, service_id, date, start, duration
    Set ws = NewSheet("db_appointments", True)
    HeaderRow ws, 1, "id,client_id,service_id,date,start,duration"
    Widths ws, "8,12,12,14,12,12"
    ws.Columns(4).NumberFormat = "@"
    ws.Columns(5).NumberFormat = "@"
    DataRow ws, 2, "1,1,1,2026-09-21,09:00,30"
    DataRow ws, 3, "2,2,2,2026-09-22,10:30,90"
    DataRow ws, 4, "3,3,3,2026-09-24,14:00,45"

    ' ---- About: fill the three values up front (26.09.2026) ---------------
    ' The app rewrites them whenever the About page is opened, but the sheet should not
    ' look empty until then. The version is read from the sources, so there is still a
    ' single place to bump it (modConfig.APP_VERSION).
    Set ws = gWb.Worksheets("About")
    ws.Range("B3").Value = AppVersionFromSources()
    ws.Range("B4").Value = gWb.Name
    ws.Range("B5").Value = CStr(gWb.Worksheets.Count)
    ws.Range("B3:B5").NumberFormat = "@"
End Sub

' "1.1.0" - the application version, read out of src\app\modConfig.bas so the generator
' never carries a second copy of it.
Function AppVersionFromSources()
    AppVersionFromSources = ""

    Dim fso, scriptDir, configPath, stream, text, position, closing

    Set fso = CreateObject("Scripting.FileSystemObject")
    scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
    configPath = fso.BuildPath(fso.BuildPath(scriptDir, ".."), "src\app\modConfig.bas")
    If Not fso.FileExists(configPath) Then
        WScript.Echo "warning: cannot read the version, " & configPath & " is not there"
        Exit Function
    End If

    Set stream = fso.OpenTextFile(configPath, 1, False, -2)
    text = stream.ReadAll
    stream.Close

    position = InStr(1, text, "APP_VERSION As String = """, vbBinaryCompare)
    If position = 0 Then Exit Function
    position = position + Len("APP_VERSION As String = """)
    closing = InStr(position, text, """")
    If closing <= position Then Exit Function
    AppVersionFromSources = Mid(text, position, closing - position)
End Function
