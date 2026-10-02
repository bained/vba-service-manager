Attribute VB_Name = "modNames"
Option Explicit

' =====================================================================
'  Service_Manager - the workbook contract (Phase 1)
'
'  Every named cell of the template is declared here as a constant and reached
'  through ReadName / WriteName, the only procedures that read or write a UI
'  cell. VerifyContract walks DeclaredNames() and reports what the template and
'  the code disagree on, so a renamed cell is a clear message instead of a run
'  time surprise.
'
'  A hyperlink cell button is identified by its own name (Target.SubAddress);
'  the dynamic cells of the dashboard body are identified by their position,
'  they have no name (see modDashboard).
' =====================================================================

' ---- dashboard --------------------------------------------------------
Public Const NAME_VIEW As String = "o_view"
Public Const NAME_PERIOD As String = "o_period"
Public Const NAME_COUNT As String = "o_count"
Public Const NAME_REVENUE As String = "o_revenue"
Public Const NAME_FREE_SLOTS As String = "o_free_slots"
Public Const NAME_STATUS_DASHBOARD As String = "o_status_dashboard"
Public Const NAME_CAL_TITLE As String = "cal_title"
Public Const NAME_CAL_FIRST As String = "cal_first"
Public Const NAME_CAL_LAST As String = "cal_last"
Public Const NAME_STATE_VIEW As String = "s_view"
Public Const NAME_STATE_DATE As String = "s_date"
Public Const NAME_BTN_VIEW_DAILY As String = "b_view_daily"
Public Const NAME_BTN_VIEW_WEEKLY As String = "b_view_weekly"
Public Const NAME_BTN_VIEW_MONTHLY As String = "b_view_monthly"
Public Const NAME_BTN_PREV As String = "b_prev"
Public Const NAME_BTN_TODAY As String = "b_today"
Public Const NAME_BTN_NEXT As String = "b_next"
Public Const NAME_BTN_REFRESH As String = "b_refresh"
Public Const NAME_BTN_NEW_APPOINTMENT As String = "b_new_appointment"

' ---- navigation -------------------------------------------------------
Public Const NAME_BTN_GO_CLIENTS As String = "b_go_clients"
Public Const NAME_BTN_GO_SERVICES As String = "b_go_services"
Public Const NAME_BTN_GO_APPOINTMENTS As String = "b_go_appointments"
Public Const NAME_BTN_GO_REPORTS As String = "b_go_reports"
Public Const NAME_BTN_GO_CONFIG As String = "b_go_config"
Public Const NAME_BTN_GO_ABOUT As String = "b_go_about"
Public Const NAME_BTN_HOME_BOOKING As String = "b_home_booking"
Public Const NAME_BTN_HOME_CLIENTS As String = "b_home_clients"
Public Const NAME_BTN_HOME_SERVICES As String = "b_home_services"
Public Const NAME_BTN_HOME_APPOINTMENTS As String = "b_home_appointments"
Public Const NAME_BTN_HOME_REPORTS As String = "b_home_reports"
Public Const NAME_BTN_HOME_CONFIG As String = "b_home_config"
Public Const NAME_BTN_HOME_ABOUT As String = "b_home_about"

' ---- booking ----------------------------------------------------------
Public Const NAME_BOOKING_DATE As String = "i_booking_date"
Public Const NAME_BOOKING_CLIENT As String = "i_booking_client"
Public Const NAME_BOOKING_SERVICE As String = "i_booking_service"
Public Const NAME_BOOKING_START_H As String = "i_booking_start_h"
Public Const NAME_BOOKING_START_M As String = "i_booking_start_m"
Public Const NAME_BOOKING_DUR_H As String = "i_booking_dur_h"
Public Const NAME_BOOKING_DUR_M As String = "i_booking_dur_m"
Public Const NAME_BOOKING_CLIENT_ID As String = "o_booking_client_id"
Public Const NAME_BOOKING_PRICE As String = "o_booking_price"
Public Const NAME_BOOKING_SERVICE_DURATION As String = "o_booking_service_duration"
Public Const NAME_BOOKING_START As String = "o_booking_start"
Public Const NAME_BOOKING_DURATION As String = "o_booking_dur"
Public Const NAME_BOOKING_END As String = "o_booking_end"
Public Const NAME_BOOKING_MODE As String = "o_booking_mode"
Public Const NAME_STATE_EDIT_ID As String = "s_edit_id"
Public Const NAME_STATUS_BOOKING As String = "o_status_booking"
Public Const NAME_BTN_BOOKING_DATE_PICK As String = "b_booking_date_pick"
Public Const NAME_BTN_BOOKING_START_PICK As String = "b_booking_start_pick"
Public Const NAME_BTN_BOOKING_DUR_PICK As String = "b_booking_dur_pick"
' The typed client search next to the client field of Booking (§9.5).
Public Const NAME_BTN_BOOKING_CLIENT_PICK As String = "b_booking_client_pick"
Public Const NAME_BTN_BOOKING_SAVE As String = "b_booking_save"
Public Const NAME_BTN_BOOKING_CLEAR As String = "b_booking_clear"
Public Const NAME_BTN_BOOKING_CANCEL As String = "b_booking_cancel"

' ---- clients ----------------------------------------------------------
Public Const NAME_CLIENT_NAME As String = "i_client_name"
Public Const NAME_CLIENT_PHONE As String = "i_client_phone"
Public Const NAME_CLIENT_EMAIL As String = "i_client_email"
Public Const NAME_CLIENT_COMMENT As String = "i_client_comment"
Public Const NAME_STATUS_CLIENTS As String = "o_status_clients"
Public Const NAME_CLIENTS_TITLE As String = "o_clients_title"
Public Const NAME_BTN_CLIENTS_ADD As String = "b_clients_add"
Public Const NAME_BTN_CLIENTS_CLEAR As String = "b_clients_clear"
' The row the form is editing (empty = a new one, §6.10): the [Edit] link of the
' list writes the id here and the add button turns into an update button. It lives
' on the same sheet as the form, the way s_edit_id does on Booking.
Public Const NAME_STATE_EDIT_CLIENT As String = "s_edit_client_id"
Public Const NAME_GRID_CLIENTS_FIRST As String = "grid_clients_first"
Public Const NAME_GRID_CLIENTS_LAST As String = "grid_clients_last"

' ---- services ---------------------------------------------------------
Public Const NAME_SERVICE_NAME As String = "i_service_name"
Public Const NAME_SERVICE_PRICE As String = "i_service_price"
Public Const NAME_SERVICE_DURATION As String = "i_service_duration"
Public Const NAME_STATUS_SERVICES As String = "o_status_services"
Public Const NAME_SERVICES_TITLE As String = "o_services_title"
Public Const NAME_BTN_SERVICES_ADD As String = "b_services_add"
Public Const NAME_BTN_SERVICES_CLEAR As String = "b_services_clear"
Public Const NAME_STATE_EDIT_SERVICE As String = "s_edit_service_id"
Public Const NAME_GRID_SERVICES_FIRST As String = "grid_services_first"
Public Const NAME_GRID_SERVICES_LAST As String = "grid_services_last"

' ---- all appointments -------------------------------------------------
Public Const NAME_STATUS_APPOINTMENTS As String = "o_status_appointments"
Public Const NAME_APPOINTMENTS_TITLE As String = "o_appointments_title"
Public Const NAME_BTN_APPOINTMENTS_REFRESH As String = "b_appointments_refresh"
Public Const NAME_GRID_APPOINTMENTS_FIRST As String = "grid_appointments_first"

' ---- reports ----------------------------------------------------------
Public Const NAME_REPORT_FROM As String = "i_report_from"
Public Const NAME_REPORT_TO As String = "i_report_to"
Public Const NAME_STATUS_REPORTS As String = "o_status_reports"
Public Const NAME_REPORT_TITLE As String = "o_report_title"
Public Const NAME_BTN_REPORT_RENDER As String = "b_report_render"
Public Const NAME_BTN_REPORT_PREV As String = "b_report_prev"
Public Const NAME_BTN_REPORT_NEXT As String = "b_report_next"
' The two date pickers of the report period (From and To), §9.2.
Public Const NAME_BTN_REPORT_FROM_PICK As String = "b_report_from_pick"
Public Const NAME_BTN_REPORT_TO_PICK As String = "b_report_to_pick"
Public Const NAME_GRID_REPORTS_FIRST As String = "grid_reports_first"
Public Const NAME_GRID_REPORTS_LAST As String = "grid_reports_last"

' ---- config -----------------------------------------------------------
Public Const NAME_CONFIG_DAY_START As String = "i_config_day_start"
Public Const NAME_CONFIG_DAY_END As String = "i_config_day_end"
Public Const NAME_CONFIG_STEP As String = "i_config_step"
Public Const NAME_CONFIG_DURATION As String = "i_config_duration"
Public Const NAME_CONFIG_CURRENCY As String = "i_config_currency"
Public Const NAME_CONFIG_NUMBER_FORMAT As String = "i_config_number_format"
Public Const NAME_CONFIG_NONWORK_DAYS As String = "i_config_nonwork_days"
Public Const NAME_CONFIG_NONWORK_DATES As String = "i_config_nonwork_dates"
Public Const NAME_STATUS_CONFIG As String = "o_status_config"
Public Const NAME_STATE_SCHEMA As String = "s_schema_version"
Public Const NAME_STATE_PERIOD_FROM As String = "s_period_from"
Public Const NAME_STATE_PERIOD_TO As String = "s_period_to"
Public Const NAME_BTN_CONFIG_APPLY As String = "b_config_apply"
Public Const NAME_BTN_CONFIG_DEFAULTS As String = "b_config_defaults"
Public Const NAME_BTN_LOCK_DATA As String = "b_lock_data"
Public Const NAME_BTN_UNLOCK_DATA As String = "b_unlock_data"
Public Const NAME_BTN_CONFIG_NONWORK_PICK As String = "b_config_nonwork_pick"
Public Const NAME_BTN_CONFIG_NONWORK_CLEAR As String = "b_config_nonwork_clear"

' ---- log, about, lists ------------------------------------------------
Public Const NAME_LOG_FIRST As String = "log_first"
Public Const NAME_ABOUT_VERSION As String = "o_about_version"
Public Const NAME_ABOUT_FILE As String = "o_about_file"
Public Const NAME_ABOUT_SHEETS As String = "o_about_sheets"
Public Const NAME_LIST_CLIENTS As String = "list_clients"
Public Const NAME_LIST_SERVICES As String = "list_services"
Public Const NAME_LIST_TIMES As String = "list_times"

' =====================================================================
'  procedures
' =====================================================================

' Every name the template must carry, in one list. VerifyContract walks it, and
' the list is built from the constants above, so there is one place to add a
' name.
'
' The text is built in several statements on purpose: the VBE accepts at most 24
' line continuations in one statement ("Too many line continuations").
Public Function DeclaredNames() As String
    Dim text As String

    text = NAME_VIEW & "," & NAME_PERIOD & "," & NAME_COUNT & "," & NAME_REVENUE & "," & _
           NAME_FREE_SLOTS & "," & NAME_STATUS_DASHBOARD & "," & NAME_CAL_TITLE & "," & _
           NAME_CAL_FIRST & "," & NAME_CAL_LAST & "," & NAME_STATE_VIEW & "," & NAME_STATE_DATE

    text = text & "," & NAME_BTN_VIEW_DAILY & "," & NAME_BTN_VIEW_WEEKLY & "," & _
           NAME_BTN_VIEW_MONTHLY & "," & NAME_BTN_PREV & "," & NAME_BTN_TODAY & "," & _
           NAME_BTN_NEXT & "," & NAME_BTN_REFRESH & "," & NAME_BTN_NEW_APPOINTMENT & "," & _
           NAME_BTN_GO_CLIENTS & "," & NAME_BTN_GO_SERVICES & "," & _
           NAME_BTN_GO_APPOINTMENTS & "," & NAME_BTN_GO_REPORTS & "," & _
           NAME_BTN_GO_CONFIG & "," & NAME_BTN_GO_ABOUT

    text = text & "," & NAME_BTN_HOME_BOOKING & "," & NAME_BTN_HOME_CLIENTS & "," & _
           NAME_BTN_HOME_SERVICES & "," & NAME_BTN_HOME_APPOINTMENTS & "," & _
           NAME_BTN_HOME_REPORTS & "," & _
           NAME_BTN_HOME_CONFIG & "," & NAME_BTN_HOME_ABOUT

    text = text & "," & NAME_BOOKING_DATE & "," & NAME_BOOKING_CLIENT & "," & _
           NAME_BOOKING_SERVICE & "," & NAME_BOOKING_START_H & "," & NAME_BOOKING_START_M & "," & _
           NAME_BOOKING_DUR_H & "," & NAME_BOOKING_DUR_M & "," & NAME_BOOKING_CLIENT_ID & "," & _
           NAME_BOOKING_PRICE & "," & NAME_BOOKING_SERVICE_DURATION & "," & NAME_BOOKING_START

    text = text & "," & NAME_BOOKING_DURATION & "," & NAME_BOOKING_END & "," & _
           NAME_BOOKING_MODE & "," & NAME_STATE_EDIT_ID & "," & NAME_STATUS_BOOKING & "," & _
           NAME_BTN_BOOKING_DATE_PICK & "," & NAME_BTN_BOOKING_START_PICK & "," & _
           NAME_BTN_BOOKING_DUR_PICK & "," & NAME_BTN_BOOKING_CLIENT_PICK & "," & _
           NAME_BTN_BOOKING_SAVE & "," & _
           NAME_BTN_BOOKING_CLEAR & "," & NAME_BTN_BOOKING_CANCEL

    text = text & "," & NAME_CLIENT_NAME & "," & NAME_CLIENT_PHONE & "," & _
           NAME_CLIENT_EMAIL & "," & NAME_CLIENT_COMMENT & "," & NAME_STATUS_CLIENTS & "," & _
           NAME_CLIENTS_TITLE & "," & NAME_BTN_CLIENTS_ADD & "," & NAME_BTN_CLIENTS_CLEAR & "," & _
           NAME_GRID_CLIENTS_FIRST & "," & NAME_GRID_CLIENTS_LAST & "," & NAME_STATE_EDIT_CLIENT

    text = text & "," & NAME_SERVICE_NAME & "," & NAME_SERVICE_PRICE & "," & _
           NAME_SERVICE_DURATION & "," & NAME_STATUS_SERVICES & "," & NAME_SERVICES_TITLE & "," & _
           NAME_BTN_SERVICES_ADD & "," & NAME_BTN_SERVICES_CLEAR & "," & _
           NAME_GRID_SERVICES_FIRST & "," & NAME_GRID_SERVICES_LAST & "," & NAME_STATE_EDIT_SERVICE

    text = text & "," & NAME_STATUS_APPOINTMENTS & "," & NAME_APPOINTMENTS_TITLE & "," & _
           NAME_BTN_APPOINTMENTS_REFRESH & "," & NAME_GRID_APPOINTMENTS_FIRST

    text = text & "," & NAME_REPORT_FROM & "," & NAME_REPORT_TO & "," & NAME_STATUS_REPORTS & "," & _
           NAME_REPORT_TITLE & "," & NAME_BTN_REPORT_RENDER & "," & NAME_BTN_REPORT_PREV & "," & _
           NAME_BTN_REPORT_NEXT & "," & NAME_GRID_REPORTS_FIRST & "," & NAME_GRID_REPORTS_LAST & "," & _
           NAME_BTN_REPORT_FROM_PICK & "," & NAME_BTN_REPORT_TO_PICK

    text = text & "," & NAME_CONFIG_DAY_START & "," & NAME_CONFIG_DAY_END & "," & _
           NAME_CONFIG_STEP & "," & NAME_CONFIG_DURATION & "," & NAME_CONFIG_CURRENCY & "," & _
           NAME_CONFIG_NUMBER_FORMAT & "," & NAME_CONFIG_NONWORK_DAYS & "," & _
           NAME_CONFIG_NONWORK_DATES & "," & NAME_STATUS_CONFIG & "," & NAME_STATE_SCHEMA & "," & _
           NAME_STATE_PERIOD_FROM & "," & NAME_STATE_PERIOD_TO

    text = text & "," & NAME_BTN_CONFIG_APPLY & "," & NAME_BTN_CONFIG_DEFAULTS & "," & _
           NAME_BTN_LOCK_DATA & "," & NAME_BTN_UNLOCK_DATA & "," & _
           NAME_BTN_CONFIG_NONWORK_PICK & "," & NAME_BTN_CONFIG_NONWORK_CLEAR

    text = text & "," & NAME_LOG_FIRST & "," & NAME_ABOUT_VERSION & "," & NAME_ABOUT_FILE & "," & _
           NAME_ABOUT_SHEETS & "," & NAME_LIST_CLIENTS & "," & NAME_LIST_SERVICES & "," & _
           NAME_LIST_TIMES

    DeclaredNames = text
End Function

' The range a name points to. Raises with a clear message when the name is
' missing, so the caller does not have to check twice. The message of the
' lookup itself is part of the error text: "missing" without it cost a whole
' build cycle once (PROJECT_CONCEPT §13).
Public Function RangeOfName(ByVal nameText As String) As Range
    Dim target As Range
    Dim lookupError As Long
    Dim lookupText As String

    lookupError = 0
    lookupText = vbNullString
    Err.Clear
    On Error Resume Next
    Set target = ThisWorkbook.Names(nameText).RefersToRange
    lookupError = Err.Number
    lookupText = Err.Description
    On Error GoTo 0
    Err.Clear

    If target Is Nothing Then
        ' The message has to survive a broken lookup, so it may only use things
        ' that cannot raise: the name, the sheet, the number of names in the book
        ' and what the failed lookup itself said. The probes that touch RefersTo
        ' (NameFound / NameRefersTo) must not be called here - they throw their own
        ' error when the name is missing, which is exactly this case (24.09.2026).
        Err.Raise 5, "modNames.RangeOfName", _
                  modText.TextFormat(modText.KEY_ERROR_NAME_MISSING, nameText, _
                                     SheetOfName(nameText)) & _
                  " [book=" & ThisWorkbook.Name & _
                  " names=" & CStr(ThisWorkbook.Names.Count) & _
                  " lookup=" & CStr(lookupError) & " " & lookupText & "]"
    End If

    Set RangeOfName = target
End Function

' What the name points at, as text (part of the error message above).
Private Function NameRefersTo(ByVal nameText As String) As String
    Dim text As String

    text = vbNullString
    Err.Clear
    On Error Resume Next
    text = CStr(ThisWorkbook.Names(nameText).RefersTo)
    On Error GoTo 0
    Err.Clear

    NameRefersTo = text
End Function

' "yes", "no" or the error of the name lookup itself: it tells a missing name
' from a name lookup that failed for another reason.
Private Function NameFound(ByVal nameText As String) As String
    Dim flag As String
    Dim found As Name

    flag = "no"
    Err.Clear
    On Error Resume Next
    Set found = ThisWorkbook.Names(nameText)
    On Error GoTo 0
    If Err.Number <> 0 Then
        flag = "err " & CStr(Err.Number)
    ElseIf Not found Is Nothing Then
        flag = "yes"
    End If
    Err.Clear

    NameFound = flag
End Function

' The text of a named cell, trimmed.
Public Function ReadName(ByVal nameText As String) As String
    ReadName = Trim$(CStr(RangeOfName(nameText).Value))
End Function

' The value of a named cell as it is written (a number stays a number).
Public Function ReadNameValue(ByVal nameText As String) As Variant
    ReadNameValue = RangeOfName(nameText).Value
End Function

' Writes a UI cell. This is the only place that writes one: the events are
' switched off while writing, so a write cannot call back the action that caused
' it. The previous state is restored, not forced to True: inside AppRouter the
' events are already off, and re-enabling them in the middle of an action would
' let a later write raise an event again.
Public Sub WriteName(ByVal nameText As String, ByVal value As Variant)
    Dim eventsWereOn As Boolean

    eventsWereOn = Application.EnableEvents
    Application.EnableEvents = False

    On Error GoTo Failed
    RangeOfName(nameText).Value = value
    GoTo Cleanup

Failed:
    Application.EnableEvents = eventsWereOn
    Err.Raise Err.Number, "modNames.WriteName", nameText & ": " & Err.Description

Cleanup:
    Application.EnableEvents = eventsWereOn
End Sub

' The sheet a name lives on ("" when the name does not exist).
Public Function SheetOfName(ByVal nameText As String) As String
    Dim sheetName As String

    sheetName = vbNullString
    Err.Clear
    On Error Resume Next
    sheetName = ThisWorkbook.Names(nameText).RefersToRange.Parent.Name
    On Error GoTo 0

    SheetOfName = sheetName
End Function

' The name of a cell, "" when the cell has none. The dynamic cells of the
' dashboard body have no name, so the event handlers fall back to the position.
Public Function NameOfCell(ByRef ws As Worksheet, ByRef target As Range) As String
    Dim definedName As Name
    Dim hit As String

    hit = vbNullString
    Err.Clear
    On Error Resume Next
    For Each definedName In ThisWorkbook.Names
        If StrComp(definedName.RefersToRange.Parent.Name, ws.Name, vbTextCompare) = 0 Then
            If StrComp(CStr(definedName.RefersToRange.Address), CStr(target.Address), _
                       vbTextCompare) = 0 Then
                hit = CStr(definedName.Name)
                Exit For
            End If
        End If
    Next definedName
    On Error GoTo 0
    Err.Clear

    NameOfCell = hit
End Function

' True when a name points to (a part of) the given cell.
Public Function IsNameAt(ByRef ws As Worksheet, ByRef target As Range, _
                         ByVal nameText As String) As Boolean
    Dim definedRange As Range

    Set definedRange = Nothing
    Err.Clear
    On Error Resume Next
    Set definedRange = ThisWorkbook.Names(nameText).RefersToRange
    On Error GoTo 0
    If definedRange Is Nothing Then Exit Function

    IsNameAt = Not (Intersect(target, definedRange) Is Nothing)
End Function

' The status cell of a sheet: every UI sheet has its own, so the name is unique.
Public Function StatusNameFor(ByVal sheetName As String) As String
    Select Case sheetName
        Case SHEET_DASHBOARD
            StatusNameFor = NAME_STATUS_DASHBOARD
        Case SHEET_BOOKING
            StatusNameFor = NAME_STATUS_BOOKING
        Case SHEET_CLIENTS_ADMIN
            StatusNameFor = NAME_STATUS_CLIENTS
        Case SHEET_SERVICES_ADMIN
            StatusNameFor = NAME_STATUS_SERVICES
        Case SHEET_APPOINTMENTS
            StatusNameFor = NAME_STATUS_APPOINTMENTS
        Case SHEET_REPORTS
            StatusNameFor = NAME_STATUS_REPORTS
        Case SHEET_CONFIG
            StatusNameFor = NAME_STATUS_CONFIG
        Case Else
            StatusNameFor = vbNullString
    End Select
End Function

' The button that belongs to a view ("day" -> b_view_daily) + the reverse.
Public Function ViewButtonName(ByVal viewKey As String) As String
    Select Case LCase$(Trim$(viewKey))
        Case VIEW_DAY
            ViewButtonName = NAME_BTN_VIEW_DAILY
        Case VIEW_WEEK
            ViewButtonName = NAME_BTN_VIEW_WEEKLY
        Case VIEW_MONTH
            ViewButtonName = NAME_BTN_VIEW_MONTHLY
        Case Else
            ViewButtonName = vbNullString
    End Select
End Function

Public Function ViewKeyOfButton(ByVal buttonName As String) As String
    Select Case LCase$(Trim$(buttonName))
        Case LCase$(NAME_BTN_VIEW_DAILY)
            ViewKeyOfButton = VIEW_DAY
        Case LCase$(NAME_BTN_VIEW_WEEKLY)
            ViewKeyOfButton = VIEW_WEEK
        Case LCase$(NAME_BTN_VIEW_MONTHLY)
            ViewKeyOfButton = VIEW_MONTH
        Case Else
            ViewKeyOfButton = vbNullString
    End Select
End Function

' The text of a cell (used for the dynamic cells of the dashboard body).
Public Function CellText(ByRef ws As Worksheet, ByVal rowIndex As Long, _
                         ByVal columnIndex As Long) As String
    CellText = Trim$(CStr(ws.Cells(rowIndex, columnIndex).Text))
End Function

' True for a name the application declares.
Public Function IsDeclaredName(ByVal nameText As String) As Boolean
    IsDeclaredName = (InStr(1, "," & DeclaredNames() & ",", "," & nameText & ",", _
                            vbTextCompare) > 0)
End Function

' =====================================================================
'  contract check
' =====================================================================
' Everything the workbook must provide. One problem per line, "" when the
' template and the code agree; app.diagnostics and Test_Contract walk it.
Public Function VerifyContract() As String
    Dim problems As String
    Dim sheetList As Variant
    Dim nameList As Variant
    Dim item As Variant
    Dim definedName As Name
    Dim target As Range
    Dim itemText As String

    ' 1. the sheets
    sheetList = Split(modConfig.UiSheetNames() & "," & modConfig.DataSheetNames() & "," & _
                      SHEET_LOG & "," & SHEET_LISTS, ",")
    For Each item In sheetList
        itemText = Trim$(CStr(item))
        If Not SheetExists(itemText) Then
            problems = problems & modText.TextFormat(modText.KEY_ERROR_NO_SHEET, itemText) & vbCrLf
        End If
    Next item

    ' 2. every declared name exists and points to one cell (the dropdown lists
    '    are ranges on purpose)
    nameList = Split(DeclaredNames(), ",")
    For Each item In nameList
        itemText = Trim$(CStr(item))
        Set target = Nothing
        Err.Clear
        On Error Resume Next
        Set target = ThisWorkbook.Names(itemText).RefersToRange
        On Error GoTo 0

        If target Is Nothing Then
            problems = problems & modText.TextFormat(modText.KEY_ERROR_NAME_MISSING, _
                                                     itemText, vbNullString) & vbCrLf
        ElseIf target.Cells.Count > 1 Then
            If InStr(1, itemText, PREFIX_LIST, vbTextCompare) <> 1 Then
                problems = problems & modText.TextFormat(modText.KEY_ERROR_NAME_NOT_CELL, _
                                                         itemText) & vbCrLf
            End If
        End If
    Next item

    ' 3. names with our prefixes that the code does not know: the template was
    '    edited (renamed, added) without telling the code
    Err.Clear
    On Error Resume Next
    For Each definedName In ThisWorkbook.Names
        itemText = CStr(definedName.Name)
        If HasOurPrefix(itemText) Then
            If Not IsDeclaredName(itemText) Then
                problems = problems & "'" & itemText & "' is not declared in modNames." & vbCrLf
            End If
        End If
    Next definedName
    On Error GoTo 0
    Err.Clear

    VerifyContract = problems
End Function

' True for a name that starts with one of the prefixes of the application.
Public Function HasOurPrefix(ByVal nameText As String) As Boolean
    Dim prefixList As Variant
    Dim item As Variant

    prefixList = Split(PREFIX_BUTTON & "," & PREFIX_INPUT & "," & PREFIX_OUTPUT & "," & _
                       PREFIX_STATE & "," & PREFIX_LIST & "," & PREFIX_CALENDAR & "," & _
                       PREFIX_GRID & "," & PREFIX_LOG, ",")
    For Each item In prefixList
        If InStr(1, nameText, CStr(item), vbTextCompare) = 1 Then
            HasOurPrefix = True
            Exit Function
        End If
    Next item
End Function

' True when the sheet exists in this workbook.
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    If Len(sheetName) = 0 Then Exit Function

    Set ws = Nothing
    Err.Clear
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not (ws Is Nothing)
End Function
