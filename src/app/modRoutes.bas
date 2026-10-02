Attribute VB_Name = "modRoutes"
Option Explicit

' =====================================================================
'  Service_Manager - the action register (Phase 1)
'
'  One place that says what the application can do: an action name ("domain.verb")
'  is mapped to a procedure, and a cell button is mapped to an action plus its
'  payload. AppRouter uses the first table, the event handlers the second, and
'  VerifyRoutes checks that every row points to a procedure that exists (a typo
'  must be found by a test, not by the first click of the user).
' =====================================================================

' ---- action names ----------------------------------------------------
Public Const ACTION_APP_REFRESH As String = "app.refresh"
Public Const ACTION_APP_NAVIGATE As String = "app.navigate"
Public Const ACTION_APP_DIAGNOSTICS As String = "app.diagnostics"
Public Const ACTION_APP_ABOUT As String = "app.about"
Public Const ACTION_DASHBOARD_VIEW As String = "dashboard.view"
Public Const ACTION_DASHBOARD_MOVE As String = "dashboard.move"
Public Const ACTION_DASHBOARD_OPEN_DAY As String = "dashboard.open_day"
Public Const ACTION_CLIENTS_ADD As String = "clients.add"
Public Const ACTION_CLIENTS_DELETE As String = "clients.delete"
Public Const ACTION_CLIENTS_LIST As String = "clients.list"
' The [Edit] link of a row in the two admin lists (§6.10): it loads that row into
' the form above; the save then goes through clients.add / services.add, which see
' the state cell and update instead of inserting.
Public Const ACTION_CLIENTS_EDIT As String = "clients.edit"
Public Const ACTION_SERVICES_EDIT As String = "services.edit"
Public Const ACTION_SERVICES_ADD As String = "services.add"
Public Const ACTION_SERVICES_DELETE As String = "services.delete"
Public Const ACTION_SERVICES_LIST As String = "services.list"
Public Const ACTION_APPOINTMENTS_LIST As String = "appointments.list"
Public Const ACTION_FORM_CLEAR As String = "form.clear"
Public Const ACTION_APPOINTMENTS_SAVE As String = "appointments.save"
Public Const ACTION_APPOINTMENTS_NEW As String = "appointments.new"
Public Const ACTION_APPOINTMENTS_EDIT As String = "appointments.edit"
Public Const ACTION_APPOINTMENTS_DELETE As String = "appointments.delete"
Public Const ACTION_BOOKING_CANCEL As String = "booking.cancel"
Public Const ACTION_FORM_PICK_DATE As String = "form.pick_date"
Public Const ACTION_FORM_PICK_TIME As String = "form.pick_time"
' The typed client search on Booking (§9.5).
Public Const ACTION_FORM_PICK_CLIENT As String = "form.pick_client"
Public Const ACTION_REPORTS_RENDER As String = "reports.render"
' The date pickers of the report period (From / To) - the Booking page keeps
' form.pick_date, because its picker fills the form and its previews.
Public Const ACTION_REPORTS_PICK_DATE As String = "reports.pick_date"
Public Const ACTION_CONFIG_APPLY As String = "config.apply"
Public Const ACTION_CONFIG_DEFAULTS As String = "config.defaults"
Public Const ACTION_CONFIG_LOCK_DATA As String = "config.lock_data"
Public Const ACTION_CONFIG_UNLOCK_DATA As String = "config.unlock_data"
Public Const ACTION_CONFIG_PICK_NONWORK As String = "config.pick_nonwork"
Public Const ACTION_CONFIG_CLEAR_NONWORK As String = "config.clear_nonwork"

' ---- payload keys ----------------------------------------------------
Public Const KEY_VIEW As String = "view"
Public Const KEY_STEP As String = "step"
Public Const KEY_SHEET As String = "sheet"
Public Const KEY_PATH As String = "path"
Public Const KEY_ID As String = "id"
Public Const KEY_DATE As String = "date"
Public Const KEY_FROM As String = "from"
Public Const KEY_TO As String = "to"
Public Const KEY_DRY As String = "dry"
Public Const KEY_CELL As String = "cell"
Public Const KEY_ACTION As String = "action"
Public Const KEY_FIELD As String = "field"
Public Const KEY_VALUE As String = "value"

' ---- result keys -----------------------------------------------------
Public Const RESULT_ACTION As String = "action"
Public Const RESULT_OK As String = "ok"
Public Const RESULT_MESSAGE As String = "message"
Public Const RESULT_DATA As String = "data"
Public Const RESULT_ERROR As String = "error"

' ---- the tables (built once per session) -----------------------------
Private mRoutes As Object
Private mClickActions As Object

' =====================================================================
'  procedures
' =====================================================================

' action -> "module.Procedure"
Private Function RouteTable() As Object
    If mRoutes Is Nothing Then
        Set mRoutes = CreateObject("Scripting.Dictionary")
        mRoutes.Add ACTION_APP_REFRESH, "modApp.Refresh"
        mRoutes.Add ACTION_APP_NAVIGATE, "modApp.Navigate"
        mRoutes.Add ACTION_APP_DIAGNOSTICS, "modApp.ActionDiagnostics"
        mRoutes.Add ACTION_APP_ABOUT, "modApp.ShowAbout"
        mRoutes.Add ACTION_APPOINTMENTS_LIST, "modUiAppointments.OnList"
        mRoutes.Add ACTION_DASHBOARD_VIEW, "modDashboard.SetView"
        mRoutes.Add ACTION_DASHBOARD_MOVE, "modDashboard.Move"
        mRoutes.Add ACTION_DASHBOARD_OPEN_DAY, "modDashboard.OpenDay"
        mRoutes.Add ACTION_CLIENTS_ADD, "modUiClients.OnAdd"
        mRoutes.Add ACTION_CLIENTS_DELETE, "modClients.Delete"
        mRoutes.Add ACTION_CLIENTS_LIST, "modUiClients.OnList"
        mRoutes.Add ACTION_CLIENTS_EDIT, "modUiClients.OnEdit"
        mRoutes.Add ACTION_SERVICES_ADD, "modUiServices.OnAdd"
        mRoutes.Add ACTION_SERVICES_DELETE, "modServices.Delete"
        mRoutes.Add ACTION_SERVICES_LIST, "modUiServices.OnList"
        mRoutes.Add ACTION_SERVICES_EDIT, "modUiServices.OnEdit"
        mRoutes.Add ACTION_FORM_CLEAR, "modUi.FormClear"
        mRoutes.Add ACTION_APPOINTMENTS_SAVE, "modUiBooking.OnSave"
        mRoutes.Add ACTION_APPOINTMENTS_NEW, "modAppointments.NewBooking"
        mRoutes.Add ACTION_APPOINTMENTS_EDIT, "modAppointments.EditBooking"
        mRoutes.Add ACTION_APPOINTMENTS_DELETE, "modAppointments.DeleteBooking"
        mRoutes.Add ACTION_BOOKING_CANCEL, "modUiBooking.OnCancel"
        mRoutes.Add ACTION_FORM_PICK_DATE, "modUiBooking.OnPickDate"
        mRoutes.Add ACTION_FORM_PICK_TIME, "modUiBooking.OnPickTime"
        mRoutes.Add ACTION_FORM_PICK_CLIENT, "modUiBooking.OnPickClient"
        mRoutes.Add ACTION_REPORTS_RENDER, "modReports.Render"
        mRoutes.Add ACTION_REPORTS_PICK_DATE, "modReports.OnPickDate"
        mRoutes.Add ACTION_CONFIG_APPLY, "modUiConfig.Apply"
        mRoutes.Add ACTION_CONFIG_DEFAULTS, "modUiConfig.Defaults"
        mRoutes.Add ACTION_CONFIG_LOCK_DATA, "modProtection.HideDataSheets"
        mRoutes.Add ACTION_CONFIG_UNLOCK_DATA, "modProtection.ShowDataSheets"
        mRoutes.Add ACTION_CONFIG_PICK_NONWORK, "modUiConfig.OnPickNonwork"
        mRoutes.Add ACTION_CONFIG_CLEAR_NONWORK, "modUiConfig.OnClearNonwork"
    End If

    Set RouteTable = mRoutes
End Function

' cell name -> action
Private Function ClickTable() As Object
    If mClickActions Is Nothing Then
        Set mClickActions = CreateObject("Scripting.Dictionary")
        mClickActions.Add NAME_BTN_VIEW_DAILY, ACTION_DASHBOARD_VIEW
        mClickActions.Add NAME_BTN_VIEW_WEEKLY, ACTION_DASHBOARD_VIEW
        mClickActions.Add NAME_BTN_VIEW_MONTHLY, ACTION_DASHBOARD_VIEW
        mClickActions.Add NAME_BTN_PREV, ACTION_DASHBOARD_MOVE
        mClickActions.Add NAME_BTN_TODAY, ACTION_DASHBOARD_MOVE
        mClickActions.Add NAME_BTN_NEXT, ACTION_DASHBOARD_MOVE
        mClickActions.Add NAME_BTN_REFRESH, ACTION_APP_REFRESH
        mClickActions.Add NAME_BTN_GO_CLIENTS, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_GO_SERVICES, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_GO_APPOINTMENTS, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_GO_REPORTS, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_GO_CONFIG, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_GO_ABOUT, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_HOME_BOOKING, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_HOME_CLIENTS, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_HOME_SERVICES, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_HOME_APPOINTMENTS, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_HOME_REPORTS, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_HOME_CONFIG, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_HOME_ABOUT, ACTION_APP_NAVIGATE
        mClickActions.Add NAME_BTN_CLIENTS_ADD, ACTION_CLIENTS_ADD
        mClickActions.Add NAME_BTN_CLIENTS_CLEAR, ACTION_FORM_CLEAR
        mClickActions.Add NAME_BTN_SERVICES_ADD, ACTION_SERVICES_ADD
        mClickActions.Add NAME_BTN_SERVICES_CLEAR, ACTION_FORM_CLEAR
        mClickActions.Add NAME_BTN_APPOINTMENTS_REFRESH, ACTION_APPOINTMENTS_LIST
        mClickActions.Add NAME_BTN_NEW_APPOINTMENT, ACTION_APPOINTMENTS_NEW
        mClickActions.Add NAME_BTN_BOOKING_SAVE, ACTION_APPOINTMENTS_SAVE
        mClickActions.Add NAME_BTN_BOOKING_CLEAR, ACTION_FORM_CLEAR
        mClickActions.Add NAME_BTN_BOOKING_CANCEL, ACTION_BOOKING_CANCEL
        mClickActions.Add NAME_BTN_BOOKING_DATE_PICK, ACTION_FORM_PICK_DATE
        mClickActions.Add NAME_BTN_BOOKING_START_PICK, ACTION_FORM_PICK_TIME
        mClickActions.Add NAME_BTN_BOOKING_DUR_PICK, ACTION_FORM_PICK_TIME
        mClickActions.Add NAME_BTN_BOOKING_CLIENT_PICK, ACTION_FORM_PICK_CLIENT
        mClickActions.Add NAME_BTN_REPORT_RENDER, ACTION_REPORTS_RENDER
        mClickActions.Add NAME_BTN_REPORT_PREV, ACTION_REPORTS_RENDER
        mClickActions.Add NAME_BTN_REPORT_NEXT, ACTION_REPORTS_RENDER
        mClickActions.Add NAME_BTN_REPORT_FROM_PICK, ACTION_REPORTS_PICK_DATE
        mClickActions.Add NAME_BTN_REPORT_TO_PICK, ACTION_REPORTS_PICK_DATE
        mClickActions.Add NAME_BTN_CONFIG_APPLY, ACTION_CONFIG_APPLY
        mClickActions.Add NAME_BTN_CONFIG_DEFAULTS, ACTION_CONFIG_DEFAULTS
        mClickActions.Add NAME_BTN_LOCK_DATA, ACTION_CONFIG_LOCK_DATA
        mClickActions.Add NAME_BTN_UNLOCK_DATA, ACTION_CONFIG_UNLOCK_DATA
        mClickActions.Add NAME_BTN_CONFIG_NONWORK_PICK, ACTION_CONFIG_PICK_NONWORK
        mClickActions.Add NAME_BTN_CONFIG_NONWORK_CLEAR, ACTION_CONFIG_CLEAR_NONWORK
    End If

    Set ClickTable = mClickActions
End Function

' The procedure behind an action, "" for an unknown one.
Public Function RouteFor(ByVal action As String) As String
    Dim key As String

    key = LCase$(Trim$(action))
    If RouteTable().Exists(key) Then RouteFor = CStr(RouteTable()(key))
End Function

' The action behind a cell button, "" when the cell is not a button yet.
Public Function ActionForCell(ByVal cellName As String) As String
    Dim key As String

    key = LCase$(Trim$(cellName))
    If ClickTable().Exists(key) Then ActionForCell = CStr(ClickTable()(key))
End Function

Public Function IsClickAction(ByVal cellName As String) As Boolean
    IsClickAction = (Len(ActionForCell(cellName)) > 0)
End Function

' The payload a cell button sends with its action (never Nothing).
Public Function PayloadForCell(ByVal cellName As String) As Object
    Dim payload As Object
    Dim key As String

    Set payload = CreateObject("Scripting.Dictionary")
    key = LCase$(Trim$(cellName))

    Select Case key
        Case LCase$(NAME_BTN_VIEW_DAILY)
            payload(KEY_VIEW) = VIEW_DAY
        Case LCase$(NAME_BTN_VIEW_WEEKLY)
            payload(KEY_VIEW) = VIEW_WEEK
        Case LCase$(NAME_BTN_VIEW_MONTHLY)
            payload(KEY_VIEW) = VIEW_MONTH
        Case LCase$(NAME_BTN_PREV)
            payload(KEY_STEP) = "-1"
        Case LCase$(NAME_BTN_TODAY)
            payload(KEY_STEP) = "0"
        Case LCase$(NAME_BTN_NEXT)
            payload(KEY_STEP) = "1"
        Case LCase$(NAME_BTN_GO_CLIENTS)
            payload(KEY_SHEET) = SHEET_CLIENTS_ADMIN
        Case LCase$(NAME_BTN_GO_SERVICES)
            payload(KEY_SHEET) = SHEET_SERVICES_ADMIN
        Case LCase$(NAME_BTN_GO_APPOINTMENTS)
            payload(KEY_SHEET) = SHEET_APPOINTMENTS
        Case LCase$(NAME_BTN_GO_REPORTS)
            payload(KEY_SHEET) = SHEET_REPORTS
        Case LCase$(NAME_BTN_GO_CONFIG)
            payload(KEY_SHEET) = SHEET_CONFIG
        Case LCase$(NAME_BTN_GO_ABOUT)
            payload(KEY_SHEET) = SHEET_ABOUT
        ' The "Dashboard" buttons of the other sheets (the tip in the template says
        ' "Back to the Dashboard"): home is always the Dashboard, whichever sheet the
        ' button sits on. They used to send their own sheet, so a click did nothing.
        Case LCase$(NAME_BTN_HOME_BOOKING), LCase$(NAME_BTN_HOME_CLIENTS), _
             LCase$(NAME_BTN_HOME_SERVICES), LCase$(NAME_BTN_HOME_APPOINTMENTS), _
             LCase$(NAME_BTN_HOME_REPORTS), _
             LCase$(NAME_BTN_HOME_CONFIG), LCase$(NAME_BTN_HOME_ABOUT)
            payload(KEY_SHEET) = SHEET_DASHBOARD
        Case LCase$(NAME_BTN_CLIENTS_CLEAR)
            payload(KEY_SHEET) = SHEET_CLIENTS_ADMIN
        Case LCase$(NAME_BTN_SERVICES_CLEAR)
            payload(KEY_SHEET) = SHEET_SERVICES_ADMIN
        Case LCase$(NAME_BTN_BOOKING_CLEAR)
            payload(KEY_SHEET) = SHEET_BOOKING
        Case LCase$(NAME_BTN_BOOKING_CANCEL)
            ' the same cell is Cancel (a new appointment) and Delete (while one is
            ' being edited): the cell travels with the payload, so the first click
            ' of the confirmation can paint it (§6.6)
            payload(KEY_SHEET) = SHEET_BOOKING
            payload(KEY_CELL) = NAME_BTN_BOOKING_CANCEL
        Case LCase$(NAME_BTN_BOOKING_DATE_PICK)
            payload(KEY_CELL) = NAME_BOOKING_DATE
        Case LCase$(NAME_BTN_BOOKING_START_PICK)
            payload(KEY_FIELD) = FIELD_START
        Case LCase$(NAME_BTN_BOOKING_DUR_PICK)
            payload(KEY_FIELD) = FIELD_DURATION
        Case LCase$(NAME_BTN_REPORT_PREV)
            payload(KEY_STEP) = "-1"
        Case LCase$(NAME_BTN_REPORT_NEXT)
            payload(KEY_STEP) = "1"
        Case LCase$(NAME_BTN_REPORT_FROM_PICK)
            payload(KEY_CELL) = NAME_REPORT_FROM
        Case LCase$(NAME_BTN_REPORT_TO_PICK)
            payload(KEY_CELL) = NAME_REPORT_TO
    End Select

    Set PayloadForCell = payload
End Function

' The keys of a payload, as text (diagnostics and tests, no arrays cross a
' module boundary).
Public Function PayloadText(ByRef payload As Object) As String
    Dim key As Variant
    Dim text As String

    If payload Is Nothing Then Exit Function
    For Each key In payload.Keys
        If Len(text) > 0 Then text = text & ", "
        text = text & CStr(key) & "=" & CStr(payload(CStr(key)))
    Next key

    PayloadText = text
End Function

' "action -> module.Procedure" for every route (diagnostics, README, tests).
Public Function RouteLines() As Collection
    Dim lines As Collection
    Dim key As Variant

    Set lines = New Collection
    For Each key In RouteTable().Keys
        lines.Add CStr(key) & " -> " & CStr(RouteTable()(CStr(key)))
    Next key

    Set RouteLines = lines
End Function

' The cell buttons that have no action yet, one name per line. The list is the
' work that is still open, phase by phase, and a test asserts it exactly.
Public Function PendingButtons() As String
    Dim names As Variant
    Dim item As Variant
    Dim text As String
    Dim nameText As String

    names = Split(modNames.DeclaredNames(), ",")
    For Each item In names
        nameText = Trim$(CStr(item))
        If InStr(1, nameText, PREFIX_BUTTON, vbTextCompare) = 1 Then
            If Len(ActionForCell(nameText)) = 0 Then
                text = text & nameText & vbCrLf
            End If
        End If
    Next item

    PendingButtons = text
End Function

' Problems of the routing table: a route that points to a procedure which does
' not exist. "" when everything is fine.
Public Function VerifyRoutes() As String
    Dim problems As String
    Dim key As Variant
    Dim target As String

    For Each key In RouteTable().Keys
        target = CStr(RouteTable()(CStr(key)))
        If Not ProcedureExists(target) Then
            problems = problems & modText.TextFormat(modText.KEY_ERROR_ROUTE, _
                                                     CStr(key), target) & vbCrLf
        End If
    Next key

    VerifyRoutes = problems
End Function

' True when "module.Procedure" exists in this project. The VBA project is read
' through VBComponents; when it is not accessible (a machine without VBOM
' access), the check says yes and the routes are trusted.
Private Function ProcedureExists(ByVal target As String) As Boolean
    Dim dot As Long
    Dim moduleName As String
    Dim procedureName As String
    Dim component As Object
    Dim codeModule As Object
    Dim lineIndex As Long
    Dim lineText As String
    Dim searchText As String

    ProcedureExists = True
    dot = InStr(1, target, ".", vbBinaryCompare)
    If dot = 0 Then Exit Function

    moduleName = Left$(target, dot - 1)
    procedureName = Mid$(target, dot + 1)

    Set component = Nothing
    Err.Clear
    On Error Resume Next
    Set component = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0
    If Err.Number <> 0 Then Exit Function
    If component Is Nothing Then
        ProcedureExists = False
        Exit Function
    End If

    Set codeModule = component.CodeModule
    searchText = "Sub " & procedureName & "("
    For lineIndex = 1 To codeModule.CountOfLines
        lineText = CStr(codeModule.Lines(lineIndex, 1))
        If InStr(1, lineText, searchText, vbTextCompare) > 0 Or _
           InStr(1, lineText, "Function " & procedureName & "(", vbTextCompare) > 0 Then
            Exit Function
        End If
    Next lineIndex

    ProcedureExists = False
End Function
