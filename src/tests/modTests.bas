Attribute VB_Name = "modTests"
Option Explicit

' =====================================================================
'  Service_Manager - the tests of phase 1
'
'  Every test enters through AppRouter (or through the event handler the user
'  triggers), never through a domain procedure: what is not reachable through
'  the router does not exist for the user either.
'
'  Run it without opening Excel by hand:
'      cscript //nologo tools\build.vbs test
' =====================================================================

Public Function RunAllTests() As String
    Dim started As Single
    Dim result As Object

    AssertReset
    started = Timer

    modUi.SetInteractiveUi False        ' a test run never shows a dialog

    ' Every test runs even when one of them raises: the report lists them all
    ' instead of stopping at the first problem with a modal dialog.
    On Error Resume Next

    Test_Contract
    CheckError "Test_Contract"
    Test_Routes
    CheckError "Test_Routes"
    Test_PendingButtons
    CheckError "Test_PendingButtons"

    Test_ClientSearch
    CheckError "Test_ClientSearch"

    Test_StartPickerWindow
    CheckError "Test_StartPickerWindow"

    Test_PhoneStaysText
    CheckError "Test_PhoneStaysText"

    Test_ConfirmDialog
    CheckError "Test_ConfirmDialog"

    Test_AppointmentsOrder
    CheckError "Test_AppointmentsOrder"
    Test_TextCoverage
    CheckError "Test_TextCoverage"
    Test_AppRouterBasics
    CheckError "Test_AppRouterBasics"
    Test_NamesAccess
    CheckError "Test_NamesAccess"
    Test_StatusCells
    CheckError "Test_StatusCells"
    Test_TimeHelpers
    CheckError "Test_TimeHelpers"
    Test_MoneyHelper
    CheckError "Test_MoneyHelper"
    Test_SheetEvents
    CheckError "Test_SheetEvents"
    Test_HyperlinkClick
    CheckError "Test_HyperlinkClick"
    Test_SheetEventsAttached
    CheckError "Test_SheetEventsAttached"

    TraceLine "step: app.refresh"
    Set result = modApp.AppRouter(modRoutes.ACTION_APP_REFRESH)
    CheckError "AppRouter(app.refresh)"

    Test_Refresh result
    CheckError "Test_Refresh"
    Test_DashboardViews result
    CheckError "Test_DashboardViews"

    ' ---- phase 2: clients and services ------------------------------------
    ' The copies are switched off for the run (a test deletes rows on purpose);
    ' Test_Backup switches them on again for its own checks.
    modBackup.SetEnabled False

    Test_ClientsAdd
    CheckError "Test_ClientsAdd"
    Test_ClientsValidation
    CheckError "Test_ClientsValidation"
    Test_ClientsDelete
    CheckError "Test_ClientsDelete"
    Test_ClientsDeleteGuard
    CheckError "Test_ClientsDeleteGuard"
    Test_ClientsList
    CheckError "Test_ClientsList"
    Test_ServicesAdd
    CheckError "Test_ServicesAdd"
    Test_ServicesValidation
    CheckError "Test_ServicesValidation"
    Test_ServicesDelete
    CheckError "Test_ServicesDelete"
    Test_ServicesDeleteGuard
    CheckError "Test_ServicesDeleteGuard"
    Test_ServicesList
    CheckError "Test_ServicesList"
    Test_ListsAndClear
    CheckError "Test_ListsAndClear"
    Test_Backup
    CheckError "Test_Backup"
    Test_HyperlinkWithoutSubAddress
    CheckError "Test_HyperlinkWithoutSubAddress"

    ' ---- phase 3: appointments, the booking page, the pickers --------------
    ' No dialogs and no copies while the tests create and remove appointments;
    ' the pickers are driven through data("value") instead (??9.4).
    modBackup.SetEnabled False
    modUi.SetInteractiveUi False
    modApp.ConfirmReset

    Test_AppointmentTimes
    CheckError "Test_AppointmentTimes"
    Test_PickersPure
    CheckError "Test_PickersPure"
    Test_PickerForms
    CheckError "Test_PickerForms"
    Test_BookingForm
    CheckError "Test_BookingForm"
    Test_BookingDropdowns
    CheckError "Test_BookingDropdowns"
    Test_AppointmentsSave
    CheckError "Test_AppointmentsSave"
    Test_AppointmentValidation
    CheckError "Test_AppointmentValidation"
    Test_Conflicts
    CheckError "Test_Conflicts"
    Test_AppointmentEdit
    CheckError "Test_AppointmentEdit"
    Test_AppointmentsList
    CheckError "Test_AppointmentsList"
    Test_AppointmentDelete
    CheckError "Test_AppointmentDelete"
    Test_BookingFormFlow
    CheckError "Test_BookingFormFlow"
    Test_ServiceFillsDuration
    CheckError "Test_ServiceFillsDuration"
    Test_PickActions
    CheckError "Test_PickActions"

    modBackup.SetEnabled False
    modUi.SetInteractiveUi False
    modApp.ConfirmReset

    ' ---- phase 4: the body of the dashboard --------------------------------
    modBackup.SetEnabled False
    modUi.SetInteractiveUi False
    modApp.ConfirmReset

    Test_DashboardDaily
    CheckError "Test_DashboardDaily"
    Test_DashboardWeekly
    CheckError "Test_DashboardWeekly"
    Test_DashboardMonthly
    CheckError "Test_DashboardMonthly"
    Test_DashboardOpenDay
    CheckError "Test_DashboardOpenDay"
    Test_DashboardRenderTimes
    CheckError "Test_DashboardRenderTimes"
    Test_DashboardResolution
    CheckError "Test_DashboardResolution"

    ' ---- phase 5: the reports ----------------------------------------------
    Test_Reports
    CheckError "Test_Reports"
    Test_ReportPickers
    CheckError "Test_ReportPickers"

    ' ---- phase 6: the Config page ------------------------------------------
    Test_ConfigPage
    CheckError "Test_ConfigPage"
    Test_NonworkDates
    CheckError "Test_NonworkDates"
    Test_DataSheetVisibility
    CheckError "Test_DataSheetVisibility"
    Test_Schema
    CheckError "Test_Schema"
    Test_Performance
    CheckError "Test_Performance"

    Test_Log
    CheckError "Test_Log"
    Test_NoFormControls
    CheckError "Test_NoFormControls"
    Test_Diagnostics
    CheckError "Test_Diagnostics"

    ' ---- the cross check of tools\sim --------------------------------------
    ' The script model answers the same cases of src\tests\crosscheck.txt; this
    ' test writes the answers of the workbook next to them and tools\build.vbs
    ' compares the two files. It comes last, because it puts the fixtures of that
    ' file into the workbook - the tests before it work on the sample data.
    Test_CrossCheck
    CheckError "Test_CrossCheck"

    On Error GoTo 0

    modUi.SetInteractiveUi True
    modUi.ResetStatuses

    Note "pending cell buttons: " & CStr(modApp.PendingButtonCount())
    Note "elapsed: " & FormatElapsed(Timer - started)

    RunAllTests = AssertSummary()
End Function

' Turns the error of a test into a failing assertion, so the report names it.
Private Sub CheckError(ByVal testName As String)
    ' The trace names every test that finished: a run that hangs or dies shows the
    ' last one that came through, which is where to look (a dialog, a loop).
    TraceLine "test: " & testName

    If Err.Number = 0 Then Exit Sub

    FailTest testName & " raised " & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
End Sub

' ---- the workbook agrees with the code -------------------------------

Private Sub Test_Contract()
    Dim problems As String

    problems = modNames.VerifyContract()
    If Len(problems) > 0 Then Note "contract problems: " & Replace(problems, vbCrLf, " | ")
    AssertEqualsText problems, vbNullString, "modNames.VerifyContract reports no problem"

    AssertTrue modNames.SheetExists(SHEET_DASHBOARD), "Dashboard exists"
    AssertTrue modNames.SheetExists(SHEET_DB_APPOINTMENTS), "db_appointments exists"
    AssertEqualsLong modNames.RangeOfName(NAME_CAL_FIRST).Row, 10, "cal_first is row 10"
    AssertEqualsText modNames.SheetOfName(NAME_STATE_VIEW), SHEET_DASHBOARD, _
                     "s_view lives on the Dashboard"
End Sub

Private Sub Test_Routes()
    Dim problems As String

    problems = modRoutes.VerifyRoutes()
    If Len(problems) > 0 Then Note "route problems: " & Replace(problems, vbCrLf, " | ")
    AssertEqualsText problems, vbNullString, "every route points to a procedure"

    AssertEqualsText modRoutes.RouteFor(modRoutes.ACTION_APP_REFRESH), "modApp.Refresh", _
                     "app.refresh -> modApp.Refresh"
    AssertEqualsText modRoutes.RouteFor("nonsense"), vbNullString, "an unknown action has no route"

    AssertEqualsText modRoutes.ActionForCell(NAME_BTN_VIEW_WEEKLY), _
                     modRoutes.ACTION_DASHBOARD_VIEW, "b_view_weekly is a cell button"
    AssertTrue modRoutes.IsClickAction(NAME_BTN_REPORT_RENDER), _
               "b_report_render runs the report"

    AssertEqualsText modRoutes.PayloadText(modRoutes.PayloadForCell(NAME_BTN_VIEW_MONTHLY)), _
                     "view=month", "the monthly button sends view=month"
    AssertEqualsText modRoutes.PayloadText(modRoutes.PayloadForCell(NAME_BTN_NEXT)), _
                     "step=1", "the next button sends step=1"
    ' Every sheet carries a "Dashboard" button and every one of them goes home,
    ' whichever sheet it sits on (the tip of the button says "Back to the Dashboard").
    Dim homeNames As Variant
    Dim homeIndex As Long
    homeNames = Array(NAME_BTN_HOME_BOOKING, NAME_BTN_HOME_CLIENTS, _
                      NAME_BTN_HOME_SERVICES, NAME_BTN_HOME_APPOINTMENTS, NAME_BTN_HOME_REPORTS, _
                      NAME_BTN_HOME_CONFIG, NAME_BTN_HOME_ABOUT)
    For homeIndex = 0 To UBound(homeNames)
        AssertEqualsText modRoutes.PayloadText( _
                         modRoutes.PayloadForCell(CStr(homeNames(homeIndex)))), _
                         "sheet=Dashboard", _
                         CStr(homeNames(homeIndex)) & " takes the user to the Dashboard"
    Next homeIndex

    ' And a real click on one of them (from the Reports page, the last one added).
    modUi.Navigate SHEET_REPORTS
    modApp.RunCellAction SHEET_REPORTS, NAME_BTN_HOME_REPORTS
    AssertEqualsText ActiveSheet.Name, SHEET_DASHBOARD, "the click really comes home"
    AssertEqualsText modRoutes.PayloadText(modRoutes.PayloadForCell(NAME_BTN_REPORT_NEXT)), _
                     "step=1", "the report's next button sends step=1"
    AssertEqualsText modRoutes.PayloadText(modRoutes.PayloadForCell(NAME_BTN_REPORT_PREV)), _
                     "step=-1", "and the previous one step=-1"
    AssertEqualsText modRoutes.PayloadText(modRoutes.PayloadForCell(NAME_BTN_GO_APPOINTMENTS)), _
                     "sheet=Appointments", "the Dashboard opens the Appointments page"
End Sub

Private Sub Test_PendingButtons()
    Dim listText As String

    ' Every cell button of the template has an action since phase 6, and the list
    ' must stay empty: a new button that nobody wired fails here instead of doing
    ' nothing when it is pressed.
    listText = Replace(modRoutes.PendingButtons(), vbCrLf, ",")
    Note "pending cell buttons: " & CStr(modApp.PendingButtonCount())

    AssertEqualsText listText, vbNullString, "no cell button is left without an action"
    AssertEqualsLong modApp.PendingButtonCount(), 0, _
                     "the number of cell buttons without an action"
End Sub

' ---- the texts -------------------------------------------------------

Private Sub Test_TextCoverage()
    Dim keys As Variant
    Dim item As Variant
    Dim key As String

    keys = Split(modText.ALL_KEYS, ",")
    For Each item In keys
        key = Trim$(CStr(item))
        If Len(key) > 0 Then
            AssertTrue (Len(modText.Text(key)) > 0), "the text of '" & key & "' is not empty"
        End If
    Next item

    AssertEqualsText modText.Text("no.such.key"), vbNullString, "an unknown key gives no text"
    AssertEqualsText modText.TextFormat(modText.KEY_ERROR_UNKNOWN_ACTION, "x.y"), _
                     "'x.y' is not a known action.", "TextFormat fills {1}"
    AssertEqualsText modText.MonthShortName(9), "Sep", "September is Sep"
    AssertEqualsText modText.MonthLongName(9), "September", "September is September"
    AssertEqualsText modText.WeekdayShortName(1), "Mon", "day 1 is Monday"
End Sub

' ---- the router ------------------------------------------------------

Private Sub Test_AppRouterBasics()
    Dim result As Object
    Dim busyResult As Object

    Set result = modApp.AppRouter(vbNullString)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "an empty action is refused"
    AssertEqualsText result(modRoutes.RESULT_MESSAGE), modText.Text(modText.KEY_ERROR_NO_ACTION), _
                     "the empty action has its own message"
    AssertEqualsLong result(modRoutes.RESULT_ERROR), 0, "no error number for a refusal"

    Set result = modApp.AppRouter("nonsense.action")
    AssertFalse (result(modRoutes.RESULT_OK) = True), "an unknown action is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "not a known action") > 0), _
               "the message names the problem"
    AssertEqualsText result(modRoutes.RESULT_ACTION), "nonsense.action", "the action is echoed"

    AssertTrue result.Exists(modRoutes.RESULT_DATA), "the result carries a data slot"

    modApp.SetBusy True
    Set busyResult = modApp.AppRouter(modRoutes.ACTION_APP_REFRESH)
    modApp.SetBusy False
    AssertEqualsText busyResult(modRoutes.RESULT_MESSAGE), modText.Text(modText.KEY_ERROR_BUSY), _
                     "a second action is refused while one runs"
    AssertFalse modApp.IsBusy(), "the busy flag is cleared afterwards"
End Sub

' ---- the named cells -------------------------------------------------

Private Sub Test_NamesAccess()
    Dim ws As Worksheet
    Dim cell As Range
    Dim missing As String

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)

    modNames.WriteName NAME_STATE_VIEW, "day"
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), "day", "WriteName / ReadName round trip"

    ' By name, not by address: the layout belongs to the template, so a redecorated
    ' Dashboard must not break the tests.
    Set cell = modNames.RangeOfName(NAME_BTN_VIEW_DAILY)
    AssertEqualsText modNames.NameOfCell(ws, cell), NAME_BTN_VIEW_DAILY, _
                     "the daily button is found by its name"
    AssertTrue modNames.IsNameAt(ws, cell, NAME_BTN_VIEW_DAILY), "IsNameAt finds the button"
    AssertFalse modNames.IsNameAt(ws, cell, NAME_BTN_VIEW_WEEKLY), _
                "IsNameAt does not mix the buttons up"

    missing = vbNullString
    On Error Resume Next
    missing = modNames.ReadName("b_no_such_button")
    On Error GoTo 0
    AssertEqualsText missing, vbNullString, "a missing name gives no value"

    AssertEqualsText modNames.StatusNameFor(SHEET_CLIENTS_ADMIN), NAME_STATUS_CLIENTS, _
                     "every page knows its status cell"
    AssertEqualsText modNames.ViewButtonName(VIEW_MONTH), NAME_BTN_VIEW_MONTHLY, _
                     "the month view has its button"
    AssertEqualsText modNames.ViewKeyOfButton(NAME_BTN_VIEW_MONTHLY), VIEW_MONTH, _
                     "and the button knows its view"
    AssertTrue modNames.IsDeclaredName(NAME_GRID_CLIENTS_LAST), "grid_clients_last is declared"
    AssertFalse modNames.IsDeclaredName("db_something"), "db_something is not declared"
End Sub

Private Sub Test_StatusCells()
    Dim cell As Range

    modUi.SetStatusOn SHEET_DASHBOARD, "unit test", False
    Set cell = modNames.RangeOfName(NAME_STATUS_DASHBOARD)
    AssertEqualsText cell.Text, "unit test", "the status cell shows the message"
    AssertEqualsDouble cell.Interior.Color, CDbl(COLOR_STATUS_ERROR), 0, "a refusal is red"

    modUi.ResetStatuses
    AssertEqualsText modNames.ReadName(NAME_STATUS_DASHBOARD), modText.Text(modText.KEY_STATUS_READY), _
                     "Ready. after a reset"
    AssertEqualsDouble cell.Interior.Color, CDbl(COLOR_STATUS_OK), 0, "Ready. is green"
End Sub

' ---- the small helpers -----------------------------------------------

Private Sub Test_TimeHelpers()
    AssertEqualsLong modApp.TimeToMinutes("09:30"), 570, "09:30 is 570 minutes"
    AssertEqualsLong modApp.TimeToMinutes("9:05"), 545, "9:05 is 545 minutes"
    AssertEqualsLong modApp.TimeToMinutes("00:00"), 0, "00:00 is 0"
    AssertEqualsLong modApp.TimeToMinutes("24:00"), -1, "24:00 does not exist"
    AssertEqualsLong modApp.TimeToMinutes("09:75"), -1, "75 minutes is not a time"
    AssertEqualsLong modApp.TimeToMinutes("nine"), -1, "text is not a time"

    AssertEqualsText modApp.MinutesToTime(570), "09:30", "570 minutes is 09:30"
    AssertEqualsText modApp.MinutesToTime(0), "00:00", "0 minutes is 00:00"
    AssertEqualsText modApp.MinutesToTime(-1), vbNullString, "a negative time has no text"
End Sub

Private Sub Test_MoneyHelper()
    Dim amount As Double

    AssertTrue modApp.TryMoneyValue("60", amount), "'60' is money"
    AssertEqualsDouble amount, 60, 0.001, "60 is 60"
    AssertTrue modApp.TryMoneyValue("60,50", amount), "'60,50' is money"
    AssertEqualsDouble amount, 60.5, 0.001, "the comma is the decimal separator"
    AssertTrue modApp.TryMoneyValue("", amount), "an empty amount is 0"
    AssertEqualsDouble amount, 0, 0.001, "no amount is 0"

    AssertFalse modApp.TryMoneyValue("12.345", amount), "three decimals are refused"
    AssertFalse modApp.TryMoneyValue("abc", amount), "text is not money"
End Sub

' ---- the events ------------------------------------------------------

Private Sub Test_SheetEvents()
    Dim dashboard As Worksheet
    Dim link As Object
    Dim cancel As Boolean

    Set dashboard = ThisWorkbook.Worksheets(SHEET_DASHBOARD)

    Set link = HyperlinkOf(dashboard, NAME_BTN_VIEW_DAILY)
    AssertTrue (Not link Is Nothing), "the daily button carries a hyperlink"

    ' A click on a button cell runs its action.
    modApp.RunCellAction SHEET_DASHBOARD, NAME_BTN_VIEW_WEEKLY
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_WEEK, "the weekly button set the view"

    ' A plain selection of the same cell must not run it a second time.
    modApp.OnSelection SHEET_DASHBOARD, modNames.RangeOfName(NAME_BTN_VIEW_WEEKLY)
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_WEEK, _
                     "a plain selection does not repeat the action"

    ' A double click is the second way to a button.
    cancel = False
    modApp.OnDoubleClick SHEET_DASHBOARD, modNames.RangeOfName(NAME_BTN_VIEW_MONTHLY), cancel
    AssertTrue cancel, "the double click cancels the cell edit"
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_MONTH, _
                     "the double click ran the action"

    ' Every button of the template answers with its own action now: this one hides
    ' the data sheets (Test_DataSheetVisibility checks the details and shows them
    ' again). The message of an action lands in the status cell of the sheet in
    ' front, so the test stands on Config - the way a user presses the button.
    modUi.Navigate SHEET_CONFIG
    modApp.RunCellAction SHEET_CONFIG, NAME_BTN_LOCK_DATA
    AssertEqualsText modNames.ReadName(modNames.StatusNameFor(SHEET_CONFIG)), _
                     modText.Text(modText.KEY_CONFIG_DATA_HIDDEN), "a wired button answers"
    modApp.RunCellAction SHEET_CONFIG, NAME_BTN_UNLOCK_DATA
    ' back where the test started: an action writes its message into the status of
    ' the sheet in front, and the tests of the refresh expect the Dashboard.
    modUi.Navigate SHEET_DASHBOARD

    ' A selection on a cell without a name changes nothing. The cell is taken from
    ' the body block (cal_first + one row), so it moves with the layout.
    modApp.OnSelection SHEET_DASHBOARD, _
        ThisWorkbook.Worksheets(SHEET_DASHBOARD).Cells(modDashboard.BodyFirstRow() + 1, _
                                                       DASH_COL_CLIENT)
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_MONTH, "an unnamed cell changes nothing"
End Sub

Private Sub Test_HyperlinkClick()
    Dim dashboard As Worksheet
    Dim link As Object

    Set dashboard = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    Set link = HyperlinkOf(dashboard, NAME_BTN_VIEW_DAILY)
    If link Is Nothing Then
        FailTest "the daily button has no hyperlink"
        Exit Sub
    End If

    ' The link points at the name of its own cell, so a click never moves the
    ' view (PROJECT_CONCEPT ??6.2.2), and the name is what the handler reads.
    AssertEqualsText link.SubAddress, NAME_BTN_VIEW_DAILY, "the hyperlink points at its own name"
    AssertEqualsText modApp.HyperlinkName(link), NAME_BTN_VIEW_DAILY, _
                     "the name of the link is read back"

    ' The dispatch by name, and then the handler the sheet events call.
    modApp.RunCellAction SHEET_DASHBOARD, NAME_BTN_VIEW_DAILY
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_DAY, _
                     "the name of the button runs the action"

    modApp.OnHyperlink SHEET_DASHBOARD, link
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_DAY, _
                     "the hyperlink of the button ran the action"

    ' "" means the link was read without a problem (modApp records the last one).
    AssertEqualsText modApp.LastHyperlinkProblem, vbNullString, _
                     "reading the hyperlink had no problem"
    modApp.LastHyperlinkProblem = vbNullString
End Sub

Private Sub Test_SheetEventsAttached()
    modApp.AttachSheetEvents
    AssertEqualsLong modApp.AttachedSheetCount(), 8, "one event instance per UI sheet"

    modApp.DetachSheetEvents
    AssertEqualsLong modApp.AttachedSheetCount(), 0, "DetachSheetEvents releases them"

    modApp.AttachSheetEvents
    AssertEqualsLong modApp.AttachedSheetCount(), 8, "attaching twice does not double the count"
End Sub

' ---- the data behind the header --------------------------------------

Private Sub Test_Refresh(ByRef result As Object)
    Dim info As Object
    Dim viewKey As String
    Dim fromText As String
    Dim toText As String

    AssertTrue (result(modRoutes.RESULT_OK) = True), "app.refresh succeeded"

    Set info = result(modRoutes.RESULT_DATA)
    AssertEqualsLong info("clients"), 3, "3 clients in the sample data"
    AssertEqualsLong info("services"), 3, "3 services in the sample data"
    AssertEqualsLong info("appointments"), 3, "3 appointments in the sample data"
    AssertEqualsLong modNames.RangeOfName(NAME_GRID_APPOINTMENTS_FIRST).Worksheet.Index, _
                     ThisWorkbook.Worksheets(SHEET_APPOINTMENTS).Index, _
                     "the all-appointments anchor belongs to Appointments"
    AssertEqualsText modNames.ReadName(NAME_APPOINTMENTS_TITLE), "Appointments (3)", _
                     "app.refresh renders the all-appointments list"

    AssertTrue modData.IsModelLoaded(), "the tables are loaded"
    AssertEqualsLong modData.ClientsRowCount(), 3, "the clients table has 3 rows"
    AssertEqualsText modData.ClientName(1), "Ivan Petrov", "client 1 is Ivan Petrov"
    AssertEqualsText modData.ClientName(999), vbNullString, "an unknown client has no name"
    AssertEqualsText modData.ServiceName(2), "Hair coloring", "service 2 is Hair coloring"
    AssertEqualsDouble modData.ServicePrice(2), 60, 0.001, "service 2 costs 60"
    AssertEqualsLong modData.ServiceDuration(3), 45, "service 3 lasts 45 minutes"
    AssertEqualsLong modData.NextId(modData.ClientsTable()), 4, "the next client id is 4"
    AssertEqualsText modData.TableProblems(), vbNullString, "the three tables are complete"

    AssertEqualsLong modData.AppointmentsOn("2026-09-21"), 1, "one appointment on 21 Sep 2026"
    AssertEqualsLong modData.AppointmentsInPeriod("2026-09-01", "2026-09-30"), 3, _
                     "three appointments in September 2026"
    AssertEqualsDouble modData.RevenueInPeriod("2026-09-01", "2026-09-30"), 115, 0.001, _
                       "the revenue of September is 115"

    viewKey = modDashboard.CurrentView()
    modDashboard.PeriodBounds viewKey, modDashboard.CurrentDate(), fromText, toText

    AssertEqualsText modNames.ReadName(NAME_VIEW), modDashboard.ViewText(viewKey), _
                     "the header shows the view"
    AssertEqualsText modNames.ReadName(NAME_PERIOD), _
                     modDashboard.PeriodTitle(viewKey, modDashboard.CurrentDate()), _
                     "the header shows the period"
    AssertEqualsText modNames.ReadName(NAME_COUNT), _
                     CStr(modData.AppointmentsInPeriod(fromText, toText)), "the header counts"
    AssertEqualsText modNames.ReadName(NAME_REVENUE), _
                     modDashboard.FormatMoney(modData.RevenueInPeriod(fromText, toText)), _
                     "the header adds the revenue up"
    AssertEqualsText modNames.ReadName(NAME_FREE_SLOTS), _
                     CStr(modDashboard.FreeSlotsFor(viewKey, modDashboard.CurrentDate())), _
                     "the header counts the free slots"
    AssertTrue (Len(modNames.ReadName(NAME_CAL_TITLE)) > 0), "the body has a title"
    AssertEqualsText modNames.ReadName(NAME_STATUS_DASHBOARD), _
                     CStr(result(modRoutes.RESULT_MESSAGE)), "the status cell carries the answer"
End Sub

Private Sub Test_DashboardViews(ByRef refreshResult As Object)
    Dim result As Object
    Dim payload As Object
    Dim anchorText As String
    Dim lastText As String

    ' The day view.
    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_VIEW) = VIEW_DAY
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_VIEW, payload)

    AssertTrue (result(modRoutes.RESULT_OK) = True), "the day view is accepted"
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_DAY, "s_view = day"
    AssertEqualsText modNames.ReadName(NAME_VIEW), VIEW_TEXT_DAY, "o_view = Daily"
    AssertEqualsDouble modNames.RangeOfName(NAME_BTN_VIEW_DAILY).Interior.Color, _
                       CDbl(COLOR_BUTTON_ACTIVE), 0, "the daily button is the active one"
    AssertEqualsDouble modNames.RangeOfName(NAME_BTN_VIEW_MONTHLY).Interior.Color, _
                       CDbl(COLOR_BUTTON), 0, "the monthly button is not active"

    ' A period of one day: first and last date are the same.
    modDashboard.PeriodBounds VIEW_DAY, modDashboard.CurrentDate(), anchorText, lastText
    AssertEqualsText anchorText, lastText, "the day view shows one date"

    ' One day forward, then back to today.
    anchorText = modNames.ReadName(NAME_STATE_DATE)
    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_STEP) = "1"
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_MOVE, payload)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "moving the period is accepted"
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), _
                     modDashboard.FormatDateText( _
                         DateAdd("d", 1, modDashboard.DateOfText(anchorText))), "one day forward"

    payload(modRoutes.KEY_STEP) = "0"
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_MOVE, payload)
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), modDashboard.FormatDateText(Date), _
                     "Today goes back to today"

    ' The week view, one week back: the anchor moves seven days.
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_VIEW, _
                                  modRoutes.PayloadForCell(NAME_BTN_VIEW_WEEKLY))
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the week view is accepted"
    payload(modRoutes.KEY_STEP) = "-1"
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_MOVE, payload)
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), _
                     modDashboard.FormatDateText(DateAdd("d", -7, Date)), "one week back"

    ' A view that does not exist is refused.
    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_VIEW) = "hour"
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_VIEW, payload)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "an unknown view is refused"
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_WEEK, "the view did not change"
End Sub

' ---- the log ---------------------------------------------------------

Private Sub Test_Log()
    Dim lines As Collection
    Dim text As String

    CheckLogAnchor

    modLog.Clear
    AssertEqualsLong modLog.RowCount(), 0, "the log is empty after Clear"

    modApp.AppRouter modRoutes.ACTION_APP_REFRESH
    AssertEqualsLong modLog.RowCount(), 1, "one row after one action"

    modApp.AppRouter "unknown.action"
    AssertEqualsLong modLog.RowCount(), 2, "a refused action is logged as well"

    Set lines = modLog.Tail(2)
    AssertEqualsLong lines.Count, 2, "Tail gives the two newest rows"
    AssertTrue (InStr(1, CStr(lines.Item(1)), "unknown.action") > 0), "the newest row is the refusal"

    modLog.TrimTo 1
    AssertEqualsLong modLog.RowCount(), 1, "TrimTo keeps one row"
    text = CStr(modLog.Tail(1).Item(1))
    AssertTrue (InStr(1, text, "unknown.action") > 0), "the kept row is the newest"

    ' The anchor name has to survive the trimming: deleting rows would break it.
    AssertEqualsText modNames.SheetOfName(NAME_LOG_FIRST), SHEET_LOG, _
                     "log_first still resolves after TrimTo"

    modLog.Clear
    AssertEqualsLong modLog.RowCount(), 0, "Clear empties the log"
End Sub

' ---- the workbook itself ---------------------------------------------

Private Sub Test_NoFormControls()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim item As Variant
    Dim shapes As Long

    sheetNames = Split(modConfig.UiSheetNames(), ",")
    For Each item In sheetNames
        Set ws = ThisWorkbook.Worksheets(Trim$(CStr(item)))
        shapes = -1
        On Error Resume Next
        shapes = ws.Shapes.Count
        On Error GoTo 0
        AssertEqualsLong shapes, 0, "no form controls or shapes on " & ws.Name
    Next item
End Sub

Private Sub Test_Diagnostics()
    Dim targetPath As String
    Dim fileNumber As Integer
    Dim text As String
    Dim line As String

    ' The file stays in build\ on purpose: it is the artefact of the test run
    ' (contract, routes, counters, log) and it is what makes a failure readable.
    targetPath = ThisWorkbook.Path & Application.PathSeparator & "test-diagnostics.txt"
    modApp.WriteDiagnostics targetPath

    AssertTrue (Len(Dir$(targetPath)) > 0), "the diagnostics file exists"

    fileNumber = FreeFile
    Open targetPath For Input As #fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        text = text & line & vbCrLf
    Loop
    Close #fileNumber

    If InStr(1, text, "contract  : ok") = 0 Then
        Note "diagnostics / contract: " & Mid$(text, InStr(1, text, "contract  :"), 400)
    End If

    AssertTrue (InStr(1, text, "contract  : ok") > 0), "the report says the contract is fine"
    AssertTrue (InStr(1, text, "routes    : ok") > 0), "the report says the routes are fine"
    AssertTrue (InStr(1, text, "app.refresh -> modApp.Refresh") > 0), "the report lists the routes"
    AssertTrue (InStr(1, text, "state     : finished") > 0), "the report is complete"
End Sub

' =====================================================================
'  phase 2: clients
' =====================================================================

Private Sub Test_ClientsAdd()
    Dim before As Long
    Dim idValue As Long
    Dim formId As Long
    Dim result As Object

    before = modData.ClientsRowCount()

    ' The clients page in front, so the automatic render after the action
    ' (??5.5) draws this sheet.
    modUi.Navigate SHEET_CLIENTS_ADMIN
    idValue = AddClientForTest("Test Client", "0888 123 456", "test@example.com", "from data")

    AssertTrue (idValue > 0), "clients.add made a new id"
    AssertEqualsLong modData.ClientsRowCount(), before + 1, "one more client in the table"
    AssertEqualsText modData.ClientName(idValue), "Test Client", "the client is in the table"
    AssertEqualsLong modData.ClientIdByName("Test Client"), idValue, "the id is found by its name"
    AssertTrue (InStr(1, modNames.ReadName(NAME_STATUS_CLIENTS), "was added") > 0), _
               "the status carries the message of the action"
    AssertTrue (InStr(1, modNames.ReadName(NAME_CLIENTS_TITLE), "Clients (") > 0), _
               "the router rendered the list after the action"

    ' The form path: the payload is empty, so the form is read (and cleared).
    modNames.WriteName NAME_CLIENT_NAME, "Form Client"
    modNames.WriteName NAME_CLIENT_PHONE, "0899 000 111"
    modNames.WriteName NAME_CLIENT_EMAIL, vbNullString
    modNames.WriteName NAME_CLIENT_COMMENT, vbNullString

    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_ADD)
    AssertTrue (result(modRoutes.RESULT_OK) = True), _
               "the form path added a client: " & CStr(result(modRoutes.RESULT_MESSAGE))
    formId = modData.ClientIdByName("Form Client")
    AssertTrue (formId > 0), "the client from the form is in the table"
    AssertEqualsText modNames.ReadName(NAME_CLIENT_NAME), vbNullString, "the form was cleared"

    ' The row count is a contract for the tests: clean up.
    modUi.SetInteractiveUi False
    AssertTrue DeleteClientForTest(idValue), "the first test client was removed"
    AssertTrue DeleteClientForTest(formId), "the second test client was removed"
    AssertEqualsLong modData.ClientsRowCount(), before, "the row count is back"
End Sub

Private Sub Test_ClientsValidation()
    Dim before As Long
    Dim result As Object

    before = modData.ClientsRowCount()

    ' No name.
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_ADD, ClientValues(vbNullString, "0888", "", ""))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a client without a name is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "required") > 0), _
               "the message says what is missing"

    ' Phone too short.
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_ADD, ClientValues("Short Phone", "12", "", ""))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a two digit phone is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "digits") > 0), _
               "the message names the rule"

    ' E-mail without @.
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_ADD, ClientValues("Bad Mail", "", "nope", ""))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "an e-mail without @ is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "@") > 0), _
               "the message quotes the address"

    AssertEqualsLong modData.ClientsRowCount(), before, "nothing was written"
End Sub

Private Sub Test_ClientsDelete()
    Dim idValue As Long
    Dim before As Long
    Dim result As Object
    Dim data As Object

    before = modData.ClientsRowCount()
    idValue = AddClientForTest("Delete Me", "0888 555 555", "", "")
    AssertTrue (idValue > 0), "the client to delete exists"

    ' No confirmation in a test run (modUi.Interactive = False): one call deletes.
    modUi.SetInteractiveUi False
    AssertTrue DeleteClientForTest(idValue), "the client was deleted without a dialog"
    AssertEqualsLong modData.ClientsRowCount(), before, "the row count is back"
    AssertFalse modData.RowExists(SHEET_DB_CLIENTS, idValue), "the row is gone"

    ' The two step confirmation of ??6.6, with the dialogs on.
    idValue = AddClientForTest("Confirm Me", "0888 666 666", "", "")
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_DELETE, data)

    AssertTrue (result(modRoutes.RESULT_OK) = True), _
               "the confirmation is answered at once without a dialog"
    AssertFalse modData.RowExists(SHEET_DB_CLIENTS, idValue), "the client is gone"
    ' A second row is deleted too, so nothing is remembered between two deletes:
    ' the dialog decides every time.
    idValue = AddClientForTest("Other Row", "0888 777 777", "", "")
    data(modRoutes.KEY_ID) = CStr(idValue)
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_DELETE, data)
    AssertFalse modData.RowExists(SHEET_DB_CLIENTS, idValue), "the other row is gone too"
    AssertEqualsLong modData.ClientsRowCount(), before, "the row count is back"
End Sub

Private Sub Test_ClientsDeleteGuard()
    Dim data As Object
    Dim result As Object
    Dim before As Long

    before = modData.ClientsRowCount()

    ' The sample client of the template is used by a sample appointment.
    AssertTrue (modClients.References(1) > 0), "the sample client is used by an appointment"

    modUi.SetInteractiveUi False
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = "1"
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_DELETE, data)

    AssertFalse (result(modRoutes.RESULT_OK) = True), "a client in use is not deleted"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "is used by") > 0), _
               "the message says how many appointments use the client"
    AssertTrue modData.RowExists(SHEET_DB_CLIENTS, 1), "the client is still in the table"
    AssertEqualsLong modData.ClientsRowCount(), before, "the row count did not change"

    ' An id that does not exist.
    data(modRoutes.KEY_ID) = "9999"
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_DELETE, data)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "an unknown id is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "does not exist") > 0), _
               "the message says that the client does not exist"

    ' An id that is not a number at all.
    data(modRoutes.KEY_ID) = "abc"
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_DELETE, data)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a broken id is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "not valid") > 0), _
               "the message names the id"
End Sub

Private Sub Test_ClientsList()
    Dim idValue As Long
    Dim before As Long
    Dim rowIndex As Long
    Dim firstCell As Range
    Dim actionCell As Range
    Dim link As Object
    Dim result As Object
    Dim info As Object
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS_ADMIN)
    modUi.Navigate SHEET_CLIENTS_ADMIN

    before = modData.ClientsRowCount()
    idValue = AddClientForTest("Row Click", "0888 444 444", "", "")
    AssertTrue (idValue > 0), "the client of the click test exists"

    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_LIST)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "clients.list succeeded"
    Set info = result(modRoutes.RESULT_DATA)
    AssertEqualsLong info("rows"), before + 1, "the list reports the rows of the table"

    Set firstCell = modNames.RangeOfName(NAME_GRID_CLIENTS_FIRST)
    rowIndex = firstCell.Row + before

    AssertEqualsText ws.Cells(rowIndex, CLIENTS_COL_ID).Text, CStr(idValue), _
                     "the rendered row shows the id"
    AssertEqualsText ws.Cells(rowIndex, CLIENTS_COL_NAME).Text, "Row Click", _
                     "the rendered row shows the name"
    AssertEqualsText ws.Cells(rowIndex, CLIENTS_COL_ACTION).Text, LIST_EDIT_TEXT, _
                     "the rendered row has an [Edit] cell"
    AssertEqualsText ws.Cells(rowIndex, CLIENTS_COL_DELETE).Text, LIST_ACTION_TEXT, _
                     "and a [Delete] cell of its own"

    Set actionCell = ws.Cells(rowIndex, CLIENTS_COL_DELETE)
    Set link = Nothing
    On Error Resume Next
    Set link = actionCell.Hyperlinks(1)
    On Error GoTo 0
    AssertTrue (Not link Is Nothing), "the [Delete] cell carries a hyperlink"
    If Not link Is Nothing Then
        AssertEqualsText link.SubAddress, NAME_GRID_CLIENTS_FIRST, _
                         "the link of a rendered row points at the grid anchor"
        AssertTrue (Not link.Range Is Nothing), "the link knows the cell it belongs to"
        If Not link.Range Is Nothing Then
            AssertEqualsLong link.Range.Row, rowIndex, "the link belongs to the clicked row"
        End If
    End If

    ' The click is recognized by position (??6.10) and deletes the row.
    modUi.SetInteractiveUi False
    modApp.OnDynamicClick SHEET_CLIENTS_ADMIN, actionCell
    AssertFalse modData.RowExists(SHEET_DB_CLIENTS, idValue), "the click deleted the client"
    AssertEqualsLong modData.ClientsRowCount(), before, "the row count is back"

    ' A click on a data column means nothing: only the two link columns act.
    Set actionCell = ws.Cells(rowIndex, CLIENTS_COL_NAME)
    modApp.OnDynamicClick SHEET_CLIENTS_ADMIN, actionCell
    AssertEqualsLong modData.ClientsRowCount(), before, "a click on a data column does nothing"
End Sub

' =====================================================================
'  phase 2: services
' =====================================================================

Private Sub Test_ServicesAdd()
    Dim before As Long
    Dim firstId As Long
    Dim secondId As Long

    before = modData.ServicesRowCount()
    modUi.Navigate SHEET_SERVICES_ADMIN

    firstId = AddServiceForTest("Test Service", "25,50", "20")
    AssertTrue (firstId > 0), "services.add made a new id"
    AssertEqualsLong modData.ServicesRowCount(), before + 1, "one more service in the table"
    AssertEqualsText modData.ServiceName(firstId), "Test Service", "the service is in the table"
    AssertEqualsDouble modData.ServicePrice(firstId), 25.5, 0.001, _
                       "the comma was read as the decimal separator"
    AssertEqualsLong modData.ServiceDuration(firstId), 20, "the duration was stored"
    AssertTrue (InStr(1, modNames.ReadName(NAME_STATUS_SERVICES), "was added") > 0), _
               "the status carries the message of the action"
    AssertTrue (InStr(1, modNames.ReadName(NAME_SERVICES_TITLE), "Services (") > 0), _
               "the router rendered the list after the action"

    ' The defaults of ??16.1: an empty price is 0, an empty duration 10 minutes.
    secondId = AddServiceForTest("Default Service", "", "")
    AssertTrue (secondId > 0), "a service with the defaults was added"
    AssertEqualsDouble modData.ServicePrice(secondId), 0, 0.001, "an empty price becomes 0"
    AssertEqualsLong modData.ServiceDuration(secondId), DEFAULT_DURATION_MINUTES, _
                     "an empty duration becomes the default"

    modUi.SetInteractiveUi False
    AssertTrue DeleteServiceForTest(firstId), "the first test service was removed"
    AssertTrue DeleteServiceForTest(secondId), "the second test service was removed"
    AssertEqualsLong modData.ServicesRowCount(), before, "the row count is back"
End Sub

Private Sub Test_ServicesValidation()
    Dim before As Long
    Dim result As Object

    before = modData.ServicesRowCount()

    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_ADD, ServiceValues(vbNullString, "10", "10"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a service without a name is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "required") > 0), _
               "the message says what is missing"

    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_ADD, ServiceValues("Bad Price", "abc", "10"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a price that is not an amount is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "not an amount") > 0), _
               "the message quotes the price"

    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_ADD, ServiceValues("Bad Duration", "10", "abc"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a duration that is not a number is refused"

    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_ADD, ServiceValues("Zero Duration", "10", "0"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a zero duration is refused"

    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_ADD, ServiceValues("Huge Duration", "10", "2000"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a duration longer than a day is refused"

    AssertEqualsLong modData.ServicesRowCount(), before, "nothing was written"
End Sub

Private Sub Test_ServicesDelete()
    Dim idValue As Long
    Dim before As Long
    Dim result As Object
    Dim data As Object

    before = modData.ServicesRowCount()
    idValue = AddServiceForTest("Delete Service", "5", "15")
    AssertTrue (idValue > 0), "the service to delete exists"

    ' Without dialogs (tests) one call deletes.
    modUi.SetInteractiveUi False
    AssertTrue DeleteServiceForTest(idValue), "the service was deleted without a dialog"
    AssertEqualsLong modData.ServicesRowCount(), before, "the row count is back"
    AssertFalse modData.RowExists(SHEET_DB_SERVICES, idValue), "the row is gone"

    ' The two step confirmation of ??6.6.
    idValue = AddServiceForTest("Confirm Service", "5", "15")
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)

    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_DELETE, data)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the delete needs no second click"
    AssertFalse modData.RowExists(SHEET_DB_SERVICES, idValue), "the service is gone"

    modUi.SetInteractiveUi False
    modApp.ConfirmReset
    AssertEqualsLong modData.ServicesRowCount(), before, "the row count is back"
End Sub

Private Sub Test_ServicesDeleteGuard()
    Dim data As Object
    Dim result As Object
    Dim before As Long

    before = modData.ServicesRowCount()

    AssertTrue (modServices.References(1) > 0), "the sample service is used by an appointment"

    modUi.SetInteractiveUi False
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = "1"
    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_DELETE, data)

    AssertFalse (result(modRoutes.RESULT_OK) = True), "a service in use is not deleted"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "is used by") > 0), _
               "the message says how many appointments use the service"
    AssertTrue modData.RowExists(SHEET_DB_SERVICES, 1), "the service is still in the table"
    AssertEqualsLong modData.ServicesRowCount(), before, "the row count did not change"

    data(modRoutes.KEY_ID) = "9999"
    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_DELETE, data)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "an unknown service id is refused"
End Sub

Private Sub Test_ServicesList()
    Dim idValue As Long
    Dim before As Long
    Dim rowIndex As Long
    Dim actionCell As Range
    Dim link As Object
    Dim result As Object
    Dim info As Object
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_SERVICES_ADMIN)
    modUi.Navigate SHEET_SERVICES_ADMIN

    before = modData.ServicesRowCount()
    idValue = AddServiceForTest("Row Service", "12,50", "30")
    AssertTrue (idValue > 0), "the service of the click test exists"

    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_LIST)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "services.list succeeded"
    Set info = result(modRoutes.RESULT_DATA)
    AssertEqualsLong info("rows"), before + 1, "the list reports the rows of the table"

    rowIndex = modNames.RangeOfName(NAME_GRID_SERVICES_FIRST).Row + before
    AssertEqualsText ws.Cells(rowIndex, SERVICES_COL_ID).Text, CStr(idValue), _
                     "the rendered row shows the id"
    AssertEqualsText ws.Cells(rowIndex, SERVICES_COL_NAME).Text, "Row Service", _
                     "the rendered row shows the service"
    AssertEqualsDouble ws.Cells(rowIndex, SERVICES_COL_PRICE).Value, 12.5, 0.001, _
                       "the rendered row shows the price as a number"
    AssertEqualsLong ws.Cells(rowIndex, SERVICES_COL_DURATION).Value, 30, _
                     "the rendered row shows the duration"

    AssertEqualsText ws.Cells(rowIndex, SERVICES_COL_ACTION).Text, LIST_EDIT_TEXT, _
                     "the rendered row has an [Edit] cell"

    Set actionCell = ws.Cells(rowIndex, SERVICES_COL_DELETE)
    Set link = Nothing
    On Error Resume Next
    Set link = actionCell.Hyperlinks(1)
    On Error GoTo 0
    AssertTrue (Not link Is Nothing), "the [Delete] cell of a service carries a hyperlink"
    If Not link Is Nothing Then
        AssertEqualsText link.SubAddress, NAME_GRID_SERVICES_FIRST, _
                         "the link points at the services grid anchor"
    End If

    modUi.SetInteractiveUi False
    modApp.OnDynamicClick SHEET_SERVICES_ADMIN, actionCell
    AssertFalse modData.RowExists(SHEET_DB_SERVICES, idValue), "the click deleted the service"
    AssertEqualsLong modData.ServicesRowCount(), before, "the row count is back"
End Sub

' =====================================================================
'  phase 2: the lists, the forms and the copies
' =====================================================================

Private Sub Test_ListsAndClear()
    Dim result As Object
    Dim data As Object
    Dim listRange As Range

    ' The dropdown lists are filled from the tables (??6.5.4).
    modData.RefreshLists
    Set listRange = modNames.RangeOfName(NAME_LIST_CLIENTS)
    AssertEqualsText listRange.Cells(1, 1).Text, modData.ClientName(1), _
                     "the first entry of list_clients is the first client"
    AssertEqualsText listRange.Cells(2, 1).Text, modData.ClientName(2), _
                     "the second entry follows the table"

    Set listRange = modNames.RangeOfName(NAME_LIST_SERVICES)
    AssertEqualsText listRange.Cells(1, 1).Text, modData.ServiceName(1), _
                     "the first entry of list_services is the first service"
    AssertTrue (Len(listRange.Cells(4, 1).Text) = 0), "the list is not longer than the table"

    ' form.clear: the ClientsAdmin form.
    modNames.WriteName NAME_CLIENT_NAME, "To Be Cleared"
    modNames.WriteName NAME_CLIENT_COMMENT, "and this"
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_SHEET) = SHEET_CLIENTS_ADMIN
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_CLEAR, data)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "form.clear worked for the clients"
    AssertEqualsText modNames.ReadName(NAME_CLIENT_NAME), vbNullString, "the name field is empty"
    AssertEqualsText modNames.ReadName(NAME_CLIENT_COMMENT), vbNullString, "the comment is empty"

    ' form.clear: the ServicesAdmin form.
    modNames.WriteName NAME_SERVICE_NAME, "To Be Cleared"
    data(modRoutes.KEY_SHEET) = SHEET_SERVICES_ADMIN
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_CLEAR, data)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "form.clear worked for the services"
    AssertEqualsText modNames.ReadName(NAME_SERVICE_NAME), vbNullString, "the service field is empty"

    ' form.clear on the Booking form (phase 3).
    modNames.WriteName NAME_BOOKING_CLIENT, "To Be Cleared"
    data(modRoutes.KEY_SHEET) = SHEET_BOOKING
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_CLEAR, data)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "form.clear worked for the booking"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT), vbNullString, "the booking is clean"

    ' form.clear on a sheet that has no form at all.
    data(modRoutes.KEY_SHEET) = SHEET_REPORTS
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_CLEAR, data)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a sheet without a form is refused"
End Sub

Private Sub Test_Backup()
    Dim idValue As Long
    Dim result As Object
    Dim data As Object
    Dim copiesBefore As Long
    Dim goodFolder As String

    goodFolder = modBackup.FolderPath()

    ' Switched off: nothing is written, and the caller may continue.
    modBackup.SetEnabled False
    AssertTrue modBackup.SaveCopy("switched off"), "a switched off copy answers OK"
    AssertEqualsText modBackup.LastPath(), vbNullString, "nothing was written"

    ' Switched on: a real copy lands next to the workbook.
    modBackup.SetEnabled True
    copiesBefore = modBackup.CountCopies()
    AssertTrue modBackup.SaveCopy("unit-test"), "the copy was written"
    AssertTrue (Len(modBackup.LastPath()) > 0), "the copy has a path"
    AssertTrue (Len(Dir$(modBackup.LastPath())) > 0), "the copy file exists"
    ' The folder keeps at most BACKUP_KEEP copies, and the newest one survives the
    ' trim (it is the copy that was just written).
    AssertTrue (modBackup.CountCopies() > 0), "the folder still holds copies"

    ' A folder that cannot be created: the copy fails and a delete is refused.
    modBackup.SetFolder "Z:\service-manager-does-not-exist\"
    modUi.SetInteractiveUi False
    idValue = AddClientForTest("Cannot Delete", "0888 333 333", "", "")
    AssertTrue (idValue > 0), "the client for the backup test exists"

    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_DELETE, data)

    AssertFalse (result(modRoutes.RESULT_OK) = True), "without a copy nothing is deleted"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "backup") > 0), _
               "the message explains why"
    AssertTrue modData.RowExists(SHEET_DB_CLIENTS, idValue), "the client is still there"

    ' Restore the settings and clean up.
    modBackup.SetFolder goodFolder
    modBackup.SetEnabled False
    AssertTrue DeleteClientForTest(idValue), "the client of the backup test was removed"
End Sub

Private Sub Test_HyperlinkWithoutSubAddress()
    Dim ws As Worksheet
    Dim cell As Range
    Dim link As Object

    ' The open question of ??6.10: does a hyperlink with neither Address nor
    ' SubAddress survive? The answer is written into the report, so the next
    ' phase can use it for the dynamic cells.
    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS_ADMIN)
    Set cell = ws.Cells(1, CLIENTS_COL_ID + 20)      ' a spare cell far to the right
    cell.Hyperlinks.Delete
    cell.ClearContents

    On Error Resume Next
    cell.Hyperlinks.Add cell, vbNullString, vbNullString, "no address at all", "empty"
    On Error GoTo 0
    Err.Clear

    Set link = Nothing
    On Error Resume Next
    Set link = cell.Hyperlinks(1)
    On Error GoTo 0

    If link Is Nothing Then
        Note "a hyperlink without Address and SubAddress is dropped by Excel"
    Else
        Note "a hyperlink without Address and SubAddress survives (SubAddress='" & _
             CStr(link.SubAddress) & "')"
    End If

    cell.Hyperlinks.Delete
    cell.ClearContents
End Sub

' =====================================================================
'  phase 3: appointments, booking page, pickers
' =====================================================================

Private Sub Test_AppointmentTimes()
    ' EndTime
    AssertEqualsText modAppointments.EndTime("09:30", 60), "10:30", "09:30 + 60 is 10:30"
    AssertEqualsText modAppointments.EndTime("23:30", 60), "24:30", "the end may run past midnight (and is refused later)"
    AssertEqualsText modAppointments.EndTime("", 60), vbNullString, "no start, no end"
    AssertEqualsText modAppointments.EndTime("09:30", 0), vbNullString, "no duration, no end"

    ' Overlaps: touching intervals do not clash (??8.2.7)
    AssertTrue modAppointments.Overlaps(540, 60, 570, 60), "09:00+60 and 09:30+60 clash"
    AssertFalse modAppointments.Overlaps(540, 60, 600, 30), "09:00+60 and 10:00+30 do not clash"
    AssertFalse modAppointments.Overlaps(600, 30, 540, 60), "and the other way round"
    AssertTrue modAppointments.Overlaps(540, 60, 599, 30), "one minute before the end clashes"

    ' IsOnStep
    AssertTrue modAppointments.IsOnStep(555, 5), "555 minutes sit on a 5 minute step"
    AssertFalse modAppointments.IsOnStep(547, 5), "547 minutes do not"
    AssertTrue modAppointments.IsOnStep(547, 0), "a step of 0 means no rule"

    ' DurationText
    AssertEqualsText modAppointments.DurationText(10), "10 min", "10 minutes"
    AssertEqualsText modAppointments.DurationText(90), "1 h 30 min", "90 minutes"

    ' The working day of Config
    AssertEqualsLong modAppointments.DayStartMinutes(), 540, "the day starts at 09:00"
    AssertEqualsLong modAppointments.DayEndMinutes(), 1140, "the day ends at 19:00"
    AssertEqualsLong modAppointments.StepMinutes(), 5, "the grid step is 5 minutes"
    AssertTrue modAppointments.IsWorkingDay("2026-09-21"), "a Monday is a working day"
    AssertFalse modAppointments.IsWorkingDay("2026-09-19"), "a Saturday is not"
    AssertFalse modAppointments.IsWorkingDay("not a date"), "text is not a day"
End Sub

Private Sub Test_PickersPure()
    ' the date helpers
    AssertEqualsText modPickers.NormaliseDateText("2026-09-21"), "2026-09-21", "ISO text stays"
    AssertTrue modPickers.IsIsoDate("2026-09-21"), "ISO text is recognised"
    AssertFalse modPickers.IsIsoDate("21.09.2026"), "a locale date is not ISO"
    AssertEqualsText modPickers.MonthTitle(2026, 9), "September 2026", "the month title"
    AssertEqualsLong modPickers.MonthStartOffset(2026, 9), 1, "1 September 2026 is a Tuesday"
    AssertEqualsText modPickers.GridDateText(2026, 9, 1), vbNullString, "the first cell is padding"
    AssertEqualsText modPickers.GridDateText(2026, 9, 2), "2026-09-01", "the second cell is the first"
    AssertEqualsLong modPickers.GridIndexOf(2026, 9, "2026-09-21"), 22, "the grid position of 21 September"

    ' the screen form of a date: European, day first (??9.2)
    AssertEqualsText modPickers.DateFormatText("2026-09-17"), "17.09.2026", "the European date"
    AssertEqualsText modPickers.DateFormatText("2026-12-01"), "01.12.2026", "day and month are padded"
    AssertEqualsText modPickers.DateFormatText("17.09.2026"), vbNullString, "only ISO text is formatted"

    ' what the user may type: ISO or European, the data is always ISO
    AssertEqualsText modPickers.DateTextToIso("2026-09-25"), "2026-09-25", "ISO text is kept"
    AssertEqualsText modPickers.DateTextToIso("25.09.2026"), "2026-09-25", "25.09.2026 is read"
    AssertEqualsText modPickers.DateTextToIso("25/09/2026"), "2026-09-25", "25/09/2026 is read"
    AssertEqualsText modPickers.DateTextToIso("25-9-26"), "2026-09-25", "25-9-26 is read"
    AssertEqualsText modPickers.DateTextToIso("31.02.2026"), vbNullString, "February has 28 days"
    AssertEqualsText modPickers.DateTextToIso("17.13.2026"), vbNullString, "there is no month 13"
    AssertEqualsText modPickers.DateTextToIso("hello"), vbNullString, "text is not a date"
    AssertFalse modPickers.IsIsoDate("2026-02-30"), "30 February is not a date"

    ' the window of the picker: the year before this one up to ten years ahead
    AssertEqualsText modPickers.PickerFirstDate(), _
                     modDashboard.FormatDateText(DateSerial(Year(Date) - PICKER_YEARS_BACK, 1, 1)), _
                     "the first date of the window"
    AssertEqualsText modPickers.PickerLastDate(), _
                     modDashboard.FormatDateText(DateSerial(Year(Date) + PICKER_YEARS_FORWARD, 12, 31)), _
                     "the last date of the window"
    AssertTrue modPickers.IsInPickerRange(modDashboard.FormatDateText(Date)), _
               "today is inside the window"
    AssertTrue modPickers.IsInPickerRange(modPickers.DateFormatText(modDashboard.FormatDateText(Date))), _
               "the European form of today is inside too"
    AssertFalse modPickers.IsInPickerRange("1999-12-17"), "17 December 1999 is outside"
    AssertFalse modPickers.IsInPickerRange("01.01.1900"), "1900 is outside"
    AssertFalse modPickers.IsInPickerRange("17.09.2099"), "99 years ahead is outside"
    AssertFalse modPickers.IsInPickerRange("not a date"), "text is not a date of the window"

    ' the time helpers (??9.3, ??9.4)
    AssertEqualsText modPickers.ParseTimeText("8:30"), "08:30", "8:30 is read"
    AssertEqualsText modPickers.ParseTimeText("830"), "08:30", "830 is read"
    AssertEqualsText modPickers.ParseTimeText("8,30"), "08:30", "8,30 is read"
    AssertEqualsText modPickers.ParseTimeText("8"), "08:00", "a bare hour is read"
    AssertEqualsText modPickers.ParseTimeText("9:75"), vbNullString, "75 minutes are refused"
    AssertEqualsText modPickers.ParseTimeText("24:00"), vbNullString, "24:00 does not exist"
    AssertEqualsText modPickers.HourText("7"), "07", "the hour 7 becomes 07"
    AssertEqualsText modPickers.HourText("24"), vbNullString, "24 is not an hour"
    AssertEqualsText modPickers.MinuteText("5"), "05", "the minute 5 becomes 05"
    AssertEqualsText modPickers.MinuteText("7"), vbNullString, "7 is not on the 5 minute step"
    AssertEqualsText modPickers.MinuteText("55"), "55", "55 is the last minute of the grid"

    ' the busy check of the picker
    AssertFalse modPickers.IsSlotBusy(540, 60, "10:00+30"), "09:00 + 60 stops exactly at 10:00"
    AssertTrue modPickers.IsSlotBusy(600, 60, "10:00+30"), "10:00 + 60 runs into the 10:00 booking"
    AssertFalse modPickers.IsSlotBusy(480, 60, "09:00+30"), "08:00 + 60 stops at 09:00"
    AssertFalse modPickers.IsSlotBusy(540, 0, "09:00+30"), "without a duration nothing is busy"
End Sub

' The two picker forms are built and driven headless: no dialog is shown, the
' public interface of the form is the same one the clicks use (??9.1).
Private Sub Test_PickerForms()
    Dim datePicker As frmDatePicker
    Dim timePicker As frmTimePicker
    Dim index As Long

    Set datePicker = New frmDatePicker
    datePicker.InitialiseWithDate "2026-09-21"
    AssertEqualsLong datePicker.ShownYear(), 2026, "the shown year"
    AssertEqualsLong datePicker.ShownMonth(), 9, "the shown month"
    AssertEqualsText datePicker.ShownMonthTitle(), "September 2026", "the title of the calendar"
    AssertEqualsLong datePicker.DayCount(), 42, "six weeks of days"
    AssertEqualsLong datePicker.SelectedGridIndex(), 22, "21 September is highlighted"
    AssertEqualsText datePicker.DayCaptionAt(22), "21", "the caption of that cell"
    AssertFalse datePicker.DayEnabledAt(1), "a padding cell is disabled"
    AssertTrue datePicker.DayEnabledAt(22), "a day of the month is enabled"

    ' Every day button carries the value of the shown month: a click sends that
    ' value (clsPickerButton), never the caption. The bug of the first version was
    ' exactly here - the grid was built while the shown year was still 0, so every
    ' click answered a 1899-12-xx date.
    For index = 1 To datePicker.DayCount()
        AssertEqualsText datePicker.GridValueAt(index), datePicker.DayValueAt(index), _
                         "day button " & CStr(index) & " carries its own date"
    Next index

    ' the click of a day (the path a real click takes) answers that day
    datePicker.ClickDayAt(22)
    AssertEqualsText datePicker.PickedDateText, "2026-09-21", "the clicked day is the result"
    AssertEqualsText datePicker.SelectedDateDisplay(), "21.09.2026", _
                     "the footer shows the selected date the European way"

    ' a value that cannot be a date is refused by the form
    datePicker.InitialiseWithDate "2026-09-21"
    datePicker.ButtonClicked ROLE_DAY, "21"
    AssertEqualsText datePicker.PickedDateText, vbNullString, "a caption is not a date"

    datePicker.ButtonClicked ROLE_NEXT_MONTH, vbNullString
    AssertEqualsLong datePicker.ShownMonth(), 10, "the next month is shown"
    AssertEqualsText datePicker.ShownMonthTitle(), "October 2026", "the title followed"

    ' after the shift every button follows the new month
    For index = 1 To datePicker.DayCount()
        AssertEqualsText datePicker.GridValueAt(index), datePicker.DayValueAt(index), _
                         "day button " & CStr(index) & " follows the new month"
    Next index
    datePicker.ClickDayAt(modPickers.GridIndexOf(2026, 10, "2026-10-05"))
    AssertEqualsText datePicker.PickedDateText, "2026-10-05", "the click answers the shown month"

    datePicker.ButtonClicked ROLE_PREV_YEAR, vbNullString
    AssertEqualsLong datePicker.ShownYear(), 2025, "a whole year back"

    ' the window of the picker stops the navigation (PICKER_YEARS_BACK)
    datePicker.ButtonClicked ROLE_PREV_YEAR, vbNullString
    datePicker.ButtonClicked ROLE_PREV_YEAR, vbNullString
    AssertEqualsLong datePicker.ShownYear(), Year(Date) - PICKER_YEARS_BACK, _
                     "the picker cannot leave its window"
    AssertEqualsLong datePicker.ShownMonth(), 1, "it stops on January"

    ' an old date (the bug of the first version) is shown as today, highlighted
    datePicker.InitialiseWithDate "1999-12-17"
    AssertEqualsLong datePicker.ShownYear(), Year(Date), "an old date falls back to this year"
    AssertEqualsLong datePicker.ShownMonth(), Month(Date), "and to this month"
    AssertEqualsLong datePicker.SelectedGridIndex(), _
                     modPickers.GridIndexOf(Year(Date), Month(Date), _
                                            modDashboard.FormatDateText(Date)), _
                     "today is highlighted on opening"
    AssertEqualsText datePicker.SelectedDateDisplay(), _
                     modPickers.DateFormatText(modDashboard.FormatDateText(Date)), _
                     "the footer shows today the European way"

    ' the time picker in the start mode: two dropdowns, 10:00 is taken by a
    ' 30 minute booking
    Set timePicker = New frmTimePicker
    timePicker.InitialiseWithTime FIELD_START, "09:30", 60, "10:00+30"
    AssertEqualsText timePicker.FieldKey(), FIELD_START, "the picker knows its field"
    AssertEqualsLong timePicker.HourOptionCount(), 24, "every hour from 00 to 23 is offered"
    AssertEqualsText timePicker.HourOptionAt(0), "00", "the first hour is 00"
    AssertEqualsText timePicker.HourOptionAt(23), "23", "the last hour is 23"
    AssertEqualsLong timePicker.MinuteOptionCount(), 12, "the minutes run every 5 minutes"
    AssertEqualsText timePicker.MinuteOptionAt(6), "30", "the seventh minute is 30"
    AssertEqualsLong timePicker.Hours(), 9, "the hour of the initial value is chosen"
    AssertEqualsLong timePicker.Minutes(), 30, "the minute of the initial value is chosen"
    AssertEqualsText timePicker.CurrentText(), "09:30", "the current choice"

    ' a taken slot is refused, a free one is the answer
    timePicker.ChooseHour 10
    timePicker.ChooseMinute 0
    timePicker.ButtonClicked ROLE_OK, vbNullString
    AssertEqualsText timePicker.PickedTimeText, vbNullString, "a taken slot is not accepted"
    AssertTrue (Len(timePicker.HintText()) > 0), "the picker explains why"

    timePicker.ChooseHour 11
    timePicker.ButtonClicked ROLE_OK, vbNullString
    AssertEqualsText timePicker.PickedTimeText, "11:00", "the free hour is the result"

    ' the slot that only touches the booking is free (10:00 + 30 ends at 10:30)
    timePicker.InitialiseWithTime FIELD_START, "09:30", 60, "10:00+30"
    timePicker.ChooseHour 10
    timePicker.ChooseMinute 30
    timePicker.ButtonClicked ROLE_OK, vbNullString
    AssertEqualsText timePicker.PickedTimeText, "10:30", "touching intervals do not clash"

    ' the duration mode: every 5 minutes and nothing is busy; an empty field opens
    ' on 00:00 here as well
    timePicker.InitialiseWithTime FIELD_DURATION, vbNullString, 0, "09:00+30"
    AssertEqualsText timePicker.FieldKey(), FIELD_DURATION, "the duration mode"
    AssertEqualsText timePicker.CurrentText(), "00:00", "an empty duration opens on 00:00"
    AssertEqualsLong timePicker.MinuteOptionCount(), 12, "a duration may be any 5 minutes"
    timePicker.ChooseHour 1
    timePicker.ChooseMinute 30
    timePicker.ButtonClicked ROLE_OK, vbNullString
    AssertEqualsText timePicker.PickedTimeText, "01:30", "1 h 30 min is the answer"

    ' Cancel answers nothing (the caller clears the field)
    timePicker.InitialiseWithTime FIELD_START, "09:00", 30, vbNullString
    timePicker.ButtonClicked ROLE_CANCEL, vbNullString
    AssertEqualsText timePicker.PickedTimeText, vbNullString, "Cancel answers nothing"

    ' A field that was still empty opens on 00:00, so OK answers 00:00 and not
    ' "no value": the dropdowns of an empty form are never blank (??9.3)
    timePicker.InitialiseWithTime FIELD_START, vbNullString, 30, vbNullString
    AssertEqualsLong timePicker.Hours(), 0, "an empty start opens on the hour 00"
    AssertEqualsLong timePicker.Minutes(), 0, "an empty start opens on the minute 00"
    timePicker.ButtonClicked ROLE_OK, vbNullString
    AssertEqualsText timePicker.PickedTimeText, "00:00", "an empty form answers 00:00"
End Sub

' The Booking form: the two identical time fields (??7.5, ??9.4).
Private Sub Test_BookingForm()
    Dim values As Object
    Dim data As Object
    Dim result As Object

    modUi.Navigate SHEET_BOOKING
    modUiBooking.ClearForm

    AssertEqualsText modNames.ReadName(NAME_BOOKING_MODE), modText.Text(modText.KEY_BOOKING_NEW), _
                     "the mode of a clean form"

    ' writing a time field fills both cells and the previews
    modUiBooking.WriteTimeField FIELD_START, 9, 30
    AssertEqualsText modNames.ReadName(NAME_BOOKING_START_H), "09", "the hour cell"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_START_M), "30", "the minute cell"
    AssertEqualsText modUiBooking.ReadTimeField(FIELD_START), "09:30", "the field reads back"
    AssertEqualsLong modUiBooking.TimeFieldMinutes(FIELD_START), 570, "09:30 is 570 minutes"

    modUiBooking.WriteTimeField FIELD_DURATION, 1, 30
    AssertEqualsLong modUiBooking.TimeFieldMinutes(FIELD_DURATION), 90, "the duration is 90 minutes"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_DURATION), "1 h 30 min", "the duration preview"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_END), "11:00", "the end preview"

    ' an unreadable field is empty, not wrong
    modNames.WriteName NAME_BOOKING_DUR_M, "7"
    AssertEqualsText modUiBooking.ReadTimeField(FIELD_DURATION), vbNullString, _
                     "7 is not on the 5 minute step"
    AssertEqualsLong modUiBooking.TimeFieldMinutes(FIELD_DURATION), -1, "an unreadable field is -1"

    ' ApplyTimeText accepts what a user may type
    AssertTrue modUiBooking.ApplyTimeText(FIELD_START, "8:05"), "8:05 is accepted"
    AssertEqualsText modUiBooking.ReadTimeField(FIELD_START), "08:05", "and normalised"
    AssertFalse modUiBooking.ApplyTimeText(FIELD_START, "9:75"), "9:75 is refused"

    ' the values of the form for the domain
    modNames.WriteName NAME_BOOKING_DATE, "2026-09-21"
    modNames.WriteName NAME_BOOKING_CLIENT, "Ivan Petrov"
    modNames.WriteName NAME_BOOKING_SERVICE, "Haircut"
    modUiBooking.ApplyTimeText FIELD_START, "11:00"
    modUiBooking.WriteTimeField FIELD_DURATION, 0, 30
    Set values = modUiBooking.ReadForm()

    AssertEqualsText values(COL_DATE), "2026-09-21", "the date of the form"
    AssertEqualsText values(COL_CLIENT_ID), "1", "Ivan Petrov is client 1"
    AssertEqualsText values(COL_SERVICE_ID), "1", "Haircut is service 1"
    AssertEqualsText values(COL_START), "11:00", "the start of the form"
    AssertEqualsText values(COL_SERVICE_DURATION), "30", "the duration of the form"
    AssertFalse values.Exists(modRoutes.KEY_ID), "a new appointment has no id"

    modUiBooking.ClearForm
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT), vbNullString, "the form is clean"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_START_H), vbNullString, "the start is clean"
End Sub
' lists of _Lists (??6.5.4), never typed. The template adds those validations in a
' second pass, because the names of _Lists must exist first; this test guards it,
' since a validation that fails is silently missing and the form cannot be filled.
Private Sub Test_BookingDropdowns()
    Dim clientCell As Range
    Dim serviceCell As Range
    Dim clientValidation As Object
    Dim serviceValidation As Object
    Dim clientFormula As String
    Dim serviceFormula As String

    Set clientValidation = Nothing
    Set serviceValidation = Nothing
    clientFormula = vbNullString
    serviceFormula = vbNullString

    ' The range goes into a local before its Validation is read: a property read
    ' on the result of a chain is not reliable (??13).
    Set clientCell = modNames.RangeOfName(NAME_BOOKING_CLIENT)
    Set serviceCell = modNames.RangeOfName(NAME_BOOKING_SERVICE)

    On Error Resume Next
    Set clientValidation = clientCell.Validation
    Set serviceValidation = serviceCell.Validation
    clientFormula = CStr(clientValidation.Formula1)
    serviceFormula = CStr(serviceValidation.Formula1)
    On Error GoTo 0
    Err.Clear

    AssertTrue (Not clientValidation Is Nothing), "the client cell has a validation"
    AssertTrue (Not serviceValidation Is Nothing), "the service cell has a validation"
    If clientValidation Is Nothing Or serviceValidation Is Nothing Then Exit Sub

    ' 3 = xlValidateList; InCellDropdown = the arrow of the dropdown is shown
    AssertEqualsLong clientValidation.Type, 3, "the client cell validates a list"
    AssertEqualsLong serviceValidation.Type, 3, "the service cell validates a list"
    AssertTrue clientValidation.InCellDropdown, "the client cell shows the arrow"
    AssertTrue serviceValidation.InCellDropdown, "the service cell shows the arrow"
    AssertEqualsText clientFormula, "=" & NAME_LIST_CLIENTS, "the client list is list_clients"
    AssertEqualsText serviceFormula, "=" & NAME_LIST_SERVICES, "the service list is list_services"

    ' the ranges the dropdowns point at are the technical lists of _Lists
    AssertEqualsLong modNames.RangeOfName(NAME_LIST_CLIENTS).Rows.Count, 200, _
                     "list_clients holds 200 rows"
    AssertEqualsLong modNames.RangeOfName(NAME_LIST_CLIENTS).Row, DB_FIRST_ROW, _
                     "the first value of list_clients is on the first data row"
End Sub

' The tests of phase 3 need the documented working day of Config (??16.1).
Private Sub EnsureConfigDefaults()
    modNames.WriteName NAME_CONFIG_DAY_START, DEFAULT_DAY_START
    modNames.WriteName NAME_CONFIG_DAY_END, DEFAULT_DAY_END
    modNames.WriteName NAME_CONFIG_STEP, CStr(DEFAULT_STEP_MINUTES)
    modNames.WriteName NAME_CONFIG_NONWORK_DAYS, DEFAULT_NONWORK_WEEKDAYS
End Sub

Private Function AppointmentValues(ByVal dateText As String, ByVal startText As String, _
                                   ByVal durationText As String, ByVal clientId As String, _
                                   ByVal serviceId As String) As Object
    Dim data As Object

    Set data = CreateObject("Scripting.Dictionary")
    data(COL_DATE) = dateText
    data(COL_START) = startText
    data(COL_SERVICE_DURATION) = durationText
    data(COL_CLIENT_ID) = clientId
    data(COL_SERVICE_ID) = serviceId

    Set AppointmentValues = data
End Function

' Adds an appointment through the router and answers its new id (0 on failure).
Private Function AddAppointmentForTest(ByVal dateText As String, ByVal startText As String, _
                                       ByVal durationText As String, ByVal clientId As String, _
                                       ByVal serviceId As String) As Long
    Dim data As Object
    Dim result As Object

    Set data = AppointmentValues(dateText, startText, durationText, clientId, serviceId)
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, data)

    If result(modRoutes.RESULT_OK) = True Then
        AddAppointmentForTest = CLng(Val(CStr(result(modRoutes.RESULT_DATA)(COL_ID))))
    End If
End Function

Private Function DeleteAppointmentForTest(ByVal idValue As Long) As Boolean
    Dim data As Object
    Dim result As Object

    modUi.SetInteractiveUi False
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)

    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_DELETE, data)
    DeleteAppointmentForTest = (result(modRoutes.RESULT_OK) = True)
End Function

Private Sub Test_AppointmentsSave()
    Dim before As Long
    Dim values As Object
    Dim result As Object
    Dim idValue As Long
    Dim secondId As Long

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    before = modData.AppointmentsRowCount()

    ' The sample of the template: Ivan Petrov / Haircut on Monday 21 September
    ' 2026 at 09:00 for 30 minutes, so 11:00 is free.
    Set values = AppointmentValues("2026-09-21", "11:00", "45", "1", "3")
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, values)

    AssertTrue (result(modRoutes.RESULT_OK) = True), _
               "the appointment was saved: " & CStr(result(modRoutes.RESULT_MESSAGE))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "was added") > 0), _
               "the message says it was added"

    idValue = CLng(Val(CStr(result(modRoutes.RESULT_DATA)(COL_ID))))
    AssertEqualsLong modData.AppointmentsRowCount(), before + 1, "one more row"

    ' the values are in the table, normalised
    AssertEqualsText modData.AppointmentCell(idValue, COL_DATE), "2026-09-21", _
                     "the date is stored as ISO text"
    AssertEqualsText modData.AppointmentCell(idValue, COL_START), "11:00", _
                     "the start is stored as HH:MM"
    AssertEqualsText modData.AppointmentCell(idValue, COL_SERVICE_DURATION), "45", _
                     "the duration is stored in minutes"

    ' a second appointment on the same day, on a free 5 minute slot
    secondId = AddAppointmentForTest("2026-09-21", "13:15", "15", "2", "1")
    AssertTrue (secondId > 0), "a second appointment was accepted"
    AssertEqualsText modData.AppointmentCell(secondId, COL_START), "13:15", _
                     "the start is stored in the table"

    AssertTrue DeleteAppointmentForTest(secondId), "the second test appointment was removed"
    AssertTrue DeleteAppointmentForTest(idValue), "the first test appointment was removed"
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the row count is back"
End Sub

Private Sub Test_AppointmentValidation()
    Dim before As Long
    Dim result As Object

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    before = modData.AppointmentsRowCount()

    ' no client
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "11:00", "30", "0", "1"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a booking without a client is refused"

    ' a client that does not exist
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "11:00", "30", "99", "1"))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "does not exist") > 0), _
               "an unknown client is named"

    ' no service
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "11:00", "30", "1", "0"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a booking without a service is refused"

    ' no date
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("", "11:00", "30", "1", "1"))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "date") > 0), _
               "the missing date is named"

    ' a non working day: Saturday 19 September 2026
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-19", "11:00", "30", "1", "1"))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "working day") > 0), _
               "a Saturday is refused"

    ' no start
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "", "30", "1", "1"))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "start time") > 0), _
               "the missing start is named"

    ' a start that is not a time
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "nine", "30", "1", "1"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "'nine' is not a start"

    ' a start that is not on the step of the grid
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "09:07", "30", "1", "1"))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "5 minute step") > 0), _
               "the step rule is quoted"

    ' before the working day
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "08:00", "30", "1", "1"))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "must fit between") > 0), _
               "08:00 is before the day starts"

    ' past the end of the working day
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "18:45", "60", "1", "1"))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "must fit between") > 0), _
               "18:45 + 60 leaves the day"

    ' no duration and a zero duration
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "11:00", "", "1", "1"))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "duration") > 0), _
               "the missing duration is named"

    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "11:00", "0", "1", "1"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a zero duration is refused"

    AssertEqualsLong modData.AppointmentsRowCount(), before, "nothing was written"
End Sub

Private Sub Test_Conflicts()
    Dim before As Long
    Dim result As Object
    Dim firstId As Long
    Dim secondId As Long

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    before = modData.AppointmentsRowCount()

    ' Monday 21 September 2026 already holds 09:00 + 30 (the sample).
    ' 10:00 + 60 starts exactly when that one ends (??8.2.7: no clash).
    firstId = AddAppointmentForTest("2026-09-21", "10:00", "60", "1", "1")
    AssertTrue (firstId > 0), "an appointment right after another one is allowed"

    ' 09:30 + 45 runs into the 09:00 booking and into the 10:00 booking.
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "09:30", "45", "2", "1"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "an overlapping appointment is refused"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "overlap") > 0), _
               "the message says that the time is taken"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "10:00") > 0), _
               "the message names the first appointment in the way"

    ' the same time on another day is free (Wednesday 23 September)
    secondId = AddAppointmentForTest("2026-09-23", "09:30", "45", "2", "1")
    AssertTrue (secondId > 0), "the same time on another day is free"

    ' a start in the middle of a booking is refused as well
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE, _
                                  AppointmentValues("2026-09-21", "09:15", "45", "3", "1"))
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a start inside a booking is refused"

    AssertEqualsLong modData.AppointmentsRowCount(), before + 2, "the two test rows are there"

    AssertTrue DeleteAppointmentForTest(secondId), "the second test appointment was removed"
    AssertTrue DeleteAppointmentForTest(firstId), "the first test appointment was removed"
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the row count is back"
End Sub

Private Sub Test_AppointmentEdit()
    Dim before As Long
    Dim result As Object
    Dim data As Object
    Dim idValue As Long
    Dim values As Object

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    before = modData.AppointmentsRowCount()

    idValue = AddAppointmentForTest("2026-09-23", "11:00", "30", "3", "2")
    AssertTrue (idValue > 0), "the appointment to edit exists"

    ' appointments.edit fills the form and shows the sheet
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_EDIT, data)

    AssertTrue (result(modRoutes.RESULT_OK) = True), "the appointment was loaded for editing"
    AssertEqualsText modNames.ReadName(NAME_STATE_EDIT_ID), CStr(idValue), "the form is in edit mode"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_DATE), "23.09.2026", _
                     "the date of the row is shown the European way"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT), "Elena Georgieva", "the client of the row"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_SERVICE), "Hair coloring", "the service of the row"
    AssertEqualsText modUiBooking.ReadTimeField(FIELD_START), "11:00", "the start of the row"
    AssertEqualsLong modUiBooking.TimeFieldMinutes(FIELD_DURATION), 30, "the duration of the row"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_MODE), _
                     modText.TextFormat(modText.KEY_BOOKING_EDIT, CStr(idValue)), "the mode shows Editing"
    AssertEqualsText ActiveSheet.Name, SHEET_BOOKING, "the Booking sheet is in front"

    ' the form path saves the change in place (??8.2.10)
    modUiBooking.ApplyTimeText FIELD_START, "12:00"
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE)

    AssertTrue (result(modRoutes.RESULT_OK) = True), _
               "the edit was saved: " & CStr(result(modRoutes.RESULT_MESSAGE))
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "was updated") > 0), _
               "the message says it was updated"
    AssertEqualsLong modData.AppointmentsRowCount(), before + 1, "no row was added"
    AssertEqualsText modData.AppointmentCell(idValue, COL_START), "12:00", _
                     "the start was changed"
    AssertEqualsText modNames.ReadName(NAME_STATE_EDIT_ID), vbNullString, "the form left the edit mode"

    ' the row does not clash with itself while it is edited
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_EDIT, data)
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "an unchanged appointment can be saved again"

    AssertTrue DeleteAppointmentForTest(idValue), "the test appointment was removed"
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the row count is back"

    ' an id that does not exist
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = "999"
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_EDIT, data)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "editing an unknown appointment fails"
End Sub

Private Sub Test_AppointmentsList()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim clientRow As Long
    Dim before As Long
    Dim idValue As Long
    Dim result As Object
    Dim payload As Object
    Dim info As Object
    Dim actionCell As Range
    Dim link As Object

    Set ws = ThisWorkbook.Worksheets(SHEET_APPOINTMENTS)
    modUi.SetInteractiveUi False
    modBackup.SetEnabled False
    before = modData.AppointmentsRowCount()
    firstRow = modNames.RangeOfName(NAME_GRID_APPOINTMENTS_FIRST).Row

    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_LIST)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "appointments.list succeeds"
    Set info = result(modRoutes.RESULT_DATA)
    AssertEqualsLong info("rows"), before, "the list reports every appointment"
    AssertEqualsText ws.Cells(firstRow, APPOINTMENTS_COL_ID).Text, "3", _
                     "the newest appointment is listed first"
    AssertEqualsText ws.Cells(firstRow, APPOINTMENTS_COL_CLIENT).Text, "Elena Georgieva", _
                     "the client name is joined"
    AssertEqualsText ws.Cells(firstRow, APPOINTMENTS_COL_PHONE).Text, "0899123456", _
                     "the phone retains its leading zero"
    AssertEqualsText ws.Cells(firstRow, APPOINTMENTS_COL_EMAIL).Text, "elena@example.com", _
                     "the email is joined"
    AssertEqualsText ws.Cells(firstRow, APPOINTMENTS_COL_END).Text, "14:45", _
                     "the end time is calculated"

    modUi.Navigate SHEET_APPOINTMENTS
    modApp.RunCellAction SHEET_APPOINTMENTS, NAME_BTN_APPOINTMENTS_REFRESH
    AssertEqualsText ActiveSheet.Name, SHEET_APPOINTMENTS, "refresh stays on Appointments"
    AssertEqualsText modNames.ReadName(NAME_APPOINTMENTS_TITLE), _
                     "Appointments (" & CStr(before) & ")", "Refresh redraws the table"

    Set actionCell = ws.Cells(firstRow, APPOINTMENTS_COL_EDIT)
    Set link = HyperlinkAt(ws, firstRow, APPOINTMENTS_COL_EDIT)
    AssertTrue (Not link Is Nothing), "the appointment Edit cell has a hyperlink"
    If Not link Is Nothing Then
        AssertEqualsText link.SubAddress, NAME_GRID_APPOINTMENTS_FIRST, _
                         "the Edit link points to the permanent table anchor"
    End If
    Set payload = modApp.ClickPayloadFor(SHEET_APPOINTMENTS, actionCell)
    AssertEqualsText payload(modRoutes.KEY_ACTION), modRoutes.ACTION_APPOINTMENTS_EDIT, _
                     "the Edit column dispatches appointments.edit"
    AssertEqualsText payload(modRoutes.KEY_ID), "3", _
                     "the Edit payload carries the id of the newest appointment"

    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_EDIT, payload)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "Edit opens the Booking form"
    AssertEqualsText ActiveSheet.Name, SHEET_BOOKING, "Edit navigates to Booking"
    AssertEqualsText modNames.ReadName(NAME_STATE_EDIT_ID), "3", "the booking is in edit mode"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT), "Elena Georgieva", "the client is loaded"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_DATE), "24.09.2026", "the date is loaded"

    modUi.Navigate SHEET_DASHBOARD
    idValue = AddAppointmentForTest("2026-09-25", "16:00", "30", "1", "1")
    AssertTrue (idValue > 0), "the appointment used for automatic refresh was added"
    AssertEqualsText modNames.ReadName(NAME_APPOINTMENTS_TITLE), _
                     "Appointments (" & CStr(before + 1) & ")", _
                     "a successful booking refreshes Appointments automatically"

    firstRow = modNames.RangeOfName(NAME_GRID_APPOINTMENTS_FIRST).Row
    ' The newest appointment is the first row of the list, and the one added just above
    ' (25 September) is the newest of all, so its Delete link sits on the top row. It
    ' must be the row that is deleted: clicking the bottom one would remove the sample
    ' appointment of 21 September and the tests that follow would lose the day.
    Set actionCell = ws.Cells(firstRow, APPOINTMENTS_COL_DELETE)
    Set payload = modApp.ClickPayloadFor(SHEET_APPOINTMENTS, actionCell)
    AssertEqualsText payload(modRoutes.KEY_ACTION), modRoutes.ACTION_APPOINTMENTS_DELETE, _
                     "the Delete column dispatches appointments.delete"
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_DELETE, payload)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the appointment was deleted from the list"
    AssertFalse modData.RowExists(SHEET_DB_APPOINTMENTS, idValue), "the appointment row is gone"
    AssertEqualsText modNames.ReadName(NAME_APPOINTMENTS_TITLE), _
                     "Appointments (" & CStr(before) & ")", "delete refreshes the list"

    Set payload = CreateObject("Scripting.Dictionary")
    payload(modRoutes.KEY_ID) = "1"
    modUi.Navigate SHEET_CLIENTS_ADMIN
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_EDIT, payload)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the client form opens for editing"
    modNames.WriteName NAME_CLIENT_PHONE, "0777000000"
    modNames.WriteName NAME_CLIENT_EMAIL, "ivan.updated@example.com"
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_ADD)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the client contact details were updated"
    ' Ivan Petrov is client 1 and his appointment is the oldest one, so under the rule
    ' "the newest date on top" it sits on the last row of the list - the top row is the
    ' newest appointment of all.
    clientRow = firstRow + before - 1
    AssertEqualsText ws.Cells(clientRow, APPOINTMENTS_COL_PHONE).Text, "0777000000", _
                     "client updates refresh the phone in Appointments"
    AssertEqualsText ws.Cells(clientRow, APPOINTMENTS_COL_EMAIL).Text, _
                     "ivan.updated@example.com", "client updates refresh the email"

    payload(modRoutes.KEY_ID) = "1"
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_EDIT, payload)
    modNames.WriteName NAME_CLIENT_PHONE, "0888123456"
    modNames.WriteName NAME_CLIENT_EMAIL, "ivan@example.com"
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_ADD)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "sample client contact details were restored"

    modUi.Navigate SHEET_DASHBOARD
End Sub

Private Sub Test_AppointmentDelete()
    Dim before As Long
    Dim result As Object
    Dim data As Object
    Dim idValue As Long

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    before = modData.AppointmentsRowCount()

    idValue = AddAppointmentForTest("2026-09-23", "15:00", "30", "1", "1")
    AssertTrue (idValue > 0), "the appointment to delete exists"

    ' the two step confirmation of ??6.6
    modApp.ConfirmReset
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)

    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_DELETE, data)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the delete needs no second click"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "was deleted") > 0), _
               "the message says it was deleted"
    AssertFalse modData.RowExists(SHEET_DB_APPOINTMENTS, idValue), "the row is gone"
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the row count is back"

    ' a second delete finds nothing
    modUi.SetInteractiveUi False
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_DELETE, data)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "a deleted appointment cannot go twice"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "does not exist") > 0), _
               "the message says it is gone"

    ' an id that is not a number
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = "abc"
    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_DELETE, data)
    AssertFalse (result(modRoutes.RESULT_OK) = True), "'abc' is not an id"

    ' leave no dialogs switched on for the tests that follow
    modUi.SetInteractiveUi False
End Sub

Private Sub Test_BookingFormFlow()
    Dim before As Long
    Dim result As Object
    Dim idValue As Long

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    before = modData.AppointmentsRowCount()

    ' the user fills the form by hand and presses Save
    modUi.Navigate SHEET_BOOKING
    modUiBooking.ClearForm
    modNames.WriteName NAME_BOOKING_DATE, "2026-09-24"
    modNames.WriteName NAME_BOOKING_CLIENT, "Ivan Petrov"
    modNames.WriteName NAME_BOOKING_SERVICE, "Manicure"
    modUiBooking.WriteTimeField FIELD_START, 16, 0
    modUiBooking.WriteTimeField FIELD_DURATION, 0, 30

    ' the previews follow the typed values
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT_ID), "1", "the client id preview"
    AssertTrue (Len(modNames.ReadName(NAME_BOOKING_PRICE)) > 0), "the price preview is filled"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_SERVICE_DURATION), "45 min", _
                     "the duration of the chosen service"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_END), "16:30", "the end preview"

    Set result = modApp.AppRouter(modRoutes.ACTION_APPOINTMENTS_SAVE)
    AssertTrue (result(modRoutes.RESULT_OK) = True), _
               "the form was saved: " & CStr(result(modRoutes.RESULT_MESSAGE))

    idValue = modData.AppointmentIdAt("2026-09-24", "16:00")
    AssertTrue (idValue > 0), "the appointment of the form is in the table"
    AssertEqualsText modData.AppointmentCell(idValue, COL_CLIENT_ID), "1", _
                     "the client of the form"
    AssertEqualsText modData.AppointmentCell(idValue, COL_SERVICE_ID), "3", _
                     "the service of the form"
    AssertEqualsText modData.AppointmentCell(idValue, COL_SERVICE_DURATION), "30", _
                     "the duration of the form"

    ' the form is ready for the next appointment
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT), vbNullString, "the form was cleared"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_MODE), _
                     modText.Text(modText.KEY_BOOKING_NEW), "and back in the new mode"

    AssertTrue DeleteAppointmentForTest(idValue), "the test appointment was removed"
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the row count is back"
End Sub

Private Sub Test_ServiceFillsDuration()
    EnsureConfigDefaults
    modUi.SetInteractiveUi False

    ' an empty duration is filled from the service (??7.5)
    modUi.Navigate SHEET_BOOKING
    modUiBooking.ClearForm
    modNames.WriteName NAME_BOOKING_DATE, "2026-09-23"
    modNames.WriteName NAME_BOOKING_CLIENT, "Ivan Petrov"
    modNames.WriteName NAME_BOOKING_SERVICE, "Hair coloring"
    modApp.OnChange SHEET_BOOKING, modNames.RangeOfName(NAME_BOOKING_SERVICE)

    AssertEqualsText modNames.ReadName(NAME_BOOKING_DUR_H), "01", "the hour cell of the service"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_DUR_M), "30", "the minute cell of the service"
    AssertEqualsLong modUiBooking.TimeFieldMinutes(FIELD_DURATION), 90, "90 minutes from the service"

    ' a duration the user chose is not overwritten
    modUiBooking.WriteTimeField FIELD_DURATION, 0, 45
    modApp.OnChange SHEET_BOOKING, modNames.RangeOfName(NAME_BOOKING_SERVICE)
    AssertEqualsLong modUiBooking.TimeFieldMinutes(FIELD_DURATION), 45, "45 minutes stay"

    ' a change in another sheet does nothing
    modApp.OnChange SHEET_CLIENTS_ADMIN, modNames.RangeOfName(NAME_CLIENT_NAME)
    AssertEqualsLong modUiBooking.TimeFieldMinutes(FIELD_DURATION), 45, "another sheet is ignored"

    modUiBooking.ClearForm
End Sub

Private Sub Test_PickActions()
    Dim result As Object
    Dim data As Object
    Dim info As Object

    EnsureConfigDefaults
    modUi.SetInteractiveUi False

    modUi.Navigate SHEET_BOOKING
    modUiBooking.ClearForm

    ' form.pick_date with a value (the picker stays closed - ??9.4)
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_VALUE) = "2026-09-23"
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_PICK_DATE, data)

    AssertTrue (result(modRoutes.RESULT_OK) = True), "the date was taken over"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_DATE), "23.09.2026", _
                     "the date cell shows the European form"
    Set info = result(modRoutes.RESULT_DATA)
    AssertEqualsText info(COL_DATE), "2026-09-23", "the payload carries the ISO date"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "23.09.2026") > 0), _
               "the message quotes the date the European way"

    ' form.pick_time for the start
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_FIELD) = FIELD_START
    data(modRoutes.KEY_VALUE) = "09:30"
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_PICK_TIME, data)

    AssertTrue (result(modRoutes.RESULT_OK) = True), "the start was taken over"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_START_H), "09", "the hour cell"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_START_M), "30", "the minute cell"

    ' form.pick_time for the duration: 01:30 means 1 h 30 min
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_FIELD) = FIELD_DURATION
    data(modRoutes.KEY_VALUE) = "01:30"
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_PICK_TIME, data)

    AssertTrue (result(modRoutes.RESULT_OK) = True), "the duration was taken over"
    AssertEqualsLong modUiBooking.TimeFieldMinutes(FIELD_DURATION), 90, "90 minutes"

    ' a value that is not a time is refused
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_FIELD) = FIELD_START
    data(modRoutes.KEY_VALUE) = "9:75"
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_PICK_TIME, data)

    AssertFalse (result(modRoutes.RESULT_OK) = True), "9:75 is not a time"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_START_M), "30", "the cells were not changed"

    ' a cancelled picker changes nothing
    modUiBooking.ClearForm
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_FIELD) = FIELD_DURATION
    Set result = modApp.AppRouter(modRoutes.ACTION_FORM_PICK_TIME, data)

    AssertFalse (result(modRoutes.RESULT_OK) = True), "an empty picker is a cancel"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_DUR_M), vbNullString, "nothing was written"

    modUiBooking.ClearForm
End Sub

' =====================================================================
'  phase 4: the body of the dashboard
' =====================================================================

' The dashboard shows a given day, or the week / the month around it: the tests
' drive it through the router, like the user does.
Private Sub ShowDashboard(ByVal viewKey As String, ByVal dateText As String)
    Dim data As Object

    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_DATE) = dateText
    modApp.AppRouter modRoutes.ACTION_DASHBOARD_OPEN_DAY, data

    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_VIEW) = viewKey
    modApp.AppRouter modRoutes.ACTION_DASHBOARD_VIEW, data

    modUi.Navigate SHEET_DASHBOARD
End Sub

' The hyperlink of one cell of a rendered row (Nothing when it has none).
Private Function HyperlinkAt(ByRef ws As Worksheet, ByVal rowIndex As Long, _
                             ByVal columnIndex As Long) As Object
    Dim link As Object

    Set link = Nothing
    On Error Resume Next
    Set link = ws.Cells(rowIndex, columnIndex).Hyperlinks(1)
    On Error GoTo 0

    Set HyperlinkAt = link
End Function

' The daily view: one row per slot of the working day, drawn at the resolution of
' Config (??6.9). A booked slot carries the appointment, the slots after it the
' continuation mark, a free slot a button that books a new appointment there.
' The sheet row a slot of the dashboard body sits in: the first row of the block
' plus the number of steps between the start of the working day and the minute
' asked for (the resolution of Config decides how many slots there are, ??6.9).
Private Function SlotRowOfTest(ByVal minutesValue As Long) As Long
    SlotRowOfTest = modNames.RangeOfName(NAME_CAL_FIRST).Row + _
        (minutesValue - modAppointments.DayStartMinutes()) \ modAppointments.StepMinutes()
End Function

Private Sub Test_DashboardDaily()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim slotRow As Long
    Dim idValue As Long
    Dim before As Long
    Dim link As Object

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    modBackup.SetEnabled False
    before = modData.AppointmentsRowCount()

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    firstRow = modNames.RangeOfName(NAME_CAL_FIRST).Row

    ' A day of its own: Friday 25 September 2026 (the samples sit elsewhere).
    idValue = AddAppointmentForTest("2026-09-25", "15:00", "45", "1", "3")
    AssertTrue (idValue > 0), "the appointment of the daily view exists"

    ShowDashboard VIEW_DAY, "2026-09-25"
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_DAY, "the daily view is on"
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), "2026-09-25", "s_date is the day"

    ' The resolution of Config (5 minutes by default) is what the day is drawn in:
    ' 10 hours of 5 minute slots = 120 rows, and the block holds all of them.
    AssertEqualsLong modDashboard.SlotRowCount(), 120, "the day has 120 slots of 5 minutes"
    AssertTrue (Not modDashboard.SlotsAreCapped()), "the whole working day fits the block"
    AssertEqualsText ws.Cells(firstRow, DASH_COL_START).Text, "09:00", _
                     "the first row is the first slot of the day"
    AssertEqualsText ws.Cells(firstRow + 119, DASH_COL_START).Text, "18:55", _
                     "the last row is the last slot of the day"
    AssertEqualsText ws.Cells(firstRow + 120, DASH_COL_START).Text, vbNullString, _
                     "below the grid the block is empty"

    ' the labels of the body
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_COL_START).Text, _
                     modText.Text(modText.KEY_DASH_COL_TIME), "the first label is the time"
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_COL_PRICE).Text, _
                     modText.Text(modText.KEY_DASH_COL_PRICE), "the price has its label"
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_COL_EDIT).Text, _
                     modText.Text(modText.KEY_DASH_COL_EDIT), "the Edit column is labelled"
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_COL_ID).Text, _
                     modText.Text(modText.KEY_DASH_COL_ID), "the id column is labelled"

    ' the slot the appointment starts in (15:00 = the 73rd slot of the day)
    slotRow = SlotRowOfTest(900)
    AssertEqualsText ws.Cells(slotRow, DASH_COL_START).Text, "15:00", "the start of the slot"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_END).Text, "15:45", "the end of the appointment"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_CLIENT).Text, "Ivan Petrov", "the client"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_SERVICE).Text, "Manicure", "the service"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_DURATION).Text, "45 min", "the duration"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_PRICE).Text, "30.00", "the price"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_ID).Text, CStr(idValue), "the id of the slot"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_EDIT).Text, LIST_EDIT_TEXT, "the slot has [Edit]"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_DELETE).Text, LIST_ACTION_TEXT, _
                     "the slot has [Delete]"

    ' the button points at the anchor of the body (??6.10)
    Set link = HyperlinkAt(ws, slotRow, DASH_COL_EDIT)
    AssertTrue (Not link Is Nothing), "the [Edit] cell carries a hyperlink"
    If Not link Is Nothing Then
        AssertEqualsText link.SubAddress, NAME_CAL_FIRST, "the link points at cal_first"
    End If

    ' The nine slots of the 45 minutes: the ones after the start are painted and
    ' carry the continuation mark, the slot after them is free again.
    AssertEqualsText ws.Cells(slotRow + 1, DASH_COL_START).Text, "15:05", _
                     "the slot after the start is drawn"
    AssertEqualsText ws.Cells(slotRow + 1, DASH_COL_CLIENT).Text, _
                     modText.TextFormat(modText.KEY_DASH_TAIL, CStr(idValue)), _
                     "the slot after the start continues the appointment"
    AssertEqualsDouble ws.Cells(slotRow + 1, DASH_COL_SERVICE).Interior.Color, _
                       CDbl(COLOR_BOOKED), 0, "the whole row of a continuation is painted"
    AssertEqualsText ws.Cells(slotRow + 9, DASH_COL_START).Text, "15:45", _
                     "the slot after the appointment starts at 15:45"
    AssertEqualsText ws.Cells(slotRow + 9, DASH_COL_CLIENT).Text, vbNullString, _
                     "a free slot carries no appointment"

    ' a free slot is a button: its time books a new appointment in that slot
    AssertEqualsLong ws.Cells(firstRow, DASH_COL_START).Hyperlinks.Count, 1, _
                     "a free slot carries a link"
    Set link = HyperlinkAt(ws, firstRow, DASH_COL_START)
    AssertTrue (Not link Is Nothing), "the free slot is clickable"
    If Not link Is Nothing Then
        AssertEqualsText link.SubAddress, NAME_CAL_FIRST, "the free slot links to the anchor"
        AssertTrue (InStr(1, CStr(link.ScreenTip), "09:00") > 0), _
                   "the tooltip of the free slot names its time"
    End If

    ' [Edit] of the slot loads the appointment into Booking
    modApp.OnDynamicClick SHEET_DASHBOARD, ws.Cells(slotRow, DASH_COL_EDIT)
    AssertEqualsText modNames.ReadName(NAME_STATE_EDIT_ID), CStr(idValue), _
                     "the click loaded the appointment for editing"
    AssertEqualsText ActiveSheet.Name, SHEET_BOOKING, "Booking is in front"
    AssertEqualsText modNames.ReadName(NAME_BTN_BOOKING_CANCEL), _
                     modText.Text(modText.KEY_BTN_DELETE), _
                     "Booking offers Delete while an appointment is being edited"
    modUiBooking.ClearForm
    modUi.Navigate SHEET_DASHBOARD

    ' a continuation row is painted, not clickable: only the first row of an
    ' appointment carries its link, so nobody takes a busy slot for a free one
    ShowDashboard VIEW_DAY, "2026-09-25"
    AssertEqualsLong ws.Cells(slotRow + 1, DASH_COL_CLIENT).Hyperlinks.Count, 0, _
                     "the continuation row of the day carries no link"
    AssertEqualsLong ws.Cells(slotRow + 1, DASH_COL_START).Hyperlinks.Count, 0, _
                     "and neither does the time of that row"

    ' a free slot books a new appointment on the shown day and at its time
    ShowDashboard VIEW_DAY, "2026-09-25"
    modApp.OnDynamicClick SHEET_DASHBOARD, ws.Cells(firstRow, DASH_COL_START)
    AssertEqualsText modNames.ReadName(NAME_BOOKING_DATE), "25.09.2026", _
                     "the free slot sends the shown day (European)"
    AssertEqualsText modUiBooking.ReadTimeField(FIELD_START), "09:00", _
                     "the free slot sends its time"
    AssertEqualsLong modData.AppointmentsRowCount(), before + 1, _
                     "opening the form did not save anything"
    modUiBooking.ClearForm
    modUi.Navigate SHEET_DASHBOARD

    ' [Delete] asks once and removes the row on the second click (??6.6)
    ShowDashboard VIEW_DAY, "2026-09-25"
    modApp.OnDynamicClick SHEET_DASHBOARD, ws.Cells(slotRow, DASH_COL_DELETE)
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the click deletes at once"
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the second click deletes the row"
    modUi.SetInteractiveUi False
    modApp.ConfirmReset

    ' the slot is free again and books a new appointment
    ShowDashboard VIEW_DAY, "2026-09-25"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_START).Text, "15:00", "the slot keeps its time"
    AssertEqualsText ws.Cells(slotRow, DASH_COL_CLIENT).Text, vbNullString, _
                     "the deleted appointment left the slot"
    AssertEqualsLong ws.Cells(slotRow, DASH_COL_START).Hyperlinks.Count, 1, _
                     "the slot is a button again"
End Sub

Private Sub Test_DashboardWeekly()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim idValue As Long
    Dim before As Long
    Dim step As Long
    Dim slotRow As Long
    Dim dayColumn As Long
    Dim saturdayColumn As Long
    Dim link As Object

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    modBackup.SetEnabled False
    before = modData.AppointmentsRowCount()

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    firstRow = modNames.RangeOfName(NAME_CAL_FIRST).Row

    ' Tuesday 29 September 2026: one appointment of an hour at 10:00.
    idValue = AddAppointmentForTest("2026-09-29", "10:00", "60", "2", "2")
    AssertTrue (idValue > 0), "the appointment of the weekly view exists"

    ShowDashboard VIEW_WEEK, "2026-09-28"
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_WEEK, "the weekly view is on"
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), "2026-09-28", "the anchor is the Monday"

    ' The grid: one row per slot at the resolution of Config (5 minutes by
    ' default), so the whole working day (10 hours = 120 rows) fits the block.
    step = modAppointments.StepMinutes()
    AssertEqualsLong step, 5, "the grid is drawn at the resolution of Config"
    AssertEqualsLong modDashboard.SlotRowCount(), 120, "120 slots cover the day"
    AssertTrue (Not modDashboard.SlotsAreCapped()), "the whole week fits the block"
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_WEEK_TIME_COL).Text, _
                     modText.Text(modText.KEY_DASH_COL_TIME), "the time column is labelled"
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_WEEK_FIRST_DAY_COL + 1).Text, _
                     "Tue 29", "the second column is Tuesday"
    ' The grid ends with the seven days: there is no action column any more, the
    ' appointment is deleted in Booking (??6.9).
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_WEEK_FIRST_DAY_COL + 6).Text, _
                     "Sun 4", "the last column of the grid is Sunday"
    AssertEqualsText ws.Cells(firstRow, DASH_WEEK_FIRST_DAY_COL + 7).Text, vbNullString, _
                     "there is no action column"

    slotRow = SlotRowOfTest(600)                  ' 10:00 on Tuesday
    dayColumn = DASH_WEEK_FIRST_DAY_COL + 1
    saturdayColumn = DASH_WEEK_FIRST_DAY_COL + 5

    AssertEqualsText ws.Cells(slotRow, DASH_WEEK_TIME_COL).Text, "10:00", "the row is 10:00"
    AssertEqualsText ws.Cells(slotRow, dayColumn).Text, _
                     modText.TextFormat(modText.KEY_DASH_SLOT, _
                         modText.TextFormat(modText.KEY_DASH_RANGE, "10:00", "11:00"), _
                         modText.TextFormat(modText.KEY_DASH_WHO, "Maria Ivanova", _
                                            "Hair coloring"), _
                         CStr(idValue)), "the booked cell names the appointment"
    AssertEqualsText ws.Cells(slotRow + 1, dayColumn).Text, _
                     modText.TextFormat(modText.KEY_DASH_TAIL, CStr(idValue)), _
                     "the next row continues the appointment"
    AssertEqualsText ws.Cells(slotRow + 11, dayColumn).Text, _
                     modText.TextFormat(modText.KEY_DASH_TAIL, CStr(idValue)), _
                     "the last slot of the hour continues it too"
    AssertEqualsDouble ws.Cells(slotRow + 11, dayColumn).Interior.Color, _
                       CDbl(COLOR_BOOKED), 0, "a booked cell is painted as booked"
    AssertEqualsText ws.Cells(slotRow + 12, dayColumn).Text, " ", _
                     "the slot after the appointment is free again"

    ' a free cell of a working day is clickable, a Saturday cell is not
    AssertEqualsLong ws.Cells(slotRow, DASH_WEEK_FIRST_DAY_COL).Hyperlinks.Count, 1, _
                     "a free cell of a working day carries a link"
    Set link = HyperlinkAt(ws, slotRow, DASH_WEEK_FIRST_DAY_COL)
    AssertTrue (Not link Is Nothing), "the free cell is clickable"
    If Not link Is Nothing Then
        AssertEqualsText link.SubAddress, NAME_CAL_FIRST, "the free cell links to the anchor"
    End If

    AssertEqualsLong ws.Cells(slotRow, saturdayColumn).Hyperlinks.Count, 0, _
                     "a day that is not a working day cannot be booked"
    AssertEqualsDouble ws.Cells(slotRow, saturdayColumn).Interior.Color, _
                       CDbl(COLOR_NOWORK), 0, "the Saturday column is muted"

    ' The tooltip of a booked cell names the appointment (client, service, day,
    ' time, id), so a click on it says what it will do.
    Set link = HyperlinkAt(ws, slotRow, dayColumn)
    AssertTrue (Not link Is Nothing), "the booked cell carries a hyperlink"
    If Not link Is Nothing Then
        AssertTrue (InStr(1, CStr(link.ScreenTip), "Maria Ivanova") > 0), _
                   "the tooltip of the booked cell names the client"
        AssertTrue (InStr(1, CStr(link.ScreenTip), "29.09.2026") > 0), _
                   "the tooltip of the booked cell names the day"
    End If

    ' one appointment reads with its client, service, day and time (??6.9)
    AssertEqualsText modAppointments.AppointmentLabel(idValue), _
                     "Maria Ivanova - Hair coloring, Tue 29.09.2026 10:00", _
                     "the label of an appointment"

    ' the continuation cells are painted, not clickable: only the first row of the
    ' appointment carries its link (a free cell books, a busy one does not)
    AssertEqualsLong ws.Cells(slotRow + 1, dayColumn).Hyperlinks.Count, 0, _
                     "a continuation cell of the week carries no link"
    AssertEqualsLong ws.Cells(slotRow + 11, dayColumn).Hyperlinks.Count, 0, _
                     "and neither does the last row of the appointment"

    ' a free cell books a new appointment at the time of the row
    ShowDashboard VIEW_WEEK, "2026-09-28"
    modApp.OnDynamicClick SHEET_DASHBOARD, ws.Cells(slotRow, DASH_WEEK_FIRST_DAY_COL)
    AssertEqualsText modNames.ReadName(NAME_BOOKING_DATE), "28.09.2026", _
                     "the free cell of Monday sends its day (European)"
    AssertEqualsText modUiBooking.ReadTimeField(FIELD_START), "10:00", _
                     "the free cell of Monday sends its time"
    modUiBooking.ClearForm
    modUi.Navigate SHEET_DASHBOARD

    ' An appointment is deleted in Booking: the cell next to the time fields asks
    ' "Delete" while one is being edited and "Cancel" on a new appointment (??6.9).
    modApp.OnDynamicClick SHEET_DASHBOARD, ws.Cells(slotRow, dayColumn)
    AssertEqualsText modNames.ReadName(NAME_STATE_EDIT_ID), CStr(idValue), _
                     "the booked cell opened the appointment"
    AssertEqualsText modNames.ReadName(NAME_BTN_BOOKING_CANCEL), _
                     modText.Text(modText.KEY_BTN_DELETE), "the button says Delete"
    Set link = HyperlinkAt(ThisWorkbook.Worksheets(SHEET_BOOKING), _
                           modNames.RangeOfName(NAME_BTN_BOOKING_CANCEL).Row, _
                           modNames.RangeOfName(NAME_BTN_BOOKING_CANCEL).Column)
    AssertTrue (Not link Is Nothing), "the relabelled button is still a button"

    modApp.RunCellAction SHEET_BOOKING, NAME_BTN_BOOKING_CANCEL
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the Delete button deletes at once"
    AssertEqualsLong modData.AppointmentsRowCount(), before, "the second click deletes the row"
    modUi.SetInteractiveUi False
    modApp.ConfirmReset

    ' the form left the edit mode with the appointment and offers Cancel again
    AssertEqualsText modNames.ReadName(NAME_STATE_EDIT_ID), vbNullString, _
                     "the form forgot the deleted appointment"
    AssertEqualsText modNames.ReadName(NAME_BTN_BOOKING_CANCEL), _
                     modText.Text(modText.KEY_BTN_CANCEL), "the button says Cancel again"
    modUi.Navigate SHEET_DASHBOARD
End Sub

Private Sub Test_DashboardMonthly()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim gridRow As Long
    Dim gridColumn As Long
    Dim todayText As String
    Dim todayIndex As Long

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    firstRow = modNames.RangeOfName(NAME_CAL_FIRST).Row

    ' September 2026: the month of the sample data (1 September is a Tuesday, so
    ' the grid starts with 31 August).
    ShowDashboard VIEW_MONTH, "2026-09-21"
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_MONTH, "the monthly view is on"
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_MONTH_FIRST_COL).Text, "Mon", _
                     "the first column of the calendar is Monday"
    AssertEqualsText ws.Cells(modDashboard.BodyHeaderRow(), DASH_MONTH_FIRST_COL + 6).Text, "Sun", _
                     "the last column of the calendar is Sunday"

    ' 21 September 2026 is the Monday of the fourth row of the grid
    gridRow = firstRow + 3
    gridColumn = DASH_MONTH_FIRST_COL
    AssertEqualsText ws.Cells(gridRow, gridColumn).Text, _
                     modText.TextFormat(modText.KEY_DASH_MONTH_DAY, "21", "1", "25.00"), _
                     "the cell of the day counts the appointments and the revenue"
    AssertEqualsDouble modData.RevenueOn("2026-09-21"), 25, 0.001, "the revenue of that day"
    AssertEqualsText ws.Cells(gridRow, gridColumn + 1).Text, _
                     modText.TextFormat(modText.KEY_DASH_MONTH_DAY, "22", "1", "60.00"), _
                     "Tuesday 22 September shows its own numbers"
    AssertEqualsText ws.Cells(gridRow, gridColumn + 3).Text, _
                     modText.TextFormat(modText.KEY_DASH_MONTH_DAY, "24", "1", "30.00"), _
                     "Thursday 24 September shows its own numbers"

    ' the padding cell in front of the month belongs to August and is muted
    AssertEqualsText ws.Cells(firstRow, gridColumn).Text, _
                     modText.TextFormat(modText.KEY_DASH_MONTH_DAY, "31", "0", "0.00"), _
                     "the first cell of the grid is 31 August"
    AssertEqualsDouble ws.Cells(firstRow, gridColumn).Font.Color, CDbl(COLOR_MUTED), 0, _
                       "a day of another month is muted"

    ' The calendar of the current month marks today.
    todayText = modDashboard.FormatDateText(Date)
    ShowDashboard VIEW_MONTH, todayText
    todayIndex = modPickers.GridIndexOf(Year(Date), Month(Date), todayText)
    AssertTrue (todayIndex > 0), "today sits in the grid of its month"
    gridRow = firstRow + (todayIndex - 1) \ DAYS_PER_WEEK
    gridColumn = DASH_MONTH_FIRST_COL + ((todayIndex - 1) Mod DAYS_PER_WEEK)
    AssertEqualsDouble ws.Cells(gridRow, gridColumn).Interior.Color, CDbl(COLOR_TODAY), 0, _
                       "the cell of today is marked"
    AssertEqualsText ws.Cells(gridRow, gridColumn).Text, _
                     modText.TextFormat(modText.KEY_DASH_MONTH_DAY, CStr(Day(Date)), _
                         CStr(modData.AppointmentsOn(todayText)), _
                         modDashboard.FormatMoney(modData.RevenueOn(todayText))), _
                     "the cell of today counts its own appointments"

    ' A click on a day of the month hands the day over to the daily view.
    ShowDashboard VIEW_MONTH, "2026-09-21"
    gridRow = firstRow + 3
    gridColumn = DASH_MONTH_FIRST_COL
    modApp.OnDynamicClick SHEET_DASHBOARD, ws.Cells(gridRow, gridColumn)
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_DAY, "the click switched to daily"
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), "2026-09-21", "the click chose the day"
    AssertEqualsText ws.Cells(firstRow, DASH_COL_START).Text, "09:00", _
                     "the daily view of the day shows the sample appointment"

    ' The muted cells of the neighbouring month work as well.
    ShowDashboard VIEW_MONTH, "2026-09-21"
    modApp.OnDynamicClick SHEET_DASHBOARD, ws.Cells(firstRow, DASH_MONTH_FIRST_COL)
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), "2026-08-31", _
                     "the cell in front of the month opens 31 August"
End Sub

Private Sub Test_DashboardOpenDay()
    Dim data As Object
    Dim result As Object
    Dim info As Object

    EnsureConfigDefaults
    modUi.SetInteractiveUi False

    AssertEqualsText modRoutes.RouteFor(modRoutes.ACTION_DASHBOARD_OPEN_DAY), _
                     "modDashboard.OpenDay", "dashboard.open_day is routed"

    ' a date that cannot be read falls back to today
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_DATE) = "not a date"
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_OPEN_DAY, data)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "open_day answers a broken date"
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), modDashboard.FormatDateText(Date), _
                     "the broken date falls back to today"
    AssertEqualsText modNames.ReadName(NAME_STATE_VIEW), VIEW_DAY, "the view is daily"
    Set info = result(modRoutes.RESULT_DATA)
    AssertEqualsText info(modRoutes.KEY_VIEW), VIEW_DAY, "the answer carries the view"

    ' no date at all: today as well
    Set data = CreateObject("Scripting.Dictionary")
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_OPEN_DAY, data)
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), modDashboard.FormatDateText(Date), _
                     "an empty payload shows today"

    ' the day of the sample data
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_DATE) = "2026-09-21"
    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_OPEN_DAY, data)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the sample day is accepted"
    AssertEqualsText modNames.ReadName(NAME_STATE_DATE), "2026-09-21", "s_date = the day"
    AssertEqualsText result(modRoutes.RESULT_MESSAGE), _
                     modText.TextFormat(modText.KEY_STATUS_VIEW, _
                         modDashboard.ViewText(VIEW_DAY), _
                         modDashboard.PeriodTitle(VIEW_DAY, "2026-09-21")), _
                     "the message names the day"
End Sub

' One full draw of each view, with the time it needs. The concept asks for less
' than 0.3 s per view; the grids are drawn at the resolution of Config, so the
' weekly one (120 rows x 7 days of links) is reported as well as the others.
Private Sub Test_DashboardRenderTimes()
    Dim result As Object
    Dim started As Single

    EnsureConfigDefaults
    modUi.SetInteractiveUi False

    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_VIEW, _
                                  modRoutes.PayloadForCell(NAME_BTN_VIEW_DAILY))
    started = Timer
    modDashboard.Render
    Note "daily render: " & Format$(Timer - started, "0.000") & " s"
    AssertTrue (Timer - started < 0.3), "the daily render stays under 0.3 s"

    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_VIEW, _
                                  modRoutes.PayloadForCell(NAME_BTN_VIEW_WEEKLY))
    started = Timer
    modDashboard.Render
    Note "weekly render: " & Format$(Timer - started, "0.000") & " s"
    AssertTrue (Timer - started < 0.3), "the weekly render stays under 0.3 s"

    Set result = modApp.AppRouter(modRoutes.ACTION_DASHBOARD_VIEW, _
                                  modRoutes.PayloadForCell(NAME_BTN_VIEW_MONTHLY))
    started = Timer
    modDashboard.Render
    Note "monthly render: " & Format$(Timer - started, "0.000") & " s"
    AssertTrue (Timer - started < 0.3), "the monthly render stays under 0.3 s"

    AssertTrue (result(modRoutes.RESULT_OK) = True), "the three views drew themselves"
End Sub

' The resolution of Config (5 / 15 / 30 / 60 minutes, ??6.9) is what the dashboard
' draws: the daily and the weekly grid have one row per slot, and a start has to
' sit on the step, so the minute list of the start picker follows it as well.
Private Sub Test_DashboardResolution()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim idValue As Long
    Dim before As Long
    Dim picker As frmTimePicker

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    modBackup.SetEnabled False
    before = modData.AppointmentsRowCount()

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    firstRow = modNames.RangeOfName(NAME_CAL_FIRST).Row

    idValue = AddAppointmentForTest("2026-09-25", "15:00", "45", "1", "3")
    AssertTrue (idValue > 0), "the appointment of the resolution test exists"

    ShowDashboard VIEW_DAY, "2026-09-25"
    AssertEqualsLong modDashboard.SlotRowCount(), 120, "the default resolution is 5 minutes"

    ' 30 minutes: the day has 20 slots of 30 minutes
    modNames.WriteName NAME_CONFIG_STEP, "30"
    modDashboard.Render
    AssertEqualsLong modDashboard.SlotRowCount(), 20, "a 30 minute resolution makes 20 rows"
    AssertEqualsText ws.Cells(firstRow, DASH_COL_START).Text, "09:00", "the first slot"
    AssertEqualsText ws.Cells(firstRow + 1, DASH_COL_START).Text, "09:30", "the second slot"
    AssertEqualsText ws.Cells(firstRow + 12, DASH_COL_CLIENT).Text, "Ivan Petrov", _
                     "15:00 is the 13th slot of the day"
    AssertEqualsText ws.Cells(firstRow + 13, DASH_COL_START).Text, "15:30", _
                     "the slot after it starts at 15:30"
    AssertEqualsText ws.Cells(firstRow + 13, DASH_COL_CLIENT).Text, _
                     modText.TextFormat(modText.KEY_DASH_TAIL, CStr(idValue)), _
                     "45 minutes reach into the next slot"

    ' the weekly grid is drawn at the same resolution; Friday 25.09.2026 sits in
    ' the week of Monday 21.09.2026 (column B + 4), the 13th slot of 30 minutes
    ShowDashboard VIEW_WEEK, "2026-09-21"
    AssertEqualsLong modDashboard.SlotRowCount(), 20, "the weekly grid follows the resolution"
    AssertEqualsText ws.Cells(firstRow + 12, DASH_WEEK_FIRST_DAY_COL + 4).Text, _
                     modText.TextFormat(modText.KEY_DASH_SLOT, _
                         modText.TextFormat(modText.KEY_DASH_RANGE, "15:00", "15:45"), _
                         modText.TextFormat(modText.KEY_DASH_WHO, "Ivan Petrov", "Manicure"), _
                         CStr(idValue)), _
                     "the appointment sits in the 13th slot of the column"

    ' the minute list of the start picker holds the multiples of the step; the
    ' duration keeps its 5 minutes (a duration is free)
    Set picker = New frmTimePicker
    picker.InitialiseWithTime FIELD_START, "09:00", 30, vbNullString
    AssertEqualsLong picker.MinuteOptionCount(), 2, "30 minutes leave two minutes to start on"
    AssertEqualsText picker.MinuteOptionAt(1), "30", "the second of them is 30"
    AssertEqualsLong picker.Minutes(), 0, "the field keeps its own minute"

    picker.InitialiseWithTime FIELD_DURATION, "30", 0, vbNullString
    AssertEqualsLong picker.MinuteOptionCount(), 12, "a duration may be any 5 minutes"
    Unload picker

    ' a start that does not sit on the step cannot be saved
    AssertFalse modAppointments.IsOnStep(555, modAppointments.StepMinutes()), _
                "555 minutes do not sit on a 30 minute step"

    ' back to the default: the day has its 120 slots again
    EnsureConfigDefaults
    ShowDashboard VIEW_DAY, "2026-09-25"
    AssertEqualsLong modDashboard.SlotRowCount(), 120, "the default resolution is back"
    AssertEqualsLong modAppointments.StepMinutes(), 5, "the step of Config is 5 again"

    AssertTrue DeleteAppointmentForTest(idValue), "the resolution test removed its appointment"
    modBackup.SetEnabled False
    AssertEqualsLong modData.AppointmentsRowCount(), before, _
                     "the resolution test left no appointment behind"
End Sub

' The probe of the first version of the weekly grid (which hyperlink calls cost
' what) was run once: 280 cells with a hyperlink each were 1.0 s, and the fixed
' version - no fill, no number format and no repeated Config read per cell, plus
' one bulk paint for the columns that cannot be booked - is 0.39 s on the same
' machine. The numbers of Test_DashboardRenderTimes are the record.
' The values of a client as a router payload.
Private Function ClientValues(ByVal nameText As String, ByVal phoneText As String, _
                              ByVal emailText As String, ByVal commentText As String) As Object
    Dim values As Object

    Set values = CreateObject("Scripting.Dictionary")
    values(COL_CLIENT_NAME) = nameText
    values(COL_CLIENT_PHONE) = phoneText
    values(COL_CLIENT_EMAIL) = emailText
    values(COL_CLIENT_COMMENT) = commentText

    Set ClientValues = values
End Function

' The values of a service as a router payload.
Private Function ServiceValues(ByVal nameText As String, ByVal priceText As String, _
                               ByVal durationText As String) As Object
    Dim values As Object

    Set values = CreateObject("Scripting.Dictionary")
    values(COL_SERVICE_NAME) = nameText
    values(COL_SERVICE_PRICE) = priceText
    values(COL_SERVICE_DURATION) = durationText

    Set ServiceValues = values
End Function

' Adds a client through the router and returns its id (0 when it failed).
Private Function AddClientForTest(ByVal nameText As String, ByVal phoneText As String, _
                                  ByVal emailText As String, ByVal commentText As String) As Long
    Dim data As Object
    Dim result As Object
    Dim info As Object

    Set data = CreateObject("Scripting.Dictionary")
    data(COL_CLIENT_NAME) = nameText
    data(COL_CLIENT_PHONE) = phoneText
    data(COL_CLIENT_EMAIL) = emailText
    data(COL_CLIENT_COMMENT) = commentText

    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_ADD, data)
    If result(modRoutes.RESULT_OK) = True Then
        Set info = result(modRoutes.RESULT_DATA)
        AddClientForTest = CLng(Val(CStr(info(COL_ID))))
    End If
End Function

Private Function DeleteClientForTest(ByVal idValue As Long) As Boolean
    Dim data As Object
    Dim result As Object

    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)
    Set result = modApp.AppRouter(modRoutes.ACTION_CLIENTS_DELETE, data)
    modApp.ConfirmReset

    DeleteClientForTest = (result(modRoutes.RESULT_OK) = True)
End Function

Private Function AddServiceForTest(ByVal nameText As String, ByVal priceText As String, _
                                   ByVal durationText As String) As Long
    Dim data As Object
    Dim result As Object
    Dim info As Object

    Set data = CreateObject("Scripting.Dictionary")
    data(COL_SERVICE_NAME) = nameText
    data(COL_SERVICE_PRICE) = priceText
    data(COL_SERVICE_DURATION) = durationText

    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_ADD, data)
    If result(modRoutes.RESULT_OK) = True Then
        Set info = result(modRoutes.RESULT_DATA)
        AddServiceForTest = CLng(Val(CStr(info(COL_ID))))
    End If
End Function

Private Function DeleteServiceForTest(ByVal idValue As Long) As Boolean
    Dim data As Object
    Dim result As Object

    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_ID) = CStr(idValue)
    Set result = modApp.AppRouter(modRoutes.ACTION_SERVICES_DELETE, data)
    modApp.ConfirmReset

    DeleteServiceForTest = (result(modRoutes.RESULT_OK) = True)
End Function

' The hyperlink of a cell button (matched by the name it points at, which is
' what modApp.OnHyperlink reads).
Private Function HyperlinkOf(ByRef ws As Worksheet, ByVal cellName As String) As Object
    Dim index As Long
    Dim item As Object
    Dim hit As Object

    Set hit = Nothing
    On Error Resume Next
    For index = 1 To ws.Hyperlinks.Count
        Set item = ws.Hyperlinks(index)
        If StrComp(CStr(item.SubAddress), cellName, vbTextCompare) = 0 Then
            Set hit = item
            Exit For
        End If
    Next index
    On Error GoTo 0

    Set HyperlinkOf = hit
End Function

' The log anchor: a declared name is not enough, it must resolve. A name that
' does not is reported as a note with the full message of the lookup.
Private Sub CheckLogAnchor()
    Dim anchor As Range
    Dim problem As String

    Set anchor = Nothing
    On Error Resume Next
    Set anchor = modNames.RangeOfName(NAME_LOG_FIRST)
    problem = Err.Description
    On Error GoTo 0
    Err.Clear

    If anchor Is Nothing Then
        Note "log_first is declared but does not resolve: " & problem
    Else
        AssertEqualsLong anchor.Row, DB_FIRST_ROW, "log_first is the first data row of the log"
    End If
End Sub

' =====================================================================
'  phase 5: the reports
' =====================================================================

' The report through the router, like the user: the two cells carry the period,
' Show writes the report, the arrows move it by whole months.
Private Function ShowReport(ByVal fromText As String, ByVal toText As String, _
                            ByVal stepValue As Long) As Object
    Dim data As Object

    Set data = CreateObject("Scripting.Dictionary")
    If Len(fromText) > 0 Then data(modRoutes.KEY_FROM) = fromText
    If Len(toText) > 0 Then data(modRoutes.KEY_TO) = toText
    If stepValue <> 0 Then data(modRoutes.KEY_STEP) = CStr(stepValue)

    Set ShowReport = modApp.AppRouter(modRoutes.ACTION_REPORTS_RENDER, data)
End Function

' The page and its numbers: the period lands in the cells (an empty pair means the
' current month), the title names it, the grid holds the days, the services, the
' clients and the total - and every number agrees with the data of the workbook.
' The exact rows are compared with the model by Test_CrossCheck; here the test
' stays independent of the sample data.
Private Sub Test_Reports()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim rowIndex As Long
    Dim groupText As String
    Dim isoText As String
    Dim totalsRow As Long
    Dim result As Object
    Dim countValue As Long

    modNames.WriteName NAME_REPORT_FROM, vbNullString
    modNames.WriteName NAME_REPORT_TO, vbNullString
    Set result = ShowReport(vbNullString, vbNullString, 0)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "an empty period renders as the current month"
    AssertEqualsText modNames.ReadName(NAME_REPORT_FROM), _
                     modPickers.DateFormatText(modReports.MonthFirstIso()), _
                     "Show fills an empty period with the first day of the month"
    AssertEqualsText modNames.ReadName(NAME_REPORT_TO), _
                     modPickers.DateFormatText(modReports.MonthLastIso()), _
                     "and with its last day"

    Set result = ShowReport("01.09.2026", "30.09.2026", 0)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the report of September renders"
    AssertEqualsText modNames.ReadName(NAME_REPORT_TITLE), "Report 01.09.2026 - 30.09.2026", _
                     "the title names the period"

    Set ws = ThisWorkbook.Worksheets(SHEET_REPORTS)
    firstRow = modNames.RangeOfName(NAME_GRID_REPORTS_FIRST).Row
    AssertEqualsText CStr(ws.Cells(firstRow, REPORT_COL_GROUP).Value), _
                     modText.Text(modText.KEY_REPORT_DAYS), "the grid starts with the days"

    ' Every day row of the grid (the loop stops at the first row that is not a day:
    ' the empty separator, the next heading or the end of the grid).
    For rowIndex = firstRow + 1 To modNames.RangeOfName(NAME_GRID_REPORTS_LAST).Row
        groupText = CStr(ws.Cells(rowIndex, REPORT_COL_GROUP).Value)
        isoText = modPickers.DateTextToIso(groupText)
        If Len(groupText) = 0 Then Exit For
        If Len(isoText) = 0 Then Exit For

        countValue = modData.AppointmentsOn(isoText)
        AssertEqualsDouble ws.Cells(rowIndex, REPORT_COL_COUNT).Value, CDbl(countValue), 0.001, _
                           "the count of " & groupText
        AssertEqualsDouble ws.Cells(rowIndex, REPORT_COL_REVENUE).Value, _
                           modData.RevenueOn(isoText), 0.001, "the revenue of " & groupText
    Next rowIndex

    ' The total row: the count and the money of the whole period.
    totalsRow = 0
    For rowIndex = firstRow To modNames.RangeOfName(NAME_GRID_REPORTS_LAST).Row
        If StrComp(CStr(ws.Cells(rowIndex, REPORT_COL_GROUP).Value), _
                   modText.Text(modText.KEY_REPORT_TOTAL), vbTextCompare) = 0 Then
            totalsRow = rowIndex
        End If
    Next rowIndex
    AssertTrue (totalsRow > 0), "the report ends with a total"
    AssertEqualsDouble ws.Cells(totalsRow, REPORT_COL_COUNT).Value, _
                       CDbl(modData.AppointmentsInPeriod("2026-09-01", "2026-09-30")), 0.001, _
                       "the total counts the period"
    AssertEqualsDouble ws.Cells(totalsRow, REPORT_COL_REVENUE).Value, _
                       modData.RevenueInPeriod("2026-09-01", "2026-09-30"), 0.001, _
                       "and it sums its money"

    ' The two arrows move the whole window by a month: September (01.09 - 30.09)
    ' becomes August (01.08 - 30.08) and back.
    Set result = ShowReport(vbNullString, vbNullString, -1)
    AssertEqualsText modNames.ReadName(NAME_REPORT_FROM), "01.08.2026", "a month back"
    AssertEqualsText modNames.ReadName(NAME_REPORT_TO), "30.08.2026", _
                     "the same window a month back"
    Set result = ShowReport(vbNullString, vbNullString, 1)
    AssertEqualsText modNames.ReadName(NAME_REPORT_FROM), "01.09.2026", "and a month forward again"

    ' A period that is not a period: the reason in the status, nothing rendered.
    Set result = ShowReport("01.10.2026", "30.09.2026", 0)
    AssertEqualsText result(modRoutes.RESULT_MESSAGE), _
                     modText.TextFormat(modText.KEY_ERROR_PERIOD, "01.10.2026", "30.09.2026"), _
                     "the ends the wrong way round are refused"
    Set result = ShowReport("nonsense", "30.09.2026", 0)
    AssertEqualsText result(modRoutes.RESULT_MESSAGE), _
                     modText.TextFormat(modText.KEY_ERROR_REQUIRED, "period start"), _
                     "a start that is not a date is refused"
End Sub

' The Pick button of a report period cell: the payload of the button plus the day a
' test feeds instead of the calendar (the picker forms have their own test).
Private Function PickReportDate(ByVal buttonName As String, ByVal isoText As String) As Object
    Dim payload As Object
    Dim data As Object
    Dim key As Variant

    Set payload = modRoutes.PayloadForCell(buttonName)
    Set data = CreateObject("Scripting.Dictionary")
    For Each key In payload.Keys
        data(CStr(key)) = payload(CStr(key))
    Next key
    data(modRoutes.KEY_VALUE) = isoText

    Set PickReportDate = modApp.AppRouter(modRoutes.ACTION_REPORTS_PICK_DATE, data)
End Function

' From and To of the report period have a Pick button and a double click on the
' cell, exactly like the date of Booking (§7.5, §9.2).
Private Sub Test_ReportPickers()
    Dim ws As Worksheet
    Dim result As Object
    Dim cancel As Boolean

    Set ws = ThisWorkbook.Worksheets(SHEET_REPORTS)

    AssertTrue modRoutes.IsClickAction(NAME_BTN_REPORT_FROM_PICK), "b_report_from_pick is wired"
    AssertTrue modRoutes.IsClickAction(NAME_BTN_REPORT_TO_PICK), "b_report_to_pick is wired"
    AssertEqualsText modRoutes.PayloadText(modRoutes.PayloadForCell(NAME_BTN_REPORT_FROM_PICK)), _
                     "cell=" & NAME_REPORT_FROM, "the start button knows its cell"
    AssertEqualsText modRoutes.PayloadText(modRoutes.PayloadForCell(NAME_BTN_REPORT_TO_PICK)), _
                     "cell=" & NAME_REPORT_TO, "the end button knows its cell"

    ' Through the router: the picked day lands in the cell in the European form the
    ' page shows, ready for Show.
    modNames.WriteName NAME_REPORT_FROM, vbNullString
    Set result = PickReportDate(NAME_BTN_REPORT_FROM_PICK, "2026-09-10")
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the start of the period was picked"
    AssertEqualsText modNames.ReadName(NAME_REPORT_FROM), "10.09.2026", "and the cell shows it"

    Set result = PickReportDate(NAME_BTN_REPORT_TO_PICK, "2026-09-20")
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the end of the period was picked"
    AssertEqualsText modNames.ReadName(NAME_REPORT_TO), "20.09.2026", "and that cell too"

    ' Show then renders the period the pickers wrote.
    Set result = ShowReport(vbNullString, vbNullString, 0)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "Show renders the picked period"
    AssertEqualsText modNames.ReadName(NAME_REPORT_TITLE), "Report 10.09.2026 - 20.09.2026", _
                     "and the title names it"

    ' A double click on either cell is the same as its Pick button (the cancel is
    ' what keeps Excel from opening the cell for typing). With the dialogs off the
    ' picker answers the day it started with, like every picker in a test run. The
    ' cells come from their names, so a new layout does not need a change here.
    modNames.WriteName NAME_REPORT_FROM, "10.09.2026"
    cancel = False
    modApp.OnDoubleClick SHEET_REPORTS, modNames.RangeOfName(NAME_REPORT_FROM), cancel
    AssertTrue cancel, "the double click on From is handled as a click"
    AssertEqualsText modNames.ReadName(NAME_REPORT_FROM), "10.09.2026", _
                     "and the calendar keeps the day the cell holds"

    modNames.WriteName NAME_REPORT_TO, "20.09.2026"
    cancel = False
    modApp.OnDoubleClick SHEET_REPORTS, modNames.RangeOfName(NAME_REPORT_TO), cancel
    AssertTrue cancel, "the double click on To is handled as a click"
    AssertEqualsText modNames.ReadName(NAME_REPORT_TO), "20.09.2026", _
                     "and that calendar too"
End Sub

' =====================================================================
'  phase 6: the Config page
' =====================================================================

' The two buttons of the page through the router: Apply checks the settings, stores
' them in their canonical form and draws the grids again; Restore defaults puts the
' built in values back. The page is read from its cells, so a test drives it like a
' user does.
Private Sub Test_ConfigPage()
    Dim result As Object
    Dim blank As Object
    Dim defaults As Object
    Dim key As Variant

    Set blank = CreateObject("Scripting.Dictionary")

    AssertTrue modRoutes.IsClickAction(NAME_BTN_CONFIG_APPLY), "b_config_apply is wired"
    AssertTrue modRoutes.IsClickAction(NAME_BTN_CONFIG_DEFAULTS), "b_config_defaults is wired"

    ' a fresh workbook opens with the built in values of §16.1: the template writes
    ' them, and this guard fails the day the two lists drift apart
    Set defaults = modUiConfig.DefaultValues()
    For Each key In defaults.Keys
        AssertEqualsText modNames.ReadName(CStr(key)), CStr(defaults(key)), _
                         "a fresh page holds the built in value of '" & CStr(key) & "'"
    Next key

    ' a page the user typed: the times become canonical, the dates ISO (that is what
    ' the working-day rule compares, whatever the cell shows)
    modNames.WriteName NAME_CONFIG_DAY_START, "8:30"
    modNames.WriteName NAME_CONFIG_DAY_END, "18:00"
    modNames.WriteName NAME_CONFIG_STEP, "015"
    modNames.WriteName NAME_CONFIG_DURATION, "20"
    modNames.WriteName NAME_CONFIG_CURRENCY, "eur"
    modNames.WriteName NAME_CONFIG_NUMBER_FORMAT, "#,##0.00"
    modNames.WriteName NAME_CONFIG_NONWORK_DAYS, " 6 , 7 "
    modNames.WriteName NAME_CONFIG_NONWORK_DATES, "25.09.2026, 2026-10-05"

    Set result = modApp.AppRouter(modRoutes.ACTION_CONFIG_APPLY, blank)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "Apply stores a page that is fine"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_DAY_START), "08:30", _
                     "the opening time is canonical"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_STEP), "15", _
                     "the step lost its leading zero"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_CURRENCY), "EUR", "the currency went up"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_NONWORK_DAYS), "6,7", _
                     "the weekdays are trimmed"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_NONWORK_DATES), "2026-09-25,2026-10-05", _
                     "and the dates are ISO"

    ' the rest of the application follows the page at once
    AssertFalse modAppointments.IsWorkingDay("2026-10-05"), "the new day off is not a working day"
    AssertEqualsLong modAppointments.StepMinutes(), 15, "the grids follow the step"
    AssertEqualsLong modAppointments.DayStartMinutes(), 510, "and the opening time"

    ' a page with a problem is refused and nothing is stored
    modNames.WriteName NAME_CONFIG_DAY_START, "19:00"
    Set result = modApp.AppRouter(modRoutes.ACTION_CONFIG_APPLY, blank)
    AssertTrue (result(modRoutes.RESULT_OK) <> True), "a day that opens when it closes is refused"
    AssertEqualsText result(modRoutes.RESULT_MESSAGE), _
                     modText.TextFormat(modText.KEY_ERROR_DAY_ORDER, "19:00", "18:00"), _
                     "and the message names the two times"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_STEP), "15", "nothing was stored"

    ' back to the built in values of §16.1
    Set result = modApp.AppRouter(modRoutes.ACTION_CONFIG_DEFAULTS, blank)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "Restore defaults works"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_DAY_START), "09:00", "the opening time is back"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_DAY_END), "19:00", "so is the closing time"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_STEP), "5", "and the step"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_DURATION), "10", "and the duration"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_CURRENCY), "EUR", "and the currency"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_NONWORK_DATES), vbNullString, _
                     "the non working dates are cleared"
    AssertEqualsLong modAppointments.StepMinutes(), 5, "and the grids follow again"

    ' an empty list of non working weekdays means none: the owner works on Saturday
    ' (a fallback to "6,7" here would quietly ignore the setting)
    modNames.WriteName NAME_CONFIG_NONWORK_DAYS, vbNullString
    modNames.WriteName NAME_CONFIG_NONWORK_DATES, vbNullString
    AssertTrue modAppointments.IsWorkingDay("2026-09-26"), _
               "with no non working weekday a Saturday is a working day"
    AssertEqualsText modAppointments.RuleCode( _
                         AppointmentValues("2026-09-26", "15:00", "45", "1", "1")), _
                     vbNullString, "and a booking on that Saturday is accepted"
    EnsureConfigDefaults
End Sub

' =====================================================================
'  phase 6: the non working dates (the calendar and the Clear button)
' =====================================================================

' The two buttons beside i_config_nonwork_dates: the calendar adds one day to the
' list (it does not replace it - the days pile up) and Clear empties it. Both touch
' the cell only, so the user presses Apply settings, exactly like with the pickers
' of Reports (§16.1).
Private Sub Test_NonworkDates()
    Dim result As Object
    Dim data As Object

    AssertTrue modRoutes.IsClickAction(NAME_BTN_CONFIG_NONWORK_PICK), _
                "b_config_nonwork_pick is wired"
    AssertTrue modRoutes.IsClickAction(NAME_BTN_CONFIG_NONWORK_CLEAR), _
                "b_config_nonwork_clear is wired"

    modNames.WriteName NAME_CONFIG_NONWORK_DATES, "2026-09-25"

    ' the calendar answers with a day and it is added; the old one stays
    Set data = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_CELL) = NAME_CONFIG_NONWORK_DATES
    data(modRoutes.KEY_VALUE) = "20.09.2026"

    Set result = modApp.AppRouter(modRoutes.ACTION_CONFIG_PICK_NONWORK, data)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the calendar added a day"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_NONWORK_DATES), _
                     "2026-09-20,2026-09-25", "the list is in order and nothing is lost"
    AssertTrue (InStr(1, CStr(result(modRoutes.RESULT_MESSAGE)), "20.09.2026", _
                      vbTextCompare) > 0), "the message names the day"

    ' the same day a second time changes nothing
    Set result = modApp.AppRouter(modRoutes.ACTION_CONFIG_PICK_NONWORK, data)
    AssertEqualsText modNames.ReadName(NAME_CONFIG_NONWORK_DATES), _
                     "2026-09-20,2026-09-25", "a day that is already in the list is not repeated"

    ' Clear drops the whole list (and Apply would store the empty cell)
    Set result = modApp.AppRouter(modRoutes.ACTION_CONFIG_CLEAR_NONWORK)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "Clear emptied the list"
    AssertEqualsText modNames.ReadName(NAME_CONFIG_NONWORK_DATES), vbNullString, _
                     "the cell of the non working dates is empty"

    ' and the page is still fine after all that: the rule accepts it
    AssertEqualsText modUiConfig.RuleCode(modUiConfig.Read()), vbNullString, _
                     "the settings are still accepted"
End Sub

' =====================================================================
'  phase 6: the data sheets out of sight
' =====================================================================

' The two buttons of the Config page hide and show the three db_* sheets. Nothing
' is protected, so the application must not care: the test loads the model again
' while they are hidden (§6.5 - the owner chose hiding over protection).
Private Sub Test_DataSheetVisibility()
    Dim result As Object

    AssertTrue modRoutes.IsClickAction(NAME_BTN_LOCK_DATA), "b_lock_data is wired"
    AssertTrue modRoutes.IsClickAction(NAME_BTN_UNLOCK_DATA), "b_unlock_data is wired"

    Set result = modApp.AppRouter(modRoutes.ACTION_CONFIG_LOCK_DATA)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the data sheets are hidden"
    AssertTrue modProtection.DataSheetsHidden(), "all three db_ sheets are out of sight"
    AssertTrue (Len(modNames.ReadName(NAME_STATUS_CONFIG)) > 0), _
                "the status of Config says what happened"

    ' the data is still there: the tables are loaded again while they are hidden
    modData.UnloadData
    modData.LoadData
    AssertTrue (modData.ClientsRowCount() > 0), "a hidden data sheet is still read"
    AssertTrue (modData.AppointmentsRowCount() > 0), "and so is the appointments table"

    Set result = modApp.AppRouter(modRoutes.ACTION_CONFIG_UNLOCK_DATA)
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the data sheets are shown again"
    AssertFalse modProtection.DataSheetsHidden(), "the sheets are visible for a manual edit"
End Sub

' =====================================================================
'  phase 6: the version of the schema
' =====================================================================

' A file carries both the code and the data, so an old file has to keep working:
' the version cell says what the file is, an older number is migrated (a copy and a
' Log row come first) and a newer one is refused with a sentence (§8.3).
Private Sub Test_Schema()
    Dim ws As Worksheet
    Dim originalColumn As Long
    Dim addedColumn As Long

    modBackup.SetEnabled False

    ' the template of this build writes the number of the code
    AssertEqualsLong modSchema.FileVersion(), modSchema.CURRENT_SCHEMA_VERSION, _
                     "the file carries the schema version of this build"

    ' an old file: a column of a table is gone and the number says "1"
    Set ws = ThisWorkbook.Worksheets(SHEET_DB_SERVICES)
    originalColumn = ColumnOfHeader(ws, COL_SERVICE_DURATION)
    AssertTrue (originalColumn > 0), "the services table carries its duration column"
    ws.Cells(1, originalColumn).Value = vbNullString
    modNames.WriteName NAME_STATE_SCHEMA, "1"

    AssertEqualsText modSchema.Problem(), vbNullString, "an older file is migrated"
    addedColumn = ColumnOfHeader(ws, COL_SERVICE_DURATION)
    AssertTrue (addedColumn > 0), "the missing column is there again"
    AssertEqualsLong modSchema.FileVersion(), modSchema.CURRENT_SCHEMA_VERSION, _
                     "and the file carries the number of this build now"

    ' the header goes back where it was, so the tables stay as the other tests (and
    ' the fixture of the cross check) expect them
    ws.Cells(1, addedColumn).Value = vbNullString
    ws.Cells(1, originalColumn).Value = COL_SERVICE_DURATION
    AssertEqualsLong ColumnOfHeader(ws, COL_SERVICE_DURATION), originalColumn, _
                     "the header is where it was"

    ' a file of a newer build is refused, and nothing is written
    modNames.WriteName NAME_STATE_SCHEMA, CStr(modSchema.CURRENT_SCHEMA_VERSION + 7)
    AssertTrue (Len(modSchema.Problem()) > 0), "a newer file is refused with a sentence"
    modNames.WriteName NAME_STATE_SCHEMA, CStr(modSchema.CURRENT_SCHEMA_VERSION)
    AssertEqualsLong modSchema.FileVersion(), modSchema.CURRENT_SCHEMA_VERSION, _
                     "the version cell is back"
End Sub

' The column of a header text in the first row of a sheet (0 when it is not there).
Private Function ColumnOfHeader(ByRef ws As Worksheet, ByVal headerText As String) As Long
    Dim lastColumn As Long
    Dim index As Long

    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For index = 1 To lastColumn
        If StrComp(Trim$(CStr(ws.Cells(1, index).Value)), headerText, vbTextCompare) = 0 Then
            ColumnOfHeader = index
            Exit Function
        End If
    Next index
End Function

' =====================================================================
'  phase 6: performance - a month of a busy business
' =====================================================================

' The application has to stay usable when the table grows. This writes a thousand
' appointments straight into db_appointments (one array write, the way a paste
' arrives), loads the model from cold and measures what a user does: the load, the
' report of the month and the two grids. The numbers land in the trace
' (build\xls-test-trace.txt) and each one has a ceiling, so a regression fails the
' build instead of being felt months later.
Private Sub Test_Performance()
    Dim result As Object
    Dim started As Single
    Dim rowCount As Long

    modBackup.SetEnabled False
    modUi.SetInteractiveUi False

    rowCount = WritePerformanceAppointments()
    modData.UnloadData

    started = Timer
    modData.LoadData
    Note "perf: load " & CStr(rowCount) & " appointments: " & Format$(Timer - started, "0.000") & " s"
    AssertEqualsLong modData.AppointmentsRowCount(), rowCount, "the whole table is loaded"
    AssertTrue (Timer - started < 10), "loading " & CStr(rowCount) & " appointments stays under 10 s"

    started = Timer
    Set result = ShowReport("01.09.2026", "30.09.2026", 0)
    Note "perf: report of the month: " & Format$(Timer - started, "0.000") & " s"
    AssertTrue (result(modRoutes.RESULT_OK) = True), "the report of the full month is drawn"
    AssertTrue (Timer - started < 5), "the report over " & CStr(rowCount) & " rows stays under 5 s"

    ShowDashboard VIEW_DAY, "2026-09-14"
    started = Timer
    modDashboard.Render
    Note "perf: daily render of a busy day: " & Format$(Timer - started, "0.000") & " s"
    AssertTrue (Timer - started < 1), "the daily render of a busy day stays under 1 s"

    ShowDashboard VIEW_WEEK, "2026-09-14"
    started = Timer
    modDashboard.Render
    Note "perf: weekly render of a busy week: " & Format$(Timer - started, "0.000") & " s"
    AssertTrue (Timer - started < 2), "the weekly render of a busy week stays under 2 s"
End Sub

' A thousand appointments of September 2026, written in one go and answered with
' their count. They overlap on purpose: a hand edited table does that too and the
' renderer has to show it (§6.9), so this is the hardest month, not a friendly one.
Private Function WritePerformanceAppointments() As Long
    Const perfRows As Long = 1000
    Dim ws As Worksheet
    Dim rows() As Variant
    Dim rowIndex As Long
    Dim dayOffset As Long
    Dim appointmentIndex As Long
    Dim monthStart As Date

    Set ws = ThisWorkbook.Worksheets(SHEET_DB_APPOINTMENTS)
    monthStart = DateSerial(2026, 9, 1)

    modUi.BeginBulk
    ws.Range(ws.Cells(DB_FIRST_ROW, 1), _
             ws.Cells(DB_FIRST_ROW + perfRows + 10, 6)).ClearContents

    ReDim rows(1 To perfRows, 1 To 6)
    For rowIndex = 1 To perfRows
        dayOffset = (rowIndex - 1) Mod 30                      ' the days of the month
        appointmentIndex = (rowIndex - 1) \ 30                 ' ~33 appointments a day
        rows(rowIndex, 1) = CStr(9000 + rowIndex)              ' id
        rows(rowIndex, 2) = CStr(1 + (rowIndex Mod 2))         ' client_id
        rows(rowIndex, 3) = CStr(1 + (rowIndex Mod 3))         ' service_id
        rows(rowIndex, 4) = Format$(DateAdd("d", dayOffset, monthStart), "yyyy-mm-dd")
        rows(rowIndex, 5) = modApp.MinutesToTime(540 + (appointmentIndex * 15))
        rows(rowIndex, 6) = "30"
    Next rowIndex

    ws.Range(ws.Cells(DB_FIRST_ROW, 1), _
             ws.Cells(DB_FIRST_ROW + perfRows - 1, 6)).Value = rows
    modUi.EndBulk

    WritePerformanceAppointments = perfRows
End Function

' =====================================================================
'  the cross check: the workbook against the script model (tools\sim)
' =====================================================================

' The fixtures AND the cases live in one file, src\tests\crosscheck.txt, which the
' model under cscript reads as well (sim-data.vbs, sim-crosscheck.vbs). This test
' puts those fixtures into the workbook, answers every case with the code of the
' workbook and writes the answers to build\xls-crosscheck.txt; tools\build.vbs
' compares that file with build\sim-crosscheck.txt line by line. So a rule that
' drifts in one of the two implementations - another code, another slot plan,
' another sum - stops the build instead of being found later.
' It runs last: it replaces the sample data of the workbook with the fixtures.
Private Sub Test_CrossCheck()
    Dim data As Object
    Dim tables As Object
    Dim lines As Variant
    Dim index As Long
    Dim answers As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(CrosscheckPath()) Then
        Note "the cross check is skipped: " & CrosscheckPath() & " is missing"
        Exit Sub
    End If

    Set data = CrosscheckFixtures()
    LoadCrosscheckFixtures data

    Set tables = data("tables")
    CheckFixtureHeader SHEET_DB_CLIENTS, tables
    CheckFixtureHeader SHEET_DB_SERVICES, tables
    CheckFixtureHeader SHEET_DB_APPOINTMENTS, tables

    lines = Split(CStr(data("cases")), vbCrLf)
    For index = 0 To UBound(lines)
        answers = answers & CrosscheckAnswer(CaseKindOf(CStr(lines(index))), _
                                             CaseNameOf(CStr(lines(index))), _
                                             CaseArgumentsOf(CStr(lines(index))))
    Next index

    WriteCrosscheckAnswers answers, UBound(lines) + 1

    ' The geometry of the body as the renderer left it: a plan that comes back empty
    ' is a body that was not drawn where the test reads it, and this note says where.
    Note "cross check body: rows " & CStr(modDashboard.BodyFirstRow()) & ".." & _
         CStr(modDashboard.BodyLastRow()) & ", first slot A=" & _
         CStr(ThisWorkbook.Worksheets(SHEET_DASHBOARD).Cells(modDashboard.BodyFirstRow(), _
              DASH_COL_START).Text) & ", view=" & modNames.ReadName("s_view") & _
         ", date=" & modNames.ReadName("s_date")

    AssertTrue Len(answers) > 0, "the workbook answered the cases of the cross check"
    ' The count has to follow the file: a case the workbook does not read would make
    ' the two cross check files differ, and this names it at once.
    AssertEqualsLong UBound(lines) + 1, 53, "the cases of the fixture file"
    AssertEqualsLong modData.AppointmentsOn("2026-09-21"), 1, "the fixtures are in the workbook"
    AssertEqualsText modData.ClientName(1), "Ivan Petrov", "client 1 of the fixtures"
    AssertEqualsText modData.ServiceName(3), "Manicure", "service 3 of the fixtures"
End Sub

' ---- the fixture file of the cross check ------------------------------

' <root>\src\tests\crosscheck.txt: the workbook of a test build sits in <root>\build.
Private Function CrosscheckPath() As String
    Dim fso As Object
    Dim root As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    root = fso.GetParentFolderName(ThisWorkbook.Path)
    CrosscheckPath = fso.BuildPath(fso.BuildPath(root, "src\tests"), "crosscheck.txt")
End Function

' The file, split the way the model splits it: config("name"), the header of a
' table ("header:db_clients"), the rows of a table ("rows:db_clients", CrLf joined)
' and the cases ("cases", CrLf joined).
Private Function CrosscheckFixtures() As Object
    Dim data As Object
    Dim config As Object
    Dim tables As Object
    Dim fileNumber As Integer
    Dim line As String
    Dim section As String
    Dim casesText As String
    Dim separator As Long

    Set data = CreateObject("Scripting.Dictionary")
    Set config = CreateObject("Scripting.Dictionary")
    Set tables = CreateObject("Scripting.Dictionary")

    fileNumber = FreeFile
    Open CrosscheckPath() For Input As #fileNumber
    Do Until EOF(fileNumber)
        Line Input #fileNumber, line
        line = Trim$(line)
        If Len(line) > 0 Then
            If Left$(line, 1) = "[" Then
                section = Mid$(line, 2, Len(line) - 2)
            ElseIf Left$(line, 1) <> "#" Then
                If section = "config" Then
                    separator = InStr(line, "=")
                    If separator > 0 Then
                        config(Trim$(Left$(line, separator - 1))) = _
                            Trim$(Mid$(line, separator + 1))
                    End If
                ElseIf section = "cases" Then
                    If Len(casesText) > 0 Then casesText = casesText & vbCrLf
                    casesText = casesText & line
                ElseIf Left$(section, 3) = "db_" Then
                    If Not tables.Exists("header:" & section) Then
                        tables("header:" & section) = line
                    Else
                        If Len(CStr(tables("rows:" & section))) > 0 Then _
                            tables("rows:" & section) = tables("rows:" & section) & vbCrLf
                        tables("rows:" & section) = tables("rows:" & section) & line
                    End If
                End If
            End If
        End If
    Loop
    Close #fileNumber

    Set data("config") = config
    Set data("tables") = tables
    data("cases") = casesText

    Set CrosscheckFixtures = data
End Function

' Puts the fixtures into the workbook: the Config cells, the three db_ tables
' (cleared first, so the workbook holds exactly the fixtures) and the cached model
' is loaded again.
Private Sub LoadCrosscheckFixtures(ByRef data As Object)
    Dim config As Object
    Dim tables As Object
    Dim key As Variant

    Set config = data("config")
    For Each key In config.Keys
        modNames.WriteName CStr(key), config(key)
    Next key

    Set tables = data("tables")
    ClearFixtureTable SHEET_DB_APPOINTMENTS, 6
    ClearFixtureTable SHEET_DB_CLIENTS, 5
    ClearFixtureTable SHEET_DB_SERVICES, 4

    WriteFixtureTable SHEET_DB_APPOINTMENTS, tables
    WriteFixtureTable SHEET_DB_CLIENTS, tables
    WriteFixtureTable SHEET_DB_SERVICES, tables

    modData.LoadData
End Sub

' Everything below the header goes, so not one row of the sample data survives.
Private Sub ClearFixtureTable(ByVal sheetName As String, ByVal columnCount As Long)
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(sheetName)
    ' Enough for the tables of the performance test as well: the cross check has to
    ' start from an empty table, whatever the test before it left behind.
    ws.Range(ws.Cells(DB_FIRST_ROW, 1), _
             ws.Cells(DB_FIRST_ROW + 1200, columnCount)).ClearContents
End Sub

' One CSV table of the fixture file: every line is one row, the columns are in the
' order of the header (which CheckFixtureHeader compares with the sheet).
Private Sub WriteFixtureTable(ByVal sheetName As String, ByRef tables As Object)
    Dim rowsText As String
    Dim lines As Variant
    Dim fields As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    If Not tables.Exists("rows:" & sheetName) Then Exit Sub

    rowsText = CStr(tables("rows:" & sheetName))
    lines = Split(rowsText, vbCrLf)
    For rowIndex = 0 To UBound(lines)
        fields = Split(CStr(lines(rowIndex)), ",")
        For columnIndex = 0 To UBound(fields)
            If Len(fields(columnIndex)) > 0 Then
                modUi.WriteCell sheetName, DB_FIRST_ROW + rowIndex, columnIndex + 1, _
                                fields(columnIndex)
            End If
        Next columnIndex
    Next rowIndex
End Sub

' The header line of the fixture file against the header row of the sheet: both
' lanes must look at the same columns in the same order.
Private Sub CheckFixtureHeader(ByVal sheetName As String, ByRef tables As Object)
    Dim expected As String
    Dim actual As String
    Dim fields As Variant
    Dim columnIndex As Long
    Dim ws As Worksheet

    If Not tables.Exists("header:" & sheetName) Then Exit Sub

    expected = CStr(tables("header:" & sheetName))
    fields = Split(expected, ",")
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For columnIndex = 0 To UBound(fields)
        If Len(actual) > 0 Then actual = actual & ","
        actual = actual & Trim$(CStr(ws.Cells(DB_HEADER_ROW, columnIndex + 1).Value))
    Next columnIndex

    AssertEqualsText actual, expected, "the header of " & sheetName & " is the header of the fixtures"
End Sub

' ---- the answers of the workbook --------------------------------------

' "rule|off-step|client_id=1;..." - the three parts of a case line.
Private Function CaseKindOf(ByVal caseLine As String) As String
    CaseKindOf = Split(caseLine, "|")(0)
End Function

Private Function CaseNameOf(ByVal caseLine As String) As String
    Dim parts As Variant

    parts = Split(caseLine, "|")
    If UBound(parts) >= 1 Then CaseNameOf = CStr(parts(1))
End Function

' Everything after the second "|" (the arguments of the case, "" when it has none).
Private Function CaseArgumentsOf(ByVal caseLine As String) As String
    Dim first As Long
    Dim second As Long

    first = InStr(caseLine, "|")
    If first = 0 Then Exit Function
    second = InStr(first + 1, caseLine, "|")
    If second = 0 Then Exit Function

    CaseArgumentsOf = Mid$(caseLine, second + 1)
End Function

' One answer line per case, in the format the model writes (kind|name|value...); a
' plan answers with one line per slot of the rendered grid.
' The answer of one case, in the format kind|name|value... A "plan" answers with one
' line per slot, in the order the grid draws them.
Private Function CrosscheckAnswer(ByVal kind As String, ByVal caseName As String, _
                                  ByVal arguments As String) As String
    Dim values As Object
    Dim parts As Variant
    Dim text As String

    Select Case kind
        Case "rule"
            Set values = CrosscheckPayload(arguments)
            CrosscheckAnswer = "rule|" & caseName & "|" & _
                               modAppointments.RuleCode(values) & vbCrLf

        Case "normalise"
            ' The validator fills the normalised values into the payload (ISO date,
            ' HH:MM, minutes) - the case shows what a writer would store.
            Set values = CrosscheckPayload(arguments)
            modAppointments.RuleCode values
            CrosscheckAnswer = "normalise|" & caseName & "|" & _
                               modApp.PayloadText(values, COL_DATE) & "|" & _
                               modApp.PayloadText(values, COL_START) & "|" & _
                               modApp.PayloadText(values, COL_SERVICE_DURATION) & vbCrLf

        Case "count"
            CrosscheckAnswer = "count|" & caseName & "|" & _
                               CStr(modData.AppointmentsOn(caseName)) & vbCrLf

        Case "revenue"
            ' Money crosses the lane as cents, so the answer does not depend on the
            ' regional settings of the machine that runs the two lanes (section 13).
            CrosscheckAnswer = "revenue|" & caseName & "|" & _
                               CStr(CLng(modData.RevenueOn(caseName) * 100)) & vbCrLf

        Case "freeday"
            CrosscheckAnswer = "freeday|" & caseName & "|" & _
                               CStr(modDashboard.FreeSlotsFor(VIEW_DAY, caseName)) & vbCrLf

        Case "label"
            CrosscheckAnswer = "label|" & caseName & "|" & _
                               modAppointments.AppointmentLabel(CLng(Val(caseName))) & vbCrLf

        Case "busy"
            CrosscheckAnswer = "busy|" & caseName & "|" & _
                               modAppointments.BusyIntervalsOn(caseName, 0) & vbCrLf

        Case "conflict-start"
            parts = Split(arguments, "|")
            CrosscheckAnswer = "conflict-start|" & caseName & "|" & _
                modAppointments.FirstConflictStart(caseName, _
                    modApp.TimeToMinutes(CStr(parts(0))), CLng(Val(CStr(parts(1)))), 0) & vbCrLf

        Case "plan"
            ShowDashboard VIEW_DAY, caseName
            CrosscheckAnswer = CrosscheckPlan(caseName)

        Case "report"
            ' The period of the arguments lands in the two cells, exactly as the page
            ' does it, and the report is rendered; then its grid is read back row by row.
            parts = Split(arguments, "|")
            modNames.WriteName NAME_REPORT_FROM, modPickers.DateFormatText(CStr(parts(0)))
            modNames.WriteName NAME_REPORT_TO, modPickers.DateFormatText(CStr(parts(1)))
            ShowReport vbNullString, vbNullString, 0
            CrosscheckAnswer = CrosscheckReport(caseName)

        Case "config-rule"
            ' The rule code of a page: "" while the settings are fine, otherwise the
            ' first problem (the sentence the user reads is built from it in modText).
            Set values = CrosscheckConfig(arguments)
            CrosscheckAnswer = "config-rule|" & caseName & "|" & _
                               modUiConfig.RuleCode(values) & vbCrLf

        Case "config-values"
            ' What the page stores: the canonical form (HH:MM, whole numbers, the money
            ' code up, the lists trimmed, the dates ISO). Only a page that is fine is
            ' normalised, so every case for this kind carries values the rules accept.
            Set values = CrosscheckConfig(arguments)
            modUiConfig.Normalise values
            CrosscheckAnswer = "config-values|" & caseName & "|" & _
                               ConfigValuesLine(values) & vbCrLf

        Case "config-defaults"
            ' The built in values of section 16.1, in the order of the sheet.
            CrosscheckAnswer = "config-defaults|" & caseName & "|" & _
                               ConfigValuesLine(modUiConfig.DefaultValues()) & vbCrLf

        Case "config-dates"
            ' The Pick button of the non working dates: one day is added to the list
            ' (ISO, in the order of the calendar, without repeats). The list travels in
            ' the arguments, so the case does not depend on what a cell holds.
            parts = Split(arguments, "|")
            text = vbNullString
            If UBound(parts) >= 1 Then
                text = modUiConfig.AppendIsoDate(CStr(parts(1)), CStr(parts(0)))
            End If
            CrosscheckAnswer = "config-dates|" & caseName & "|" & text & vbCrLf

        Case Else
            Note "the cross check met an unknown case kind: " & kind
    End Select
End Function

' "client_id=1;service_id=1;date=25.09.2026;start=9:5;duration=045" - the payload of
' a booking as a case writes it (the keys are the columns of db_appointments).
Private Function CrosscheckPayload(ByVal arguments As String) As Object
    Dim values As Object
    Dim pairs As Variant
    Dim item As String
    Dim separator As Long
    Dim index As Long

    Set values = CreateObject("Scripting.Dictionary")
    If Len(arguments) > 0 Then
        pairs = Split(arguments, ";")
        For index = 0 To UBound(pairs)
            item = CStr(pairs(index))
            separator = InStr(item, "=")
            If separator > 0 Then
                values(Trim$(Left$(item, separator - 1))) = Mid$(item, separator + 1)
            End If
        Next index
    End If

    Set CrosscheckPayload = values
End Function

' The Config page of the fixtures, with the arguments of a case on top of it:
' "i_config_step=15" is the cell the user typed over, the fixture file holds the
' rest. The keys are the names of the cells, the same ones the fixture uses, so a
' case reads like the page it describes. Nothing is written back to the sheet, so
' one case cannot leak into the next.
Private Function CrosscheckConfig(ByVal arguments As String) As Object
    Dim values As Object
    Dim pairs As Variant
    Dim item As String
    Dim separator As Long
    Dim index As Long

    Set values = modUiConfig.Read()
    If Len(arguments) > 0 Then
        pairs = Split(arguments, ";")
        For index = 0 To UBound(pairs)
            item = CStr(pairs(index))
            separator = InStr(item, "=")
            If separator > 0 Then
                values(Trim$(Left$(item, separator - 1))) = Mid$(item, separator + 1)
            End If
        Next index
    End If

    Set CrosscheckConfig = values
End Function

' The eight values of a page in the order of the sheet, joined by "|" - the shape
' the model answers with as well (both read the page in the same order).
Private Function ConfigValuesLine(ByRef values As Object) As String
    Dim keys As Variant
    Dim index As Long
    Dim text As String

    keys = values.Keys
    For index = 0 To UBound(keys)
        If index > 0 Then text = text & "|"
        text = text & CStr(values(CStr(keys(index))))
    Next index

    ConfigValuesLine = text
End Function

' The plan of a day, read from the cells the daily renderer painted: one line per
' slot of the body - "HH:MM|free", "HH:MM|tail|id" or
' "HH:MM|start|id|client|service". The test does not repeat the decision of the
' renderer; it reads what the user sees on the sheet (a booked row carries the id
' in the last column, a continuation row only the mark in the client cell).
Private Function CrosscheckPlan(ByVal isoDate As String) As String
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim timeText As String
    Dim clientText As String
    Dim idValue As Long
    Dim lines As String

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    firstRow = modNames.RangeOfName(NAME_CAL_FIRST).Row
    lastRow = modNames.RangeOfName(NAME_CAL_LAST).Row

    For rowIndex = firstRow To lastRow
        timeText = CellTimeText(ws.Cells(rowIndex, DASH_COL_START))
        ' The renderer draws one row per slot of the working day and the block holds
        ' more rows than a short day needs (the cap of section 6.9): the first row
        ' time is the end of the grid, so the plan is not padded with rows nobody sees.
        If Len(timeText) = 0 Then Exit For

        clientText = Trim$(CStr(ws.Cells(rowIndex, DASH_COL_CLIENT).Text))
        idValue = CLng(Val(CStr(ws.Cells(rowIndex, DASH_COL_ID).Value)))

        If idValue > 0 Then
            lines = lines & "plan|" & isoDate & "|" & timeText & "|start|" & _
                    CStr(idValue) & "|" & clientText & "|" & _
                    Trim$(CStr(ws.Cells(rowIndex, DASH_COL_SERVICE).Text)) & vbCrLf
        ElseIf Left$(clientText, 1) = modText.ContinuationMark() Then
            lines = lines & "plan|" & isoDate & "|" & timeText & "|tail|" & _
                    Mid$(clientText, InStr(clientText, "#") + 1) & vbCrLf
        Else
            lines = lines & "plan|" & isoDate & "|" & timeText & "|free" & vbCrLf
        End If
    Next rowIndex

    CrosscheckPlan = lines
End Function

' The slot time of a cell, whatever Excel made of it: the text of the cell, or -
' when a format turned it into a time - the time it means.
Private Function CellTimeText(ByRef cell As Range) As String
    Dim value As Variant

    value = cell.Value
    If IsDate(value) Then
        CellTimeText = modApp.MinutesToTime(CLng(Hour(CDate(value))) * 60 + _
                                            CLng(Minute(CDate(value))))
        Exit Function
    End If

    CellTimeText = Trim$(CStr(cell.Text))
End Function

' The rows of the rendered report grid, in the format the model writes
' ("report|name|group|count|cents"): the group is the text of the cell, the count
' and the revenue are read back as numbers (the cell holds 25 while the user reads
' "25.00", so the format of Config is not part of the comparison). The dump stops
' at the total row, which is the last row the model writes as well.
Private Function CrosscheckReport(ByVal caseName As String) As String
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim groupText As String
    Dim countText As String
    Dim centsText As String
    Dim lines As String
    Dim cellValue As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_REPORTS)
    firstRow = modNames.RangeOfName(NAME_GRID_REPORTS_FIRST).Row
    lastRow = modNames.RangeOfName(NAME_GRID_REPORTS_LAST).Row

    For rowIndex = firstRow To lastRow
        groupText = Trim$(CStr(ws.Cells(rowIndex, REPORT_COL_GROUP).Value))

        ' An empty cell must stay empty: IsNumeric(Empty) is True in VBA, so the
        ' emptiness is checked first (a heading row has no numbers in either lane).
        countText = vbNullString
        cellValue = ws.Cells(rowIndex, REPORT_COL_COUNT).Value
        If Not IsEmpty(cellValue) Then
            If IsNumeric(cellValue) Then countText = CStr(CLng(CDbl(cellValue)))
        End If

        centsText = vbNullString
        cellValue = ws.Cells(rowIndex, REPORT_COL_REVENUE).Value
        If Not IsEmpty(cellValue) Then
            If IsNumeric(cellValue) Then centsText = CStr(CLng(CDbl(cellValue) * 100))
        End If

        lines = lines & "report|" & caseName & "|" & groupText & "|" & countText & _
                "|" & centsText & vbCrLf

        If StrComp(groupText, modText.Text(modText.KEY_REPORT_TOTAL), vbTextCompare) = 0 Then Exit For
    Next rowIndex

    CrosscheckReport = lines
End Function
' build\xls-crosscheck.txt next to the workbook: the answers of the workbook, for
' tools\build.vbs to compare with build\sim-crosscheck.txt.
Private Sub WriteCrosscheckAnswers(ByVal answers As String, ByVal caseCount As Long)
    Dim fileNumber As Integer
    Dim lines As Variant
    Dim index As Long
    Dim targetPath As String

    targetPath = ThisWorkbook.Path & Application.PathSeparator & "xls-crosscheck.txt"

    fileNumber = FreeFile
    Open targetPath For Output As #fileNumber
    Print #fileNumber, "# " & CStr(caseCount) & " cases of " & CrosscheckPath()
    Print #fileNumber, "# answers of the workbook, " & Now

    lines = Split(answers, vbCrLf)
    For index = 0 To UBound(lines)
        If Len(CStr(lines(index))) > 0 Then Print #fileNumber, CStr(lines(index))
    Next index
    Close #fileNumber

    Note "cross check: " & CStr(caseCount) & " cases -> " & targetPath
End Sub

' =====================================================================
'  the typed client search (§9.5)
' =====================================================================

' With a few hundred clients the dropdown of Booking cannot be used any more and
' the owner cannot remember whether he stored the first name or the surname
' first. The search matches every word of the term anywhere in the name or in the
' phone, so the order of the words in the record does not matter.
Public Sub Test_ClientSearch()
    Dim values As Object
    Dim data As Object
    Dim result As Object
    Dim rows As Collection
    Dim picker As frmClientPicker
    Dim clientId As Long
    Dim clientsBefore As Long
    Dim clientName As String

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    clientsBefore = modData.ClientsRowCount()

    clientName = "Zaharieva Violeta"
    Set values = CreateObject("Scripting.Dictionary")
    values(COL_CLIENT_NAME) = clientName
    values(COL_CLIENT_PHONE) = "0888555123"
    values(COL_CLIENT_EMAIL) = "violeta@example.com"
    values(COL_CLIENT_COMMENT) = "the client of the search test"
    clientId = modData.AppendClient(values)
    AssertTrue (clientId > 0), "the client of the search test exists"

    ' the pure rule first: a fragment, the other word order, the phone
    AssertTrue modClients.Matches(clientName, "0888555123", "zaha"), _
               "a fragment of the first name matches"
    AssertTrue modClients.Matches(clientName, "0888555123", "violeta"), _
               "the first name alone matches"
    AssertTrue modClients.Matches(clientName, "0888555123", "zaharieva violeta"), _
               "the words in the stored order match"
    AssertTrue modClients.Matches(clientName, "0888555123", "violeta zaharieva"), _
               "the words in the other order match too - that is the whole point"
    AssertTrue modClients.Matches(clientName, "0888555123", "0888555123"), _
               "the phone matches"
    AssertTrue Not modClients.Matches(clientName, "0888555123", "zaha stoianov"), _
               "every word has to appear somewhere"
    AssertTrue Not modClients.Matches(clientName, "0888555123", vbNullString), _
               "an empty term matches nobody (the caller decides what that means)"

    ' and the rows the picker shows
    Set rows = modClients.SearchRows("violeta zaharieva")
    AssertTrue ContainsRow(rows, clientId), "the search finds the client by the reversed words"

    Set rows = modClients.SearchRows("555123")
    AssertTrue ContainsRow(rows, clientId), "the search finds the client by a fragment of the phone"

    Set rows = modClients.SearchRows("qwertyuiop")
    AssertEqualsLong rows.Count, 0, "a term without a match answers no rows"

    Set rows = modClients.SearchRows(vbNullString)
    AssertEqualsLong rows.Count, clientsBefore + 1, "an empty term answers every client"

    ' the form: it builds itself without a dialog, narrows with every change of the
    ' box and answers the picked name
    Set picker = New frmClientPicker
    picker.InitialiseWithSearch "zaharieva"
    AssertEqualsLong picker.ResultCount(), 1, "the form shows the one match"
    AssertTrue (InStr(1, picker.ResultAt(0), clientName, vbTextCompare) > 0), _
               "the result reads as the name of the client"

    picker.SearchChanged "violeta zaharieva"
    AssertEqualsLong picker.ResultCount(), 1, "the reversed order narrows to the same client"

    picker.SearchChanged "qwertyuiop"
    AssertEqualsLong picker.ResultCount(), 0, "no match leaves the list empty"
    AssertTrue (Len(picker.HintText()) > 0), "and the form says so"

    picker.SearchChanged "violeta"
    AssertEqualsLong picker.ResultCount(), 1, "the search follows the box"
    picker.ChooseResult 0
    picker.ButtonClicked ROLE_OK, vbNullString
    AssertEqualsText picker.PickedClientText, clientName, "OK answers the name of the client"

    ' form.pick_client: the action fills the field, and the id follows from the name
    modUiBooking.ClearForm
    Set data = CreateObject("Scripting.Dictionary")
    Set result = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_VALUE) = clientName
    modUiBooking.OnPickClient data, result
    AssertTrue result(modRoutes.RESULT_OK), "the picker action reports success"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT), clientName, _
                     "the client field carries the picked name"
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT_ID), CStr(clientId), _
                     "and the id is filled in from the name"

    ' a cancellation changes nothing: the field keeps what it had
    Set result = CreateObject("Scripting.Dictionary")
    data(modRoutes.KEY_VALUE) = vbNullString
    modNames.WriteName NAME_BOOKING_CLIENT, clientName
    modUiBooking.OnPickClient data, result
    AssertEqualsText modNames.ReadName(NAME_BOOKING_CLIENT), clientName, _
                     "a cancelled search leaves the field alone"

    ' the button is visible and clickable with nothing chosen: it must sit on a cell
    ' of its own (a cell that another name writes into loses the caption - the bug
    ' of 19.09.2026, when it was put on o_booking_client_id)
    AssertEqualsText modNames.ReadName(NAME_BTN_BOOKING_CLIENT_PICK), "Search", _
                     "the search button reads Search while the field is empty"
    AssertTrue (StrComp(modNames.RangeOfName(NAME_BTN_BOOKING_CLIENT_PICK).Address, _
                        modNames.RangeOfName(NAME_BOOKING_CLIENT_ID).Address, _
                        vbTextCompare) <> 0), _
               "and it is not the cell of the client id preview"

    ' the wiring and the texts
    AssertEqualsText modRoutes.RouteFor(ACTION_FORM_PICK_CLIENT), "modUiBooking.OnPickClient", _
                     "form.pick_client points at the action of the search"
    AssertEqualsText modText.Text(modText.KEY_PICKER_CLIENT_TITLE), "Find a client", _
                     "the picker has a title"
    AssertEqualsText modText.Text(modText.KEY_PICKER_CLIENT_SEARCH), _
                     "Type any part of the name or the phone", "and a hint for the box"
    AssertTrue (InStr(1, modText.TextFormat(modText.KEY_OK_PICKED_CLIENT, "A", "7"), "A", _
                     vbBinaryCompare) > 0), "the confirmation names the client"

    ' the test cleans up after itself: the sample data of the template has to stay
    ' the way the other tests expect it (this one added a client)
    AssertTrue modData.DeleteRowById(SHEET_DB_CLIENTS, clientId), _
               "the client of the search test is removed again"
    AssertEqualsLong modData.ClientsRowCount(), clientsBefore, _
                     "and the sample data is back"

    modUiBooking.ClearForm
    modUi.SetInteractiveUi True
End Sub

' True when one of the rows ("id|name|phone") belongs to this client.
Private Function ContainsRow(ByRef rows As Collection, ByVal clientId As Long) As Boolean
    Dim item As Variant
    Dim idText As String
    Dim nameText As String
    Dim phoneText As String

    If rows Is Nothing Then Exit Function

    For Each item In rows
        modPickers.SplitSearchRow CStr(item), idText, nameText, phoneText
        If CLng(Val(idText)) = clientId Then
            ContainsRow = True
            Exit Function
        End If
    Next item
End Function

' =====================================================================
'  the start picker follows the working day (§9.3, 20.09.2026)
' =====================================================================

' With 09:00-19:00 in Config the picker offers 09..18 and nothing else, so it
' cannot suggest a start the domain would refuse. A duration is not a clock time
' and keeps the whole day.
Public Sub Test_StartPickerWindow()
    Dim picker As frmTimePicker

    AssertEqualsLong modPickers.StartHourFirst(540, 1140), 9, "09:00 is the first hour"
    AssertEqualsLong modPickers.StartHourLast(540, 1140), 18, _
                     "18 is the last hour of a day that ends at 19:00"
    AssertEqualsLong modPickers.StartHourFirst(0, 1440), 0, "the whole day starts at 00"
    AssertEqualsLong modPickers.StartHourLast(0, 1440), 23, "and ends at 23"
    AssertEqualsLong modPickers.StartHourFirst(600, 600), 0, _
                     "a window without hours falls back to 00"
    AssertEqualsLong modPickers.StartHourLast(600, 600), 23, "and to 23"
    AssertEqualsLong modPickers.StartHourFirst(570, 1140), 9, "half past nine is still hour 09"

    ' the form follows the window it is given
    Set picker = New frmTimePicker
    picker.InitialiseWithTime FIELD_START, vbNullString, 30, vbNullString, 540, 1140
    AssertEqualsLong picker.HourOptionCount(), 10, "09:00-19:00 leaves ten hours to start on"
    AssertEqualsText picker.HourOptionAt(0), "09", "the first of them is 09"
    AssertEqualsText picker.HourOptionAt(9), "18", "the last of them is 18"

    picker.InitialiseWithTime FIELD_DURATION, vbNullString, 0, vbNullString, 540, 1140
    AssertEqualsLong picker.HourOptionCount(), 24, "a duration keeps every hour"

    ' an appointment outside the window (the working day was changed after it was
    ' booked) still opens on its own hour instead of falling to the first slot
    picker.InitialiseWithTime FIELD_START, "08:00", 30, vbNullString, 540, 1140
    AssertEqualsLong picker.Hours(), 8, "08:00 stays the choice of the form"

    ' and the day of Config is what the picker is given, not a constant
    AssertEqualsLong modAppointments.DayStartMinutes(), 540, "the default day starts at 09:00"
    AssertEqualsLong modAppointments.DayEndMinutes(), 1140, "and ends at 19:00"
End Sub

' =====================================================================
'  the phone is text, not a number (24.09.2026)
' =====================================================================

' In a General cell Excel turns 0888123456 into 888123456 and the leading zero is
' gone - the list of clients showed exactly that. The renderer writes the phone into
' a cell formatted as text (and the template carries the same format), so the zero
' survives; this test would catch a restyle that drops it.
Public Sub Test_PhoneStaysText()
    Dim values As Object
    Dim ws As Worksheet
    Dim firstCell As Range
    Dim clientId As Long
    Dim rowIndex As Long
    Dim rowsBefore As Long

    EnsureConfigDefaults
    modUi.SetInteractiveUi False
    rowsBefore = modData.ClientsRowCount()

    Set values = CreateObject("Scripting.Dictionary")
    values(COL_CLIENT_NAME) = "Phone Test"
    values(COL_CLIENT_PHONE) = "0888123456"
    values(COL_CLIENT_EMAIL) = "phone@example.com"
    values(COL_CLIENT_COMMENT) = "leading zero"
    clientId = modData.AppendClient(values)
    AssertTrue (clientId > 0), "the client with a leading zero exists"

    Set ws = ThisWorkbook.Worksheets(SHEET_CLIENTS_ADMIN)
    modUiClients.Render

    Set firstCell = modNames.RangeOfName(NAME_GRID_CLIENTS_FIRST)
    rowIndex = firstCell.Row + rowsBefore

    AssertEqualsText ws.Cells(rowIndex, CLIENTS_COL_PHONE).Text, "0888123456", _
                     "the phone keeps its leading zero in the list"
    AssertEqualsText CStr(ws.Cells(rowIndex, CLIENTS_COL_PHONE).NumberFormat), "@", _
                     "because the cell is formatted as text"

    AssertTrue modData.DeleteRowById(SHEET_DB_CLIENTS, clientId), "the test client is removed"
    modUiClients.Render
    AssertEqualsLong modData.ClientsRowCount(), rowsBefore, "and the list is back to what it was"
End Sub

' =====================================================================
'  the modal confirmation (24.09.2026)
' =====================================================================

' The two step confirmation (click, then click again within five seconds) became
' one modal question with OK and Cancel. The dialog builds itself without being
' shown - the pickers are tested the same way - and answers what was clicked.
Public Sub Test_ConfirmDialog()
    Dim ask As frmConfirm

    Set ask = New frmConfirm
    ask.InitialiseWithQuestion "Are you sure?"
    AssertEqualsText ask.MessageText(), "Are you sure?", "the dialog shows the question it was given"
    AssertFalse ask.Confirmed, "nothing is confirmed before an answer"

    ask.ButtonClicked ROLE_CANCEL, vbNullString
    AssertFalse ask.Confirmed, "Cancel answers no"

    ask.ButtonClicked ROLE_OK, vbNullString
    AssertTrue ask.Confirmed, "OK answers yes"

    ' the question every delete uses, and the answer without a dialog
    AssertEqualsText modText.Text(modText.KEY_CONFIRM_ASK), "Are you sure?", _
                     "the shared question is the one in modText"

    modUi.SetInteractiveUi False
    AssertTrue modApp.ConfirmNow("any.key"), "without a dialog the confirmation is a yes"
End Sub

' =====================================================================
'  the order of the appointments list (25.09.2026)
' =====================================================================

' The rule of the owner: the newest date is on top, and inside one day the latest
' start. The order is a pure function on the table, so the test walks it and checks
' that every following key is not newer than the one above - whatever the sample
' data happens to be (no date is hard coded).
Public Sub Test_AppointmentsOrder()
    Dim order As Collection
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim previous As String
    Dim current As String
    Dim before As Long

    before = modData.AppointmentsRowCount()
    Set order = modData.AppointmentRowsNewestFirst()
    AssertEqualsLong order.Count, before, "every appointment is in the order"

    Set table = modData.AppointmentsTable()
    previous = vbNullString
    For rowIndex = 1 To order.Count
        current = modData.DateTextOfRow(table, order(rowIndex)) & "|" & _
                  modData.StartTextOfRow(table, order(rowIndex))
        If Len(previous) > 0 Then
            AssertTrue (previous >= current), _
                       "row " & CStr(rowIndex) & " is not newer than the one above"
        End If
        previous = current
    Next rowIndex

    AssertEqualsLong modData.AppointmentsRowCount(), before, _
                     "reading the order does not touch the table"
End Sub

