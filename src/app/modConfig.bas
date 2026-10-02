Attribute VB_Name = "modConfig"
Option Explicit

' =====================================================================
'  Service_Manager - configuration (Phase 1)
'
'  Every constant of the application lives here, in the declarations section,
'  before the first procedure. The POC learned this the hard way: a Const after
'  the first procedure is reported as "Only comments may appear after End Sub"
'  and the names after it become invisible for the whole project.
'
'  The cells and the sheets are named in modNames; the texts the user reads are
'  in modText; the actions are in modRoutes.
' =====================================================================

' ---- application -----------------------------------------------------
Public Const APP_VERSION As String = "1.1.0"
Public Const APP_NAME As String = "Service Manager"

' ---- sheets ----------------------------------------------------------
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_BOOKING As String = "Booking"
Public Const SHEET_APPOINTMENTS As String = "Appointments"
Public Const SHEET_CLIENTS_ADMIN As String = "ClientsAdmin"
Public Const SHEET_SERVICES_ADMIN As String = "ServicesAdmin"
Public Const SHEET_REPORTS As String = "Reports"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_LOG As String = "Log"
Public Const SHEET_ABOUT As String = "About"
Public Const SHEET_LISTS As String = "_Lists"
Public Const SHEET_DB_CLIENTS As String = "db_clients"
Public Const SHEET_DB_SERVICES As String = "db_services"
Public Const SHEET_DB_APPOINTMENTS As String = "db_appointments"

' ---- columns of the data sheets --------------------------------------
Public Const COL_ID As String = "id"
Public Const COL_CLIENT_NAME As String = "name"
Public Const COL_CLIENT_PHONE As String = "phone"
Public Const COL_CLIENT_EMAIL As String = "email"
Public Const COL_CLIENT_COMMENT As String = "comment"
Public Const COL_CLIENT_ID As String = "client_id"
Public Const COL_SERVICE_NAME As String = "service"
Public Const COL_SERVICE_PRICE As String = "price"
Public Const COL_SERVICE_DURATION As String = "duration"
Public Const COL_SERVICE_ID As String = "service_id"
Public Const COL_DATE As String = "date"
Public Const COL_START As String = "start"
Public Const COL_LOG_TIME As String = "timestamp"
Public Const COL_LOG_ACTION As String = "action"
Public Const COL_LOG_OK As String = "ok"
Public Const COL_LOG_MESSAGE As String = "message"

' ---- data ranges -----------------------------------------------------
Public Const DB_HEADER_ROW As Long = 1
Public Const DB_FIRST_ROW As Long = 2

' ---- defaults (the Config sheet holds the live values) ---------------
Public Const DEFAULT_DAY_START As String = "09:00"
Public Const DEFAULT_DAY_END As String = "19:00"
Public Const DEFAULT_STEP_MINUTES As Long = 5
Public Const DEFAULT_DURATION_MINUTES As Long = 10
Public Const DEFAULT_CURRENCY As String = "EUR"
Public Const DEFAULT_NUMBER_FORMAT As String = "#,##0.00"
Public Const DEFAULT_NONWORK_WEEKDAYS As String = "6,7"

' ---- the rules of the booking validator (§8.2.7) ---------------------
' The code is the identity of the rule; the parameters follow after a "|"
' ("error.conflict|1|15:00"). modAppointments.RuleCode answers with the code,
' Problems turns it into the sentence the user reads and the cross check of
' tools\sim compares the two implementations on the codes.
Public Const RULE_REQUIRED As String = "error.required"
Public Const RULE_NOT_FOUND As String = "error.not_found"
Public Const RULE_NONWORK As String = "error.nonwork"
Public Const RULE_TIME As String = "error.time"
Public Const RULE_RANGE As String = "error.range"
Public Const RULE_STEP As String = "error.step"
Public Const RULE_CONFLICT As String = "error.conflict"
Public Const RULE_PERIOD As String = "error.period"

' ---- the rules of the Config page (phase 6) --------------------------
Public Const RULE_DAY_ORDER As String = "error.day_order"
Public Const RULE_FORMAT As String = "error.format"
Public Const RULE_WEEKDAY As String = "error.weekday"
Public Const RULE_DATES As String = "error.dates"

' ---- the report grid (phase 5) ---------------------------------------
' The three columns of the report grid (the header row of the template says
' group, count, revenue) and the format of its count column; the revenue column
' uses the number format of Config.
Public Const REPORT_COL_GROUP As Long = 1
Public Const REPORT_COL_COUNT As Long = 2
Public Const REPORT_COL_REVENUE As Long = 3
Public Const REPORT_COUNT_FORMAT As String = "0"

' ---- log -------------------------------------------------------------
Public Const LOG_MAX_ROWS As Long = 1000

' ---- dashboard views -------------------------------------------------
Public Const VIEW_DAY As String = "day"
Public Const VIEW_WEEK As String = "week"
Public Const VIEW_MONTH As String = "month"
Public Const VIEW_TEXT_DAY As String = "Daily"
Public Const VIEW_TEXT_WEEK As String = "Weekly"
Public Const VIEW_TEXT_MONTH As String = "Monthly"

' ---- dashboard body (phase 4, PROJECT_CONCEPT §6.9) ------------------
' The body is the block cal_first..cal_last; the row above cal_first carries the
' labels and belongs to the renderer, because their meaning changes with the view.
Public Const DASH_BODY_FIRST_COLUMN As Long = 1
Public Const DASH_BODY_LAST_COLUMN As Long = 9
Public Const DASH_HEADER_OFFSET As Long = 1

' Daily: one appointment per row; the last row of the block is the footer (free
' slots and [+ New]), so a click on it means "book another one".
Public Const DASH_COL_START As Long = 1
Public Const DASH_COL_END As Long = 2
Public Const DASH_COL_CLIENT As Long = 3
Public Const DASH_COL_SERVICE As Long = 4
Public Const DASH_COL_DURATION As Long = 5
Public Const DASH_COL_PRICE As Long = 6
Public Const DASH_COL_EDIT As Long = 7
Public Const DASH_COL_DELETE As Long = 8
Public Const DASH_COL_ID As Long = 9

' Weekly: the time column, then the seven days (the grid has no action column any
' more - deleting happens in Booking, §6.9).
Public Const DASH_WEEK_TIME_COL As Long = 1
Public Const DASH_WEEK_FIRST_DAY_COL As Long = 2

' Monthly: seven columns (Monday first), MONTH_GRID_WEEKS rows.
Public Const DASH_MONTH_FIRST_COL As Long = 1

' ---- date and time ---------------------------------------------------
Public Const ISO_DATE_LENGTH As Long = 10
Public Const TIME_TEXT_LENGTH As Long = 5
Public Const MINUTES_PER_HOUR As Long = 60
Public Const MINUTES_PER_DAY As Long = 1440
Public Const DAYS_PER_WEEK As Long = 7
Public Const MONTH_GRID_WEEKS As Long = 6

' ---- colours ---------------------------------------------------------
Public Const COLOR_STATUS_OK As Long = &HD8F0D8        ' RGB(216,240,216)
Public Const COLOR_STATUS_ERROR As Long = &HCEC7FF     ' RGB(255,199,206)
Public Const COLOR_TEXT As Long = 0
Public Const COLOR_WARNING_FONT As Long = &H0000C0     ' RGB(192,0,0)
Public Const COLOR_MUTED As Long = &H808080            ' RGB(128,128,128)
Public Const COLOR_FREE As Long = &HFFFFFF
Public Const COLOR_BOOKED As Long = &HD8F0D8
Public Const COLOR_CONFLICT As Long = &HCEC7FF
Public Const COLOR_TODAY As Long = &HCCFFFF
Public Const COLOR_NOWORK As Long = &HF0F0F0
Public Const COLOR_GRID_HEADER As Long = &HD9D9D9
Public Const COLOR_BUTTON As Long = &HE8E0D0           ' RGB(208,224,232)
Public Const COLOR_BUTTON_ACTIVE As Long = &H90EE90    ' RGB(144,238,144)
Public Const COLOR_PENDING As Long = &HFFFF            ' RGB(255,255,0)

' ---- prefixes of the names -------------------------------------------
Public Const PREFIX_BUTTON As String = "b_"
Public Const PREFIX_INPUT As String = "i_"
Public Const PREFIX_OUTPUT As String = "o_"
Public Const PREFIX_STATE As String = "s_"
Public Const PREFIX_LIST As String = "list_"
Public Const PREFIX_CALENDAR As String = "cal_"
Public Const PREFIX_GRID As String = "grid_"
Public Const PREFIX_LOG As String = "log_"

' ---- the admin lists (ClientsAdmin / ServicesAdmin) ------------------
' The grid geometry is fixed (the template declares the anchors); the code
' addresses the columns of a row by number, because a rendered row has no name.
' Every row carries two link cells: [Edit] loads that row into the form above and
' [Delete] removes it - one column each, because a cell can hold one hyperlink
' only (19.09.2026).
Public Const CLIENTS_COL_ID As Long = 1
Public Const CLIENTS_COL_NAME As Long = 2
Public Const CLIENTS_COL_PHONE As Long = 3
Public Const CLIENTS_COL_EMAIL As Long = 4
Public Const CLIENTS_COL_COMMENT As Long = 5
Public Const CLIENTS_COL_ACTION As Long = 6
Public Const CLIENTS_COL_DELETE As Long = 7
Public Const SERVICES_COL_ID As Long = 1
Public Const SERVICES_COL_NAME As Long = 2
Public Const SERVICES_COL_PRICE As Long = 3
Public Const SERVICES_COL_DURATION As Long = 4
Public Const SERVICES_COL_ACTION As Long = 5
Public Const SERVICES_COL_DELETE As Long = 6
Public Const APPOINTMENTS_COL_ID As Long = 1
Public Const APPOINTMENTS_COL_DATE As Long = 2
Public Const APPOINTMENTS_COL_START As Long = 3
Public Const APPOINTMENTS_COL_END As Long = 4
Public Const APPOINTMENTS_COL_DURATION As Long = 5
Public Const APPOINTMENTS_COL_CLIENT As Long = 6
Public Const APPOINTMENTS_COL_PHONE As Long = 7
Public Const APPOINTMENTS_COL_EMAIL As Long = 8
Public Const APPOINTMENTS_COL_SERVICE As Long = 9
Public Const APPOINTMENTS_COL_EDIT As Long = 10
Public Const APPOINTMENTS_COL_DELETE As Long = 11
Public Const LIST_ACTION_TEXT As String = "[Delete]"
Public Const LIST_EDIT_TEXT As String = "[Edit]"

' ---- confirmation of a destructive action (§6.6) ---------------------
' (the second click and its five seconds are gone: the confirmation is a modal
'  dialog now, §6.6 - 24.09.2026)

' ---- backups (§8.2.4) ------------------------------------------------
Public Const BACKUP_FOLDER As String = "backups"
Public Const BACKUP_KEEP As Long = 10

' ---- validation ------------------------------------------------------
Public Const MIN_PHONE_DIGITS As Long = 3

' ---- paint states of a cell button (modUi.PaintCell) -----------------
Public Const PAINT_NORMAL As String = "normal"
Public Const PAINT_ACTIVE As String = "active"
Public Const PAINT_MUTED As String = "muted"
Public Const PAINT_PENDING As String = "pending"

' ---- the time fields of Booking (§7.5) -------------------------------
Public Const FIELD_START As String = "start"
Public Const FIELD_DURATION As String = "duration"

' ---- the appointment rules (§8.2) ------------------------------------
Public Const MAX_DURATION_MINUTES As Long = 1440

' ---- pickers: geometry and behaviour (§9) ----------------------------
Public Const PICKER_MARGIN As Single = 6
Public Const PICKER_GAP As Single = 6
Public Const PICKER_LABEL_HEIGHT As Single = 14
Public Const PICKER_CELL_WIDTH As Single = 34
Public Const PICKER_CELL_HEIGHT As Single = 22
Public Const PICKER_HEADER_HEIGHT As Single = 24
Public Const PICKER_WEEKDAY_HEIGHT As Single = 18
Public Const PICKER_NAV_WIDTH As Single = 28
Public Const PICKER_GRID_DAYS As Long = 7
Public Const PICKER_GRID_WEEKS As Long = 6
Public Const PICKER_START_POSITION As Long = 1        ' CenterScreen
Public Const TIME_STEP_MINUTES As Long = 5
Public Const COLOR_PICKER_DISABLED As Long = &HE0E0E0  ' RGB(224,224,224)

' The window of the date picker (§9.2): the previous year (a booking made in
' December for January) up to ten years ahead. Nothing older can be picked - the
' first version of the picker could hand out 1899 dates, because every day
' button got its value while the shown month was still 0.
Public Const PICKER_YEARS_BACK As Long = 1
Public Const PICKER_YEARS_FORWARD As Long = 10

' ---- picker roles: the name of a button decides what its click does ---
Public Const ROLE_PREV_MONTH As String = "prev_month"
Public Const ROLE_NEXT_MONTH As String = "next_month"
Public Const ROLE_PREV_YEAR As String = "prev_year"
Public Const ROLE_NEXT_YEAR As String = "next_year"
Public Const ROLE_TODAY As String = "today"
Public Const ROLE_DAY As String = "day"
Public Const ROLE_OK As String = "ok"
Public Const ROLE_CANCEL As String = "cancel"

' ---- the client picker of Booking (§9.5) -----------------------------
' A typed search over the clients: the term is matched against the name and the
' phone, so a client is found whichever half of the name was typed first.
Public Const ROLE_ITEM As String = "item"
Public Const PICKER_CLIENT_SEARCH_WIDTH As Single = 250
Public Const PICKER_CLIENT_LIST_HEIGHT As Single = 160

' The confirmation dialog (§6.6): one question, two buttons.
Public Const PICKER_CONFIRM_WIDTH As Single = 250
Public Const PICKER_QUESTION_HEIGHT As Single = 44

' =====================================================================
'  procedures (everything below is code, all constants live above)
' =====================================================================

' The UI sheets as a comma separated list: no array crosses a module boundary.
Public Function UiSheetNames() As String
    UiSheetNames = SHEET_DASHBOARD & "," & SHEET_BOOKING & "," & SHEET_CLIENTS_ADMIN & "," & _
                   SHEET_SERVICES_ADMIN & "," & SHEET_APPOINTMENTS & "," & SHEET_REPORTS & "," & _
                   SHEET_CONFIG & "," & SHEET_ABOUT
End Function

' The data sheets as a comma separated list.
Public Function DataSheetNames() As String
    DataSheetNames = SHEET_DB_CLIENTS & "," & SHEET_DB_SERVICES & "," & SHEET_DB_APPOINTMENTS
End Function

' True for a sheet the user works on (the event class attaches to these).
Public Function IsUiSheet(ByVal sheetName As String) As Boolean
    IsUiSheet = (InStr(1, "," & UiSheetNames() & ",", "," & sheetName & ",", vbTextCompare) > 0)
End Function

' True for a sheet that holds data.
Public Function IsDataSheet(ByVal sheetName As String) As Boolean
    IsDataSheet = (InStr(1, "," & DataSheetNames() & ",", "," & sheetName & ",", vbTextCompare) > 0)
End Function
