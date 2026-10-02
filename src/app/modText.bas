Attribute VB_Name = "modText"
Option Explicit

' =====================================================================
'  Service_Manager - every text the user can see (Phase 1)
'
'  The application speaks English. No string that reaches the screen is written
'  anywhere else: the code calls modText.Text("key") or modText.TextFormat(
'  "key", arg1, ...). The keys are constants of this module, so a test can walk
'  them (Test_TextCoverage) and a second language would be one more Select Case.
' =====================================================================

' ---- keys ------------------------------------------------------------
Public Const KEY_STATUS_READY As String = "ok.ready"
Public Const KEY_STATUS_DONE As String = "ok.done"
Public Const KEY_STATUS_REFRESHED As String = "ok.refreshed"
Public Const KEY_STATUS_NAVIGATED As String = "ok.navigated"
Public Const KEY_STATUS_DIAGNOSTICS As String = "ok.diagnostics"
Public Const KEY_STATUS_VIEW As String = "ok.view"

Public Const KEY_ERROR_NO_ACTION As String = "error.no_action"
Public Const KEY_ERROR_UNKNOWN_ACTION As String = "error.unknown_action"
Public Const KEY_ERROR_BUSY As String = "error.busy"
Public Const KEY_ERROR_FAILED As String = "error.failed"
Public Const KEY_ERROR_NAME_MISSING As String = "error.name_missing"
Public Const KEY_ERROR_NAME_NOT_CELL As String = "error.name_not_cell"
Public Const KEY_ERROR_CONTRACT As String = "error.contract"
Public Const KEY_ERROR_ROUTE As String = "error.route"
Public Const KEY_ERROR_NO_SHEET As String = "error.no_sheet"
Public Const KEY_ERROR_NO_DATA As String = "error.no_data"
Public Const KEY_ERROR_READ_ONLY As String = "error.read_only"

Public Const KEY_WARN_PENDING As String = "warn.pending_routes"
Public Const KEY_WARN_EMPTY As String = "warn.empty"

' ---- the dashboard body (phase 4) ------------------------------------
Public Const KEY_DASH_COL_TIME As String = "dash.col_time"
Public Const KEY_DASH_COL_END As String = "dash.col_end"
Public Const KEY_DASH_COL_CLIENT As String = "dash.col_client"
Public Const KEY_DASH_COL_SERVICE As String = "dash.col_service"
Public Const KEY_DASH_COL_DURATION As String = "dash.col_duration"
Public Const KEY_DASH_COL_PRICE As String = "dash.col_price"
Public Const KEY_DASH_COL_EDIT As String = "dash.col_edit"
Public Const KEY_DASH_COL_DELETE As String = "dash.col_delete"
Public Const KEY_DASH_COL_ID As String = "dash.col_id"
Public Const KEY_DASH_DAY_HEADER As String = "dash.day_header"
Public Const KEY_DASH_RANGE As String = "dash.range"
Public Const KEY_DASH_WHO As String = "dash.who"
Public Const KEY_DASH_SLOT As String = "dash.slot"
Public Const KEY_DASH_TAIL As String = "dash.tail"
Public Const KEY_DASH_MORE As String = "dash.more"
Public Const KEY_DASH_MONTH_DAY As String = "dash.month_day"
Public Const KEY_DASH_TIP_EDIT As String = "dash.tip_edit"
Public Const KEY_DASH_TIP_DELETE As String = "dash.tip_delete"
Public Const KEY_DASH_TIP_NEW As String = "dash.tip_new"
Public Const KEY_DASH_TIP_OPEN_DAY As String = "dash.tip_open_day"
Public Const KEY_BTN_CANCEL As String = "booking.btn_cancel"
Public Const KEY_BTN_DELETE As String = "booking.btn_delete"
Public Const KEY_LABEL_PERIOD_START As String = "label.period_start"
Public Const KEY_LABEL_PERIOD_END As String = "label.period_end"

' ---- the Config page (phase 6) ---------------------------------------
Public Const KEY_CONFIG_APPLIED As String = "config.applied"
Public Const KEY_CONFIG_DEFAULTS As String = "config.defaults"
Public Const KEY_CONFIG_DATA_HIDDEN As String = "config.data_hidden"
Public Const KEY_CONFIG_DATA_SHOWN As String = "config.data_shown"
Public Const KEY_CONFIG_DATE_ADDED As String = "config.date_added"
Public Const KEY_CONFIG_DATES_CLEARED As String = "config.dates_cleared"
Public Const KEY_ERROR_DAY_ORDER As String = "error.day_order"
Public Const KEY_ERROR_FORMAT As String = "error.format"
Public Const KEY_ERROR_WEEKDAY As String = "error.weekday"
Public Const KEY_ERROR_DATES As String = "error.dates"
Public Const KEY_ERROR_SCHEMA_NEWER As String = "error.schema_newer"

Public Const KEY_ABOUT_VERSION As String = "about.version"
Public Const KEY_ABOUT_FILE As String = "about.file"
Public Const KEY_ABOUT_SHEETS As String = "about.sheets"

Public Const KEY_CLIENTS_ADDED As String = "clients.added"
Public Const KEY_CLIENTS_DELETED As String = "clients.deleted"
Public Const KEY_CLIENTS_UPDATED As String = "clients.updated"
Public Const KEY_CLIENTS_TITLE As String = "clients.title"
Public Const KEY_CLIENTS_NONE As String = "clients.none"
Public Const KEY_CLIENTS_BTN_ADD As String = "clients.btn_add"
Public Const KEY_CLIENTS_BTN_UPDATE As String = "clients.btn_update"
Public Const KEY_LIST_CAPPED As String = "list.capped"
Public Const KEY_SERVICES_ADDED As String = "services.added"
Public Const KEY_SERVICES_DELETED As String = "services.deleted"
Public Const KEY_SERVICES_UPDATED As String = "services.updated"
Public Const KEY_SERVICES_TITLE As String = "services.title"
Public Const KEY_SERVICES_NONE As String = "services.none"
Public Const KEY_SERVICES_BTN_ADD As String = "services.btn_add"
Public Const KEY_SERVICES_BTN_UPDATE As String = "services.btn_update"
Public Const KEY_APPOINTMENTS_TITLE As String = "appointments.title"
Public Const KEY_APPOINTMENTS_NONE As String = "appointments.none"
Public Const KEY_STATUS_APPOINTMENTS_REFRESHED As String = "ok.appointments_refreshed"
' The two texts every admin form needs while a row is being edited (§6.10).
Public Const KEY_FORM_MODE_EDIT As String = "form.mode_edit"
Public Const KEY_FORM_LOADED As String = "form.loaded"
Public Const KEY_CONFIRM_DELETE_AT As String = "confirm.delete_at"
Public Const KEY_CONFIRM_BUTTON As String = "confirm.button"
' The modal confirmation of a destructive action (§6.6, 24.09.2026).
Public Const KEY_CONFIRM_ASK As String = "confirm.ask"
Public Const KEY_OK_CLEARED As String = "ok.cleared"
Public Const KEY_OK_LIST As String = "ok.list"
Public Const KEY_OK_BACKUP As String = "ok.backup"

Public Const KEY_ERROR_REQUIRED As String = "error.required"
Public Const KEY_ERROR_PHONE As String = "error.phone"
Public Const KEY_ERROR_EMAIL As String = "error.email"
Public Const KEY_ERROR_PRICE As String = "error.price"
Public Const KEY_ERROR_DURATION As String = "error.duration"
Public Const KEY_ERROR_BAD_ID As String = "error.bad_id"
Public Const KEY_ERROR_NOT_FOUND As String = "error.not_found"
Public Const KEY_ERROR_IN_USE As String = "error.in_use"
Public Const KEY_ERROR_BACKUP As String = "error.backup"
Public Const KEY_ERROR_WRITE As String = "error.write"
Public Const KEY_ERROR_NONWORK As String = "error.nonwork"
Public Const KEY_ERROR_TIME As String = "error.time"
Public Const KEY_ERROR_STEP As String = "error.step"
Public Const KEY_ERROR_RANGE As String = "error.range"
Public Const KEY_ERROR_CONFLICT As String = "error.conflict"
Public Const KEY_ERROR_PERIOD As String = "error.period"

' ---- the reports (phase 5) -------------------------------------------
Public Const KEY_REPORT_TITLE As String = "report.title"
Public Const KEY_REPORT_DAYS As String = "report.heading_days"
Public Const KEY_REPORT_SERVICES As String = "report.heading_services"
Public Const KEY_REPORT_CLIENTS As String = "report.heading_clients"
Public Const KEY_REPORT_TOTAL As String = "report.total"
Public Const KEY_REPORT_RENDERED As String = "report.rendered"

Public Const KEY_BOOKING_ADDED As String = "booking.added"
Public Const KEY_BOOKING_UPDATED As String = "booking.updated"
Public Const KEY_BOOKING_DELETED As String = "booking.deleted"
Public Const KEY_BOOKING_NEW As String = "booking.mode_new"
Public Const KEY_BOOKING_EDIT As String = "booking.mode_edit"
Public Const KEY_BOOKING_READY As String = "booking.ready"
Public Const KEY_OK_PICKED As String = "ok.picked"
Public Const KEY_OK_CANCELLED As String = "ok.cancelled"
Public Const KEY_TIME_MINUTES As String = "time.minutes"
Public Const KEY_TIME_HOURS_MINUTES As String = "time.hours_minutes"

Public Const KEY_PICKER_DATE_TITLE As String = "picker.date_title"
Public Const KEY_PICKER_TIME_TITLE As String = "picker.time_title"
Public Const KEY_PICKER_DURATION_TITLE As String = "picker.duration_title"
Public Const KEY_PICKER_TODAY As String = "picker.today"
Public Const KEY_PICKER_CANCEL As String = "picker.cancel"
Public Const KEY_PICKER_OK As String = "picker.ok"
Public Const KEY_PICKER_HOURS As String = "picker.hours"
Public Const KEY_PICKER_MINUTES As String = "picker.minutes"
Public Const KEY_PICKER_BUSY As String = "picker.busy"
Public Const KEY_PICKER_CLIENT_TITLE As String = "picker.client_title"
Public Const KEY_PICKER_CLIENT_SEARCH As String = "picker.client_search"
Public Const KEY_PICKER_CLIENT_NONE As String = "picker.client_none"
Public Const KEY_OK_PICKED_CLIENT As String = "ok.picked_client"
Public Const KEY_LABEL_DATE As String = "label.date"

' Every key, as a comma separated list (Test_TextCoverage walks it).
' Three parts, joined at the end: VBE refuses one long statement (the line
' continuation limit, error 40192 0x9D00 at import - 24.09.2026), and the list
' only grows. Each part stays short, the joined value is exactly the list.
Private Const ALL_KEYS_PART1 As String = _
    "ok.ready,ok.done,ok.refreshed,ok.navigated,ok.diagnostics,ok.view,ok.cleared,ok.list," & _
    "ok.backup," & _
    "error.no_action,error.unknown_action,error.busy,error.failed,error.name_missing," & _
    "error.name_not_cell,error.contract,error.route,error.no_sheet,error.no_data," & _
    "error.read_only,error.required,error.phone,error.email,error.price,error.duration," & _
    "error.bad_id,error.not_found,error.in_use,error.backup,error.write," & _
    "error.nonwork,error.time,error.step,error.range,error.conflict," & _
    "booking.added,booking.updated,booking.deleted,booking.mode_new,booking.mode_edit," & _
    "booking.ready,time.minutes,time.hours_minutes,ok.picked,ok.cancelled,"

Private Const ALL_KEYS_PART2 As String = _
    "picker.date_title,picker.time_title,picker.duration_title,picker.today," & _
    "picker.cancel,picker.ok,picker.hours,picker.minutes,picker.busy,label.date," & _
    "warn.pending_routes,warn.empty,dash.col_time,dash.col_end,dash.col_client," & _
    "dash.col_service,dash.col_duration,dash.col_price,dash.col_edit,dash.col_delete," & _
    "dash.col_id,dash.day_header,dash.range,dash.who,dash.slot,dash.tail," & _
    "dash.more,dash.month_day,dash.tip_edit,dash.tip_delete,dash.tip_new," & _
    "dash.tip_open_day,booking.btn_cancel,booking.btn_delete,"

Private Const ALL_KEYS_PART3 As String = _
    "clients.added,clients.deleted,clients.updated,clients.title,clients.none,list.capped,clients.btn_add,clients.btn_update," & _
    "services.added,services.deleted,services.updated,services.title,services.none,services.btn_add,services.btn_update," & _
    "appointments.title,appointments.none,ok.appointments_refreshed,form.mode_edit,form.loaded," & _
    "confirm.delete_at,confirm.button,confirm.ask,about.version,about.file,about.sheets,picker.client_title,picker.client_search,picker.client_none,ok.picked_client," & _
    "error.period,report.title,report.heading_days,report.heading_services," & _
    "report.heading_clients,report.total,report.rendered,label.period_start," & _
    "label.period_end,config.applied,config.defaults,config.data_hidden," & _
    "config.data_shown,config.date_added,config.dates_cleared,error.day_order," & _
    "error.format,error.weekday,error.dates,error.schema_newer"

Public Const ALL_KEYS As String = _
    ALL_KEYS_PART1 & ALL_KEYS_PART2 & ALL_KEYS_PART3

' ---- month and weekday names ----------------------------------------
' Kept here, not taken from Format(): the name of a month must not depend on the
' regional settings of the machine.
Private Const MONTH_SHORT_NAMES As String = _
    "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"
Private Const MONTH_LONG_NAMES As String = _
    "January,February,March,April,May,June,July,August,September,October,November,December"
Private Const WEEKDAY_SHORT_NAMES As String = "Mon,Tue,Wed,Thu,Fri,Sat,Sun"

' =====================================================================
'  procedures
' =====================================================================

' The text of a key ("" for an unknown key, so a test can walk all of them).
Public Function Text(ByVal key As String) As String
    Select Case LCase$(Trim$(key))
        Case KEY_STATUS_READY
            Text = "Ready."
        Case KEY_STATUS_DONE
            Text = "Done."
        Case KEY_STATUS_REFRESHED
            Text = "Refreshed: {1} clients, {2} services, {3} appointments."
        Case KEY_STATUS_NAVIGATED
            Text = "Opened {1}."
        Case KEY_STATUS_DIAGNOSTICS
            Text = "Diagnostics written to {1}."
        Case KEY_STATUS_VIEW
            Text = "{1} view - {2}"

        Case KEY_PICKER_CLIENT_TITLE
            Text = "Find a client"
        Case KEY_PICKER_CLIENT_SEARCH
            Text = "Type any part of the name or the phone"
        Case KEY_PICKER_CLIENT_NONE
            Text = "No client matches {1}."
        Case KEY_OK_PICKED_CLIENT
            Text = "Client: {1} (id {2})."

        Case KEY_CLIENTS_UPDATED
            Text = "Client #{1} ({2}) updated."
        Case KEY_SERVICES_UPDATED
            Text = "Service #{1} ({2}) updated."
        Case KEY_CLIENTS_BTN_ADD
            Text = "Add client"
        Case KEY_CLIENTS_BTN_UPDATE
            Text = "Update client"
        Case KEY_SERVICES_BTN_ADD
            Text = "Add service"
        Case KEY_SERVICES_BTN_UPDATE
            Text = "Update service"
        Case KEY_FORM_MODE_EDIT
            Text = "Editing #{1}"
        Case KEY_FORM_LOADED
            Text = "{1} #{2} is in the form - change it and press the update button."

        Case KEY_ERROR_NO_ACTION
            Text = "No action was given."
        Case KEY_ERROR_UNKNOWN_ACTION
            Text = "'{1}' is not a known action."
        Case KEY_ERROR_BUSY
            Text = "One moment - another action is still running."
        Case KEY_ERROR_FAILED
            Text = "The action failed ({1}): {2}"
        Case KEY_ERROR_NAME_MISSING
            Text = "The cell name '{1}' is missing on {2}."
        Case KEY_ERROR_NAME_NOT_CELL
            Text = "The name '{1}' does not point to a single cell."
        Case KEY_ERROR_CONTRACT
            Text = "The workbook does not match the code: {1}"
        Case KEY_ERROR_ROUTE
            Text = "The route for '{1}' points to '{2}', which does not exist."
        Case KEY_ERROR_NO_SHEET
            Text = "There is no sheet called '{1}'."
        Case KEY_ERROR_NO_DATA
            Text = "The data is not loaded: {1}"
        Case KEY_ERROR_READ_ONLY
            Text = "This build only shows the shell; the action comes in a later phase."

        Case KEY_WARN_PENDING
            Text = "{1} cell buttons are not wired yet (a later phase)."
        Case KEY_WARN_EMPTY
            Text = "Nothing to show."

        Case KEY_DASH_COL_TIME
            Text = "Time"
        Case KEY_DASH_COL_END
            Text = "End"
        Case KEY_DASH_COL_CLIENT
            Text = "Client"
        Case KEY_DASH_COL_SERVICE
            Text = "Service"
        Case KEY_DASH_COL_DURATION
            Text = "Duration"
        Case KEY_DASH_COL_PRICE
            Text = "Price"
        Case KEY_DASH_COL_EDIT
            Text = "Edit"
        Case KEY_DASH_COL_DELETE
            Text = "Delete"
        Case KEY_DASH_COL_ID
            Text = "Id"
        Case KEY_DASH_DAY_HEADER
            Text = "{1} {2}"
        Case KEY_DASH_RANGE
            Text = "{1}-{2}"
        Case KEY_DASH_WHO
            Text = "{1} / {2}"
        Case KEY_DASH_SLOT
            Text = "{1} {2} (#{3})"
        Case KEY_DASH_TAIL
            Text = ContinuationMark() & " #{1}"
        Case KEY_DASH_MORE
            Text = "+{1} more"
        Case KEY_DASH_MONTH_DAY
            Text = "{1}: {2} / {3}"
        Case KEY_DASH_TIP_EDIT
            Text = "Edit #{1}: {2}"
        Case KEY_DASH_TIP_DELETE
            Text = "Delete #{1}: {2}"
        Case KEY_DASH_TIP_NEW
            Text = "Book a new appointment on {1} at {2}"
        Case KEY_DASH_TIP_OPEN_DAY
            Text = "Show {1}"
        Case KEY_BTN_CANCEL
            Text = "Cancel"
        Case KEY_BTN_DELETE
            Text = "Delete"

        Case KEY_ABOUT_VERSION
            Text = "Version {1}"
        Case KEY_ABOUT_FILE
            Text = "{1}"
        Case KEY_ABOUT_SHEETS
            Text = "{1} sheets"

        Case KEY_CLIENTS_ADDED
            Text = "Client #{1} ({2}) was added."
        Case KEY_CLIENTS_DELETED
            Text = "Client #{1} ({2}) was deleted."
        Case KEY_CLIENTS_TITLE
            Text = "Clients ({1})"
        Case KEY_CLIENTS_NONE
            Text = "No clients yet - fill the form and press Add client."
        Case KEY_LIST_CAPPED
            Text = "{1} - only the first {2} rows are shown"
        Case KEY_SERVICES_ADDED
            Text = "Service #{1} ({2}) was added."
        Case KEY_SERVICES_DELETED
            Text = "Service #{1} ({2}) was deleted."
        Case KEY_SERVICES_TITLE
            Text = "Services ({1})"
        Case KEY_SERVICES_NONE
            Text = "No services yet - fill the form and press Add service."
        Case KEY_APPOINTMENTS_TITLE
            Text = "Appointments ({1})"
        Case KEY_APPOINTMENTS_NONE
            Text = "No appointments have been booked yet."
        Case KEY_STATUS_APPOINTMENTS_REFRESHED
            Text = "Refreshed {1} appointments."
        Case KEY_CONFIRM_ASK
            Text = "Are you sure?"
        Case KEY_CONFIRM_DELETE_AT
            Text = "Click again to delete #{1}: {2}."
        Case KEY_CONFIRM_BUTTON
            Text = "Confirm delete #{1}"
        Case KEY_OK_CLEARED
            Text = "The form was cleared."
        Case KEY_OK_LIST
            Text = "Listed {1} rows."
        Case KEY_OK_BACKUP
            Text = "Backup written: {1}"

        Case KEY_ERROR_REQUIRED
            Text = "The {1} is required."
        Case KEY_ERROR_PHONE
            Text = "The phone number needs at least {1} digits."
        Case KEY_ERROR_EMAIL
            Text = "The e-mail '{1}' has no @."
        Case KEY_ERROR_PRICE
            Text = "The price '{1}' is not an amount."
        Case KEY_ERROR_DURATION
            Text = "The duration '{1}' must be a positive multiple of {2} minutes."
        Case KEY_ERROR_BAD_ID
            Text = "The id '{1}' is not valid."
        Case KEY_ERROR_NOT_FOUND
            Text = "{1} #{2} does not exist."
        Case KEY_ERROR_IN_USE
            Text = "{1} #{2} is used by {3} appointment(s) - remove them first."
        Case KEY_ERROR_BACKUP
            Text = "The backup could not be written ({1}), so nothing was deleted."
        Case KEY_ERROR_WRITE
            Text = "The change was not written: {1}"
        Case KEY_ERROR_NONWORK
            Text = "{1} is not a working day."
        Case KEY_ERROR_TIME
            Text = "The {1} must be a time like 09:30."
        Case KEY_ERROR_STEP
            Text = "The start time must sit on a {1} minute step."
        Case KEY_ERROR_RANGE
            Text = "The appointment must fit between {1} and {2}."
        Case KEY_ERROR_CONFLICT
            Text = "{1} appointment(s) already overlap that time (the first starts at {2})."

        Case KEY_BOOKING_ADDED
            Text = "Appointment #{1} for {2} on {3} at {4} was added."
        Case KEY_BOOKING_UPDATED
            Text = "Appointment #{1} was updated: {2} on {3} at {4}."
        Case KEY_BOOKING_DELETED
            Text = "Appointment #{1} was deleted."
        Case KEY_BOOKING_NEW
            Text = "New appointment"
        Case KEY_BOOKING_EDIT
            Text = "Editing #{1}"
        Case KEY_BOOKING_READY
            Text = "Ready for a new appointment."
        Case KEY_OK_PICKED
            Text = "{1} is now {2}."
        Case KEY_OK_CANCELLED
            Text = "Cancelled - nothing changed."
        Case KEY_TIME_MINUTES
            Text = "{1} min"
        Case KEY_TIME_HOURS_MINUTES
            Text = "{1} h {2} min"

        Case KEY_PICKER_DATE_TITLE
            Text = "Pick a date"
        Case KEY_PICKER_TIME_TITLE
            Text = "Start time"
        Case KEY_PICKER_DURATION_TITLE
            Text = "Duration"
        Case KEY_PICKER_TODAY
            Text = "Today"
        Case KEY_PICKER_CANCEL
            Text = "Cancel"
        Case KEY_PICKER_OK
            Text = "OK"
        Case KEY_PICKER_HOURS
            Text = "Hour"
        Case KEY_PICKER_MINUTES
            Text = "Minutes"
        Case KEY_PICKER_BUSY
            Text = "That time is taken."
        Case KEY_LABEL_DATE
            Text = "The date"
        Case KEY_LABEL_PERIOD_START
            Text = "the start of the period"
        Case KEY_LABEL_PERIOD_END
            Text = "the end of the period"

        Case KEY_CONFIG_APPLIED
            Text = "Settings applied: {1} to {2}, step {3} min."
        Case KEY_CONFIG_DEFAULTS
            Text = "Settings are back to the built in values."
        Case KEY_CONFIG_DATA_HIDDEN
            Text = "The data sheets are hidden. Unlock them if you want to edit by hand."
        Case KEY_CONFIG_DATA_SHOWN
            Text = "The data sheets are visible. Lock them again before you hand the file over."
        Case KEY_CONFIG_DATE_ADDED
            Text = "Added {1} to the non working dates. Press Apply settings to store the page."
        Case KEY_CONFIG_DATES_CLEARED
            Text = "The non working dates are cleared. Press Apply settings to store the page."
        Case KEY_ERROR_DAY_ORDER
            Text = "The day must open before it closes (from {1} to {2})."
        Case KEY_ERROR_FORMAT
            Text = "The number format '{1}' has no 0 or # in it."
        Case KEY_ERROR_WEEKDAY
            Text = "The non working days are numbers from 1 (Monday) to 7 (Sunday); '{1}' is not."
        Case KEY_ERROR_DATES
            Text = "'{1}' is not a date. Write the days as 25.09.2026 or 2026-09-25."
        Case KEY_ERROR_SCHEMA_NEWER
            Text = "This file was written by a newer version of the application (schema {1}); this build knows {2}. Open it with the newer version instead."
            Text = "The non working dates are real dates like 25.09.2026; '{1}' is not."

        Case KEY_ERROR_PERIOD
            Text = "The period must start before it ends (from {1} to {2})."

        Case KEY_REPORT_TITLE
            Text = "Report {1} - {2}"
        Case KEY_REPORT_DAYS
            Text = "Days"
        Case KEY_REPORT_SERVICES
            Text = "Services"
        Case KEY_REPORT_CLIENTS
            Text = "Clients"
        Case KEY_REPORT_TOTAL
            Text = "Total"
        Case KEY_REPORT_RENDERED
            Text = "Reported {1} rows for {2} - {3}."

        Case Else
            Text = vbNullString
    End Select
End Function

' The text of a key with {1} / {2} / {3} / {4} replaced.
Public Function TextFormat(ByVal key As String, ByVal arg1 As String, _
                           Optional ByVal arg2 As String = vbNullString, _
                           Optional ByVal arg3 As String = vbNullString, _
                           Optional ByVal arg4 As String = vbNullString) As String
    Dim message As String

    message = Text(key)
    message = Replace(message, "{1}", arg1)
    message = Replace(message, "{2}", arg2)
    message = Replace(message, "{3}", arg3)
    message = Replace(message, "{4}", arg4)

    TextFormat = message
End Function

' The mark in front of a continuation cell of the weekly grid ("this cell is
' still the appointment of the row above"). ChrW keeps the source ASCII, which
' the harness requires (§13).
Public Function ContinuationMark() As String
    ContinuationMark = ChrW(&H21B3)
End Function

Public Function MonthShortName(ByVal monthNumber As Long) As String
    MonthShortName = NameAt(MONTH_SHORT_NAMES, monthNumber)
End Function

Public Function MonthLongName(ByVal monthNumber As Long) As String
    MonthLongName = NameAt(MONTH_LONG_NAMES, monthNumber)
End Function

' 1 = Monday ... 7 = Sunday.
Public Function WeekdayShortName(ByVal weekdayNumber As Long) As String
    WeekdayShortName = NameAt(WEEKDAY_SHORT_NAMES, weekdayNumber)
End Function

' The n-th entry of a comma separated list, "" when it does not exist.
Private Function NameAt(ByVal listText As String, ByVal index As Long) As String
    Dim parts As Variant

    If index < 1 Then Exit Function

    parts = Split(listText, ",")
    If index > UBound(parts) + 1 Then Exit Function

    NameAt = CStr(parts(index - 1))
End Function
