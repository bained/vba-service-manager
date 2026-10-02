Option Explicit

' =====================================================================
'  Service_Manager - the tests of the model
'
'      cscript //nologo tools\sim\run.wsf
'
'  The cases cover the same rules the VBA batch checks (build\test-report.txt),
'  but here they run in milliseconds, without Excel - that is the point of the
'  model. The report goes to build\sim-test-report.txt.
' =====================================================================

Dim gPassed, gFailed, gLog

Sub AddLine(text)
    gLog = gLog & text & vbCrLf
    WScript.Echo text
End Sub

Sub CheckText(caseName, actual, expected)
    If CStr(actual) = CStr(expected) Then
        gPassed = gPassed + 1
    Else
        gFailed = gFailed + 1
        AddLine "FAIL " & caseName & " (expected '" & CStr(expected) & "', actual '" & CStr(actual) & "')"
    End If
End Sub

Sub CheckNumber(caseName, actual, expected)
    If CDbl(actual) = CDbl(expected) Then
        gPassed = gPassed + 1
    Else
        gFailed = gFailed + 1
        AddLine "FAIL " & caseName & " (expected " & CStr(expected) & ", actual " & CStr(actual) & ")"
    End If
End Sub

Sub CheckTrue(caseName, condition)
    If condition Then
        gPassed = gPassed + 1
    Else
        gFailed = gFailed + 1
        AddLine "FAIL " & caseName
    End If
End Sub

' A booking as the form would carry it (the same keys the payload of
' appointments.save has).
Function BookingValues(clientId, serviceId, dateText, startText, durationText)
    Dim values
    Set values = CreateObject("Scripting.Dictionary")
    values("client_id") = clientId
    values("service_id") = serviceId
    values("date") = dateText
    values("start") = startText
    values("duration") = durationText
    Set BookingValues = values
End Function

Sub RunModelTests()

    Dim plan, book, cell, values, result, i

    gPassed = 0
    gFailed = 0
    gLog = ""

    ModelInit
    Set book = ModelBook()

    ' ---- the workbook in memory ---------------------------------------
    AddLine "the model: " & CStr(book.Sheet(SHEET_CONFIG).NameCount()) & " Config cells, " & _
            CStr(book.Sheet(SHEET_BOOKING).NameCount()) & " Booking cells, " & _
            CStr(ModelTable(SHEET_DB_APPOINTMENTS).RowCount()) & " appointments"

    ' a named cell is read and written like in Excel
    CheckTrue "the contract brought the names of the template", _
              UBound(ContractNames()) + 1 >= 100
    CheckNumber "cal_first is row 10 (the body of the Dashboard)", _
                book.RangeOfName("cal_first").Row, 10
    CheckNumber "s_view sits in column 12 of the Dashboard", _
                book.RangeOfName("s_view").Column, 12
    CheckTrue "the model has no name that the template does not have", _
              book.RangeOfName("no_such_name") Is Nothing
    CheckText "s_edit_id starts empty", book.ReadName("s_edit_id"), ""
    book.WriteName "s_edit_id", "7"
    CheckText "s_edit_id after the write", book.ReadName("s_edit_id"), "7"
    book.WriteName "s_edit_id", ""

    ' a cell button: caption + hyperlink (what modUi.WriteLinkCell writes)
    Set cell = book.RangeOfName("b_booking_cancel")
    cell.WriteLink "Delete", "b_booking_cancel", "Delete #7"
    CheckTrue "the button is a link", cell.Linked
    CheckText "the button keeps its anchor", cell.Link, "b_booking_cancel"

    ' ---- the time helpers ---------------------------------------------
    CheckNumber "TimeToMinutes(09:30)", TimeToMinutes("09:30"), 570
    CheckText "MinutesToTime(570)", MinutesToTime(570), "09:30"
    CheckNumber "TimeToMinutes(24:00) is refused", TimeToMinutes("24:00"), -1
    CheckNumber "TimeToMinutes(nonsense) is refused", TimeToMinutes("nonsense"), -1
    CheckText "MinutesToTime(-1) is empty", MinutesToTime(-1), ""
    CheckText "EndTime(09:00, 60)", EndTime("09:00", 60), "10:00"
    CheckText "EndTime without a start", EndTime("", 30), ""
    CheckText "EndTime without a duration", EndTime("09:00", 0), ""
    CheckTrue "touching intervals do not clash", Not Overlaps(540, 60, 600, 30)
    CheckTrue "a real overlap clashes", Overlaps(540, 60, 570, 30)
    CheckTrue "a zero length never clashes", Not Overlaps(540, 0, 540, 30)
    CheckTrue "555 sits on the 5 minute step", IsOnStep(555, 5)
    CheckTrue "547 does not sit on the 5 minute step", Not IsOnStep(547, 5)
    CheckTrue "a step of 0 means no rule", IsOnStep(547, 0)

    ' ---- dates ---------------------------------------------------------
    CheckTrue "2026-09-25 is an ISO date", IsIsoDate("2026-09-25")
    CheckTrue "2026-02-30 is not", Not IsIsoDate("2026-02-30")
    CheckText "the screen shows Europe", DateFormatText("2026-09-25"), "25.09.2026"
    CheckText "25.09.2026 is read", DateTextToIso("25.09.2026"), "2026-09-25"
    CheckText "25/09/2026 is read too", DateTextToIso("25/09/2026"), "2026-09-25"
    CheckText "25.9.26 counts as 2026", DateTextToIso("25.9.26"), "2026-09-25"
    CheckText "nonsense is not a date", DateTextToIso("nonsense"), ""
    CheckTrue "Monday is a working day", IsWorkingDay("2026-09-21")
    CheckTrue "Saturday is not", Not IsWorkingDay("2026-09-19")
    CheckTrue "a non working date of Config is not", Not IsWorkingDay("2026-09-19")
    CheckTrue "text that is not a date is not a working day", Not IsWorkingDay("garbage")

    ' ---- the tables ----------------------------------------------------
    CheckNumber "the appointments of 21.09.2026", AppointmentsOn("2026-09-21"), 1
    CheckNumber "the appointments of 25.09.2026", AppointmentsOn("2026-09-25"), 0
    CheckNumber "the revenue of 21.09.2026", RevenueOn("2026-09-21"), 25
    CheckNumber "the revenue of 22.09.2026", RevenueOn("2026-09-22"), 60
    CheckNumber "the revenue of the month so far", RevenueInPeriod("2026-09-01", "2026-09-30"), 115
    CheckText "client 1", ClientName(1), "Ivan Petrov"
    CheckText "service 3", ServiceName(3), "Manicure"
    CheckNumber "the price of service 3", ServicePrice(3), 30
    CheckNumber "the duration of service 3", ServiceDuration(3), 45
    CheckTrue "client 9 does not exist", Not RowExists(SHEET_DB_CLIENTS, 9)
    CheckText "the rows of a day are sorted", AppointmentRowsOn("2026-09-21"), "1"

    ' ---- the rules of a booking ----------------------------------------
    Set values = BookingValues(1, 1, "2026-09-25", "15:00", "45")
    CheckText "a fine booking has no problem", RuleCode(values), ""
    Set values = BookingValues(0, 1, "2026-09-25", "15:00", "45")
    CheckText "a missing client", RuleCode(values), "error.required|client"
    Set values = BookingValues(9, 1, "2026-09-25", "15:00", "45")
    CheckText "an unknown client", RuleCode(values), "error.not_found|Client|9"
    Set values = BookingValues(1, 0, "2026-09-25", "15:00", "45")
    CheckText "a missing service", RuleCode(values), "error.required|service"
    Set values = BookingValues(1, 1, "nonsense", "15:00", "45")
    CheckText "a date that is no date", RuleCode(values), "error.required|date"
    Set values = BookingValues(1, 1, "2026-09-19", "15:00", "45")
    CheckText "a Saturday", RuleCode(values), "error.nonwork|2026-09-19"
    Set values = BookingValues(1, 1, "2026-09-25", "nonsense", "45")
    CheckText "a start that is no time", RuleCode(values), "error.time|start time"
    Set values = BookingValues(1, 1, "2026-09-25", "15:00", "")
    CheckText "a missing duration", RuleCode(values), "error.required|duration"
    Set values = BookingValues(1, 1, "2026-09-25", "15:07", "45")
    CheckText "a start off the step", RuleCode(values), "error.step|5"
    Set values = BookingValues(1, 1, "2026-09-25", "08:00", "45")
    CheckText "a start before the working day", RuleCode(values), _
              "error.range|09:00|19:00"
    Set values = BookingValues(1, 1, "2026-09-25", "18:45", "60")
    CheckText "the end leaves the day", RuleCode(values), _
              "error.range|09:00|19:00"
    Set values = BookingValues(1, 1, "2026-09-21", "09:15", "45")
    CheckText "a clash with the sample appointment", RuleCode(values), _
              "error.conflict|1|09:00"
    Set values = BookingValues(1, 1, "2026-09-21", "09:30", "45")
    CheckText "touching the sample appointment is fine", RuleCode(values), ""

    ' ---- what the user reads -------------------------------------------
    CheckText "the label of appointment 2", AppointmentLabel(2), _
              "Maria Ivanova - Hair coloring, Tue 22.09.2026 10:00"
    CheckText "45 minutes read as", DurationText(45), "45 min"
    CheckText "90 minutes read as", DurationText(90), "1 h 30 min"
    CheckText "the busy intervals of 21.09", BusyIntervalsOn("2026-09-21", 0), "09:00+30"
    CheckText "the busy intervals without row 1", BusyIntervalsOn("2026-09-21", 1), ""
    CheckNumber "the free slots of 21.09 (5 minute step)", FreeSlotsOn("2026-09-21"), 119
    CheckText "the first clash of 09:15", FirstConflictStart("2026-09-21", 555, 45, 0), "09:00"
    CheckNumber "the clashing rows of 09:15", ConflictCount("2026-09-21", 555, 45, 0), 1

    ' ---- the daily grid (the decision layer) ---------------------------
    ' 21.09.2026 has one 30 minute appointment at 09:00: at a 5 minute
    ' resolution that is one start row, five continuation rows and 114 free
    ' ones - exactly what modDashboard.RenderDaily draws in Excel.
    plan = DailyPlan("2026-09-21")
    CheckNumber "the day has 120 slot rows", UBound(plan) + 1, 120
    CheckText "the first slot starts the appointment", plan(0), _
              "09:00|start|1|Ivan Petrov|Haircut"
    CheckText "the slot after it continues it", plan(1), "09:05|tail|1"
    CheckText "the last continuation slot", plan(5), "09:25|tail|1"
    CheckText "and then the day is free again", plan(6), "09:30|free"
    CheckText "the last slot of the day", plan(119), "18:55|free"

    ' the resolution is a setting: a 30 minute step redraws the same day as 20
    ' slots, and 09:07 is no longer a legal start
    book.WriteName "i_config_step", "30"
    CheckNumber "the resolution is 30 minutes now", StepMinutes(), 30
    plan = DailyPlan("2026-09-21")
    CheckNumber "the day has 20 slot rows", UBound(plan) + 1, 20
    CheckText "the appointment is still the first slot", plan(0), _
              "09:00|start|1|Ivan Petrov|Haircut"
    CheckText "30 minutes later the day is free", plan(1), "09:30|free"
    Set values = BookingValues(1, 1, "2026-09-25", "15:07", "45")
    CheckText "09:07 is off the 30 minute step", RuleCode(values), "error.step|30"
    book.WriteName "i_config_step", "5"

    ' ---- a booking through the form, the rules, the table and the grid --
    ' The whole flow without Excel: fill the named cells of the form, read the
    ' payload, let the rules decide, write the row, look at the day and the grid.
    ModelReloadTables

    book.WriteName "i_booking_client", "Ivan Petrov"
    book.WriteName "i_booking_service", "Manicure"
    book.WriteName "i_booking_date", "25.09.2026"
    book.WriteName "i_booking_start_h", "15"
    book.WriteName "i_booking_start_m", "00"
    book.WriteName "i_booking_dur_h", "00"
    book.WriteName "i_booking_dur_m", "45"

    Set values = ValuesFromForm(book)
    CheckText "the form finds the client by name", values("client_id"), "1"
    CheckText "and the service", values("service_id"), "3"
    CheckText "the duration in minutes", values("duration"), "45"
    CheckText "the form normalised the date it shows", values("date"), "2026-09-25"

    Set result = SaveBookingResult(values)
    CheckTrue "the booking was saved", result("ok")
    CheckNumber "it got the next id", result("id"), 4
    CheckText "the date went into the table as ISO", AppointmentCell(4, "date"), "2026-09-25"
    CheckText "the start as HH:MM", AppointmentCell(4, "start"), "15:00"
    CheckText "the duration as minutes", AppointmentCell(4, "duration"), "45"
    CheckNumber "the day has one appointment now", AppointmentsOn("2026-09-25"), 1
    CheckNumber "and its revenue", RevenueOn("2026-09-25"), 30
    CheckNumber "the id of the row of 15:00", AppointmentIdAt("2026-09-25", "15:00"), 4
    CheckText "the label of the new appointment", AppointmentLabel(4), _
              "Ivan Petrov - Manicure, Fri 25.09.2026 15:00"

    plan = DailyPlan("2026-09-25")
    CheckText "the grid starts the appointment at 15:00", plan(72), _
              "15:00|start|4|Ivan Petrov|Manicure"
    CheckText "and continues it", plan(73), "15:05|tail|4"
    CheckText "the slot after it is free", plan(81), "15:45|free"

    ' the same slot again: the rules refuse it and nothing is written
    book.WriteName "i_booking_dur_m", "30"
    Set values = ValuesFromForm(book)
    Set result = SaveBookingResult(values)
    CheckTrue "a second booking at 15:00 is refused", Not result("ok")
    CheckText "and the clash is named", result("rule"), "error.conflict|1|15:00"
    CheckNumber "the table did not grow", ModelTable(SHEET_DB_APPOINTMENTS).RowCount(), 4

    ' editing: s_edit_id is what makes the form an edit (the same cell the
    ' Cancel/Delete button of ??7.5 uses)
    book.WriteName "s_edit_id", "4"
    book.WriteName "i_booking_start_h", "16"
    book.WriteName "i_booking_start_m", "00"
    book.WriteName "i_booking_dur_m", "45"
    Set values = ValuesFromForm(book)
    CheckText "the form carries the id", PayloadText(values, "id"), "4"
    Set result = SaveBookingResult(values)
    CheckTrue "the edit was saved", result("ok")
    CheckNumber "it stayed the same row", result("id"), 4
    CheckNumber "and the table did not grow", ModelTable(SHEET_DB_APPOINTMENTS).RowCount(), 4
    CheckText "the row moved to 16:00", AppointmentCell(4, "start"), "16:00"
    plan = DailyPlan("2026-09-25")
    CheckText "15:00 is free again", plan(72), "15:00|free"
    CheckText "and 16:00 carries the appointment", plan(84), _
              "16:00|start|4|Ivan Petrov|Manicure"

    ' an edit of an appointment that is gone is refused
    book.WriteName "s_edit_id", "9"
    Set result = SaveBookingResult(ValuesFromForm(book))
    CheckText "an unknown id is named", result("rule"), "error.not_found|Appointment|9"
    book.WriteName "s_edit_id", ""

    ' the delete flow of the weekly view: the form that edited the appointment
    ' goes back to "new appointment" (the Cancel button became Delete, ??7.5)
    book.WriteName "s_edit_id", "4"
    CheckTrue "the delete worked", DeleteBookingById(book, 4)
    CheckText "the form forgot the appointment", book.ReadName("s_edit_id"), ""
    CheckTrue "deleting it again does nothing", Not DeleteBookingById(book, 4)
    CheckNumber "the day is empty again", AppointmentsOn("2026-09-25"), 0
    CheckNumber "and so is its revenue", RevenueOn("2026-09-25"), 0
    CheckNumber "the id is free for the next booking", _
                NextId(ModelTable(SHEET_DB_APPOINTMENTS)), 4

    ' deleting a row moves the rows below it up, exactly like Excel
    CheckNumber "the first sample appointment is deleted", _
                DeleteRowById(SHEET_DB_APPOINTMENTS, 1), 1
    CheckText "its day is empty", AppointmentRowsOn("2026-09-21"), ""
    CheckNumber "the appointment of 22.09 is still there", AppointmentsOn("2026-09-22"), 1
    CheckNumber "with its revenue", RevenueOn("2026-09-22"), 60
    CheckNumber "the appointments of client 2", CountAppointmentsFor("client_id", 2), 1
    CheckNumber "an id that is gone deletes nothing", _
                DeleteRowById(SHEET_DB_APPOINTMENTS, 99), 0

    ' and the fixtures are back for the next case
    ModelReloadTables
    CheckNumber "the fixtures are back", ModelTable(SHEET_DB_APPOINTMENTS).RowCount(), 3
    CheckNumber "with their first appointment", AppointmentsOn("2026-09-21"), 1

    ' ---- the reports (phase 5) ------------------------------------------
    ' The period: the cells of the sheet win, an empty pair means the current month
    ' (§16.1). The fixtures live in September, so the explicit periods below use it.
    book.WriteName "i_report_from", ""
    book.WriteName "i_report_to", ""
    CheckText "an empty period is the current month", ReportFromIso(), MonthFirstIso()
    CheckText "and it ends with its last day", ReportToIso(), MonthLastIso()

    book.WriteName "i_report_from", "01.09.2026"
    book.WriteName "i_report_to", "30.09.2026"
    CheckText "the cells decide the period", ReportFromIso(), "2026-09-01"
    CheckText "and both ends are read", ReportToIso(), "2026-09-30"

    CheckText "a period that is fine", PeriodRuleCode("2026-09-01", "2026-09-30"), ""
    CheckText "a missing start", PeriodRuleCode("", "2026-09-30"), "error.required|period start"
    CheckText "a missing end", PeriodRuleCode("2026-09-01", "garbage"), "error.required|period end"
    CheckText "the ends the wrong way round", PeriodRuleCode("2026-10-01", "2026-09-30"), _
              "error.period|01.10.2026|30.09.2026"

    CheckText "a month back", ShiftPeriodIso("2026-09-30", -1), "2026-08-30"
    CheckText "a month forward over the year", ShiftPeriodIso("2026-12-15", 1), "2027-01-15"
    CheckText "a day the target month does not have", ShiftPeriodIso("2026-03-31", -1), "2026-02-28"
    CheckText "nothing moved", ShiftPeriodIso("2026-09-30", 0), "2026-09-30"

    CheckNumber "the period holds the fixtures", AppointmentsInPeriod("2026-09-01", "2026-09-30"), 3
    CheckNumber "BookingsInPeriod answers the same", BookingsInPeriod("2026-09-01", "2026-09-30"), 3
    CheckNumber "a single day", AppointmentsInPeriod("2026-09-21", "2026-09-21"), 1
    CheckNumber "the days before the fixtures", AppointmentsInPeriod("2026-09-01", "2026-09-20"), 0

    ' the whole report: the days of September, the services, the clients, the total
    CheckText "the report of September", ReportRows("2026-09-01", "2026-09-30"), _
              "Days" & vbCrLf & _
              "21.09.2026|1|2500" & vbCrLf & _
              "22.09.2026|1|6000" & vbCrLf & _
              "24.09.2026|1|3000" & vbCrLf & _
              "" & vbCrLf & _
              "Services" & vbCrLf & _
              "Hair coloring|1|6000" & vbCrLf & _
              "Manicure|1|3000" & vbCrLf & _
              "Haircut|1|2500" & vbCrLf & _
              "" & vbCrLf & _
              "Clients" & vbCrLf & _
              "Maria Ivanova|1|6000" & vbCrLf & _
              "Ivan Petrov|2|5500" & vbCrLf & _
              "Total|3|11500"
    CheckNumber "the report has 14 rows", ReportLineCount(), 14
    CheckTrue "and it fits in the grid", Not ReportIsCapped()

    ' an empty period is not an error: the three headings and a total of nothing
    CheckText "nothing in the first days of September", ReportRows("2026-09-01", "2026-09-05"), _
              "Days" & vbCrLf & "" & vbCrLf & "Services" & vbCrLf & "" & vbCrLf & _
              "Clients" & vbCrLf & "Total|0|0"
    CheckNumber "six rows of nothing", ReportLineCount(), 6

    ' the grid of the template holds 50 rows and a longer report is capped (§6.9)
    Dim reportDay, reportText
    CheckNumber "the grid holds 50 rows", ReportGridCapacity(), 50
    For i = 0 To 59
        Set values = BookingValues(1, 1, DateValueIso(DateAdd("d", i, DateSerial(2026, 1, 1))), _
                                   "09:00", "30")
        AppendAppointment values
    Next
    reportText = ReportRows("2026-01-01", "2026-12-31")
    CheckText "a long report still starts with the days", Left(reportText, 4), "Days"
    CheckNumber "it stops at the grid", ReportLineCount(), 50
    CheckTrue "and it is marked as capped", ReportIsCapped()
    ModelReloadTables
    CheckNumber "the fixtures are back", ModelTable(SHEET_DB_APPOINTMENTS).RowCount(), 3

    ' ---- the Config page (phase 6) ---------------------------------------
    ' The built in values (§16.1) and the rules of the page. The tests below write
    ' the config cells, so they come last: nothing after them needs the fixtures.
    Set values = ConfigDefaultValues()
    CheckText "the default opening time", values(NAME_CONFIG_DAY_START), "09:00"
    CheckText "the default closing time", values(NAME_CONFIG_DAY_END), "19:00"
    CheckText "the default step", values(NAME_CONFIG_STEP), "5"
    CheckText "the default duration", values(NAME_CONFIG_DURATION), "10"
    CheckText "the default currency", values(NAME_CONFIG_CURRENCY), "EUR"
    CheckText "the default number format", values(NAME_CONFIG_NUMBER_FORMAT), "#,##0.00"
    CheckText "the default non working weekdays", values(NAME_CONFIG_NONWORK_DAYS), "6,7"
    CheckText "and no non working dates", values(NAME_CONFIG_NONWORK_DATES), ""

    ' a set the user typed: the code is empty and the normalised form is canonical
    book.WriteName NAME_CONFIG_DAY_START, "8:30"
    book.WriteName NAME_CONFIG_DAY_END, "18:00"
    book.WriteName NAME_CONFIG_STEP, "015"
    book.WriteName NAME_CONFIG_DURATION, "20"
    book.WriteName NAME_CONFIG_CURRENCY, "eur"
    book.WriteName NAME_CONFIG_NONWORK_DAYS, " 6 , 7 "
    book.WriteName NAME_CONFIG_NONWORK_DATES, "25.09.2026, 2026-10-05"

    Set values = ConfigValues()
    CheckText "a set that is fine has no rule", ConfigRuleCode(values), ""
    ConfigNormalise values
    CheckText "the opening time is canonical", values(NAME_CONFIG_DAY_START), "08:30"
    CheckText "the step lost its leading zero", values(NAME_CONFIG_STEP), "15"
    CheckText "the duration is a number", values(NAME_CONFIG_DURATION), "20"
    CheckText "the currency went up", values(NAME_CONFIG_CURRENCY), "EUR"
    CheckText "the weekdays are trimmed", values(NAME_CONFIG_NONWORK_DAYS), "6,7"
    CheckText "and the dates are ISO", values(NAME_CONFIG_NONWORK_DATES), "2026-09-25,2026-10-05"

    ' the rest of the application reads the same cells
    CheckNumber "the grid follows the opening time", DayStartMinutes(), 510
    CheckNumber "and the closing time", DayEndMinutes(), 1080
    CheckNumber "and the step", StepMinutes(), 15
    CheckTrue "the new working day reaches the rules", IsWorkingDay("2026-10-05") = False

    ' an empty list of non working weekdays means none (the owner works on Saturdays)
    book.WriteName NAME_CONFIG_NONWORK_DAYS, ""
    book.WriteName NAME_CONFIG_NONWORK_DATES, ""
    CheckTrue "an empty list means no non working weekday", IsWorkingDay("2026-09-26") = True
    book.WriteName NAME_CONFIG_NONWORK_DAYS, "6,7"

    ' one problem at a time, each named
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_DAY_START) = ""
    CheckText "an empty opening time", ConfigRuleCode(values), "error.required|opening time"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_DAY_END) = "25:00"
    CheckText "a closing time that is not a time", ConfigRuleCode(values), "error.time|closing time"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_DAY_START) = "19:00"
    CheckText "a day that opens when it closes", ConfigRuleCode(values), "error.day_order|19:00|19:00"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_STEP) = "0"
    CheckText "a step of nothing", ConfigRuleCode(values), "error.required|slot step"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_STEP) = "five"
    CheckText "a step that is not a number", ConfigRuleCode(values), "error.required|slot step"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_DURATION) = "abc"
    CheckText "a duration that is not a number", ConfigRuleCode(values), _
              "error.required|default duration"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_CURRENCY) = ""
    CheckText "no currency", ConfigRuleCode(values), "error.required|currency"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_NUMBER_FORMAT) = "AAA"
    CheckText "a format without a digit", ConfigRuleCode(values), "error.format|AAA"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_NONWORK_DAYS) = "6,8"
    CheckText "a weekday that does not exist", ConfigRuleCode(values), "error.weekday|8"
    Set values = ConfigDefaultValues()
    values(NAME_CONFIG_NONWORK_DATES) = "2026-09-19, nonsense"
    CheckText "a date that is not a date", ConfigRuleCode(values), "error.dates|nonsense"

    ' ---- the report ----------------------------------------------------
    AddLine ""
    If gFailed = 0 Then
        AddLine "PASSED - assertions passed: " & CStr(gPassed) & ", failed: 0"
    Else
        AddLine "FAILED - assertions passed: " & CStr(gPassed) & ", failed: " & CStr(gFailed)
    End If
    WriteModelReport
End Sub

' build\sim-test-report.txt, next to the report of the Excel batch, so a build
' can compare the two implementations over the same fixtures.
Sub WriteModelReport()
    Dim fso, root, buildDir, path, file

    Set fso = CreateObject("Scripting.FileSystemObject")
    root = fso.GetParentFolderName(fso.GetParentFolderName( _
           fso.GetParentFolderName(WScript.ScriptFullName)))
    buildDir = fso.BuildPath(root, "build")
    If Not fso.FolderExists(buildDir) Then fso.CreateFolder buildDir

    path = fso.BuildPath(buildDir, "sim-test-report.txt")
    Set file = fso.CreateTextFile(path, True, True)
    file.WriteLine "Service_Manager - script model report"
    file.WriteLine "generated : " & Now
    file.WriteLine "engine    : cscript (no Excel, no workbook)"
    file.WriteLine "fixtures  : " & CStr(ModelTable(SHEET_DB_APPOINTMENTS).RowCount()) & _
                   " appointments, " & CStr(ModelTable(SHEET_DB_CLIENTS).RowCount()) & _
                   " clients, " & CStr(ModelTable(SHEET_DB_SERVICES).RowCount()) & " services"
    file.WriteLine ""
    file.WriteLine "result    : assertions passed: " & CStr(gPassed) & _
                   ", failed: " & CStr(gFailed)
    file.WriteLine ""
    file.Write gLog
    file.Close

    WScript.Echo "report: " & path
End Sub

Function TestExitCode()
    If gFailed = 0 Then
        TestExitCode = 0
    Else
        TestExitCode = 1
    End If
End Function

