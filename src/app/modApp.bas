Attribute VB_Name = "modApp"
Option Explicit

' =====================================================================
'  Service_Manager - the entry point (Phase 1)
'
'  AppRouter(action, data) is the only way into the application: it looks the
'  action up in modRoutes, calls the procedure, writes a log row, keeps the
'  status cell in sync and never lets an error escape (a modal dialog would
'  block a headless run).
'
'  The event handlers of the sheets end here as well: OnHyperlink accepts the
'  click of a cell button, OnSelection and OnDoubleClick the clicks inside a
'  sheet (the pickers and the dynamic cells of the dashboard body).
' =====================================================================

Private mBusy As Boolean
Private mSheetEvents As Collection
Private mAppEvents As clsAppEvents
Private mContractChecked As Boolean
Private mPendingKey As String
Private mPendingAt As Single

' The last problem a hyperlink read had ("" when the last read was fine). The
' diagnostics and the tests show it, so a broken link is not a silent event.
' Declarations live here, never after a procedure (a lesson, PROJECT_CONCEPT §13).
Public LastHyperlinkProblem As String

' =====================================================================
'  the router
' =====================================================================

Public Function AppRouter(ByVal action As String, Optional ByRef data As Variant) As Object
    Dim payload As Object
    Dim result As Object
    Dim routeText As String
    Dim actionKey As String

    Set result = NewResult(action)
    actionKey = CStr(result(modRoutes.RESULT_ACTION))

    If Len(actionKey) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_ERROR_NO_ACTION)
        Set AppRouter = result
        Exit Function
    End If

    If mBusy Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_ERROR_BUSY)
        Set AppRouter = result
        Exit Function
    End If

    Set payload = PayloadFrom(data)
    routeText = modRoutes.RouteFor(actionKey)

    If Len(routeText) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_UNKNOWN_ACTION, actionKey)
        modLog.WriteRow actionKey, False, CStr(result(modRoutes.RESULT_MESSAGE))
        Set AppRouter = result
        Exit Function
    End If

    mBusy = True
    On Error GoTo Failed

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' The data layer belongs to the contract of every action: the router loads
    ' it when it is not loaded yet, so no procedure has to ask.
    If Not modData.IsModelLoaded() Then modData.LoadData

    ' The registry is the only place that knows what the application can do; a
    ' typo in it is found by modRoutes.VerifyRoutes, not by the first click.
    Application.Run routeText, payload, result

    If Len(CStr(result(modRoutes.RESULT_MESSAGE))) = 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_STATUS_DONE)
    End If

    ' A successful action leaves the screen in step with the data (§5.5):
    ' app.refresh draws itself, every other action is followed by a render.
    If result(modRoutes.RESULT_OK) = True Then
        If StrComp(actionKey, modRoutes.ACTION_APPOINTMENTS_LIST, vbTextCompare) <> 0 Then
            RefreshAppointmentListFor actionKey
        End If
        If StrComp(actionKey, modRoutes.ACTION_APP_REFRESH, vbTextCompare) <> 0 And _
           StrComp(actionKey, modRoutes.ACTION_APPOINTMENTS_LIST, vbTextCompare) <> 0 Then
            RefreshScreens
        End If
    End If

    modLog.WriteRow actionKey, (result(modRoutes.RESULT_OK) = True), _
                CStr(result(modRoutes.RESULT_MESSAGE))
    GoTo Cleanup

Failed:
    result(modRoutes.RESULT_OK) = False
    result(modRoutes.RESULT_ERROR) = Err.Number
    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_FAILED, _
                                                          CStr(Err.Number), Err.Description)
    modLog.WriteRow actionKey, False, CStr(result(modRoutes.RESULT_MESSAGE))

Cleanup:
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    mBusy = False
    modUi.SetStatus CStr(result(modRoutes.RESULT_MESSAGE)), (result(modRoutes.RESULT_OK) = True)
    On Error GoTo 0

    Set AppRouter = result
End Function

' An empty result: ok = False until a procedure says otherwise. All five slots
' exist from the start, so a caller never has to test the shape.
Public Function NewResult(ByVal action As String) As Object
    Dim result As Object

    Set result = CreateObject("Scripting.Dictionary")
    result(modRoutes.RESULT_ACTION) = LCase$(Trim$(action))
    result(modRoutes.RESULT_OK) = False
    result(modRoutes.RESULT_MESSAGE) = vbNullString
    result(modRoutes.RESULT_ERROR) = 0
    result(modRoutes.RESULT_DATA) = vbNullString

    Set NewResult = result
End Function

' The payload of an action: the data of the caller, or an empty dictionary.
Public Function PayloadFrom(ByRef data As Variant) As Object
    Dim payload As Object

    Set payload = CreateObject("Scripting.Dictionary")

    If IsMissing(data) Then
        Set PayloadFrom = payload
        Exit Function
    End If
    If IsObject(data) Then
        If Not data Is Nothing Then Set payload = data
    End If

    Set PayloadFrom = payload
End Function

' One text of a payload ("" when the key is missing or empty).
Public Function PayloadText(ByRef data As Object, ByVal key As String) As String
    If data Is Nothing Then Exit Function
    If Not data.Exists(key) Then Exit Function

    PayloadText = Trim$(CStr(data(key)))
End Function

' True when the payload asks for a dry run (validate, do not write).
Public Function PayloadDry(ByRef data As Object) As Boolean
    PayloadDry = (StrComp(PayloadText(data, modRoutes.KEY_DRY), "yes", vbTextCompare) = 0)
End Function

' A numeric payload value (0 when the key is missing or not a number).
Public Function PayloadNumber(ByRef data As Object, ByVal key As String) As Long
    Dim text As String

    text = PayloadText(data, key)
    If Not IsNumeric(text) Then Exit Function

    PayloadNumber = CLng(Val(text))
End Function

' A number for a cell: the value when it is numeric, otherwise the fallback.
Public Function NumberValue(ByVal value As Variant, ByVal fallback As Double) As Variant
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        NumberValue = fallback
        Exit Function
    End If
    If Not IsNumeric(value) Then
        NumberValue = fallback
        Exit Function
    End If

    NumberValue = CDbl(value)
End Function

' =====================================================================
'  the confirmation of a destructive action (§6.6)
' =====================================================================

' One modal question - "Are you sure?" with OK and Cancel (24.09.2026). It replaced
' the two step confirmation (click, then click again within CONFIRM_SECONDS): a
' dialog cannot be hit by accident, which is the whole reason the second click
' existed, and it says out loud what it is about to do.
'
' The key stays in the signature so the callers did not have to change; it is no
' longer remembered anywhere. With the dialogs off (a test, the harness) the answer
' is yes at once, so a test stays deterministic.
Public Function ConfirmNow(ByVal keyText As String) As Boolean
    If Not modUi.InteractiveUi() Then
        ConfirmNow = True
        Exit Function
    End If

    ConfirmNow = modUi.AskConfirm(modText.Text(modText.KEY_CONFIRM_ASK))
End Function

' Kept for the callers that reset the confirmation state between two actions; with
' a modal dialog there is no state left to reset.
Public Sub ConfirmReset()
    mPendingKey = vbNullString
    mPendingAt = 0
End Sub

' While one action runs, the next one is refused: writing into a sheet raises
' events, and without the guard the action would call itself.
Public Function IsBusy() As Boolean
    IsBusy = mBusy
End Function

Public Sub SetBusy(ByVal value As Boolean)
    mBusy = value
End Sub

' =====================================================================
'  start and stop
' =====================================================================

' The one call Workbook_Open makes (src\doc\ThisWorkbook.txt). It attaches the
' event handlers, sets "Ready." on every page and refreshes the dashboard.
Public Sub Boot()
    mContractChecked = False
    modUi.SetInteractiveUi True

    AttachSheetEvents
    Set mAppEvents = New clsAppEvents
    Set mAppEvents.App = Application

    modUi.ResetStatuses
    AppRouter modRoutes.ACTION_APP_REFRESH
    modUi.Navigate SHEET_DASHBOARD
End Sub

' Workbook_BeforeClose: the event instances go away with the session.
Public Sub Shutdown()
    DetachSheetEvents
    Set mAppEvents = Nothing
End Sub

' One clsSheetEvents per UI sheet. The instances must live in a module variable:
' a local one would be released at the end of this procedure and the events
' would stop arriving (the trap from the concept, §6.2.2).
Public Sub AttachSheetEvents()
    Dim sheetEvents As clsSheetEvents
    Dim sheetNames As Variant
    Dim item As Variant
    Dim sheetName As String

    DetachSheetEvents
    Set mSheetEvents = New Collection

    sheetNames = Split(modConfig.UiSheetNames(), ",")
    For Each item In sheetNames
        sheetName = Trim$(CStr(item))
        If modNames.SheetExists(sheetName) Then
            Set sheetEvents = New clsSheetEvents
            Set sheetEvents.Ws = ThisWorkbook.Worksheets(sheetName)
            mSheetEvents.Add sheetEvents
        End If
    Next item
End Sub

Public Sub DetachSheetEvents()
    Dim index As Long

    If mSheetEvents Is Nothing Then Exit Sub

    For index = 1 To mSheetEvents.Count
        mSheetEvents.Item(index).Detach
    Next index

    Set mSheetEvents = Nothing
End Sub

' How many sheets are attached (a test checks that it is all of them).
Public Function AttachedSheetCount() As Long
    If mSheetEvents Is Nothing Then Exit Function

    AttachedSheetCount = mSheetEvents.Count
End Function

' =====================================================================
'  drawing the screen
' =====================================================================

' Loads the data, refreshes the technical lists and draws the sheet in front.
' The router calls it after a successful action (§5.5), app.refresh and
' app.navigate use it, so there is one render path.
Public Sub RefreshScreens()
    If Not modData.IsModelLoaded() Then modData.LoadData

    modData.RefreshLists
    RenderSheet ActiveSheetName()
End Sub

' The all-appointments list is a second view of the same records. Keep it current
' after any client, service or appointment mutation, even when another sheet is
' active; client/service changes also update the displayed contact/name columns.
Private Sub RefreshAppointmentListFor(ByVal actionKey As String)
    Select Case LCase$(actionKey)
        Case LCase$(modRoutes.ACTION_CLIENTS_ADD), LCase$(modRoutes.ACTION_CLIENTS_DELETE), _
             LCase$(modRoutes.ACTION_SERVICES_ADD), LCase$(modRoutes.ACTION_SERVICES_DELETE), _
             LCase$(modRoutes.ACTION_APPOINTMENTS_SAVE), LCase$(modRoutes.ACTION_APPOINTMENTS_DELETE), _
             LCase$(modRoutes.ACTION_BOOKING_CANCEL), LCase$(modRoutes.ACTION_APP_REFRESH)
            If modData.IsModelLoaded() And modNames.SheetExists(SHEET_APPOINTMENTS) Then
                modUiAppointments.Render
            End If
    End Select
End Sub

' Draws the sheet in front.
Public Sub RenderSheet(ByVal sheetName As String)
    If Len(sheetName) = 0 Then Exit Sub
    If Not modNames.SheetExists(sheetName) Then Exit Sub

    Select Case sheetName
        Case SHEET_DASHBOARD
            modDashboard.Render
        Case SHEET_CLIENTS_ADMIN
            modUiClients.Render
        Case SHEET_SERVICES_ADMIN
            modUiServices.Render
        Case SHEET_APPOINTMENTS
            modUiAppointments.Render
        Case SHEET_BOOKING
            modUiBooking.Render
        Case Else
            ' Nothing to draw yet.
    End Select
End Sub

' The name of the sheet in front ("" when Excel says nothing).
Public Function ActiveSheetName() As String
    Dim sheetName As String

    sheetName = vbNullString
    On Error Resume Next
    sheetName = ActiveSheet.Name
    On Error GoTo 0

    ActiveSheetName = sheetName
End Function

' A sheet came to the front by a click on its tab: draw it. The guard keeps the
' activation inside app.navigate from rendering twice (§6.7).
Public Sub OnActivate(ByVal sheetName As String)
    If IsBusy() Then Exit Sub
    If Not modData.IsModelLoaded() Then Exit Sub

    RenderSheet sheetName
End Sub

' A manual change in a cell of a UI sheet: the sheet decides what follows (the
' booking form refreshes its previews and fills the duration, §7.5, §9.4).
Public Sub OnChange(ByVal sheetName As String, ByRef target As Range)
    If IsBusy() Then Exit Sub

    Select Case sheetName
        Case SHEET_BOOKING
            modUiBooking.OnChange sheetName, target
        Case Else
            ' The lists and the dashboard have nothing to follow.
    End Select
End Sub

' =====================================================================
'  actions of the shell
' =====================================================================

' app.refresh: check the contract once, load the data, render the dashboard.
Public Sub Refresh(ByRef data As Object, ByRef result As Object)
    Dim problems As String
    Dim info As Object

    If Not mContractChecked Then
        problems = modNames.VerifyContract()
        If Len(problems) > 0 Then
            result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
                modText.KEY_ERROR_CONTRACT, OneLine(problems))
            Exit Sub
        End If
        mContractChecked = True
    End If

    ' The schema of the file: an older file is migrated once (a copy and a Log row
    ' come first), a newer one is refused instead of being written into (§8.3).
    problems = modSchema.Problem()
    If Len(problems) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = OneLine(problems)
        Exit Sub
    End If

    modData.LoadData

    problems = modData.TableProblems()
    If Len(problems) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_CONTRACT, OneLine(problems))
        Exit Sub
    End If

    ' The render path is shared with the router, so a refresh and an action
    ' draw the sheet in front the same way.
    RefreshScreens

    Set info = CreateObject("Scripting.Dictionary")
    info("view") = modDashboard.CurrentView()

    info("date") = modDashboard.CurrentDate()
    info("clients") = CStr(modData.ClientsRowCount())
    info("services") = CStr(modData.ServicesRowCount())
    info("appointments") = CStr(modData.AppointmentsRowCount())
    info("pending_buttons") = CStr(PendingButtonCount())
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_STATUS_REFRESHED, _
        CStr(modData.ClientsRowCount()), CStr(modData.ServicesRowCount()), _
        CStr(modData.AppointmentsRowCount()))
    result(modRoutes.RESULT_OK) = True
End Sub

Private Function OneLine(ByVal text As String) As String
    OneLine = Trim$(Replace(Replace(text, vbCrLf, " | "), "  ", " "))
End Function

' app.navigate: show another page.
Public Sub Navigate(ByRef data As Object, ByRef result As Object)
    Dim sheetName As String
    Dim info As Object

    sheetName = PayloadText(data, modRoutes.KEY_SHEET)
    If Len(sheetName) = 0 Then sheetName = PayloadText(data, modRoutes.KEY_VIEW)

    If Not modConfig.IsUiSheet(sheetName) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
            modText.KEY_ERROR_NO_SHEET, sheetName)
        Exit Sub
    End If

    modUi.Navigate sheetName
    RefreshScreens

    Set info = CreateObject("Scripting.Dictionary")
    info(modRoutes.KEY_SHEET) = sheetName
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_STATUS_NAVIGATED, sheetName)
    result(modRoutes.RESULT_OK) = True
End Sub

' app.about: the version page.
Public Sub ShowAbout(ByRef data As Object, ByRef result As Object)
    modNames.WriteName NAME_ABOUT_VERSION, _
        modText.TextFormat(modText.KEY_ABOUT_VERSION, APP_VERSION)
    modNames.WriteName NAME_ABOUT_FILE, _
        modText.TextFormat(modText.KEY_ABOUT_FILE, ThisWorkbook.Name)
    modNames.WriteName NAME_ABOUT_SHEETS, _
        modText.TextFormat(modText.KEY_ABOUT_SHEETS, CStr(ThisWorkbook.Worksheets.Count))

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_STATUS_NAVIGATED, SHEET_ABOUT)
    result(modRoutes.RESULT_OK) = True
End Sub

' app.diagnostics: the report the inspect mode of the harness writes.
Public Sub ActionDiagnostics(ByRef data As Object, ByRef result As Object)
    Dim targetPath As String

    targetPath = PayloadText(data, modRoutes.KEY_PATH)
    If Len(targetPath) = 0 Then targetPath = DefaultDiagnosticsPath()

    WriteDiagnostics targetPath

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_STATUS_DIAGNOSTICS, targetPath)
    result(modRoutes.RESULT_OK) = True
End Sub

Public Function DefaultDiagnosticsPath() As String
    DefaultDiagnosticsPath = ThisWorkbook.Path & Application.PathSeparator & "diagnostics.txt"
End Function

' How many cell buttons are still without an action: the work of the next
' phases, and a row of the diagnostics report. The list ends with a line break,
' so the empty entry is not counted.
Public Function PendingButtonCount() As Long
    Dim listText As String
    Dim lines As Variant
    Dim index As Long
    Dim count As Long

    listText = modRoutes.PendingButtons()
    ' Every button is wired since phase 6, so the list is normally empty - and
    ' Split("") has no elements at all, so UBound would fail on it.
    If Len(listText) = 0 Then Exit Function

    lines = Split(listText, vbCrLf)
    For index = 0 To UBound(lines)
        If Len(Trim$(CStr(lines(index)))) > 0 Then count = count + 1
    Next index

    PendingButtonCount = count
End Function

' The same helpers the domain will use: one place for money and for time.
Public Function TryMoneyValue(ByVal amountText As String, ByRef amount As Double) As Boolean
    Dim text As String
    Dim separator As Long
    Dim wholeText As String
    Dim fractionText As String

    amount = 0
    text = Trim$(amountText)
    If Len(text) = 0 Then
        TryMoneyValue = True
        Exit Function
    End If

    separator = InStr(1, text, ".", vbBinaryCompare)
    If separator = 0 Then separator = InStr(1, text, ",", vbBinaryCompare)

    If separator > 0 Then
        If InStr(separator + 1, text, ".", vbBinaryCompare) > 0 Then Exit Function
        If InStr(separator + 1, text, ",", vbBinaryCompare) > 0 Then Exit Function

        wholeText = Left$(text, separator - 1)
        fractionText = Mid$(text, separator + 1)
        If Len(fractionText) = 0 Or Len(fractionText) > 2 Then Exit Function
    Else
        wholeText = text
        fractionText = vbNullString
    End If

    If Len(wholeText) = 0 Then wholeText = "0"
    If Len(wholeText) > 9 Then Exit Function
    If Not IsDigits(wholeText) Then Exit Function
    If Len(fractionText) > 0 Then
        If Not IsDigits(fractionText) Then Exit Function
    End If

    amount = CDbl(CLng(wholeText))
    If Len(fractionText) = 1 Then fractionText = fractionText & "0"
    If Len(fractionText) = 2 Then amount = amount + CDbl(CLng(fractionText)) / 100#

    TryMoneyValue = True
End Function

Public Function IsDigits(ByVal text As String) As Boolean
    Dim index As Long

    If Len(text) = 0 Then Exit Function

    For index = 1 To Len(text)
        If InStr(1, "0123456789", Mid$(text, index, 1), vbBinaryCompare) = 0 Then Exit Function
    Next index

    IsDigits = True
End Function

' "09:30" -> 570 minutes; -1 when the text is not a time.
Public Function TimeToMinutes(ByVal text As String) As Long
    Dim parts As Variant
    Dim hours As Long
    Dim minutes As Long

    TimeToMinutes = -1
    If Len(Trim$(text)) = 0 Then Exit Function

    parts = Split(Trim$(text), ":")
    If UBound(parts) <> 1 Then Exit Function

    If Not IsNumeric(parts(0)) Then Exit Function
    If Not IsNumeric(parts(1)) Then Exit Function

    hours = CLng(Val(parts(0)))
    minutes = CLng(Val(parts(1)))
    If hours < 0 Or hours > 23 Then Exit Function
    If minutes < 0 Or minutes > 59 Then Exit Function

    TimeToMinutes = hours * MINUTES_PER_HOUR + minutes
End Function

' 570 -> "09:30".
Public Function MinutesToTime(ByVal minutes As Long) As String
    Dim hours As Long

    If minutes < 0 Then Exit Function

    hours = minutes \ MINUTES_PER_HOUR
    MinutesToTime = Format$(hours, "00") & ":" & Format$(minutes Mod MINUTES_PER_HOUR, "00")
End Function

' =====================================================================
'  diagnostics (also the inspect mode of the harness)
' =====================================================================

Public Sub WriteDiagnostics(ByVal targetPath As String)
    Dim fileNumber As Integer
    Dim routes As Collection
    Dim logLines As Collection
    Dim index As Long
    Dim problems As String
    Dim viewKey As String
    Dim anchorText As String
    Dim fromText As String
    Dim toText As String

    If Len(Trim$(targetPath)) = 0 Then
        Err.Raise 5, "modApp.WriteDiagnostics", "The target path cannot be empty."
    End If

    If Not modData.IsModelLoaded() Then modData.LoadData

    fileNumber = FreeFile
    Open targetPath For Output As #fileNumber

    Print #fileNumber, APP_NAME & " - diagnostics"
    Print #fileNumber, "generated : " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #fileNumber, "version   : " & APP_VERSION
    Print #fileNumber, "file      : " & ThisWorkbook.Name
    Print #fileNumber, "folder    : " & ThisWorkbook.Path
    Print #fileNumber, "sheets    : " & CStr(ThisWorkbook.Worksheets.Count)
    Print #fileNumber, "attached  : " & CStr(AttachedSheetCount()) & " UI sheets"

    problems = modNames.VerifyContract()
    If Len(problems) = 0 Then
        Print #fileNumber, "contract  : ok"
    Else
        Print #fileNumber, "contract  : PROBLEMS"
        Print #fileNumber, problems
    End If

    problems = modRoutes.VerifyRoutes()
    If Len(problems) = 0 Then
        Print #fileNumber, "routes    : ok"
    Else
        Print #fileNumber, "routes    : PROBLEMS"
        Print #fileNumber, problems
    End If

    Print #fileNumber, "pending   : " & CStr(PendingButtonCount()) & _
                      " cell buttons without an action"
    Print #fileNumber, modRoutes.PendingButtons()

    Print #fileNumber, "data      : clients " & CStr(modData.ClientsRowCount()) & _
                      ", services " & CStr(modData.ServicesRowCount()) & _
                      ", appointments " & CStr(modData.AppointmentsRowCount())
    problems = modData.TableProblems()
    If Len(problems) > 0 Then
        Print #fileNumber, "table     : PROBLEMS"
        Print #fileNumber, problems
    Else
        Print #fileNumber, "table     : ok"
    End If

    viewKey = modDashboard.CurrentView()
    anchorText = modDashboard.CurrentDate()
    modDashboard.PeriodBounds viewKey, anchorText, fromText, toText
    Print #fileNumber, "period    : " & viewKey & " " & anchorText & _
                      " (" & fromText & ".." & toText & ")"
    Print #fileNumber, "counters  : " & CStr(modData.AppointmentsInPeriod(fromText, toText)) & _
                      " appointments, revenue " & modDashboard.FormatMoney( _
                      modData.RevenueInPeriod(fromText, toText)) & _
                      ", free slots " & CStr(modDashboard.FreeSlotsFor(viewKey, anchorText))

    Print #fileNumber, "body      : labels row " & CStr(modDashboard.BodyHeaderRow()) & _
                      ", block " & CStr(modDashboard.BodyFirstRow()) & ":" & _
                      CStr(modDashboard.BodyLastRow()) & ", " & _
                      CStr(modDashboard.SlotRowCount()) & " slots of " & _
                      CStr(modAppointments.StepMinutes()) & " min"

    Print #fileNumber, "config    : day " & _
                      modDashboard.ConfigText(NAME_CONFIG_DAY_START, DEFAULT_DAY_START) & "-" & _
                      modDashboard.ConfigText(NAME_CONFIG_DAY_END, DEFAULT_DAY_END) & _
                      ", step " & CStr(modDashboard.ConfigNumber(NAME_CONFIG_STEP, _
                                                                 DEFAULT_STEP_MINUTES)) & _
                      ", duration " & CStr(modDashboard.ConfigNumber(NAME_CONFIG_DURATION, _
                                                                     DEFAULT_DURATION_MINUTES))

    Print #fileNumber, ""
    Print #fileNumber, "routes:"
    Set routes = modRoutes.RouteLines()
    For index = 1 To routes.Count
        Print #fileNumber, "  " & CStr(routes.Item(index))
    Next index

    Print #fileNumber, ""
    Print #fileNumber, "log (newest first):"
    Set logLines = modLog.Tail(10)
    For index = 1 To logLines.Count
        Print #fileNumber, "  " & CStr(logLines.Item(index))
    Next index

    Print #fileNumber, ""
    Print #fileNumber, "state     : finished"
    Close #fileNumber
End Sub

' =====================================================================
'  events of the sheets (called by clsSheetEvents)
' =====================================================================

' A click on a hyperlink cell: the name of the cell is the action. A rendered
' cell has no name - its hyperlink points at a grid anchor - so a click that is
' not an action is handed to the sheet, which knows what the position means
' (§6.10).
Public Sub OnHyperlink(ByVal sheetName As String, ByVal target As Variant)
    Dim cellName As String
    Dim cell As Range

    cellName = HyperlinkName(target)
    If Len(cellName) = 0 Then Exit Sub

    If modRoutes.IsClickAction(cellName) Then
        RunCellAction sheetName, cellName
        Exit Sub
    End If

    Set cell = Nothing
    On Error Resume Next
    Set cell = target.Range
    On Error GoTo 0
    If cell Is Nothing Then Exit Sub

    OnDynamicClick sheetName, cell
End Sub

' A click on a rendered cell: the sheet turns the position into an action.
Public Sub OnDynamicClick(ByVal sheetName As String, ByRef cell As Range)
    Dim payload As Object
    Dim actionText As String
    Dim result As Object

    Set payload = ClickPayloadFor(sheetName, cell)
    If payload Is Nothing Then Exit Sub

    actionText = PayloadText(payload, modRoutes.KEY_ACTION)
    If Len(actionText) = 0 Then Exit Sub
    payload.Remove modRoutes.KEY_ACTION

    Set result = AppRouter(actionText, payload)
End Sub

' The payload a click on a rendered cell means (Nothing when it means nothing).
Public Function ClickPayloadFor(ByVal sheetName As String, ByRef cell As Range) As Object
    Set ClickPayloadFor = Nothing

    Select Case sheetName
        Case SHEET_CLIENTS_ADMIN
            Set ClickPayloadFor = modUiClients.ClickPayload(cell)
        Case SHEET_SERVICES_ADMIN
            Set ClickPayloadFor = modUiServices.ClickPayload(cell)
        Case SHEET_APPOINTMENTS
            Set ClickPayloadFor = modUiAppointments.ClickPayload(cell)
        Case SHEET_DASHBOARD
            Set ClickPayloadFor = modDashboard.ClickPayload(cell)
        Case Else
            ' The dashboard body is the only dynamic block outside the two lists;
            ' nothing else answers a click on a cell without a name.
    End Select
End Function

' The action of a named cell button. A button that is not wired yet answers
' "later phase" instead of staying silent.
Public Sub RunCellAction(ByVal sheetName As String, ByVal cellName As String)
    Dim actionText As String
    Dim payload As Object
    Dim result As Object

    actionText = modRoutes.ActionForCell(cellName)

    If Len(actionText) = 0 Then
        modUi.SetStatusOn sheetName, modText.Text(modText.KEY_ERROR_READ_ONLY), False
        Exit Sub
    End If

    Set payload = modRoutes.PayloadForCell(cellName)
    Set result = AppRouter(actionText, payload)
End Sub

' A single click inside a sheet. A cell that carries an action is NOT handled
' here: the hyperlink already ran it, and running it twice would book twice.
Public Sub OnSelection(ByVal sheetName As String, ByRef target As Range)
    Dim cellName As String

    If target Is Nothing Then Exit Sub
    If target.Cells.Count > 1 Then Exit Sub

    cellName = modNames.NameOfCell(target.Worksheet, target)
    If Len(cellName) = 0 Then Exit Sub
    If modRoutes.IsClickAction(cellName) Then Exit Sub

    ' Phase 3 adds the input cells that open a picker (the date and time cells).
End Sub

' A double click behaves like a click on a cell button: a cell that has a picker
' opens it (the date of Booking, From and To of the report period) and any other
' cell that carries a named button runs that button. Without the cancel Excel would
' open the cell for typing.
Public Sub OnDoubleClick(ByVal sheetName As String, ByRef target As Range, ByRef cancel As Boolean)
    Dim cellName As String

    If target Is Nothing Then Exit Sub
    If target.Cells.Count > 1 Then Exit Sub

    ' The input cells that open a calendar: the double click is the same as pressing
    ' their Pick button, so the rule lives in one place (§7.5).
    cellName = modNames.NameOfCell(target.Worksheet, target)
    Select Case LCase$(cellName)
        Case LCase$(NAME_BOOKING_DATE)
            cellName = NAME_BTN_BOOKING_DATE_PICK
        Case LCase$(NAME_REPORT_FROM)
            cellName = NAME_BTN_REPORT_FROM_PICK
        Case LCase$(NAME_REPORT_TO)
            cellName = NAME_BTN_REPORT_TO_PICK
        Case LCase$(NAME_CONFIG_NONWORK_DATES)
            cellName = NAME_BTN_CONFIG_NONWORK_PICK

    End Select

    If Len(cellName) = 0 Then Exit Sub
    If Not modRoutes.IsClickAction(cellName) Then Exit Sub

    cancel = True
    RunCellAction sheetName, cellName
End Sub

' The name a hyperlink points at. The template anchors every button at its own
' name (SubAddress), which never moves the view (PROJECT_CONCEPT §6.2.2); an
' older, hand made link is accepted as well. Public, because the shell and the
' tests both read a link the same way.
'
' The object is copied into an Object variable before it is used: a late bound
' property read on a Variant raises "Type mismatch" (a lesson, §13). The
' procedure never raises either: it is called from a sheet event.
Public Function HyperlinkName(ByVal target As Variant) As String
    Dim link As Object
    Dim nameText As String
    Dim stepText As String

    stepText = "start"
    On Error GoTo Failed

    stepText = "not an object"
    If Not IsObject(target) Then Exit Function

    stepText = "copy the reference"
    Set link = target
    If link Is Nothing Then Exit Function

    stepText = "read SubAddress"
    nameText = Trim$(CStr(link.SubAddress))

    If Len(nameText) = 0 Then
        stepText = "read the cell name"
        nameText = modNames.NameOfCell(link.Range.Worksheet, link.Range)
    End If

    stepText = "clean the name up"
    nameText = Replace(nameText, "#", vbNullString)
    If InStr(1, nameText, "!", vbBinaryCompare) > 0 Then
        nameText = Mid$(nameText, InStrRev(nameText, "!") + 1)
    End If

    HyperlinkName = nameText
    Exit Function

Failed:
    HyperlinkName = vbNullString
    LastHyperlinkProblem = stepText & ": " & CStr(Err.Number) & " " & Err.Description
End Function
