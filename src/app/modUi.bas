Attribute VB_Name = "modUi"
Option Explicit

' =====================================================================
'  Service_Manager - the screen (Phase 1)
'
'  Everything the user sees goes through here: the status cell of a sheet, the
'  move to another sheet, the paint state of a cell button and the three
'  dialogs. The guard flag is formulated so that its default value allows the
'  dialogs: if a caller forgets to set it, the application still works instead
'  of staying silent (a lesson from the POC).
' =====================================================================

Private mDialogsSuppressed As Boolean

' A bulk write (a render) switches the per-cell toggles off for the whole block:
' EnableEvents was saved, cleared and restored for every single cell, and the
' weekly grid is ~400 cells - that alone was a third of the 0.8 s a render took
' (§6.9). Nesting is counted, so a render inside a render is safe.
Private mBulkDepth As Long
Private mScreenWasOn As Boolean
Private mEventsWereOn As Boolean
Private mCalculationWas As Long

' =====================================================================
'  bulk writes
' =====================================================================

' Everything that writes many cells at once (the dashboard body, the lists, the
' report grid) runs between these two: ScreenUpdating, EnableEvents and the
' calculation mode are set once and restored at the end.
Public Sub BeginBulk()
    If mBulkDepth = 0 Then
        mScreenWasOn = Application.ScreenUpdating
        mEventsWereOn = Application.EnableEvents
        mCalculationWas = Application.Calculation
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    End If
    mBulkDepth = mBulkDepth + 1
End Sub

Public Sub EndBulk()
    If mBulkDepth = 0 Then Exit Sub
    mBulkDepth = mBulkDepth - 1
    If mBulkDepth = 0 Then
        Application.Calculation = mCalculationWas
        Application.EnableEvents = mEventsWereOn
        Application.ScreenUpdating = mScreenWasOn
    End If
End Sub

Public Function InBulk() As Boolean
    InBulk = (mBulkDepth > 0)
End Function

' The events off for one cell - or already off, when a bulk write did it. What it
' returns is what EventsBack takes back.
Public Function EventsOff() As Boolean
    EventsOff = Application.EnableEvents
    If mBulkDepth = 0 Then Application.EnableEvents = False
End Function

Public Sub EventsBack(ByVal wasOn As Boolean)
    If mBulkDepth = 0 Then Application.EnableEvents = wasOn
End Sub

' =====================================================================
'  dialogs
' =====================================================================

Public Function InteractiveUi() As Boolean
    InteractiveUi = Not mDialogsSuppressed
End Function

Public Sub SetInteractiveUi(ByVal value As Boolean)
    mDialogsSuppressed = Not value
End Sub

' A yes/no question. Without dialogs (tests, headless runs) the answer is yes,
' so a flow that asks for confirmation stays deterministic.
Public Function Confirm(ByVal question As String) As Boolean
    If mDialogsSuppressed Then
        Confirm = True
        Exit Function
    End If

    Confirm = (MsgBox(question, vbYesNo + vbQuestion, APP_NAME) = vbYes)
End Function

Public Sub ShowError(ByVal message As String)
    If mDialogsSuppressed Then Exit Sub

    MsgBox message, vbExclamation, APP_NAME
End Sub

Public Sub ShowInfo(ByVal message As String, ByVal title As String)
    If mDialogsSuppressed Then Exit Sub

    MsgBox message, vbInformation, title
End Sub

' =====================================================================
'  the sheet
' =====================================================================

' Writes the status cell of a sheet: green for a success, red for a refusal.
Public Sub SetStatusOn(ByVal sheetName As String, ByVal message As String, ByVal ok As Boolean)
    Dim statusName As String
    Dim cell As Range

    statusName = modNames.StatusNameFor(sheetName)
    If Len(statusName) = 0 Then Exit Sub
    If Not modNames.SheetExists(sheetName) Then Exit Sub

    Set cell = modNames.RangeOfName(statusName)
    cell.NumberFormat = "@"
    cell.Value = message

    If ok Then
        cell.Interior.Color = COLOR_STATUS_OK
        cell.Font.Color = COLOR_TEXT
    Else
        cell.Interior.Color = COLOR_STATUS_ERROR
        cell.Font.Color = COLOR_WARNING_FONT
    End If
End Sub

' The status cell of the sheet that is in front (used by the router).
Public Sub SetStatus(ByVal message As String, ByVal ok As Boolean)
    Dim sheetName As String

    sheetName = vbNullString
    On Error Resume Next
    sheetName = ActiveSheet.Name
    On Error GoTo 0

    If Len(sheetName) = 0 Then Exit Sub
    SetStatusOn sheetName, message, ok
End Sub

' The status of every UI sheet back to "Ready." (the boot does it once).
Public Sub ResetStatuses()
    Dim sheetNames As Variant
    Dim item As Variant

    sheetNames = Split(modConfig.UiSheetNames(), ",")
    For Each item In sheetNames
        SetStatusOn Trim$(CStr(item)), modText.Text(modText.KEY_STATUS_READY), True
    Next item
End Sub

' Moves the view. The only place that activates a sheet on purpose.
Public Sub Navigate(ByVal sheetName As String)
    If Not modNames.SheetExists(sheetName) Then Exit Sub

    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Activate
    On Error GoTo 0
End Sub

' The paint state of a cell button: normal, "active" (the view in front),
' "pending" (waiting for a confirmation, §6.6) or "muted" (nothing to do here).
' The caption changes only when one is given.
Public Sub PaintCell(ByRef cell As Range, ByVal stateText As String, _
                     Optional ByVal caption As String = vbNullString)
    Dim eventsWereOn As Boolean

    If cell Is Nothing Then Exit Sub

    eventsWereOn = EventsOff()
    On Error Resume Next
    Select Case LCase$(Trim$(stateText))
        Case PAINT_ACTIVE
            cell.Interior.Color = COLOR_BUTTON_ACTIVE
            cell.Font.Color = COLOR_TEXT
        Case PAINT_PENDING
            cell.Interior.Color = COLOR_PENDING
            cell.Font.Color = COLOR_TEXT
        Case PAINT_MUTED
            cell.Interior.Color = COLOR_NOWORK
            cell.Font.Color = COLOR_MUTED
        Case Else
            cell.Interior.Color = COLOR_BUTTON
            cell.Font.Color = COLOR_TEXT
    End Select
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
    If Len(caption) > 0 Then cell.Value = caption
    On Error GoTo 0
    Err.Clear
    EventsBack eventsWereOn
End Sub

' The paint state of a cell button by name (used by the dashboard).
Public Sub PaintButton(ByVal sheetName As String, ByVal cellName As String, ByVal stateText As String)
    Dim cell As Range

    If Not modNames.SheetExists(sheetName) Then Exit Sub

    Set cell = modNames.RangeOfName(cellName)
    PaintCell cell, stateText
End Sub

' Paints the cell a click came from: the payload carries sheet and cell, so a
' rendered row (which has no name) can be painted as well.
Public Sub PaintCellFrom(ByRef data As Object, ByVal stateText As String, ByVal caption As String)
    Dim sheetName As String
    Dim addressText As String
    Dim cell As Range

    If data Is Nothing Then Exit Sub

    sheetName = modApp.PayloadText(data, modRoutes.KEY_SHEET)
    addressText = modApp.PayloadText(data, modRoutes.KEY_CELL)
    If Len(sheetName) = 0 Or Len(addressText) = 0 Then Exit Sub
    If Not modNames.SheetExists(sheetName) Then Exit Sub

    Set cell = Nothing
    On Error Resume Next
    Set cell = ThisWorkbook.Worksheets(sheetName).Range(addressText)
    On Error GoTo 0
    If cell Is Nothing Then Exit Sub

    PaintCell cell, stateText, caption
End Sub

' One cell of a rendered block: value, colours, font. The dynamic cells of a
' rendered list or of the dashboard body have no name, so everything they carry
' is written here (with the events off, like every write of this module).
Public Sub WriteCellStyled(ByRef cell As Range, ByVal value As Variant, _
                           ByVal fillColor As Long, ByVal fontColor As Long, _
                           Optional ByVal bold As Boolean = False, _
                           Optional ByVal italic As Boolean = False, _
                           Optional ByVal centered As Boolean = False)
    Dim eventsWereOn As Boolean

    If cell Is Nothing Then Exit Sub

    eventsWereOn = EventsOff()
    On Error Resume Next
    cell.NumberFormat = "@"
    cell.Value = value
    cell.Interior.Color = fillColor
    cell.Font.Color = fontColor
    cell.Font.Bold = bold
    cell.Font.Italic = italic
    If centered Then
        cell.HorizontalAlignment = xlCenter
    Else
        cell.HorizontalAlignment = xlLeft
    End If
    On Error GoTo 0
    Err.Clear
    EventsBack eventsWereOn
End Sub

' One cell of a rendered block that carries a NUMBER: the value stays a number (so
' Excel can sum it and the format decides what the user reads - the revenue column
' of the report shows 25.00 with the format of Config, while the cell holds 25).
Public Sub WriteNumberCell(ByRef cell As Range, ByVal value As Variant, _
                           ByVal numberFormat As String)
    Dim eventsWereOn As Boolean

    If cell Is Nothing Then Exit Sub

    eventsWereOn = EventsOff()
    On Error Resume Next
    cell.NumberFormat = numberFormat
    cell.Value = value
    On Error GoTo 0
    Err.Clear
    EventsBack eventsWereOn
End Sub


' The same cell, but clickable: the hyperlink points at a permanent anchor, so
' the click arrives as FollowHyperlink and the position of the cell decides what
' it meant (§6.10). An empty caption would be a link nobody can hit, so the cell
' gets a space - the free cells of the weekly grid must still look empty.
Public Sub WriteLinkCellEx(ByRef cell As Range, ByVal subAddress As String, _
                           ByVal caption As String, ByVal tip As String, _
                           ByVal fillColor As Long, ByVal fontColor As Long, _
                           Optional ByVal bold As Boolean = False, _
                           Optional ByVal italic As Boolean = False, _
                           Optional ByVal centered As Boolean = False)
    Dim eventsWereOn As Boolean
    Dim displayText As String
    Dim linkList As Object

    If cell Is Nothing Then Exit Sub

    displayText = caption
    If Len(displayText) = 0 Then displayText = " "

    eventsWereOn = EventsOff()
    On Error Resume Next
    ' The collection goes into an Object local first: a late bound property read
    ' on the result of a chain is not reliable (§13).
    Set linkList = cell.Hyperlinks
    If linkList.Count > 0 Then linkList.Delete
    WriteCellStyled cell, displayText, fillColor, fontColor, bold, italic, centered
    cell.Hyperlinks.Add cell, vbNullString, subAddress, tip, displayText
    ' A hyperlink is blue and underlined by default; a rendered cell is not.
    cell.Font.Underline = xlUnderlineStyleNone
    cell.Font.Color = fontColor
    On Error GoTo 0
    Err.Clear
    EventsBack eventsWereOn
End Sub

' A clickable cell that stays empty. The weekly grid offers a new appointment in
' every free slot, so a full week writes hundreds of them: only what the link
' itself needs is written, because the fill and the font of a cleared cell are
' already what the grid wants.
Public Sub WriteEmptyLinkCell(ByRef cell As Range, ByVal subAddress As String, ByVal tip As String)
    Dim eventsWereOn As Boolean

    If cell Is Nothing Then Exit Sub

    eventsWereOn = EventsOff()
    On Error Resume Next
    cell.Value = " "
    cell.Hyperlinks.Add cell, vbNullString, subAddress, tip, " "
    ' A hyperlink is blue and underlined by default; an empty cell is not.
    cell.Font.Underline = xlUnderlineStyleNone
    cell.Font.Color = COLOR_TEXT
    On Error GoTo 0
    Err.Clear
    EventsBack eventsWereOn
End Sub

' A clickable cell that carries a short text (the time of a free slot of the
' daily grid). The cell is empty before the render, so only what the link, the
' text, the format and the alignment need is written - a whole day is 120 of them
' and every property set costs time (§6.9). The text format matters: without it
' Excel reads "09:00" as a time value and shows "9:00".
Public Sub WriteTextLinkCell(ByRef cell As Range, ByVal subAddress As String, _
                             ByVal caption As String, ByVal tip As String)
    Dim eventsWereOn As Boolean

    If cell Is Nothing Then Exit Sub

    eventsWereOn = EventsOff()
    On Error Resume Next
    cell.NumberFormat = "@"
    cell.Value = caption
    cell.Hyperlinks.Add cell, vbNullString, subAddress, tip, caption
    ' A hyperlink is blue and underlined by default; a rendered cell is not.
    cell.Font.Underline = xlUnderlineStyleNone
    cell.Font.Color = COLOR_MUTED
    cell.HorizontalAlignment = xlCenter
    On Error GoTo 0
    Err.Clear
    EventsBack eventsWereOn
End Sub

' A cell that behaves like a button: a caption and a hyperlink anchored at a
' name, so the click arrives as FollowHyperlink (§6.1, §6.10). Used for the
' buttons of the template and for every row of a rendered list.
Public Sub WriteLinkCell(ByRef cell As Range, ByVal subAddress As String, _
                         ByVal caption As String, ByVal tip As String)
    WriteLinkCellEx cell, subAddress, caption, tip, COLOR_BUTTON, COLOR_TEXT, True, False, True
End Sub

' form.clear: the sheet in the payload decides whose form is emptied. Booking
' joins in phase 3.
Public Sub FormClear(ByRef data As Object, ByRef result As Object)
    Dim sheetName As String

    sheetName = modApp.PayloadText(data, modRoutes.KEY_SHEET)

    Select Case sheetName
        Case SHEET_CLIENTS_ADMIN
            modUiClients.ClearForm
        Case SHEET_SERVICES_ADMIN
            modUiServices.ClearForm
        Case SHEET_BOOKING
            modUiBooking.ClearForm
        Case Else
            result(modRoutes.RESULT_MESSAGE) = modText.TextFormat( _
                modText.KEY_ERROR_NO_SHEET, sheetName)
            Exit Sub
    End Select

    result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CLEARED)
    result(modRoutes.RESULT_OK) = True
End Sub

' Clears the block between two named cells (the dashboard body). The events are
' switched off while writing and their previous state is restored. The links go
' as well: a rendered block is written from scratch, and a link that stays from
' the previous render would make a free cell look booked (§6.10).
Public Sub ClearBlock(ByVal firstName As String, ByVal lastName As String)
    Dim firstCell As Range
    Dim lastCell As Range
    Dim target As Range
    Dim linkList As Object
    Dim eventsWereOn As Boolean

    Set firstCell = modNames.RangeOfName(firstName)
    Set lastCell = modNames.RangeOfName(lastName)

    eventsWereOn = EventsOff()
    On Error Resume Next
    Set target = ThisWorkbook.Worksheets(firstCell.Worksheet.Name).Range(firstCell, lastCell)
    target.Clear
    ' The collection goes into an Object local first: a late bound property read
    ' on the result of a chain is not reliable (§13).
    Set linkList = target.Hyperlinks
    linkList.Delete
    On Error GoTo 0
    EventsBack eventsWereOn
End Sub

' Writes one cell of a block that has no name (the dashboard body).
Public Sub WriteCell(ByVal sheetName As String, ByVal rowIndex As Long, _
                     ByVal columnIndex As Long, ByVal value As Variant)
    Dim eventsWereOn As Boolean

    If Not modNames.SheetExists(sheetName) Then Exit Sub

    eventsWereOn = EventsOff()
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Cells(rowIndex, columnIndex).Value = value
    On Error GoTo 0
    EventsBack eventsWereOn
End Sub

' Writes the status of a sheet without touching the interior colour (the
' dashboard header uses it for the counters as well).
Public Sub WriteOutput(ByVal nameText As String, ByVal value As Variant)
    modNames.WriteName nameText, value
End Sub

' =====================================================================
'  the modal confirmation (§6.6, 24.09.2026)
' =====================================================================

' One question with OK and Cancel. Everything that asks before it destroys
' something goes through here, so the wording and the buttons are the same
' everywhere. With the dialogs off the answer is yes, so a test and the harness
' never open a window.
Public Function AskConfirm(ByVal messageText As String) As Boolean
    Dim ask As frmConfirm

    If Not InteractiveUi() Then
        AskConfirm = True
        Exit Function
    End If

    On Error GoTo Failed
    Set ask = New frmConfirm
    ask.InitialiseWithQuestion messageText
    ask.Show vbModal
    AskConfirm = ask.Confirmed
    Unload ask
    Set ask = Nothing
    Exit Function

Failed:
    ' A dialog that cannot be shown must not approve a delete.
    Err.Clear
    AskConfirm = False
End Function
