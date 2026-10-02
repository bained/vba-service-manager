Option Explicit

' =====================================================================
'  Service_Manager - the date picker (Phase 3, §9.2)
'
'  A small modal calendar: month/year navigation, 6 x 7 day buttons, "Today" and
'  "Cancel". Picking a day stores the ISO text in PickedDateText and closes the
'  form; "" means cancelled.
'
'  The designer of this form is empty on purpose: tools\build.vbs creates the
'  UserForm component from this file, every control is added in
'  UserForm_Initialize, and the clicks arrive through clsPickerButton
'  (WithEvents) - a control the designer never held gets no compiled event
'  procedure.
'
'  Usage:
'      Set picker = New frmDatePicker
'      picker.InitialiseWithDate "2026-09-21"
'      picker.Show vbModal
'      If Len(picker.PickedDateText) > 0 Then ...
' =====================================================================

Public PickedDateText As String

Private mShownYear As Long
Private mShownMonth As Long
Private mSelectedText As String
Private mTitleLabel As MSForms.Label
Private mSelectedLabel As MSForms.Label
Private mWrappers As Collection          ' every wrapper of the form
Private mDayWrappers As Collection       ' the 42 day buttons, in grid order

' =====================================================================
'  public interface (also used by the tests)
' =====================================================================

Public Sub InitialiseWithDate(ByVal isoText As String)
    Dim dateText As String

    EnsureControls
    PickedDateText = vbNullString

    ' The cell may hold anything ("" on a new form). A date outside the window of
    ' the picker - an old one, say - is shown as today: what counts on opening is
    ' the highlight on a day the user can really pick.
    dateText = modPickers.DateTextToIso(isoText)
    If Len(dateText) = 0 Then dateText = modDashboard.FormatDateText(Date)
    If Not modPickers.IsInPickerRange(dateText) Then dateText = modDashboard.FormatDateText(Date)

    mSelectedText = dateText
    ' The ISO text is read part by part: Year()/Month() of CDate would follow the
    ' regional settings.
    mShownYear = CLng(Left$(dateText, 4))
    mShownMonth = CLng(Mid$(dateText, 6, 2))

    RenderMonth
End Sub

Public Function ShownYear() As Long
    ShownYear = mShownYear
End Function

Public Function ShownMonth() As Long
    ShownMonth = mShownMonth
End Function

Public Function ShownMonthTitle() As String
    ShownMonthTitle = modPickers.MonthTitle(mShownYear, mShownMonth)
End Function

' The ISO date a grid button carries ("" for the padding buttons).
Public Function DayValueAt(ByVal buttonIndex As Long) As String
    DayValueAt = modPickers.GridDateText(mShownYear, mShownMonth, buttonIndex)
End Function

Public Function DayCaptionAt(ByVal buttonIndex As Long) As String
    DayCaptionAt = DayButtonAt(buttonIndex).Caption
End Function

Public Function DayEnabledAt(ByVal buttonIndex As Long) As Boolean
    DayEnabledAt = DayButtonAt(buttonIndex).Enabled
End Function

Public Function DayCount() As Long
    If mDayWrappers Is Nothing Then Exit Function

    DayCount = mDayWrappers.Count
End Function

' Grid position of the highlighted date, 0 when nothing is highlighted.
Public Function SelectedGridIndex() As Long
    SelectedGridIndex = modPickers.GridIndexOf(mShownYear, mShownMonth, mSelectedText)
End Function

' The one click handler of every button of the form (§9.2).
Public Sub ButtonClicked(ByVal roleText As String, ByVal valueText As String)
    Select Case Trim$(roleText)
        Case ROLE_PREV_MONTH
            ShiftMonth -1
        Case ROLE_NEXT_MONTH
            ShiftMonth 1
        Case ROLE_PREV_YEAR
            ShiftMonth -12
        Case ROLE_NEXT_YEAR
            ShiftMonth 12
        Case ROLE_TODAY
            AcceptDate modDashboard.FormatDateText(Date)
        Case ROLE_CANCEL
            PickedDateText = vbNullString
            Hide
        Case ROLE_DAY
            AcceptDate valueText
    End Select
End Sub

Private Sub AcceptDate(ByVal isoText As String)
    If Len(isoText) < ISO_DATE_LENGTH Then Exit Sub

    PickedDateText = isoText
    Hide
End Sub

' =====================================================================
'  building the frame and the controls
' =====================================================================

Private Sub UserForm_Initialize()
    EnsureControls
End Sub

Private Sub EnsureControls()
    If Not (mWrappers Is Nothing) Then Exit Sub

    Me.Caption = modText.Text(modText.KEY_PICKER_DATE_TITLE)
    Me.StartUpPosition = PICKER_START_POSITION
    SetFrameSize

    Set mWrappers = New Collection
    Set mDayWrappers = New Collection

    ' The shown month comes first: BuildDayGrid hands every day button its date,
    ' and a grid built while the year was still 0 carried 1899-12-xx. A click
    ' answers the value of the button (clsPickerButton), not its caption, so that
    ' stale value was what the user got - whatever he clicked.
    mShownYear = CLng(Year(Date))
    mShownMonth = CLng(Month(Date))
    mSelectedText = modDashboard.FormatDateText(Date)

    BuildHeader
    BuildWeekdayRow
    BuildDayGrid
    BuildFooter

    mTitleLabel.Caption = ShownMonthTitle()
End Sub

' InsideWidth/InsideHeight are read only, so the frame is measured as the
' difference between the outside and the inside size at this moment.
Private Sub SetFrameSize()
    Dim borderWidth As Single
    Dim borderHeight As Single

    borderWidth = Me.Width - Me.InsideWidth
    borderHeight = Me.Height - Me.InsideHeight
    If borderWidth < 0 Then borderWidth = 0
    If borderHeight < 0 Then borderHeight = 0

    Me.Width = FrameWidth() + borderWidth
    Me.Height = FrameHeight() + borderHeight
End Sub

Private Function FrameWidth() As Single
    FrameWidth = PICKER_MARGIN * 2 + PICKER_GRID_DAYS * PICKER_CELL_WIDTH
End Function

Private Function GridTop() As Single
    GridTop = PICKER_MARGIN + PICKER_HEADER_HEIGHT + PICKER_GAP + PICKER_WEEKDAY_HEIGHT
End Function

Private Function FooterTop() As Single
    FooterTop = GridTop() + PICKER_GRID_WEEKS * PICKER_CELL_HEIGHT + PICKER_GAP
End Function

Private Function FrameHeight() As Single
    FrameHeight = FooterTop() + PICKER_HEADER_HEIGHT + PICKER_MARGIN
End Function

Private Function AddLabel(ByVal controlName As String, ByVal leftValue As Single, _
                          ByVal topValue As Single, ByVal widthValue As Single, _
                          ByVal heightValue As Single) As MSForms.Label
    Dim captionLabel As MSForms.Label

    Set captionLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    captionLabel.Left = leftValue
    captionLabel.Top = topValue
    captionLabel.Width = widthValue
    captionLabel.Height = heightValue
    captionLabel.Font.Size = 8

    Set AddLabel = captionLabel
End Function

' Every button of the form is created here and wrapped by clsPickerButton, so
' the click arrives although the designer never held the control (§9.2).
Private Function AddButton(ByVal roleText As String, ByVal valueText As String, _
                           ByVal captionText As String, ByVal leftValue As Single, _
                           ByVal topValue As Single, ByVal widthValue As Single, _
                           ByVal heightValue As Single) As MSForms.CommandButton
    Dim button As MSForms.CommandButton
    Dim wrapper As clsPickerButton

    Set button = Me.Controls.Add("Forms.CommandButton.1", _
                                 "ctl" & CStr(mWrappers.Count + 1), True)
    button.Caption = captionText
    button.Left = leftValue
    button.Top = topValue
    button.Width = widthValue
    button.Height = heightValue
    button.Font.Size = 8
    button.TakeFocusOnClick = False
    button.Tag = valueText

    Set wrapper = New clsPickerButton
    wrapper.Attach button, Me, roleText, valueText
    mWrappers.Add wrapper

    Set AddButton = button
End Function

Private Sub BuildHeader()
    Dim rightEdge As Single

    rightEdge = FrameWidth() - PICKER_MARGIN

    AddButton ROLE_PREV_YEAR, vbNullString, "<<", PICKER_MARGIN, PICKER_MARGIN, _
              PICKER_NAV_WIDTH, PICKER_HEADER_HEIGHT
    AddButton ROLE_PREV_MONTH, vbNullString, "<", PICKER_MARGIN + PICKER_NAV_WIDTH, _
              PICKER_MARGIN, PICKER_NAV_WIDTH, PICKER_HEADER_HEIGHT

    Set mTitleLabel = AddLabel("lblTitle", PICKER_MARGIN + PICKER_NAV_WIDTH * 2, _
                               PICKER_MARGIN, _
                               FrameWidth() - (PICKER_MARGIN + PICKER_NAV_WIDTH * 2) * 2, _
                               PICKER_HEADER_HEIGHT)
    mTitleLabel.TextAlign = 2          ' fmTextAlignCenter
    mTitleLabel.Font.Bold = True
    mTitleLabel.Font.Size = 10

    AddButton ROLE_NEXT_MONTH, vbNullString, ">", rightEdge - PICKER_NAV_WIDTH * 2, _
              PICKER_MARGIN, PICKER_NAV_WIDTH, PICKER_HEADER_HEIGHT
    AddButton ROLE_NEXT_YEAR, vbNullString, ">>", rightEdge - PICKER_NAV_WIDTH, _
              PICKER_MARGIN, PICKER_NAV_WIDTH, PICKER_HEADER_HEIGHT
End Sub

Private Sub BuildWeekdayRow()
    Dim index As Long
    Dim captionLabel As MSForms.Label

    For index = 1 To PICKER_GRID_DAYS
        Set captionLabel = AddLabel("wd" & CStr(index), _
                                    PICKER_MARGIN + (index - 1) * PICKER_CELL_WIDTH, _
                                    PICKER_MARGIN + PICKER_HEADER_HEIGHT + PICKER_GAP, _
                                    PICKER_CELL_WIDTH, PICKER_WEEKDAY_HEIGHT)
        captionLabel.Caption = modText.WeekdayShortName(index)
        captionLabel.TextAlign = 2
        captionLabel.Font.Bold = True
    Next index
End Sub

Private Sub BuildDayGrid()
    Dim index As Long
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim button As MSForms.CommandButton
    Dim isoText As String

    For index = 1 To PICKER_GRID_DAYS * PICKER_GRID_WEEKS
        columnIndex = (index - 1) Mod PICKER_GRID_DAYS
        rowIndex = (index - 1) \ PICKER_GRID_DAYS
        isoText = modPickers.GridDateText(mShownYear, mShownMonth, index)

        Set button = AddButton(ROLE_DAY, isoText, vbNullString, _
                               PICKER_MARGIN + columnIndex * PICKER_CELL_WIDTH, _
                               GridTop() + rowIndex * PICKER_CELL_HEIGHT, _
                               PICKER_CELL_WIDTH, PICKER_CELL_HEIGHT)

        mDayWrappers.Add mWrappers.Item(mWrappers.Count)
    Next index
End Sub

Private Function DayButtonAt(ByVal buttonIndex As Long) As MSForms.CommandButton
    Set DayButtonAt = DayWrapperAt(buttonIndex).PickerButton
End Function

Private Function DayWrapperAt(ByVal buttonIndex As Long) As clsPickerButton
    EnsureControls

    If buttonIndex < 1 Or buttonIndex > mDayWrappers.Count Then
        Err.Raise 5, "frmDatePicker.DayWrapperAt", _
                  "Day button out of range: " & CStr(buttonIndex)
    End If

    Set DayWrapperAt = mDayWrappers.Item(buttonIndex)
End Function

' The value the click of a day button carries. clsPickerButton passes this on
' (not the caption), so a test walks exactly the path of a real click.
Public Function GridValueAt(ByVal buttonIndex As Long) As String
    GridValueAt = DayWrapperAt(buttonIndex).ValueText
End Function

' The same as clicking a day cell (without the mouse).
Public Sub ClickDayAt(ByVal buttonIndex As Long)
    ButtonClicked ROLE_DAY, GridValueAt(buttonIndex)
End Sub

' The selected date the way the user reads it: "17.09.2026".
Public Function SelectedDateDisplay() As String
    SelectedDateDisplay = modPickers.DateFormatText(mSelectedText)
End Function

Private Sub BuildFooter()
    Dim labelWidth As Single
    Dim buttonWidth As Single
    Dim usedWidth As Single

    ' The left part of the footer shows the selected date the European way
    ' ("17.09.2026"): the data of the workbook stays ISO text.
    usedWidth = FrameWidth() - PICKER_MARGIN * 2
    labelWidth = usedWidth * 0.3
    buttonWidth = (usedWidth - labelWidth - PICKER_GAP * 2) / 2

    Set mSelectedLabel = AddLabel("lblSelected", PICKER_MARGIN, FooterTop(), labelWidth, _
                                  PICKER_HEADER_HEIGHT)
    mSelectedLabel.Font.Bold = True
    mSelectedLabel.Font.Size = 9
    mSelectedLabel.Caption = SelectedDateDisplay()

    AddButton ROLE_TODAY, vbNullString, modText.Text(modText.KEY_PICKER_TODAY), _
              PICKER_MARGIN + labelWidth + PICKER_GAP, FooterTop(), buttonWidth, _
              PICKER_HEADER_HEIGHT
    AddButton ROLE_CANCEL, vbNullString, modText.Text(modText.KEY_PICKER_CANCEL), _
              PICKER_MARGIN + labelWidth + PICKER_GAP * 2 + buttonWidth, FooterTop(), _
              buttonWidth, PICKER_HEADER_HEIGHT
End Sub

' Draws the shown month: captions, the disabled cells (padding and dates outside
' the window of the picker) and the highlight of the selected day. Every button
' gets its value again: after a month shift the old value would answer the click.
Private Sub RenderMonth()
    Dim index As Long
    Dim isoText As String
    Dim button As MSForms.CommandButton
    Dim selected As Long

    EnsureControls

    mTitleLabel.Caption = ShownMonthTitle()
    mSelectedLabel.Caption = SelectedDateDisplay()
    selected = SelectedGridIndex()

    For index = 1 To mDayWrappers.Count
        isoText = modPickers.GridDateText(mShownYear, mShownMonth, index)
        Set button = DayButtonAt(index)

        If Len(isoText) = 0 Then
            button.Caption = vbNullString
            button.Enabled = False
            button.BackColor = COLOR_PICKER_DISABLED
            SetPickerValue index, vbNullString
        Else
            button.Caption = CStr(CLng(Mid$(isoText, 9, 2)))
            button.Enabled = modPickers.IsInPickerRange(isoText)
            SetPickerValue index, isoText

            If Not button.Enabled Then
                button.BackColor = COLOR_PICKER_DISABLED
            ElseIf index = selected Then
                button.BackColor = COLOR_BUTTON_ACTIVE
            Else
                button.BackColor = COLOR_FREE
            End If
        End If
    Next index
End Sub

Private Sub ShiftMonth(ByVal months As Long)
    Dim total As Long
    Dim firstTotal As Long
    Dim lastTotal As Long

    ' The window of the picker: the year before this one up to
    ' PICKER_YEARS_FORWARD years ahead (§9.2).
    firstTotal = (Year(Date) - PICKER_YEARS_BACK) * 12
    lastTotal = (Year(Date) + PICKER_YEARS_FORWARD) * 12 + 11

    total = (mShownYear * 12 + mShownMonth - 1) + months
    If total < firstTotal Then total = firstTotal
    If total > lastTotal Then total = lastTotal

    mShownYear = total \ 12
    mShownMonth = (total Mod 12) + 1

    RenderMonth
End Sub

' The value a real click sends: the wrapper and the Tag must follow the shown
' month, otherwise a click after a month shift answers the month of the opening.
Private Sub SetPickerValue(ByVal buttonIndex As Long, ByVal isoText As String)
    Dim wrapper As clsPickerButton

    Set wrapper = DayWrapperAt(buttonIndex)
    wrapper.ValueText = isoText
    wrapper.PickerButton.Tag = isoText
End Sub
