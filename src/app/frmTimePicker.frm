Option Explicit

' =====================================================================
'  Service_Manager - the time picker (Phase 3, §9.3)
'
'  Two dropdowns and nothing else: the hour (00..23) and the minute. The form
'  answers "HH:MM"; "" means cancelled or "no value at all".
'
'  The list of minutes follows the field (not the user): a start has to sit on
'  the grid step of Config (step 15 -> 00/15/30/45, §8.2.7), a duration may be
'  any 5 minute value (§9.4). So the picker cannot offer a value that the domain
'  would refuse.
'
'  Same technique as frmDatePicker: the designer is empty, the controls are built
'  in UserForm_Initialize and the clicks arrive through clsPickerButton
'  (WithEvents) - a control the designer never held gets no compiled event.
'
'  Usage:
'      Set picker = New frmTimePicker
'      picker.InitialiseWithTime "start", "09:30", 60, "10:00+30"
'      picker.Show vbModal
'      If Len(picker.PickedTimeText) > 0 Then ...
' =====================================================================

Public PickedTimeText As String

Private mField As String                 ' FIELD_START or FIELD_DURATION
Private mDurationMinutes As Long         ' the duration the start has to fit into
Private mBusyText As String              ' "09:00+30,10:30+90"
Private mWindowStart As Long             ' the working day of Config, in minutes
Private mWindowEnd As Long
Private mHourBox As MSForms.ComboBox
Private mMinuteBox As MSForms.ComboBox
Private mHintLabel As MSForms.Label
Private mWrappers As Collection

' =====================================================================
'  public interface (also used by the tests)
' =====================================================================

' The working day (in minutes from midnight) the hours of a start are taken from;
' 0 and 1440 mean "the whole day" and are what a duration uses.
Public Sub InitialiseWithTime(ByVal fieldKey As String, ByVal initialText As String, _
                              ByVal durationMinutes As Long, ByVal busyText As String, _
                              Optional ByVal windowStartMinutes As Long = 0, _
                              Optional ByVal windowEndMinutes As Long = 1440)
    Dim timeText As String
    Dim total As Long
    Dim windowHour As Long

    EnsureControls
    PickedTimeText = vbNullString
    SetHint vbNullString

    mField = LCase$(Trim$(fieldKey))
    If StrComp(mField, FIELD_DURATION, vbTextCompare) <> 0 Then mField = FIELD_START

    mDurationMinutes = durationMinutes
    mBusyText = busyText
    mWindowStart = windowStartMinutes
    mWindowEnd = windowEndMinutes

    FillHourOptions
    FillMinuteOptions

    timeText = modPickers.NormaliseTimeText(initialText)
    If Len(timeText) > 0 Then
        total = modApp.TimeToMinutes(timeText)
        ChooseHour total \ MINUTES_PER_HOUR
        ChooseMinute total Mod MINUTES_PER_HOUR

        ' An appointment that sits outside the working day (the day was changed
        ' after it was booked) still opens on its own hour: it is added to the list
        ' instead of falling back to the first slot and being lost on OK.
        If Hours() < 0 Then
            windowHour = total \ MINUTES_PER_HOUR
            mHourBox.AddItem Format$(windowHour, "00")
            ChooseHour windowHour
        End If
    Else
        ' Nothing set yet (a new appointment): the two dropdowns open on their
        ' first value ("00" and "00"). An empty form looks broken and keeps the
        ' user guessing, and OK without a choice would answer nothing at all.
        ChooseFirst
    End If

    Me.Caption = FieldTitle()
End Sub

Public Function FieldKey() As String
    FieldKey = mField
End Function

' The chosen hour; -1 while nothing is chosen.
Public Function Hours() As Long
    Hours = ChosenValue(mHourBox)
End Function

' The chosen minute; -1 while nothing is chosen.
Public Function Minutes() As Long
    Minutes = ChosenValue(mMinuteBox)
End Function

' "09:30" of the current choice, "" when neither dropdown has a choice.
Public Function CurrentText() As String
    Dim hoursValue As Long
    Dim minutesValue As Long

    hoursValue = Hours()
    minutesValue = Minutes()
    If hoursValue < 0 And minutesValue < 0 Then Exit Function

    If hoursValue < 0 Then hoursValue = 0
    If minutesValue < 0 Then minutesValue = 0

    CurrentText = modApp.MinutesToTime(hoursValue * MINUTES_PER_HOUR + minutesValue)
End Function

' The line under the dropdowns; empty until something has to be explained.
Public Function HintText() As String
    HintText = mHintLabel.Caption
End Function

Public Function HourOptionCount() As Long
    EnsureControls
    HourOptionCount = mHourBox.ListCount
End Function

Public Function HourOptionAt(ByVal optionIndex As Long) As String
    EnsureControls
    HourOptionAt = OptionText(mHourBox, optionIndex)
End Function

Public Function MinuteOptionCount() As Long
    EnsureControls
    MinuteOptionCount = mMinuteBox.ListCount
End Function

Public Function MinuteOptionAt(ByVal optionIndex As Long) As String
    EnsureControls
    MinuteOptionAt = OptionText(mMinuteBox, optionIndex)
End Function

' Picks one of the offered values (the tests use this instead of a mouse).
Public Sub ChooseHour(ByVal hourValue As Long)
    EnsureControls
    mHourBox.ListIndex = OptionIndex(mHourBox, Format$(hourValue, "00"))
End Sub

Public Sub ChooseMinute(ByVal minuteValue As Long)
    EnsureControls
    mMinuteBox.ListIndex = OptionIndex(mMinuteBox, Format$(minuteValue, "00"))
End Sub

' Both dropdowns on their first option: "00" and "00". This is what a field that
' is still empty opens with (§9.3).
Public Sub ChooseFirst()
    EnsureControls
    mHourBox.ListIndex = 0
    mMinuteBox.ListIndex = 0
End Sub

' Both dropdowns empty ("no value"). The form is only left this way by ClearForm
' of the caller - on opening it never is (§9.3).
Public Sub ClearChoice()
    EnsureControls
    mHourBox.ListIndex = -1
    mMinuteBox.ListIndex = -1
End Sub

' The one click handler of the form.
Public Sub ButtonClicked(ByVal roleText As String, ByVal valueText As String)
    Select Case Trim$(roleText)
        Case ROLE_OK
            Accept
        Case ROLE_CANCEL
            PickedTimeText = vbNullString
            Hide
    End Select
End Sub

' =====================================================================
'  the answer
' =====================================================================

Private Sub Accept()
    Dim text As String

    SetHint vbNullString
    text = CurrentText()

    ' Nothing chosen means "no value": the caller clears the field.
    If Len(text) = 0 Then
        PickedTimeText = vbNullString
        Hide
        Exit Sub
    End If

    ' A start must not run into an appointment of the day (§9.3).
    If StrComp(mField, FIELD_START, vbTextCompare) = 0 Then
        If modPickers.IsSlotBusy(modApp.TimeToMinutes(text), mDurationMinutes, mBusyText) Then
            SetHint modText.Text(modText.KEY_PICKER_BUSY)
            Exit Sub
        End If
    End If

    PickedTimeText = text
    Hide
End Sub

' =====================================================================
'  the two lists
' =====================================================================

' The hours the picker offers. A start is limited to the working day it has to
' live in (§9.3, 20.09.2026): 09:00-19:00 offers 09..18 and nothing before or after
' it, so the picker itself cannot suggest a start the domain would refuse. A
' duration is not a clock time and keeps the whole day.
Private Sub FillHourOptions()
    Dim hourValue As Long
    Dim firstHour As Long
    Dim lastHour As Long

    ' The whole day is the default; only a start is narrowed to the working day.
    firstHour = 0
    lastHour = 23

    If StrComp(mField, FIELD_START, vbTextCompare) = 0 Then
        firstHour = modPickers.StartHourFirst(mWindowStart, mWindowEnd)
        lastHour = modPickers.StartHourLast(mWindowStart, mWindowEnd)
    End If

    ' A window that leaves no hour at all falls back to the whole day.
    If lastHour < firstHour Then
        firstHour = 0
        lastHour = 23
    End If

    mHourBox.Clear
    For hourValue = firstHour To lastHour
        mHourBox.AddItem Format$(hourValue, "00")
    Next hourValue
End Sub

' The minutes the field offers: a start has to sit on the step of Config - the
' resolution the user set (5 / 15 / 30 / 60) - while a duration may be any 5
' minutes (§6.9, §9.3).
Private Sub FillMinuteOptions()
    Dim minuteValue As Long
    Dim minuteStep As Long

    If StrComp(mField, FIELD_DURATION, vbTextCompare) = 0 Then
        minuteStep = TIME_STEP_MINUTES
    Else
        minuteStep = modAppointments.StepMinutes()
    End If
    If minuteStep < 1 Then minuteStep = TIME_STEP_MINUTES
    If minuteStep > MINUTES_PER_HOUR Then minuteStep = MINUTES_PER_HOUR

    mMinuteBox.Clear
    For minuteValue = 0 To MINUTES_PER_HOUR - 1 Step minuteStep
        mMinuteBox.AddItem Format$(minuteValue, "00")
    Next minuteValue
End Sub

' =====================================================================
'  small helpers for the dropdowns
' =====================================================================

' The chosen number of a dropdown; -1 while there is no choice.
Private Function ChosenValue(ByRef box As MSForms.ComboBox) As Long
    ChosenValue = -1
    If box Is Nothing Then Exit Function
    If box.ListIndex < 0 Then Exit Function

    ChosenValue = CLng(Val(box.Value))
End Function

Private Function OptionText(ByRef box As MSForms.ComboBox, ByVal optionIndex As Long) As String
    If box Is Nothing Then Exit Function
    If optionIndex < 0 Or optionIndex >= box.ListCount Then Exit Function

    OptionText = CStr(box.List(optionIndex))
End Function

Private Function OptionIndex(ByRef box As MSForms.ComboBox, ByVal wantedText As String) As Long
    Dim index As Long

    OptionIndex = -1
    If box Is Nothing Then Exit Function

    For index = 0 To box.ListCount - 1
        If StrComp(CStr(box.List(index)), wantedText, vbTextCompare) = 0 Then
            OptionIndex = index
            Exit Function
        End If
    Next index
End Function

Private Sub SetHint(ByVal text As String)
    If mHintLabel Is Nothing Then Exit Sub

    mHintLabel.Caption = text
End Sub

Private Function FieldTitle() As String
    If StrComp(mField, FIELD_DURATION, vbTextCompare) = 0 Then
        FieldTitle = modText.Text(modText.KEY_PICKER_DURATION_TITLE)
    Else
        FieldTitle = modText.Text(modText.KEY_PICKER_TIME_TITLE)
    End If
End Function

' =====================================================================
'  building the frame and the four controls
' =====================================================================

Private Sub UserForm_Initialize()
    EnsureControls
End Sub

Private Sub EnsureControls()
    If Not (mWrappers Is Nothing) Then Exit Sub

    Me.Caption = modText.Text(modText.KEY_PICKER_TIME_TITLE)
    Me.StartUpPosition = PICKER_START_POSITION
    SetFrameSize

    Set mWrappers = New Collection
    BuildFields
    BuildFooter
End Sub

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

' Two rows of (label, dropdown), one line for a message, one row of buttons.
Private Function FrameWidth() As Single
    FrameWidth = PICKER_MARGIN * 2 + PICKER_CELL_WIDTH * 4
End Function

Private Function FieldTop(ByVal rowIndex As Long) As Single
    FieldTop = PICKER_MARGIN + rowIndex * (PICKER_CELL_HEIGHT + PICKER_GAP)
End Function

Private Function HintTop() As Single
    HintTop = FieldTop(2)
End Function

Private Function FooterTop() As Single
    FooterTop = HintTop() + PICKER_LABEL_HEIGHT + PICKER_GAP
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
    captionLabel.Font.Size = 9

    Set AddLabel = captionLabel
End Function

Private Function AddBox(ByVal controlName As String, ByVal leftValue As Single, _
                        ByVal topValue As Single, ByVal widthValue As Single, _
                        ByVal heightValue As Single) As MSForms.ComboBox
    Dim box As MSForms.ComboBox

    Set box = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    box.Left = leftValue
    box.Top = topValue
    box.Width = widthValue
    box.Height = heightValue
    box.Font.Size = 10
    box.Style = 2                      ' fmStyleDropDownList: only the list is offered

    Set AddBox = box
End Function

' Every button of the form is created here and wrapped by clsPickerButton, so the
' click arrives although the designer never held the control.
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
    button.Font.Size = 9
    button.TakeFocusOnClick = False
    button.Tag = valueText

    Set wrapper = New clsPickerButton
    wrapper.Attach button, Me, roleText, valueText
    mWrappers.Add wrapper

    Set AddButton = button
End Function

Private Sub BuildFields()
    Dim captionLabel As MSForms.Label
    Dim labelWidth As Single
    Dim boxLeft As Single

    labelWidth = PICKER_CELL_WIDTH * 2
    boxLeft = PICKER_MARGIN + labelWidth

    Set captionLabel = AddLabel("lblHour", PICKER_MARGIN, FieldTop(0), labelWidth, _
                                PICKER_CELL_HEIGHT)
    captionLabel.Caption = modText.Text(modText.KEY_PICKER_HOURS)

    Set mHourBox = AddBox("cbHour", boxLeft, FieldTop(0), labelWidth, PICKER_CELL_HEIGHT)

    Set captionLabel = AddLabel("lblMinute", PICKER_MARGIN, FieldTop(1), labelWidth, _
                                PICKER_CELL_HEIGHT)
    captionLabel.Caption = modText.Text(modText.KEY_PICKER_MINUTES)

    Set mMinuteBox = AddBox("cbMinute", boxLeft, FieldTop(1), labelWidth, PICKER_CELL_HEIGHT)

    Set mHintLabel = AddLabel("lblHint", PICKER_MARGIN, HintTop(), _
                              FrameWidth() - PICKER_MARGIN * 2, PICKER_LABEL_HEIGHT)
    mHintLabel.Font.Bold = True
End Sub

Private Sub BuildFooter()
    Dim halfWidth As Single

    halfWidth = (FrameWidth() - PICKER_MARGIN * 2 - PICKER_GAP) / 2

    AddButton ROLE_OK, vbNullString, modText.Text(modText.KEY_PICKER_OK), _
              PICKER_MARGIN, FooterTop(), halfWidth, PICKER_HEADER_HEIGHT
    AddButton ROLE_CANCEL, vbNullString, modText.Text(modText.KEY_PICKER_CANCEL), _
              PICKER_MARGIN + halfWidth + PICKER_GAP, FooterTop(), halfWidth, _
              PICKER_HEADER_HEIGHT
End Sub
