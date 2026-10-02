Option Explicit

' =====================================================================
'  Service_Manager - the confirmation dialog (§6.6)
'
'  One question and two answers: "Are you sure?" with OK and Cancel. It replaced
'  the two step confirmation (click, then click again within five seconds,
'  24.09.2026): a modal window cannot be hit by accident and it says what it wants.
'
'  Same technique as the pickers: the designer is empty, the controls are built in
'  InitialiseWithQuestion and the clicks arrive through clsPickerButton (WithEvents)
'  - a control the designer never held gets no compiled event procedure.
'
'  Usage:
'      Set ask = New frmConfirm
'      ask.InitialiseWithQuestion "Are you sure?"
'      ask.Show vbModal
'      If ask.Confirmed Then ...
' =====================================================================

Public Confirmed As Boolean

Private mMessage As String
Private mLabel As MSForms.Label
Private mWrappers As Collection
Private mBuilt As Boolean

' =====================================================================
'  public interface (also used by the tests)
' =====================================================================

Public Sub InitialiseWithQuestion(ByVal messageText As String)
    EnsureControls

    Confirmed = False
    mMessage = messageText
    mLabel.Caption = messageText

    Me.Caption = APP_NAME
    ' The title bar of a UserForm takes about twenty points of the height away, so the
    ' OK/Cancel row needs that much room on top of its own height (26.09.2026).
    Me.Width = FormWidth() + 12
    Me.Height = FooterTop() + PICKER_HEADER_HEIGHT + 46
End Sub

Public Function MessageText() As String
    MessageText = mMessage
End Function

' The one click handler of the form: OK answers yes, Cancel (and closing) no.
Public Sub ButtonClicked(ByVal roleText As String, ByVal valueText As String)
    Select Case Trim$(roleText)
        Case ROLE_OK
            Confirmed = True
            Hide
        Case ROLE_CANCEL
            Confirmed = False
            Hide
    End Select
End Sub

' =====================================================================
'  the controls (the designer is empty, so all of them are built here)
' =====================================================================

Private Sub EnsureControls()
    If mBuilt Then Exit Sub
    mBuilt = True

    Set mWrappers = New Collection

    Set mLabel = Me.Controls.Add("Forms.Label.1", "lblQuestion", True)
    mLabel.Left = PICKER_MARGIN
    mLabel.Top = PICKER_MARGIN
    mLabel.Width = FormWidth() - PICKER_MARGIN * 2
    mLabel.Height = PICKER_QUESTION_HEIGHT
    mLabel.Font.Size = 10
    mLabel.WordWrap = True

    BuildFooter
End Sub

Private Sub BuildFooter()
    Dim halfWidth As Single

    halfWidth = (FormWidth() - PICKER_MARGIN * 2 - PICKER_GAP) / 2

    AddButton ROLE_OK, modText.Text(modText.KEY_PICKER_OK), _
              PICKER_MARGIN, FooterTop(), halfWidth, PICKER_HEADER_HEIGHT
    AddButton ROLE_CANCEL, modText.Text(modText.KEY_PICKER_CANCEL), _
              PICKER_MARGIN + halfWidth + PICKER_GAP, FooterTop(), halfWidth, _
              PICKER_HEADER_HEIGHT
End Sub

' =====================================================================
'  geometry
' =====================================================================

Private Function FormWidth() As Single
    FormWidth = PICKER_CONFIRM_WIDTH
End Function

Private Function FooterTop() As Single
    FooterTop = PICKER_MARGIN + PICKER_QUESTION_HEIGHT + PICKER_GAP
End Function

Private Function AddButton(ByVal roleText As String, ByVal captionText As String, _
                           ByVal leftValue As Single, ByVal topValue As Single, _
                           ByVal widthValue As Single, ByVal heightValue As Single) As MSForms.CommandButton
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
    button.Tag = vbNullString

    Set wrapper = New clsPickerButton
    wrapper.Attach button, Me, roleText, vbNullString
    mWrappers.Add wrapper

    Set AddButton = button
End Function
