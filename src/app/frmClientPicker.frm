Option Explicit

' =====================================================================
'  Service_Manager - the client picker (Phase 8, §9.5)
'
'  A typed search over the clients: the box takes any part of the name or of the
'  phone and the list under it narrows with every keystroke. The form answers the
'  name of the picked client ("" when it was cancelled).
'
'  The reason it exists: with a few hundred clients the dropdown of the Booking
'  sheet is unusable, and Excel's own autocomplete only matches the beginning of
'  the entry - which is exactly what fails when one cannot remember whether the
'  first name or the surname was typed first.
'
'  Same technique as the other two pickers: the designer is empty, the controls
'  are built in InitialiseWithSearch, and the events arrive through the runtime
'  wrappers - clsPickerBox for the typing, clsPickerButton for the clicks. A
'  control the designer never held gets no compiled event procedure.
'
'  Usage:
'      Set picker = New frmClientPicker
'      picker.InitialiseWithSearch "pet"
'      picker.Show vbModal
'      If Len(picker.PickedClientText) > 0 Then ...
' =====================================================================

Public PickedClientText As String

Private mRows As Collection                  ' the shown results, "id|name|phone"
Private mBoxes As Collection                 ' clsPickerBox wrappers
Private mWrappers As Collection              ' clsPickerButton wrappers
Private mSearchBox As MSForms.TextBox
Private mListBox As MSForms.ListBox
Private mHintLabel As MSForms.Label
Private mBuilt As Boolean

' =====================================================================
'  public interface (also used by the tests)
' =====================================================================

Public Sub InitialiseWithSearch(ByVal searchText As String)
    EnsureControls
    PickedClientText = vbNullString
    SetHint vbNullString

    Me.Caption = modText.Text(modText.KEY_PICKER_CLIENT_TITLE)
    ' The title bar of a UserForm takes about twenty points of the height away, so the
    ' footer buttons need that much room on top of their own height - with a margin of
    ' 22 the bottom buttons were half cut off (the owner, 26.09.2026).
    Me.Width = FormWidth() + 12
    Me.Height = FooterTop() + PICKER_HEADER_HEIGHT + 46

    mSearchBox.Text = Trim$(searchText)
    FillResults Trim$(searchText)
    If mListBox.ListCount > 0 Then mListBox.ListIndex = 0
End Sub

Public Function ResultCount() As Long
    EnsureControls
    ResultCount = mListBox.ListCount
End Function

' What a shown result reads as (the tests use it, the form only shows it).
Public Function ResultAt(ByVal itemIndex As Long) As String
    EnsureControls
    ResultAt = modPickers.SearchRowText(RowAt(itemIndex))
End Function

' The whole result line ("id|name|phone") - the tests read the id from it.
Public Function ResultRowAt(ByVal itemIndex As Long) As String
    EnsureControls
    ResultRowAt = RowAt(itemIndex)
End Function

' Picks a result without a mouse (the tests).
Public Sub ChooseResult(ByVal itemIndex As Long)
    EnsureControls
    mListBox.ListIndex = itemIndex
End Sub

Public Function SelectedRowText() As String
    EnsureControls
    SelectedRowText = RowAt(mListBox.ListIndex)
End Function

Public Function SearchText() As String
    EnsureControls
    SearchText = mSearchBox.Text
End Function

Public Function HintText() As String
    EnsureControls
    HintText = mHintLabel.Caption
End Function

' Every keystroke of the search box arrives here (clsPickerBox).
Public Sub SearchChanged(ByVal text As String)
    EnsureControls
    FillResults text
    If mListBox.ListCount > 0 Then mListBox.ListIndex = 0
End Sub

' The one click handler of the form.
Public Sub ButtonClicked(ByVal roleText As String, ByVal valueText As String)
    Select Case Trim$(roleText)
        Case ROLE_OK
            Accept
        Case ROLE_CANCEL
            PickedClientText = vbNullString
            Hide
    End Select
End Sub

' =====================================================================
'  the answer
' =====================================================================

Private Sub Accept()
    Dim rowText As String
    Dim idText As String
    Dim nameText As String
    Dim phoneText As String

    SetHint vbNullString
    rowText = SelectedRowText()
    If Len(rowText) > 0 Then
        modPickers.SplitSearchRow rowText, idText, nameText, phoneText
    End If

    If Len(nameText) = 0 Then
        SetHint modText.TextFormat(modText.KEY_PICKER_CLIENT_NONE, Trim$(mSearchBox.Text))
        Exit Sub
    End If

    PickedClientText = nameText
    Hide
End Sub

' =====================================================================
'  the list of results
' =====================================================================

Private Sub FillResults(ByVal term As String)
    Dim rows As Collection
    Dim item As Variant
    Dim rowText As String
    Dim idText As String
    Dim nameText As String
    Dim phoneText As String

    Set mRows = New Collection
    mListBox.Clear

    Set rows = modClients.SearchRows(term)
    For Each item In rows
        rowText = CStr(item)
        mRows.Add rowText
        modPickers.SplitSearchRow rowText, idText, nameText, phoneText
        mListBox.AddItem nameText
        mListBox.List(mListBox.ListCount - 1, 1) = phoneText
    Next item

    If Len(Trim$(term)) > 0 And mRows.Count = 0 Then
        SetHint modText.TextFormat(modText.KEY_PICKER_CLIENT_NONE, Trim$(term))
    Else
        SetHint vbNullString
    End If
End Sub

Private Function RowAt(ByVal itemIndex As Long) As String
    If mRows Is Nothing Then Exit Function
    If itemIndex < 0 Or itemIndex >= mRows.Count Then Exit Function

    RowAt = CStr(mRows(itemIndex + 1))
End Function

Private Sub SetHint(ByVal textText As String)
    If mHintLabel Is Nothing Then Exit Sub

    mHintLabel.Caption = textText
End Sub

' =====================================================================
'  the controls (the designer is empty, so all of them are built here)
' =====================================================================

Private Sub EnsureControls()
    If mBuilt Then Exit Sub
    mBuilt = True

    Set mRows = New Collection
    Set mBoxes = New Collection
    Set mWrappers = New Collection

    BuildSearch
    BuildList
    BuildFooter
End Sub

Private Sub BuildSearch()
    Dim captionLabel As MSForms.Label

    Set captionLabel = AddLabel("lblSearchHint", PICKER_MARGIN, PICKER_MARGIN, _
                                FormWidth() - PICKER_MARGIN * 2, PICKER_LABEL_HEIGHT)
    captionLabel.Caption = modText.Text(modText.KEY_PICKER_CLIENT_SEARCH)

    Set mSearchBox = AddSearchBox("txtSearch", PICKER_MARGIN, _
                                  PICKER_MARGIN + PICKER_LABEL_HEIGHT + 2, _
                                  PICKER_CLIENT_SEARCH_WIDTH, PICKER_CELL_HEIGHT)

    Set mHintLabel = AddLabel("lblHint", PICKER_MARGIN, HintTop(), _
                              FormWidth() - PICKER_MARGIN * 2, PICKER_LABEL_HEIGHT)
    mHintLabel.Font.Bold = True
End Sub

Private Sub BuildList()
    Set mListBox = Me.Controls.Add("Forms.ListBox.1", "lstClients", True)
    mListBox.Left = PICKER_MARGIN
    mListBox.Top = ListTop()
    mListBox.Width = FormWidth() - PICKER_MARGIN * 2
    mListBox.Height = PICKER_CLIENT_LIST_HEIGHT
    mListBox.Font.Size = 9
    mListBox.ColumnCount = 2
    mListBox.ColumnWidths = "170;80"
End Sub

Private Sub BuildFooter()
    Dim halfWidth As Single

    halfWidth = (FormWidth() - PICKER_MARGIN * 2 - PICKER_GAP) / 2

    AddButton ROLE_OK, vbNullString, modText.Text(modText.KEY_PICKER_OK), _
              PICKER_MARGIN, FooterTop(), halfWidth, PICKER_HEADER_HEIGHT
    AddButton ROLE_CANCEL, vbNullString, modText.Text(modText.KEY_PICKER_CANCEL), _
              PICKER_MARGIN + halfWidth + PICKER_GAP, FooterTop(), halfWidth, _
              PICKER_HEADER_HEIGHT
End Sub

' =====================================================================
'  geometry
' =====================================================================

Private Function FormWidth() As Single
    FormWidth = PICKER_CLIENT_SEARCH_WIDTH + PICKER_MARGIN * 2
End Function

Private Function ListTop() As Single
    ListTop = PICKER_MARGIN + PICKER_LABEL_HEIGHT + 2 + PICKER_CELL_HEIGHT + PICKER_GAP
End Function

Private Function HintTop() As Single
    HintTop = ListTop() + PICKER_CLIENT_LIST_HEIGHT + 2
End Function

Private Function FooterTop() As Single
    FooterTop = HintTop() + PICKER_LABEL_HEIGHT + PICKER_GAP
End Function

' =====================================================================
'  building one control
' =====================================================================

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

' The search box is wrapped by clsPickerBox: without the wrapper its Change event
' would never fire, because the designer never held the control (§9.5).
Private Function AddSearchBox(ByVal controlName As String, ByVal leftValue As Single, _
                              ByVal topValue As Single, ByVal widthValue As Single, _
                              ByVal heightValue As Single) As MSForms.TextBox
    Dim box As MSForms.TextBox
    Dim wrapper As clsPickerBox

    Set box = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    box.Left = leftValue
    box.Top = topValue
    box.Width = widthValue
    box.Height = heightValue
    box.Font.Size = 10

    Set wrapper = New clsPickerBox
    wrapper.Attach box, Me
    mBoxes.Add wrapper

    Set AddSearchBox = box
End Function

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
