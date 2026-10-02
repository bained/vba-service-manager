VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPickerButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Service_Manager - one runtime button of a picker (Phase 3)
'
'  The designer of the picker forms is empty on purpose: the code builds every
'  control in UserForm_Initialize (a hand written .frm cannot carry the binary
'  .frx blob of the designer). A control that the designer never held gets no
'  compiled event procedure, so the click arrives through this class
'  (WithEvents) and is passed on as a role text - the form decides what the role
'  means (§9.2, §9.3).
' =====================================================================

Public WithEvents PickerButton As MSForms.CommandButton

' The role of the button (ROLE_* of modConfig) and the value it carries (the ISO
' date of a day, "09" of an hour, "30" of a minute, ...).
Public Role As String
Public ValueText As String

Private mOwner As Object

Public Sub Attach(ByRef button As MSForms.CommandButton, ByRef ownerForm As Object, _
                  ByVal roleText As String, ByVal valueText As String)
    Set PickerButton = button
    Set mOwner = ownerForm
    Role = roleText
    ValueText = valueText
End Sub

' The one click handler of every button of the form. The value lives in the
' wrapper; the Tag (and then the caption) is the fallback that keeps a click
' working even when the wrapper was (re)built later.
Private Sub PickerButton_Click()
    On Error Resume Next
    If Len(ValueText) = 0 Then ValueText = PickerButton.Tag
    If Len(ValueText) = 0 Then ValueText = PickerButton.Caption
    mOwner.ButtonClicked Role, ValueText
    On Error GoTo 0
End Sub
