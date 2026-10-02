VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPickerBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Service_Manager - one runtime search box of a picker (§9.5)
'
'  The same reason as clsPickerButton: a control the designer never held gets no
'  compiled event procedure. clsPickerButton carries the clicks, this class
'  carries the typing - a search has to follow every keystroke, so a click-only
'  wrapper is not enough.
' =====================================================================

Public WithEvents PickerBox As MSForms.TextBox

Private mOwner As Object

Public Sub Attach(ByRef box As MSForms.TextBox, ByRef ownerForm As Object)
    Set PickerBox = box
    Set mOwner = ownerForm
End Sub

Private Sub PickerBox_Change()
    On Error Resume Next
    mOwner.SearchChanged PickerBox.Text
    On Error GoTo 0
End Sub
