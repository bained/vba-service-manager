VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Service_Manager - the events of one UI sheet (Phase 1)
'
'  Variant B of PROJECT_CONCEPT §6.2.2: a single class with WithEvents instead
'  of one stub module per sheet. modApp.AttachSheetEvents creates one instance
'  per UI sheet and keeps them in a module level collection: a local instance
'  would be released at the end of the procedure and the events would stop
'  arriving.
'
'  The three handlers only pass the name of the sheet on; all decisions are in
'  modApp, so the same flow is reachable from the harness and from a test.
' =====================================================================

Public WithEvents Ws As Worksheet

' A click on a cell button (a hyperlink anchored at the name of the cell).
Private Sub Ws_FollowHyperlink(ByVal Target As Hyperlink)
    If modApp.IsBusy() Then Exit Sub
    On Error Resume Next
    modApp.OnHyperlink Ws.Name, Target
    On Error GoTo 0
End Sub

' A single click inside the sheet.
Private Sub Ws_SelectionChange(ByVal Target As Range)
    If modApp.IsBusy() Then Exit Sub
    On Error Resume Next
    modApp.OnSelection Ws.Name, Target
    On Error GoTo 0
End Sub

' The sheet came to the front: it is drawn again (§6.7).
Private Sub Ws_Activate()
    If modApp.IsBusy() Then Exit Sub
    On Error Resume Next
    modApp.OnActivate Ws.Name
    On Error GoTo 0
End Sub

' A cell was edited by hand: the sheet decides what follows (§9.4).
Private Sub Ws_Change(ByVal Target As Range)
    If modApp.IsBusy() Then Exit Sub
    On Error Resume Next
    modApp.OnChange Ws.Name, Target
    On Error GoTo 0
End Sub

' A double click inside the sheet.
Private Sub Ws_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If modApp.IsBusy() Then Exit Sub
    On Error Resume Next
    modApp.OnDoubleClick Ws.Name, Target, Cancel
    On Error GoTo 0
End Sub

' Releases the sheet: called by modApp.DetachSheetEvents.
Public Sub Detach()
    Set Ws = Nothing
End Sub
