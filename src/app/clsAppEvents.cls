VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Service_Manager - the events of the application (Phase 1)
'
'  Two jobs Excel only offers on the Application object: the log is trimmed
'  before every save (so the file cannot grow without a limit) and the session
'  is closed cleanly when the workbook goes away.
'
'  Both handlers check the workbook name: the Application object is shared with
'  every other open workbook.
' =====================================================================

Public WithEvents App As Application

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, _
                                   Cancel As Boolean)
    On Error Resume Next
    If StrComp(Wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then
        modLog.TrimTo LOG_MAX_ROWS
    End If
    On Error GoTo 0
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    If StrComp(Wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then
        modApp.Shutdown
    End If
    On Error GoTo 0
End Sub
