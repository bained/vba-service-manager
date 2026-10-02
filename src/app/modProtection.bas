Attribute VB_Name = "modProtection"
Option Explicit

' =====================================================================
'  Service_Manager - the data sheets in and out of sight (Phase 6)
'
'  §6.5 once planned sheet protection, with PrepareWrite/FinishWrite around every
'  write of modData. The owner decided against it, and the decision holds up:
'  protection costs a pair of calls per write, it stops a hand made correction,
'  and it does not fit a hyperlink driven UI. Hiding does the same job here - the
'  application reads and writes a hidden sheet without noticing anything, and the
'  three db_* sheets are out of the way of the user.
'
'  The two buttons of the Config page are the way in and out; a hidden sheet can
'  also be shown by hand (right click -> Unhide), which is what a single owner
'  doing maintenance wants. Nothing is protected, so nothing becomes slower.
' =====================================================================

' b_lock_data: the data sheets out of sight. Not "Lock": that is a statement of
' VBA (file locking), and a procedure cannot carry it - the lint names it now.
Public Sub HideDataSheets(ByRef data As Object, ByRef result As Object)
    ApplyVisibility xlSheetHidden

    result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_CONFIG_DATA_HIDDEN)
    result(modRoutes.RESULT_OK) = True
End Sub

' b_unlock_data: back in sight, for a look or a manual edit.
Public Sub ShowDataSheets(ByRef data As Object, ByRef result As Object)
    ApplyVisibility xlSheetVisible

    result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_CONFIG_DATA_SHOWN)
    result(modRoutes.RESULT_OK) = True
End Sub

' True while every data sheet is hidden (the tests ask; the status text says it).
Public Function DataSheetsHidden() As Boolean
    Dim names As Variant
    Dim index As Long

    names = DataSheetList()
    DataSheetsHidden = True
    For index = 0 To UBound(names)
        If SheetOf(CStr(names(index))).Visible <> xlSheetHidden Then
            DataSheetsHidden = False
            Exit Function
        End If
    Next index
End Function

' The data sheets of the application, as the names they have in the workbook.
Private Function DataSheetList() As Variant
    DataSheetList = Split(modConfig.DataSheetNames(), ",")
End Function

Private Function SheetOf(ByVal sheetName As String) As Worksheet
    Set SheetOf = ThisWorkbook.Worksheets(sheetName)
End Function

Private Function IsDataSheet(ByVal sheetName As String) As Boolean
    Dim names As Variant
    Dim index As Long

    names = DataSheetList()
    For index = 0 To UBound(names)
        If StrComp(CStr(names(index)), sheetName, vbTextCompare) = 0 Then
            IsDataSheet = True
            Exit Function
        End If
    Next index
End Function

' Excel refuses to hide the sheet that is in front. The user presses the button on
' Config, but a test - or a future action - may stand somewhere else.
Private Sub LeaveDataSheet()
    If Not modNames.SheetExists(SHEET_CONFIG) Then Exit Sub
    If IsDataSheet(ThisWorkbook.ActiveSheet.Name) Then
        ThisWorkbook.Worksheets(SHEET_CONFIG).Activate
    End If
End Sub

Private Sub ApplyVisibility(ByVal visibility As Long)
    Dim names As Variant
    Dim index As Long

    LeaveDataSheet
    names = DataSheetList()

    modUi.BeginBulk
    On Error Resume Next
    For index = 0 To UBound(names)
        SheetOf(CStr(names(index))).Visible = visibility
    Next index
    On Error GoTo 0
    Err.Clear
    modUi.EndBulk
End Sub
