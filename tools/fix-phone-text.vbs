Option Explicit

' =====================================================================
'  Service_Manager - makes the phone cells text in the template
'
'  A phone like 0888123456 written into a General cell becomes the number
'  888123456 and the leading zero is gone (24.09.2026). This script sets the "text"
'  format where a phone lives:
'    - the input cell i_client_phone,
'    - the phone column of the ClientsAdmin list (found by its header, not by a
'      fixed address, so a restyled sheet is fine),
'    - the phone column of db_clients (the same way).
'
'      cscript //nologo tools\fix-phone-text.vbs
'
'  It is idempotent: cells that already hold text are reported and nothing changes.
' =====================================================================

Const TEMPLATE_FILE = "Service_Manager.xlsx"
Const TEXT_FORMAT = "@"
Const PHONE_HEADER = "phone"

Dim gFso, gRoot, gXl, gWb
Dim gChanged

Set gFso = CreateObject("Scripting.FileSystemObject")
gRoot = gFso.GetParentFolderName(gFso.GetParentFolderName(WScript.ScriptFullName))
Set gXl = CreateObject("Excel.Application")
gXl.DisplayAlerts = False
gXl.ScreenUpdating = False
Set gWb = gXl.Workbooks.Open(gFso.BuildPath(gRoot, TEMPLATE_FILE))

gChanged = 0

FixNamedCell "i_client_phone"
FixHeaderColumn "ClientsAdmin", "grid_clients_first", "grid_clients_last"
FixHeaderColumn "db_clients", "", ""

gWb.Save
WScript.Echo "phone cells made text: " & CStr(gChanged)
WScript.Echo "template saved: " & gFso.BuildPath(gRoot, TEMPLATE_FILE)
gWb.Close False
gXl.Quit

' =====================================================================
'  the work
' =====================================================================

Sub FixNamedCell(ByVal nameText)
    Dim rng

    On Error Resume Next
    Set rng = Nothing
    Set rng = gWb.Names(nameText).RefersToRange
    On Error GoTo 0

    If rng Is Nothing Then
        WScript.Echo "  no name " & nameText
        Exit Sub
    End If

    If StrComp(CStr(rng.NumberFormat), TEXT_FORMAT, vbTextCompare) = 0 Then
        WScript.Echo "  " & nameText & " is text already"
        Exit Sub
    End If

    rng.NumberFormat = TEXT_FORMAT
    gChanged = gChanged + 1
    WScript.Echo "  " & nameText & " (" & rng.Address(False, False) & ") -> text"
End Sub

' The column whose header says "phone", between the two anchors (or from the header
' row down when there is no anchor). Addresses are never assumed: the header is the
' contract.
Sub FixHeaderColumn(ByVal sheetName, ByVal firstAnchor, ByVal lastAnchor)
    Dim ws, headerRow, firstRow, lastRow, colIndex, headerText, rng

    On Error Resume Next
    Set ws = Nothing
    Set ws = gWb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        WScript.Echo "  no sheet " & sheetName
        Exit Sub
    End If

    If Len(firstAnchor) > 0 Then
        firstRow = gWb.Names(firstAnchor).RefersToRange.Row
        lastRow = gWb.Names(lastAnchor).RefersToRange.Row
        headerRow = firstRow - 1
    Else
        headerRow = 1
        firstRow = 2
        lastRow = ws.UsedRange.Rows.Count
    End If

    For colIndex = 1 To 30
        headerText = LCase$(Trim$(CStr(ws.Cells(headerRow, colIndex).Text)))
        If headerText = PHONE_HEADER Then
            Set rng = ws.Range(ws.Cells(firstRow, colIndex), ws.Cells(lastRow, colIndex))
            If StrComp(CStr(rng.NumberFormat), TEXT_FORMAT, vbTextCompare) = 0 Then
                WScript.Echo "  " & sheetName & "!" & rng.Address(False, False) & " is text already"
            Else
                rng.NumberFormat = TEXT_FORMAT
                gChanged = gChanged + 1
                WScript.Echo "  " & sheetName & "!" & rng.Address(False, False) & " -> text"
            End If
            Exit Sub
        End If
    Next

    WScript.Echo "  no column with the header '" & PHONE_HEADER & "' on " & sheetName
End Sub
