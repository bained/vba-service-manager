Option Explicit

' =====================================================================
'  Service_Manager - adds the Appointments sheet to the template
'
'  The code of the page (modUiAppointments, the routes, the texts, the names) is
'  already there; the template still has to carry the sheet and its anchors. This
'  script writes exactly that and nothing else, so the design stays yours.
'
'      cscript //nologo tools\add-appointments-sheet.vbs
'
'  What it writes:
'    - the sheet Appointments, right after ServicesAdmin;
'    - the title (o_appointments_title), the status line (o_status_appointments)
'      and the first grid row (grid_appointments_first);
'    - three cell buttons with the look of an existing one: Refresh, Appointments
'      and Dashboard (b_appointments_refresh / b_go_appointments /
'      b_home_appointments);
'    - the header row of eleven columns, exactly the order of APPOINTMENTS_COL_*;
'    - the text format on date, start, end, duration and phone, so a leading zero
'      and an ISO date survive (24.09.2026).
'
'  Run it twice: the second run only reports that the sheet is there.
' =====================================================================

Const TEMPLATE_FILE = "Service_Manager.xlsx"
Const SHEET_NAME = "Appointments"
Const AFTER_SHEET = "ServicesAdmin"
Const STYLE_NAME = "b_booking_clear"
Const TEXT_FORMAT = "@"

Dim gFso, gRoot, gXl, gWb, gWs, gStyle, index

Set gFso = CreateObject("Scripting.FileSystemObject")
gRoot = gFso.GetParentFolderName(gFso.GetParentFolderName(WScript.ScriptFullName))

Set gXl = CreateObject("Excel.Application")
gXl.DisplayAlerts = False
gXl.ScreenUpdating = False
Set gWb = gXl.Workbooks.Open(gFso.BuildPath(gRoot, TEMPLATE_FILE))

If Not SheetExists(AFTER_SHEET) Then
    WScript.Echo "the template carries no sheet " & AFTER_SHEET & "; nothing was changed"
    Finish 1
End If

If SheetExists(SHEET_NAME) Then
    WScript.Echo SHEET_NAME & " is already in the template; nothing was changed"
    Finish 0
End If

Set gStyle = Nothing
If NameExists(STYLE_NAME) Then Set gStyle = gWb.Names(STYLE_NAME).RefersToRange

Set gWs = gWb.Worksheets.Add(, gWb.Worksheets(AFTER_SHEET))
gWs.Name = SHEET_NAME

gWs.Range("A1").Value = SHEET_NAME
gWs.Range("A1").Font.Bold = True
gWs.Range("A1").Font.Size = 14
gWs.Range("A2").Value = "Every appointment, newest first. Refresh rebuilds the list."
gWs.Range("A3").Value = "Ready."

AddName "o_appointments_title", gWs.Range("A1")
AddName "o_status_appointments", gWs.Range("A3")
AddButton "b_appointments_refresh", "D2", "Refresh", "Rebuild the whole list"
AddButton "b_go_appointments", "D3", "Appointments", "Go to the appointments"
AddButton "b_home_appointments", "D4", "Dashboard", "Back to the Dashboard"

Dim headers
headers = Array("id", "date", "start", "end", "duration", "client", "phone", _
                "email", "service", "[Edit]", "[Delete]")
For index = 0 To UBound(headers)
    gWs.Cells(5, index + 1).Value = headers(index)
    gWs.Cells(5, index + 1).Font.Bold = True
Next

' Date, start, end, duration and phone are text: a General cell turns 09:00 into a
' time and 0888123456 into a number without its leading zero.
For index = 2 To 7
    gWs.Columns(index).NumberFormat = TEXT_FORMAT
Next

AddName "grid_appointments_first", gWs.Range("A6")
gWs.Activate
gWs.Range("A6").Select

gWb.Save
WScript.Echo "sheet written: " & SHEET_NAME & " (after " & AFTER_SHEET & ")"
WScript.Echo "names: o_appointments_title, o_status_appointments, b_appointments_refresh,"
WScript.Echo "       b_go_appointments, b_home_appointments, grid_appointments_first"
WScript.Echo "next : cscript //nologo tools\\build.vbs inspect"
Finish 0

' =====================================================================
'  helpers
' =====================================================================

Sub AddName(ByVal nameText, ByRef target)
    Err.Clear
    On Error Resume Next
    gWb.Names(nameText).Delete
    On Error GoTo 0

    Err.Clear
    gWb.Names.Add nameText, target
    If Err.Number <> 0 Then
        WScript.Echo "  cannot add " & nameText & ": " & Err.Number & " " & Err.Description
    Else
        WScript.Echo "  " & nameText & " -> " & target.Address(False, False)
    End If
End Sub

' A cell button: the caption, the look of an existing button and the hyperlink
' whose SubAddress is its own name (that is how every button of this workbook works).
Sub AddButton(ByVal nameText, ByVal cellAddress, ByVal captionText, ByVal tipText)
    Dim target

    Set target = gWs.Range(cellAddress)

    If Not gStyle Is Nothing Then
        gStyle.Copy
        target.PasteSpecial -4122
        gXl.CutCopyMode = False
    End If

    target.ClearContents
    target.Value = captionText

    On Error Resume Next
    target.Hyperlinks.Delete
    On Error GoTo 0

    Err.Clear
    target.Hyperlinks.Add target, "", nameText, tipText, captionText
    If Err.Number <> 0 Then WScript.Echo "  cannot add the hyperlink of " & nameText

    AddName nameText, target
End Sub

Function SheetExists(ByVal sheetName)
    Dim item

    SheetExists = False
    For Each item In gWb.Worksheets
        If StrComp(CStr(item.Name), sheetName, vbTextCompare) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

Function NameExists(ByVal nameText)
    Dim item

    NameExists = False
    For Each item In gWb.Names
        If StrComp(CStr(item.Name), nameText, vbTextCompare) = 0 Then
            NameExists = True
            Exit Function
        End If
    Next
End Function

Sub Finish(ByVal exitCode)
    gWb.Close False
    gXl.Quit
    WScript.Quit exitCode
End Sub
