Option Explicit

' =====================================================================
'  Service_Manager - adds the two edit-state cells to the template
'
'  The [Edit] link of the two admin lists needs one cell per sheet to remember
'  which row the form is editing - the same trick as s_edit_id on Booking (6.10).
'  The template is designed by hand now, so this script adds those two cells and
'  nothing else. Run it twice and the second run only reports what is there.
'
'      cscript //nologo tools\add-admin-edit.vbs
'
'  Each state cell goes on the first cell to the right of the status line of its
'  sheet that belongs to nobody, styled like s_edit_id.
' =====================================================================

Const XL_PASTE_FORMATS = -4122
Const SEARCH_LIMIT = 5
Const STYLE_NAME = "s_edit_id"

Dim gFso, gRoot, gPath, gXl, gWb, gTargets, gIndex

Set gFso = CreateObject("Scripting.FileSystemObject")
gRoot = gFso.GetParentFolderName(gFso.GetParentFolderName(WScript.ScriptFullName))
gPath = gFso.BuildPath(gRoot, "Service_Manager.xlsx")

If Not gFso.FileExists(gPath) Then
    WScript.Echo "template not found: " & gPath
    WScript.Quit 1
End If

Set gXl = CreateObject("Excel.Application")
gXl.DisplayAlerts = False
gXl.ScreenUpdating = False

Err.Clear
Set gWb = gXl.Workbooks.Open(gPath)
If Err.Number <> 0 Then
    WScript.Echo "cannot open " & gPath & ": " & Err.Number & " " & Err.Description
    gXl.Quit
    WScript.Quit 1
End If

gTargets = Array( _
    Array("s_edit_client_id", "o_status_clients"), _
    Array("s_edit_service_id", "o_status_services"))

For gIndex = 0 To UBound(gTargets)
    AddStateCell CStr(gTargets(gIndex)(0)), CStr(gTargets(gIndex)(1))
Next

gWb.Save
WScript.Echo "template saved: " & gPath
gWb.Close False
gXl.Quit

' =====================================================================
'  the work
' =====================================================================

Sub AddStateCell(ByVal nameText, ByVal anchorName)
    Dim anchorRange, target, styleCell, index, candidate

    If NameExists(nameText) Then
        WScript.Echo "  " & nameText & " is already there (" & NameAddress(nameText) & ")"
        Exit Sub
    End If

    If Not NameExists(anchorName) Then
        WScript.Echo "  no name " & anchorName & " - " & nameText & " was not added"
        Exit Sub
    End If

    Set anchorRange = gWb.Names(anchorName).RefersToRange
    Set target = Nothing

    For index = 1 To SEARCH_LIMIT
        Set candidate = anchorRange.Worksheet.Cells(anchorRange.Row, anchorRange.Column + index)
        If IsFreeCell(candidate) Then
            Set target = candidate
            Exit For
        End If
    Next

    If target Is Nothing Then
        WScript.Echo "  no free cell next to " & anchorName & " (" & anchorRange.Address & _
                     ") - add " & nameText & " by hand"
        Exit Sub
    End If

    Set styleCell = Nothing
    If NameExists(STYLE_NAME) Then Set styleCell = gWb.Names(STYLE_NAME).RefersToRange

    If Not styleCell Is Nothing Then
        styleCell.Copy
        target.PasteSpecial XL_PASTE_FORMATS
        gXl.CutCopyMode = False
    End If
    target.ClearContents
    target.Value = vbNullString

    Err.Clear
    gWb.Names.Add nameText, target
    If Err.Number <> 0 Then
        WScript.Echo "  cannot add " & nameText & ": " & Err.Number & " " & Err.Description
        Exit Sub
    End If

    WScript.Echo "  " & nameText & " on " & target.Address(False, False) & _
                 " (" & anchorRange.Worksheet.Name & ")"
End Sub

' =====================================================================
'  helpers
' =====================================================================

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

Function NameAddress(ByVal nameText)
    NameAddress = CStr(gWb.Names(nameText).RefersToRange.Address(False, False))
End Function

' True when no name of the workbook points at (part of) this cell, the cell is not
' merged, holds no formula and carries no hyperlink yet.
Function IsFreeCell(ByRef cell)
    Dim item, rng

    IsFreeCell = False
    If cell.MergeCells Then Exit Function
    If cell.Hyperlinks.Count > 0 Then Exit Function
    If Len(Trim(CStr(cell.Formula))) > 0 Then Exit Function

    For Each item In gWb.Names
        On Error Resume Next
        Set rng = Nothing
        Set rng = item.RefersToRange
        On Error GoTo 0
        If Not rng Is Nothing Then
            If Overlaps(rng, cell) Then Exit Function
        End If
    Next

    IsFreeCell = True
End Function

' Application.Intersect is not reachable through late binding in VBScript, so the
' overlap of two ranges is worked out by hand.
Function Overlaps(ByRef firstRange, ByRef secondRange)
    Overlaps = False
    If firstRange.Worksheet.Name <> secondRange.Worksheet.Name Then Exit Function
    If firstRange.Row > secondRange.Row + secondRange.Rows.Count - 1 Then Exit Function
    If secondRange.Row > firstRange.Row + firstRange.Rows.Count - 1 Then Exit Function
    If firstRange.Column > secondRange.Column + secondRange.Columns.Count - 1 Then Exit Function
    If secondRange.Column > firstRange.Column + firstRange.Columns.Count - 1 Then Exit Function

    Overlaps = True
End Function
