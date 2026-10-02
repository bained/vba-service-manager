Option Explicit

' =====================================================================
'  Service_Manager - puts the client search button into the template
'
'  The template is designed by hand in Excel now, so this script does not
'  regenerate it: it adds the one cell the typed client search (9.5) needs and
'  touches nothing else. Run it twice and the second run only reports that the
'  button is already there.
'
'      cscript //nologo tools\add-client-search.vbs
'
'  What it writes:
'    1. the name b_booking_client_pick on the first free cell of the row that
'       holds i_booking_client (found by name, never by address),
'    2. the format of an existing cell button (copied, so the new one looks like
'       the buttons around it),
'    3. the hyperlink whose SubAddress is the name of the cell - the way every
'       other cell button of this workbook works,
'    4. the workbook scoped name itself, which modNames.VerifyContract checks.
'
'  The same cell is added by tools\make-template.vbs, so a regenerated template
'  carries it as well.
' =====================================================================

Const TEMPLATE_FILE = "Service_Manager.xlsx"
Const BUTTON_NAME = "b_booking_client_pick"
Const BUTTON_CAPTION = "Search"
Const BUTTON_TIP = "Search the clients by any part of the name or phone"
Const ANCHOR_NAME = "i_booking_client"
Const STYLE_NAME = "b_booking_dur_pick"
Const XL_PASTE_FORMATS = -4122
Const SEARCH_LIMIT = 5

Dim gFso, gRoot, gPath, gXl, gWb, gChanged

Set gFso = CreateObject("Scripting.FileSystemObject")
gRoot = gFso.GetParentFolderName(gFso.GetParentFolderName(WScript.ScriptFullName))
gPath = gFso.BuildPath(gRoot, TEMPLATE_FILE)

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

' A button on a cell that another name owns is the bug of 19.09.2026: the caption
' is overwritten (o_booking_client_id writes an id into its own cell) and an empty
' cell cannot be clicked. So: take whatever is there away and place it again.
If NameExists(BUTTON_NAME) Then
    WScript.Echo "moving " & BUTTON_NAME & " away from " & NameAddress(BUTTON_NAME)
    RemoveButton
End If

If Not NameExists(ANCHOR_NAME) Then
    WScript.Echo "the template carries no name " & ANCHOR_NAME & "; nothing was changed"
    Finish 1
End If

Dim anchorRange, target, styleCell, index, candidate, prefix, subText, bangPos

Set anchorRange = gWb.Names(ANCHOR_NAME).RefersToRange
Set target = Nothing
Set styleCell = Nothing

' The first cell to the right of the client field that belongs to nobody: empty,
' not merged, no hyperlink, no formula and - the check that was missing - not
' covered by any other name of the workbook.
For index = 1 To SEARCH_LIMIT
    Set candidate = anchorRange.Worksheet.Cells(anchorRange.Row, anchorRange.Column + index)
    If IsFreeCell(candidate) Then
        Set target = candidate
        Exit For
    End If
Next

If target Is Nothing Then
    WScript.Echo "no free cell in the row of " & ANCHOR_NAME & " (" & anchorRange.Address & ")"
    WScript.Echo "put the button somewhere by hand and add the name " & BUTTON_NAME & " to that cell"
    Finish 1
End If

If NameExists(STYLE_NAME) Then Set styleCell = gWb.Names(STYLE_NAME).RefersToRange

If Not styleCell Is Nothing Then
    styleCell.Copy
    target.PasteSpecial XL_PASTE_FORMATS
    gXl.CutCopyMode = False
    target.ClearContents
End If

target.Value = BUTTON_CAPTION

' The other buttons of the workbook link to their own name; the workbook prefix
' of that SubAddress is taken from an existing button, so both forms of the link
' behave the same.
prefix = ""
If Not styleCell Is Nothing Then
    If styleCell.Hyperlinks.Count > 0 Then
        subText = CStr(styleCell.Hyperlinks(1).SubAddress)
        bangPos = InStr(subText, "!")
        If bangPos > 0 Then prefix = Left(subText, bangPos)
    End If
End If

On Error Resume Next
target.Hyperlinks.Delete
On Error GoTo 0

Err.Clear
target.Hyperlinks.Add target, "", prefix & BUTTON_NAME, BUTTON_TIP, BUTTON_CAPTION
If Err.Number <> 0 Then
    WScript.Echo "cannot add the hyperlink: " & Err.Number & " " & Err.Description
    Finish 1
End If

Err.Clear
gWb.Names.Add BUTTON_NAME, target
If Err.Number <> 0 Then
    WScript.Echo "cannot add the name: " & Err.Number & " " & Err.Description
    Finish 1
End If

gWb.Save
If Err.Number <> 0 Then
    WScript.Echo "cannot save " & gPath & ": " & Err.Number & " " & Err.Description
    Finish 1
End If

WScript.Echo "button written: " & BUTTON_NAME & " on " & target.Address(False, False) & _
             " (" & anchorRange.Worksheet.Name & ")"
WScript.Echo "link          : " & CStr(target.Hyperlinks(1).SubAddress)
WScript.Echo "name points to: " & NameAddress(BUTTON_NAME)
WScript.Echo "next          : cscript //nologo tools\\build.vbs inspect"
Finish 0

' =====================================================================
'  helpers
' =====================================================================

' Takes an old button away: its hyperlink, its text and its name. The cell is left
' empty, the way it was before the button was put there.
Sub RemoveButton()
    Dim oldRange

    Set oldRange = gWb.Names(BUTTON_NAME).RefersToRange
    On Error Resume Next
    oldRange.Hyperlinks.Delete
    oldRange.ClearContents
    On Error GoTo 0
    gWb.Names(BUTTON_NAME).Delete
End Sub

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

' True when two ranges of the same sheet share at least one cell. Written by hand,
' because Application.Intersect is not reachable through late binding in VBScript
' ("Unable to get the Intersect property of the Application class").
Function Overlaps(ByRef firstRange, ByRef secondRange)
    Overlaps = False
    If firstRange.Worksheet.Name <> secondRange.Worksheet.Name Then Exit Function
    If firstRange.Row > secondRange.Row + secondRange.Rows.Count - 1 Then Exit Function
    If secondRange.Row > firstRange.Row + firstRange.Rows.Count - 1 Then Exit Function
    If firstRange.Column > secondRange.Column + secondRange.Columns.Count - 1 Then Exit Function
    If secondRange.Column > firstRange.Column + firstRange.Columns.Count - 1 Then Exit Function

    Overlaps = True
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

Function NameAddress(ByVal nameText)
    NameAddress = CStr(gWb.Names(nameText).RefersToRange.Address(False, False))
End Function

Sub Finish(ByVal exitCode)
    gWb.Close False
    gXl.Quit
    WScript.Quit exitCode
End Sub
