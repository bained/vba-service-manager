Option Explicit

' Reads the module and the line that the VBE has selected, so the location of
' a compile error can be read while the modal "Compile error" dialog is open.
' Run it next to tools\dump-vbe-dialog.ps1 (which prints the dialog text):
'
'     cscript //T:20 //nologo tools\vbe-error-location.vbs
'
' The output goes to build\vbe-error-location.txt. The call works while the
' dialog is up, because Application.VBE lives in the same process.

Dim fso, root, outPath, outFile, xl, pane, startLine, startColumn, endLine, endColumn
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
outPath = fso.BuildPath(fso.BuildPath(root, "build"), "vbe-error-location.txt")

Set outFile = fso.CreateTextFile(outPath, True, True)
outFile.WriteLine "VBE error location at " & Now

On Error Resume Next
Set xl = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    outFile.WriteLine "no running Excel instance: " & Err.Number & " " & Err.Description
    Err.Clear
Else
    Set pane = xl.VBE.ActiveCodePane
    If Err.Number <> 0 Then
        outFile.WriteLine "ActiveCodePane failed: " & Err.Number & " " & Err.Description
        Err.Clear
    Else
        pane.GetSelection startLine, startColumn, endLine, endColumn
        outFile.WriteLine "module : " & pane.CodeModule.Name
        outFile.WriteLine "line   : " & startLine & " (column " & startColumn & ")"
        outFile.WriteLine "text   : " & pane.CodeModule.Lines(startLine, 1)
    End If
End If
On Error GoTo 0

outFile.WriteLine "state  : finished"
outFile.Close

WScript.Echo "written: " & outPath
