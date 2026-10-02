Option Explicit

' Calls one macro of the generated workbook and records how far it got. This
' is the tool for the "VBA compiles procedures lazily" trap: a compile error
' in a procedure that no test calls shows up as a modal VBE dialog, and a
' headless run then hangs until the harness watchdog kills Excel.
'
'     cscript //T:30 //nologo tools\probe-macro.vbs modServiceManager.ModelReport
'     cscript //T:30 //nologo tools\probe-macro.vbs modDashboard.RefreshDashboard
'
' The result file build\probe-result.txt says "started" and, when the macro
' returned, also "finished". A missing "finished" line means the run was
' blocked (dump the dialog with tools\dump-vbe-dialog.ps1 and read the error
' location with tools\vbe-error-location.vbs).
'
' Exit codes: 0 = the probe finished, 1 = the probe was killed mid run.

Dim fso, root, target, resultPath, macroName, macroArgument, xl, wb, resultFile
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))

macroName = ""
macroArgument = ""
If WScript.Arguments.Count > 0 Then macroName = Trim(WScript.Arguments(0))
If WScript.Arguments.Count > 1 Then macroArgument = WScript.Arguments(1)

resultPath = fso.BuildPath(fso.BuildPath(root, "build"), "probe-result.txt")
Set resultFile = fso.CreateTextFile(resultPath, True, True)
resultFile.WriteLine "macro  : " & macroName & IIfText(Len(macroArgument) > 0, " " & macroArgument, "") & "  at " & Now
resultFile.WriteLine "state  : started"
resultFile.Close

If Len(macroName) = 0 Then
    WScript.Echo "usage: cscript //T:30 //nologo tools\probe-macro.vbs <module.procedure> [argument]"
    WScript.Quit 2
End If

target = fso.BuildPath(fso.BuildPath(root, "build"), "Service_Manager.xlsm")
If Not fso.FileExists(target) Then
    target = fso.BuildPath(fso.BuildPath(root, "build"), "Service_Manager.test.xlsm")
End If

If Not fso.FileExists(target) Then
    WScript.Echo "no generated workbook in " & fso.BuildPath(root, "build") & _
                 "; run 'cscript //nologo tools\build.vbs build' first"
    WScript.Quit 2
End If

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.ScreenUpdating = False
xl.EnableEvents = False
xl.AutomationSecurity = 1

Set wb = xl.Workbooks.Open(target)
If Len(macroArgument) > 0 Then
    xl.Run "'" & wb.Name & "'!" & macroName, macroArgument
Else
    xl.Run "'" & wb.Name & "'!" & macroName
End If

Set resultFile = fso.OpenTextFile(resultPath, 8, True, True)
resultFile.WriteLine "state  : finished"
resultFile.Close

wb.Close False
xl.Quit

WScript.Echo "probe finished: " & macroName & " (see " & resultPath & ")"
WScript.Quit 0

Function IIfText(ByVal condition, ByVal whenTrue, ByVal whenFalse)
    If condition Then
        IIfText = whenTrue
    Else
        IIfText = whenFalse
    End If
End Function
