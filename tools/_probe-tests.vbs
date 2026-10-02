Option Explicit

' =====================================================================
'  Service_Manager - probe: runs the test batch of the generated workbook
'
'      cscript //T:120 //nologo tools\_probe-tests.vbs
'
'  Unlike tools\build.vbs this probe keeps the Excel window visible and reports
'  the error Application.Run answers with, so a failure that the harness only
'  shows as "Exception occurred" can be named. While a modal dialog is up the
'  run blocks and tools\dump-vbe-dialog.ps1 / tools\vbe-error-location.vbs can
'  read it. It is a diagnostic tool of a work session, not part of the build.
' =====================================================================

Dim fso, root, target, xl, wb, answer, errNumber, errText, logPath, logFile, mimic

Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
target = fso.BuildPath(fso.BuildPath(root, "build"), "Service_Manager.test.xlsm")
logPath = fso.BuildPath(fso.BuildPath(root, "build"), "probe-tests.txt")

' "harness" mimics the session of tools\build.vbs: the same application flags and
' no Workbook_Open, because the harness opens the macro free template and imports
' the code into it.
mimic = ""
If WScript.Arguments.Count > 0 Then mimic = LCase(WScript.Arguments(0))

Set logFile = fso.CreateTextFile(logPath, True, True)
logFile.WriteLine "probe at " & Now
logFile.WriteLine "mode  : " & mimic

If Not fso.FileExists(target) Then
    logFile.WriteLine "no test workbook: run 'cscript //nologo tools\build.vbs test' first"
    logFile.Close
    WScript.Quit 2
End If

Set xl = CreateObject("Excel.Application")
If mimic = "harness" Then
    xl.Visible = True
    xl.Interactive = True
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    xl.EnableEvents = False
    xl.AskToUpdateLinks = False
    xl.AutomationSecurity = 1
End If

Set wb = xl.Workbooks.Open(target)
logFile.WriteLine "opened " & wb.Name
logFile.Close

Set logFile = fso.OpenTextFile(logPath, 8, True)

On Error Resume Next
answer = xl.Run("'" & wb.Name & "'!modTests.RunAllTests")
errNumber = Err.Number
errText = Err.Description
On Error GoTo 0

If errNumber <> 0 Then
    logFile.WriteLine "run failed : " & CStr(errNumber) & " " & errText
Else
    logFile.WriteLine "run ok     : " & CStr(answer)
End If

logFile.WriteLine "finished at " & Now
logFile.Close
wb.Close False
xl.Quit
