Option Explicit

' =====================================================================
'  Service_Manager - import probe (diagnostic)
'
'  Copies nothing and changes nothing of the template: it opens the staged copy
'  of the last build and imports every library and application source, one by
'  one, printing the first failure with its error number. Use it when
'  tools\build.vbs stops at VBComponents.Import with "Unknown runtime error".
'
'      cscript //nologo tools\build.vbs build      (creates the staged copy)
'      cscript //nologo tools\probe-import.vbs
' =====================================================================

Dim gFso, gRoot, gXl, gWb, gFolderNames, gIndex, gFile, gComponent
Dim gFailures, gCount

Set gFso = CreateObject("Scripting.FileSystemObject")
gRoot = gFso.GetParentFolderName(gFso.GetParentFolderName(WScript.ScriptFullName))

Set gXl = CreateObject("Excel.Application")
gXl.DisplayAlerts = False
gXl.Visible = False
gXl.AutomationSecurity = 1

Set gWb = gXl.Workbooks.Open(gFso.BuildPath(gFso.BuildPath(gRoot, "build"), "_template_staging.xlsx"))

gFolderNames = Array("lib", "src\app", "src\tests")
gCount = 0

For gIndex = 0 To UBound(gFolderNames)
    For Each gFile In gFso.GetFolder(gFso.BuildPath(gRoot, CStr(gFolderNames(gIndex)))).Files
        If IsModuleFile(gFile.Path) Then
            ImportOne gFile.Path
        End If
    Next
Next

WScript.Echo ""
WScript.Echo "imported: " & CStr(gCount) & ", failures: " & CStr(gFailures)

gWb.Close False
gXl.Quit
WScript.Quit 0

Function IsModuleFile(ByVal filePath)
    Dim extension
    extension = LCase(gFso.GetExtensionName(filePath))
    IsModuleFile = (extension = "bas" Or extension = "cls")
End Function

Sub ImportOne(ByVal filePath)
    Dim component, errNumber, errDescription, componentName, before

    componentName = vbNullString
    errNumber = 0
    errDescription = vbNullString
    before = gWb.VBProject.VBComponents.Count

    Err.Clear
    On Error Resume Next
    Set component = gWb.VBProject.VBComponents.Import(filePath)
    errNumber = Err.Number
    errDescription = Err.Description
    If Not component Is Nothing Then componentName = component.Name
    On Error GoTo 0

    If errNumber <> 0 Or Len(componentName) = 0 Then
        gFailures = gFailures + 1
        WScript.Echo "FAILED  " & gFso.GetFileName(filePath) & " : " & CStr(errNumber) & _
                     " 0x" & Hex(errNumber) & " " & errDescription & _
                     " (components before: " & CStr(before) & _
                     ", now: " & CStr(gWb.VBProject.VBComponents.Count) & ")"
        Err.Clear
    Else
        gCount = gCount + 1
        WScript.Echo "ok      " & gFso.GetFileName(filePath) & " -> " & componentName & _
                     " (lines: " & CStr(gWb.VBProject.VBComponents(componentName).CodeModule.CountOfLines) & ")"
    End If
End Sub
