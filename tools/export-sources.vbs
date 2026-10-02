Option Explicit

' =====================================================================
'  Service_Manager - export the VBA sources of a workbook back to disk
'
'  Why this exists:
'    The build harness injects .cls/.bas files into a workbook, so the
'    workbook is the only place where the imported VBA code is stored
'    after a build. This script performs the reverse direction: it opens
'    an .xlsm/.xlsx file and writes every code component back as a
'    .cls/.bas file, which is how the library sources (SqlTable.cls,
'    SqlQuery.cls, SqlMutation.cls) were recovered on the home machine.
'
'  Usage:
'    cscript //nologo tools\export-sources.vbs <workbook> <targetFolder> [/force]
'
'  Exit codes: 0 = success, 2 = harness/environment error
'
'  Requirements:
'    - Excel installed, "Trust access to the VBA project object model"
'      enabled (File > Options > Trust Center > Trust Center Settings >
'      Macro Settings)
' =====================================================================

Const EXIT_OK = 0
Const EXIT_ABORTED = 2

Const XL_UPDATE_LINKS_NEVER = 0
const XL_AUTOMATION_SECURITY_LOW = 1

Const VBE_CT_STANDARD = 1       ' vbext_ct_StdModule  -> .bas
Const VBE_CT_CLASS = 2          ' vbext_ct_ClassModule -> .cls
Const VBE_CT_MSFORM = 3         ' vbext_ct_MSForm     -> .frm
Const VBE_CT_DOCUMENT = 100     ' sheet / ThisWorkbook -> skipped

Dim gFso, gXl, gWb, gPath, gTarget, gForce, gExported

Set gFso = CreateObject("Scripting.FileSystemObject")
gExported = 0

Main
WScript.Quit EXIT_OK

Sub Main()
    Dim arguments
    Set arguments = WScript.Arguments

    If arguments.Count < 2 Then
        WScript.Echo "usage: cscript //nologo tools\export-sources.vbs <workbook> <targetFolder> [/force]"
        WScript.Quit EXIT_ABORTED
    End If

    gPath = gFso.GetAbsolutePathName(arguments(0))
    gTarget = gFso.GetAbsolutePathName(arguments(1))
    gForce = False
    If arguments.Count > 2 Then gForce = (LCase(arguments(2)) = "/force")

    If Not gFso.FileExists(gPath) Then Abort "workbook not found: " & gPath
    If gFso.FolderExists(gTarget) And Not gForce Then
        Abort "target folder already exists: " & gTarget & " (pass /force to overwrite files)"
    End If
    If Not gFso.FolderExists(gTarget) Then
        Err.Clear
        gFso.CreateFolder gTarget
        If Err.Number <> 0 Then Abort "cannot create " & gTarget & ": " & Err.Description
    End If

    If Not StartExcel() Then Exit Sub
    If Not OpenWorkbook() Then Exit Sub

    ExportComponents
    Log "exported " & CStr(gExported) & " component(s) to " & gTarget

    CleanupExcel
End Sub

Function StartExcel()
    StartExcel = False
    Err.Clear
    Set gXl = CreateObject("Excel.Application")
    If Err.Number <> 0 Then Abort "cannot start Excel: " & Err.Number & " " & Err.Description

    On Error Resume Next
    gXl.Visible = False
    gXl.DisplayAlerts = False
    gXl.EnableEvents = False
    gXl.AskToUpdateLinks = False
    gXl.AutomationSecurity = XL_AUTOMATION_SECURITY_LOW
    gXl.Interactive = False
    On Error GoTo 0

    Log "Excel started: version " & gXl.Version
    StartExcel = True
End Function

Function OpenWorkbook()
    OpenWorkbook = False
    Set gWb = Nothing
    Err.Clear
    Set gWb = gXl.Workbooks.Open(gPath, XL_UPDATE_LINKS_NEVER, True)
    If Err.Number <> 0 Then Abort "cannot open " & gPath & ": " & Err.Number & " " & Err.Description

    ' VBProject access raises an untrappable error when the Trust Center
    ' setting is off, so it must be wrapped in On Error Resume Next.
    Dim projectName
    projectName = ""
    Err.Clear
    On Error Resume Next
    projectName = gWb.VBProject.Name
    On Error GoTo 0
    If Err.Number <> 0 Then
        Abort "VBProject is not accessible. Enable 'Trust access to the VBA project object model'. (" & Err.Description & ")"
    End If

    Log "opened " & gWb.Name & " (VBProject: " & projectName & ")"
    OpenWorkbook = True
End Function

Sub ExportComponents()
    Dim components, index, component, fileName, path, extension

    Set components = gWb.VBProject.VBComponents
    For index = 1 To components.Count
        Set component = components.Item(index)

        extension = ""
        Select Case component.Type
            Case VBE_CT_STANDARD
                extension = ".bas"
            Case VBE_CT_CLASS
                extension = ".cls"
            Case VBE_CT_MSFORM
                extension = ".frm"
            Case Else
                Log "skipped document component " & component.Name
        End Select

        If extension <> "" Then
            path = gFso.BuildPath(gTarget, component.Name & extension)
            Err.Clear
            component.Export path
            If Err.Number <> 0 Then Abort "export failed for " & component.Name & ": " & Err.Number & " " & Err.Description

            gExported = gExported + 1
            Log "exported " & component.Name & extension & " (lines: " & component.CodeModule.CountOfLines & ")"
        End If
    Next
End Sub

Sub Log(ByVal text)
    WScript.Echo text
End Sub

Sub Abort(ByVal text)
    WScript.Echo "ABORTED: " & text
    CleanupExcel
    WScript.Quit EXIT_ABORTED
End Sub

Sub CleanupExcel()
    On Error Resume Next

    If Not gWb Is Nothing Then
        gWb.Close False
        Set gWb = Nothing
    End If
    If Not gXl Is Nothing Then
        gXl.Quit
        Set gXl = Nothing
    End If

    On Error GoTo 0
End Sub
