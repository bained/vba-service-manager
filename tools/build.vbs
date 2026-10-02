Option Explicit

' =====================================================================
'  Service_Manager - build and test harness (Windows Script Host, VBScript)
'
'  What it does:
'    1. copies the macro-free template Service_Manager.xlsx to build\
'    2. injects the VBA sources into the copy through the VBProject
'       automation API (no manual VBE work, no AutoHotkey/AutoIt)
'    3. saves the copy as a macro-enabled workbook (.xlsm)
'    4. in "test" mode runs modTests.RunAllTests inside
'       the workbook and writes build\test-report.txt
'
'  Sources:
'    lib\             vendored sql_emulate library (imported first)
'    src\app\         application modules
'    src\doc\         code of document modules (src\doc\ThisWorkbook.txt), injected
'                     into the matching component of the workbook
'    src\tests\       test modules (test mode only)
'    tools\export-sources.vbs exports the VBA code of a workbook back to disk.
'
'  Usage:
'    cscript //nologo tools\build.vbs build    -> build\Service_Manager.xlsm
'    cscript //nologo tools\build.vbs test     -> tests on build\Service_Manager.test.xlsm
'    cscript //nologo tools\build.vbs inspect  -> build\inspect-report.txt (modApp.WriteDiagnostics)
'    cscript //nologo tools\build.vbs release  -> release\Service_Manager_pre_<version>.xlsm + README.txt
'                                                 (started once, so the shipped file already shows the UI)
'
'  Exit codes: 0 = success, 1 = test failures, 2 = harness/environment error
'
'  A second instance of this script runs as a watchdog: if the test run
'  blocks (a modal VBA dialog from a compile error or an unhandled run-time
'  error), Excel is terminated after WATCHDOG_SECONDS and build\_timeout.txt
'  is written instead of hanging the harness forever.
'
'  Requirements:
'    - Excel installed, "Trust access to the VBA project object model" enabled
'      (File > Options > Trust Center > Trust Center Settings > Macro Settings)
'    - VBA sources must use CRLF line endings and ASCII characters
' =====================================================================

' ---- exit codes
Const EXIT_OK = 0
Const EXIT_TEST_FAILED = 1
Const EXIT_ABORTED = 2

' ---- Excel constants
Const XL_MACRO_WORKBOOK = 52          ' xlOpenXMLWorkbookMacroEnabled
Const XL_AUTOMATION_SECURITY_LOW = 1  ' msoAutomationSecurityLow
Const XL_UPDATE_LINKS_NEVER = 0
Const VBE_CT_DOCUMENT = 100
Const VBE_CT_FORM = 3                 ' vbext_ct_MSForm

' ---- test entry point inside the workbook
Const TEST_MODULE = "modTests"
Const TEST_ENTRY = "RunAllTests"

' ---- inspect entry point inside the workbook
Const INSPECT_MODULE = "modApp"
Const INSPECT_ENTRY = "WriteDiagnostics"

' ---- release build of the pre-production version
' APP_VERSION must match modConfig.APP_VERSION: the harness checks it, so the
' file name, the README and the model report can never disagree.
Const APP_VERSION = "1.1.0"
Const RELEASE_README = "README.txt"
' The draft of a phase is published with this checklist next to it (the file is
' written by hand, never by the build).
Const DRAFT_CHECKLIST = "DRAFT_CHECKLIST.md"
' The sheet (and cell) the published workbook shows when it is opened. The cell is
' found by NAME: cal_first is the first row of the dashboard body, so a new layout
' may move it anywhere. The address is only the fallback for a workbook where the
' name is gone (the contract test would report that anyway).
Const SHEET_DASHBOARD_NAME = "Dashboard"
Const RELEASE_VIEW_NAME = "cal_first"
' The release view cell is looked up by its name (ReleaseViewCell below) - the
' layout of the sheets belongs to the template, so the harness keeps no address.

' ---- watchdog: seconds before a blocked Excel instance is terminated
' This is a hard limit for one harness run, not a progress watchdog: a long
' test batch may legitimately take a while (the batch grew with every phase, so
' the limit is raised with it).
' When Excel is really blocked by a modal dialog, the run is stopped here.
Const WATCHDOG_SECONDS = 600

Dim gFso, gMode, gRoot, gLibDir, gSrcDir, gDocDir, gOutDir, gReleaseDir, gTemplate
Dim gStage, gTarget, gReportPath, gExitCode, gImportedCount
Dim gSentinelPath, gTimeoutPath, gShell, gInspectPath, gExcelVersion, gRunId
Dim gXl, gWb, gLog

Set gFso = CreateObject("Scripting.FileSystemObject")
Set gLog = CreateObject("Scripting.Dictionary")
gExitCode = EXIT_OK
gImportedCount = 0

Main
WScript.Quit gExitCode

' =====================================================================
'  main flow
' =====================================================================
Sub Main()
    Dim testResult
    Dim crosscheckProblem

    ' The watchdog is a second, detached instance of this script. It kills a
    ' blocked Excel instance (modal VBA dialog) instead of hanging forever.
    If WScript.Arguments.Count > 0 Then
        If LCase(WScript.Arguments(0)) = "watchdog" Then
            RunWatchdog
            Exit Sub
        End If
    End If

    If Not ParseCommandLine() Then Exit Sub
    If Not ResolvePaths() Then Exit Sub
    If Not PrepareOutput() Then Exit Sub

    ' The fast lane runs first, on every mode that tests (a release is one of them:
    ' the file is handed over, so it passes the same gate as a draft, Phase 7).
    If gMode = "test" Or gMode = "inspect" Or gMode = "draft" Or gMode = "release" Then
        If Not RunFastLane() Then Exit Sub
    End If

    If gMode = "test" Or gMode = "inspect" Or gMode = "draft" Or gMode = "release" Then StartWatchdog

    If Not StartExcel() Then Exit Sub
    If Not OpenTarget() Then Exit Sub
    If Not RemoveCodeComponents() Then Exit Sub
    If Not ImportModules() Then Exit Sub
    If Not InjectDocumentCode() Then Exit Sub
    If Not SaveAsMacroWorkbook() Then Exit Sub

    ' A draft is published before the tests run: a fresh session publishes cleanly
    ' (the test batch leaves the runtime in its own state). When the tests fail the
    ' draft is removed again, so "a draft exists only when the tests are green"
    ' still holds. A release cannot do that: the file carries no tests (the harness
    ' skips src\tests in release mode), so its gate is the test lane's report -
    ' see VerifyTestReportIsFresh.
    If gMode = "test" Or gMode = "draft" Then
        ' Everything that runs code (the draft's modApp.Boot and the test batch)
        ' happens in a fresh Excel instance: the build session - which imported the
        ' modules and used the VBE - died on the first Application.Run here, while
        ' a fresh instance runs the same file without a problem.
        If Not RestartExcelForTests() Then Exit Sub

        If gMode = "draft" Then
            If Not PublishRelease() Then Exit Sub
        End If

        testResult = RunTests()
        If testResult = "" Then Exit Sub

        ' The two lanes answered the same cases of src\tests\crosscheck.txt; the
        ' harness compares them. A difference is a failing build, not a note.
        crosscheckProblem = CompareCrosscheck()
        If Len(crosscheckProblem) > 0 Then
            testResult = Replace(testResult, "PASSED - ", "FAILED - ", 1, 1) & _
                         "  FAIL: cross check: " & crosscheckProblem & vbCrLf
        End If

        WriteReport testResult
        If InStr(1, testResult, "PASSED - ", vbBinaryCompare) = 1 Then
            gExitCode = EXIT_OK
        Else
            gExitCode = EXIT_TEST_FAILED
        End If
        If gMode = "draft" Then
            If gExitCode <> EXIT_OK Then
                RemovePublishedDraft
                Log "the draft was removed again: the tests did not pass"
            Else
                Log "draft is ready for a manual test"
            End If
        End If
    ElseIf gMode = "inspect" Then
        If Not RunInspect() Then Exit Sub
    ElseIf gMode = "release" Then
        ' The release file carries no tests (src\tests is skipped in release mode -
        ' the file is the deliverable), so the batch cannot run here. The gate is
        ' the report of the lane that does have them, and it must be fresh: a
        ' release cut from sources that no green lane has seen is not a release.
        If Not VerifyTestReportIsFresh() Then Exit Sub
        If Not PublishRelease() Then Exit Sub
        Log "release is ready for a manual test"
    Else
        Log "build finished: " & gTarget
    End If

    StopWatchdog
    CleanupExcel
    Log "exit code: " & CStr(gExitCode)
End Sub

' =====================================================================
'  steps
' =====================================================================
Function ParseCommandLine()
    gMode = "test"
    If WScript.Arguments.Count > 0 Then gMode = LCase(WScript.Arguments(0))

    If gMode = "?" Or gMode = "-?" Or gMode = "/?" Or gMode = "help" Then
        WScript.Echo "usage: cscript //nologo tools\build.vbs [build|test|inspect|release|draft]"
        gExitCode = EXIT_ABORTED
        Exit Function
    End If
    If gMode <> "build" And gMode <> "test" And gMode <> "inspect" And gMode <> "release" _
       And gMode <> "draft" Then
        Abort "unknown mode '" & gMode & "'; use build, test, inspect, release or draft"
    End If
    Log "mode: " & gMode
    ParseCommandLine = True
End Function

Function ResolvePaths()
    Dim scriptPath
    scriptPath = WScript.ScriptFullName
    gRoot = gFso.GetParentFolderName(gFso.GetParentFolderName(scriptPath))  ' parent of tools\
    gSrcDir = gFso.BuildPath(gRoot, "src")
    gDocDir = gFso.BuildPath(gSrcDir, "doc")
    gOutDir = gFso.BuildPath(gRoot, "build")
    gReleaseDir = gFso.BuildPath(gRoot, "release")
    gTemplate = gFso.BuildPath(gRoot, "Service_Manager.xlsx")

    ' The vendored sql_emulate library (SqlTable/SqlQuery/SqlMutation) is
    ' imported before the application sources. Older copies of the project
    ' kept it next to the project folder (the sql_emulate layout), so that
    ' location is still accepted as a fallback.
    gLibDir = gFso.BuildPath(gRoot, "lib")
    If Not gFso.FolderExists(gLibDir) Then gLibDir = gFso.GetParentFolderName(gRoot)
    If Not gFso.FolderExists(gLibDir) Then Abort "library source folder not found: " & gFso.BuildPath(gRoot, "lib")

    If gMode = "build" Then
        gTarget = gFso.BuildPath(gOutDir, "Service_Manager.xlsm")
    ElseIf gMode = "inspect" Then
        gTarget = gFso.BuildPath(gOutDir, "Service_Manager.inspect.xlsm")
    ElseIf gMode = "release" Then
        gTarget = gFso.BuildPath(gOutDir, "Service_Manager.release.xlsm")
    ElseIf gMode = "draft" Then
        gTarget = gFso.BuildPath(gOutDir, "Service_Manager.draft.xlsm")
    Else
        gTarget = gFso.BuildPath(gOutDir, "Service_Manager.test.xlsm")
    End If
    gReportPath = gFso.BuildPath(gOutDir, "test-report.txt")
    gInspectPath = gFso.BuildPath(gOutDir, "inspect-report.txt")
    gStage = gFso.BuildPath(gOutDir, "_template_staging.xlsx")
    gSentinelPath = gFso.BuildPath(gOutDir, "_harness.flag")
    gTimeoutPath = gFso.BuildPath(gOutDir, "_timeout.txt")

    If Not gFso.FileExists(gTemplate) Then Abort "template not found: " & gTemplate
    Log "root: " & gRoot
    Log "library sources: " & gLibDir
    Log "application sources: " & gSrcDir
    ResolvePaths = True
End Function

Function PrepareOutput()
    If Not gFso.FolderExists(gOutDir) Then
        Err.Clear
        gFso.CreateFolder gOutDir
        If Err.Number <> 0 Then Abort "cannot create " & gOutDir & ": " & Err.Description
        Log "created " & gOutDir
    End If

    DeleteIfExists gTarget
    DeleteIfExists gStage

    ' The template is xlsx content, so it must be opened as .xlsx and then
    ' saved as .xlsm; Excel refuses to open xlsx payloads with an .xlsm name.
    Err.Clear
    gFso.CopyFile gTemplate, gStage, True
    If Err.Number <> 0 Then Abort "cannot copy template to " & gStage & ": " & Err.Description
    Log "template staged as " & gStage
    PrepareOutput = True
End Function

Sub DeleteIfExists(ByVal filePath)
    If Not gFso.FileExists(filePath) Then Exit Sub

    Err.Clear
    gFso.DeleteFile filePath, True
    If Err.Number <> 0 Then Abort "cannot delete " & filePath & ": " & Err.Description
    Log "removed previous " & gFso.GetFileName(filePath)
End Sub

Function StartExcel()
    Err.Clear
    Set gXl = CreateObject("Excel.Application")
    If Err.Number <> 0 Then Abort "cannot start Excel: " & Err.Number & " " & Err.Description
    If gXl Is Nothing Then Abort "cannot start Excel: CreateObject returned nothing"

    On Error Resume Next
    ' A test run builds the picker UserForms, and MSForms needs a visible host;
    ' the sheets stay behind the harness window, so nothing flashes. Interactive
    ' stays False: a stray dialog is an error in the log instead of a hang.
    If gMode = "test" Or gMode = "draft" Then
        gXl.Visible = True
    Else
        gXl.Visible = False
    End If
    gXl.Interactive = False
    gXl.DisplayAlerts = False
    gXl.ScreenUpdating = False
    gXl.EnableEvents = False
    gXl.AskToUpdateLinks = False
    gXl.AutomationSecurity = XL_AUTOMATION_SECURITY_LOW
    On Error GoTo 0

    gExcelVersion = gXl.Version
    Log "Excel started: version " & gExcelVersion
    StartExcel = True
End Function

Function OpenTarget()
    Set gWb = Nothing
    Err.Clear
    Set gWb = gXl.Workbooks.Open(gStage, XL_UPDATE_LINKS_NEVER, False)
    If Err.Number <> 0 Then Abort "cannot open " & gStage & ": " & Err.Number & " " & Err.Description
    If gWb Is Nothing Then Abort "cannot open " & gStage & ": no workbook returned"

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
    OpenTarget = True
End Function

Function RemoveCodeComponents()
    Dim components, index, component
    Set components = gWb.VBProject.VBComponents

    For index = components.Count To 1 Step -1
        Set component = components.Item(index)
        If component.Type <> VBE_CT_DOCUMENT Then
            Err.Clear
            components.Remove component
            If Err.Number <> 0 Then Abort "cannot remove component " & component.Name & ": " & Err.Description
            Log "removed existing component " & component.Name
        End If
    Next

    RemoveCodeComponents = True
End Function

Function ImportModules()
    ImportModules = ImportFolder(gLibDir, False)
    If Not ImportModules Then Exit Function
    ImportModules = ImportFolder(gSrcDir, True)
End Function

Function ImportFolder(ByVal folderPath, ByVal recurse)
    ImportFolder = True
    If Not gFso.FolderExists(folderPath) Then
        Log "source folder skipped (missing): " & folderPath
        Exit Function
    End If

    ' Test modules belong to the test build only; the production build must
    ' not contain them (see TODO.md, production release phase).
    If gMode <> "test" And gMode <> "draft" And LCase(gFso.GetFileName(folderPath)) = "tests" Then
        Log "test folder skipped in " & gMode & " mode: " & folderPath
        Exit Function
    End If

    Dim file, extension, subFolder
    For Each file In gFso.GetFolder(folderPath).Files
        extension = LCase(gFso.GetExtensionName(file.Path))
        If extension = "cls" Or extension = "bas" Then
            If Not ImportSourceFile(file.Path) Then
                ImportFolder = False
                Exit Function
            End If
        ElseIf extension = "frm" Then
            If Not ImportFormFile(file.Path) Then
                ImportFolder = False
                Exit Function
            End If
        End If
    Next

    If recurse Then
        For Each subFolder In gFso.GetFolder(folderPath).SubFolders
            If Not ImportFolder(subFolder.Path, True) Then
                ImportFolder = False
                Exit Function
            End If
        Next
    End If
End Function

Function ImportSourceFile(ByVal filePath)
    If Not VerifySourceFile(filePath) Then
        ImportSourceFile = False
        Exit Function
    End If

    Dim component
    Set component = Nothing
    Err.Clear
    Set component = gWb.VBProject.VBComponents.Import(filePath)
    If Err.Number <> 0 Then Abort "import failed for " & filePath & ": " & Err.Number & " " & Err.Description
    If component Is Nothing Then Abort "import returned no component for " & filePath

    gImportedCount = gImportedCount + 1
    Log "imported " & gFso.GetFileName(filePath) & " as " & component.Name & " (lines: " & component.CodeModule.CountOfLines & ")"
    ImportSourceFile = True
End Function

' A UserForm is written as code only: an exported .frm carries a binary .frx
' designer blob and a hand written .frm without it cannot be imported ("System
' Error &H80004005"). The build therefore creates the form component and
' injects the lines of the file into its code module; the controls are built at
' run time by the form itself (frmDatePicker.UserForm_Initialize), so nothing
' is lost by the empty designer.
Function ImportFormFile(ByVal filePath)
    ImportFormFile = False

    If Not VerifySourceFile(filePath) Then Exit Function

    Dim formName, component, module, text, lines, index, line, code
    formName = gFso.GetBaseName(filePath)

    Set component = Nothing
    Err.Clear
    Set component = gWb.VBProject.VBComponents.Add(VBE_CT_FORM)
    If Err.Number <> 0 Then Abort "cannot create form component " & formName & ": " & Err.Number & " " & Err.Description
    If component Is Nothing Then Abort "cannot create form component " & formName

    component.Name = formName
    Set module = component.CodeModule

    text = ReadTextFile(filePath)
    lines = Split(text, vbCrLf)
    code = ""

    ' Attribute lines belong to the exported form of the VBE, not to the code.
    For index = 0 To UBound(lines)
        line = lines(index)
        If InStr(1, line, "Attribute ", vbTextCompare) = 1 Then
            line = ""
        End If
        code = code & line & vbCrLf
    Next

    module.AddFromString code
    If Err.Number <> 0 Then Abort "cannot add the code of " & filePath & ": " & Err.Number & " " & Err.Description

    gImportedCount = gImportedCount + 1
    Log "created form " & component.Name & " from " & gFso.GetFileName(filePath) & _
        " (code lines: " & CStr(module.CountOfLines) & ")"
    ImportFormFile = True
End Function

Function ReadTextFile(ByVal filePath)
    Dim stream

    ReadTextFile = ""
    Set stream = Nothing

    Err.Clear
    Set stream = gFso.OpenTextFile(filePath, 1, False, -2)   ' ForReading, TristateUseDefault
    If Err.Number <> 0 Then Abort "cannot read " & filePath & ": " & Err.Description

    ReadTextFile = stream.ReadAll
    stream.Close
End Function

Function VerifySourceFile(ByVal filePath)
    VerifySourceFile = False

    Dim stream, text, remainder
    Set stream = Nothing
    Err.Clear
    Set stream = gFso.OpenTextFile(filePath, 1, False, -2)   ' ForReading, TristateUseDefault
    If Err.Number <> 0 Then Abort "cannot read " & filePath & ": " & Err.Description
    text = stream.ReadAll
    stream.Close

    If Len(text) = 0 Then Abort "source file is empty: " & filePath

    ' A UTF-8 byte order mark costs a whole cycle: VBA hides the Attribute line behind
    ' it and imports the file as "Module1". The check reads the file as ISO-8859-1,
    ' where every byte is one character, so the three BOM bytes are the same three
    ' characters whatever the system code page happens to be.
    If Left(RawTextOfFile(filePath), 3) = ChrW(&HEF) & ChrW(&HBB) & ChrW(&HBF) Then
        Abort "source file starts with a UTF-8 byte order mark (VBA would import it " & _
              "as Module1): " & filePath
    End If

    ' VBA takes the module name of a .bas/.cls from the Attribute line, not from the
    ' file name, so a mismatch quietly renames the component (modRoutes.bas once arrived
    ' as Module1). Forms are exempt: a .frm in the sources carries no Attribute line,
    ' ImportFormFile writes the designer blob (and with it the name) at import time.
    Dim extension, declaredName, baseName, position
    extension = LCase(gFso.GetExtensionName(filePath))
    If extension = "bas" Or extension = "cls" Then
        baseName = gFso.GetBaseName(filePath)
        position = InStr(1, text, "Attribute VB_Name = """, vbBinaryCompare)
        If position = 0 Then
            Abort "source file has no 'Attribute VB_Name' line (VBA would import it as " & _
                  "Module1): " & filePath
        End If
        position = position + Len("Attribute VB_Name = """)
        declaredName = Mid(text, position, InStr(position, text, """") - position)
        If StrComp(declaredName, baseName, vbTextCompare) <> 0 Then
            Abort "source file declares the module '" & declaredName & "' but is named '" & _
                  baseName & "': the import would carry the wrong name: " & filePath
        End If
    End If

    ' A stray "End Sub" behind an "End Function" (or the other way round) is a compile
    ' error in VBA ("Only comments may appear after End Sub...") that shows only when
    ' the module is first executed - so the harness must name it before Excel does.
    ' Blank lines and comments between the two closers are ignored, as VBA ignores them.
    Dim sourceLines, lineIndex, endLineSeen, trimmed
    sourceLines = Split(text, vbCrLf)
    endLineSeen = False
    For lineIndex = 0 To UBound(sourceLines)
        trimmed = Trim(sourceLines(lineIndex))
        If Len(trimmed) > 0 And Left(trimmed, 1) <> "'" Then
            If IsClosingEndLine(trimmed) Then
                If endLineSeen Then
                    Abort "source file has a stray '" & trimmed & "' behind the end of " & _
                          "a procedure (VBA would refuse the module): " & filePath & _
                          " (line " & CStr(lineIndex + 1) & ")"
                End If
                endLineSeen = True
            Else
                endLineSeen = False
            End If
        End If
    Next

    remainder = Replace(text, vbCrLf, "")
    If InStr(1, remainder, vbLf, vbBinaryCompare) > 0 Then
        Abort "source file uses LF line endings; VBA import needs CRLF: " & filePath
    End If
    If InStr(1, remainder, vbCr, vbBinaryCompare) > 0 Then
        Abort "source file uses CR line endings; VBA import needs CRLF: " & filePath
    End If

    VerifySourceFile = True
End Function

' The whole file read as ISO-8859-1: every byte is exactly one character, so the
' first characters of the result are the first bytes of the file (used by the byte
' order mark check of VerifySourceFile). Without ADO the check is skipped with a note.
Function RawTextOfFile(ByVal filePath)
    RawTextOfFile = ""

    Dim stream
    Set stream = Nothing
    Err.Clear
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        Log "the byte order mark of " & filePath & " was not checked (no ADODB.Stream)"
        Exit Function
    End If

    stream.Type = 2                 ' adTypeText
    stream.Charset = "iso-8859-1"
    stream.Open
    stream.LoadFromFile filePath
    RawTextOfFile = stream.ReadText(-1)   ' adReadAll
    stream.Close
End Function

' True for the lines that close a procedure: "End Sub", "End Function" and
' "End Property" (used by the structural check of VerifySourceFile).
Function IsClosingEndLine(ByVal lineText)
    Dim pattern

    Set pattern = New RegExp
    pattern.Pattern = "^End (Sub|Function|Property)(\s|$)"
    IsClosingEndLine = pattern.Test(lineText)
End Function

Function SaveAsMacroWorkbook()
    Err.Clear
    gWb.SaveAs gTarget, XL_MACRO_WORKBOOK
    If Err.Number <> 0 Then Abort "SaveAs failed for " & gTarget & ": " & Err.Number & " " & Err.Description
    Log "saved as macro workbook (" & CStr(gImportedCount) & " modules imported)"

    DeleteIfExists gStage
    SaveAsMacroWorkbook = True
End Function

' The build session is not the one the tests run in: the instance that imported
' the code (and used the VBE) died on the first Application.Run here, while a
' fresh instance runs the same file without any problem. So the built workbook is
' closed, Excel is started again and the file is opened fresh - the batch then
' runs in exactly the session a user gets (Workbook_Open runs modApp.Boot).
Function RestartExcelForTests()
    RestartExcelForTests = False

    Err.Clear
    On Error Resume Next
    gWb.Close False
    Set gWb = Nothing
    On Error GoTo 0
    If Err.Number <> 0 Then Abort "cannot close the built workbook: " & CStr(Err.Number) & " " & Err.Description

    CleanupExcel
    WScript.Sleep 1000
    If Not StartExcel() Then Exit Function
    If Not OpenBuiltTarget() Then Exit Function

    Log "Excel restarted for the test run"
    RestartExcelForTests = True
End Function

Function OpenBuiltTarget()
    OpenBuiltTarget = False

    Err.Clear
    On Error Resume Next
    Set gWb = gXl.Workbooks.Open(gTarget, XL_UPDATE_LINKS_NEVER, False)
    On Error GoTo 0
    If Err.Number <> 0 Or gWb Is Nothing Then
        Abort "cannot open " & gTarget & ": " & CStr(Err.Number) & " " & Err.Description
        Exit Function
    End If

    Log "opened " & gWb.Name & " for the tests"
    OpenBuiltTarget = True
End Function

Function RunTests()
    Dim macroName, result
    macroName = "'" & gWb.Name & "'!" & TEST_MODULE & "." & TEST_ENTRY

    result = ""
    Err.Clear
    On Error Resume Next

    result = gXl.Run(macroName)
    If Err.Number <> 0 Then
        ' The state of the VBE names the module of a compile error; it goes into
        ' its own file, because the script stops here and a buffered log would
        ' lose it.
        WriteNote "run failed: " & CStr(Err.Number) & " " & Err.Description
        WriteNote "workbook  : " & gWb.Name
        WriteNote "excel     : " & ExcelProcessList()
        WriteNote "vbe state : " & RunStateOf(gXl)
        On Error GoTo 0
        Abort "running " & macroName & " failed: " & Err.Number & " " & Err.Description & WatchdogHint()
    End If
    On Error GoTo 0

    If IsEmpty(result) Or IsNull(result) Then result = ""

    Log "tests executed: " & macroName
    RunTests = CStr(result)
End Function

' One line into build\run-error.txt, flushed at once (a failing run quits the
' script before the buffered log reaches the disk).
Sub WriteNote(text)
    Dim noteFile

    On Error Resume Next
    Set noteFile = gFso.OpenTextFile(gFso.BuildPath(gOutDir, "run-error.txt"), 8, True)
    If Not noteFile Is Nothing Then
        noteFile.WriteLine text
        noteFile.Close
    End If
    On Error GoTo 0
End Sub

' The EXCEL.EXE processes with their start time, so a dying instance can be told
' from one that was replaced by another.
Function ExcelProcessList()
    Dim services, processes, process, text

    text = ""
    On Error Resume Next
    Set services = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = services.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
    For Each process In processes
        text = text & CStr(process.ProcessId) & "@" & CStr(process.CreationDate) & " "
    Next
    On Error GoTo 0

    If Len(text) = 0 Then text = "(none)"
    ExcelProcessList = text
End Function

' What the VBE is showing while a Run failed: the module and the line the error
' dialog points at, so a compile error of an imported project can be named.
Function RunStateOf(xl)
    Dim pane, stateText, startLine, startColumn, endLine, endColumn

    stateText = ""
    startLine = 0
    startColumn = 0
    endLine = 0
    endColumn = 0

    On Error Resume Next
    stateText = "project=" & CStr(xl.VBE.ActiveVBProject.Name)
    Set pane = xl.VBE.ActiveCodePane
    If Err.Number = 0 Then
        If Not pane Is Nothing Then
            pane.GetSelection startLine, startColumn, endLine, endColumn
            stateText = stateText & ", module=" & CStr(pane.CodeModule.Name) & _
                        ", line=" & CStr(startLine) & "-" & CStr(endLine)
        End If
    End If
    On Error GoTo 0

    RunStateOf = stateText
End Function

' Runs modServiceManager.WriteModelReport inside the workbook. The function
' writes build\inspect-report.txt itself, because Application.Run truncates
' a returned string after 255 characters.
Function RunInspect()
    Dim macroName
    macroName = "'" & gWb.Name & "'!" & INSPECT_MODULE & "." & INSPECT_ENTRY

    DeleteIfExists gInspectPath

    Err.Clear
    gXl.Run macroName, gInspectPath
    If Err.Number <> 0 Then
        Abort "running " & macroName & " failed: " & Err.Number & " " & Err.Description & WatchdogHint()
    End If

    If Not gFso.FileExists(gInspectPath) Then
        Abort "the inspect run did not write " & gInspectPath
    End If

    Log "inspect report written: " & gInspectPath
    EchoFile gInspectPath
    RunInspect = True
End Function

Sub EchoFile(ByVal filePath)
    Dim stream, line
    On Error Resume Next
    Set stream = gFso.OpenTextFile(filePath, 1, False, -2)   ' ForReading, TristateUseDefault
    If Err.Number <> 0 Then Exit Sub

    Do While Not stream.AtEndOfStream
        line = stream.ReadLine
        WScript.Echo "  " & line
    Loop
    stream.Close
    On Error GoTo 0
End Sub

' =====================================================================
'  document modules
' =====================================================================
' Document modules (ThisWorkbook, the sheets) cannot be imported like .cls or
' .bas files. Their code is kept in src\doc\<ComponentName>.txt and injected
' into the matching component; the workbook template is macro free, so the
' injection replaces whatever the component holds.
Function InjectDocumentCode()
    InjectDocumentCode = True
    If Not gFso.FolderExists(gDocDir) Then
        Log "document code folder skipped (missing): " & gDocDir
        Exit Function
    End If

    Dim file, componentName, component, module, text
    For Each file In gFso.GetFolder(gDocDir).Files
        If LCase(gFso.GetExtensionName(file.Path)) = "txt" Then
            If Not VerifySourceFile(file.Path) Then
                InjectDocumentCode = False
                Exit Function
            End If

            componentName = gFso.GetBaseName(file.Path)
            Set component = Nothing
            Err.Clear
            On Error Resume Next
            Set component = gWb.VBProject.VBComponents(componentName)
            On Error GoTo 0

            If component Is Nothing Then
                Log "document code " & gFso.GetFileName(file.Path) & _
                    " skipped: no component named " & componentName
            Else
                text = ReadTextFile(file.Path)
                Set module = component.CodeModule
                If module.CountOfLines > 0 Then
                    module.DeleteLines 1, module.CountOfLines
                    Log "cleared the existing code of " & componentName
                End If
                Err.Clear
                module.AddFromString text
                If Err.Number <> 0 Then Abort "cannot inject " & file.Path & ": " & Err.Number & " " & Err.Description
                Log "injected " & gFso.GetFileName(file.Path) & " into " & componentName & _
                    " (code lines: " & CStr(module.CountOfLines) & ")"
            End If
        End If
    Next
End Function

' =====================================================================
'  release build of the pre-production version
' =====================================================================
' Publishes the workbook: the built (and therefore clean) file is started once
' (modApp.Boot is the same call Workbook_Open uses), saved and copied to
' release\. The README next to it is generated from this run, so the folder
' always describes the file that lies in it.
' The cell the published file shows: the first row of the dashboard body, taken by
' name, so a new layout may move it anywhere. The constant is the fallback.
Function ReleaseViewCell()
    Dim nameObject, addressText

    ' By name only (the constant above used to keep a copy of the address): a cell
    ' that moves with a redesign would otherwise break the release alone, which is
    ' the worst way to find out.
    ReleaseViewCell = ""
    addressText = ""

    Err.Clear
    On Error Resume Next
    Set nameObject = gWb.Names(RELEASE_VIEW_NAME)
    If Not nameObject Is Nothing Then addressText = nameObject.RefersToRange.Address(False, False)
    On Error GoTo 0

    If Len(CStr(addressText)) = 0 Then
        Abort "the workbook carries no name " & RELEASE_VIEW_NAME & _
              " - the release view needs it (the template writes it)"
    End If

    ReleaseViewCell = CStr(addressText)
End Function

Function PublishRelease()
    PublishRelease = False

    If Not VerifyVersion() Then Exit Function

    Err.Clear
    gXl.Run "'" & gWb.Name & "'!modApp.Boot"
    If Err.Number <> 0 Then
        Abort "StartServiceManager failed: " & Err.Number & " " & Err.Description & WatchdogHint()
    End If
    Log "the application started inside the built workbook"

    Err.Clear
    gWb.Save
    If Err.Number <> 0 Then Abort "cannot save " & gTarget & ": " & Err.Number & " " & Err.Description

    ' The published file must open on the Dashboard, also when the user does not
    ' enable macros: the active sheet is stored in the file. An Activate from
    ' inside VBA is not reliable while the build window is hidden, so the harness
    ' does it (and puts the cursor on the week cell).
    Err.Clear
    gWb.Worksheets(SHEET_DASHBOARD_NAME).Activate
    releaseView = ReleaseViewCell()
    gWb.Worksheets(SHEET_DASHBOARD_NAME).Range(releaseView).Select
    If Err.Number <> 0 Then
        Log "cannot set the view of the release file: " & Err.Description
        Err.Clear
    Else
        Log "release view: " & SHEET_DASHBOARD_NAME & "!" & releaseView
    End If
    Err.Clear
    gWb.Save
    If Err.Number <> 0 Then Abort "cannot save " & gTarget & " again: " & Err.Number & " " & Err.Description

    EnsureFolder gReleaseDir
    Dim releasePath, draftMode, releaseView
    draftMode = (gMode = "draft")
    If draftMode Then
        releasePath = gFso.BuildPath(gReleaseDir, "Service_Manager_draft_" & APP_VERSION & ".xlsm")
    Else
        ' The file that goes out is a release, not a pre-release: the name says so.
        releasePath = gFso.BuildPath(gReleaseDir, "Service_Manager_" & APP_VERSION & ".xlsm")
    End If
    DeleteIfExists releasePath
    Err.Clear
    gFso.CopyFile gTarget, releasePath, True
    If Err.Number <> 0 Then Abort "cannot write " & releasePath & ": " & Err.Description
    Log "release copy: " & releasePath

    ' Both files in both modes: a release is read by the owner as well, and the
    ' checklist is what "done" means (§15, Phase 7).
    WriteDraftChecklist releasePath
    WriteReleaseReadme releasePath
    PublishRelease = True
End Function

' The version of the harness and the version of the application must agree,
' otherwise the file name and the README would lie.
Function VerifyVersion()
    VerifyVersion = False

    Dim configPath, text
    configPath = gFso.BuildPath(gFso.BuildPath(gRoot, "src"), "app\modConfig.bas")
    If Not gFso.FileExists(configPath) Then Abort "modConfig.bas not found: " & configPath

    text = ReadTextFile(configPath)
    If InStr(1, text, "APP_VERSION As String = """ & APP_VERSION & """", vbBinaryCompare) = 0 Then
        Abort "the version of this harness (" & APP_VERSION & _
              ") does not match modConfig.APP_VERSION in " & configPath
    End If

    Log "version: " & APP_VERSION
    VerifyVersion = True
End Function

' The harness skips src\tests in release mode, so the file that is handed over has
' no test batch of its own: the gate of a release is the report of the lane that
' does have one. It counts only when it is green and newer than every file under
' src\ and lib\ - otherwise the release would carry untested code.
Function VerifyTestReportIsFresh()
    VerifyTestReportIsFresh = False

    Dim reportPath, text, reportTime, newestTime, newestName
    reportPath = gFso.BuildPath(gOutDir, "test-report.txt")

    If Not gFso.FileExists(reportPath) Then
        Abort "no test report (" & reportPath & _
              "): run ""cscript //nologo tools\build.vbs draft"" first"
    End If

    text = ReadTextFile(reportPath)
    If InStr(1, text, "result    : PASSED - ", vbBinaryCompare) = 0 Then
        Abort "the report in " & reportPath & _
              " is not green: run ""cscript //nologo tools\build.vbs draft"" first"
    End If

    reportTime = gFso.GetFile(reportPath).DateLastModified
    newestTime = reportTime
    newestName = ""
    NewestSourceTime gSrcDir, newestTime, newestName
    NewestSourceTime gLibDir, newestTime, newestName

    If Len(newestName) > 0 Then
        Abort "the report in " & reportPath & " is older than " & newestName & _
              ": run ""cscript //nologo tools\build.vbs draft"" first"
    End If

    Log "test report is fresh: " & gFso.GetFileName(reportPath)
    VerifyTestReportIsFresh = True
End Function

' The newest modification time of anything under folderPath; newestName receives
' the path of the file that holds it, so the message can name the culprit.
Sub NewestSourceTime(ByVal folderPath, newestTime, newestName)
    Dim folder, file, item

    If Not gFso.FolderExists(folderPath) Then Exit Sub

    Set folder = gFso.GetFolder(folderPath)
    For Each file In folder.Files
        If file.DateLastModified > newestTime Then
            newestTime = file.DateLastModified
            newestName = file.Path
        End If
    Next
    For Each item In folder.SubFolders
        NewestSourceTime item.Path, newestTime, newestName
    Next
End Sub

Sub EnsureFolder(ByVal folderPath)
    If gFso.FolderExists(folderPath) Then Exit Sub

    Err.Clear
    gFso.CreateFolder folderPath
    If Err.Number <> 0 Then Abort "cannot create " & folderPath & ": " & Err.Description
    Log "created " & folderPath
End Sub

' Removes a draft that was published while the tests still had to prove it.
Sub RemovePublishedDraft()
    Dim draftPath, publishedPath, checklistPath, releaseChecklistPath, readmePath

    draftPath = gFso.BuildPath(gReleaseDir, "Service_Manager_draft_" & APP_VERSION & ".xlsm")
    publishedPath = gFso.BuildPath(gReleaseDir, "Service_Manager_" & APP_VERSION & ".xlsm")
    checklistPath = gFso.BuildPath(gReleaseDir, "DRAFT_" & APP_VERSION & ".txt")
    releaseChecklistPath = gFso.BuildPath(gReleaseDir, "CHECKLIST_" & APP_VERSION & ".txt")
    readmePath = gFso.BuildPath(gReleaseDir, RELEASE_README)

    ' Everything the failed run published goes: a draft or a release exists only
    ' when the tests behind it are green.
    If gFso.FileExists(draftPath) Then DeleteIfExists draftPath
    If gFso.FileExists(publishedPath) Then DeleteIfExists publishedPath
    If gFso.FileExists(checklistPath) Then DeleteIfExists checklistPath
    If gFso.FileExists(releaseChecklistPath) Then DeleteIfExists releaseChecklistPath
    If gFso.FileExists(readmePath) Then DeleteIfExists readmePath
End Sub

' The draft of a phase: the workbook plus the checklist of what to try by hand.
' The checklist itself lives in the repository (DRAFT_CHECKLIST.md); the harness
' only adds the header with the file name, the version and the moment of the build.
Sub WriteDraftChecklist(ByVal draftPath)
    Dim sourcePath, targetPath, header, text, file, stream

    sourcePath = gFso.BuildPath(gRoot, DRAFT_CHECKLIST)
    ' The checklist follows the mode: a draft is tested by the owner day by day,
    ' a release is the file that goes to a clean machine (§15).
    If gMode = "draft" Then
        targetPath = gFso.BuildPath(gReleaseDir, "DRAFT_" & APP_VERSION & ".txt")
    Else
        targetPath = gFso.BuildPath(gReleaseDir, "CHECKLIST_" & APP_VERSION & ".txt")
    End If

    header = "Service_Manager - " & gMode & " for a manual test" & vbCrLf & _
             "=============================================" & vbCrLf & vbCrLf & _
             "workbook : " & draftPath & vbCrLf & _
             "version  : " & APP_VERSION & vbCrLf & _
             "built    : " & FormatDateTime(Now, vbGeneralDate) & vbCrLf & _
             "tests    : build\test-report.txt (the draft is published only when they pass)" & _
             vbCrLf & vbCrLf

    text = vbNullString
    If gFso.FileExists(sourcePath) Then
        ' The checklist is written in UTF-8; the harness reads it with the right
        ' charset and writes the copy as Unicode, so the Cyrillic text survives.
        Err.Clear
        Set stream = CreateObject("ADODB.Stream")
        If Err.Number = 0 Then
            stream.Type = 2                      ' adTypeText
            stream.Charset = "utf-8"
            stream.Open
            stream.LoadFromFile sourcePath
            text = stream.ReadText
            stream.Close
        End If
        If Len(text) = 0 Then text = ReadTextFile(sourcePath)
    Else
        text = "(the checklist file " & DRAFT_CHECKLIST & " is missing)"
    End If

    Set file = Nothing
    Err.Clear
    Set file = gFso.CreateTextFile(targetPath, True, True)   ' Unicode
    If Err.Number <> 0 Then
        Log "cannot write " & targetPath & ": " & Err.Description
        Exit Sub
    End If

    file.Write header & text
    file.Close
    Log "checklist: " & targetPath
End Sub

Sub WriteReleaseReadme(ByVal releasePath)
    Dim readme, readmePath, file

    readmePath = gFso.BuildPath(gReleaseDir, RELEASE_README)
    Set file = Nothing
    Err.Clear
    Set file = gFso.CreateTextFile(readmePath, True, False)
    If Err.Number <> 0 Then
        Log "cannot write " & readmePath & ": " & Err.Description
        Exit Sub
    End If

    readme = gFso.GetFileName(releasePath)
    file.WriteLine "Service_Manager - build " & APP_VERSION & " (" & gMode & ")"
    file.WriteLine "built     : " & CStr(Now)
    file.WriteLine "file      : " & readme
    file.WriteLine "modules   : " & CStr(gImportedCount) & " VBA modules (library + application, no tests)"
    file.WriteLine "source    : " & gRoot
    file.WriteLine "tests     : every lane starts with the lint, the script model and the cross"
    file.WriteLine "            check; a release is cut only when the report of the test lane"
    file.WriteLine "            (build\test-report.txt) is green and newer than the sources"
    file.WriteLine ""
    file.WriteLine "What it is"
    file.WriteLine "----------"
    file.WriteLine "A hairdresser's diary in one workbook: appointments with client, service, date,"
    file.WriteLine "start and duration; three views of the plan; reports for a period; and a Config"
    file.WriteLine "page for the working hours. The data lives in the same file (the db_* sheets)."
    file.WriteLine ""
    file.WriteLine "How to start"
    file.WriteLine "------------"
    file.WriteLine "1. Keep the file in a folder you may write to: the application creates a"
    file.WriteLine "   ""backups"" sub folder next to it."
    file.WriteLine "2. If Windows marked the file as downloaded, unblock it first:"
    file.WriteLine "   right click the file -> Properties -> Unblock."
    file.WriteLine "3. Open the workbook and press ""Enable Content"" (""Enable Macros""). It opens on"
    file.WriteLine "   the Dashboard. If macros are blocked, run it by hand:"
    file.WriteLine "   Alt+F8 -> modApp.Boot -> Run."
    file.WriteLine ""
    file.WriteLine "What to try"
    file.WriteLine "-----------"
    file.WriteLine "- Dashboard: Daily / Weekly / Monthly with the cells at the top, Prev / Today /"
    file.WriteLine "  Next for the period. A free slot books a new appointment; the first row of an"
    file.WriteLine "  appointment opens it for editing (the rows below are the same hour going on)."
    file.WriteLine "- Booking: the Pick buttons open the calendar and the time fields, a service"
    file.WriteLine "  fills the duration, and Cancel becomes Delete while an appointment is edited."
    file.WriteLine "- ClientsAdmin / ServicesAdmin: a form at the top, the whole table below."
    file.WriteLine "- Reports: From / To with the Pick buttons, Show renders the period by day,"
    file.WriteLine "  service and client, with the total."
    file.WriteLine "- Config: opening and closing time, slot step, default duration, currency and"
    file.WriteLine "  number format, non working weekdays; Pick a day adds a non working date,"
    file.WriteLine "  Clear dates drops the list, Apply settings stores the page and redraws the"
    file.WriteLine "  grids. Lock data sheets hides the three db_* sheets (Unlock shows them again,"
    file.WriteLine "  or right click a sheet tab -> Unhide)."
    file.WriteLine ""
    file.WriteLine "Conventions of the file"
    file.WriteLine "-----------------------"
    file.WriteLine "- db_clients, db_services and db_appointments hold the data: header in row 1,"
    file.WriteLine "  dates as YYYY-MM-DD and times as HH:MM - text, so Excel does not rewrite them"
    file.WriteLine "  with the locale of the machine."
    file.WriteLine "- Every action lands in the Log sheet; every delete writes a copy of the"
    file.WriteLine "  workbook into ""backups"" (the newest 10 are kept)."
    file.WriteLine "- s_schema_version in Config says which layout the file has: an older file is"
    file.WriteLine "  migrated when it is opened, a newer one is refused with a sentence."
    file.WriteLine ""
    file.WriteLine "How this file was made"
    file.WriteLine "----------------------"
    file.WriteLine "cscript //nologo tools\build.vbs release"
    file.WriteLine "The harness runs the lint of the sources, the script model (tools\sim, no"
    file.WriteLine "Excel), the cross check of the two implementations and the whole Excel test"
    file.WriteLine "batch; only then does it publish, with the workbook opened on the Dashboard."
    file.WriteLine ""
    file.WriteLine "Not in this build"
    file.WriteLine "-----------------"
    file.WriteLine "- No ribbon, no menu bar and no signed macro project: the navigation cells are"
    file.WriteLine "  the menu, and Excel asks about macros because the project is not signed."
    file.WriteLine "- Reports are not exported to PDF automatically; print the sheet by hand."
    file.WriteLine "- The VBA project is not password protected: the harness needs it open to"
    file.WriteLine "  import the sources for the next build."
    file.WriteLine "- The checklist of what to try by hand sits next to this file (DRAFT_ or"
    file.WriteLine "  CHECKLIST_<version>.txt)."
    file.Close

    Log "release readme: " & readmePath
End Sub

' Extra hint for the two failure modes the watchdog protects against.
Function WatchdogHint()
    If gFso.FileExists(gTimeoutPath) Then
        WatchdogHint = " Excel was terminated by the watchdog; the run blocked on a modal VBA dialog" & _
                       " (compile error or unhandled run-time error). Open " & gTarget & " in Excel and check the VBE."
    Else
        WatchdogHint = ""
    End If
End Function

' =====================================================================
'  watchdog
' =====================================================================
Sub StartWatchdog()
    Set gShell = CreateObject("WScript.Shell")

    DeleteIfExists gTimeoutPath
    ' The flag file names the run: the watchdog stops as soon as the file belongs
    ' to another run, so the watchdog of an abandoned run can never kill the Excel
    ' instance of a later one (that is how a fresh run used to die after exactly
    ' the watchdog time of the previous one).
    gRunId = CStr(Timer) & "-" & CStr(WScript.ScriptName)
    WriteTextFile gSentinelPath, "run " & gRunId & vbCrLf & "started " & CStr(Now) & vbCrLf

    gShell.Run "cscript //nologo """ & WScript.ScriptFullName & """ watchdog """ & _
               gSentinelPath & """ """ & gRunId & """", 0, False
    Log "watchdog started (Excel is terminated after " & CStr(WATCHDOG_SECONDS) & " s without progress)"
End Sub

Sub StopWatchdog()
    If Not gFso.FileExists(gSentinelPath) Then Exit Sub

    On Error Resume Next
    gFso.DeleteFile gSentinelPath, True
    On Error GoTo 0
End Sub

Sub RunWatchdog()
    Dim sentinel, runId, elapsed, flagText
    If WScript.Arguments.Count < 3 Then Exit Sub
    sentinel = WScript.Arguments(1)
    runId = WScript.Arguments(2)

    elapsed = 0
    Do While elapsed < WATCHDOG_SECONDS
        WScript.Sleep 2000
        elapsed = elapsed + 2
        If Not gFso.FileExists(sentinel) Then Exit Sub   ' the harness finished
        flagText = ReadTextFile(sentinel)
        If InStr(1, flagText, "run " & runId, vbBinaryCompare) = 0 Then Exit Sub   ' another run owns it
    Loop

    ' The harness is still running, so Excel is blocked by a modal dialog.
    WriteTextFile Replace(sentinel, "_harness.flag", "_timeout.txt"), _
                  "Excel did not respond within " & CStr(WATCHDOG_SECONDS) & " seconds and was terminated." & vbCrLf & _
                  "The usual cause is a modal VBA dialog: a compile error or an unhandled run-time error." & vbCrLf & _
                  "Open the generated .xlsm in Excel, press Alt+F11 and press F8/F5 to see the offending line." & vbCrLf

    KillExcelProcesses
    On Error Resume Next
    If gFso.FileExists(sentinel) Then gFso.DeleteFile sentinel, True
    On Error GoTo 0
End Sub

Sub KillExcelProcesses()
    On Error Resume Next

    Dim services, processes, process
    Set services = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = services.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
    For Each process In processes
        process.Terminate
    Next

    On Error GoTo 0
End Sub

Sub WriteTextFile(ByVal filePath, ByVal text)
    Dim file
    Set file = Nothing
    Err.Clear
    Set file = gFso.CreateTextFile(filePath, True, True)
    If Err.Number <> 0 Then Exit Sub
    file.Write text
    file.Close
End Sub

' =====================================================================
'  infrastructure
' =====================================================================
Sub Log(ByVal text)
    gLog.Add gLog.Count + 1, text
    WScript.Echo text
End Sub

Sub WriteReport(ByVal testResult)
    If Not gFso.FolderExists(gOutDir) Then
        Err.Clear
        gFso.CreateFolder gOutDir
        If Err.Number <> 0 Then Exit Sub
    End If

    Dim report, line, index
    Set report = Nothing
    Err.Clear
    Set report = gFso.CreateTextFile(gReportPath, True, True)   ' overwrite, Unicode
    If Err.Number <> 0 Then
        WScript.Echo "cannot write " & gReportPath & ": " & Err.Description
        Exit Sub
    End If

    report.WriteLine "Service_Manager - build and test report"
    report.WriteLine "generated : " & CStr(Now)
    report.WriteLine "mode      : " & gMode
    report.WriteLine "target    : " & gTarget
    ' gXl is already released when the report is written from Abort, so the
    ' version string is captured once when Excel starts.
    report.WriteLine "excel     : " & gExcelVersion
    report.WriteLine "modules   : " & CStr(gImportedCount)
    report.WriteLine ""
    report.WriteLine "result    : " & testResult
    report.WriteLine ""
    report.WriteLine "log"

    For index = 1 To gLog.Count
        line = gLog(index)
        report.WriteLine "  " & line
    Next
    report.Close

    WScript.Echo "report written: " & gReportPath
End Sub

Sub Abort(ByVal text)
    WScript.Echo "ABORTED: " & text
    StopWatchdog
    CleanupExcel
    WriteReport "ABORTED - " & text
    WScript.Quit EXIT_ABORTED
End Sub

' =====================================================================
'  the fast lane: the script model of tools\sim
' =====================================================================

' The lint of the sources (tools\lint.vbs). VBA refuses a module as a whole, so one
' bad name anywhere stops every macro of the workbook and the harness reports
' "ABORTED" without naming the line (the trace even keeps the lines of the previous
' run). The lint reads the sources and names it, in a fraction of a second, before
' Excel is opened at all.
Function RunLint()
    Dim lintScript, lintReport, lintExit

    RunLint = True
    lintScript = gFso.BuildPath(gFso.BuildPath(gRoot, "tools"), "lint.vbs")
    lintReport = gFso.BuildPath(gOutDir, "lint-report.txt")
    If Not gFso.FileExists(lintScript) Then
        Log "the lint is not here: " & lintScript
        Exit Function
    End If

    lintExit = RunSimJob(lintScript, "")
    If lintExit <> 0 Then
        EchoFile lintReport
        Abort "the lint of the sources failed (exit code " & CStr(lintExit) & _
              "), see build\lint-report.txt"
        RunLint = False
        Exit Function
    End If

    Log "lint: no problems in the sources"
End Function

' The tests of the model (~0.1 s, no Excel) and its answers to the cases of
' src\tests\crosscheck.txt. False stops the build: a rule that is red in the model
' is red in the workbook as well, and finding it here costs a second instead of a
' minute and a half.
Function RunFastLane()
    Dim simScript
    Dim simExit

    RunFastLane = False
    If Not RunLint() Then Exit Function

    simScript = gFso.BuildPath(gFso.BuildPath(gRoot, "tools\sim"), "run.wsf")
    If Not gFso.FileExists(simScript) Then
        Log "the script model is not here: " & simScript
        RunFastLane = True
        Exit Function
    End If

    simExit = RunSimJob(simScript, "")
    If simExit <> 0 Then
        Abort "the script model of tools\sim failed (exit code " & CStr(simExit) & _
              "), see build\sim-test-report.txt"
        Exit Function
    End If
    Log "script model: green, see build\sim-test-report.txt"

    simExit = RunSimJob(simScript, "crosscheck")
    If simExit <> 0 Then
        Abort "the cross check of the script model failed (exit code " & CStr(simExit) & ")"
        Exit Function
    End If

    RunFastLane = True
End Function

' One job of the fast lane (cscript with a hidden window, waited for): its exit
' code, or -1 when it could not run at all.
Function RunSimJob(ByVal scriptPath, ByVal jobName)
    Dim command, exitCode, shell

    command = "cscript //nologo """ & scriptPath & """"
    If Len(jobName) > 0 Then command = command & " " & jobName

    exitCode = -1
    Err.Clear
    On Error Resume Next
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run(command, 0, True)
    If Err.Number <> 0 Then
        Log "the script model did not run: " & CStr(Err.Number) & " " & Err.Description
        exitCode = -1
    End If
    On Error GoTo 0

    RunSimJob = exitCode
End Function

' The two lanes answered the same cases of the same file (src\tests\
' crosscheck.txt); the answers are compared line by line - comments and empty
' lines go. "" when they agree, otherwise the first difference as a text and the
' build fails.
Function CompareCrosscheck()
    Dim simPath, xlsPath
    Dim simLines, xlsLines
    Dim simCount, xlsCount
    Dim index

    CompareCrosscheck = ""
    simPath = gFso.BuildPath(gOutDir, "sim-crosscheck.txt")
    xlsPath = gFso.BuildPath(gOutDir, "xls-crosscheck.txt")

    If Not gFso.FileExists(simPath) Then
        CompareCrosscheck = "the answers of the script model are missing: " & simPath
        Exit Function
    End If
    If Not gFso.FileExists(xlsPath) Then
        CompareCrosscheck = "the answers of the workbook are missing: " & xlsPath
        Exit Function
    End If

    Set simLines = ReadAnswerLines(simPath)
    Set xlsLines = ReadAnswerLines(xlsPath)
    simCount = simLines.Count
    xlsCount = xlsLines.Count

    For index = 0 To simCount - 1
        If index >= xlsCount Then
            CompareCrosscheck = "the workbook answered " & CStr(xlsCount) & _
                                " lines, the model " & CStr(simCount)
            Exit Function
        End If
        If StrComp(simLines(index), xlsLines(index), vbBinaryCompare) <> 0 Then
            CompareCrosscheck = "line " & CStr(index + 1) & ": the model says '" & _
                                simLines(index) & "', the workbook says '" & _
                                xlsLines(index) & "'"
            Exit Function
        End If
    Next

    If xlsCount > simCount Then
        CompareCrosscheck = "the workbook answered " & CStr(xlsCount) & _
                            " lines, the model " & CStr(simCount)
        Exit Function
    End If

    Log "cross check: " & CStr(simCount) & " lines, both lanes agree"
End Function

' The answer lines of one cross check file, in the order they were written.
Function ReadAnswerLines(ByVal filePath)
    Dim file, line, list

    Set list = CreateObject("Scripting.Dictionary")
    Set file = gFso.OpenTextFile(filePath, 1, False)
    Do Until file.AtEndOfStream
        line = Trim(CStr(file.ReadLine))
        If Len(line) > 0 Then
            If Left(line, 1) <> "#" Then list.Add list.Count, line
        End If
    Loop
    file.Close

    Set ReadAnswerLines = list
End Function

Sub CleanupExcel()
    On Error Resume Next

    If Not gWb Is Nothing Then
        gWb.Close False          ' never save from the harness; the build was already saved
        Set gWb = Nothing
    End If
    If Not gXl Is Nothing Then
        gXl.Quit
        Set gXl = Nothing
    End If

    On Error GoTo 0
End Sub
