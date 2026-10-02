' =====================================================================
'  Service_Manager - the lint of the VBA sources
'
'      cscript //nologo tools\lint.vbs [sourceFolder] [reportPath]
'
'  VBA compiles a module as a whole: one line it refuses stops every macro of the
'  workbook, and the harness then reports "ABORTED" without naming the line (the
'  trace file even keeps the lines of the previous run, which sends the search in
'  the wrong direction). This script reads the sources and names those lines
'  before Excel is opened, in a fraction of a second:
'
'      RESERVED  a reserved word used as a name      Dim empty As Object
'      SHADOW    a local that hides a procedure      Dim problems As String
'                of the same module                  next to Function Problems
'      FOREACH   For Each with a scalar control      Dim key As String
'                                                     For Each key In ...
'      DUPLICATE the same name twice in one scope    Dim value As Long  (twice)
'
'  Defaults: the folder src and the report build\lint-report.txt. The exit code
'  is 0 only when nothing was found; tools\build.vbs runs it in the fast lane.
' =====================================================================

Option Explicit

Dim gFso, gRoot, gSourceDir, gReportPath
Dim gProblems, gFileCount
Dim gReserved, gScalars, gStatements

Set gFso = CreateObject("Scripting.FileSystemObject")
Set gProblems = CreateObject("Scripting.Dictionary")
Set gReserved = CreateObject("Scripting.Dictionary")
Set gScalars = CreateObject("Scripting.Dictionary")
Set gStatements = CreateObject("Scripting.Dictionary")

Main

' =====================================================================
'  the run
' =====================================================================

Sub Main()
    gRoot = gFso.GetParentFolderName(gFso.GetParentFolderName(WScript.ScriptFullName))
    gSourceDir = gFso.BuildPath(gRoot, "src")
    gReportPath = gFso.BuildPath(gFso.BuildPath(gRoot, "build"), "lint-report.txt")

    If WScript.Arguments.Count > 0 Then gSourceDir = ResolvePath(WScript.Arguments(0))
    If WScript.Arguments.Count > 1 Then gReportPath = ResolvePath(WScript.Arguments(1))

    FillReserved
    FillScalars
    FillStatements

    If Not gFso.FolderExists(gSourceDir) Then
        WScript.Echo "lint: folder not found: " & gSourceDir
        WScript.Quit 2
    End If

    gFileCount = 0
    ScanFolder gSourceDir, True
    WriteReport
End Sub

' A path from the command line: a full path as it is, a relative one against the
' root of the project (the harness may run the script from anywhere).
Function ResolvePath(ByVal argument)
    If Mid(argument, 2, 1) = ":" Then
        ResolvePath = argument
    ElseIf Left(argument, 2) = "\\" Then
        ResolvePath = argument
    Else
        ResolvePath = gFso.BuildPath(gRoot, argument)
    End If
End Function

' =====================================================================
'  the tables of the lint
' =====================================================================

' The words VBA never accepts as a name. Words that only look reserved are left
' out on purpose (Date, Time, Name, Error, Step, Value, String, ... all compile as
' a variable name, and some of them are used that way in these sources).
Sub FillReserved()
    AddWords gReserved, "and,as,attribute,byref,byval,call,case,const,declare," & _
                        "dim,do,each,else,elseif,empty,end,enum,eqv,erase,exit," & _
                        "false,for,friend,function,get,global,goto,if,imp,in,is," & _
                        "let,like,loop,me,mod,new,next,nothing,not,null,on,option," & _
                        "optional,or,paramarray,preserve,private,property,public," & _
                        "redim,rem,resume,select,set,static,stop,sub,then,to,true," & _
                        "type,until,wend,while,with,xor"
End Sub

' The types a For Each control variable cannot have.
Sub FillScalars()
    AddWords gScalars, "boolean,byte,currency,date,decimal,double,integer,long," & _
                       "longlong,single,string"
End Sub

' The words a procedure may not be named after: VBA statements. "Sub Write(...)"
' and "Sub Lock(...)" do not compile (lesson 1 of TODO.md, and the hiding of the
' data sheets ran into Lock/Unlock), and the harness only says "Exception
' occurred" - so the names are checked here.
Sub FillStatements()
    AddWords gStatements, "write,print,open,close,lock,unlock,name,line,input," & _
                          "seek,put,kill,chdir,mkdir,rmdir,filecopy,beep,stop," & _
                          "load,save,get,set,let,width,height,scale,rem"
End Sub

Sub AddWords(ByRef table, ByVal wordsText)
    Dim parts, index

    parts = Split(wordsText, ",")
    For index = 0 To UBound(parts)
        table(Trim(parts(index))) = True
    Next
End Sub

' =====================================================================
'  walking the sources
' =====================================================================

Sub ScanFolder(ByVal folderPath, ByVal recurse)
    Dim folder, file, subFolder

    Set folder = gFso.GetFolder(folderPath)
    For Each file In folder.Files
        If IsSourceFile(file.Name) Then
            gFileCount = gFileCount + 1
            ScanFile file.Path
        End If
    Next
    If recurse Then
        For Each subFolder In folder.SubFolders
            ScanFolder subFolder.Path, True
        Next
    End If
End Sub

Function IsSourceFile(ByVal fileName)
    Dim extension

    extension = LCase(gFso.GetExtensionName(fileName))
    IsSourceFile = (extension = "bas" Or extension = "cls" Or extension = "frm")
End Function

' One report line per problem, in the order they were found.
Sub AddProblem(ByVal kind, ByVal filePath, ByVal lineNumber, ByVal message, ByVal codeLine)
    Dim text

    text = kind & "  " & ShortPath(filePath) & ":" & CStr(lineNumber) & "  " & message
    If Len(Trim(codeLine)) > 0 Then text = text & ": " & Trim(codeLine)
    gProblems.Add gProblems.Count, text
End Sub

Function ShortPath(ByVal filePath)
    If LCase(Left(filePath, Len(gRoot))) = LCase(gRoot) Then
        ShortPath = Mid(filePath, Len(gRoot) + 2)
    Else
        ShortPath = filePath
    End If
End Function

' =====================================================================
'  the report
' =====================================================================

Sub WriteReport()
    Dim report, index, result, text

    If gProblems.Count = 0 Then
        result = "OK - no problems"
    Else
        result = "FAILED - " & CStr(gProblems.Count) & " problem(s)"
    End If

    Set report = gFso.CreateTextFile(gReportPath, True, True)
    report.WriteLine "Service_Manager - lint of the VBA sources"
    report.WriteLine "generated : " & CStr(Now)
    report.WriteLine "folder    : " & gSourceDir
    report.WriteLine "files     : " & CStr(gFileCount) & " (bas, cls, frm)"
    report.WriteLine ""
    report.WriteLine "result    : " & result
    If gProblems.Count > 0 Then
        report.WriteLine ""
        report.WriteLine "problems"
        For index = 0 To gProblems.Count - 1
            text = gProblems(index)
            report.WriteLine "  " & text
            WScript.Echo "  " & text
        Next
    End If
    report.Close

    WScript.Echo "lint: " & result & " in " & CStr(gFileCount) & " files"
    WScript.Echo "report written: " & gReportPath

    If gProblems.Count = 0 Then
        WScript.Quit 0
    Else
        WScript.Quit 1
    End If
End Sub

' =====================================================================
'  one file
' =====================================================================

Sub ScanFile(ByVal filePath)
    Dim file, text

    Err.Clear
    On Error Resume Next
    Set file = gFso.OpenTextFile(filePath, 1, False)
    If Err.Number <> 0 Then
        On Error GoTo 0
        AddProblem "UNREADABLE", filePath, 0, "the file could not be read", ""
        Exit Sub
    End If
    text = file.ReadAll
    file.Close
    On Error GoTo 0

    If Len(text) > 0 Then ScanLines filePath, Split(text, vbCrLf)
End Sub

' The lines of one source. Three scopes matter: the procedure (a local lives from
' its Sub to its End Sub), the module and the names of the procedures themselves.
Sub ScanLines(ByVal filePath, ByRef lines)
    Dim procedures, moduleScope, localScope
    Dim moduleLines, localLines
    Dim index, code, name, typeName, key
    Dim names, params
    Dim inProcedure
    Dim scopeCode

    Set procedures = ProcedureNames(lines)
    Set moduleScope = NewScope()
    Set moduleLines = NewScope()
    Set localScope = NewScope()
    Set localLines = NewScope()
    inProcedure = False
    scopeCode = ""

    For index = 0 To UBound(lines)
        code = CodeOf(lines(index))

        If Len(Trim(code)) > 0 Then

            ' 1. the end of the procedure that is open
            If inProcedure And IsProcedureEnd(code) Then
                CheckShadows filePath, localScope, localLines, scopeCode, procedures, False, lines
                inProcedure = False
                Set localScope = NewScope()
                Set localLines = NewScope()
                scopeCode = ""
            End If

            ' 2. the start of a procedure: its signature names the parameters
            If Not inProcedure Then
                name = ProcedureStartOf(code)
                If Len(name) > 0 Then
                    CheckProcedureName filePath, index + 1, name, lines(index)
                    inProcedure = True
                    Set localScope = NewScope()
                    Set localLines = NewScope()
                    scopeCode = ""
                    Set params = ParamsOf(code)
                    For Each key In params
                        CheckName filePath, index + 1, key, params(key), localScope, localLines, lines(index)
                    Next
                End If
            End If

            ' 3. a declaration, in the scope that is open
            If Not IsProcedureLine(code) Then
                Set names = DeclaredNames(code)
                If names.Count > 0 Then
                    If inProcedure Then
                        For Each key In names
                            CheckName filePath, index + 1, key, names(key), localScope, localLines, lines(index)
                        Next
                    Else
                        For Each key In names
                            CheckName filePath, index + 1, key, names(key), moduleScope, moduleLines, lines(index)
                        Next
                    End If
                End If
            End If

            ' 4. a For Each: its control variable must be Variant or an object
            name = ForEachVariableOf(code)
            If Len(name) > 0 Then
                If inProcedure Then
                    typeName = TypeIn(localScope, name)
                    If Len(typeName) = 0 Then typeName = TypeIn(moduleScope, name)
                Else
                    typeName = TypeIn(moduleScope, name)
                End If
                If gScalars.Exists(LCase(typeName)) Then
                    AddProblem "FOREACH", filePath, index + 1, _
                               "'" & name & "' is " & typeName & ": the control variable of For Each must be Variant or Object", _
                               lines(index)
                End If
            End If

            If inProcedure Then scopeCode = scopeCode & code & vbLf

        End If
    Next

    CheckShadows filePath, localScope, localLines, scopeCode, procedures, False, lines
    CheckShadows filePath, moduleScope, moduleLines, "", procedures, True, lines
End Sub

' =====================================================================
'  the checks of one declared name
' =====================================================================

' A reserved word, a name that the same scope declares twice: VBA refuses those
' wherever they stand. A local that hides a procedure needs the whole scope (the
' call is what breaks), so that check runs in CheckShadows, when the scope closes.
Sub CheckName(ByVal filePath, ByVal lineNumber, ByVal name, ByVal typeName, ByRef scope, _
              ByRef scopeLines, ByVal codeLine)
    Dim key

    key = LCase(name)
    If gReserved.Exists(key) Then
        AddProblem "RESERVED", filePath, lineNumber, _
                   "'" & name & "' is a reserved word of VBA", codeLine
    End If

    If scope.Exists(key) Then
        AddProblem "DUPLICATE", filePath, lineNumber, _
                   "'" & name & "' is declared twice in the same scope", codeLine
    Else
        scope(key) = typeName
        scopeLines(key) = lineNumber
    End If
End Sub

' The names of a closed scope against the procedures of the module. A module level
' name that repeats a procedure is "Ambiguous name detected" for VBA, whatever the
' code does; a local or a parameter only breaks the build when the scope calls the
' name - then VBA reads the call as an array index ("Expected array", the error
' that cost an afternoon).
Sub CheckShadows(ByVal filePath, ByRef scope, ByRef scopeLines, ByVal scopeCode, _
                 ByRef procedures, ByVal isModuleScope, ByRef lines)
    Dim name, key

    For Each name In scopeLines.Keys
        key = LCase(name)
        If procedures.Exists(key) Then
            If isModuleScope Or IsCalled(scopeCode, name) Then
                AddProblem "SHADOW", filePath, scopeLines(name), _
                           "'" & name & "' hides the procedure of the same module", _
                           lines(scopeLines(name) - 1)
            End If
        End If
    Next
End Sub

' True when the code calls that name - "name(" and not "member.name(" - because a
' call of a name that a local hides is what VBA refuses.
Function IsCalled(ByVal scopeCode, ByVal name)
    IsCalled = NewRegExp("(^|[^A-Za-z0-9_.])" & name & "\s*\(").Test(scopeCode)
End Function

' The type a scope knows for a name ("" when the name is not there or is Variant).
Function TypeIn(ByRef scope, ByVal name)
    TypeIn = ""
    If scope.Exists(LCase(name)) Then TypeIn = scope(LCase(name))
End Function

' A procedure may not carry the name of a VBA statement (see FillStatements).
Sub CheckProcedureName(ByVal filePath, ByVal lineNumber, ByVal name, ByVal codeLine)
    If gStatements.Exists(LCase(name)) Then
        AddProblem "RESERVED", filePath, lineNumber, _
                   "'" & name & "' is a statement of VBA, so it cannot name a procedure", codeLine
    End If
End Sub

' =====================================================================
'  reading one line
' =====================================================================

' The line without its string literals and without a trailing comment, so a text
' that names a problem ("Dim empty" inside a message) is not taken for code. The
' literals become spaces, so the columns stay where VBA sees them.
Function CodeOf(ByVal rawLine)
    Dim text, index, character, inString

    text = ""
    inString = False
    For index = 1 To Len(rawLine)
        character = Mid(rawLine, index, 1)
        If character = """" Then
            inString = Not inString
            text = text & " "
        ElseIf inString Then
            text = text & " "
        ElseIf character = "'" Then
            Exit For
        Else
            text = text & character
        End If
    Next
    CodeOf = text
End Function

Function FirstWord(ByVal text)
    FirstWord = Match1(NewRegExp("^\s*([A-Za-z_][A-Za-z0-9_]*)"), text)
End Function

' The names of every procedure of the module, lower case: a local that hides one
' of them is the error the compiler reports as a wrong array index.
Function ProcedureNames(ByRef lines)
    Dim names, index, name

    Set names = NewScope()
    For index = 0 To UBound(lines)
        name = ProcedureStartOf(CodeOf(lines(index)))
        If Len(name) > 0 Then names(LCase(name)) = True
    Next
    Set ProcedureNames = names
End Function

Function ProcedureStartOf(ByVal codeLine)
    Dim name

    name = Match1(NewRegExp("^\s*(?:(?:Public|Private|Friend|Static)\s+)?(?:Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)"), codeLine)
    If Len(name) = 0 Then
        name = Match1(NewRegExp("^\s*(?:(?:Public|Private|Friend|Static)\s+)?Property\s+(?:Get|Let|Set)\s+([A-Za-z_][A-Za-z0-9_]*)"), codeLine)
    End If
    ProcedureStartOf = name
End Function

Function IsProcedureEnd(ByVal codeLine)
    IsProcedureEnd = NewRegExp("^\s*End\s+(?:Sub|Function|Property)\s*$").Test(codeLine)
End Function

' A line that opens a procedure declares no variable, whatever its first word
' (Private Sub ... is not a Private variable).
Function IsProcedureLine(ByVal codeLine)
    IsProcedureLine = (Len(ProcedureStartOf(codeLine)) > 0)
End Function

Function IsDeclKeyword(ByVal word)
    IsDeclKeyword = (word = "dim" Or word = "static" Or word = "private" Or _
                     word = "public" Or word = "global" Or word = "const")
End Function

Function IsProcedureWord(ByVal word)
    IsProcedureWord = (word = "sub" Or word = "function" Or word = "property" Or _
                       word = "declare" Or word = "enum" Or word = "type" Or word = "event")
End Function

' The first name of a declarator or a parameter ("ByVal a As Long" -> "a").
Function NameOf(ByVal item)
    NameOf = Match1(NewRegExp("^\s*(?:(?:ByVal|ByRef|Optional|ParamArray)\s+)*([A-Za-z_][A-Za-z0-9_]*)"), item)
End Function

' The type of a declarator ("" while the line does not say As ...). Not TypeOf:
' that word belongs to VBScript.
Function TypeNameOf(ByVal item)
    TypeNameOf = Match1(NewRegExp("\bAs\s+([A-Za-z_][A-Za-z0-9_]*)"), item)
End Function

' The control variable of a For Each line ("" while the line is not one).
Function ForEachVariableOf(ByVal codeLine)
    ForEachVariableOf = Match1(NewRegExp("^\s*For\s+Each\s+([A-Za-z_][A-Za-z0-9_]*)\s+In\b"), codeLine)
End Function

' The parameters of a signature: name (lower case) -> type.
Function ParamsOf(ByVal codeLine)
    Dim params, inside, parts, index, item, name

    Set params = NewScope()
    inside = Match1(NewRegExp("\(([^)]*)\)"), codeLine)
    If Len(inside) > 0 Then
        parts = Split(inside, ",")
        For index = 0 To UBound(parts)
            item = Trim(parts(index))
            name = NameOf(item)
            If Len(name) > 0 Then
                If Not params.Exists(LCase(name)) Then params(LCase(name)) = TypeNameOf(item)
            End If
        Next
    End If

    Set ParamsOf = params
End Function

' The names a declaration line declares, lower case, with their types: "Dim a As
' String, b As Long" gives two of them. Empty while the line declares nothing.
Function DeclaredNames(ByVal codeLine)
    Dim names, rest, parts, index, item, name

    Set names = NewScope()
    If Not IsDeclKeyword(LCase(FirstWord(codeLine))) Then
        Set DeclaredNames = names
        Exit Function
    End If

    rest = Trim(Mid(Trim(codeLine), Len(FirstWord(codeLine)) + 1))
    If LCase(Left(rest, 6)) = "const " Then rest = Trim(Mid(rest, 7))
    If Len(rest) = 0 Then
        Set DeclaredNames = names
        Exit Function
    End If
    If IsDeclKeyword(LCase(FirstWord(rest))) Or IsProcedureWord(LCase(FirstWord(rest))) Then
        Set DeclaredNames = names
        Exit Function
    End If

    parts = Split(rest, ",")
    For index = 0 To UBound(parts)
        item = Trim(parts(index))
        name = NameOf(item)
        If Len(name) > 0 Then
            If Not names.Exists(LCase(name)) Then names(LCase(name)) = TypeNameOf(item)
        End If
    Next

    Set DeclaredNames = names
End Function

' A scope: names lower case -> type ("" = Variant). Text compare, because VBA does
' not care about the case of a name either.
Function NewScope()
    Dim scope

    Set scope = CreateObject("Scripting.Dictionary")
    scope.CompareMode = 1
    Set NewScope = scope
End Function

Function NewRegExp(ByVal pattern)
    Dim re

    Set re = New RegExp
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    Set NewRegExp = re
End Function

' The first capture of a pattern ("" when the pattern does not match).
Function Match1(ByRef re, ByVal text)
    Dim matches

    Match1 = ""
    If re.Test(text) Then
        Set matches = re.Execute(text)
        If matches.Count > 0 Then
            If matches(0).SubMatches.Count > 0 Then Match1 = matches(0).SubMatches(0)
        End If
    End If
End Function
