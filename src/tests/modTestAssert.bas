Attribute VB_Name = "modTestAssert"
Option Explicit

' Minimal assertion helpers for the Service_Manager test modules.
' The counters are module level so that every test module can share them.

Public AssertPassed As Long
Public AssertFailed As Long
Public AssertLog As String
Public AssertNotes As String

' Every note and every failure also goes into this file, flushed at once: a run
' that dies (Excel gone, a modal dialog, the watchdog) loses the report, but this
' trace says how far it came and what it complained about.
' The file is written through the FileSystemObject: it shares the file, so reading
' the trace while a run is going cannot block the run (VBA's own Open waits for a
' reader and the two would sit there for ever).
Public Const TRACE_FILE As String = "xls-test-trace.txt"

Public Sub TraceLine(ByVal text As String)
    Dim fso As Object
    Dim file As Object
    Dim path As String

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = ThisWorkbook.path & Application.PathSeparator & TRACE_FILE
    Set file = fso.OpenTextFile(path, 8, True)      ' ForAppending, create when missing
    If Not file Is Nothing Then
        file.WriteLine text
        file.Close
    End If
    Set file = Nothing
    Set fso = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub AssertReset()
    Dim path As String

    AssertPassed = 0
    AssertFailed = 0
    AssertLog = vbNullString
    AssertNotes = vbNullString

    ' A fresh trace per run.
    On Error Resume Next
    path = ThisWorkbook.path & Application.PathSeparator & TRACE_FILE
    Kill path
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub Note(ByVal text As String)
    AssertNotes = AssertNotes & "  note: " & text & vbCrLf
    TraceLine "note: " & text
End Sub

Public Function AssertSummary() As String
    Dim text As String
    text = "assertions passed: " & CStr(AssertPassed) & ", failed: " & CStr(AssertFailed)

    If AssertFailed > 0 Then
        AssertSummary = "FAILED - " & text & vbCrLf & AssertNotes & AssertLog
    Else
        AssertSummary = "PASSED - " & text & vbCrLf & AssertNotes
    End If
End Function

Public Sub FailTest(ByVal message As String)
    AssertFailed = AssertFailed + 1
    AssertLog = AssertLog & "  FAIL: " & message & vbCrLf
    TraceLine "FAIL: " & message
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If condition Then
        AssertPassed = AssertPassed + 1
    Else
        FailTest message & " (expected True)"
    End If
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, ByVal message As String)
    If condition Then
        FailTest message & " (expected False)"
    Else
        AssertPassed = AssertPassed + 1
    End If
End Sub

Public Sub AssertEqualsText(ByVal actual As Variant, ByVal expected As String, ByVal message As String)
    If IsNull(actual) Or IsEmpty(actual) Or IsError(actual) Then
        FailTest message & " (no value; expected '" & expected & "')"
        Exit Sub
    End If

    If StrComp(CStr(actual), expected, vbTextCompare) = 0 Then
        AssertPassed = AssertPassed + 1
    Else
        FailTest message & " (expected '" & expected & "', actual '" & CStr(actual) & "')"
    End If
End Sub

Public Sub AssertEqualsLong(ByVal actual As Variant, ByVal expected As Long, ByVal message As String)
    If IsNull(actual) Or IsEmpty(actual) Or IsError(actual) Or Not IsNumeric(actual) Then
        FailTest message & " (not numeric; expected " & CStr(expected) & ")"
        Exit Sub
    End If

    If CLng(actual) = expected Then
        AssertPassed = AssertPassed + 1
    Else
        FailTest message & " (expected " & CStr(expected) & ", actual " & CStr(actual) & ")"
    End If
End Sub

Public Sub AssertEqualsDouble(ByVal actual As Variant, ByVal expected As Double, _
                              ByVal tolerance As Double, ByVal message As String)
    If IsNull(actual) Or IsEmpty(actual) Or IsError(actual) Or Not IsNumeric(actual) Then
        FailTest message & " (not numeric; expected " & CStr(expected) & ")"
        Exit Sub
    End If

    If Abs(CDbl(actual) - expected) <= tolerance Then
        AssertPassed = AssertPassed + 1
    Else
        FailTest message & " (expected " & CStr(expected) & ", actual " & CStr(actual) & ")"
    End If
End Sub

Public Sub AssertRowCount(ByRef rows As Variant, ByVal expectedDataRows As Long, ByVal message As String)
    Dim dataRows As Long
    dataRows = -1
    On Error Resume Next
    dataRows = UBound(rows, 1) - 1
    On Error GoTo 0

    If dataRows < 0 Then
        FailTest message & " (result is not a 2D array)"
        Exit Sub
    End If
    AssertEqualsLong dataRows, expectedDataRows, message
End Sub

Public Function FormatElapsed(ByVal seconds As Single) As String
    FormatElapsed = CStr(Seconds) & " s"
End Function
