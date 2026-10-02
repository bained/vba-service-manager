Attribute VB_Name = "modBad"
Option Explicit
' A deliberately broken module for the self test of tools\lint.vbs:
'
'     cscript //nologo tools\lint.vbs tools\lint-selftest
'
' It must report exactly the four problems below (RESERVED, SHADOW, FOREACH,
' DUPLICATE) and nothing else - the comment and the text literals at the end name
' the same words on purpose.

Private mKeep As Object

Public Function Problems(ByVal values As Object) As String
    Problems = ""
End Function

Public Sub Apply(ByRef values As Object)
    Dim problems As String
    Dim empty As Object
    Dim key As String
    Dim value As Long
    Dim value As Long
    Dim index As Variant

    problems = Problems(values)
    Set empty = Nothing
    For Each key In values.Keys
        value = 1
    Next
    For Each index In values.Keys
        value = 2
    Next
    ' For Each value In values: a comment, not code
    Debug.Print "Dim empty As Object and For Each key In nothing"
    If value = 1 Then Debug.Print "ok"
End Sub
