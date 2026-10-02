Attribute VB_Name = "modBackup"
Option Explicit

' =====================================================================
'  Service_Manager - a copy before a destructive action (Phase 2)
'
'  Rule (§8.2.4): a delete writes a copy of the workbook into "backups" next to
'  it first; when the copy cannot be written, the delete is refused. Only the
'  newest BACKUP_KEEP copies stay, so the folder cannot grow without an end.
'
'  The tests switch the copies off (SetEnabled False), because a test run deletes
'  rows on purpose and must not fill a folder; one test switches them on to prove
'  that the path works.
' =====================================================================

' Suppressed, not Enabled: the default value of the flag is the safe one.
Private mSuppressed As Boolean
Private mFolder As String
Private mLastPath As String

' =====================================================================
'  settings
' =====================================================================

Public Function IsEnabled() As Boolean
    IsEnabled = Not mSuppressed
End Function

Public Sub SetEnabled(ByVal value As Boolean)
    mSuppressed = Not value
End Sub

' Where the copies go; "" restores the folder next to the workbook.
Public Sub SetFolder(ByVal folderPath As String)
    mFolder = Trim$(folderPath)
End Sub

Public Function FolderPath() As String
    If Len(mFolder) > 0 Then
        FolderPath = mFolder
    Else
        FolderPath = ThisWorkbook.Path & Application.PathSeparator & BACKUP_FOLDER
    End If
End Function

' The file of the last successful copy ("" when there is none).
Public Function LastPath() As String
    LastPath = mLastPath
End Function

Public Function CountCopies() As Long
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim count As Long

    CountCopies = 0
    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FolderPath()) Then Exit Function

    Set folder = fso.GetFolder(FolderPath())
    For Each file In folder.Files
        count = count + 1
    Next file

    CountCopies = count
    Exit Function

Failed:
    Err.Clear
End Function

' =====================================================================
'  the copy
' =====================================================================

' Writes a copy of the workbook. False means "the copy could not be written" and
' the caller has to refuse the destructive action. A switched off backup answers
' True on purpose: nothing is wrong, the caller may continue.
Public Function SaveCopy(ByVal reason As String) As Boolean
    Dim fso As Object
    Dim targetFolder As String
    Dim targetPath As String

    SaveCopy = False
    If mSuppressed Then
        SaveCopy = True
        Exit Function
    End If

    mLastPath = vbNullString
    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")
    targetFolder = FolderPath()
    If Len(targetFolder) = 0 Then Exit Function

    If Not fso.FolderExists(targetFolder) Then fso.CreateFolder targetFolder
    If Not fso.FolderExists(targetFolder) Then Exit Function

    targetPath = fso.BuildPath(targetFolder, CopyName(fso, reason))
    If fso.FileExists(targetPath) Then fso.DeleteFile targetPath, True

    ThisWorkbook.SaveCopyAs targetPath
    If Err.Number <> 0 Then Exit Function
    If Not fso.FileExists(targetPath) Then Exit Function

    mLastPath = targetPath
    TrimTo BACKUP_KEEP
    SaveCopy = True
    Exit Function

Failed:
    mLastPath = vbNullString
    Err.Clear
End Function

' "Service_Manager_2026-09-17_14-05-11_clients-delete-3.xlsm"
Private Function CopyName(ByRef fso As Object, ByVal reason As String) As String
    Dim baseName As String
    Dim extension As String

    baseName = fso.GetBaseName(ThisWorkbook.Name)
    If Len(baseName) = 0 Then baseName = "workbook"

    extension = fso.GetExtensionName(ThisWorkbook.Name)
    If Len(extension) = 0 Then extension = "xlsm"

    CopyName = baseName & "_" & Format$(Now, "yyyy-mm-dd_hh-nn-ss") & "_" & _
               SafeText(reason) & "." & extension
End Function

' Only letters, digits, "-" and "_" survive, everything else becomes "-".
Private Function SafeText(ByVal text As String) As String
    Dim index As Long
    Dim character As String
    Dim result As String

    For index = 1 To Len(text)
        character = Mid$(text, index, 1)
        If InStr(1, "abcdefghijklmnopqrstuvwxyz0123456789-_", LCase$(character), _
                  vbBinaryCompare) > 0 Then
            result = result & character
        Else
            result = result & "-"
        End If
    Next index

    If Len(result) = 0 Then result = "copy"
    If Len(result) > 40 Then result = Left$(result, 40)

    SafeText = result
End Function

' Keeps the newest maxCopies files: the one with the oldest modification time is
' removed first. (Sorting by name does not work: every name starts with the name
' of the workbook, not with the timestamp.)
Private Sub TrimTo(ByVal maxCopies As Long)
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim count As Long
    Dim victimPath As String
    Dim oldestTime As Date

    If maxCopies < 1 Then Exit Sub
    On Error GoTo Failed

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(FolderPath()) Then Exit Sub

    Do
        Set folder = fso.GetFolder(FolderPath())
        count = 0
        victimPath = vbNullString
        oldestTime = 0

        For Each file In folder.Files
            count = count + 1
            If Len(victimPath) = 0 Then
                victimPath = CStr(file.Path)
                oldestTime = file.DateLastModified
            ElseIf file.DateLastModified < oldestTime Then
                victimPath = CStr(file.Path)
                oldestTime = file.DateLastModified
            End If
        Next file

        If count <= maxCopies Then Exit Do
        If Len(victimPath) = 0 Then Exit Do

        fso.DeleteFile victimPath, True
    Loop

    Exit Sub

Failed:
    Err.Clear
End Sub
