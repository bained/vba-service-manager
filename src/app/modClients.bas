Attribute VB_Name = "modClients"
Option Explicit

' =====================================================================
'  Service_Manager - the clients domain (Phase 2)
'
'  Validation, insert, delete and the "is it in use?" question. The cells of the
'  form are read by modUiClients, so only a dictionary of column values arrives
'  here: the same code runs from a click, from a test and from a script (§5.3).
'
'  Nothing here touches the screen: the result dictionary carries the message and
'  the router writes it into the status cell.
' =====================================================================

' clients.add
Public Sub Add(ByRef data As Object, ByRef result As Object)
    Dim values As Object
    Dim problems As String
    Dim newId As Long
    Dim nameText As String
    Dim info As Object

    Set values = modApp.PayloadFrom(data)
    nameText = Trim$(modApp.PayloadText(values, COL_CLIENT_NAME))

    problems = Validate(values)
    If Len(problems) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = problems
        Exit Sub
    End If

    newId = modData.AppendClient(values)

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(newId)
    info(COL_CLIENT_NAME) = nameText
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_CLIENTS_ADDED, _
                                                          CStr(newId), nameText)
    result(modRoutes.RESULT_OK) = True
End Sub

' clients.edit / clients.add while a row is being edited: the values of the form go
' to the row with this id - it is not inserted again (§6.10). Only the four columns
' of the form are written, so a payload key like "action" cannot reach the sheet.
Public Sub Update(ByRef data As Object, ByRef result As Object)
    Dim values As Object
    Dim problems As String
    Dim idText As String
    Dim idValue As Long
    Dim nameText As String
    Dim info As Object

    idText = modApp.PayloadText(data, modRoutes.KEY_ID)
    idValue = modApp.PayloadNumber(data, modRoutes.KEY_ID)
    If idValue < 1 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_BAD_ID, idText)
        Exit Sub
    End If

    Set values = CreateObject("Scripting.Dictionary")
    values(COL_CLIENT_NAME) = modApp.PayloadText(data, COL_CLIENT_NAME)
    values(COL_CLIENT_PHONE) = modApp.PayloadText(data, COL_CLIENT_PHONE)
    values(COL_CLIENT_EMAIL) = modApp.PayloadText(data, COL_CLIENT_EMAIL)
    values(COL_CLIENT_COMMENT) = modApp.PayloadText(data, COL_CLIENT_COMMENT)

    nameText = Trim$(CStr(values(COL_CLIENT_NAME)))

    problems = Validate(values)
    If Len(problems) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = problems
        Exit Sub
    End If

    If Not modData.RowExists(SHEET_DB_CLIENTS, idValue) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_NOT_FOUND, _
                                                             "Client", CStr(idValue))
        Exit Sub
    End If

    If Not modData.UpdateClient(idValue, values) Then Exit Sub

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(idValue)
    info(COL_CLIENT_NAME) = nameText
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_CLIENTS_UPDATED, _
                                                          CStr(idValue), nameText)
    result(modRoutes.RESULT_OK) = True
End Sub

' clients.delete. The order is deliberate: exists, not in use, confirm, backup,
' delete. Every step that fails leaves the data untouched.
Public Sub Delete(ByRef data As Object, ByRef result As Object)
    Dim idText As String
    Dim idValue As Long
    Dim label As String
    Dim used As Long
    Dim info As Object

    idText = modApp.PayloadText(data, modRoutes.KEY_ID)
    idValue = modApp.PayloadNumber(data, modRoutes.KEY_ID)

    If idValue < 1 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_BAD_ID, idText)
        Exit Sub
    End If

    If Not modData.RowExists(SHEET_DB_CLIENTS, idValue) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_NOT_FOUND, _
                                                             "Client", CStr(idValue))
        Exit Sub
    End If

    used = References(idValue)
    If used > 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_IN_USE, _
                                                             "Client", CStr(idValue), CStr(used))
        Exit Sub
    End If

    If Not Confirmed(data, result, idValue) Then Exit Sub

    If Not modBackup.SaveCopy("clients-delete-" & CStr(idValue)) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_BACKUP, _
                                                             modBackup.FolderPath())
        Exit Sub
    End If

    label = modData.ClientName(idValue)
    modData.DeleteRowById SHEET_DB_CLIENTS, idValue

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(idValue)
    info(COL_CLIENT_NAME) = label
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_CLIENTS_DELETED, _
                                                          CStr(idValue), label)
    result(modRoutes.RESULT_OK) = True
End Sub

' How many appointments use this client (§8.2.5).
Public Function References(ByVal idValue As Variant) As Long
    References = modData.CountAppointmentsFor(COL_CLIENT_ID, idValue)
End Function

' The problems of the values, as one text for the status cell ("" = fine).
Public Function Validate(ByRef values As Object) As String
    Dim nameText As String
    Dim phoneText As String
    Dim emailText As String

    nameText = Trim$(modApp.PayloadText(values, COL_CLIENT_NAME))
    If Len(nameText) = 0 Then
        Validate = modText.TextFormat(modText.KEY_ERROR_REQUIRED, "name")
        Exit Function
    End If
    values(COL_CLIENT_NAME) = nameText

    phoneText = Trim$(modApp.PayloadText(values, COL_CLIENT_PHONE))
    If Len(phoneText) > 0 Then
        If DigitCount(phoneText) < MIN_PHONE_DIGITS Then
            Validate = modText.TextFormat(modText.KEY_ERROR_PHONE, CStr(MIN_PHONE_DIGITS))
            Exit Function
        End If
    End If
    values(COL_CLIENT_PHONE) = phoneText

    emailText = Trim$(modApp.PayloadText(values, COL_CLIENT_EMAIL))
    If Len(emailText) > 0 Then
        If InStr(1, emailText, "@", vbBinaryCompare) = 0 Then
            Validate = modText.TextFormat(modText.KEY_ERROR_EMAIL, emailText)
            Exit Function
        End If
    End If
    values(COL_CLIENT_EMAIL) = emailText

    values(COL_CLIENT_COMMENT) = Trim$(modApp.PayloadText(values, COL_CLIENT_COMMENT))
End Function

' =====================================================================
'  finding a client (the typed search of the booking form)
' =====================================================================

' True when every word of the term appears somewhere in the name or in the phone
' ("pet" finds "Ivan Petrov", "pet ivan" finds it too, "0888" finds the phone).
' The words are compared against one haystack, so the order in the record does
' not matter - which is the whole point: the owner cannot remember whether he
' typed the first name or the surname first.
Public Function Matches(ByVal nameText As String, ByVal phoneText As String, _
                        ByVal term As String) As Boolean
    Dim words As Variant
    Dim word As Variant
    Dim wordText As String
    Dim haystack As String

    term = Trim$(term)
    If Len(term) = 0 Then Exit Function

    haystack = LCase$(nameText & " " & phoneText)

    Matches = True
    words = Split(term, " ")
    For Each word In words
        wordText = Trim$(CStr(word))
        If Len(wordText) > 0 Then
            If InStr(1, haystack, LCase$(wordText), vbBinaryCompare) = 0 Then
                Matches = False
                Exit Function
            End If
        End If
    Next word
End Function

' The matching clients as "id|name|phone" lines, in table order. An empty term
' answers every client: that is the list the picker opens with.
Public Function SearchRows(ByVal term As String) As Collection
    Dim rows As Collection
    Dim table As SqlTable
    Dim rowIndex As Long
    Dim nameText As String
    Dim phoneText As String
    Dim idText As String

    Set rows = New Collection
    term = Trim$(term)

    Set table = modData.ClientsTable()
    For rowIndex = 1 To table.RowCount
        nameText = modData.CellTextOf(table, rowIndex, COL_CLIENT_NAME)
        phoneText = modData.CellTextOf(table, rowIndex, COL_CLIENT_PHONE)
        If Len(term) = 0 Or Matches(nameText, phoneText, term) Then
            idText = modData.CellTextOf(table, rowIndex, COL_ID)
            rows.Add idText & "|" & nameText & "|" & phoneText
        End If
    Next rowIndex

    Set SearchRows = rows
End Function

' =====================================================================
'  internals
' =====================================================================

' The two step confirmation of §6.6. With the dialogs off (tests) the answer is
' always yes, so a test is deterministic.
Private Function Confirmed(ByRef data As Object, ByRef result As Object, _
                           ByVal idValue As Long) As Boolean
    Dim keyText As String

    keyText = "clients.delete." & CStr(idValue)
    Confirmed = modApp.ConfirmNow(keyText)

    If Not Confirmed Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CANCELLED)
    End If
End Function

Private Function DigitCount(ByVal text As String) As Long
    Dim index As Long
    Dim count As Long

    For index = 1 To Len(text)
        If InStr(1, "0123456789", Mid$(text, index, 1), vbBinaryCompare) > 0 Then
            count = count + 1
        End If
    Next index

    DigitCount = count
End Function
