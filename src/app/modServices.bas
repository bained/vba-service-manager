Attribute VB_Name = "modServices"
Option Explicit

' =====================================================================
'  Service_Manager - the services domain (Phase 2)
'
'  The same shape as modClients: validate, insert, delete, references. The one
'  difference is what a service carries: a price (money, §8.2.8) and a duration
'  in minutes (the default comes from §16.1: 10 minutes).
'
'  The duration of a service is a suggestion for a booking; it has to be
'  positive, but it does NOT have to be a multiple of the grid step: the owner
'  fixed 10 minutes as the default while the grid step is 15. The step rule
'  belongs to the start of an appointment (phase 3).
' =====================================================================

' services.add
Public Sub Add(ByRef data As Object, ByRef result As Object)
    Dim values As Object
    Dim problems As String
    Dim newId As Long
    Dim nameText As String
    Dim info As Object

    Set values = modApp.PayloadFrom(data)
    nameText = Trim$(modApp.PayloadText(values, COL_SERVICE_NAME))

    problems = Validate(values)
    If Len(problems) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = problems
        Exit Sub
    End If

    newId = modData.AppendService(values)

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(newId)
    info(COL_SERVICE_NAME) = nameText
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_SERVICES_ADDED, _
                                                          CStr(newId), nameText)
    result(modRoutes.RESULT_OK) = True
End Sub

' services.edit / services.add while a row is being edited: the row with this id
' is updated instead of a second one being inserted (§6.10).
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
    values(COL_SERVICE_NAME) = modApp.PayloadText(data, COL_SERVICE_NAME)
    values(COL_SERVICE_PRICE) = modApp.PayloadText(data, COL_SERVICE_PRICE)
    values(COL_SERVICE_DURATION) = modApp.PayloadText(data, COL_SERVICE_DURATION)

    nameText = Trim$(CStr(values(COL_SERVICE_NAME)))

    problems = Validate(values)
    If Len(problems) > 0 Then
        result(modRoutes.RESULT_MESSAGE) = problems
        Exit Sub
    End If

    If Not modData.RowExists(SHEET_DB_SERVICES, idValue) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_NOT_FOUND, _
                                                             "Service", CStr(idValue))
        Exit Sub
    End If

    If Not modData.UpdateService(idValue, values) Then Exit Sub

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(idValue)
    info(COL_SERVICE_NAME) = nameText
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_SERVICES_UPDATED, _
                                                          CStr(idValue), nameText)
    result(modRoutes.RESULT_OK) = True
End Sub

' services.delete: exists, not in use, confirm, backup, delete.
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

    If Not modData.RowExists(SHEET_DB_SERVICES, idValue) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_NOT_FOUND, _
                                                             "Service", CStr(idValue))
        Exit Sub
    End If

    used = References(idValue)
    If used > 0 Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_IN_USE, _
                                                             "Service", CStr(idValue), CStr(used))
        Exit Sub
    End If

    If Not Confirmed(data, result, idValue) Then Exit Sub

    If Not modBackup.SaveCopy("services-delete-" & CStr(idValue)) Then
        result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_ERROR_BACKUP, _
                                                             modBackup.FolderPath())
        Exit Sub
    End If

    label = modData.ServiceName(idValue)
    modData.DeleteRowById SHEET_DB_SERVICES, idValue

    Set info = CreateObject("Scripting.Dictionary")
    info(COL_ID) = CStr(idValue)
    info(COL_SERVICE_NAME) = label
    Set result(modRoutes.RESULT_DATA) = info

    result(modRoutes.RESULT_MESSAGE) = modText.TextFormat(modText.KEY_SERVICES_DELETED, _
                                                          CStr(idValue), label)
    result(modRoutes.RESULT_OK) = True
End Sub

' How many appointments use this service (§8.2.5).
Public Function References(ByVal idValue As Variant) As Long
    References = modData.CountAppointmentsFor(COL_SERVICE_ID, idValue)
End Function

' The problems of the values, as one text for the status cell ("" = fine). The
' empty price becomes 0 and the empty duration the default of §16.1.
Public Function Validate(ByRef values As Object) As String
    Dim nameText As String
    Dim priceText As String
    Dim durationText As String
    Dim amount As Double
    Dim durationValue As Long

    nameText = Trim$(modApp.PayloadText(values, COL_SERVICE_NAME))
    If Len(nameText) = 0 Then
        Validate = modText.TextFormat(modText.KEY_ERROR_REQUIRED, "service")
        Exit Function
    End If
    values(COL_SERVICE_NAME) = nameText

    priceText = Trim$(modApp.PayloadText(values, COL_SERVICE_PRICE))
    If Len(priceText) = 0 Then
        amount = 0
    ElseIf Not modApp.TryMoneyValue(priceText, amount) Then
        Validate = modText.TextFormat(modText.KEY_ERROR_PRICE, priceText)
        Exit Function
    End If
    values(COL_SERVICE_PRICE) = amount

    durationText = Trim$(modApp.PayloadText(values, COL_SERVICE_DURATION))
    If Len(durationText) = 0 Then
        values(COL_SERVICE_DURATION) = DEFAULT_DURATION_MINUTES
        Exit Function
    End If

    If Not IsNumeric(durationText) Then
        Validate = modText.TextFormat(modText.KEY_ERROR_DURATION, durationText, _
                                      CStr(DEFAULT_STEP_MINUTES))
        Exit Function
    End If

    durationValue = CLng(Val(durationText))
    If durationValue < 1 Or durationValue > MINUTES_PER_DAY Then
        Validate = modText.TextFormat(modText.KEY_ERROR_DURATION, durationText, _
                                      CStr(DEFAULT_STEP_MINUTES))
        Exit Function
    End If
    values(COL_SERVICE_DURATION) = durationValue
End Function

' =====================================================================
'  internals
' =====================================================================

' The confirmation of §6.6 (see modClients.Confirmed): one modal question.
Private Function Confirmed(ByRef data As Object, ByRef result As Object, _
                           ByVal idValue As Long) As Boolean
    Dim keyText As String

    keyText = "services.delete." & CStr(idValue)
    Confirmed = modApp.ConfirmNow(keyText)

    If Not Confirmed Then
        result(modRoutes.RESULT_MESSAGE) = modText.Text(modText.KEY_OK_CANCELLED)
    End If
End Function
