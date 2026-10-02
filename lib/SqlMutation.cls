VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlMutation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' In-memory INSERT/UPDATE facade.
' Changes are committed to Excel only by calling Commit.

Private mSource As Variant
Private mHeaders() As String
Private mHeaderIndex As Object
Private mRowCount As Long
Private mColumnCount As Long
Private mTargetRange As Range
Private mRowsAffected As Long
Private mIsOpen As Boolean
Private mHasPendingChanges As Boolean
Private mOriginalSource As Variant
Private mOriginalRowCount As Long
Private mCommitLog As Collection
Private mFilterFields As Collection
Private mFilterOperators As Collection
Private mFilterValues As Collection
Private mFilterTokens As Collection
Private mGroupOpen As Boolean
Private mFilterTypes As Collection
Private mConversionPolicy As String
Private mDiagnosticsEnabled As Boolean
Private mConversionErrors As Collection
Private mDiagnosticKeys As Object
Private mInsertPolicy As String
Private mRequiredInsertFields As Object
Private mCommitFailureLogged As Boolean

Private Sub Class_Initialize()
    Set mHeaderIndex = CreateObject("Scripting.Dictionary")
    mHeaderIndex.CompareMode = vbTextCompare
    Set mCommitLog = New Collection
    Set mFilterFields = New Collection
    Set mFilterOperators = New Collection
    Set mFilterValues = New Collection
    Set mFilterTokens = New Collection
    Set mFilterTypes = New Collection
    mConversionPolicy = "STRICT"
    mDiagnosticsEnabled = False
    Set mConversionErrors = New Collection
    Set mDiagnosticKeys = CreateObject("Scripting.Dictionary")
    mInsertPolicy = "LENIENT"
    mCommitFailureLogged = False
    Set mRequiredInsertFields = CreateObject("Scripting.Dictionary")
    mRequiredInsertFields.CompareMode = vbTextCompare
End Sub

Public Function OpenRange(ByVal sourceRange As Range) As SqlMutation
    On Error GoTo ErrorHandler
    If sourceRange Is Nothing Then Err.Raise 5, "SqlMutation.OpenRange", "The source range is Nothing."
    Set mTargetRange = sourceRange
    Dim sourceArray As Variant
    sourceArray = sourceRange.Value2
    LoadSource sourceArray, "SqlMutation.OpenRange"
    Set OpenRange = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.OpenRange", Err.Description
End Function

Public Function AutoRange(ByVal sheetName As String) As SqlMutation
    On Error GoTo ErrorHandler
    If Len(Trim$(sheetName)) = 0 Then Err.Raise 5, "SqlMutation.AutoRange", "The worksheet name cannot be empty."

    Dim sourceSheet As Worksheet
    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    If sourceSheet Is Nothing Then Err.Raise 9, "SqlMutation.AutoRange", "Worksheet not found: " & sheetName

    Dim lastRowCell As Range, lastColumnCell As Range
    Set lastRowCell = sourceSheet.Cells.Find(What:="*", After:=sourceSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    Set lastColumnCell = sourceSheet.Cells.Find(What:="*", After:=sourceSheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If lastRowCell Is Nothing Or lastColumnCell Is Nothing Then Err.Raise 5, "SqlMutation.AutoRange", "Worksheet contains no data: " & sheetName

    Set mTargetRange = sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastRowCell.Row, lastColumnCell.column))
    Dim sourceArray As Variant
    sourceArray = mTargetRange.Value2
    LoadSource sourceArray, "SqlMutation.AutoRange"
    Set AutoRange = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.AutoRange", Err.Description
End Function

Public Function Insert(ByVal values As Object) As SqlMutation
    On Error GoTo ErrorHandler
    EnsureOpen "Insert"
    If values Is Nothing Then Err.Raise 5, "SqlMutation.Insert", "The values dictionary is Nothing."
    ValidateInsertValues values

    Dim output() As Variant
    ReDim output(1 To mRowCount + 2, 1 To mColumnCount)
    CopySource output, mRowCount + 1

    Dim columnNumber As Long, headerText As String
    For columnNumber = 1 To mColumnCount
        headerText = mHeaders(columnNumber)
        If values.Exists(headerText) Then
            output(mRowCount + 2, columnNumber) = values(headerText)
        Else
            output(mRowCount + 2, columnNumber) = Empty
        End If
    Next columnNumber

    mSource = output
    mRowCount = mRowCount + 1
    mRowsAffected = 1
    mHasPendingChanges = True
    Set Insert = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.Insert", Err.Description
End Function

Public Function SetInsertPolicy(ByVal policy As String) As SqlMutation
    policy = UCase$(Trim$(policy))
    If policy <> "LENIENT" And policy <> "STRICT" Then Err.Raise 5, "SqlMutation.SetInsertPolicy", "Insert policy must be LENIENT or STRICT."
    mInsertPolicy = policy
    Set SetInsertPolicy = Me
End Function

Public Function RequireInsertColumns(ParamArray fields() As Variant) As SqlMutation
    Dim field As Variant
    For Each field In fields
        ResolveColumn field, "RequireInsertColumns"
        mRequiredInsertFields(CStr(field)) = True
    Next field
    Set RequireInsertColumns = Me
End Function

Public Function UpdateWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal values As Object) As SqlMutation
    On Error GoTo ErrorHandler
    ClearFilters
    AddFilter field, comparisonOperator, expectedValue
    Set UpdateWhere = Update(values)
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.UpdateWhere", Err.Description
End Function

Public Function Where(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlMutation
    On Error GoTo ErrorHandler
    If Not mGroupOpen Then ClearFilters
    AddFilter field, comparisonOperator, expectedValue
    Set Where = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.Where", Err.Description
End Function

Public Function AndWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlMutation
    On Error GoTo ErrorHandler
    If mFilterTokens.Count > 0 Then AddFilterOperator "AND"
    AddFilter field, comparisonOperator, expectedValue
    Set AndWhere = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.AndWhere", Err.Description
End Function

Public Function SetConversionPolicy(ByVal policy As String) As SqlMutation
    policy = UCase$(Trim$(policy))
    If policy <> "STRICT" And policy <> "IGNORE" And policy <> "NULL" Then Err.Raise 5, "SqlMutation.SetConversionPolicy", "Policy must be STRICT, IGNORE, or NULL."
    mConversionPolicy = policy
    Set SetConversionPolicy = Me
End Function

Public Function EnableDiagnostics(Optional ByVal enabled As Boolean = True) As SqlMutation
    mDiagnosticsEnabled = enabled
    Set EnableDiagnostics = Me
End Function

Public Function ClearDiagnostics() As SqlMutation
    Set mConversionErrors = New Collection
    Set mDiagnosticKeys = CreateObject("Scripting.Dictionary")
    Set ClearDiagnostics = Me
End Function

Public Function GetConversionErrors() As Variant
    Dim result() As Variant, item As Long, entry As Variant
    ReDim result(1 To mConversionErrors.Count + 1, 1 To 5)
    result(1, 1) = "SourceRow": result(1, 2) = "Context": result(1, 3) = "ValueType": result(1, 4) = "Policy": result(1, 5) = "Value"
    For item = 1 To mConversionErrors.Count
        entry = mConversionErrors(item)
        result(item + 1, 1) = entry(0): result(item + 1, 2) = entry(1): result(item + 1, 3) = entry(2): result(item + 1, 4) = entry(3): result(item + 1, 5) = entry(4)
    Next item
    GetConversionErrors = result
End Function

Public Function WhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlMutation
    ClearFilters
    AddFilterTyped field, comparisonOperator, expectedValue, valueType
    Set WhereTyped = Me
End Function

Public Function AndWhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlMutation
    If mFilterTokens.Count > 0 Then AddFilterOperator "AND"
    AddFilterTyped field, comparisonOperator, expectedValue, valueType
    Set AndWhereTyped = Me
End Function

Public Function OrWhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlMutation
    If mFilterTokens.Count > 0 Then AddFilterOperator "OR"
    AddFilterTyped field, comparisonOperator, expectedValue, valueType
    Set OrWhereTyped = Me
End Function

Public Function NotWhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlMutation
    AddFilterOperator "NOT"
    AddFilterTyped field, comparisonOperator, expectedValue, valueType
    Set NotWhereTyped = Me
End Function

Public Function UpdateWhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal values As Object, ByVal valueType As String) As SqlMutation
    Set UpdateWhereTyped = WhereTyped(field, comparisonOperator, expectedValue, valueType).Update(values)
End Function

Public Function DeleteWhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlMutation
    Set DeleteWhereTyped = WhereTyped(field, comparisonOperator, expectedValue, valueType).Delete()
End Function

Public Function OrWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlMutation
    On Error GoTo ErrorHandler
    If mFilterTokens.Count > 0 Then AddFilterOperator "OR"
    AddFilter field, comparisonOperator, expectedValue
    Set OrWhere = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.OrWhere", Err.Description
End Function

Public Function NotWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlMutation
    On Error GoTo ErrorHandler
    AddFilterOperator "NOT"
    AddFilter field, comparisonOperator, expectedValue
    Set NotWhere = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.NotWhere", Err.Description
End Function

Public Function BeginGroup() As SqlMutation
    On Error GoTo ErrorHandler
    mFilterTokens.Add "("
    mGroupOpen = True
    Set BeginGroup = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.BeginGroup", Err.Description
End Function

Public Function EndGroup() As SqlMutation
    On Error GoTo ErrorHandler
    mFilterTokens.Add ")"
    mGroupOpen = False
    Set EndGroup = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.EndGroup", Err.Description
End Function

Public Function Update(ByVal values As Object) As SqlMutation
    On Error GoTo ErrorHandler
    EnsureOpen "Update"
    If mFilterFields.Count = 0 Then Err.Raise 5, "SqlMutation.Update", "Update requires at least one filter."
    If values Is Nothing Then Err.Raise 5, "SqlMutation.Update", "The values dictionary is Nothing."
    mRowsAffected = ApplyUpdate(values)
    Set Update = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.Update", Err.Description
End Function

Public Function Delete() As SqlMutation
    On Error GoTo ErrorHandler
    EnsureOpen "Delete"
    If mFilterFields.Count = 0 Then Err.Raise 5, "SqlMutation.Delete", "Delete requires at least one filter. Use Truncate True to remove all data rows."
    mRowsAffected = ApplyDelete()
    Set Delete = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.Delete", Err.Description
End Function

Public Function MatchedCount() As Long
    On Error GoTo ErrorHandler
    EnsureOpen "MatchedCount"
    If mFilterFields.Count = 0 Then Err.Raise 5, "SqlMutation.MatchedCount", "MatchedCount requires at least one filter."
    Dim rowNumber As Long
    For rowNumber = 2 To mRowCount + 1
        If RowMatches(rowNumber) Then MatchedCount = MatchedCount + 1
    Next rowNumber
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.MatchedCount", Err.Description
End Function

Public Function Count() As Long
    Count = MatchedCount()
End Function

Public Function Preview() As Variant
    On Error GoTo ErrorHandler
    EnsureOpen "Preview"
    If mFilterTokens.Count = 0 Then Err.Raise 5, "SqlMutation.Preview", "Preview requires at least one filter."
    Dim result() As Variant, rowNumber As Long, columnNumber As Long, outputRow As Long
    Dim matchCount As Long
    For rowNumber = 2 To mRowCount + 1
        If RowMatches(rowNumber) Then matchCount = matchCount + 1
    Next rowNumber
    ReDim result(1 To matchCount + 1, 1 To mColumnCount)
    For columnNumber = 1 To mColumnCount
        result(1, columnNumber) = mHeaders(columnNumber)
    Next columnNumber
    For rowNumber = 2 To mRowCount + 1
        If RowMatches(rowNumber) Then
            outputRow = outputRow + 1
            For columnNumber = 1 To mColumnCount
                result(outputRow + 1, columnNumber) = mSource(rowNumber, LBound(mSource, 2) + columnNumber - 1)
            Next columnNumber
        End If
    Next rowNumber
    Preview = result
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.Preview", Err.Description
End Function

Public Function DeleteWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlMutation
    On Error GoTo ErrorHandler
    ClearFilters
    AddFilter field, comparisonOperator, expectedValue
    Set DeleteWhere = Delete()
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.DeleteWhere", Err.Description
End Function

Public Function Truncate(Optional ByVal confirm As Boolean = False) As SqlMutation
    On Error GoTo ErrorHandler
    EnsureOpen "Truncate"
    If Not confirm Then Err.Raise 5, "SqlMutation.Truncate", "Truncate requires explicit confirmation: Truncate True."

    Dim output() As Variant
    ReDim output(1 To 1, 1 To mColumnCount)
    Dim columnNumber As Long
    For columnNumber = 1 To mColumnCount
        output(1, columnNumber) = mSource(LBound(mSource, 1), LBound(mSource, 2) + columnNumber - 1)
    Next columnNumber

    mRowsAffected = mRowCount
    mSource = output
    mRowCount = 0
    mHasPendingChanges = (mRowsAffected > 0)
    Set Truncate = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.Truncate", Err.Description
End Function

Public Function Commit() As SqlMutation
    On Error GoTo ErrorHandler
    EnsureOpen "Commit"
    If mTargetRange Is Nothing Then Err.Raise 5, "SqlMutation.Commit", "No worksheet destination is attached."

    Dim oldRange As Range
    Dim newRange As Range
    Dim oldRowCount As Long
    Dim newRowCount As Long
    Dim backup As Variant
    Dim worksheetName As String
    Dim rangeAddress As String
    Set oldRange = mTargetRange
    mCommitFailureLogged = False
    oldRowCount = oldRange.rows.Count
    newRowCount = mRowCount + 1
    worksheetName = oldRange.Worksheet.Name
    rangeAddress = oldRange.Address(False, False)
    On Error GoTo ValidationFailure
    ValidateCommitTarget oldRange
    backup = oldRange.Value2
    On Error GoTo CommitRestore
    Set newRange = oldRange.Resize(newRowCount, mColumnCount)
    newRange.value = mSource
    If newRowCount < oldRowCount Then DeleteTrailingWorksheetRows oldRange, newRowCount, oldRowCount
    Set mTargetRange = oldRange.Worksheet.Range(oldRange.Cells(1, 1), oldRange.Cells(newRowCount, mColumnCount))
    mOriginalSource = mSource
    mOriginalRowCount = mRowCount
    mHasPendingChanges = False
    mRowsAffected = 0
    AddCommitLog worksheetName, rangeAddress, newRowCount - oldRowCount, "SUCCESS", ""
    Set Commit = Me
    Exit Function
ValidationFailure:
    Dim validationErrorNumber As Long
    Dim validationErrorDescription As String
    validationErrorNumber = Err.Number
    validationErrorDescription = Err.Description
    AddCommitLog worksheetName, rangeAddress, newRowCount - oldRowCount, _
                 "FAILED", validationErrorDescription
    mCommitFailureLogged = True
    Err.Raise validationErrorNumber, "SqlMutation.Commit", validationErrorDescription
CommitRestore:
    Dim commitErrorNumber As Long
    Dim commitErrorDescription As String
    commitErrorNumber = Err.Number
    commitErrorDescription = Err.Description
    AddCommitLog worksheetName, rangeAddress, newRowCount - oldRowCount, _
                 "FAILED", commitErrorDescription
    mCommitFailureLogged = True
    On Error Resume Next
    oldRange.Resize(oldRowCount, mColumnCount).value = backup
    On Error GoTo 0
    Err.Raise commitErrorNumber, "SqlMutation.Commit", commitErrorDescription
ErrorHandler:
    If Not mCommitFailureLogged Then
        If Len(worksheetName) > 0 Then
            AddCommitLog worksheetName, rangeAddress, newRowCount - oldRowCount, _
                         "FAILED", Err.Description
            mCommitFailureLogged = True
        End If
    End If
    Err.Raise Err.Number, "SqlMutation.Commit", Err.Description
End Function

Private Sub ValidateCommitTarget(ByVal targetRange As Range)
    On Error GoTo ValidationError
    If targetRange Is Nothing Then Err.Raise 5, "SqlMutation.ValidateCommitTarget", "The target range is Nothing."
    If targetRange.Worksheet Is Nothing Then Err.Raise 5, "SqlMutation.ValidateCommitTarget", "The target worksheet is invalid."
    If targetRange.columns.Count <> mColumnCount Then _
        Err.Raise 5, "SqlMutation.ValidateCommitTarget", "The target column count has changed."
    If targetRange.rows.Count <> mOriginalRowCount + 1 Then _
        Err.Raise 5, "SqlMutation.ValidateCommitTarget", "The target row count has changed."
    If Not ArraysMatch(targetRange.Value2, mOriginalSource) Then _
        Err.Raise 5, "SqlMutation.ValidateCommitTarget", _
                  "The target range was changed outside SqlMutation."
    Exit Sub
ValidationError:
    Err.Raise Err.Number, "SqlMutation.ValidateCommitTarget", Err.Description
End Sub

Private Function ArraysMatch(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    On Error GoTo NotEqual
    If Not IsArray(leftValue) Or Not IsArray(rightValue) Then
        ArraysMatch = (CStr(leftValue) = CStr(rightValue))
        Exit Function
    End If
    If UBound(leftValue, 1) - LBound(leftValue, 1) <> UBound(rightValue, 1) - LBound(rightValue, 1) Then Exit Function
    If UBound(leftValue, 2) - LBound(leftValue, 2) <> UBound(rightValue, 2) - LBound(rightValue, 2) Then Exit Function

    Dim rowNumber As Long, columnNumber As Long
    For rowNumber = 0 To UBound(leftValue, 1) - LBound(leftValue, 1)
        For columnNumber = 0 To UBound(leftValue, 2) - LBound(leftValue, 2)
            If Not ValuesEqualForCommit( _
                leftValue(LBound(leftValue, 1) + rowNumber, LBound(leftValue, 2) + columnNumber), _
                rightValue(LBound(rightValue, 1) + rowNumber, LBound(rightValue, 2) + columnNumber)) Then Exit Function
        Next columnNumber
    Next rowNumber
    ArraysMatch = True
    Exit Function
NotEqual:
    ArraysMatch = False
End Function

Private Function ValuesEqualForCommit(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsError(leftValue) Or IsError(rightValue) Then
        ValuesEqualForCommit = IsError(leftValue) And IsError(rightValue)
    ElseIf IsNull(leftValue) Or IsNull(rightValue) Then
        ValuesEqualForCommit = IsNull(leftValue) And IsNull(rightValue)
    ElseIf IsEmpty(leftValue) Or IsEmpty(rightValue) Then
        ValuesEqualForCommit = IsEmpty(leftValue) And IsEmpty(rightValue)
    Else
        ValuesEqualForCommit = (CStr(leftValue) = CStr(rightValue))
    End If
End Function

Private Sub DeleteTrailingWorksheetRows(ByVal oldRange As Range, ByVal newRowCount As Long, ByVal oldRowCount As Long)
    Dim firstDeletedRow As Long
    Dim lastDeletedRow As Long
    firstDeletedRow = oldRange.Row + newRowCount
    lastDeletedRow = oldRange.Row + oldRowCount - 1
    If firstDeletedRow <= lastDeletedRow Then
        oldRange.Worksheet.rows(firstDeletedRow & ":" & lastDeletedRow).Delete
    End If
End Sub

Public Function Rollback() As SqlMutation
    On Error GoTo ErrorHandler
    EnsureOpen "Rollback"
    If mTargetRange Is Nothing Then Err.Raise 5, "SqlMutation.Rollback", "No worksheet source is attached."
    Dim sourceArray As Variant
    sourceArray = mTargetRange.Value2
    LoadSource sourceArray, "SqlMutation.Rollback"
    mRowsAffected = 0
    mHasPendingChanges = False
    Set Rollback = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.Rollback", Err.Description
End Function

Public Property Get RowsAffected() As Long
    RowsAffected = mRowsAffected
End Property

Public Property Get HasPendingChanges() As Boolean
    HasPendingChanges = mHasPendingChanges
End Property

Public Function GetCommitLog() As Variant
    On Error GoTo ErrorHandler
    Dim result() As Variant
    Dim item As Long
    ReDim result(1 To mCommitLog.Count + 1, 1 To 6)
    result(1, 1) = "Timestamp"
    result(1, 2) = "Worksheet"
    result(1, 3) = "Range"
    result(1, 4) = "RowDelta"
    result(1, 5) = "Status"
    result(1, 6) = "Error"

    Dim entry As Variant
    For item = 1 To mCommitLog.Count
        entry = mCommitLog(item)
        result(item + 1, 1) = entry(0)
        result(item + 1, 2) = entry(1)
        result(item + 1, 3) = entry(2)
        result(item + 1, 4) = entry(3)
        result(item + 1, 5) = entry(4)
        result(item + 1, 6) = entry(5)
    Next item
    GetCommitLog = result
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.GetCommitLog", Err.Description
End Function

Public Function ClearCommitLog() As SqlMutation
    On Error GoTo ErrorHandler
    Set mCommitLog = New Collection
    Set ClearCommitLog = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlMutation.ClearCommitLog", Err.Description
End Function

Private Sub AddCommitLog(ByVal worksheetName As String, ByVal rangeAddress As String, ByVal rowDelta As Long, ByVal status As String, ByVal errorText As String)
    Dim entry(0 To 5) As Variant
    entry(0) = Now
    entry(1) = worksheetName
    entry(2) = rangeAddress
    entry(3) = rowDelta
    entry(4) = status
    entry(5) = errorText
    mCommitLog.Add entry
End Sub

Private Sub LoadSource(ByRef sourceArray As Variant, ByVal context As String)
    If Not IsArray(sourceArray) Then Err.Raise 5, context, "The source must be a 2D array."
    If GetArrayDimensionCount(sourceArray) <> 2 Then Err.Raise 5, context, "The source must be a 2D array."
    mSource = sourceArray
    mOriginalSource = sourceArray
    mOriginalRowCount = UBound(sourceArray, 1) - LBound(sourceArray, 1)
    mRowCount = UBound(mSource, 1) - LBound(mSource, 1)
    mColumnCount = UBound(mSource, 2) - LBound(mSource, 2) + 1
    If mRowCount < 0 Or mColumnCount < 1 Then Err.Raise 5, context, "The source array is empty."

    ReDim mHeaders(1 To mColumnCount)
    Set mHeaderIndex = CreateObject("Scripting.Dictionary")
    mHeaderIndex.CompareMode = vbTextCompare
    Dim columnNumber As Long, headerText As String
    For columnNumber = 1 To mColumnCount
        headerText = CStr(mSource(LBound(mSource, 1), LBound(mSource, 2) + columnNumber - 1))
        If Len(Trim$(headerText)) = 0 Then Err.Raise 5, context, "Header column " & columnNumber & " is empty."
        If mHeaderIndex.Exists(headerText) Then Err.Raise 5, context, "Duplicate header: " & headerText
        mHeaders(columnNumber) = headerText
        mHeaderIndex.Add headerText, columnNumber
    Next columnNumber
    mIsOpen = True
End Sub

Private Sub CopySource(ByRef target() As Variant, ByVal lastSourceRow As Long)
    Dim rowNumber As Long, columnNumber As Long
    For rowNumber = 1 To lastSourceRow
        For columnNumber = 1 To mColumnCount
            target(rowNumber, columnNumber) = mSource(LBound(mSource, 1) + rowNumber - 1, LBound(mSource, 2) + columnNumber - 1)
        Next columnNumber
    Next rowNumber
End Sub

Private Function ResolveColumn(ByVal field As Variant, ByVal context As String) As Long
    If IsNumeric(field) Then
        ResolveColumn = CLng(field)
        If ResolveColumn < 1 Or ResolveColumn > mColumnCount Then Err.Raise 5, "SqlMutation." & context, "Column index is out of range."
    ElseIf mHeaderIndex.Exists(CStr(field)) Then
        ResolveColumn = mHeaderIndex(CStr(field))
    Else
        Err.Raise 5, "SqlMutation." & context, "Unknown column: " & CStr(field)
    End If
End Function

Private Sub ClearFilters()
    Set mFilterFields = New Collection
    Set mFilterOperators = New Collection
    Set mFilterValues = New Collection
    Set mFilterTokens = New Collection
    Set mFilterTypes = New Collection
    mGroupOpen = False
End Sub

Private Sub AddFilter(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant)
    AddFilterTyped field, comparisonOperator, expectedValue, "AUTO"
End Sub

Private Sub ValidateInsertValues(ByVal values As Object)
    Dim key As Variant
    Dim field As Variant
    If mInsertPolicy = "STRICT" Then
        For Each key In values.Keys
            If Not mHeaderIndex.Exists(CStr(key)) Then Err.Raise 5, "SqlMutation.Insert", "Unknown INSERT column: " & CStr(key)
        Next key
    End If
    For Each field In mRequiredInsertFields.Keys
        If Not values.Exists(CStr(field)) Then Err.Raise 5, "SqlMutation.Insert", "Missing required INSERT column: " & CStr(field)
    Next field
End Sub

Private Sub AddFilterTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String)
    ResolveColumn field, "AddFilter"
    valueType = NormalizeValueType(valueType)
    mFilterFields.Add field
    mFilterOperators.Add comparisonOperator
    mFilterValues.Add expectedValue
    mFilterTypes.Add valueType
    Dim typedPredicate(0 To 3) As Variant
    typedPredicate(0) = field
    typedPredicate(1) = comparisonOperator
    typedPredicate(2) = expectedValue
    typedPredicate(3) = valueType
    mFilterTokens.Add typedPredicate
End Sub

Private Sub AddFilterOperator(ByVal operatorText As String)
    operatorText = UCase$(Trim$(operatorText))
    If mFilterTokens.Count = 0 Then
        If operatorText <> "NOT" And operatorText <> "(" Then _
            Err.Raise 5, "SqlMutation.AddFilterOperator", _
                      "A logical operator requires a previous filter."
    ElseIf IsFilterToken(mFilterTokens.Count, "AND") Or _
           IsFilterToken(mFilterTokens.Count, "OR") Or _
           IsFilterToken(mFilterTokens.Count, "NOT") Then
        Err.Raise 5, "SqlMutation.AddFilterOperator", _
                  "A logical operator cannot follow another logical operator."
    End If
    mFilterTokens.Add operatorText
End Sub

Private Function RowMatches(ByVal sourceRow As Long) As Boolean
    Dim position As Long
    position = 1
    RowMatches = ParseOrExpression(sourceRow, position)
End Function

Private Function PredicateMatches(ByVal sourceRow As Long, ByVal predicate As Variant) As Boolean
    Dim actualValue As Variant
    actualValue = GetSourceValue(sourceRow, predicate(0))
    If predicate(3) <> "AUTO" Then
        Dim convertedActual As Variant
        If Not TryConvert(actualValue, predicate(3), sourceRow, convertedActual) Then Exit Function
        Dim expectedValue As Variant
        expectedValue = predicate(2)
        Dim convertedExpected As Variant
        If Not TryConvert(expectedValue, predicate(3), 0, convertedExpected) Then Exit Function
        PredicateMatches = ValuesMatch(convertedActual, predicate(1), convertedExpected)
    Else
        PredicateMatches = ValuesMatch(actualValue, predicate(1), predicate(2))
    End If
End Function

Private Function NormalizeValueType(ByVal valueType As String) As String
    valueType = UCase$(Trim$(valueType))
    Select Case valueType
        Case "AUTO", "TEXT", "LONG", "DOUBLE", "CURRENCY", "DATE", "BOOLEAN": NormalizeValueType = valueType
        Case Else: Err.Raise 5, "SqlMutation.NormalizeValueType", "Unsupported value type: " & valueType
    End Select
End Function

Private Function TryConvert(ByVal value As Variant, ByVal valueType As String, _
                            ByVal sourceRow As Long, ByRef convertedValue As Variant) As Boolean
    On Error GoTo Failed
    If IsError(value) Then Err.Raise 13
    If IsNull(value) Then Err.Raise 13
    If IsEmpty(value) Then Err.Raise 13
    Select Case valueType
        Case "TEXT": convertedValue = CStr(value)
        Case "LONG": convertedValue = CLng(value)
        Case "DOUBLE": convertedValue = CDbl(value)
        Case "CURRENCY": convertedValue = CCur(value)
        Case "DATE": convertedValue = CDate(value)
        Case "BOOLEAN": convertedValue = CBool(value)
        Case Else: convertedValue = value
    End Select
    TryConvert = True
    Exit Function
Failed:
    RecordConversionError sourceRow, "MUTATION_FILTER", valueType, value
    If mConversionPolicy = "STRICT" Then
        Err.Raise 13, "SqlMutation.TryConvert", _
                  "Cannot convert selected value [" & SafeValueText(value) & _
                  "] in source row " & sourceRow & " to " & valueType & "."
    End If
End Function

Private Function SafeValueText(ByVal value As Variant) As String
    If IsError(value) Then
        SafeValueText = "#ERROR"
    ElseIf IsNull(value) Then
        SafeValueText = "Null"
    ElseIf IsEmpty(value) Then
        SafeValueText = "Empty"
    Else
        SafeValueText = CStr(value)
    End If
End Function

Private Sub RecordConversionError(ByVal sourceRow As Long, ByVal context As String, ByVal valueType As String, ByVal value As Variant)
    If Not mDiagnosticsEnabled Then Exit Sub
    Dim valueText As String
    If IsError(value) Then
        valueText = "#ERROR"
    ElseIf IsNull(value) Then
        valueText = "Null"
    ElseIf IsEmpty(value) Then
        valueText = "Empty"
    Else
        valueText = CStr(value)
    End If
    Dim key As String
    key = CStr(sourceRow) & "|" & context & "|" & valueType & "|" & valueText
    If mDiagnosticKeys.Exists(key) Then Exit Sub
    mDiagnosticKeys.Add key, True
    Dim entry(0 To 4) As Variant
    entry(0) = sourceRow: entry(1) = context: entry(2) = valueType: entry(3) = mConversionPolicy: entry(4) = valueText
    mConversionErrors.Add entry
End Sub

Private Function ParseOrExpression(ByVal sourceRow As Long, ByRef position As Long) As Boolean
    ParseOrExpression = ParseAndExpression(sourceRow, position)
    Do While IsFilterToken(position, "OR")
        position = position + 1
        ParseOrExpression = ParseOrExpression Or ParseAndExpression(sourceRow, position)
    Loop
End Function

Private Function ParseAndExpression(ByVal sourceRow As Long, ByRef position As Long) As Boolean
    ParseAndExpression = ParseUnaryExpression(sourceRow, position)
    Do While IsFilterToken(position, "AND")
        position = position + 1
        ParseAndExpression = ParseAndExpression And ParseUnaryExpression(sourceRow, position)
    Loop
End Function

Private Function ParseUnaryExpression(ByVal sourceRow As Long, ByRef position As Long) As Boolean
    If IsFilterToken(position, "NOT") Then
        position = position + 1
        ParseUnaryExpression = Not ParseUnaryExpression(sourceRow, position)
    ElseIf IsFilterToken(position, "(") Then
        position = position + 1
        ParseUnaryExpression = ParseOrExpression(sourceRow, position)
        If IsFilterToken(position, ")") Then position = position + 1
    ElseIf position <= mFilterTokens.Count Then
        ParseUnaryExpression = PredicateMatches(sourceRow, mFilterTokens(position))
        position = position + 1
    End If
End Function

Private Function IsFilterToken(ByVal position As Long, ByVal expectedToken As String) As Boolean
    If position < 1 Or position > mFilterTokens.Count Then Exit Function
    If IsArray(mFilterTokens(position)) Then Exit Function
    IsFilterToken = (UCase$(CStr(mFilterTokens(position))) = UCase$(expectedToken))
End Function

Private Function GetSourceValue(ByVal sourceRow As Long, ByVal field As Variant) As Variant
    GetSourceValue = mSource(sourceRow, LBound(mSource, 2) + ResolveColumn(field, "GetSourceValue") - 1)
End Function

Private Function ApplyUpdate(ByVal values As Object) As Long
    Dim rowNumber As Long, sourceColumn As Long, headerText As String
    For rowNumber = 2 To mRowCount + 1
        If RowMatches(rowNumber) Then
            For sourceColumn = 1 To mColumnCount
                headerText = mHeaders(sourceColumn)
                If values.Exists(headerText) Then mSource(rowNumber, LBound(mSource, 2) + sourceColumn - 1) = values(headerText)
            Next sourceColumn
            ApplyUpdate = ApplyUpdate + 1
        End If
    Next rowNumber
    If ApplyUpdate > 0 Then mHasPendingChanges = True
End Function

Private Function ApplyDelete() As Long
    Dim remainingRows() As Variant
    ReDim remainingRows(1 To mRowCount + 1, 1 To mColumnCount)
    Dim rowNumber As Long, sourceColumn As Long, outputRow As Long
    For sourceColumn = 1 To mColumnCount
        remainingRows(1, sourceColumn) = mSource(LBound(mSource, 1), LBound(mSource, 2) + sourceColumn - 1)
    Next sourceColumn
    outputRow = 1
    For rowNumber = 2 To mRowCount + 1
        If RowMatches(rowNumber) Then
            ApplyDelete = ApplyDelete + 1
        Else
            outputRow = outputRow + 1
            For sourceColumn = 1 To mColumnCount
                remainingRows(outputRow, sourceColumn) = mSource(rowNumber, LBound(mSource, 2) + sourceColumn - 1)
            Next sourceColumn
        End If
    Next rowNumber
    If ApplyDelete > 0 Then
        Dim compacted() As Variant
        ReDim compacted(1 To outputRow, 1 To mColumnCount)
        For rowNumber = 1 To outputRow
            For sourceColumn = 1 To mColumnCount
                compacted(rowNumber, sourceColumn) = remainingRows(rowNumber, sourceColumn)
            Next sourceColumn
        Next rowNumber
        mSource = compacted
        mRowCount = outputRow - 1
        mHasPendingChanges = True
    End If
End Function

Private Function ValuesMatch(ByVal actualValue As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As Boolean
    comparisonOperator = UCase$(Trim$(comparisonOperator))
    If comparisonOperator = "IS NULL" Then ValuesMatch = IsEmpty(actualValue) Or IsNull(actualValue): Exit Function
    If comparisonOperator = "IS NOT NULL" Then ValuesMatch = Not (IsEmpty(actualValue) Or IsNull(actualValue)): Exit Function
    If comparisonOperator = "LIKE" Then ValuesMatch = (CStr(actualValue) Like CStr(expectedValue)): Exit Function

    Dim comparison As Long
    If IsNumeric(actualValue) And IsNumeric(expectedValue) Then
        If CDbl(actualValue) < CDbl(expectedValue) Then
            comparison = -1
        ElseIf CDbl(actualValue) > CDbl(expectedValue) Then
            comparison = 1
        End If
    Else
        comparison = StrComp(CStr(actualValue), CStr(expectedValue), vbTextCompare)
    End If
    Select Case comparisonOperator
        Case "=": ValuesMatch = comparison = 0
        Case "<>": ValuesMatch = comparison <> 0
        Case ">": ValuesMatch = comparison > 0
        Case "<": ValuesMatch = comparison < 0
        Case ">=": ValuesMatch = comparison >= 0
        Case "<=": ValuesMatch = comparison <= 0
        Case Else: Err.Raise 5, "SqlMutation.ValuesMatch", "Unsupported comparison operator: " & comparisonOperator
    End Select
End Function

Private Sub EnsureOpen(ByVal context As String)
    If Not mIsOpen Then Err.Raise 91, "SqlMutation." & context, "The mutation table has not been opened."
End Sub

Private Function GetArrayDimensionCount(ByRef value As Variant) As Long
    Dim dimension As Long
    Dim lowerBound As Long
    On Error GoTo Finished
    Do
        dimension = dimension + 1
        lowerBound = LBound(value, dimension)
    Loop
Finished:
    GetArrayDimensionCount = dimension - 1
End Function
