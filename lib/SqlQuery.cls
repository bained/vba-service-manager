VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Lightweight, in-memory SQL-like query object.
' The first row of the source is always treated as the header row.

Private mSource As Variant
Private mHeaders() As String
Private mHeaderIndex As Object
Private mRowCount As Long
Private mColumnCount As Long

Private mSelectFields As Collection
Private mSelectAliases As Collection
Private mSelectAliasIndex As Object
Private mComputedExpressions As Collection
Private mComputedAliases As Collection
Private mWhereFields As Collection
Private mWhereOperators As Collection
Private mWhereValues As Collection
Private mWhereTypes As Collection
Private mWhereIndexes As Collection
Private mWhereTokens As Collection
Private mOrderFields As Collection
Private mOrderDirections As Collection
Private mOrderTypes As Collection
Private mDistinct As Boolean
Private mLimit As Long
Private mOffset As Long
Private mConversionPolicy As String
Private mDiagnosticsEnabled As Boolean
Private mConversionErrors As Collection
Private mDiagnosticKeys As Object
Private mGroupFields As Collection
Private mAggregateFields As Collection
Private mAggregateFunctions As Collection
Private mAggregateAliases As Collection
Private mAggregateTypes As Collection

Private Sub Class_Initialize()
    Set mSelectFields = New Collection
    Set mSelectAliases = New Collection
    Set mSelectAliasIndex = CreateObject("Scripting.Dictionary")
    mSelectAliasIndex.CompareMode = vbTextCompare
    Set mComputedExpressions = New Collection
    Set mComputedAliases = New Collection
    Set mWhereFields = New Collection
    Set mWhereOperators = New Collection
    Set mWhereValues = New Collection
    Set mWhereTypes = New Collection
    Set mWhereIndexes = New Collection
    Set mWhereTokens = New Collection
    Set mOrderFields = New Collection
    Set mOrderDirections = New Collection
    Set mOrderTypes = New Collection
    mLimit = -1
    mOffset = 0
    mConversionPolicy = "STRICT"
    mDiagnosticsEnabled = False
    Set mConversionErrors = New Collection
    Set mDiagnosticKeys = CreateObject("Scripting.Dictionary")
    Set mGroupFields = New Collection
    Set mAggregateFields = New Collection
    Set mAggregateFunctions = New Collection
    Set mAggregateAliases = New Collection
    Set mAggregateTypes = New Collection
End Sub

Public Function LoadRange(ByVal sourceRange As Range) As SqlQuery
    On Error GoTo ErrorHandler

    If sourceRange Is Nothing Then Err.Raise 5, "SqlQuery.LoadRange", "The source range is Nothing."
    Dim sourceArray As Variant
    sourceArray = sourceRange.Value2
    LoadArray sourceArray
    Set LoadRange = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.LoadRange", Err.Description
End Function

Public Function LoadArray(ByRef sourceArray As Variant) As SqlQuery
    On Error GoTo ErrorHandler

    If Not IsArray(sourceArray) Then Err.Raise 5, "SqlQuery.LoadArray", "The source must be a 2D array."
    If GetArrayDimensionCount(sourceArray) <> 2 Then Err.Raise 5, "SqlQuery.LoadArray", "The source must be a 2D array."

    mSource = sourceArray
    mRowCount = UBound(mSource, 1) - LBound(mSource, 1) + 1
    mColumnCount = UBound(mSource, 2) - LBound(mSource, 2) + 1
    If mRowCount < 1 Or mColumnCount < 1 Then Err.Raise 5, "SqlQuery.LoadArray", "The source array is empty."

    ReDim mHeaders(1 To mColumnCount)
    Set mHeaderIndex = CreateObject("Scripting.Dictionary")
    mHeaderIndex.CompareMode = vbTextCompare

    Dim columnNumber As Long
    Dim headerText As String
    For columnNumber = 1 To mColumnCount
        headerText = CStr(mSource(LBound(mSource, 1), LBound(mSource, 2) + columnNumber - 1))
        mHeaders(columnNumber) = headerText
        If Len(Trim$(headerText)) > 0 Then
            If Not mHeaderIndex.Exists(headerText) Then mHeaderIndex.Add headerText, columnNumber
        End If
    Next columnNumber

    Set LoadArray = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.LoadArray", Err.Description
End Function

Public Function SelectColumns(ParamArray columns() As Variant) As SqlQuery
    On Error GoTo ErrorHandler

    Dim item As Variant
    Dim fieldName As String
    Dim aliasName As String
    For Each item In columns
        If Trim$(CStr(item)) = "*" Then
            If UBound(columns) <> LBound(columns) Then Err.Raise 5, "SqlQuery.SelectColumns", "The * selector cannot be combined with other columns."
            ResetSelection
            Set SelectColumns = Me
            Exit Function
        End If
        ParseSelectItem CStr(item), fieldName, aliasName
        ResolveColumn fieldName
        If mSelectAliasIndex.Exists(aliasName) Then Err.Raise 5, "SqlQuery.SelectColumns", "Duplicate SELECT alias: " & aliasName
        mSelectFields.Add fieldName
        mSelectAliases.Add aliasName
        mSelectAliasIndex.Add aliasName, True
    Next item

    Set SelectColumns = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.SelectColumns", Err.Description
End Function

Public Function SelectAll() As SqlQuery
    On Error GoTo ErrorHandler
    ResetSelection
    Set SelectAll = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.SelectAll", Err.Description
End Function

Public Function SelectComputed(ByVal expression As String, ByVal aliasName As String) As SqlQuery
    On Error GoTo ErrorHandler
    expression = Trim$(expression)
    aliasName = Trim$(aliasName)
    If Len(expression) = 0 Or Len(aliasName) = 0 Then Err.Raise 5, "SqlQuery.SelectComputed", "Expression and alias are required."
    If mSelectAliasIndex.Exists(aliasName) Then Err.Raise 5, "SqlQuery.SelectComputed", "Duplicate SELECT alias: " & aliasName
    mSelectAliasIndex.Add aliasName, True
    mComputedExpressions.Add expression
    mComputedAliases.Add aliasName
    Set SelectComputed = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.SelectComputed", Err.Description
End Function

Public Function AddComputed(ByVal expression As String, ByVal aliasName As String) As SqlQuery
    Set AddComputed = SelectComputed(expression, aliasName)
End Function

Public Function GroupBy(ParamArray fields() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Dim field As Variant
    For Each field In fields
        ResolveColumn field
        mGroupFields.Add field
    Next field
    Set GroupBy = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.GroupBy", Err.Description
End Function

Public Function Aggregate(ByVal aggregateFunction As String, ByVal field As Variant, Optional ByVal aliasName As String = "", Optional ByVal valueType As String = "AUTO") As SqlQuery
    On Error GoTo ErrorHandler
    aggregateFunction = UCase$(Trim$(aggregateFunction))
    Select Case aggregateFunction
        Case "COUNT", "SUM", "AVG", "MIN", "MAX"
        Case Else: Err.Raise 5, "SqlQuery.Aggregate", "Unsupported aggregate function: " & aggregateFunction
    End Select
    If aggregateFunction = "COUNT" And CStr(field) <> "*" Then ResolveColumn field
    If aggregateFunction <> "COUNT" Then ResolveColumn field
    valueType = NormalizeValueType(valueType)
    If Len(Trim$(aliasName)) = 0 Then aliasName = aggregateFunction & "(" & CStr(field) & ")"
    mAggregateFields.Add field
    mAggregateFunctions.Add aggregateFunction
    mAggregateAliases.Add aliasName
    mAggregateTypes.Add valueType
    Set Aggregate = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.Aggregate", Err.Description
End Function

Public Function CountRows(Optional ByVal aliasName As String = "Count") As SqlQuery
    Set CountRows = Aggregate("COUNT", "*", aliasName)
End Function

Public Function Count(ByVal field As Variant, Optional ByVal aliasName As String = "") As SqlQuery
    Set Count = Aggregate("COUNT", field, aliasName)
End Function

Public Function Sum(ByVal field As Variant, Optional ByVal aliasName As String = "", Optional ByVal valueType As String = "DOUBLE") As SqlQuery
    Set Sum = Aggregate("SUM", field, aliasName, valueType)
End Function

Public Function Avg(ByVal field As Variant, Optional ByVal aliasName As String = "", Optional ByVal valueType As String = "DOUBLE") As SqlQuery
    Set Avg = Aggregate("AVG", field, aliasName, valueType)
End Function

Public Function Min(ByVal field As Variant, Optional ByVal aliasName As String = "", Optional ByVal valueType As String = "AUTO") As SqlQuery
    Set Min = Aggregate("MIN", field, aliasName, valueType)
End Function

Public Function Max(ByVal field As Variant, Optional ByVal aliasName As String = "", Optional ByVal valueType As String = "AUTO") As SqlQuery
    Set Max = Aggregate("MAX", field, aliasName, valueType)
End Function

Public Function Distinct() As SqlQuery
    On Error GoTo ErrorHandler
    mDistinct = True
    Set Distinct = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.Distinct", Err.Description
End Function

Public Function SelectDistinct() As SqlQuery
    On Error GoTo ErrorHandler
    mDistinct = True
    Set SelectDistinct = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.SelectDistinct", Err.Description
End Function

Public Function Limit(ByVal RowCount As Long) As SqlQuery
    On Error GoTo ErrorHandler
    If RowCount < 0 Then Err.Raise 5, "SqlQuery.Limit", "LIMIT cannot be negative."
    mLimit = RowCount
    Set Limit = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.Limit", Err.Description
End Function

Public Function offset(ByVal RowCount As Long) As SqlQuery
    On Error GoTo ErrorHandler
    If RowCount < 0 Then Err.Raise 5, "SqlQuery.Offset", "OFFSET cannot be negative."
    mOffset = RowCount
    Set offset = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.Offset", Err.Description
End Function

Public Function Page(ByVal RowCount As Long, Optional ByVal offsetRows As Long = 0) As SqlQuery
    On Error GoTo ErrorHandler
    If RowCount < 0 Then Err.Raise 5, "SqlQuery.Page", "Page size cannot be negative."
    If offsetRows < 0 Then Err.Raise 5, "SqlQuery.Page", "Page offset cannot be negative."
    mLimit = RowCount
    mOffset = offsetRows
    Set Page = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.Page", Err.Description
End Function

Public Function SetConversionPolicy(ByVal policy As String) As SqlQuery
    On Error GoTo ErrorHandler
    policy = UCase$(Trim$(policy))
    Select Case policy
        Case "STRICT", "IGNORE", "NULL"
            mConversionPolicy = policy
        Case Else
            Err.Raise 5, "SqlQuery.SetConversionPolicy", "Policy must be STRICT, IGNORE, or NULL."
    End Select
    Set SetConversionPolicy = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.SetConversionPolicy", Err.Description
End Function

Public Function EnableDiagnostics(Optional ByVal enabled As Boolean = True) As SqlQuery
    On Error GoTo ErrorHandler
    mDiagnosticsEnabled = enabled
    Set EnableDiagnostics = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.EnableDiagnostics", Err.Description
End Function

Public Function ClearDiagnostics() As SqlQuery
    On Error GoTo ErrorHandler
    Set mConversionErrors = New Collection
    Set mDiagnosticKeys = CreateObject("Scripting.Dictionary")
    Set ClearDiagnostics = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.ClearDiagnostics", Err.Description
End Function

Public Function GetConversionErrors() As Variant
    On Error GoTo ErrorHandler
    Dim result() As Variant
    Dim item As Long
    ReDim result(1 To mConversionErrors.Count + 1, 1 To 5)
    result(1, 1) = "SourceRow"
    result(1, 2) = "Context"
    result(1, 3) = "ValueType"
    result(1, 4) = "Policy"
    result(1, 5) = "Value"
    Dim entry As Variant
    For item = 1 To mConversionErrors.Count
        entry = mConversionErrors(item)
        result(item + 1, 1) = entry(0)
        result(item + 1, 2) = entry(1)
        result(item + 1, 3) = entry(2)
        result(item + 1, 4) = entry(3)
        result(item + 1, 5) = entry(4)
    Next item
    GetConversionErrors = result
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.GetConversionErrors", Err.Description
End Function

Public Function InnerJoin(ByVal rightSource As Variant, ByVal leftField As Variant, ByVal rightField As Variant, Optional ByVal keyType As String = "AUTO", Optional ByVal rightAlias As String = "") As SqlQuery
    On Error GoTo ErrorHandler
    JoinSource rightSource, leftField, rightField, keyType, False, rightAlias
    Set InnerJoin = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.InnerJoin", Err.Description
End Function

Public Function LeftJoin(ByVal rightSource As Variant, ByVal leftField As Variant, ByVal rightField As Variant, Optional ByVal keyType As String = "AUTO", Optional ByVal rightAlias As String = "") As SqlQuery
    On Error GoTo ErrorHandler
    JoinSource rightSource, leftField, rightField, keyType, True, rightAlias
    Set LeftJoin = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.LeftJoin", Err.Description
End Function

Public Function Where(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhere field, comparisonOperator, expectedValue, "AUTO"
    Set Where = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.Where", Err.Description
End Function

Public Function WhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhere field, comparisonOperator, expectedValue, valueType
    Set WhereTyped = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereTyped", Err.Description
End Function

Public Function AndWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhere field, comparisonOperator, expectedValue, "AUTO"
    Set AndWhere = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhere", Err.Description
End Function

Public Function OrWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddLogicalToken "OR"
    AddWhere field, comparisonOperator, expectedValue, "AUTO"
    Set OrWhere = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrWhere", Err.Description
End Function

Public Function NotWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddLogicalToken "NOT"
    AddWhere field, comparisonOperator, expectedValue, "AUTO"
    Set NotWhere = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.NotWhere", Err.Description
End Function

Public Function OrWhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlQuery
    AddLogicalToken "OR"
    AddWhere field, comparisonOperator, expectedValue, valueType
    Set OrWhereTyped = Me
End Function

Public Function NotWhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlQuery
    AddLogicalToken "NOT"
    AddWhere field, comparisonOperator, expectedValue, valueType
    Set NotWhereTyped = Me
End Function

Public Function BeginGroup() As SqlQuery
    AddLogicalToken "("
    Set BeginGroup = Me
End Function

Public Function EndGroup() As SqlQuery
    If mWhereTokens.Count = 0 Then Err.Raise 5, "SqlQuery.EndGroup", "No open filter group."
    mWhereTokens.Add ")"
    Set EndGroup = Me
End Function

Public Function AndWhereTyped(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhere field, comparisonOperator, expectedValue, valueType
    Set AndWhereTyped = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereTyped", Err.Description
End Function

Public Function WhereNull(ByVal field As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereNull field, False
    Set WhereNull = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereNull", Err.Description
End Function

Public Function WhereNotNull(ByVal field As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereNull field, True
    Set WhereNotNull = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereNotNull", Err.Description
End Function

Public Function AndWhereNull(ByVal field As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereNull field, False
    Set AndWhereNull = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereNull", Err.Description
End Function

Public Function AndWhereNotNull(ByVal field As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereNull field, True
    Set AndWhereNotNull = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereNotNull", Err.Description
End Function

Public Function WhereIn(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "AUTO"
    Set WhereIn = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereIn", Err.Description
End Function

Public Function WhereInTyped(ByVal field As Variant, ByVal valueType As String, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, valueType
    Set WhereInTyped = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereInTyped", Err.Description
End Function

Public Function AndWhereIn(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "AUTO"
    Set AndWhereIn = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereIn", Err.Description
End Function

Public Function AndWhereInTyped(ByVal field As Variant, ByVal valueType As String, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, valueType
    Set AndWhereInTyped = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereInTyped", Err.Description
End Function

Public Function OrderBy(ByVal field As Variant, Optional ByVal direction As String = "ASC") As SqlQuery
    On Error GoTo ErrorHandler

    ResolveOrderColumn field
    direction = UCase$(Trim$(direction))
    If direction <> "ASC" And direction <> "DESC" Then Err.Raise 5, "SqlQuery.OrderBy", "Direction must be ASC or DESC."
    mOrderFields.Add field
    mOrderDirections.Add direction
    mOrderTypes.Add "AUTO"

    Set OrderBy = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrderBy", Err.Description
End Function

Public Function OrderByTyped(ByVal field As Variant, Optional ByVal direction As String = "ASC", Optional ByVal valueType As String = "AUTO") As SqlQuery
    On Error GoTo ErrorHandler

    ResolveOrderColumn field
    direction = UCase$(Trim$(direction))
    If direction <> "ASC" And direction <> "DESC" Then Err.Raise 5, "SqlQuery.OrderByTyped", "Direction must be ASC or DESC."
    valueType = NormalizeValueType(valueType)
    mOrderFields.Add field
    mOrderDirections.Add direction
    mOrderTypes.Add valueType

    Set OrderByTyped = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrderByTyped", Err.Description
End Function

Public Function WhereAsText(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set WhereAsText = WhereTyped(field, comparisonOperator, expectedValue, "TEXT")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereAsText", Err.Description
End Function

Public Function WhereAsLong(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set WhereAsLong = WhereTyped(field, comparisonOperator, expectedValue, "LONG")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereAsLong", Err.Description
End Function

Public Function WhereAsDouble(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set WhereAsDouble = WhereTyped(field, comparisonOperator, expectedValue, "DOUBLE")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereAsDouble", Err.Description
End Function

Public Function WhereAsCurrency(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set WhereAsCurrency = WhereTyped(field, comparisonOperator, expectedValue, "CURRENCY")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereAsCurrency", Err.Description
End Function

Public Function WhereAsDate(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set WhereAsDate = WhereTyped(field, comparisonOperator, expectedValue, "DATE")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereAsDate", Err.Description
End Function

Public Function WhereAsBoolean(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set WhereAsBoolean = WhereTyped(field, comparisonOperator, expectedValue, "BOOLEAN")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereAsBoolean", Err.Description
End Function

Public Function AndWhereAsText(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set AndWhereAsText = AndWhereTyped(field, comparisonOperator, expectedValue, "TEXT")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereAsText", Err.Description
End Function

Public Function AndWhereAsLong(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set AndWhereAsLong = AndWhereTyped(field, comparisonOperator, expectedValue, "LONG")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereAsLong", Err.Description
End Function

Public Function AndWhereAsDouble(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set AndWhereAsDouble = AndWhereTyped(field, comparisonOperator, expectedValue, "DOUBLE")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereAsDouble", Err.Description
End Function

Public Function AndWhereAsCurrency(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set AndWhereAsCurrency = AndWhereTyped(field, comparisonOperator, expectedValue, "CURRENCY")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereAsCurrency", Err.Description
End Function

Public Function AndWhereAsDate(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set AndWhereAsDate = AndWhereTyped(field, comparisonOperator, expectedValue, "DATE")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereAsDate", Err.Description
End Function

Public Function AndWhereAsBoolean(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    Set AndWhereAsBoolean = AndWhereTyped(field, comparisonOperator, expectedValue, "BOOLEAN")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereAsBoolean", Err.Description
End Function

Public Function WhereInAsText(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "TEXT"
    Set WhereInAsText = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereInAsText", Err.Description
End Function

Public Function WhereInAsLong(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "LONG"
    Set WhereInAsLong = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereInAsLong", Err.Description
End Function

Public Function WhereInAsDouble(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "DOUBLE"
    Set WhereInAsDouble = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereInAsDouble", Err.Description
End Function

Public Function WhereInAsCurrency(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "CURRENCY"
    Set WhereInAsCurrency = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereInAsCurrency", Err.Description
End Function

Public Function WhereInAsDate(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "DATE"
    Set WhereInAsDate = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereInAsDate", Err.Description
End Function

Public Function WhereInAsBoolean(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "BOOLEAN"
    Set WhereInAsBoolean = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WhereInAsBoolean", Err.Description
End Function

Public Function AndWhereInAsText(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "TEXT"
    Set AndWhereInAsText = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereInAsText", Err.Description
End Function

Public Function AndWhereInAsLong(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "LONG"
    Set AndWhereInAsLong = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereInAsLong", Err.Description
End Function

Public Function AndWhereInAsDouble(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "DOUBLE"
    Set AndWhereInAsDouble = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereInAsDouble", Err.Description
End Function

Public Function AndWhereInAsCurrency(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "CURRENCY"
    Set AndWhereInAsCurrency = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereInAsCurrency", Err.Description
End Function

Public Function AndWhereInAsDate(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "DATE"
    Set AndWhereInAsDate = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereInAsDate", Err.Description
End Function

Public Function AndWhereInAsBoolean(ByVal field As Variant, ParamArray values() As Variant) As SqlQuery
    On Error GoTo ErrorHandler
    AddWhereIn field, values, "BOOLEAN"
    Set AndWhereInAsBoolean = Me
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.AndWhereInAsBoolean", Err.Description
End Function

Public Function OrderByAsText(ByVal field As Variant, Optional ByVal direction As String = "ASC") As SqlQuery
    On Error GoTo ErrorHandler
    Set OrderByAsText = OrderByTyped(field, direction, "TEXT")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrderByAsText", Err.Description
End Function

Public Function OrderByAsLong(ByVal field As Variant, Optional ByVal direction As String = "ASC") As SqlQuery
    On Error GoTo ErrorHandler
    Set OrderByAsLong = OrderByTyped(field, direction, "LONG")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrderByAsLong", Err.Description
End Function

Public Function OrderByAsDouble(ByVal field As Variant, Optional ByVal direction As String = "ASC") As SqlQuery
    On Error GoTo ErrorHandler
    Set OrderByAsDouble = OrderByTyped(field, direction, "DOUBLE")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrderByAsDouble", Err.Description
End Function

Public Function OrderByAsCurrency(ByVal field As Variant, Optional ByVal direction As String = "ASC") As SqlQuery
    On Error GoTo ErrorHandler
    Set OrderByAsCurrency = OrderByTyped(field, direction, "CURRENCY")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrderByAsCurrency", Err.Description
End Function

Public Function OrderByAsDate(ByVal field As Variant, Optional ByVal direction As String = "ASC") As SqlQuery
    On Error GoTo ErrorHandler
    Set OrderByAsDate = OrderByTyped(field, direction, "DATE")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrderByAsDate", Err.Description
End Function

Public Function OrderByAsBoolean(ByVal field As Variant, Optional ByVal direction As String = "ASC") As SqlQuery
    On Error GoTo ErrorHandler
    Set OrderByAsBoolean = OrderByTyped(field, direction, "BOOLEAN")
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.OrderByAsBoolean", Err.Description
End Function

Public Function ToArray() As Variant
    On Error GoTo ErrorHandler

    If mGroupFields.Count > 0 Or mAggregateFields.Count > 0 Then
        ToArray = BuildAggregateResult()
        Exit Function
    End If

    Dim selectedCount As Long
    selectedCount = mSelectFields.Count
    If selectedCount = 0 Then selectedCount = mColumnCount
    selectedCount = selectedCount + mComputedExpressions.Count

    Dim selectedColumns() As Long
    Dim outputHeaders() As String
    ReDim selectedColumns(1 To selectedCount)
    ReDim outputHeaders(1 To selectedCount)
    BuildSelection selectedColumns, outputHeaders

    Dim rowNumbers() As Long
    Dim matchingRows As Long
    matchingRows = BuildMatchingRows(rowNumbers)
    If mDistinct And matchingRows > 1 Then matchingRows = BuildDistinctRows(rowNumbers, matchingRows, selectedColumns, selectedCount)
    If matchingRows > 1 And mOrderFields.Count > 0 Then QuickSortRows rowNumbers, 1, matchingRows
    matchingRows = ApplyPaging(rowNumbers, matchingRows)

    Dim result() As Variant
    ReDim result(1 To matchingRows + 1, 1 To selectedCount)
    Dim outputColumn As Long
    Dim outputRow As Long
    Dim sourceSelectionCount As Long
    sourceSelectionCount = mSelectFields.Count
    If sourceSelectionCount = 0 Then sourceSelectionCount = mColumnCount
    For outputColumn = 1 To selectedCount
        result(1, outputColumn) = outputHeaders(outputColumn)
    Next outputColumn

    Dim sourceRow As Long
    For outputRow = 1 To matchingRows
        sourceRow = rowNumbers(outputRow)
        For outputColumn = 1 To selectedCount
            If outputColumn <= sourceSelectionCount Then
                result(outputRow + 1, outputColumn) = mSource(LBound(mSource, 1) + sourceRow - 1, LBound(mSource, 2) + selectedColumns(outputColumn) - 1)
            Else
                result(outputRow + 1, outputColumn) = EvaluateComputed(mComputedExpressions(outputColumn - sourceSelectionCount), sourceRow)
            End If
        Next outputColumn
    Next outputRow

    ToArray = result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.ToArray", Err.Description
End Function

Private Function BuildAggregateResult() As Variant
    If mAggregateFields.Count = 0 Then Err.Raise 5, "SqlQuery.BuildAggregateResult", "GROUP BY requires at least one aggregate."
    ValidateAggregateSelection

    Dim rows() As Long, RowCount As Long
    RowCount = BuildMatchingRows(rows)
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare

    Dim rowNumber As Long, key As String
    For rowNumber = 1 To RowCount
        key = BuildGroupKey(rows(rowNumber))
        If Not groups.Exists(key) Then
            Dim groupRows As Collection
            Set groupRows = New Collection
            groups.Add key, groupRows
        End If
        groups(key).Add rows(rowNumber)
    Next rowNumber

    Dim outputColumns As Long
    outputColumns = mGroupFields.Count + mAggregateFields.Count
    Dim result() As Variant
    ReDim result(1 To groups.Count + 1, 1 To outputColumns)

    Dim item As Long, groupKey As Variant
    For item = 1 To mGroupFields.Count
        result(1, item) = GetFieldName(mGroupFields(item))
    Next item
    For item = 1 To mAggregateFields.Count
        result(1, mGroupFields.Count + item) = mAggregateAliases(item)
    Next item

    Dim outputRow As Long, groupValues As Collection
    For Each groupKey In groups.Keys
        outputRow = outputRow + 1
        Set groupValues = groups(groupKey)
        For item = 1 To mGroupFields.Count
            result(outputRow + 1, item) = GetSourceValue(groupValues(1), mGroupFields(item))
        Next item
        For item = 1 To mAggregateFields.Count
            result(outputRow + 1, mGroupFields.Count + item) = CalculateAggregate(mAggregateFunctions(item), mAggregateFields(item), mAggregateTypes(item), groupValues)
        Next item
    Next groupKey
    ApplyAggregateResultOperations result, outputColumns
    BuildAggregateResult = result
End Function

Private Sub ApplyAggregateResultOperations(ByRef result() As Variant, ByVal ColumnCount As Long)
    Dim dataRows As Long
    dataRows = UBound(result, 1) - 1
    If dataRows <= 0 Then Exit Sub

    If mOrderFields.Count > 0 Then SortAggregateRows result, dataRows, ColumnCount
    If mDistinct Then dataRows = DistinctAggregateRows(result, dataRows, ColumnCount)

    Dim startRow As Long
    Dim finishRow As Long
    startRow = mOffset + 1
    If startRow > dataRows + 1 Then startRow = dataRows + 1
    finishRow = dataRows
    If mLimit >= 0 And finishRow > startRow - 1 + mLimit Then finishRow = startRow - 1 + mLimit

    If startRow > 1 Or finishRow < dataRows Then
        Dim paged() As Variant
        Dim rowNumber As Long, columnNumber As Long, outputRow As Long
        ReDim paged(1 To finishRow - startRow + 2, 1 To ColumnCount)
        For columnNumber = 1 To ColumnCount: paged(1, columnNumber) = result(1, columnNumber): Next columnNumber
        For rowNumber = startRow To finishRow
            outputRow = outputRow + 1
            For columnNumber = 1 To ColumnCount
                paged(outputRow + 1, columnNumber) = result(rowNumber + 1, columnNumber)
            Next columnNumber
        Next rowNumber
        result = paged
    End If
End Sub

Private Sub SortAggregateRows(ByRef result() As Variant, ByVal dataRows As Long, ByVal ColumnCount As Long)
    Dim i As Long, j As Long, keyColumn As Long, comparison As Long
    Dim temporary As Variant
    For i = 2 To dataRows
        For j = i + 1 To dataRows + 1
            If AggregateRowsCompare(result, i, j, comparison) Then
                For keyColumn = 1 To ColumnCount
                    temporary = result(i, keyColumn): result(i, keyColumn) = result(j, keyColumn): result(j, keyColumn) = temporary
                Next keyColumn
            End If
        Next j
    Next i
End Sub

Private Function AggregateRowsCompare(ByRef result() As Variant, ByVal leftRow As Long, ByVal rightRow As Long, ByRef finalComparison As Long) As Boolean
    Dim item As Long, leftColumn As Long, rightColumn As Long
    For item = 1 To mOrderFields.Count
        leftColumn = ResolveAggregateOutputColumn(mOrderFields(item), result)
        rightColumn = leftColumn
        finalComparison = CompareValues(result(leftRow, leftColumn), result(rightRow, rightColumn))
        If mOrderDirections(item) = "DESC" Then finalComparison = -finalComparison
        If finalComparison <> 0 Then AggregateRowsCompare = (finalComparison > 0): Exit Function
    Next item
End Function

Private Function ResolveAggregateOutputColumn(ByVal field As Variant, ByRef result() As Variant) As Long
    Dim columnNumber As Long
    For columnNumber = 1 To UBound(result, 2)
        If LCase$(CStr(result(1, columnNumber))) = LCase$(CStr(field)) Then ResolveAggregateOutputColumn = columnNumber: Exit Function
    Next columnNumber
    Err.Raise 5, "SqlQuery.ResolveAggregateOutputColumn", "Unknown aggregate output column: " & CStr(field)
End Function

Private Function DistinctAggregateRows(ByRef result() As Variant, ByVal dataRows As Long, ByVal ColumnCount As Long) As Long
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim rowNumber As Long, columnNumber As Long, writeRow As Long, key As String
    For rowNumber = 2 To dataRows + 1
        key = ""
        For columnNumber = 1 To ColumnCount: key = key & EncodeDistinctValue(result(rowNumber, columnNumber)): Next columnNumber
        If Not seen.Exists(key) Then
            seen.Add key, True: writeRow = writeRow + 1
            For columnNumber = 1 To ColumnCount: result(writeRow + 1, columnNumber) = result(rowNumber, columnNumber): Next columnNumber
        End If
    Next rowNumber
    DistinctAggregateRows = writeRow
End Function

Private Sub ValidateAggregateSelection()
    If mSelectFields.Count = 0 Then Exit Sub

    Dim item As Long
    Dim selectedColumn As Long
    Dim groupItem As Long
    Dim isGroupColumn As Boolean
    For item = 1 To mSelectFields.Count
        selectedColumn = ResolveColumn(mSelectFields(item))
        isGroupColumn = False
        For groupItem = 1 To mGroupFields.Count
            If ResolveColumn(mGroupFields(groupItem)) = selectedColumn Then
                isGroupColumn = True
                Exit For
            End If
        Next groupItem
        If Not isGroupColumn Then
            Err.Raise 5, "SqlQuery.ValidateAggregateSelection", _
                      "Selected column '" & CStr(mSelectFields(item)) & "' must appear in GROUP BY or be aggregated."
        End If
    Next item
End Sub

Private Function BuildGroupKey(ByVal sourceRow As Long) As String
    Dim item As Long
    For item = 1 To mGroupFields.Count
        BuildGroupKey = BuildGroupKey & EncodeDistinctValue(GetSourceValue(sourceRow, mGroupFields(item)))
    Next item
End Function

Private Function CalculateAggregate(ByVal aggregateFunction As String, ByVal field As Variant, ByVal valueType As String, ByVal rows As Collection) As Variant
    If aggregateFunction = "COUNT" Then
        If CStr(field) = "*" Then
            CalculateAggregate = rows.Count
            Exit Function
        End If
    End If

    Dim item As Long, validCount As Long
    Dim value As Variant, total As Double, current As Variant
    For item = 1 To rows.Count
        value = GetSourceValue(rows(item), field)
        If Not IsNullLikeValue(value) And Not IsError(value) Then
            If valueType <> "AUTO" Then
                If Not TryConvertValue(value, valueType, rows(item), "AGGREGATE") Then GoTo NextValue
            End If
            If aggregateFunction = "COUNT" Then
                validCount = validCount + 1
            ElseIf aggregateFunction = "SUM" Or aggregateFunction = "AVG" Then
                total = total + CDbl(value)
                validCount = validCount + 1
            ElseIf aggregateFunction = "MIN" Or aggregateFunction = "MAX" Then
                If validCount = 0 Then
                    current = value
                ElseIf aggregateFunction = "MIN" Then
                    If CompareValues(value, current, valueType) < 0 Then current = value
                ElseIf CompareValues(value, current, valueType) > 0 Then
                    current = value
                End If
                validCount = validCount + 1
            End If
        End If
NextValue:
    Next item

    Select Case aggregateFunction
        Case "COUNT": CalculateAggregate = validCount
        Case "SUM": CalculateAggregate = total
        Case "AVG": If validCount = 0 Then CalculateAggregate = Null Else CalculateAggregate = total / validCount
        Case "MIN", "MAX": If validCount = 0 Then CalculateAggregate = Null Else CalculateAggregate = current
    End Select
End Function

Private Function GetFieldName(ByVal field As Variant) As String
    If IsNumeric(field) Then GetFieldName = mHeaders(CLng(field)) Else GetFieldName = CStr(field)
End Function

Private Function ApplyPaging(ByRef rowNumbers() As Long, ByVal RowCount As Long) As Long
    If mOffset >= RowCount Then Exit Function

    Dim availableRows As Long
    availableRows = RowCount - mOffset
    ApplyPaging = availableRows
    If mLimit >= 0 And ApplyPaging > mLimit Then ApplyPaging = mLimit

    Dim item As Long
    For item = 1 To ApplyPaging
        rowNumbers(item) = rowNumbers(item + mOffset)
    Next item
End Function

Private Function BuildDistinctRows(ByRef rowNumbers() As Long, ByVal RowCount As Long, ByRef selectedColumns() As Long, ByVal selectedCount As Long) As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim distinctRows() As Long
    ReDim distinctRows(1 To RowCount)
    Dim sourceRow As Long
    Dim outputRow As Long
    Dim key As String

    For outputRow = 1 To RowCount
        sourceRow = rowNumbers(outputRow)
        key = BuildDistinctKey(sourceRow, selectedColumns, selectedCount)
        If Not seen.Exists(key) Then
            seen.Add key, True
            BuildDistinctRows = BuildDistinctRows + 1
            distinctRows(BuildDistinctRows) = sourceRow
        End If
    Next outputRow

    For outputRow = 1 To BuildDistinctRows
        rowNumbers(outputRow) = distinctRows(outputRow)
    Next outputRow
End Function

Private Function BuildDistinctKey(ByVal sourceRow As Long, ByRef selectedColumns() As Long, ByVal selectedCount As Long) As String
    Dim item As Long
    Dim value As Variant
    For item = 1 To selectedCount
        value = mSource(LBound(mSource, 1) + sourceRow - 1, LBound(mSource, 2) + selectedColumns(item) - 1)
        BuildDistinctKey = BuildDistinctKey & EncodeDistinctValue(value)
    Next item
End Function

Private Function EncodeDistinctValue(ByVal value As Variant) As String
    Dim valueText As String
    Select Case VarType(value)
        Case vbEmpty
            valueText = ""
        Case vbNull
            valueText = ""
        Case vbError
            valueText = "ERROR"
        Case Else
            valueText = CStr(value)
    End Select

    ' Type and length prefixes make composite keys collision-safe.
    EncodeDistinctValue = CStr(VarType(value)) & ":" & CStr(Len(valueText)) & ":" & valueText & "|"
End Function

Public Sub WriteToRange(ByVal destination As Range)
    On Error GoTo ErrorHandler
    If destination Is Nothing Then Err.Raise 5, "SqlQuery.WriteToRange", "The destination range is Nothing."

    Dim result As Variant
    result = ToArray()
    destination.Resize(UBound(result, 1), UBound(result, 2)).value = result
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "SqlQuery.WriteToRange", Err.Description
End Sub

Private Sub AddWhere(ByVal field As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String)
    If mWhereTokens.Count > 0 Then
        If IsArray(mWhereTokens(mWhereTokens.Count)) Or IsQueryToken(mWhereTokens.Count, ")") Then mWhereTokens.Add "AND"
    End If
    ResolveColumn field
    comparisonOperator = UCase$(Trim$(comparisonOperator))
    Select Case comparisonOperator
        Case "=", "<>", ">", "<", ">=", "<=", "LIKE"
        Case Else: Err.Raise 5, "SqlQuery.AddWhere", "Unsupported comparison operator: " & comparisonOperator
    End Select
    valueType = NormalizeValueType(valueType)
    mWhereFields.Add field
    mWhereOperators.Add comparisonOperator
    mWhereValues.Add expectedValue
    mWhereTypes.Add valueType
    mWhereIndexes.Add ""
    Dim predicate(0 To 4) As Variant
    predicate(0) = field
    predicate(1) = comparisonOperator
    predicate(2) = expectedValue
    predicate(3) = valueType
    predicate(4) = mWhereFields.Count
    mWhereTokens.Add predicate
End Sub

Private Sub AddLogicalToken(ByVal token As String)
    token = UCase$(Trim$(token))
    If token = "(" And mWhereTokens.Count > 0 Then
        If IsArray(mWhereTokens(mWhereTokens.Count)) Or IsQueryToken(mWhereTokens.Count, ")") Then mWhereTokens.Add "AND"
    End If
    If token = "OR" Or token = "AND" Then
        If mWhereTokens.Count = 0 Then Err.Raise 5, "SqlQuery.AddLogicalToken", "A logical operator requires a previous filter."
    End If
    mWhereTokens.Add token
End Sub

Private Sub JoinSource(ByVal rightSource As Variant, ByVal leftField As Variant, ByVal rightField As Variant, ByVal keyType As String, ByVal preserveLeft As Boolean, ByVal rightAlias As String)
    Dim rightArray As Variant
    rightArray = GetJoinSourceArray(rightSource)
    keyType = NormalizeValueType(keyType)

    Dim leftColumns() As Long
    Dim rightColumnsKey() As Long
    Dim keyCount As Long
    BuildJoinColumns leftField, rightField, rightArray, leftColumns, rightColumnsKey, keyCount

    Dim rightRows As Long, rightColumns As Long
    rightRows = UBound(rightArray, 1) - LBound(rightArray, 1) + 1
    rightColumns = UBound(rightArray, 2) - LBound(rightArray, 2) + 1

    Dim rightHeaders() As String
    Dim rightHeaderIndex As Object
    BuildArrayHeaders rightArray, rightHeaders, rightHeaderIndex

    Dim rightLookup As Object
    Set rightLookup = CreateObject("Scripting.Dictionary")
    rightLookup.CompareMode = vbTextCompare

    Dim rightRow As Long
    Dim joinKey As String
    For rightRow = 2 To rightRows
        If TryBuildCompositeJoinKey(rightArray, rightRow, rightColumnsKey, keyCount, keyType, 0, joinKey) Then
            If Not rightLookup.Exists(joinKey) Then
                Dim RowMatches As Collection
                Set RowMatches = New Collection
                rightLookup.Add joinKey, RowMatches
            End If
            rightLookup(joinKey).Add rightRow
        End If
    Next rightRow

    Dim outputHeaders() As String
    BuildJoinHeaders rightHeaders, rightColumns, outputHeaders, rightAlias

    Dim outputRows As Collection
    Set outputRows = New Collection
    Dim leftRow As Long
    Dim matches As Collection
    Dim matchItem As Long
    For leftRow = 2 To mRowCount
        joinKey = ""
        If TryBuildSourceCompositeJoinKey(leftRow, leftColumns, keyCount, keyType, joinKey) Then
            If rightLookup.Exists(joinKey) Then Set matches = rightLookup(joinKey) Else Set matches = Nothing
        Else
            Set matches = Nothing
        End If

        If matches Is Nothing Then
            If preserveLeft Then outputRows.Add BuildJoinedRow(leftRow, 0, rightArray, rightColumns)
        Else
            For matchItem = 1 To matches.Count
                outputRows.Add BuildJoinedRow(leftRow, matches(matchItem), rightArray, rightColumns)
            Next matchItem
        End If
    Next leftRow

    Dim result() As Variant
    Dim outputColumns As Long
    outputColumns = mColumnCount + rightColumns
    ReDim result(1 To outputRows.Count + 1, 1 To outputColumns)

    Dim columnNumber As Long
    For columnNumber = 1 To mColumnCount
        result(1, columnNumber) = mHeaders(columnNumber)
    Next columnNumber
    For columnNumber = 1 To rightColumns
        result(1, mColumnCount + columnNumber) = outputHeaders(columnNumber)
    Next columnNumber

    Dim outputRow As Long
    Dim joinedRow As Variant
    For outputRow = 1 To outputRows.Count
        joinedRow = outputRows(outputRow)
        For columnNumber = 1 To outputColumns
            result(outputRow + 1, columnNumber) = joinedRow(columnNumber)
        Next columnNumber
    Next outputRow

    LoadArray result
End Sub

Private Sub BuildJoinColumns(ByVal leftField As Variant, ByVal rightField As Variant, _
                             ByRef rightArray As Variant, ByRef leftColumns() As Long, _
                             ByRef rightColumns() As Long, ByRef keyCount As Long)
    Dim leftItems As Variant
    Dim rightItems As Variant
    If IsArray(leftField) Then leftItems = leftField Else leftItems = Array(leftField)
    If IsArray(rightField) Then rightItems = rightField Else rightItems = Array(rightField)
    If UBound(leftItems) - LBound(leftItems) <> UBound(rightItems) - LBound(rightItems) Then _
        Err.Raise 5, "SqlQuery.BuildJoinColumns", "JOIN key column counts must match."
    keyCount = UBound(leftItems) - LBound(leftItems) + 1
    ReDim leftColumns(1 To keyCount)
    ReDim rightColumns(1 To keyCount)
    Dim item As Long
    For item = 1 To keyCount
        leftColumns(item) = ResolveColumn(leftItems(LBound(leftItems) + item - 1))
        rightColumns(item) = ResolveArrayColumn(rightArray, rightItems(LBound(rightItems) + item - 1))
    Next item
End Sub

Private Function TryBuildCompositeJoinKey(ByRef sourceArray As Variant, ByVal rowNumber As Long, _
                                           ByRef columns() As Long, ByVal keyCount As Long, _
                                           ByVal keyType As String, ByVal sourceRow As Long, _
                                           ByRef joinKey As String) As Boolean
    Dim item As Long, value As Variant, part As String
    joinKey = ""
    For item = 1 To keyCount
        value = GetArrayValue(sourceArray, rowNumber, columns(item))
        If Not TryBuildJoinKey(value, keyType, sourceRow, part) Then Exit Function
        joinKey = joinKey & CStr(Len(part)) & ":" & part & "|"
    Next item
    TryBuildCompositeJoinKey = True
End Function

Private Function TryBuildSourceCompositeJoinKey(ByVal sourceRow As Long, ByRef columns() As Long, _
                                                 ByVal keyCount As Long, ByVal keyType As String, _
                                                 ByRef joinKey As String) As Boolean
    Dim item As Long, value As Variant, part As String
    joinKey = ""
    For item = 1 To keyCount
        value = mSource(LBound(mSource, 1) + sourceRow - 1, LBound(mSource, 2) + columns(item) - 1)
        If Not TryBuildJoinKey(value, keyType, sourceRow, part) Then Exit Function
        joinKey = joinKey & CStr(Len(part)) & ":" & part & "|"
    Next item
    TryBuildSourceCompositeJoinKey = True
End Function

Private Function GetJoinSourceArray(ByVal source As Variant) As Variant
    If IsObject(source) Then
        If TypeName(source) <> "Range" Then Err.Raise 5, "SqlQuery.GetJoinSourceArray", "Join source must be a Range or 2D array."
        GetJoinSourceArray = source.Value2
    Else
        GetJoinSourceArray = source
    End If
    If Not IsArray(GetJoinSourceArray) Then Err.Raise 5, "SqlQuery.GetJoinSourceArray", "Join source must be a Range or 2D array."
    If GetArrayDimensionCount(GetJoinSourceArray) <> 2 Then Err.Raise 5, "SqlQuery.GetJoinSourceArray", "Join source must be a 2D array."
End Function

Private Function ResolveArrayColumn(ByRef sourceArray As Variant, ByVal field As Variant) As Long
    If IsNumeric(field) Then
        ResolveArrayColumn = CLng(field)
        If ResolveArrayColumn < 1 Or ResolveArrayColumn > UBound(sourceArray, 2) - LBound(sourceArray, 2) + 1 Then Err.Raise 5, "SqlQuery.ResolveArrayColumn", "Right join column index is out of range."
        Exit Function
    End If

    Dim columnNumber As Long
    For columnNumber = 1 To UBound(sourceArray, 2) - LBound(sourceArray, 2) + 1
        If StrComp(CStr(sourceArray(LBound(sourceArray, 1), LBound(sourceArray, 2) + columnNumber - 1)), CStr(field), vbTextCompare) = 0 Then
            ResolveArrayColumn = columnNumber
            Exit Function
        End If
    Next columnNumber
    Err.Raise 5, "SqlQuery.ResolveArrayColumn", "Unknown right join column: " & CStr(field)
End Function

Private Sub BuildArrayHeaders(ByRef sourceArray As Variant, ByRef headers() As String, ByRef headerIndex As Object)
    Dim columnNumber As Long, ColumnCount As Long
    ColumnCount = UBound(sourceArray, 2) - LBound(sourceArray, 2) + 1
    ReDim headers(1 To ColumnCount)
    Set headerIndex = CreateObject("Scripting.Dictionary")
    headerIndex.CompareMode = vbTextCompare
    For columnNumber = 1 To ColumnCount
        headers(columnNumber) = CStr(sourceArray(LBound(sourceArray, 1), LBound(sourceArray, 2) + columnNumber - 1))
        If Not headerIndex.Exists(headers(columnNumber)) Then headerIndex.Add headers(columnNumber), columnNumber
    Next columnNumber
End Sub

Private Sub BuildJoinHeaders(ByRef rightHeaders() As String, ByVal rightColumns As Long, ByRef outputHeaders() As String, ByVal rightAlias As String)
    ReDim outputHeaders(1 To rightColumns)
    Dim columnNumber As Long, headerText As String
    rightAlias = Trim$(rightAlias)
    If InStr(1, rightAlias, ".", vbBinaryCompare) > 0 Then Err.Raise 5, "SqlQuery.BuildJoinHeaders", "Join alias cannot contain a period."
    For columnNumber = 1 To rightColumns
        headerText = rightHeaders(columnNumber)
        If Len(rightAlias) > 0 Then
            headerText = rightAlias & "." & headerText
        ElseIf mHeaderIndex.Exists(headerText) Then
            headerText = "Right." & headerText
        End If
        outputHeaders(columnNumber) = headerText
    Next columnNumber
End Sub

Private Function BuildJoinedRow(ByVal leftRow As Long, ByVal rightRow As Long, ByRef rightArray As Variant, ByVal rightColumns As Long) As Variant
    Dim joined() As Variant
    ReDim joined(1 To mColumnCount + rightColumns)
    Dim columnNumber As Long
    For columnNumber = 1 To mColumnCount
        joined(columnNumber) = mSource(LBound(mSource, 1) + leftRow - 1, LBound(mSource, 2) + columnNumber - 1)
    Next columnNumber
    For columnNumber = 1 To rightColumns
        If rightRow = 0 Then
            joined(mColumnCount + columnNumber) = Null
        Else
            joined(mColumnCount + columnNumber) = GetArrayValue(rightArray, rightRow, columnNumber)
        End If
    Next columnNumber
    BuildJoinedRow = joined
End Function

Private Function GetArrayValue(ByRef sourceArray As Variant, ByVal rowNumber As Long, ByVal columnNumber As Long) As Variant
    GetArrayValue = sourceArray(LBound(sourceArray, 1) + rowNumber - 1, LBound(sourceArray, 2) + columnNumber - 1)
End Function

Private Function TryBuildJoinKey(ByVal value As Variant, ByVal keyType As String, ByVal sourceRow As Long, ByRef joinKey As String) As Boolean
    If IsNullLikeValue(value) Or IsError(value) Then Exit Function
    If keyType <> "AUTO" Then
        If Not TryConvertValue(value, keyType, sourceRow) Then Exit Function
    End If
    joinKey = EncodeDistinctValue(value)
    TryBuildJoinKey = True
End Function

Private Sub AddWhereNull(ByVal field As Variant, ByVal negate As Boolean)
    If mWhereTokens.Count > 0 Then
        If IsArray(mWhereTokens(mWhereTokens.Count)) Or IsQueryToken(mWhereTokens.Count, ")") Then mWhereTokens.Add "AND"
    End If
    ResolveColumn field
    mWhereFields.Add field
    If negate Then
        mWhereOperators.Add "IS NOT NULL"
    Else
        mWhereOperators.Add "IS NULL"
    End If
    mWhereValues.Add Empty
    mWhereTypes.Add "AUTO"
    mWhereIndexes.Add ""
    Dim predicate(0 To 4) As Variant
    predicate(0) = field
    predicate(1) = IIf(negate, "IS NOT NULL", "IS NULL")
    predicate(2) = Empty
    predicate(3) = "AUTO"
    predicate(4) = mWhereFields.Count
    mWhereTokens.Add predicate
End Sub

Private Sub AddWhereIn(ByVal field As Variant, ByVal values As Variant, ByVal valueType As String)
    If mWhereTokens.Count > 0 Then
        If IsArray(mWhereTokens(mWhereTokens.Count)) Or IsQueryToken(mWhereTokens.Count, ")") Then mWhereTokens.Add "AND"
    End If
    ResolveColumn field
    valueType = NormalizeValueType(valueType)

    Dim valueList As Collection
    Set valueList = New Collection
    Dim item As Long
    Dim firstIndex As Long
    Dim lastIndex As Long
    On Error Resume Next
    firstIndex = LBound(values)
    lastIndex = UBound(values)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
    Else
        On Error GoTo 0
        If lastIndex >= firstIndex Then
            If lastIndex = firstIndex And IsArray(values(firstIndex)) Then
                AddArrayValues valueList, values(firstIndex)
            Else
                For item = firstIndex To lastIndex
                    valueList.Add values(item)
                Next item
            End If
        End If
    End If

    mWhereFields.Add field
    mWhereOperators.Add "IN"
    mWhereValues.Add valueList
    mWhereTypes.Add valueType

    If valueType <> "AUTO" Then
        Dim lookup As Object
        Set lookup = CreateObject("Scripting.Dictionary")
        Dim candidate As Variant
        For item = 1 To valueList.Count
            candidate = valueList(item)
            If TryConvertValue(candidate, valueType, 0) Then
                If Not lookup.Exists(EncodeDistinctValue(candidate)) Then lookup.Add EncodeDistinctValue(candidate), True
            ElseIf mConversionPolicy = "STRICT" Then
                Err.Raise 13, "SqlQuery.AddWhereIn", "Cannot convert expected value to " & valueType & "."
            End If
        Next item
        mWhereIndexes.Add lookup
    Else
        mWhereIndexes.Add ""
    End If
    Dim predicate(0 To 4) As Variant
    predicate(0) = field
    predicate(1) = "IN"
    predicate(2) = valueList
    predicate(3) = valueType
    predicate(4) = mWhereFields.Count
    mWhereTokens.Add predicate
End Sub

Private Sub AddArrayValues(ByVal target As Collection, ByVal sourceArray As Variant)
    If Not IsArray(sourceArray) Then
        target.Add sourceArray
        Exit Sub
    End If

    If GetArrayDimensionCount(sourceArray) <> 1 Then Err.Raise 5, "SqlQuery.AddArrayValues", "The IN list array must be one-dimensional."

    Dim item As Long
    For item = LBound(sourceArray) To UBound(sourceArray)
        target.Add sourceArray(item)
    Next item
End Sub

Private Sub BuildSelection(ByRef selectedColumns() As Long, ByRef outputHeaders() As String)
    Dim item As Long
    Dim fieldName As String
    If mSelectFields.Count = 0 Then
        For item = 1 To mColumnCount
            selectedColumns(item) = item
            outputHeaders(item) = mHeaders(item)
        Next item
    Else
        For item = 1 To mSelectFields.Count
            fieldName = CStr(mSelectFields(item))
            selectedColumns(item) = ResolveColumn(fieldName)
            outputHeaders(item) = CStr(mSelectAliases(item))
        Next item
    End If
    Dim computedIndex As Long
    For computedIndex = 1 To mComputedExpressions.Count
        outputHeaders(item + computedIndex - 1) = CStr(mComputedAliases(computedIndex))
    Next computedIndex
End Sub

Private Sub ResetSelection()
    Set mSelectFields = New Collection
    Set mSelectAliases = New Collection
    Set mSelectAliasIndex = CreateObject("Scripting.Dictionary")
    mSelectAliasIndex.CompareMode = vbTextCompare
    Set mComputedExpressions = New Collection
    Set mComputedAliases = New Collection
End Sub

Private Function EvaluateComputed(ByVal expression As String, ByVal sourceRow As Long) As Variant
    Dim normalized As String
    Dim fieldName As String
    Dim value As Variant
    normalized = UCase$(Trim$(expression))
    If Left$(normalized, 6) = "UPPER(" And Right$(normalized, 1) = ")" Then
        fieldName = Trim$(Mid$(expression, 7, Len(expression) - 7))
        value = GetSourceValue(sourceRow, fieldName)
        If IsNullLikeValue(value) Then EvaluateComputed = Null Else EvaluateComputed = UCase$(CStr(value))
    ElseIf Left$(normalized, 6) = "LOWER(" And Right$(normalized, 1) = ")" Then
        fieldName = Trim$(Mid$(expression, 7, Len(expression) - 7))
        value = GetSourceValue(sourceRow, fieldName)
        If IsNullLikeValue(value) Then EvaluateComputed = Null Else EvaluateComputed = LCase$(CStr(value))
    ElseIf Left$(normalized, 4) = "LEN(" And Right$(normalized, 1) = ")" Then
        fieldName = Trim$(Mid$(expression, 5, Len(expression) - 5))
        value = GetSourceValue(sourceRow, fieldName)
        If IsNullLikeValue(value) Then EvaluateComputed = Null Else EvaluateComputed = Len(CStr(value))
    Else
        Err.Raise 5, "SqlQuery.EvaluateComputed", "Unsupported computed expression: " & expression
    End If
End Function

Private Function BuildMatchingRows(ByRef rowNumbers() As Long) As Long
    Dim sourceRow As Long
    Dim Count As Long
    ReDim rowNumbers(1 To mRowCount)
    For sourceRow = 2 To mRowCount
        If RowMatches(sourceRow) Then
            Count = Count + 1
            rowNumbers(Count) = sourceRow
        End If
    Next sourceRow
    BuildMatchingRows = Count
End Function

Private Function RowMatches(ByVal sourceRow As Long) As Boolean
    If mWhereTokens.Count = 0 Then RowMatches = True: Exit Function
    Dim position As Long
    position = 1
    RowMatches = ParseOrExpression(sourceRow, position)
End Function

Private Function ParseOrExpression(ByVal sourceRow As Long, ByRef position As Long) As Boolean
    ParseOrExpression = ParseAndExpression(sourceRow, position)
    Do While IsQueryToken(position, "OR")
        position = position + 1
        ParseOrExpression = ParseOrExpression Or ParseAndExpression(sourceRow, position)
    Loop
End Function

Private Function ParseAndExpression(ByVal sourceRow As Long, ByRef position As Long) As Boolean
    ParseAndExpression = ParseUnaryExpression(sourceRow, position)
    Do While IsQueryToken(position, "AND")
        position = position + 1
        ParseAndExpression = ParseAndExpression And ParseUnaryExpression(sourceRow, position)
    Loop
End Function

Private Function ParseUnaryExpression(ByVal sourceRow As Long, ByRef position As Long) As Boolean
    If IsQueryToken(position, "NOT") Then
        position = position + 1
        ParseUnaryExpression = Not ParseUnaryExpression(sourceRow, position)
    ElseIf IsQueryToken(position, "(") Then
        position = position + 1
        ParseUnaryExpression = ParseOrExpression(sourceRow, position)
        If IsQueryToken(position, ")") Then position = position + 1
    ElseIf position <= mWhereTokens.Count Then
        ParseUnaryExpression = QueryPredicateMatches(sourceRow, mWhereTokens(position))
        position = position + 1
    End If
End Function

Private Function IsQueryToken(ByVal position As Long, ByVal expectedToken As String) As Boolean
    If position < 1 Or position > mWhereTokens.Count Then Exit Function
    If IsArray(mWhereTokens(position)) Then Exit Function
    IsQueryToken = (UCase$(CStr(mWhereTokens(position))) = UCase$(expectedToken))
End Function

Private Function QueryPredicateMatches(ByVal sourceRow As Long, ByVal predicate As Variant) As Boolean
    QueryPredicateMatches = ValuesMatch(GetSourceValue(sourceRow, predicate(0)), predicate(1), predicate(2), predicate(3), sourceRow, "")
End Function

Private Function ValuesMatch(ByVal actualValue As Variant, ByVal comparisonOperator As String, ByVal expectedValue As Variant, ByVal valueType As String, ByVal sourceRow As Long, Optional ByVal valueIndex As Variant) As Boolean
    If comparisonOperator = "IS NULL" Or comparisonOperator = "IS NOT NULL" Then
        ValuesMatch = IsNullLikeValue(actualValue)
        If comparisonOperator = "IS NOT NULL" Then ValuesMatch = Not ValuesMatch
        Exit Function
    End If

    If comparisonOperator = "IN" Then
        ValuesMatch = ValueIsInList(actualValue, expectedValue, valueType, sourceRow, valueIndex)
        Exit Function
    End If

    If valueType <> "AUTO" Then
        If Not TryConvertValue(actualValue, valueType, sourceRow, "WHERE") Then Exit Function
        If Not TryConvertValue(expectedValue, valueType, 0) Then Err.Raise 13, "SqlQuery.ValuesMatch", "Cannot convert expected value to " & valueType & "."
    End If
    If comparisonOperator = "LIKE" Then
        ValuesMatch = (CStr(actualValue) Like CStr(expectedValue))
        Exit Function
    End If

    Dim comparison As Long
    comparison = CompareValues(actualValue, expectedValue)
    Select Case comparisonOperator
        Case "=": ValuesMatch = (comparison = 0)
        Case "<>": ValuesMatch = (comparison <> 0)
        Case ">": ValuesMatch = (comparison > 0)
        Case "<": ValuesMatch = (comparison < 0)
        Case ">=": ValuesMatch = (comparison >= 0)
        Case "<=": ValuesMatch = (comparison <= 0)
    End Select
End Function

Private Function IsNullLikeValue(ByVal value As Variant) As Boolean
    IsNullLikeValue = (IsEmpty(value) Or IsNull(value))
End Function

Private Function ValueIsInList(ByVal actualValue As Variant, ByVal expectedValues As Collection, ByVal valueType As String, ByVal sourceRow As Long, Optional ByVal valueIndex As Variant) As Boolean
    If expectedValues.Count = 0 Then Exit Function

    If valueType <> "AUTO" And IsObject(valueIndex) Then
        If Not TryConvertValue(actualValue, valueType, sourceRow, "IN") Then Exit Function
        ValueIsInList = valueIndex.Exists(EncodeDistinctValue(actualValue))
        Exit Function
    End If

    Dim item As Long
    Dim candidate As Variant
    For item = 1 To expectedValues.Count
        candidate = expectedValues(item)
        If valueType <> "AUTO" Then
            If Not TryConvertValue(candidate, valueType, 0) Then Exit Function
            If Not TryConvertValue(actualValue, valueType, sourceRow, "IN") Then Exit Function
        End If

        If ValuesMatch(actualValue, "=", candidate, "AUTO", sourceRow) Then
            ValueIsInList = True
            Exit Function
        End If
    Next item
End Function

Private Function TryConvertValue(ByRef value As Variant, ByVal valueType As String, ByVal sourceRow As Long, Optional ByVal context As String = "CONVERSION") As Boolean
    On Error GoTo ConversionError
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Err.Raise 13

    Select Case valueType
        Case "TEXT": value = CStr(value)
        Case "LONG": value = CLng(value)
        Case "DOUBLE": value = CDbl(value)
        Case "CURRENCY": value = CCur(value)
        Case "DATE": value = CDate(value)
        Case "BOOLEAN": value = CBool(value)
        Case Else: TryConvertValue = True: Exit Function
    End Select
    TryConvertValue = True
    Exit Function

ConversionError:
    RecordConversionError sourceRow, context, valueType, value
    Select Case mConversionPolicy
        Case "IGNORE", "NULL"
            TryConvertValue = False
        Case Else
            If sourceRow > 1 Then
                Err.Raise 13, "SqlQuery.ConvertValue", "Cannot convert value in source row " & sourceRow & " to " & valueType & "."
            Else
                Err.Raise 13, "SqlQuery.ConvertValue", "Cannot convert expected value to " & valueType & "."
            End If
    End Select
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
    entry(0) = sourceRow
    entry(1) = context
    entry(2) = valueType
    entry(3) = mConversionPolicy
    entry(4) = valueText
    mConversionErrors.Add entry
End Sub

Private Sub QuickSortRows(ByRef rows() As Long, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long, pivot As Long, temporary As Long
    low = first: high = last: pivot = rows((first + last) \ 2)
    Do While low <= high
        Do While CompareRows(rows(low), pivot) < 0: low = low + 1: Loop
        Do While CompareRows(rows(high), pivot) > 0: high = high - 1: Loop
        If low <= high Then temporary = rows(low): rows(low) = rows(high): rows(high) = temporary: low = low + 1: high = high - 1
    Loop
    If first < high Then QuickSortRows rows, first, high
    If low < last Then QuickSortRows rows, low, last
End Sub

Private Function CompareRows(ByVal leftRow As Long, ByVal rightRow As Long) As Long
    Dim item As Long, comparison As Long
    For item = 1 To mOrderFields.Count
        comparison = CompareOrderValues(GetOrderValue(leftRow, mOrderFields(item)), GetOrderValue(rightRow, mOrderFields(item)), mOrderTypes(item), leftRow, rightRow)
        If comparison <> 0 Then
            If mOrderDirections(item) = "DESC" Then comparison = -comparison
            CompareRows = comparison
            Exit Function
        End If
    Next item
End Function

Private Function GetOrderValue(ByVal sourceRow As Long, ByVal field As Variant) As Variant
    Dim item As Long
    For item = 1 To mComputedAliases.Count
        If LCase$(CStr(mComputedAliases(item))) = LCase$(CStr(field)) Then
            GetOrderValue = EvaluateComputed(mComputedExpressions(item), sourceRow)
            Exit Function
        End If
    Next item
    GetOrderValue = GetSourceValue(sourceRow, field)
End Function

Private Function CompareOrderValues(ByVal leftValue As Variant, ByVal rightValue As Variant, ByVal valueType As String, ByVal leftRow As Long, ByVal rightRow As Long) As Long
    If valueType = "AUTO" Then
        CompareOrderValues = CompareValues(leftValue, rightValue)
        Exit Function
    End If

    Dim leftValid As Boolean
    Dim rightValid As Boolean
    leftValid = TryConvertValue(leftValue, valueType, leftRow)
    rightValid = TryConvertValue(rightValue, valueType, rightRow)

    If Not leftValid And Not rightValid Then Exit Function
    If Not leftValid Then
        CompareOrderValues = 1
        Exit Function
    End If
    If Not rightValid Then
        CompareOrderValues = -1
        Exit Function
    End If

    CompareOrderValues = CompareValues(leftValue, rightValue)
End Function

Private Function CompareValues(ByVal leftValue As Variant, ByVal rightValue As Variant, Optional ByVal valueType As String = "AUTO", Optional ByVal leftRow As Long = 0, Optional ByVal rightRow As Long = 0) As Long
    If valueType <> "AUTO" Then
        leftValue = ConvertValue(leftValue, valueType, leftRow)
        rightValue = ConvertValue(rightValue, valueType, rightRow)
    End If

    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        If CDbl(leftValue) < CDbl(rightValue) Then
            CompareValues = -1
        ElseIf CDbl(leftValue) > CDbl(rightValue) Then
            CompareValues = 1
        End If
    Else
        Dim leftText As String, rightText As String
        leftText = CStr(leftValue): rightText = CStr(rightValue)
        If StrComp(leftText, rightText, vbTextCompare) < 0 Then
            CompareValues = -1
        ElseIf StrComp(leftText, rightText, vbTextCompare) > 0 Then
            CompareValues = 1
        End If
    End If
End Function

Private Function NormalizeValueType(ByVal valueType As String) As String
    valueType = UCase$(Trim$(valueType))
    Select Case valueType
        Case "AUTO", "TEXT", "LONG", "DOUBLE", "CURRENCY", "DATE", "BOOLEAN"
            NormalizeValueType = valueType
        Case Else
            Err.Raise 5, "SqlQuery.NormalizeValueType", "Unsupported value type: " & valueType
    End Select
End Function

Private Function ConvertValue(ByVal value As Variant, ByVal valueType As String, ByVal sourceRow As Long) As Variant
    On Error GoTo ConversionError
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Err.Raise 13

    Select Case valueType
        Case "TEXT": ConvertValue = CStr(value)
        Case "LONG": ConvertValue = CLng(value)
        Case "DOUBLE": ConvertValue = CDbl(value)
        Case "CURRENCY": ConvertValue = CCur(value)
        Case "DATE": ConvertValue = CDate(value)
        Case "BOOLEAN": ConvertValue = CBool(value)
        Case Else: ConvertValue = value
    End Select
    Exit Function

ConversionError:
    If sourceRow > 1 Then
        Err.Raise 13, "SqlQuery.ConvertValue", "Cannot convert value in source row " & sourceRow & " to " & valueType & "."
    Else
        Err.Raise 13, "SqlQuery.ConvertValue", "Cannot convert expected value to " & valueType & "."
    End If
End Function

Private Function GetSourceValue(ByVal sourceRow As Long, ByVal field As Variant) As Variant
    GetSourceValue = mSource(LBound(mSource, 1) + sourceRow - 1, LBound(mSource, 2) + ResolveColumn(field) - 1)
End Function

Private Function ResolveColumn(ByVal field As Variant) As Long
    If IsNumeric(field) Then
        ResolveColumn = CLng(field)
        If ResolveColumn < 1 Or ResolveColumn > mColumnCount Then Err.Raise 5, "SqlQuery.ResolveColumn", "Column index is out of range."
    ElseIf mHeaderIndex.Exists(CStr(field)) Then
        ResolveColumn = mHeaderIndex(CStr(field))
    Else
        Err.Raise 5, "SqlQuery.ResolveColumn", "Unknown column: " & CStr(field)
    End If
End Function

Private Function ResolveOrderColumn(ByVal field As Variant) As Long
    On Error GoTo AliasLookup
    ResolveOrderColumn = ResolveColumn(field)
    Exit Function
AliasLookup:
    Err.Clear
    Dim item As Long
    For item = 1 To mAggregateAliases.Count
        If LCase$(CStr(mAggregateAliases(item))) = LCase$(CStr(field)) Then
            ResolveOrderColumn = mGroupFields.Count + item
            Exit Function
        End If
    Next item
    For item = 1 To mComputedAliases.Count
        If LCase$(CStr(mComputedAliases(item))) = LCase$(CStr(field)) Then
            ResolveOrderColumn = mColumnCount + item
            Exit Function
        End If
    Next item
    Err.Raise 5, "SqlQuery.ResolveOrderColumn", "Unknown ORDER BY column: " & CStr(field)
End Function

Private Sub ParseSelectItem(ByVal itemText As String, ByRef fieldName As String, ByRef aliasName As String)
    Dim tokens() As String
    Dim tokenCount As Long
    tokens = Split(NormalizeWhitespace(itemText))
    tokenCount = UBound(tokens) - LBound(tokens) + 1
    If tokenCount = 1 Then
        fieldName = Trim$(tokens(0))
        aliasName = fieldName
    ElseIf tokenCount = 3 And UCase$(Trim$(tokens(1))) = "AS" Then
        fieldName = Trim$(tokens(0))
        aliasName = Trim$(tokens(2))
    Else
        Err.Raise 5, "SqlQuery.ParseSelectItem", "Expected 'Field' or 'Field AS Alias'."
    End If
    If Len(fieldName) = 0 Or Len(aliasName) = 0 Then Err.Raise 5, "SqlQuery.ParseSelectItem", "Invalid SELECT column specification."
End Sub

Private Function NormalizeWhitespace(ByVal value As String) As String
    value = Replace(value, vbTab, " ")
    value = Trim$(value)
    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop
    NormalizeWhitespace = value
End Function

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
