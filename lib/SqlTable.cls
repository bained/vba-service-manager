VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Lightweight table-access facade over the SqlQuery in-memory engine.
' Public row indexes are 1-based data-row indexes; the header row is not counted.

Private mSource As Variant
Private mHeaders() As String
Private mHeaderIndex As Object
Private mRowCount As Long
Private mColumnCount As Long
Private mIsOpen As Boolean

Public Function OpenRange(ByVal sourceRange As Range) As SqlTable
    On Error GoTo ErrorHandler

    If sourceRange Is Nothing Then
        Err.Raise 5, "SqlTable.OpenRange", "The source range is Nothing."
    End If

    Dim sourceArray As Variant
    sourceArray = sourceRange.Value2
    Set OpenRange = OpenArray(sourceArray)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.OpenRange", Err.Description
End Function

Public Function AutoRange(ByVal sheetName As String) As SqlTable
    On Error GoTo ErrorHandler

    If Len(Trim$(sheetName)) = 0 Then
        Err.Raise 5, "SqlTable.AutoRange", "The worksheet name cannot be empty."
    End If

    Dim sourceSheet As Worksheet
    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo ErrorHandler
    If sourceSheet Is Nothing Then
        Err.Raise 9, "SqlTable.AutoRange", "Worksheet not found: " & sheetName
    End If

    Dim lastCellByRow As Range
    Dim lastCellByColumn As Range
    Set lastCellByRow = sourceSheet.Cells.Find(What:="*", _
                                               After:=sourceSheet.Cells(1, 1), _
                                               LookIn:=xlFormulas, _
                                               LookAt:=xlPart, _
                                               SearchOrder:=xlByRows, _
                                               SearchDirection:=xlPrevious, _
                                               MatchCase:=False)
    Set lastCellByColumn = sourceSheet.Cells.Find(What:="*", _
                                                  After:=sourceSheet.Cells(1, 1), _
                                                  LookIn:=xlFormulas, _
                                                  LookAt:=xlPart, _
                                                  SearchOrder:=xlByColumns, _
                                                  SearchDirection:=xlPrevious, _
                                                  MatchCase:=False)

    If lastCellByRow Is Nothing Or lastCellByColumn Is Nothing Then
        Err.Raise 5, "SqlTable.AutoRange", "Worksheet contains no data: " & sheetName
    End If
    If lastCellByRow.Row < 1 Or lastCellByColumn.column < 1 Then
        Err.Raise 5, "SqlTable.AutoRange", "Worksheet contains no usable table: " & sheetName
    End If

    Dim automaticRange As Range
    Set automaticRange = sourceSheet.Range(sourceSheet.Cells(1, 1), _
                                           sourceSheet.Cells(lastCellByRow.Row, lastCellByColumn.column))

    If IsError(automaticRange.Cells(1, 1).Value2) Or IsNull(automaticRange.Cells(1, 1).Value2) Or IsEmpty(automaticRange.Cells(1, 1).Value2) Then
        Err.Raise 5, "SqlTable.AutoRange", "The automatic range must have a header row beginning at A1."
    End If
    If Len(Trim$(CStr(automaticRange.Cells(1, 1).Value2))) = 0 Then
        Err.Raise 5, "SqlTable.AutoRange", "The automatic range must have a header row beginning at A1."
    End If

    Dim headerColumn As Long
    For headerColumn = 1 To automaticRange.columns.Count
        If IsError(automaticRange.Cells(1, headerColumn).Value2) Or IsNull(automaticRange.Cells(1, headerColumn).Value2) Or IsEmpty(automaticRange.Cells(1, headerColumn).Value2) Then
            Err.Raise 5, "SqlTable.AutoRange", "Header row contains an invalid title at column " & headerColumn & "."
        End If
        If Len(Trim$(CStr(automaticRange.Cells(1, headerColumn).Value2))) = 0 Then
            Err.Raise 5, "SqlTable.AutoRange", "Header row contains an empty title at column " & headerColumn & "."
        End If
    Next headerColumn

    Set AutoRange = OpenRange(automaticRange)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.AutoRange", Err.Description
End Function

Public Function OpenArray(ByRef sourceArray As Variant) As SqlTable
    On Error GoTo ErrorHandler

    If Not IsArray(sourceArray) Then
        Err.Raise 5, "SqlTable.OpenArray", "The source must be a 2D array."
    End If

    If GetArrayDimensionCount(sourceArray) <> 2 Then
        Err.Raise 5, "SqlTable.OpenArray", "The source must be a 2D array."
    End If

    Dim sourceRows As Long
    Dim sourceColumns As Long
    sourceRows = UBound(sourceArray, 1) - LBound(sourceArray, 1) + 1
    sourceColumns = UBound(sourceArray, 2) - LBound(sourceArray, 2) + 1

    If sourceRows < 1 Or sourceColumns < 1 Then
        Err.Raise 5, "SqlTable.OpenArray", "The source array is empty."
    End If

    mSource = sourceArray
    mRowCount = sourceRows - 1
    mColumnCount = sourceColumns

    ReDim mHeaders(1 To mColumnCount)
    Set mHeaderIndex = CreateObject("Scripting.Dictionary")
    mHeaderIndex.CompareMode = vbTextCompare

    Dim columnNumber As Long
    Dim headerText As String
    For columnNumber = 1 To mColumnCount
        headerText = CStr(mSource(LBound(mSource, 1), _
                                 LBound(mSource, 2) + columnNumber - 1))
        mHeaders(columnNumber) = headerText

        If Len(Trim$(headerText)) > 0 Then
            If Not mHeaderIndex.Exists(headerText) Then
                mHeaderIndex.Add headerText, columnNumber
            End If
        End If
    Next columnNumber

    mIsOpen = True
    Set OpenArray = Me
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.OpenArray", Err.Description
End Function

Public Property Get HeaderCount() As Long
    On Error GoTo ErrorHandler
    EnsureOpen "HeaderCount"
    HeaderCount = mColumnCount
    Exit Property

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.HeaderCount", Err.Description
End Property

Public Property Get RowCount() As Long
    On Error GoTo ErrorHandler
    EnsureOpen "RowCount"
    RowCount = mRowCount
    Exit Property

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.RowCount", Err.Description
End Property

Public Property Get ColumnCount() As Long
    On Error GoTo ErrorHandler
    EnsureOpen "ColumnCount"
    ColumnCount = mColumnCount
    Exit Property

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.ColumnCount", Err.Description
End Property

Public Function HasColumn(ByVal column As Variant) As Boolean
    On Error GoTo ErrorHandler
    EnsureOpen "HasColumn"

    If IsNumeric(column) Then
        HasColumn = CLng(column) >= 1 And CLng(column) <= mColumnCount
    Else
        HasColumn = mHeaderIndex.Exists(CStr(column))
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.HasColumn", Err.Description
End Function

Public Function GetCell(ByVal rowIndex As Long, ByVal column As Variant) As Variant
    On Error GoTo ErrorHandler
    EnsureOpen "GetCell"

    Dim columnIndex As Long
    columnIndex = ResolveColumn(column, "GetCell")
    ValidateDataRow rowIndex, "GetCell"
    GetCell = mSource(LBound(mSource, 1) + rowIndex, _
                      LBound(mSource, 2) + columnIndex - 1)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.GetCell", Err.Description
End Function

Public Function GetRowArray(ByVal rowIndex As Long) As Variant
    On Error GoTo ErrorHandler
    EnsureOpen "GetRowArray"
    ValidateDataRow rowIndex, "GetRowArray"

    Dim result() As Variant
    ReDim result(1 To mColumnCount)

    Dim columnNumber As Long
    For columnNumber = 1 To mColumnCount
        result(columnNumber) = mSource(LBound(mSource, 1) + rowIndex, _
                                       LBound(mSource, 2) + columnNumber - 1)
    Next columnNumber

    GetRowArray = result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.GetRowArray", Err.Description
End Function

Public Function GetRowDictionary(ByVal rowIndex As Long) As Object
    On Error GoTo ErrorHandler
    EnsureOpen "GetRowDictionary"
    ValidateDataRow rowIndex, "GetRowDictionary"

    Dim result As Object
    Set result = CreateRowDictionary(rowIndex)
    Set GetRowDictionary = result
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.GetRowDictionary", Err.Description
End Function

Public Function GetColumn(ByVal column As Variant) As Variant
    On Error GoTo ErrorHandler
    EnsureOpen "GetColumn"

    Dim columnIndex As Long
    columnIndex = ResolveColumn(column, "GetColumn")

    Dim result() As Variant
    If mRowCount > 0 Then ReDim result(1 To mRowCount)

    Dim rowIndex As Long
    For rowIndex = 1 To mRowCount
        result(rowIndex) = mSource(LBound(mSource, 1) + rowIndex, _
                                   LBound(mSource, 2) + columnIndex - 1)
    Next rowIndex

    If mRowCount = 0 Then
        GetColumn = Array()
    Else
        GetColumn = result
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.GetColumn", Err.Description
End Function

Public Function GetAll() As Variant
    On Error GoTo ErrorHandler
    EnsureOpen "GetAll"
    GetAll = mSource
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.GetAll", Err.Description
End Function

Public Function query() As SqlQuery
    On Error GoTo ErrorHandler
    EnsureOpen "Query"

    Dim queryObject As SqlQuery
    Set queryObject = New SqlQuery
    Set queryObject = queryObject.LoadArray(mSource)
    Set query = queryObject
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.Query", Err.Description
End Function

Public Function FindFirst(ByVal field As Variant, _
                          ByVal comparisonOperator As String, _
                          ByVal expectedValue As Variant) As Object
    On Error GoTo ErrorHandler

    Dim matches As Collection
    Set matches = FindAll(field, comparisonOperator, expectedValue)

    If matches.Count = 0 Then
        Set FindFirst = Nothing
    Else
        Set FindFirst = matches(1)
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.FindFirst", Err.Description
End Function

Public Function FindAll(ByVal field As Variant, _
                        ByVal comparisonOperator As String, _
                        ByVal expectedValue As Variant) As Collection
    On Error GoTo ErrorHandler

    Dim queryObject As SqlQuery
    Dim result As Variant
    Dim matches As Collection
    Set matches = New Collection

    Set queryObject = query()
    Set queryObject = queryObject.Where(field, comparisonOperator, expectedValue)
    result = queryObject.ToArray()

    Dim resultRow As Long
    Dim resultColumn As Long
    Dim rowDictionary As Object
    Dim sourceRowIndex As Long

    For resultRow = 2 To UBound(result, 1)
        Set rowDictionary = CreateObject("Scripting.Dictionary")
        rowDictionary.CompareMode = vbTextCompare

        For resultColumn = 1 To mColumnCount
            If Not rowDictionary.Exists(mHeaders(resultColumn)) Then
                rowDictionary.Add mHeaders(resultColumn), result(resultRow, resultColumn)
            End If
        Next resultColumn

        matches.Add rowDictionary
    Next resultRow

    Set FindAll = matches
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.FindAll", Err.Description
End Function

Public Function Count() As Long
    On Error GoTo ErrorHandler
    EnsureOpen "Count"
    Count = mRowCount
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.Count", Err.Description
End Function

Public Function Exists(ByVal field As Variant, _
                       ByVal comparisonOperator As String, _
                       ByVal expectedValue As Variant) As Boolean
    On Error GoTo ErrorHandler

    Dim matches As Collection
    Set matches = FindAll(field, comparisonOperator, expectedValue)
    Exists = matches.Count > 0
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "SqlTable.Exists", Err.Description
End Function

Private Function CreateRowDictionary(ByVal rowIndex As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim columnNumber As Long
    For columnNumber = 1 To mColumnCount
        If Not result.Exists(mHeaders(columnNumber)) Then
            result.Add mHeaders(columnNumber), _
                       mSource(LBound(mSource, 1) + rowIndex, _
                               LBound(mSource, 2) + columnNumber - 1)
        End If
    Next columnNumber

    Set CreateRowDictionary = result
End Function

Private Function ResolveColumn(ByVal column As Variant, ByVal context As String) As Long
    If IsNumeric(column) Then
        ResolveColumn = CLng(column)
        If ResolveColumn < 1 Or ResolveColumn > mColumnCount Then
            Err.Raise 5, "SqlTable." & context, "Column index is out of range: " & CStr(column)
        End If
    Else
        If Not mHeaderIndex.Exists(CStr(column)) Then
            Err.Raise 5, "SqlTable." & context, "Unknown column: " & CStr(column)
        End If
        ResolveColumn = CLng(mHeaderIndex(CStr(column)))
    End If
End Function

Private Sub ValidateDataRow(ByVal rowIndex As Long, ByVal context As String)
    If rowIndex < 1 Or rowIndex > mRowCount Then
        Err.Raise 5, "SqlTable." & context, _
                  "Data row index is out of range: " & CStr(rowIndex)
    End If
End Sub

Private Sub EnsureOpen(ByVal context As String)
    If Not mIsOpen Then
        Err.Raise 91, "SqlTable." & context, "The table has not been opened."
    End If
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
