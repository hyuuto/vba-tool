Function Number2Digit(ByVal ColumnNumber As Integer) As String
    'Convert To Column Letter
    Number2Digit = Split(Cells(1, ColumnNumber).Address, "$")(1)
End Function



Public Function RemoveWhiteSpace(target As String) As String
    With New RegExp
        .Pattern = "\s"
        .MultiLine = True
        .Global = True
        RemoveWhiteSpace = .Replace(target, vbNullString)
    End With
End Function



Function getFileName(filePath As String)
    Dim FSO As New FileSystemObject
    getFileName = Split(FSO.getFileName(filePath), ".")(0)
End Function



Function DeleteQueries()
    If ThisWorkbook.Queries.Count <> 0 Then
        For i = ThisWorkbook.Queries.Count To 1 Step -1
            ThisWorkbook.Queries(i).Delete
        Next
    End If
End Function



Sub DeleteConnections()
    For i = ThisWorkbook.Connections.Count To 1 Step -1
        ThisWorkbook.Connections(i).Delete
    Next
End Sub



Sub DeleteEmptyRows(Sheet As Excel.Worksheet)
    Dim Row As Range
    Dim Index As Long
    Dim Count As Long

    If Sheet Is Nothing Then Exit Sub

    ' We are iterating across a collection where we delete elements on the way.
    ' So its safe to iterate from the end to the beginning to avoid index confusion.
    For Index = Sheet.UsedRange.Rows.Count To 1 Step -1
        Set Row = Sheet.UsedRange.Rows(Index)
        ' This construct is necessary because SpecialCells(xlCellTypeBlanks)
        ' always throws runtime errors if it doesn't find any empty cell.
        Count = 0
        On Error Resume Next
        Count = Row.SpecialCells(xlCellTypeBlanks).Count
        On Error GoTo 0

        If Count = Row.Cells.Count Then Row.Delete xlUp
  Next
End Sub




