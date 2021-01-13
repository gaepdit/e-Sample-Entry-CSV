''' Table functions
Function TableIsEmpty(tableName As String) As Boolean
    TableIsEmpty = True
    If WorksheetFunction.CountA(Range(tableName)) Then TableIsEmpty = False
End Function

Function CellValue(tbl As ListObject, row As Range, columnName As String) As String
    CellValue = row.Cells(1, tbl.ListColumns(columnName).Index)
End Function

Function CellDateValue(tbl As ListObject, row As Range, columnName As String) As String
    CellDateValue = Format(row.Cells(1, tbl.ListColumns(columnName).Index), "yyyy-mm-dd")
End Function

Function CellTimeValue(tbl As ListObject, row As Range, columnName As String) As String
    CellTimeValue = Format(row.Cells(1, tbl.ListColumns(columnName).Index), "hhnn")
End Function
