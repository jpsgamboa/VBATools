Attribute VB_Name = "DataBaseFunctions"
Option Explicit

'A set of functions usefull to handle data stored in a database format in a worksheet.
'The table fields must be on line 1, and the ID should be stored on column A

'Returns a range holding every named column
Public Function getColumnsRange(ws As Worksheet) As Range
    Set getColumnsRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Range("A1").End(xlToRight).column))
End Function

'Returns a range with will all the rows containing a value
Public Function getRowsRange(ws As Worksheet) As Range
    Set getRowsRange = ws.Range(ws.Cells(1, 1), ws.Cells(ws.Range("A" & rows.Count).End(xlUp).row, 1))
End Function

'Returns a range containing the cell with the first empty column
Public Function findEmptyColumn(ws As Worksheet) As Range
    Dim lastColumn As Integer: lastColumn = ws.Cells(1, columns.Count).End(xlToLeft).column + 1
    If Len(ws.Cells(1, columns.Count).End(xlToLeft)) = 0 Then lastColumn = lastColumn - 1
    Set findEmptyColumn = ws.Range(ws.Cells(1, lastColumn), ws.Cells(1, lastColumn))
End Function

'Returns a range containing the cell with the first empty row
Public Function findEmptyRow(ws As Worksheet) As Range
    Dim lastRow As Integer: lastRow = ws.Cells(ws.Range("A" & rows.Count).End(xlUp).row, 1).row
    Set findEmptyRow = ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(lastRow + 1, 1))
End Function

'Returns a range with the cell where the field name is written
Public Function getColumnByName(name As String, ws As Worksheet) As Range
Dim c As Variant
For Each c In getColumnsRange(ws).Cells
    If strEquals(c.value, name) Then
        Set getColumnByName = ws.Range(c, c): Exit For
    Else
        Set getColumnByName = Nothing
    End If
Next
End Function

'Returns true if a column exists in a database
Public Function columnExists(columnName As String, targetWS As Worksheet) As Boolean
    Dim col As Variant: Set col = getColumnByName(columnName, targetWS)
    If Not col Is Nothing Then columnExists = True
End Function

'Convenience function to return a value from a databse
Public Function getValueAt(cRow As Integer, cColumn As String, ws As Worksheet) As Variant
    getValueAt = ws.Cells(cRow, getColumnByName(cColumn, ws).column)
End Function

'Returns true if an entry exists in a database
Public Function containsByID(ID As Variant, idColumnName As String, ws As Worksheet) As Boolean
    Dim targetColumn As Variant: Set targetColumn = getColumnByName(idColumnName, ws)(1)
    Dim r As Variant
    For Each r In getRowsRange(ws).Cells
        If strEquals(ID, ws.Cells(r.row, targetColumn.column)) Then
            containsByID = True
            Exit For
        End If
    Next
End Function

'Returns a range with the cell of the matching ID
Public Function getRowByID(ID As Variant, idColumnName As String, ws As Worksheet) As Variant
    If containsByID(ID, idColumnName, ws) Then
        Dim targetColumn As Variant: Set targetColumn = getColumnByName(idColumnName, ws)(1)
        Dim r As Variant
        For Each r In getRowsRange(ws).Cells
            If strEquals(ID, ws.Cells(r.row, targetColumn.column)) Then
                Set getRowByID = ws.Cells(r.row, targetColumn.column)
                Exit For
            End If
        Next
    End If
End Function

'Returns the value of the latest name in the given column
Public Function getLatestRecord(dateColumnName As String, ws As Worksheet) As Variant
    Dim rs As Range: Set rs = getRowsRange(ws)
    Dim col As Integer: col = getColumnByName(dateColumnName, ws).column
    
    Dim latestDate As Variant, r As Variant: latestDate = 1
    
    For Each r In rs.Cells
        If r.row > 1 Then
            Dim a As Variant: a = ws.Cells(r.row, col).value
            Dim b As Variant: b = a - latestDate
        End If
        If r.row > 1 Then If ws.Cells(r.row, col).value > latestDate Then latestDate = ws.Cells(r.row, col).value
    Next
    getLatestRecord = latestDate
End Function

'Procura, na sheet destino, a data de exportação mais recente
Private Function getLatestExportDate(ws As Worksheet) As Variant
    Dim rs As Range: Set rs = getRowsRange(ws)
    Dim col As Integer: col = getColumnByName("HoraExport", ws).column
    
    Dim latestDate As Variant, r As Variant: latestDate = 1
    
    For Each r In rs.Cells
        If r.row > 1 Then
            Dim a As Variant: a = ws.Cells(r.row, col).value
            Dim b As Variant: b = a - latestDate
        End If
        If r.row > 1 Then If ws.Cells(r.row, col).value > latestDate Then latestDate = ws.Cells(r.row, col).value
    Next
    getLatestExportDate = latestDate
End Function

'Verifica se a coluna ja existe e, caso contrário, adiciona-a
Private Sub addNewColumn(title As String, ws As Worksheet)
    If columnExists(title, ws) Then 'if this column already exists, skip it
        GoTo DoNothing:
    Else
        findEmptyColumn(ws)(1).value = title
    End If
DoNothing:
End Sub
