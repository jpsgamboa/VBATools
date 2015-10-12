Attribute VB_Name = "WorksheetHandling"
Option Explicit

'Returns the named range
Public Function getNamedRange(ByVal rangeName As String, ws As Worksheet) As Range
    Set getNamedRange = ws.Range(rangeName)
End Function

'Returns the number of cells in a range
Public Function getRangeCount(rng As Range) As Integer
    getRangeCount = WorksheetFunction.CountA(rng)
End Function

'Returns True if the range contains the specified String
Private Function rangeContains(rng As Range, ByVal str As String) As Boolean
    Dim c As Variant
    For Each c In rng.Cells
        If strEquals(c.value, str) Then rangeContains = True
    Next
End Function
