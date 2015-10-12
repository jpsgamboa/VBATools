Attribute VB_Name = "DataFunctions"
Option Explicit

'Returns true if two strings are equal
Public Function strEquals(str1, str2) As Boolean
    If StrComp(str1, str2) = 0 Then strEquals = True
End Function

'Converts from epoch time to hour
Public Function epochToHour(value As Variant) As Variant
   ' On Error GoTo handle
    ''Miliseconds to seconds
    Dim mil2sec As Variant
    mil2sec = value * 0.001
    ''Epoch to date
    Dim d As Variant
    d = DateAdd("s", mil2sec, "01/01/1970 00:00:00")
    ''Format to hour
    epochToHour = Format(d, "hh:nn:ss")
'handle:
   ' epochToHour = "Error"
End Function


'Converts from epoch time to hour
Public Function epochToDate(value As Variant) As Variant
   ' On Error GoTo handle
    ''Miliseconds to seconds
    Dim mil2sec As Variant
    mil2sec = value * 0.001
    ''Epoch to date
    Dim d As Variant
    d = DateAdd("s", mil2sec, "01/01/1970 00:00:00")
    ''Format to date
    epochToDate = Format(d, "dd-mm-yy")
'handle:
   ' epochToDate = "Error"
End Function

'Formats as date
Public Function formatAsDate(value As Variant) As Variant
    On Error GoTo handle
    formatAsDate = DateSerial(year(value), month(value), day(value))
handle:
    formatAsDate = "Error"
End Function

'Returns a string with the type of data in the cell
Function CellType(rng) As String
    Set rng = rng.Range("A1")
    Select Case True
        Case IsEmpty(rng)
            CellType = "Blank"
        Case WorksheetFunction.IsText(rng)
            CellType = "Text"
        Case WorksheetFunction.IsLogical(rng)
            CellType = "Logical"
        Case WorksheetFunction.IsErr(rng)
            CellType = "Error"
        Case IsDate(rng)
            CellType = "Date"
        Case InStr(1, rng.text, ":") <> 0
            CellType = "Time"
        Case IsNumeric(rng)
            CellType = "Value"
    End Select
End Function

'Replaces the common portuguese special chars with the non-special equivalent
Public Function replaceSpecialChars(str As String) As String
    Dim A_chars() As Variant: A_chars = Array("A", "Á", "Ã", "Â", "À", "á", "ã", "â", "à")
    Dim E_chars() As Variant: E_chars = Array("E", "É", "È", "Ê", "é", "è", "ê")
    Dim I_chars() As Variant: I_chars = Array("I", "Í", "Í", "í", "ì")
    Dim O_chars() As Variant: O_chars = Array("O", "Ó", "Õ", "Ô", "ó", "õ", "ô")
    Dim U_chars() As Variant: U_chars = Array("U", "Ú", "Ù", "ú", "ù")
    Dim C_chars() As Variant: C_chars = Array("C", "Ç", "ç")

    Dim charArrays() As Variant: charArrays = Array(A_chars, E_chars, I_chars, O_chars, U_chars, C_chars)
        
    Dim arr As Variant, chr As Variant, iter As Integer, currChr As String, newStr As String: newStr = ""
    
    For iter = 1 To Len(str) 'Loop through each char in the provided string
        currChr = Mid(str, iter, 1) 'String indexes start at 1 in VBA
        For Each arr In charArrays 'For each of the special char arrays declared above
            Dim counter As Integer: counter = 0
            For Each chr In arr 'Loop through each letter (besides the first)
                If counter > 0 Then _
                    If strEquals(chr, currChr) Then currChr = arr(0) 'Check if they match
                counter = counter + 1
            Next chr
        Next arr
        newStr = newStr & currChr 'Add either the replacement or the old letter to a new string
    Next iter
    replaceSpecialChars = newStr
End Function
