Attribute VB_Name = "MiscFunctions"
'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 2/3/2014 6:52:06 PM : from manifest:8767201 gist https://gist.github.com/brucemcpherson/3414346/raw
Option Explicit
' v2.20  3414346

 ' Acknowledgement for the microtimer procedures used here to
 ' thanks to Charles Wheeler - http://www.decisionmodels.com/
 ' ---
 #If VBA7 And Win64 Then
 Private Declare PtrSafe Function getTickCount _
     Lib "kernel32" Alias "QueryPerformanceCounter" _
     (cyTickCount As Currency) As Long
 Private Declare PtrSafe Function _
     GetDeviceCaps Lib "Gdi32" _
     (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
 #Else
 Private Declare Function getTickCount Lib "kernel32" _
 Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long


 Private Declare Function getFrequency Lib "kernel32" _
 Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
 #End If


 '----

 #If VBA7 And Win64 Then

    
     Private Declare PtrSafe Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" ( _
   ByVal hwnd As Long, _
   ByVal Operation As String, _
   ByVal fileName As String, _
   Optional ByVal Parameters As String, _
   Optional ByVal Directory As String, _
   Optional ByVal WindowStyle As Long = vbMaximizedFocus _
   ) As LongLong
  
   Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
     ByVal CodePage As LongLong, ByVal dwflags As LongLong, _
     ByVal lpWideCharStr As LongLong, ByVal cchWideChar As LongLong, _
     ByVal lpMultiByteStr As LongLong, ByVal cchMultiByte As LongLong, _
     ByVal lpDefaultChar As LongLong, ByVal lpUsedDefaultChar As LongLong) As LongLong
    
    
 #Else

 Private Declare Function ShellExecute _
   Lib "shell32.dll" Alias "ShellExecuteA" ( _
   ByVal hwnd As Long, _
   ByVal Operation As String, _
   ByVal fileName As String, _
   Optional ByVal Parameters As String, _
   Optional ByVal Directory As String, _
   Optional ByVal WindowStyle As Long = vbMaximizedFocus _
   ) As Long
  
 Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
     ByVal CodePage As Long, ByVal dwflags As Long, _
     ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
     ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, _
     ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    
 #End If

 ' note original execute shell stuff came from this post
 ' http://stackoverflow.com/questions/3166265/open-an-html-page-in-default-browser-with-vba
 ' thanks to http://stackoverflow.com/users/174718/dmr

Private Const CP_UTF8 = 65001
Public Const cFailedtoGetHandle = -1
Public Function nameExists(s As String) As name
    On Error GoTo handle
    Set nameExists = ActiveWorkbook.names(s)
    Exit Function
handle:
    Set nameExists = Nothing
End Function
Public Function whereIsThis(r As Variant) As Range
    Dim n As name
    
    If TypeName(r) = "range" Then
        Set whereIsThis = r
    Else
        Set n = nameExists(CStr(r))
        If Not n Is Nothing Then
            Set whereIsThis = n.RefersToRange
        Else
            Set whereIsThis = Range(r)
        End If
    End If
            
        
End Function
Public Function OpenUrl(url) As Boolean
    #If VBA7 And Win64 Then
    Dim lSuccess As LongLong
    #Else
    Dim lSuccess As Long
    #End If
    lSuccess = ShellExecute(0, "Open", url)
    OpenUrl = lSuccess > 32
End Function

Function firstCell(inrange As Range) As Range
    Set firstCell = inrange.Cells(1, 1)
End Function
Function lastCell(inrange As Range) As Range
    Set lastCell = inrange.Cells(inrange.rows.Count, inrange.columns.Count)
End Function
Function isSheet(o As Object) As Boolean
     Dim r As Range
     On Error GoTo handleError
        Set r = o.Cells
        isSheet = True
        Exit Function

handleError:
    isSheet = False
End Function
Public Function findShape(sName As String, Optional ws As Worksheet = Nothing) As shape
    Dim s As shape, t As shape
    If ws Is Nothing Then Set ws = ActiveSheet
    For Each s In ws.Shapes
        If makeKey(s.name) = makeKey(sName) Then
            Set t = s
            Exit For
        End If
        If s.Type = msoGroup Then
            Set t = findRecurse(sName, s.GroupItems)
            If Not t Is Nothing Then
                Exit For
            End If
        End If
    Next s
    Set findShape = t
    
End Function
Public Function findRecurse(target As String, co As GroupShapes) As shape
    Dim s As shape, t As shape
    ' only works one level down.. cant get .gtoupitems to work properly
    For Each s In co
        If makeKey(s.name) = makeKey(target) Then
            Set t = s
            Exit For
        End If
    Next s
    Set findRecurse = t
End Function
Public Sub clearHyperLinks(ws As Worksheet)
' delete all the hyperlinks on a sheet
    With ws
        While .Hyperlinks.Count > 0
           .Hyperlinks(1).Delete
        Wend
    End With
End Sub
Function sheetExists(sName As String, Optional complain As Boolean = True) As Worksheet
    
    On Error GoTo handleError
        Set sheetExists = Sheets(sName)
        Exit Function

handleError:
    If complain Then MsgBox ("Could not open sheet " & sName)
    Set sheetExists = Nothing

End Function
Function wholeSheet(wn As String) As Range
    ' return a range representing the entire used worksheet
    Set wholeSheet = wholeWs(sheetExists(wn))
End Function
Function wholeWs(ws As Worksheet) As Range
    Set wholeWs = ws.UsedRange
End Function
Function wholeRange(r As Range) As Range
    Set wholeRange = wholeWs(r.Worksheet)
End Function
Function cleanFind(x As Variant, r As Range, Optional complain As Boolean = False, _
        Optional singlecell As Boolean = False) As Range
    ' does a normal .find, but catches where range is nothing
    Dim u As Range
    Set u = Nothing

    If r Is Nothing Then
        Set u = Nothing
    Else
        Set u = r.find(x, , xlValues, xlWhole)
    End If
    
    If singlecell And Not u Is Nothing Then
        Set u = firstCell(u)
    End If
 
    If complain And u Is Nothing Then
        Call msglost(x, r)
    End If
    
    Set cleanFind = u
    
End Function
Sub msglost(x As Variant, r As Range, Optional extra As String = "")

    MsgBox ("Couldnt find " & CStr(x) & " in " & SAd(r) & " " & extra)

End Sub
Function SAd(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, _
        Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String
    Dim strA As String
    Dim r As Range
    Dim u As Range
    
    ' creates an address including the worksheet name
    strA = ""
    For Each r In rngIn.Areas
        Set u = r
        If singlecell Then
            Set u = firstCell(u)
        End If
        strA = strA + SAdOneRange(u, target, singlecell, removeRowDollar, removeColDollar) & ","
    Next r
    SAd = Left(strA, Len(strA) - 1)
End Function
Function SAdOneRange(rngIn As Range, Optional target As Range = Nothing, Optional singlecell As Boolean = False, _
                        Optional removeRowDollar As Boolean = False, Optional removeColDollar As Boolean = False) As String
    Dim strA As String
    
    ' creates an address including the worksheet name
    
    strA = AddressNoDollars(rngIn, removeRowDollar, removeColDollar)
    
    ' dont bother with worksheet name if its on the same sheet, and its been asked to do that
    
    If Not target Is Nothing Then
        If target.Worksheet Is rngIn.Worksheet Then
            SAdOneRange = strA
            Exit Function
        End If
    End If

    ' otherwise add the sheet name
    
    SAdOneRange = "'" & rngIn.Worksheet.name & "'!" & strA
        
End Function
Function AddressNoDollars(a As Range, Optional doRow As Boolean = True, Optional doColumn As Boolean = True) As String
' return address minus the dollars
    Dim st As String
    Dim p1 As Long, p2 As Long
    AddressNoDollars = a.Address
    
    If doRow And doColumn Then
        AddressNoDollars = Replace(a.Address, "$", "")
    Else
        p1 = InStr(1, a.Address, "$")
        p2 = 0
        If p1 > 0 Then
            p2 = InStr(p1 + 1, a.Address, "$")
        End If
        ' turn $A$1 into A$1
        If doColumn And p1 > 0 Then
            AddressNoDollars = Left(a.Address, p1 - 1) & Mid(a.Address, p1 + 1)
        
        ' turn $a$1 into $a1
        ElseIf doRow And p2 > 0 Then
            AddressNoDollars = Left(a.Address, p2 - 1) & Mid(a.Address, p2 + 1, p2 - p1)
    
        End If
    End If
    
    
End Function
Function isReallyEmpty(r As Range) As Boolean
    Dim b As Boolean
    b = (Application.CountBlank(r) = r.Cells.Count)

    isReallyEmpty = b
End Function
Function toEmptyRow(r As Range) As Range
    Dim o As Range, u As Range, w As Long
    ' returns to first blank row
    Set u = wholeRange(r)
    Set o = r
    w = lastCell(u).row + 1
    Do While True
        ' whats left in the sheet
        Set o = cleanFind(Empty, o.Resize(w, 1), True, True)
        If isReallyEmpty(o.Resize(1, r.columns.Count)) Then
            Exit Do
        Else
            Set o = o.Offset(1)
        End If
    Loop

    If (o.row > lastCell(r).row And r.rows.Count > 1) Then
        Set toEmptyRow = r
    Else
        If o.row > r.row Then
            Set toEmptyRow = r.Resize(o.row - r.row)
        Else
            MsgBox ("nothing on sheet")
            Set toEmptyRow = Nothing
        End If
    End If
    
End Function
Function toEmptyCol(r As Range) As Range

    Dim o As Range, u As Range, w As Long
    ' returns to first blank column
    Set u = wholeRange(r)
    Set o = r
    w = lastCell(u).column + 1
    Do While True
        Set o = cleanFind(Empty, o.Resize(1, w), True, True)
        If isReallyEmpty(toEmptyRow(o)) Then
            Exit Do
        Else
            Set o = o.Offset(, 1)
        End If
    Loop
    If (o.column > r.column) Then
        Set toEmptyCol = r.Resize(r.rows.Count, o.column - r.column)
    End If
End Function
Function toEmptyBox(r As Range) As Range
    Set toEmptyBox = toEmptyCol(toEmptyRow(r))
End Function
Public Function getLikelyColumnRange(Optional ws As Worksheet = Nothing) As Range
    ' figure out the likely default value for the refedit.
    Dim rstart As Range
    If ws Is Nothing Then
        Set rstart = wholeSheet(ActiveSheet.name)
    Else
        Set rstart = wholeSheet(ws.name)
    End If

    Set getLikelyColumnRange = toEmptyBox(rstart)
    
End Function
Sub deleteAllFromCollection(co As Collection)
    Dim o As Object, i As Long
    For i = co.Count To 1 Step -1
        co(i).Delete
    Next i
    
End Sub
Sub deleteAllShapes(r As Range, startingwith As String)
   
    Dim l As Long
    With r.Worksheet
        For l = .Shapes.Count To 1 Step -1
            If Left(.Shapes(l).name, Len(startingwith)) = startingwith Then
                .Shapes(l).Delete
            End If
        Next l
    End With
    
End Sub
Function makearangeofShapes(r As Range, startingwith As String) As shapeRange
   
    Dim s As shape
    
    Dim n() As String, sz As Long
    With r.Worksheet
        For Each s In .Shapes
            If Left(s.name, Len(startingwith)) = startingwith Then
                sz = sz + 1
                ReDim Preserve n(1 To sz) As String
                n(sz) = s.name

            End If
        Next s
        Set makearangeofShapes = .Shapes.Range(n)
    End With
    
End Function


Public Function UTF16To8(ByVal UTF16 As String) As String
Dim sBuffer As String
#If VBA7 And Win64 Then
    Dim lLength As LongLong
#Else
    Dim lLength As Long
#End If
If UTF16 <> "" Then
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, 0, 0, 0, 0)
    sBuffer = Space$(CLng(lLength))
    lLength = WideCharToMultiByte( _
        CP_UTF8, 0, StrPtr(UTF16), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF16To8 = Left$(sBuffer, CLng(lLength - 1))
Else
    UTF16To8 = ""
End If
End Function




Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False, _
   Optional UTF8Encode As Boolean = True _
) As String

Dim StringValCopy As String: StringValCopy = _
    IIf(UTF8Encode, UTF16To8(StringVal), StringVal)
Dim StringLen As Long: StringLen = Len(StringValCopy)

If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

  If SpaceAsPlus Then Space = "+" Else Space = "%20"

  For i = 1 To StringLen
    Char = Mid$(StringValCopy, i, 1)
    CharCode = asc(Char)
    Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        result(i) = Char
      Case 32
        result(i) = Space
      Case 0 To 15
        result(i) = "%0" & Hex(CharCode)
      Case Else
        result(i) = "%" & Hex(CharCode)
    End Select
  Next i
  URLEncode = join(result, "")

End If
End Function
Public Sub cloneFormat(b As Range, a As Range)
    
    ' this probably needs additional properties copied over
    With a.Interior
        .Color = b.Interior.Color
    End With
    With a.Font
        .Color = b.Font.Color
        .size = b.Font.size
    End With
    With a
        .HorizontalAlignment = b.HorizontalAlignment
        .VerticalAlignment = b.VerticalAlignment
        
    End With

End Sub
' sort a collection
Function SortColl(ByRef coll As Collection, eorder As Long) As Long
    Dim ita As Long, itb As Long
    Dim va As Variant, vb As Variant, bSwap As Boolean
    Dim x As Object, y As Object
    
    For ita = 1 To coll.Count - 1
        For itb = ita + 1 To coll.Count
            Set x = coll(ita)
            Set y = coll(itb)
            bSwap = x.needSwap(y, eorder)
            If bSwap Then
                With coll
                    Set va = coll(ita)
                    Set vb = coll(itb)
                    .add va, , itb
                    .add vb, , ita
                    .remove ita + 1
                    .remove itb + 1
                End With
            End If
        Next
    Next
End Function
Public Function getHandle(sName As String, Optional readOnly As Boolean = False) As Integer
    Dim hand As Integer
    On Error GoTo handleError
        hand = FreeFile
        If (readOnly) Then
            Open sName For Input As hand
        Else
            Open sName For Output As hand
        End If
        getHandle = hand
        Exit Function

handleError:
    MsgBox ("Could not open file " & sName)
    getHandle = cFailedtoGetHandle
End Function
Function afConcat(arr() As Variant) As String
    Dim i As Long, s As String
    s = ""
    For i = LBound(arr) To UBound(arr)
        s = s & arr(i, 1) & "|"
    Next i
    afConcat = s
End Function
Public Function quote(s As String) As String
    quote = q & s & q
End Function
Public Function q() As String
    q = chr(34)
End Function
Public Function qs() As String
    qs = chr(39)
End Function
Public Function bracket(s As String) As String
    bracket = "(" & s & ")"
End Function
Public Function list(ParamArray args() As Variant) As String
    Dim i As Long, s As String
    s = vbNullString
    For i = LBound(args) To UBound(args)
        If s <> vbNullString Then s = s & ","
        s = s & CStr(args(i))
    Next i
    list = s
End Function

Public Function qlist(ParamArray args() As Variant) As String
    Dim i As Long, s As String
    s = vbNullString
    For i = LBound(args) To UBound(args)
        If s <> vbNullString Then s = s & ","
        s = s & quote(CStr(args(i)))
    Next i
    qlist = s
End Function
Public Function diminishingReturn(val As Double, Optional s As Double = 10) As Double
    diminishingReturn = Sgn(val) * s * (Sqr(2 * (Sgn(val) * val / s) + 1) - 1)
End Function

Sub pivotCacheRefreshAll()

    Dim pc As PivotCache
    Dim ws As Worksheet

    With ActiveWorkbook
        For Each pc In .PivotCaches
            pc.refresh
        Next pc
    End With

End Sub
Public Function makeKey(v As Variant) As String
    makeKey = LCase(Trim(CStr(v)))
End Function
' The below is taken from http://stackoverflow.com/questions/496751/base64-encode-string-in-vbscript
Function Base64Encode(sText)
    Dim oXML, oNode
    Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
    Set oNode = oXML.createElement("base64")
    oNode.DataType = "bin.base64"
    oNode.nodeTypedValue = Stream_StringToBinary(sText)
    Base64Encode = oNode.text
    Set oNode = Nothing
    Set oXML = Nothing
End Function
'Stream_StringToBinary Function
'2003 Antonin Foller, http://www.motobit.com
'Text - string parameter To convert To binary data
Function Stream_StringToBinary(text)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.writeText text

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary

  'Ignore first two bytes - sign of
  BinaryStream.Position = 0

  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read

  Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'2003 Antonin Foller, http://www.motobit.com
'Binary - VT_UI1 | VT_ARRAY data To convert To a string
Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1

  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")

  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary

  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.write Binary

  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText

  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"

  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function
' Decodes a base-64 encoded string (BSTR type).
' 1999 - 2004 Antonin Foller, http://www.motobit.com
' 1.01 - solves problem with Access And 'Compare Database' (InStr)
Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function
Public Function openNewHtml(sName As String, sContent As String) As Boolean
    Dim handle As Integer

    handle = getHandle(sName)
    If (handle <> cFailedtoGetHandle) Then
        Print #handle, sContent
        Close #handle
        openNewHtml = True
    End If

End Function
Public Function readFromFile(sName As String) As String
    Dim handle As Integer
    handle = getHandle(sName, True)
    If (handle <> cFailedtoGetHandle) Then
        readFromFile = Input$(LOF(handle), #handle)
        Close #handle
    End If
End Function
Public Function arrayLength(a) As Long
    arrayLength = UBound(a) - LBound(a) + 1
End Function
Public Function getControlValue(ctl As Object) As Variant
    Select Case TypeName(ctl)
        Case "Shape"
            getControlValue = ctl.TextFrame.Characters.text
        Case "Label"
            getControlValue = ctl.Caption
        Case Else
            getControlValue = ctl.value
    End Select
End Function
Public Function setControlValue(ctl As Object, v As Variant) As Variant
    Select Case TypeName(ctl)
        Case "Shape"
            ctl.TextFrame.Characters.text = v
        Case "Label"
            ctl.Caption = v
        Case Else
            ctl.value = v
    End Select
    setControlValue = v
End Function
Public Function isinCollection(vCollect As Variant, sid As Variant) As Boolean
    Dim v As Variant
    If Not vCollect Is Nothing Then
        On Error GoTo handle
        Set v = vCollect(sid)
        isinCollection = True
        Exit Function
    End If
handle:
    isinCollection = False
End Function
'--- based on trig at http://www.movable-type.co.uk/scripts/latlong.html
Public Function getLatFromDistance(mLat As Double, d As Double, heading As Double) As Double
    Dim lat As Double
    ' convert ro radians
    lat = toRadians(mLat)
    getLatFromDistance = _
        fromRadians( _
            Application.WorksheetFunction.Asin(sIn(lat) * _
            Cos(d / earthRadius) + _
            Cos(lat) * _
            sIn(d / earthRadius) * _
            Cos(heading)))
End Function
Public Function getLonFromDistance(mLat As Double, mLon As Double, d As Double, heading As Double) As Double
    Dim lat As Double, lon As Double, newLat As Double
    ' convert ro radians
    lat = toRadians(mLat)
    lon = toRadians(mLon)
    newLat = toRadians(getLatFromDistance(mLat, d, heading))
    getLonFromDistance = _
        fromRadians( _
             (lon + Application.WorksheetFunction.Atan2(Cos(d / earthRadius) - _
            sIn(lat) * _
            sIn(newLat), _
            sIn(heading) * _
            sIn(d / earthRadius) * _
            Cos(lat))))
End Function
Public Function earthRadius() As Double
    ' earth radius in km.
    earthRadius = 6371
End Function
Public Function toRadians(deg)
    toRadians = Application.WorksheetFunction.PI / 180 * deg
End Function
Public Function fromRadians(rad) As Double
    'convert radians to degress
    fromRadians = 180 / Application.WorksheetFunction.PI * rad
End Function
Public Function dimensionCount(a As Variant) As Long
' the only way I can figure out how to do this is to keep trying till it fails
    Dim n As Long, j As Long

    n = 1
    On Error GoTo allDone
    While True
        j = UBound(a, n)
        n = n + 1
    Wend
    Debug.Assert False
    Exit Function
    
allDone:
    dimensionCount = n - 1
    Exit Function
    
End Function
Public Function min(ParamArray args() As Variant)
    min = Application.WorksheetFunction.min(args)
End Function
Public Function max(ParamArray args() As Variant)
    max = Application.WorksheetFunction.max(args)
End Function
Public Function encloseTag(tag As String, Optional newLine As Boolean = True, _
                    Optional tClass As String = vbNullString, _
                    Optional args As Variant) As String
    
    Dim i As Long, t As cStringChunker
    Set t = New cStringChunker
    ' args can be an array or a single item
    If Not IsArray(args) Then
        With t
            .add("<").add (tag)
            If tClass <> vbNullString Then .add(" class=").add (tClass)
            .add (">")
            If newLine Then .add (vbCrLf)
            .add (CStr(args))
            If newLine Then .add (vbCrLf)
            .add("</").add(tag).add (">")
            If newLine Then .add (vbCrLf)
        End With
    Else
        ' recurse for array memmbers
        For i = LBound(args) To UBound(args)
            t.add encloseTag(tag, newLine, tClass, args(i))
        Next i
    End If
    encloseTag = t.content
End Function

Public Function scrollHack() As String
    'hack for IOS
    scrollHack = _
     "<div id='wrapper' style='width:100%;height:100%;overflow-x:auto;" & _
     "overflow-y:auto;-webkit-overflow-scrolling: touch;'>"
End Function

Public Function escapeify(s As String) As String
    escapeify = _
                    Replace( _
                        Replace( _
                            Replace( _
                                Replace(s _
                                    , q, "\" & q), _
                                "%", "\" & "%"), _
                            ">", "\>"), _
                        "<", "\<")
    

    
End Function
Public Function unEscapify(s As String) As String
    unEscapify = _
                    Replace( _
                        Replace( _
                            Replace( _
                                Replace( _
                                    s, "\" & q, q), _
                                 "\" & "%", "%"), _
                             "\>", ">"), _
                         "\<", "<")
    
End Function
Public Function basicStyle() As String
    With New cStringChunker
        .add ".viewdiv {}"
        .add ".hide {"
        .add "display:none;position:absolute;"
        .add "padding:5px;background:white;color:black;"
        .add "border-radius:5px;border:1px solid black;"
        .add "}"
        basicStyle = .content
    End With

End Function
' i adapted this from some table css I found - apologies I dont have the site for crediting.
Public Function tableStyle() As String
    Dim t As cStringChunker
    Set t = New cStringChunker
t.add _
 " table {" & _
    "font-family:Arial, Helvetica, sans-serif;" & _
    "color:#666;" & _
    "font-size:10px;" & _
    "background:#eaebec;" & _
    "margin:4px;" & _
    "border:#ccc 1px solid;" & _
    "-moz-border-radius:3px;" & _
    "-webkit-border-radius:3px;" & _
    "border-radius:3px;" & _
    "-moz-box-shadow: 0 1px 2px #d1d1d1;" & _
    "-webkit-box-shadow: 0 1px 2px #d1d1d1;" & _
    "box-shadow: 0 1px 2px #d1d1d1;" & _
    "}" & _
 "table th {" & _
    "padding:8px 9px 8px 9px;" & _
    "border-top:1px solid #fafafa;" & _
    "border-bottom:1px solid #e0e0e0;" & _
    "background: #ededed;" & _
    "background: -webkit-gradient(linear, left top, left bottom, from(#ededed), to(#ebebeb));" & _
    "background: -moz-linear-gradient(top,  #ededed,  #ebebeb);" & _
    "}"
    
t.add _
 "table tr {" & _
    "text-align: left;" & _
    "padding-left:16px;" & _
    "}" & _
 "table td {" & _
    "padding:6px;" & _
    "border-top: 1px solid #ffffff;" & _
    "border-bottom:1px solid #e0e0e0;" & _
    "border-left: 1px solid #e0e0e0;" & _
    "background: #fafafa;" & _
    "}" & _
 "table tr.even td {" & _
    "background: #f6f6f6;" & _
    "}"


 
    tableStyle = t.content
End Function
Public Function is64BitExcel() As Boolean
#If VBA7 And Win64 Then
    is64BitExcel = True
#Else
    is64BitExcel = False
#End If
End Function
Public Function includeJQuery() As String
    ' include jquery source
    With New cStringChunker
        .addLine jScriptTag("http://www.google.com/jsapi")
        .addLine jScriptTag
        .addLine "google.load('jquery', '1');"
        .addLine "</script>"
        includeJQuery = .content
    End With
    
End Function
Public Function includeGoogleCallBack(c As String) As String
    ' include google call back
    With New cStringChunker
        .addLine jScriptTag
        .addLine "google.setOnLoadCallback("
        .addLine c
        .addLine ");"
        .addLine "</script>"
        includeGoogleCallBack = .content
    End With
    
End Function
Public Function jScriptTag(Optional src As String) As String
    With New cStringChunker
        .add "<script type='text/javascript'"
        If src <> vbNullString Then
            .add(" src='").add(src).addLine ("'></script>")
        Else
            .addLine ">"
        End If
        jScriptTag = .content
    End With
End Function
Public Function jDivAtMouse()
    With New cStringChunker
        .addLine "function() {"
        .add "$('a.viewdiv').mousemove("
        .addLine "function(e) {"
        .add "var targetdiv = $('#d'+this.id);"
        .add "targetdiv.css({left:(e.pageX + 20) + 'px',"
        .add "top: (Math.max(0,e.pageY - targetdiv.height()/2)) + 'px'}).show();"
        .addLine "});"
        .add "$('a.viewdiv').mouseout("
        .addLine "function(e) {"
        .add "$('#d'+this.id).hide();"
        .addLine "});"
        .addLine "}"
        jDivAtMouse = .content
    End With
End Function
Public Function toClipBoard(s As String) As String
    With New MSForms.DataObject
        .SetText s
        .PutInClipboard
    End With
End Function

Public Function importTabbed(fn As String, r As Range) As Range

    r.Worksheet.QueryTables.add(Connection:= _
        "TEXT;" + fn, destination:=r).refresh BackgroundQuery:=False

    Set importTabbed = r
End Function

Function biasedRandom(possibilities, weights) As String
    Dim w As Variant, a As Variant, p As Variant, _
        r As Double, i As Long
    ' comes in as 2 lists
    a = Split(weights, ",")
    p = Split(possibilities, ",")
    ReDim w(LBound(a) To UBound(a))

    ' create cumulative
    For i = LBound(w) To UBound(w)
        w(i) = CDbl(a(i))
        If i > LBound(w) Then w(i) = w(i - 1) + w(i)
    Next i
    
    ' get random index
    r = Rnd() * w(UBound(w))
    
    ' find its weighted position
    For i = LBound(w) To UBound(w)
        If (r <= w(i)) Then
            biasedRandom = p(i)
            Exit Function
        End If
    Next i
    
End Function

Public Sub sleep(seconds As Long)

    Application.Wait TimeSerial(hour(Now()), Minute(Now()), Second(Now()) + seconds)
End Sub
Public Function getDateFromTimestamp(s As String) As Date
    Dim d As Double
    
    If (Len(s) = 13) Then
        ' javaScript Time
        d = CDbl(Left(s, 10))
        ' may need to round for milliseconds
        If Int(Mid(s, 11, 3) >= 500) Then
            d = d + 1
        End If
        
    ElseIf (Len(s) = 10) Then
        ' unix Time
        d = CDbl(s)
    
    Else
        ' wtf time
        getDateFromTimestamp = 0
        Exit Function
    
    End If
    getDateFromTimestamp = DateAdd("s", d, DateSerial(1970, 1, 1))

End Function
Public Function dateFromUnix(s As Variant) As Variant
    Dim d As Date, sd As String
    sd = CStr(s)
    
    If (Len(sd) > 0) Then
        d = getDateFromTimestamp(sd)
        If d = 0 Then
            dateFromUnix = CVErr(xlErrValue)
        Else
            dateFromUnix = d
        End If
    Else
        dateFromUnix = Empty
    End If

End Function
Public Function isSomething(o As Object) As Boolean

    isSomething = Not o Is Nothing
End Function


Public Function tinyTime() As Double
' Returns seconds.
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    tinyTime = 0
' Get frequency.
    If cyFrequency = 0 Then getFrequency cyFrequency
' Get ticks.
    getTickCount cyTicks1
    If cyFrequency Then tinyTime = cyTicks1 / cyFrequency
End Function

Function isUndefined(value As Variant) As Boolean
    If (IsObject(value)) Then
        isUndefined = value Is Nothing
    Else
        If (IsMissing(value) Or IsEmpty(value)) Then
            isUndefined = True
        Else
            isUndefined = (value = vbNullString)
        End If
    End If
End Function
Function conditionalAssignment(condition As Boolean, a As Variant, b As Variant) As Variant
    If (condition) Then

        If IsObject(a) Then
            Set conditionalAssignment = a
        Else
            conditionalAssignment = assignHelper(a)
        End If

    Else
        If IsObject(b) Then
            Set conditionalAssignment = b
        Else
            conditionalAssignment = assignHelper(b)
        End If
    
    End If
End Function
Private Function assignHelper(a As Variant) As Variant
    If IsObject(a) Then
        Set assignHelper = a
    Else
        If Not isUndefined(a) Then
            assignHelper = a
        Else
            assignHelper = vbNullString
        End If
    End If
End Function
Public Function getTimestampFromDate(Optional dt As Date = 0) As Double
    Dim d As Double
    
    If (dt = 0) Then
        dt = Now()
    End If
    
    ' convert into time since the epoch
    d = DateDiff("s", DateSerial(1970, 1, 1) + TimeSerial(0, 0, 0), dt)
    
    ' convert to ms
    d = d * 1000#

    getTimestampFromDate = d

End Function

Public Function checkOrCreateFolder(path As String, Optional optCreate As Boolean = True) As Object
    ' doing late binding to avoid refernce for this
    
    Dim fso As Object, cleanPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    'fso not smart enough to create entire thing, so we need recurse for each
    If (optCreate) Then
        recurseCreateFolder fso, fso.GetAbsolutePathName(path)
    End If
    
    Set checkOrCreateFolder = fso.getFolder(path)
End Function
Private Function recurseCreateFolder(fso As Object, cleanPath As String) As Object
    Dim parentPathString As String

    If Not fso.folderExists(cleanPath) Then
        ' need to create the parent first
        recurseCreateFolder fso, fso.GetParentFolderName(cleanPath)
        ' now we can do it
        fso.CreateFolder (cleanPath)
    End If

End Function
Public Function writeToFolderFile(folderName As String, fileName As String, content As String) As String
    
    Dim file As Object, fso As Object
    Dim path As String
    path = fileName
    ' create the folder if we need to
    
    If (folderName <> vbNullString) Then
        path = concatFolderName(folderName, path)
        checkOrCreateFolder folderName
    End If
    ' write the data
  
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(path, 2, True)
    file.write content
    writeToFolderFile = content
End Function
Public Function getAllSubFolderPaths(folderName As String) As String
    Dim folder As Object, subFolder As Object, c As cStringChunker
    Set c = New cStringChunker
    Set folder = checkOrCreateFolder(folderName, False)
    If (isSomething(folder)) Then
        For Each subFolder In folder.subFolders
            c.add(subFolder.path).add ","
        Next subFolder
    End If
    getAllSubFolderPaths = c.chopIf(",").toString
End Function
Public Function readFromFolderFile(folderName As String, fileName As String) As String
    Dim file As Object, fso As Object
    Dim path As String
    path = fileName
    
    If (folderName <> vbNullString) Then
        path = concatFolderName(folderName, fileName)
    End If
    ' read the data
    If (fileExists(path)) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(path, 1)
        If (file.AtEndOfStream) Then
            readFromFolderFile = ""
        Else
            readFromFolderFile = file.readAll()
        End If
        file.Close
    End If
    
End Function
Public Function fileExists(path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileExists = fso.fileExists(path)
End Function
Public Function concatFolderName(folderName As String, fileName As String) As String
    Dim c As cStringChunker
    Set c = New cStringChunker
    concatFolderName = c.add(folderName).chopIf("/").chopIf("\").add("/").add(fileName).toString
    
End Function

' v0.1.1 27.3.15
Public Function straightenOutContinuations(s As String) As String
    ' just get rid of continuations and move
    straightenOutContinuations = getRidOfMultipleSpaces( _
        getRx("_\s*$\n").Replace(s, ""))
End Function
Public Function getRidOfDims(s As String) As String
    'get rid of dims which may have locals matching function names, but have to leave if new mentioned
    getRidOfDims = getRx("\s*dim(?!.*\s*new\s*).*").Replace(s, "")
End Function
Public Function getRidOfQuoted(s As String) As String
    getRidOfQuoted = getRx("(""[^""]*"")").Replace(s, "")
End Function
Public Function getRidOfComments(s As String) As String
    getRidOfComments = getRx("("".*?"")|('.*$)").Replace(s, "$1")
End Function
Public Function getRidOfMultipleSpaces(s As String) As String
    getRidOfMultipleSpaces = getRx("[\t ]{2,}", False).Replace(s, " ")
End Function
Public Function getRx(pattern As String, Optional multi As Boolean = True) As RegExp
    Dim rx As RegExp
    Set rx = New RegExp
    With rx
        .ignorecase = True
        .Global = True
        .MultiLine = multi
        .pattern = pattern
    End With
    Set getRx = rx
End Function
'/**
' *@return {} get a regex that picks out the end of a sub/function
'*/
Public Function getTheEndRx() As RegExp
    Set getTheEndRx = getRx("\bend\s*function|sub|property")
End Function
'/**
' *@return {} get a regex that picks out all lines with Dim
'*/
Public Function getDimLinesRx() As RegExp
    Set getDimLinesRx = getRx("(^\s*dim\s+.*)$")
End Function
'/**
' *@return {} get a regex that picks out all locally defined variables from a dim
'*/
Public Function getDimLocalsRx() As RegExp
    Set getDimLocalsRx = getRx("dim|(?:\s+as\s+)(\w+)")
End Function

'/**
' * since I use this all the time,may as well make it a library
' * does UrlFetch() stuff and creates standard results
' */

'/**
'* execute a get
'* @param {string} url the url
'* @param {string} optAccessToken an optional access token
'* @param {object} optOptions optional headers
'* @param {boolean} optBasic the access token is for basic auth
'* @return {object} a standard response
'*/
Public Function urlGet(url As String, _
    Optional optOptions As cJobject = Nothing, _
    Optional optAccessToken As String, _
    Optional optBasic As Boolean = False) As cJobject
    
    Set urlGet = _
        urlExecute(url, "GET", _
            vbNullString, _
            optOptions, _
            optAccessToken, _
            optBasic)
        
End Function

'/**
'* execute a post
'* @param {string} url the url
'* @param {string} optMethod the http method
'* @param {string} optPayload any payload
'* @param {string} optAccessToken an optional access token
'* @param {object} optOptions optional headers
'* @param {boolean} optBasic the access token is for basic auth
'* @return {object} a standard response
'*/
Public Function urlPost(url As String, _
    Optional optMethod As String = "POST", _
    Optional optPayload As Variant, _
    Optional optOptions As cJobject = Nothing, _
    Optional optAccessToken As String, _
    Optional optBasic As Boolean = False) As cJobject
    Dim payload As String
    If (IsObject(optPayload)) Then
        payload = optPayload.stringify
    Else
        If (isUndefined(optPayload)) Then
            payload = vbNullString
        Else
            payload = optPayload
        End If
    End If
    Set urlPost = _
        urlExecute(url, _
            optMethod, _
            payload, _
            optOptions, _
            optAccessToken, _
            optBasic)
  
End Function
'/**
'* execute a urlfetch
'* @param {string} url the url
'* @param {string} optMethod the http method
'* @param {string} optPayload any payload
'* @param {string} optAccessToken an optional access token
'* @param {object} optOptions optional headers
'* @param {boolean} optBasic the access token is for basic auth
'* @return {object} a standard response
'*/
Private Function urlExecute(url As String, _
    Optional optMethod As String = "GET", _
    Optional optPayload As String = vbNullString, _
    Optional optOptions As cJobject = Nothing, _
    Optional optAccessToken As String, _
    Optional optBasic As Boolean = False)
    
    Dim job As cJobject
    
    ' we'll need some headers
    If (optOptions Is Nothing) Then
        Set optOptions = New cJobject
        optOptions.init Nothing
    End If
    
    If (optOptions.childExists("headers") Is Nothing) Then
        optOptions.add "headers"
    End If
    
    ' apply the access token/ basic auth if there is one
    If (Not isUndefined(optAccessToken)) Then
        optOptions.child("headers").add "authorization", _
            conditionalAssignment(optBasic, "Basic ", "Bearer ") & optAccessToken
    End If
    

    ' do the operation - we're using server http .. better for cors
    Dim ohttp As MSXML2.ServerXMLHTTP60
    Set ohttp = New MSXML2.ServerXMLHTTP60
    
    With ohttp
        ' this is for some MS bug .. cant remember which now
        .setOption 2, .getOption(2) - SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID
        
        ' set it up
        .Open optMethod, url, False
        
        ' set the headers
        For Each job In optOptions.child("headers").children
            .setRequestHeader job.key, job.value
        Next job
        
        ' execute the thing
        .send optPayload
        
    End With
    
    ' turn results into standard
    Set urlExecute = makeResults(ohttp, url)
    
End Function
'/**
'* this is a standard result object to simply error checking etc.
'* @param {HTTPResponse} response the response from UrlFetchApp
'/ @param {string} optUrl the url if given
'* @return {object} the result object
'*/
Private Function makeResults(Response As Object, Optional optUrl As String = vbNullString)
    Dim result As cJobject, job As cJobject, _
        rx As RegExp, matches As MatchCollection, match As match, i
    
    ' default result
    Set result = JSONParse("{" & _
        "'success':false," & _
        "'data':null," & _
        "'code':null," & _
        "'url':'" & optUrl & "'," & _
        "'extended':'failed to parse'," & _
        "'parsed':false }")

'   // process the result
    If (Not isUndefined(Response)) Then
        result.add "code", Response.status
        result.add "headers"
        result.add "content", Response.responseText
        result.add "success", (result.cValue("code") = 200 Or result.cValue("code") = 201)

        ' parse
        Set job = JSONParse(result.cValue("content"), , False)
        If (Not isUndefined(job) And job.isValid) Then
            result.child("data").setValue job
            result.add "parsed", True
        End If
        
        ' headers - MS doesnt do this for you
        Set rx = New RegExp
        With rx
            .MultiLine = True
            .Global = True
            .pattern = "^([^:]+)\s*:\s*(.+)*"
            Set matches = .Execute(Response.getAllResponseHeaders())
        End With
        For i = 0 To matches.Count - 1
            result.child("headers").add _
                rxReplace("^\s*", matches.Item(i).SubMatches(0), ""), _
                rxReplace("\s*$", matches.Item(i).SubMatches(1), "")
        Next i
    End If
    Set makeResults = result

End Function

' v2.02
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
Public Function rxString(sName As String, s As String, Optional ignorecase As Boolean = True) As String
    Dim rx As cregXLib
    ' create a new regx
    Set rx = rxMakeRxLib(sName)
    rx.ignorecase = ignorecase
    ' extract the string that matches the requested pattern
    rxString = rx.getString(s)

End Function
Public Function rxGroup(sName As String, s As String, group As Long, Optional ignorecase As Boolean = True) As String
    Dim rx As cregXLib
    ' create a new regx
    Set rx = rxMakeRxLib(sName)
    rx.ignorecase = ignorecase
    ' extract the string that matches the requested pattern
    rxGroup = rx.getGroup(s, group)

End Function
Public Function rxTest(sName As String, s As String, Optional ignorecase As Boolean = True) As Boolean
    Dim rx As cregXLib
    ' create a new regx
    Set rx = rxMakeRxLib(sName)
    rx.ignorecase = ignorecase
    ' extract the string that matches the requested pattern
    rxTest = rx.getTest(s)

End Function
Public Function rxReplace(sName As String, sFrom As String, sTo As String, Optional ignorecase As Boolean = True) As String
    Dim rx As cregXLib
     ' create a new regx
    Set rx = rxMakeRxLib(sName)
    rx.ignorecase = ignorecase
    ' replace the string that matches the requested pattern
    rxReplace = rx.getReplace(sFrom, sTo)
    
End Function
Public Function rxPattern(sName As String) As String
    Dim rx As cregXLib
     ' create a new regx
    Set rx = rxMakeRxLib(sName)
    ' just returnthe pattern
    rxPattern = rx.pattern
    
End Function
 Function rxMakeRxLib(sName As String) As cregXLib
    Dim rx As cregXLib, s As String
    Set rx = New cregXLib
    ' normally sname points to a preselected regEX
    ' if not known, silently assume its a regex pattern
        s = Replace(UCase(sName), " ", "")
        Select Case s
            Case "POSTALCODEUK"
                rx.init s, _
                "(((^[BEGLMNS][1-9]\d?) | (^W[2-9] ) | ( ^( A[BL] | B[ABDHLNRST] | C[ABFHMORTVW] | D[ADEGHLNTY] | E[HNX] | F[KY] | G[LUY] | H[ADGPRSUX] | I[GMPV] |" & _
                " JE | K[ATWY] | L[ADELNSU] | M[EKL] | N[EGNPRW] | O[LX] | P[AEHLOR] | R[GHM] | S[AEGKL-PRSTWY] | T[ADFNQRSW] | UB | W[ADFNRSV] | YO | ZE ) \d\d?) |" & _
                " (^W1[A-HJKSTUW0-9]) | ((  (^WC[1-2])  |  (^EC[1-4]) | (^SW1)  ) [ABEHMNPRVWXY] ) ) (\s*)?  ([0-9][ABD-HJLNP-UW-Z]{2})) | (^GIR\s?0AA)"
            
            Case "POSTALCODESPAIN"
                rx.init s, _
                    "^([1-9]{2}|[0-9][1-9]|[1-9][0-9])[0-9]{3}$"
                    
            Case "PHONENUMBERUS"
                rx.init s, _
                "^\(?(?<AreaCode>[2-9]\d{2})(\)?)(-|.|\s)?(?<Prefix>[1-9]\d{2})(-|.|\s)?(?<Suffix>\d{4})$"
                
            Case "CREDITCARD" 'amex/visa/mastercard
                rx.init s, _
                "^((4\d{3})|(5[1-5]\d{2}))(-?|\040?)(\d{4}(-?|\040?)){3}|^(3[4,7]\d{2})(-?|\040?)\d{6}(-?|\040?)\d{5}"
                
            Case "NUMERIC"
                rx.init s, _
                    "[\0-9]"
            
            Case "ALPHABETIC"
                rx.init s, _
                    "[\a-zA-Z]"
                    
            Case "NONNUMERIC"
                rx.init s, _
                    "[^\0-9]"
                    
            Case "IPADDRESS"
                rx.init s, _
                "^(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])\.(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$"
            
            Case "SINGLESPACE"  ' should take a replace value of "$1 "
                rx.init s, _
                    "(\S+)\x20{2,}(?=\S+)"
            
            Case "EMAIL"
                rx.init s, _
                    "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}$"
                    
            Case "EMAILINSIDE"
                rx.init s, _
                    "\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b"
                    
            Case "NONPRINTABLE"
                rx.init s, "[\x00-\x1F\x7F]"
                
                
            Case "PUNCTUATION"
                rx.init s, "[^A-Za-z0-9\x20]+"

            Case Else
                rx.init "Adhoc", sName
        
        End Select
    
    Set rxMakeRxLib = rx
End Function


Public Function fromISODateTime(iso As String) As Date
    Dim rx As RegExp, matches As MatchCollection, d As Date, ms As Double, sec As Double
    Set rx = New RegExp
    With rx
        .ignorecase = True
        .Global = True
        .pattern = "(\d{4})-([01]\d)-([0-3]\d)T([0-2]\d):([0-5]\d):(\d*\.?\d*)Z"
    End With
    Set matches = rx.Execute(iso)
    
    ' TODO -- timeszone

    If matches.Count = 1 And matches.Item(0).SubMatches.Count = 6 Then

        With matches.Item(0)
            sec = CDbl(.SubMatches(5))
            ms = sec - Int(sec)
            d = DateSerial(.SubMatches(0), .SubMatches(1), .SubMatches(2)) + _
                TimeSerial(.SubMatches(3), .SubMatches(4), Int(sec)) + ms / 86400
        End With
    
    Else
        d = 0
    End If
    
    fromISODateTime = d
   
End Function

Public Function toISODateTime(d As Date) As String
    Dim s As String, ms As Double, adjustSecond As Long
    
    ' need to adjust if seconds are going to be rounded up
    ms = milliseconds(d)
    adjustSecond = 0
    If (ms >= 0.5) Then adjustSecond = -1
    
    ' TODO - timezone
    toISODateTime = Format(year(d), "0000") & "-" & Format(month(d), "00") & "-" & Format(day(d), "00T") & _
            Format(d, "hh:mm:") & Format(DateAdd("s", adjustSecond, d), "ss") & Format(ms, ".000Z")

    
End Function
Public Function milliseconds(d As Date) As Double
    ' extract the milliseconds from the time
    Dim t As Date
    t = (d - DateSerial(year(d), month(d), day(d)) - TimeSerial(hour(d), Minute(d), Second(d)))
    If t < 0 Then
        ' the millsecond rounded it up
        t = (d - DateSerial(year(d), month(d), day(d)) - TimeSerial(hour(d), Minute(d), Second(d) - 1))
    End If
    
    milliseconds = t * 86400
    
End Function
Public Function JSONParse(s As String, Optional jtype As eDeserializeType, Optional complain As Boolean = True) As cJobject
    Dim j As New cJobject
    Set JSONParse = j.init(Nothing).parse(s, jtype, complain)
    j.tearDown
End Function
Public Function JSONStringify(j As cJobject, Optional blf As Boolean) As String
    JSONStringify = j.stringify(blf)
End Function
Public Function jSonArgs(options As String) As cJobject
    ' takes a javaScript like options paramte and converts it to cJobject
    ' it can be accessed as job.child('argName').value or job.find('argName') etc.
    Dim job As New cJobject
    If options <> vbNullString Then
        Set jSonArgs = job.init(Nothing, "jSonArgs").deSerialize(options)
    End If
End Function
Public Function optionsExtend(givenOptions As String, _
            Optional defaultOptions As String = vbNullString) As cJobject
    Dim jGiven As cJobject, jDefault As cJobject, _
        jExtended As cJobject, cj As cJobject
    ' this works like $.extend in jQuery.
    ' given and default options arrive as a json string
    ' example -
    ' optionsExtend ("{'width':90,'color':'blue'}", "{'width':20,'height':30,'color':'red'}")
    ' would return a cJobject which serializes to
    ' "{width:90,height:30,color:blue}"
    Set jGiven = jSonArgs(givenOptions)
    Set jDefault = jSonArgs(defaultOptions)
    
    ' now we combine them
    If Not jDefault Is Nothing Then
        Set jExtended = jDefault
    Else
        Set jExtended = New cJobject
        jExtended.init Nothing
    End If
    
    ' now we merge that with whatever was given
    If Not jGiven Is Nothing Then
        jExtended.merge jGiven
    End If
    
    ' and its over
    Set optionsExtend = jExtended
End Function

'udfs to expose classes
Public Function ucJobjectMake(r As Variant) As cJobject
    Dim cj As New cJobject
    Set ucJobjectMake = cj.deSerialize(CStr(r))
End Function
Public Function ucJobjectChildValue(json As Variant, child As Variant) As String
    ucJobjectChildValue = ucJobjectMake(CStr(json)).child(CStr(child)).value
End Function
Public Function ucJobjectLint(json As Variant, Optional child As Variant) As String
    Dim cj As cJobject
    Set cj = ucJobjectMake(json)
    If Not IsMissing(child) Then
        Set cj = cj.child(CStr(child))
    End If
    ucJobjectLint = cj.serialize(True)
End Function
Public Function cleanGoogleWire(sWire As String) As String
    Dim jStart As String, p As Long, newWire As Boolean, e As Long, s As String, reg As RegExp, _
        match As match, matches As MatchCollection, v As Double, i As Long, _
        year As Long, month As Long, day As Long, hour As Long, min As Long, sec As Long, ms As Long, _
        t As cStringChunker, consumed As Long

    jStart = "table:"
    p = InStr(1, sWire, jStart)
    'there have been multiple versions of wire ...
    If p = 0 Then
        'try the other one
        jStart = q & ("table") & q & ":"
        p = InStr(1, sWire, jStart)
        newWire = True
    End If

    p = InStr(1, sWire, jStart)
    e = Len(sWire) - 1

    If p <= 0 Or e <= 0 Or p > e Then
        MsgBox " did not find table definition data"
        Exit Function
    End If
    
    If Mid(sWire, e, 2) <> ");" Then
        MsgBox ("incomplete google wire message")
        Exit Function
    End If
    ' encode the 'table:' part to a cjobject
    p = p + Len(jStart)
    s = "{" & jStart & "[" & Mid(sWire, p, e - p - 1) & "]}"
    ' google protocol doesnt have quotes round the key of key value pairs,
    ' and i also need to convert date from javascript syntax new Date()
    ' we'll force it to be a 13 digit timestamp, since cjobject knows how to make that into a date
    's = rxReplace("(new\sDate)(\()(\d+)(,)(\d+)(,)(\d+)(\))", s, "'$3/$5/$7'")
    'new\s+date\s*\(\s*(\d+)\s*(,\s*\d+)\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\)
    Set reg = New RegExp
    With reg
        .pattern = "new\s+Date\s*\(\s*(\d+)\s*(,\s*\d+)\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\s*(,\s*\d+)?\)"
        .Global = True
    End With
    Set matches = reg.Execute(s)

    
    If matches.Count > 0 Then
        Set t = New cStringChunker
        consumed = 0
        For Each match In matches
            t.add Mid(s, consumed + 1, match.FirstIndex - consumed)
            consumed = consumed + match.FirstIndex - consumed
            With match
                If .SubMatches.Count >= 2 And .SubMatches.Count <= 7 Then
                    'these are the only valid number of args to a javascript new Date()
                    day = 1
                    hour = 0
                    min = 0
                    sec = 0
                    ms = 0
                    year = .SubMatches(0)
                    month = Replace(.SubMatches(1), ",", "") + 1
                    If .SubMatches.Count > 2 And Not IsEmpty(.SubMatches(2)) Then day = Replace(.SubMatches(2), ",", "")
                    If .SubMatches.Count > 3 And Not IsEmpty(.SubMatches(3)) Then hour = Replace(.SubMatches(3), ",", "")
                    If .SubMatches.Count > 4 And Not IsEmpty(.SubMatches(4)) Then min = Replace(.SubMatches(4), ",", "")
                    If .SubMatches.Count > 5 And Not IsEmpty(.SubMatches(5)) Then sec = Replace(.SubMatches(5), ",", "")
                    If .SubMatches.Count > 6 And Not IsEmpty(.SubMatches(6)) Then ms = Replace(.SubMatches(6), ",", "")
                    ' now convert to a date and format
                    t.add(q) _
                        .add(CStr(DateSerial(year, month, day) + TimeSerial(hour, min, sec) + CDbl(ms) / 86400)) _
                        .add (q)
                    consumed = consumed + match.Length
                End If
            End With
        Next match
        If consumed < Len(s) Then t.add Mid(s, consumed + 1)
        s = t.content
        Set t = Nothing
    End If
    If Not newWire Then s = rxReplace("(\w+)(:)", s, "'$1':")
    cleanGoogleWire = s
    
End Function

Public Function xmlStringToJobject(xmlString As String, Optional complain As Boolean = True) As cJobject
    Dim doc As Object
    ' parse xml

    Set doc = CreateObject("msxml2.DOMDocument")
    doc.LoadXML xmlString
    If doc.parsed And doc.parseError = 0 Then
        Set xmlStringToJobject = docToJobject(doc, complain)
        Exit Function
    End If

    Set xmlStringToJobject = Nothing
    If complain Then
        MsgBox ("Invalid xml string - xmlparseerror code:" & doc.parseError)
    End If
    
    Exit Function
    
End Function
Public Function docToJobject(doc As Object, Optional complain As Boolean = True) As cJobject
    ' convert xml document to a cjobject
    Dim node As IXMLDOMNode, job As cJobject
    Set job = New cJobject
    job.init Nothing
       
    Set docToJobject = handleNodes(doc, job)
End Function
Private Function isArrayRoot(parent As IXMLDOMNode) As Boolean
    
    Dim node As IXMLDOMNode, n As Long, node2 As IXMLDOMNode
    
    
    isArrayRoot = False
    If parent.NodeType = NODE_ELEMENT And parent.ChildNodes.Length > 1 Then
        For Each node2 In parent.ChildNodes
            If node2.NodeType = NODE_ELEMENT Then
                n = 0
                For Each node In parent.ChildNodes
                    If node.NodeType = NODE_ELEMENT And _
                        node2.nodeName = node.nodeName Then n = n + 1
                Next node
                If n > 1 Then
                    ' this shoudl be true, but for leniency i'll comment
                    'Debug.Assert n = parent.ChildNodes.Length
                    isArrayRoot = True
                    Exit Function
                End If
            End If
        Next node2
    End If

    
End Function
Private Function handleNodes(parent As IXMLDOMNode, job As cJobject) As cJobject
    Dim node As IXMLDOMNode, joc As cJobject, attrib As IXMLDOMAttribute, i As Long, _
         arrayJob As cJobject
    
    If isArrayRoot(parent) Then
        ' we need an array associated with this this node
        ' subsequent members will need to make space for themselves
        Set joc = job.add(parent.nodeName).addArray
    Else
        Set joc = handleNode(parent, job)
    End If
    
    ' deal with any attributes
    If Not parent.Attributes Is Nothing Then
        For Each attrib In parent.Attributes
            handleNode attrib, joc
        Next attrib
    End If
    
    ' do the children
    If Not parent.ChildNodes Is Nothing And parent.ChildNodes.Length > 0 Then
        For Each node In parent.ChildNodes
            handleNodes node, joc
        Next node
    End If
    
    ' always return the level at which we arrived
    Set handleNodes = job
    
End Function
Private Function handleNode(node As IXMLDOMNode, job As cJobject, Optional arrayHead As Boolean = False) As cJobject
    Dim key As cJobject
    '' not a comprehensive convertor
    Set handleNode = job
    Debug.Print node.nodeName & node.NodeType & node.NodeValue
    Select Case node.NodeType
        Case NODE_ATTRIBUTE
            ' we cant have an array of attributes - this will silently use the latest
            job.add node.nodeName, node.NodeValue
            
        Case NODE_ELEMENT
            If job.isArrayRoot Then
                Dim b As Boolean
                b = (node.ChildNodes.Length = 1)
                If (b) Then b = node.ChildNodes(0).NodeType = NODE_TEXT
                If (b) Then
                    Set handleNode = job.add.add
                Else
                    Set handleNode = job.add.add(node.nodeName)
                End If
            Else
                Set handleNode = job.add(node.nodeName)
            End If

        Case NODE_TEXT
            job.value = node.NodeValue

            
        Case NODE_DOCUMENT, NODE_CDATA_SECTION, NODE_ENTITY_REFERENCE, _
            NODE_ENTITY, NODE_PROCESSING_INSTRUCTION, NODE_COMMENT, NODE_DOCUMENT_TYPE, _
            NODE_DOCUMENT_FRAGMENT, NODE_NOTATION
            ' just ignore these for now

        Case Else
            Debug.Assert False
    End Select
    
End Function


Public Function compareAsKey(a As Variant, b As Variant, Optional asKey As Boolean = True) As Boolean
    If (asKey And TypeName(a) = "String" And TypeName(b) = "String") Then
        compareAsKey = (makeKey(a) = makeKey(b))
    Else
        compareAsKey = (a = b)
    
    End If
End Function

Public Function superTrim(s As String) As String
    Dim c As cStringChunker
    Set c = New cStringChunker
    superTrim = c.add(s).chopSuperTrim.toString
    
End Function


