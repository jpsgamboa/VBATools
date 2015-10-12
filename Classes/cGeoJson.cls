VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGeoJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Creates and returns a GeoJSON file

Private geoJsonFile As Object
Private pfields As Variant 'An array of strings with the field names
Private firstFeature As Boolean

'Creates and saves the new json file
Public Sub newGeoJSON()

    'Dim geoJsonFile
    Set geoJsonFile = CreateObject("ADODB.Stream")   'Set geoJsonFile = fso.CreateTextFile(filePath & "\" & fileName & ".geojson", True)
    geoJsonFile.Type = 2
    geoJsonFile.Charset = "utf-8"
    geoJsonFile.Open
    
End Sub

'Begins the file scheme
Public Function begin(fields As Variant)
    wrtln ("{")
    wrtln (quote("type") & ": " & quote("FeatureCollection") & ",")
    wrtln (quote("crs") & ": { " & quote("type") & ": " & quote("name") & ", " _
                                 & quote("properties") & ": { " & quote("name") _
                                 & ": " & quote("urn:ogc:def:crs:OGC:1.3:CRS84") & " } },")
    wrtln ("")
    wrtln (quote("features") & ": [")
    
    pfields = fields
    firstFeature = True
End Function

'Finishes the file scheme
Public Function finish(fileName As String, filePath As String)
    wrtln ("]"): wrt ("}")
    
    Dim fullPath As String: fullPath = (filePath & "\" & fileName & ".geojson")
    
    Dim withBOM As Boolean: withBOM = False
        
    If withBOM Then
        geoJsonFile.SaveToFile fullPath, 2
    Else
        Dim binStream: Set binStream = CreateObject("ADODB.Stream")
        binStream.Type = 1
        binStream.Open
        geoJsonFile.Position = 3
        geoJsonFile.CopyTo binStream
        binStream.SaveToFile fullPath, 2
        binStream.Close
    End If
        
    geoJsonFile.Close
    
    
End Function

'Adds a new feature to the file
Public Function addFeature(fieldValues As Variant, ByRef coordinates As Collection)
    If Not firstFeature Then wrtln (",")
    
    wrt ("{ ")
        wrt (quote("type") & ": " & quote("Feature") & ", ")
        
        wrt (quote("properties") & ": ")
        wrt ("{ ")
        
            wrtProperties fieldValues
            
        wrt ("}, ")
            
        wrt (quote("geometry") & ": ")
        wrt ("{ ")
        
            wrtField "type", "LineString": wrt (", ")
            
            wrt (quote("coordinates") & ": ")
            wrt ("[ ")
            
                wrtCoordinates coordinates
            
            wrt ("] ")
            
        wrt ("} ")
            
    wrt ("}")
    
    firstFeature = False
End Function

'Writes a new field and it's value, enclosed in quotes if it's a string
Private Function wrtField(fieldName As String, value As String)
    Dim formattedValue As String: formattedValue = value
    If Not IsNumeric(value) Then formattedValue = quote(value)
    wrt (quote(fieldName) & ": "): wrt (formattedValue)
End Function

'Writes each field value in the order they're provided
Private Function wrtProperties(fieldValues As Variant)
    Dim fieldName As Variant, i As Integer
    For Each fieldName In pfields
        If i > 0 Then wrt (", ")
        wrtField (fieldName), (fieldValues(i))
        i = i + 1
    Next
End Function

'Writes the coordinates provided in the collection
Private Function wrtCoordinates(ByRef coordinates As Collection)
    Dim coord As cGeoCoordinate, i As Integer
    For Each coord In coordinates
        If i > 0 Then wrt (", ")
        wrt ("[ " & coord.longitude & "," & coord.latitude & " ]")
        i = i + 1
    Next coord
End Function

'Writes to the file withoung sending the cursor to the next line
Private Function wrt(str As String)
    If Not geoJsonFile Is Nothing Then geoJsonFile.writeText (str)
End Function

'Writes to the file and sets the cursor on the next line
Private Function wrtln(str As String)
    wrt (str & vbCrLf)
End Function
