Attribute VB_Name = "GoogleMapsAPI"
Option Explicit

Private Const gMapsSignature = "INSERT_YOUR_OWN"

'Returns the distance and the time of travel based on start and end adress
Public Function GoogleDistanceTime(origin As String, destination As String, _
                                    Optional avoidTolls As Boolean = False, Optional avoidHighways As Boolean = False, _
                                    Optional mode As String = "DRIVING", _
                                    Optional waypoints As Variant, Optional optimize As Boolean = False) As Variant
    'Make request
    Dim domDoc As DOMDocument60: Set domDoc = getXMLResponse(origin, destination, avoidTolls, avoidHighways, mode, waypoints, optimize)
    
    'Evaluate the response
    Dim status As String: status = handleStatus(domDoc)
    
    If strEquals(status, "OK") Then
        Dim finalDist As Double, finalTime As Double
        
        Dim legs As IXMLDOMNodeList: Set legs = domDoc.SelectNodes("//leg")
        
        Dim leg As Variant, child As Variant, node As IXMLDOMNode
        For Each leg In legs
            Dim legNode As IXMLDOMNode: Set legNode = leg 'Casting
            
            Dim distanceNode As IXMLDOMNode: Set distanceNode = getXMLNodeImmediateChild("distance", legNode)
            Dim timeNode As IXMLDOMNode: Set timeNode = getXMLNodeImmediateChild("duration", legNode)
                         
            Dim dist As Double: dist = Round(getXMLNodeImmediateChild("value", distanceNode).nodeTypedValue / 1000, 3)
            Dim time As Double: time = Round(getXMLNodeImmediateChild("value", timeNode).nodeTypedValue / 60, 2)
                        
            finalDist = finalDist + dist
            finalTime = finalTime + time
        Next leg
                            
        Dim result(1, 1 To 2) As Variant

        result(0, 1) = finalDist
        result(0, 2) = finalTime

        GoogleDistanceTime = result
    Else
        GoogleDistanceTime = status
    End If
    
    Set domDoc = Nothing

End Function

'Handles the status of the response
Private Function handleStatus(xml As DOMDocument60) As String
    Dim statusNode As IXMLDOMNode: Set statusNode = xml.SelectSingleNode("//status")
    If strEquals(statusNode.text, "OK") Then
        handleStatus = "OK"
    Else
        handleStatus = handleError(statusNode.text)
    End If
End Function

'Handles the error message
Private Function handleError(errorMessage As String) As String
    Select Case UCase(errorMessage)
        Case "ZERO_RESULTS"   'The geocode was successful but returned no results.
            handleError = "ZERO_RESULTS"
            Application.StatusBar = "No results"
        Case "NOT_FOUND"
            handleError = "NOT_FOUND"
        Case "INVALID_REQUEST"
            handleError = "INVALID_REQUEST"
        Case "MAX_WAYPOINTS_EXCEEDED"
            handleError = "MAX_WAYPOINTS_EXCEEDED"
        Case "REQUEST_DENIED"
            handleError = "REQUEST_DENIED"
        Case "OVER_QUERY_LIMIT" 'The requestor has exceeded the limit of 2500 request/day.
            handleError = "The requestor has exceeded the limit of 2500 request/day."
        Case Else
            handleError = "Error"
    End Select
End Function

'Adds the signature to the url
Private Function addSignature(url As String) As String
    Dim keyString As String
    keyString = "&key=" & gMapsSignature
    addSignature = url & keyString
End Function

'Calls google and returns the resulting XML document
Private Function getXMLResponse(origin As String, destination As String, _
                                Optional avoidTolls As Boolean = False, Optional avoidHighways As Boolean = False, _
                                Optional mode As String = "DRIVING", _
                                Optional waypoints As Variant, Optional optimize As Boolean = False) As DOMDocument60

    'Prepare URL
    Dim url As String: url = "https://maps.googleapis.com/maps/api/directions/xml?"
    
    'OD:
    url = url & "origin=" & origin & "&destination=" & destination & "&region=pt"
    
    'Waypoints
    If Not IsMissing(waypoints) Then
        url = url & "&waypoints="
        
        If optimize Then url = url & "optimize:true|"
            
        If TypeOf waypoints Is Range Then
            If waypoints.Height > 8 Or waypoints.Width > 8 Then MsgBox "Não podem ser introduzidos mais de 8 waypoints", vbInformation
            Dim c As Variant, i As Integer
            For Each c In waypoints
                If i > 0 Then url = url & "|"
                url = url & c.value
                i = i + 1
            Next c
        Else
            url = url & waypoints
        End If
    End If
    
    'Avoid
    If avoidTolls And avoidHighways Then url = url & "&avoid=tolls|highways"
    If avoidTolls And Not avoidHighways Then url = url & "&avoid=tolls"
    If Not avoidTolls And avoidHighways Then url = url & "&avoid=highways"
    
    'Travel mode
    If Not strEquals(mode, "DRIVING") Then
        url = url & "&mode="
        If strEquals(mode, "transit") Then
            url = url & "transit"
        ElseIf InStr(1, mode, "bus", vbTextCompare) <> 0 Then
            url = url & "transit" & "&transit_mode=bus"
        ElseIf InStr(1, mode, "train", vbTextCompare) <> 0 Then
            url = url & "transit" & "&transit_mode=train"
        ElseIf InStr(1, mode, "tram", vbTextCompare) <> 0 Then
            url = url & "transit" & "&transit_mode=tram"
        ElseIf InStr(1, mode, "rail", vbTextCompare) <> 0 Then
            url = url & "transit" & "&transit_mode=rail"
        Else
            url = url & mode
        End If
    End If
    
    url = addSignature(url) 'Add key
    
    Debug.Print url
    'Send request
    Dim cb As cBrowser: Set cb = New cBrowser
    Dim sWire As String: sWire = cb.httpGET(url): cb.tearDown
    Dim domDoc As DOMDocument60: Set domDoc = New DOMDocument60: domDoc.LoadXML sWire
    Set getXMLResponse = domDoc
    
End Function

'Exports a route to GeoJSON with decoded coordinates
Public Function exportRouteGeoJSON(origin As String, destination As String, _
                                    Optional avoidTolls As Boolean = False, Optional ByVal avoidHighways As Boolean = False, _
                                    Optional mode As String = "DRIVING", _
                                    Optional waypoints As Variant, Optional optimize As Boolean = False, _
                                    Optional outputName As String = "NOSPEC", Optional routeID As Variant = "") _
                                    As Variant

    Dim DEBUGMODE As Boolean: DEBUGMODE = True
    
    'Make the request with the user input and return the XML response
    Dim serverResponse As DOMDocument60: Set serverResponse = getXMLResponse(origin, destination, avoidTolls, avoidHighways, mode, waypoints, optimize)
    Dim status As String: status = handleStatus(serverResponse)
    
    If Not strEquals(status, "OK") Then GoTo ErrorReport
    
    'If the server says OK then go on..
    If Not strEquals(status, "OK") Then exportRouteGeoJSON = status: Exit Function
 
    'Check for target folder
    Dim folderPath As String
    If Not DEBUGMODE Then
        folderPath = createFolderIfNotExists(ActiveWorkbook.path, "jsonPaths")
    End If
    
    'Get the leg node
    Dim legs As IXMLDOMNodeList: Set legs = serverResponse.SelectSingleNode("//route").SelectNodes("//leg")

    Dim startloc As String: startloc = getXMLNodeImmediateChild("start_address", legs(0)).text
    Dim endloc As String: endloc = getXMLNodeImmediateChild("end_address", legs(legs.Length - 1)).text
       
    'Create GeoJSON
    Dim json As New cGeoJson: Set json = New cGeoJson
    'Init file
    json.newGeoJSON
    
    'Make scheme
    Dim fields() As Variant: fields = Array("RouteID", "StepID", "Origin", "Destination", "Distance", "TravelTime", "TravelMode", "Toll", "Instructions", _
                                            "InputOrigin", "InputDestination", "AvoidTolls", "AvoidHighways", "Waypoints")
    
    'Populate JSON with the fields and init other stuff inside
    json.begin (fields)
    Dim lindex As Integer: lindex = 0
    Dim leg As IXMLDOMNode: Set leg = serverResponse.SelectSingleNode("//route").FirstChild
    Do While Not leg Is Nothing
        If Not strEquals(leg.nodeName, "leg") Then GoTo NextLeg 'Make sure this is a leg
        'Get the steps
        Dim step As IXMLDOMNode: Set step = leg.FirstChild
        
        Do While Not step Is Nothing
            If Not strEquals(step.nodeName, "step") Then GoTo NextLeg 'Make sure this is a step
            lindex = lindex + 1
    
            Dim travelMode As String, instructions As String, encodedPolyline As String, travelTime As String, distance As String
            Dim childNode As IXMLDOMElement, subChild As IXMLDOMElement
            
            'Loop through all the children of <step> and save the wanted information
            For Each childNode In step.ChildNodes
                Dim cName As String: cName = childNode.tagName
                If strEquals(cName, "travel_mode") Then
                    travelMode = childNode.text 'TravelMode
                ElseIf strEquals(cName, "html_instructions") Then
                    instructions = formatInstructions(childNode.text) 'Instructions
                ElseIf strEquals(cName, "polyline") Then
                    encodedPolyline = childNode.text 'The actual route coordinates
                ElseIf strEquals(cName, "duration") Then
                    For Each subChild In childNode.ChildNodes
                        If strEquals(subChild.tagName, "value") Then travelTime = subChild.text 'Travel time of the step
                       ' travelTime = Round(travelTime / 60, 2)
                    Next subChild
                ElseIf strEquals(cName, "distance") Then
                    For Each subChild In childNode.ChildNodes
                        If strEquals(subChild.tagName, "value") Then distance = subChild.text 'Travel distance of the step
                        'distance = Round(distance / 1000, 3)
                    Next subChild
                End If
            Next childNode
            
            'Figure out if the road has tolls by looking at the instructions
            Dim toll As String: toll = analyseTollInstructions(instructions)
            
            'Decode the coordinates and return a collection
            Dim coords As Collection: Set coords = decodePolyline(encodedPolyline)
            
            'If there are no waypoints, write "none" to the output, else write a string with the waypoints
            If IsMissing(waypoints) Then
                waypoints = "None"
            ElseIf UBound(Split(waypoints, "|")) > 1 Then
                'Do nothing with it
            ElseIf TypeOf waypoints Is Range Then
                waypoints = "": Dim d As Variant
                For Each d In waypoints
                    waypoints = waypoints & "|" & d
                Next d
            End If
            
            'Assign an ID if none was provided
            If strEquals(routeID, "") Then routeID = countFilesInFolder(folderPath) + 1
            
            'Create array with the values. The order must be the same as the declared fields above
            Dim vals As Variant: vals = Array(routeID, lindex, startloc, endloc, distance, travelTime, travelMode, _
                                                toll, instructions, origin, destination, avoidTolls, avoidHighways, waypoints)
                    
            'Add feature to the json
            json.addFeature vals, coords
            
NextStep:
            Set step = step.NextSibling
        Loop
NextLeg:
        Set leg = leg.NextSibling
    Loop
                    
    'Prepare the output name
    'If the user specified a name, look for illegal chars in there and replace them
    If Not strEquals(outputName, "NOSPEC") Then 'This condition is true if the user didn't specify any parameter #justvbathings
        outputName = removeAcentuation(outputName)
        outputName = replaceSpecialChars(outputName)
        
    'If not, create a meaningfull name with the start and end locations
    Else
        outputName = formatFileName(startloc, endloc)
    End If
        
    'Save and close the json file
    json.finish outputName, folderPath
    
    'Function output
    Dim result(1, 1 To 4) As Variant
    
    Dim finalDist As Double, finalTime As Double
        
    Set legs = serverResponse.SelectNodes("//leg")
        
    For Each leg In legs
        Dim legNode As IXMLDOMNode: Set legNode = leg 'Casting
        
        Dim distanceNode As IXMLDOMNode: Set distanceNode = getXMLNodeImmediateChild("distance", legNode)
        Dim timeNode As IXMLDOMNode: Set timeNode = getXMLNodeImmediateChild("duration", legNode)
                     
        Dim dist As Double: dist = Round(getXMLNodeImmediateChild("value", distanceNode).nodeTypedValue / 1000, 3)
        Dim time As Double: time = Round(getXMLNodeImmediateChild("value", timeNode).nodeTypedValue / 60, 2)
                    
        finalDist = finalDist + dist
        finalTime = finalTime + time
    Next leg
    
    result(0, 1) = finalDist
    result(0, 2) = finalTime
    result(0, 3) = startloc
    result(0, 4) = endloc
    
    Set serverResponse = Nothing
    exportRouteGeoJSON = result
    
ErrorReport:
    exportRouteGeoJSON = status
    
End Function

'Investigates if this step has tolls, by looking for the word "toll" in the html_instructions node
Private Function analyseTollInstructions(htmlInstructions As String) As String
    Dim lcased As String: lcased = LCase(htmlInstructions)
    
    If InStr(1, lcased, "toll", vbTextCompare) <> 0 Or InStr(1, lcased, "portagem", vbTextCompare) <> 0 Then
        analyseTollInstructions = "Yes"
    Else
        analyseTollInstructions = "No"
    End If
End Function

'Transforms the string returned by the API in a collection of readable GeoCoordinates
Public Function decodePolyline(encodedString As String) As Collection
    
    Dim coords As Collection: Set coords = New Collection

    Dim index As Long: index = 0
    
    Dim lat As Long, lng As Long: lat = 0: lng = 0
    
    While (index < Len(encodedString))
    
        lat = lat + decodeNextOffset(encodedString, index)
        lng = lng + decodeNextOffset(encodedString, index)
        
        Dim coord As cGeoCoordinate
        Set coord = New cGeoCoordinate
        coord.latitude = lat / 100000: coord.longitude = lng / 100000
        coords.add coord, str(index)
        
    Wend
        
    Set decodePolyline = coords
End Function
'Removes <hmtl> tags
Private Function formatInstructions(instructions As String)
    Dim RegEx As Object
    Set RegEx = CreateObject("vbscript.regexp")
    
    With RegEx
        .Global = True
        .pattern = "<[^>]+>"
    End With
    
    formatInstructions = RegEx.Replace(instructions, "")
End Function

'Decodes the coordinates string returned by google
Private Function decodeNextOffset(encodedString As String, ByRef index As Long) As Long

    Dim b As Long, shift As Long, result As Long
    b = 0: shift = 0: result = 0
    
    Do
        index = index + 1
        
        b = asc(Mid(encodedString, index, 1)) - 63
        result = result Or ((b And 31) * (2 ^ shift))
        
        shift = shift + 5
    Loop While (b >= 32)
    
    Dim dCoord As Long
    If (result And 1) <> 0 Then
        dCoord = Not Int(result / 2)
    Else
        dCoord = Int(result / 2)
    End If
    
    decodeNextOffset = dCoord
End Function

Private Function removeAcentuation(str As String) As String
    str = Replace(str, ",", "")
    str = Replace(str, ".", "")
    str = Replace(str, "-", "")
removeAcentuation = str
End Function

Private Function formatFileName(ByVal startloc As String, ByVal endloc As String) As String
    Dim str As String
    str = Replace("Route " & removeAcentuation(startloc) & " To " & removeAcentuation(endloc), " ", "")
    str = Replace(str, "Portugal", "")
    str = replaceSpecialChars(str)
    formatFileName = str
End Function

'Exports a route to GeoJSON with decoded coordinates
Public Function requestBUSRouteOD2(origin As String, destination As String, busID As String, linha As String, sentido As String, oStop As String, dStop As String) _
                                    As Variant
    
    'Make the request with the user input and return the XML response
    Dim serverResponse As DOMDocument60: Set serverResponse = getXMLResponse(origin, destination)
    Dim status As String: status = handleStatus(serverResponse)

    'If the server says OK then go on..
    If Not strEquals(status, "OK") Then requestBUSRouteOD2 = status: Exit Function
     
    'Get the leg node
    Dim leg As IXMLDOMNode: Set leg = serverResponse.SelectSingleNode("//route").SelectSingleNode("//leg")

    Dim startloc As String: startloc = leg.SelectSingleNode("//start_address").text
    Dim endloc As String: endloc = leg.SelectSingleNode("//end_address").text
           
    Dim overviewPolyline As String: overviewPolyline = serverResponse.SelectSingleNode("//overview_polyline/points").text
    Dim totalDist As Double: totalDist = serverResponse.SelectSingleNode("//leg/distance/value").text: totalDist = Round(totalDist / 1000, 3)
    Dim totalTime As Double: totalTime = serverResponse.SelectSingleNode("//leg/duration/value").text: totalTime = Round(totalTime / 60, 2)
                                    
    'Create array with the values. The order must be the same as the declared fields above
    Dim vals As Variant: vals = Array(busID, linha, sentido, totalDist, totalTime, origin, destination, oStop, dStop, overviewPolyline)
        
    Set serverResponse = Nothing
    requestBUSRouteOD2 = vals

End Function


Function geoCode(place As String, Optional region As String = "PT") As Variant

place = formatKeyword(place)

Dim urlRequest As String: urlRequest = "http://maps.googleapis.com/maps/api/geocode/xml?" & "address=" & place & "&sensor=false" & "&region=" & region


'Send request
Dim cb As cBrowser: Set cb = New cBrowser
Dim sWire As String: sWire = cb.httpGET(urlRequest): cb.tearDown

Dim Response As DOMDocument60: Set Response = New DOMDocument60: Response.LoadXML sWire
Dim statusNode As IXMLDOMNode: Set statusNode = Response.SelectSingleNode("//status")

Select Case UCase(statusNode.text)
    Case "OK"
        
        Dim latitude As Double: latitude = Response.SelectSingleNode("//result/geometry/location/lat").text
        Dim longitude As Double: longitude = Response.SelectSingleNode("//result/geometry/location/lng").text
        Dim formattedAdress As String: formattedAdress = Response.SelectSingleNode("//result/formatted_address").text
        
        Dim Results(1, 1 To 2) As Variant
        Results(0, 1) = latitude & "," & longitude
        Results(0, 2) = formattedAdress
        
        geoCode = Results
              
    Case "ZERO_RESULTS"   'The geocode was successful but returned no results.
        geoCode = "ZERO_RESULTS"
        
    Case "OVER_QUERY_LIMIT" 'The requestor has exceeded the limit of 2500 request/day.
        geoCode = "The requestor has exceeded the limit of 2500 request/day."
        
    Case Else
        geoCode = "Error"
End Select

Set statusNode = Nothing

End Function

Function formatKeyword(key As String) As String
    Dim result As String
    If Len(key) > 0 Then
        result = Replace(key, " ", "+") + ","
    Else
        result = ""
    End If
    formatKeyword = result
End Function



