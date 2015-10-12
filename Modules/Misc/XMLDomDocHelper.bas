Attribute VB_Name = "XMLDomDocHelper"

Option Explicit

Public Function newXMLDoc() As MSXML2.DOMDocument60
    Set newXMLDoc = New MSXML2.DOMDocument60
End Function

Public Function saveXMLDoc(folderDest As String, fileName As String, objXML As MSXML2.DOMDocument60)
    objXML.Save folderDest & "\" & fileName & ".xml"
End Function

Public Function addNewXMLElement(elementName As String, objXML As MSXML2.DOMDocument60, _
                                 Optional fatherElement As MSXML2.IXMLDOMElement, _
                                 Optional textValue As String) As MSXML2.IXMLDOMElement
                                
    Dim element As MSXML2.IXMLDOMElement
    Set element = objXML.createElement(elementName)
    
    If Not fatherElement Is Nothing Then
        fatherElement.appendChild element
    Else
        objXML.appendChild element
    End If
    
    If Not IsMissing(textValue) Then
        If Len(textValue) > 0 Then element.text = textValue
    End If
    
    Set addNewXMLElement = element
    
End Function

Public Function addNewXMLAttribute(attrName As String, attrValue As Variant, _
                                    fatherElement As MSXML2.IXMLDOMElement, _
                                    objXML As MSXML2.DOMDocument60) As MSXML2.IXMLDOMAttribute
    Dim attr As MSXML2.IXMLDOMAttribute
    Set attr = objXML.createAttribute(attrName)
    attr.value = attrValue
    fatherElement.setAttributeNode attr
    Set addNewXMLAttribute = attr
End Function

Public Function getXMLNodeImmediateChild(nodeName As String, fatherNode As IXMLDOMNode) As IXMLDOMNode
    Dim child As IXMLDOMNode
    Set child = fatherNode.FirstChild
    
    Do While Not child Is Nothing
        If strEquals(child.nodeName, nodeName) Then
            Set getXMLNodeImmediateChild = child
            Exit Function
        End If
        Set child = child.NextSibling
    Loop
End Function
            



