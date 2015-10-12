Attribute VB_Name = "Translation"
Option Explicit

Public Function translateGoogle(ByVal text As String, ByVal language As String) As String
    Dim i As Variant, result As String
    For i = 0 To 5
        result = translationGoogleRequest(text, language)
        If Len(result) > 0 Then translateGoogle = result: Exit For
    Next i
End Function

Public Function translationGoogleRequest(ByVal text As String, ByVal language As String) As String
    Dim IE As Object: Set IE = CreateObject("InternetExplorer.application"): IE.visible = False
    
    IE.Navigate "http://translate.google.com/#" & "auto" & "/" & language & "/" & text
    Do Until IE.readyState = 4
        DoEvents
    Loop
    
    Application.Wait (Now + TimeValue("0:00:5"))
    
    Do Until IE.readyState = 4
        DoEvents
    Loop
    
    Dim result As Variant
    result = Split(Application.WorksheetFunction.Substitute(IE.Document.getElementById("result_box").innerHTML, "</SPAN>", ""), "<")
    
    Dim j As Variant, result_data As Variant
    For j = LBound(result) To UBound(result)
        result_data = result_data & Right(result(j), Len(result(j)) - InStr(result(j), ">"))
    Next

    IE.Quit
    translationGoogleRequest = result_data

End Function

