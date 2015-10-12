Attribute VB_Name = "GitHelper"
Option Explicit


Public Sub ExportSourceFiles()
 
    Dim component As VBComponent
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            Call createFolderIfNotExists(ThisWorkbook.path, ThisWorkbook.VBProject.name)
            Dim fullPath As String: fullPath = ThisWorkbook.path & "\" & ThisWorkbook.VBProject.name & "\" & component.name & ToFileExtension(component.Type)
            component.Export fullPath
    End If
    Next
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String

    Select Case vbeComponentType
        Case vbext_ComponentType.vbext_ct_ClassModule
            ToFileExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_StdModule
            ToFileExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_MSForm
            ToFileExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
        Case vbext_ComponentType.vbext_ct_Document
        Case Else
            ToFileExtension = vbNullString
    End Select
 
End Function
