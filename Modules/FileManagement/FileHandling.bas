Attribute VB_Name = "FileHandling"
Option Explicit

'Creates a folder if doesn't exist and returns the path to the new or existing folder
Public Function createFolderIfNotExists(path As String, folderName As String) As String
    If Len(Dir(path & "\" & folderName, vbDirectory)) = 0 Then MkDir (path & "\" & folderName)
    createFolderIfNotExists = path & "\" & folderName
End Function

'Creates a new Workbook if doesn't exist and returns the path to the new or existing workbook
Public Function createWorkbookIfNotExists(filePath As String, workbookName As String) As String
    If Len(Dir(filePath & "\" & workbookName & ".xlsx", vbDirectory)) = 0 Then
        Dim newWB As Workbook: Set newWB = Workbooks.add(xlWBATWorksheet)
        newWB.SaveAs filePath & "\" & workbookName, 51 'filetype xlsx: 51; xls: -4143
        newWB.Close
    End If
    createWorkbookIfNotExists = filePath & "\" & workbookName & ".xlsx"
End Function

''Create new worksheet with the process's name
Private Function createNewSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim newWS As Worksheet
    Set newWS = wb.Sheets.add()
    newWS.name = sheetName
    Set createNewSheet = newWS
End Function

'Returns True if the folder is found
Public Function folderExists(folderDir As String) As Boolean
    If Len(Dir(folderDir, vbDirectory)) <> 0 Then folderExists = True
End Function
'Returns True if the file is found
Public Function fileExists(fileDir As String) As Boolean
    If Len(Dir(fileDir)) <> 0 Then fileExists = True
End Function

'Creates and returns log file object
Public Function newLogFile(fileName As String, filePath As String) As Object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set newLogFile = fso.CreateTextFile(filePath & "\" & fileName & _
                                            year(Now) & month(Now) & day(Now) & " " & _
                                            hour(Now) & "h" & Minute(Now) & "m" & ".txt", False)
End Function

'Returns the number of files in a folder
Public Function countFilesInFolder(path As String) As Integer
    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject").getFolder(path)
    countFilesInFolder = objFSO.Files.Count
    Set objFSO = Nothing
End Function
