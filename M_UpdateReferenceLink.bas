Attribute VB_Name = "M_UpdateReferenceLink"
Option Explicit

'Sub UpdateRefAllFiles()
'
'        Dim filpath As String
'        Dim fso As Object
'        Dim fil As Object
'        Dim fldr As Object
'
'        filpath = "C:\Users\khanh.nguyen\OneDrive - xxx\Desktop\PackNGo\E-XXX"
'        Set fso = CreateObject("Scripting.FileSystemObject")
'        Set fldr = fso.GetFolder(filpath)
'
'    For Each fil In fldr.Files
'        With Application.Workbooks.Open(fil.path)
'            .RefreshAll
'            .Close True
'        End With
'    Next fil
'
'
'End Sub

'-------------------------------------------------------


Sub loopAllSubFolderSelectStartDirectory()

Dim FSOLibrary As Object
Dim FSOFolder As Object
Dim folderName As String

'Set the folder name to a variable
folderName = "C:\Users\khanh.nguyen\OneDrive - xxx\Desktop\PackNGo\E-XXX"

'Set the reference to the FSO Library
Set FSOLibrary = CreateObject("Scripting.FileSystemObject")

'Another Macro must call LoopAllSubFolders Macro to start
LoopAllSubFolders FSOLibrary.GetFolder(folderName)

End Sub

Sub LoopAllSubFolders(FSOFolder As Object)

Dim FSOSubFolder As Object
Dim FSOFile As Object

'For each subfolder call the macro
For Each FSOSubFolder In FSOFolder.subfolders
    LoopAllSubFolders FSOSubFolder
Next

Dim wsWorkbook As Workbook
Dim wsWorksheet As Worksheet
'For each file, print the name
For Each FSOFile In FSOFolder.Files

    'Insert the actions to be performed on each file
    'This example will print the full file path to the immediate window
    
    If InStrRev(FSOFile.path, ".xlsx") Then
        Debug.Print FSOFile.path
        
        Set wsWorkbook = Application.Workbooks.Open(FSOFile.path)
       
        wsWorkbook.RefreshAll
        Application.CalculateUntilAsyncQueriesDone
        wsWorkbook.Close True

    End If

Next

End Sub
