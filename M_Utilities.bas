Attribute VB_Name = "M_Utilities"
Option Explicit

Sub S_CreateFoldersFromRange(newFolderPath As String)
    Dim folderPath As String
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant
    
    Dim sFolder As String
    
    ' Loop through each cell in the range (excluding the header)
    tempFolderPath = ""
    'Split the folder path into individual folder names
    individualFolders = Split(newFolderPath, "\")
    
    For Each arrayElement In individualFolders
        'Build string of folder path
        tempFolderPath = tempFolderPath & arrayElement & "\"
        
        If F_CheckFolderExists(tempFolderPath) = False Then
            MkDir tempFolderPath
        End If
    Next arrayElement
    
    Debug.Print "Created " & newFolderPath
    
    ' Display a message
    ' MsgBox "Folders created based on the specified cell values."
End Sub

Function F_CheckFolderExists(folderPath As String) As Boolean
        
    Dim objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    If objFso.FolderExists(folderPath) Then        ' check if the folder exists.
'        MsgBox "Yes, it exist"
        F_CheckFolderExists = True
    Else
'        MsgBox "No, the folder does not exist"
        F_CheckFolderExists = False
    End If
End Function

Function checkFileExists(filePath As String) As Boolean
        
    Dim objFso As Object
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    If objFso.FileExists(filePath) Then        ' check if the folder exists.
'        MsgBox "Yes, it exist"
        checkFileExists = True
    Else
'        MsgBox "No, the folder does not exist"
        checkFileExists = False
    End If
End Function

Sub TraverseComponent _
(swComp As SldWorks.Component2, nLevel As Long)
    Dim vChildComp As Variant
    Dim swChildComp As SldWorks.Component2
    Dim swCompConfig As SldWorks.Configuration
    Dim sPadStr As String
    Dim i As Long
    For i = 0 To nLevel - 1
        sPadStr = sPadStr + "  "
    Next i
    vChildComp = swComp.GetChildren
    For i = 0 To UBound(vChildComp)
        Set swChildComp = vChildComp(i)
        TraverseComponent swChildComp, nLevel + 1
        Debug.Print sPadStr & swChildComp.Name2 & " <" & swChildComp.ReferencedConfiguration & ">"
    Next i
End Sub

