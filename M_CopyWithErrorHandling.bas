Attribute VB_Name = "M_CopyWithErrorHandling"
Option Explicit
Dim errors As Collection
Dim file As Object 'Scripting.File
Dim fldr As Object 'Scripting.Folder

Sub CopyFolderWithErrorHandling(FromPath, ToPath1 As String)
    Dim fso As Object 'Scripting.FileSystemObject
    Dim paths As Variant
    Dim path As Variant
'    Dim FromPath, ToPath1$ As String
    Dim i As Long
    Dim ToPath2$, ToPath3$, ToPath4$, ToPath5$, ToPath6$, ToPath7$, ToPath8$, ToPath9$, ToPath10$
    
    '!!!### IMPORTANT ###!!!
    '    Assign all of your "ToPath" variables here:
'    ToPath1 = "c:\some\path"
    'Etc.
    
    Set fso = CreateObject("scripting.filesystemobject")
    Set errors = New Collection
    
'    FromPath = "C:\Debug\" '## Modify as needed
    
    If fso.FolderExists(FromPath) = False Then
        MsgBox FromPath & " doesn't exist"
        Exit Sub
    End If
    
    '## Create an array of destination paths for concise coding
    paths = Array(ToPath1, ToPath2, ToPath3, ToPath4, ToPath5, ToPath6, ToPath7, ToPath8, ToPath9, ToPath10)
    
    '## Ensure each path is well-formed:
    For i = 0 To UBound(paths)
        path = paths(i)
        If Right(path, 1) = "\" Then
            path = Left(path, Len(path) - 1)
        End If
        paths(i) = path
    Next
    
    '## Attempt to delete the destination paths and identify any file locks
    For Each path In paths
        '# This funcitno will attempt to delete each file & subdirectory in the folder path
        Call DeleteFolder(fso, path)
    Next
    
    
    '## If there are no errors, then do the copy:
    If errors.count = 0 Then
        For Each path In paths
            fso.CopyFolder FromPath, path
        Next
    Else:
        '# inform you of errors, you should modify to print a text file...
        Dim str$
    
'        For Each e In errors
'            str = str & e & vbNewLine
'        Next
    
        '## Create an error log on your desktop
        'FSO.CreateTextFile(CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\errors.txt").Write str
    
    End If
    
    Set errors = Nothing
End Sub

Sub DeleteFolder(fso As Object, path As Variant)

    'Check each file in the folder
    For Each file In fso.GetFolder(path).Files
        Call DeleteFile(fso, file)
    Next
    'Check each subdirectory
    For Each fldr In fso.GetFolder(path).subfolders
        Call DeleteFolder(fso, fldr.path)
    Next

End Sub

Sub DeleteFile(fso As Object, file)
    On Error Resume Next
    Kill file.path
    If Err.Number <> 0 Then
        errors.Add file.path
    End If
End Sub

