Attribute VB_Name = "M_ModelsPopulate"
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swActiveModel As SldWorks.ModelDoc2
Dim swComp As SldWorks.Component2
Dim swConf As SldWorks.Configuration
Dim swRootComp As SldWorks.Component2
Dim swModelDocExt As SldWorks.ModelDocExtension
Dim swPackAndGo As SldWorks.PackAndGo

Dim sExcludePath As String
Dim sCurrentModelParentPath As String
Dim sTargetPath As String

Dim sParentFolderPath As String
Dim swModelPath As String

Dim arPackAndGoFileNames As Variant
Dim arPackAndGoFileStatus As Variant

Dim componentsArray(0) As String
Dim components As Variant
Dim name As String
Dim errors As Long
Dim warnings As Long

Dim searchFolders As String
Dim boolstatus As Boolean

Dim swModelFileName As String

Dim myFileName As String
Dim myExtension As String
Dim SaveName As String

Dim cMsgListener As Class_MsgListener
Dim sTextToReplace As String
    

Sub S_ModelPopulate(sParentFolderPath, sWorkbookPath, sFromPath, sUsedRange, sOriginModelFile, sOriginModel As String)


    
'/////////////////////////////////////////////////
    
'    sCurrentPath = GetLocalWorkbookName(ThisWorkbook.fullName, True)
    
'''Initiate Solidworks
    Set swApp = GetObject(, "SldWorks.Application")
 
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    swApp.Visible = True
        
''' Open the original assembly
    Dim folderPath As String
    Dim Cell As Range
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant
    Dim sFolder As String
    
    swModelPath = sFromPath & "\" & sOriginModelFile
    Set swModel = swApp.OpenDoc6(swModelPath, swDocASSEMBLY, swOpenDocOptions_LoadLightweight, "", errors, warnings)
    Set swModelDocExt = swModel.Extension
    Debug.Print "swModelPath " & swModelPath
    

    ' Active and display the model
    If Not swModel Is Nothing Then

        'Set the working directory to the document directory
        sCurrentModelParentPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\"))
        
        swApp.SetCurrentWorkingDirectory (sCurrentModelParentPath)
        Debug.Print "Current working directory is now " & swApp.GetCurrentWorkingDirectory
        'Activate the loaded document and prompt for rebuild to use getComponents
        Set swModel = swApp.ActivateDoc3(swModel.GetTitle(), False, swRebuildOnActivation_e.swRebuildActiveDoc, errors)
        Debug.Print ("Error code after document activation: " & errors)
        
''' Pack and go all components that's not located in \COMMON PARTS to another XXX-NO folder
        'Get Pack and Go object
        Debug.Print "Pack and Go"
        
        'Set folder where to save the files
        sExcludePath = sParentFolderPath + "\COMMON PARTS"
        
        sTextToReplace = sOriginModel
        
        ' Loop through each cell in the range (excluding the header)
        For Each Cell In Range(sUsedRange)
            sTargetPath = sParentFolderPath & "\" & Cell.Value & "\3D FILES\"
            S_PackAndGoExclude sTextToReplace, Cell.Value
        Next Cell
    End If
    
     swApp.CloseDoc swModelPath
     
End Sub

Sub S_TraverseComponent(swComp As SldWorks.Component2, nLevel As Long)
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

''' PackAndGo and change the model name
Sub S_PackAndGoExclude(sTextToReplace, sNewInNameText As String)
    
    Dim bStatus As Boolean
    Dim vStatuses As Variant
    Dim i As Long
'/////////////////////////////////////////////////////

    Set swPackAndGo = swModelDocExt.GetPackAndGo
    'Include any drawings
    swPackAndGo.IncludeDrawings = True
    
    ' Get file name and extension
'    swModelFileName = Mid(swModelPath, InStrRev(swModelPath, "\") + 1, InStrRev(swModelPath, ".") - InStrRev(swModelPath, "\") - 1)
'    Debug.Print (swModelFileName)
    
    ' Get current paths and filenames of the pack and go assembly documents
    bStatus = swPackAndGo.GetDocumentSaveToNames(arPackAndGoFileNames, arPackAndGoFileStatus)
    Debug.Print ""
    Debug.Print "  Add SOLIDWORKS files' paths and filenames: "
    
    Dim sCurrentPackAndGoFilePath As String
    Dim listOfNewPaths As New Collection
    
    ' Get rid of files that are located in \COMMON PARTS from the Pack and Go list
    If (Not (IsEmpty(arPackAndGoFileNames))) Then
        For i = 0 To UBound(arPackAndGoFileNames)
            sCurrentPackAndGoFilePath = arPackAndGoFileNames(i)
            If InStr(sCurrentPackAndGoFilePath, sExcludePath) <> 0 Then
                sCurrentPackAndGoFilePath = ""
            End If
            
            listOfNewPaths.Add sCurrentPackAndGoFilePath
        Next i
    End If
                  
    Debug.Print "    The path and filename to be PnG is: "
    For i = 1 To listOfNewPaths.count
        arPackAndGoFileNames(i - 1) = Replace(listOfNewPaths(i), sCurrentModelParentPath, sTargetPath)
        arPackAndGoFileNames(i - 1) = Replace(arPackAndGoFileNames(i - 1), sTextToReplace, sNewInNameText)
        
        Debug.Print arPackAndGoFileNames(i - 1)
    Next i
    
    ' Set document paths and names for Pack and Go
    bStatus = swPackAndGo.SetDocumentSaveToNames(arPackAndGoFileNames)
    Debug.Print "Save-to name successful " & bStatus
    vStatuses = swModelDocExt.SavePackAndGo(swPackAndGo)
    
    For i = 1 To UBound(vStatuses)
        Debug.Print "    PnG Status: " & vStatuses(i - 1)
        Debug.Print " "
    Next i

End Sub
