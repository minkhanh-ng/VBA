Attribute VB_Name = "M_DrawingPopulate"
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swActiveModel As SldWorks.ModelDoc2
Dim swComp As SldWorks.Component2
Dim swConf As SldWorks.Configuration
Dim swRootComp As SldWorks.Component2
Dim swDraw As SldWorks.DrawingDoc
Dim swModelDocExt As SldWorks.ModelDocExtension
Dim swSelectionMgr As SldWorks.SelectionMgr

Dim swDocSpecification  As SldWorks.DocumentSpecification
Dim componentsArray(0) As String
Dim components As Variant
Dim name As String
Dim errors As Long
Dim warnings As Long

Dim searchFolders As String
Dim bStatus As Boolean

Dim swModelFileName As String
Dim swPackAndGo As SldWorks.PackAndGo
Dim status As Boolean
Dim statuses As Variant
Dim i As Long

Dim arPackAndGoFileNames As Variant
Dim arPackAndGoFileStatus As Variant

Dim myFileName As String
Dim myExtension As String
Dim SaveName As String

Dim currentPath As String
Dim swSampleModelPath As String
Dim swModelPath As String

Dim longResponse As Long

Sub S_DrawingPopulate(usedRange, sOriginModel As String)

    Dim newFolderPath As String
    Dim folderPath As String
    Dim Cell As Range
'    Dim usedRange As String
    Dim currentModelParentPath As String
    Dim targetPath As String
    Dim sCopiedModelPath As String
    Dim sReplaceModelParentPath As String
    Dim sReplaceModelPath As String
    
    Dim sConfiguration As String
'    Dim sOriginModel As String
    
'    usedRange = "A23:A23"
'    sOriginModel = "XXX-200"
    
    currentPath = GetLocalWorkbookName(ThisWorkbook.fullName, True)
    
    'Create folders for files
'    For Each Cell In Range(usedRange)
'        newFolderPath = currentPath & "\" & Cell.Value & "\2D FILES"
'        CreateFoldersFromRange newFolderPath
'
'        newFolderPath = currentPath & "\" & Cell.Value & "\2D FILES\PDFs"
'        CreateFoldersFromRange newFolderPath
'    Next Cell

    'Copy Drawings
    swSampleModelPath = currentPath + "\MASTER FILES\ORIGINS\" + sOriginModel + "\2D FILES\" + sOriginModel + "-TANK-GA.SLDDRW"
    bStatus = checkFileExists(swSampleModelPath)

    If (bStatus) Then
        ' Set folder where to save the files
        ' currentModelParentPath = currentPath + "\MASTER FILES\2D FILES\"

        ' Loop through each cell in the range then copy and change names
        'Use variables in the FileCopy statement
        Dim xlobj As Object
        Set xlobj = CreateObject("Scripting.FileSystemObject")

        'object.copyfile,source,destination,file overright(True is default)
        For Each Cell In Range(usedRange)
            targetPath = currentPath & "\" & Cell.Value & "\2D FILES\"
            sCopiedModelPath = targetPath & Cell.Value & "-TANK-GA.SLDDRW"

            xlobj.CopyFile swSampleModelPath, sCopiedModelPath, True
            Debug.Print sCopiedModelPath & " created!"
        Next Cell
        Set xlobj = Nothing
    Else
        longResponse = MsgBox("The file is Read-Only." & Chr(13) & "Do you want to close the file without Saving?", vbCritical + vbYesNo, "FileOpenRebuildSaveClose")
    End If
    
    ' Change sheet view model
    'Set swApp = CreateObject("SldWorks.Application")
    Set swApp = GetObject(, "SldWorks.Application")
    swApp.Visible = True

    For Each Cell In Range(usedRange)
        targetPath = currentPath & "\" & Cell.Value & "\2D FILES\"
        swModelPath = targetPath & Cell.Value & "-TANK-GA.SLDDRW"

        sReplaceModelParentPath = currentPath & "\" & Cell.Value & "\3D FILES\"
        sReplaceModelPath = sReplaceModelParentPath & Cell.Value & "-TANK.SLDASM"

        Set swModel = swApp.OpenDoc6(swModelPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", errors, warnings)
'        Set swModelDocExt = swModel.Extension
        Debug.Print "Open " & swModelPath

        If Not swModel Is Nothing Then
            ' Set the working directory to the document directory
            swApp.SetCurrentWorkingDirectory (Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\")))

            ' Activate the loaded document and prompt for rebuild to use getComponents
            Set swModel = swApp.ActivateDoc3(swModel.GetTitle(), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, errors)
            Debug.Print ("Error code after document activation: " & errors)

            S_ReplaceDrwSheetName sOriginModel, Cell.Value, swModel
            S_ReplaceDrwViewName sOriginModel, Cell.Value, swModel
            
            'swModel.Save2 True
            
            S_ChangeViewReferenceDoc sOriginModel, Cell.Value, sReplaceModelPath
            
            swModel.ForceRebuild3 True
            swModel.Save2 True
      
            swApp.CloseDoc swModelPath
        End If
    Next Cell
End Sub



Sub S_ChangeDrwSheetName(prpName As String, swDraw As SldWorks.DrawingDoc)
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If
    
    Dim swSheet As SldWorks.sheet
    Set swSheet = swModel.GetCurrentSheet
    
    Dim vSheetNames As Variant
    vSheetNames = swDraw.GetSheetNames
    Debug.Print ' bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb '
'    Dim i As Integer
'    For i = 0 To UBound(vSheetNames)
'
'        Set swSheet = swModel.Sheet(vSheetNames(i))
'
'        Dim custPrpViewName As String
'        custPrpViewName = swSheet.CustomPropertyView
'
'        Dim vViews As Variant
'        vViews = swSheet.GetViews()
'
'        Dim swCustPrpView As SldWorks.View
'        Set swCustPrpView = Nothing
'
'        Dim j As Integer
'
'        For j = 0 To UBound(vViews)
'            Dim swView As SldWorks.View
'            Set swView = vViews(j)
'
'            If LCase(swView.name) = LCase(custPrpViewName) Then
'                Set swCustPrpView = swView
'                Exit For
'            End If
'        Next
'
'        If swCustPrpView Is Nothing Then
'            Set swCustPrpView = vViews(0)
'        End If
'
'        If Not swCustPrpView Is Nothing Then
'            If prpValue <> "" Then
'                swSheet.SetName (prpValue)
'            End If
            
'            Dim swRefConfName As String
'            Dim swRefDoc As SldWorks.ModelDoc2
'
'            swRefConfName = swCustPrpView.ReferencedConfiguration
'            Set swRefDoc = swCustPrpView.ReferencedDocument
'
'            If Not swRefDoc Is Nothing Then
'
'                Dim prpValue As String
'
'                prpValue = GetCustomPropertyValue(swRefDoc, swRefConfName, prpName)
'
'                If prpValue <> "" Then
'                    swSheet.SetName (prpValue)
'                End If
'
'            Else
'                MsgBox "Failed to get the model from drawing view. Make sure that the drawing is not lightweight"
'            End If
'        Else
'            MsgBox "Failed to get the view to get property from"
'        End If
        
'    Next
    
End Sub

Sub S_ReplaceDrwSheetName(sOldText, sNewText As String, swModel As SldWorks.ModelDoc2)
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If
    
    Dim swSheet As SldWorks.sheet
    Set swSheet = swModel.GetCurrentSheet
    
    Dim vSheetNames As Variant
    vSheetNames = swDraw.GetSheetNames
    
    Dim sNewSheetName As String
    Dim i As Integer
    For i = 0 To UBound(vSheetNames)
    
        Set swSheet = swModel.sheet(vSheetNames(i))
        sNewSheetName = Replace(vSheetNames(i), sOldText, sNewText)
        swSheet.SetName (sNewSheetName)
        Debug.Print "Sheet name " & vSheetNames(i) & " is changed to " & sNewSheetName
        
    Next
    
End Sub

Sub S_ReplaceDrwViewName(sOldText, sNewText As String, swModel As SldWorks.ModelDoc2)
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If

    Dim vSheets As Variant
    vSheets = swDraw.GetViews
    
    Dim i As Integer

    For i = 0 To UBound(vSheets)
        Dim nextViewIndex As Integer
        nextViewIndex = 0
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.view
        
        Set swSheetView = vViews(0)
        
        Dim j As Integer
        
        For j = 1 To UBound(vViews)
            
            Dim swView As SldWorks.view
            Set swView = vViews(j)
            
            Dim viewType As Integer
            viewType = swView.Type
            
            If viewType <> swDrawingViewTypes_e.swDrawingDetailView And viewType <> swDrawingViewTypes_e.swDrawingSectionView Then
                
                nextViewIndex = nextViewIndex + 1
                
                Dim newViewName As String
                newViewName = Replace(swView.name, sOldText, sNewText)
                
                If False = swView.SetName2(newViewName) Then
                    Debug.Print "Failed to rename " & swView.name & " to " & ""
                Else
                    Debug.Print "View name " & swView.name & " is changed to " & newViewName
                End If
            End If
            
        Next
        
    Next
    
End Sub

Sub S_ChangeViewReferenceDoc(sOriginModel, sTankModel, sReplaceModelPath As String)
'    Dim swApp As SldWorks.SldWorks

    Dim swActiveModel As SldWorks.ModelDoc2
    
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    
    Set swActiveModel = swApp.ActiveDoc
    
    Dim swDrawingDoc As SldWorks.DrawingDoc
    Dim bStatus As Boolean
    
    Set swDrawingDoc = swActiveModel
    
    If swActiveModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If
    
    
    If swDrawingDoc Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If

    Set swModelDocExt = swActiveModel.Extension
    Dim vSheets As Variant
    vSheets = swDrawingDoc.GetViews
    
    Dim i As Integer
    Dim vTankHeight As Variant
    
    vTankHeight = Array("6400") 'Array("2200", "2900", "3600", "4300", "5000", "5700", "6400")
    Dim sTankHeight As String
    Dim sTankConfigurationName As String
    
    Dim iRefModel_ID As Integer
    iRefModel_ID = 0
    
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.view
        
        Set swSheetView = vViews(0)
        
        Dim j As Integer
        
        Dim nextViewIndex As Integer
        
        nextViewIndex = 0
        sTankHeight = vTankHeight(i)
        sTankConfigurationName = sTankModel & "-" & sTankHeight & "-B-ZN"
        
        
        For j = 1 To UBound(vViews)
            
            Dim swView As SldWorks.view
            Set swView = vViews(j)
            
            Dim viewType As Integer
            viewType = swView.Type
            
            Dim swSelectedView As SldWorks.view
            Dim swDrawingComponent As SldWorks.DrawingComponent
            Dim views(0) As Object
            Dim instances(0) As Object
            
            If viewType <> swDrawingViewTypes_e.swDrawingDetailView And viewType <> swDrawingViewTypes_e.swDrawingSectionView Then
                
                nextViewIndex = nextViewIndex + 1
                
                Dim sSelectedViewName As String
'                newViewName = Replace(swView.name, sOldText, sNewText)
                'Select the view in which to replace the model
                bStatus = swModelDocExt.SelectByID2(swView.name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
                Set swSelectionMgr = swActiveModel.SelectionManager
                Set swSelectedView = swSelectionMgr.GetSelectedObject6(1, -1)
                Set views(0) = swSelectedView
'                status = swModelDocExt.SelectByID2(swView.name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
                
                sSelectedViewName = swSelectedView.name
                Debug.Print "SelectedView: " & sSelectedViewName
                'Get current referenced model of dwg view
                Dim swDrawModel As SldWorks.ModelDoc2
                Dim sDrawModelName As String
                Dim sDrawModelPath As String
                Dim sDrawModelNameAndID As String
                
                Set swDrawModel = swSelectedView.ReferencedDocument
                
                sDrawModelPath = swDrawModel.GetPathName
                sDrawModelName = Mid(sDrawModelPath, InStrRev(sDrawModelPath, "\") + 1, InStrRev(sDrawModelPath, ".") - InStrRev(sDrawModelPath, "\") - 1)
                Debug.Print "Referenced model path = " & sDrawModelPath
                Debug.Print "Referenced model name = " & sDrawModelName

''' Get component to pass to the instance() to replace it in view
''' In my case use this
                If (sOriginModel = "XXX-40") Then
                
                    If (i = 0 And j = 2) Then
                        iRefModel_ID = 3
                    Else
                        iRefModel_ID = iRefModel_ID + 1
                    End If
                
                ElseIf (sOriginModel = "XXX-200") Then
                    
                    ' No apply of ID change
                    iRefModel_ID = iRefModel_ID + 1
                
                End If
                                              
                sDrawModelNameAndID = sDrawModelName & "-" & iRefModel_ID
                Debug.Print "Referenced model name-ID = " & sDrawModelNameAndID

''' In more common cases use this
                'Get visible components in DrawingView - Used in more common case
'                Dim vComponents As Variant
'                vComponents = swSelectedView.GetVisibleComponents
'                Dim swComp As SldWorks.Component2
'                Set swComp = vComponents(0)
'                Debug.Print "Component 0 = " & swComp.Name2
'                Left ..... get the desired component

                'Select the instance of the model to replace
                bStatus = swModelDocExt.SelectByID2(sDrawModelNameAndID & "@" & sSelectedViewName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
                Set swDrawingComponent = swSelectionMgr.GetSelectedObject6(1, -1)
                Set instances(0) = swDrawingComponent.Component
                
                If (InStrRev(swDrawModel.GetPathName, sReplaceModelPath) = 0) Then
                    bStatus = swDrawingDoc.ReplaceViewModel(sReplaceModelPath, (views), (instances))
                    Debug.Print "Referenced model path = " & swDrawModel.GetPathName
                    
                    swSelectedView.ReferencedConfiguration = sTankConfigurationName
                    Debug.Print "View " & swSelectedView.name & "'s model'S configuration is changed to " & swSelectedView.ReferencedConfiguration
                    If (j = 1) Then
                        swSelectedView.DisplayState = "HALF ROOF HIDE"
                    ElseIf (j = 2) Then
                        swSelectedView.DisplayState = "Display State-1"
                    End If
                    'bStatus = swDrawingDoc.ReplaceViewModel(sReplaceModelPath, (views), 0)
                Else
                    Debug.Print "Error! Model is already referenced in drawing view."
                End If
                
                If (bStatus) Then
                    Debug.Print "View " & swSelectedView.name & "'s model is changed to " & sReplaceModelPath
                Else
                    Debug.Print "Error! Failed to change view " & swView.name & "'s model to " & sReplaceModelPath
                End If
            End If
            
            swActiveModel.ClearSelection
        Next
    Next
    
End Sub


Sub S_ChangeDrwViewName(sOldText, sNewText As String, swModel As SldWorks.ModelDoc2)
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If

    Dim vSheets As Variant
    vSheets = swDraw.GetViews
    
    Dim i As Integer

    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.view
        
        Set swSheetView = vViews(0)
        
        Dim j As Integer
        
        Dim nextViewIndex As Integer
        nextViewIndex = 0
        
        For j = 1 To UBound(vViews)
            
            Dim swView As SldWorks.view
            Set swView = vViews(j)
            
            Dim viewType As Integer
            viewType = swView.Type
            
            If viewType <> swDrawingViewTypes_e.swDrawingDetailView And viewType <> swDrawingViewTypes_e.swDrawingSectionView Then
                
                nextViewIndex = nextViewIndex + 1
                
                Dim newViewName As String
                newViewName = swSheetView.name & "(" & nextViewIndex & ")"
                
                If False = swView.SetName2(newViewName) Then
                    Err.Raise vbError, "", "Failed to rename " & swView.name & " to " & ""
                End If
            End If
            
        Next
        
    Next
    
End Sub




