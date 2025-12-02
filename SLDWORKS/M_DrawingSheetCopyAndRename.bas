Attribute VB_Name = "M_DrawingSheetCopyAndRename"
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

Sub S_SheetCopyAndRename(usedRange As String)

    Dim newFolderPath As String
    Dim folderPath As String
    Dim Cell As Range
    Dim CellConfig As Range
'    Dim usedRange As String
    Dim currentModelParentPath As String
    Dim targetPath As String
    Dim sCopiedModelPath As String
    Dim sReplaceModelParentPath As String
    Dim sReplaceModelPath As String
    Dim sConfigRange As String
    
    Dim sSheetNameTobeCopy As String
    Dim sSheetNameOfNewCopy As String
    Dim sCurrentSheetName As String
    
    Dim sConfiguration As String
    Dim sOriginModel As String
    Dim sOriginConfig As String
    
'    usedRange = "A9:A9"
    sOriginModel = ""
    sConfigRange = "D2:D7"
    sOriginConfig = "D8"
    
    currentPath = GetLocalWorkbookName(ThisWorkbook.fullName, True)
     
    ' Change sheet view model
    'Set swApp = CreateObject("SldWorks.Application")
    Set swApp = GetObject(, "SldWorks.Application")
    swApp.Visible = True

    For Each Cell In Range(usedRange)
        'targetPath = currentPath & "\" & Cell.Value & "\2D FILES\"
        targetPath = currentPath & "\" & "GA DRAWINGS\"
        swModelPath = targetPath & Cell.Value & "-TANK-GA.SLDDRW"

'        sReplaceModelParentPath = currentPath & "\" & cell.Value & "\3D FILES\"
'        sReplaceModelPath = sReplaceModelParentPath & cell.Value & "-TANK.SLDASM"

        Set swModel = swApp.OpenDoc6(swModelPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", errors, warnings)
'        Set swModelDocExt = swModel.Extension
        Debug.Print "Open " & swModelPath

        If Not swModel Is Nothing Then
            ' Set the working directory to the document directory
            swApp.SetCurrentWorkingDirectory (Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\")))

            ' Activate the loaded document and prompt for rebuild to use getComponents
            Set swModel = swApp.ActivateDoc3(swModel.GetTitle(), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, errors)
            Debug.Print ("Error code after document activation: " & errors)
            sSheetNameTobeCopy = Cell.Value & "-" & Range(sOriginConfig).Value & "-" & "B"
            Set swDraw = swModel
            
            ' Copy sheet and change the name following the first sheet for 6400
            ' Active sheet xx-6400-xx
            For Each CellConfig In Range(sConfigRange)
                'Copy sheet to new sheet
                sCurrentSheetName = F_DrawingSheetCopy(sSheetNameTobeCopy)
                'Change name of new sheet to match config
                sSheetNameOfNewCopy = Cell.Value & "-" & CellConfig.Value & "-" & "B"
                sSheetNameTobeCopy = F_ReplaceDrwSheetName(sCurrentSheetName, sCurrentSheetName, sSheetNameOfNewCopy)
            Next CellConfig
                             
            ' Change view name by sheet name
            S_MassReplaceDrwViewNameBySheetName
            S_ChangeViewConfigurationBySheetName
            
            Call M_ViewScaleProcess.ViewScaleProcess(RePositionBaseOnSeedSheet, "", "")
            M_ViewSequentialLabeling.S_ViewSequentialLabeling
            
            'S_ReplaceDrwSheetName sOriginModel, cell.Value, swModel
            'S_ReplaceDrwViewName sOriginModel, cell.Value, swModel
            
            'swModel.Save2 True
            
'            S_ChangeViewReferenceDoc sOriginModel, cell.Value, sReplaceModelPath
            
            swModel.ForceRebuild3 True
            swModel.Save2 True
      
            'swApp.CloseDoc swModelPath
        End If
    Next Cell
End Sub



Function F_DrawingSheetCopy(sSheetNameToCopy As String) As String
    
    Dim bStatus As Boolean
    Dim swDraw As SldWorks.DrawingDoc
    Dim swModel As SldWorks.ModelDoc2
    
    Set swApp = GetObject(, "SldWorks.Application")
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    
    If (swDraw Is Nothing) Then
        MsgBox " Please open a drawing document. "
        End
    End If
    Dim currentsheet As sheet
    'Active tobe-copied sheet
    swDraw.ActivateSheet sSheetNameToCopy
    
    Set currentsheet = swDraw.GetCurrentSheet
    swDraw.ActivateSheet (currentsheet.GetName)
    Debug.Print "Active sheet: " & currentsheet.GetName
    
    bStatus = swDraw.Extension.SelectByID2(sSheetNameToCopy, "SHEET", 0.09205356547875, 0.10872368523, 0, False, 0, Nothing, 0)
    swModel.EditCopy
    bStatus = swDraw.PasteSheet(swInsertOptions_e.swInsertOption_BeforeSelectedSheet, swRenameOptions_e.swRenameOption_Yes)
    
    Set currentsheet = swDraw.GetCurrentSheet
    Debug.Print "Sheet has been copied to: " & currentsheet.GetName
    F_DrawingSheetCopy = currentsheet.GetName
    
End Function


Function F_ReplaceDrwSheetName(sSheetName, sOldText, sNewText As String) As String
    
    Dim bStatus As Boolean
    Dim sNewSheetName As String
    
    Dim swApp As SldWorks.SldWorks
    Dim swDraw As SldWorks.DrawingDoc
    Dim swModel As SldWorks.ModelDoc2
    
    Set swApp = GetObject(, "SldWorks.Application")
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If
    
    Dim swSheet As SldWorks.sheet
      
    Set swSheet = swModel.sheet(sSheetName)
        sNewSheetName = Replace(sSheetName, sOldText, sNewText)
        swSheet.SetName (sNewSheetName)
    Debug.Print "Sheet name " & sSheetName & " is changed to " & sNewSheetName
        
    F_ReplaceDrwSheetName = sNewSheetName
    
End Function


Function F_ReplaceDrwViewName(sOldText, sNewText As String, swModel As SldWorks.ModelDoc2)
    
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If

    Dim vSheets As Variant
    vSheets = swDraw.GetViews       'vSheets(0)(0) = Sheet1 / vSheets(0)(1) = S1_View1 / vSheets(0)(2) = S1_View2 / vSheets(1)(0) = Sheet2
    
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
    
End Function

Function F_MassReplaceDrwSheetName(ByVal sOldText, ByVal sNewText As String)
    
    Dim bStatus As Boolean
    
    Dim swApp As SldWorks.SldWorks
    Dim swDraw As SldWorks.DrawingDoc
    Dim swModel As SldWorks.ModelDoc2
    
    Set swApp = GetObject(, "SldWorks.Application")
    Set swModel = swApp.ActiveDoc
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
    
End Function

Sub S_MassReplaceDrwViewNameBySheetName()
    
    Dim bStatus As Boolean
    
    Dim swApp As SldWorks.SldWorks
    Dim swDraw As SldWorks.DrawingDoc
    Dim swModel As SldWorks.ModelDoc2
    
    Set swApp = GetObject(, "SldWorks.Application")
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If

    Dim vSheets As Variant
    'Return array of sheets, consist of array of its views
    vSheets = swDraw.GetViews 'vSheets(0)(0) = Sheet1 / vSheets(0)(1) = S1_View1 / vSheets(0)(2) = S1_View2 / vSheets(1)(0) = Sheet2
    
    Dim i As Integer

    For i = 0 To UBound(vSheets)
        Dim nextViewIndex As Integer
        nextViewIndex = 0

        Dim vArrayViewsAndSheets As Variant
        vArrayViewsAndSheets = vSheets(i)

        Dim swSheet As SldWorks.view

        Set swSheet = vArrayViewsAndSheets(0) 'Each view0 is sheet itself
        
        Dim sSheetName As String
        sSheetName = swSheet.GetName2()

        Dim j As Integer

        For j = 1 To UBound(vArrayViewsAndSheets)

            Dim swView As SldWorks.view
            Set swView = vArrayViewsAndSheets(j)

            Dim viewType As Integer
            viewType = swView.Type

            If viewType <> swDrawingViewTypes_e.swDrawingDetailView And viewType <> swDrawingViewTypes_e.swDrawingSectionView Then

                nextViewIndex = nextViewIndex + 1

                Dim newViewName As String

                If (j = 1) Then
                    newViewName = sSheetName + " GA VIEW"
                ElseIf (j = 2) Then
                    newViewName = sSheetName + " F VIEW"
                End If

                If False = swView.SetName2(newViewName) Then
                    Debug.Print "Failed to rename " & swView.name & " to " & ""
                Else
                    Debug.Print "View name " & swView.name & " is changed to " & newViewName
                End If
            End If

        Next

    Next
    
End Sub



Sub S_ChangeViewConfigurationBySheetName()
'    Dim swApp As SldWorks.SldWorks

    Dim swActiveModel As SldWorks.ModelDoc2
    
    Set swApp = GetObject(, "SldWorks.Application")
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
    Dim sTankConfigurationName As String
    
    For i = 0 To UBound(vSheets)
        
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.view
        
        Set swSheetView = vViews(0)
        
        Dim j As Integer
        
        Dim nextViewIndex As Integer
        
        nextViewIndex = 0
        
        Dim sSheetName As String
        sSheetName = swSheetView.GetName2()
        
        sTankConfigurationName = sSheetName & "-" & "ZN"
        For j = 1 To UBound(vViews)
            
            Dim swView As SldWorks.view
            Set swView = vViews(j)
            
            Dim viewType As Integer
            viewType = swView.Type
            
            Dim swSelectedView As SldWorks.view
            Dim views(0) As Object
            
            If viewType <> swDrawingViewTypes_e.swDrawingDetailView And viewType <> swDrawingViewTypes_e.swDrawingSectionView Then
                
                nextViewIndex = nextViewIndex + 1
                
                Dim sSelectedViewName As String
                'Select the view in which to replace the model
                bStatus = swModelDocExt.SelectByID2(swView.name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
                Set swSelectionMgr = swActiveModel.SelectionManager
                Set swSelectedView = swSelectionMgr.GetSelectedObject6(1, -1)
                Set views(0) = swSelectedView
                
                sSelectedViewName = swSelectedView.name
                Debug.Print "SelectedView: " & sSelectedViewName

                'Replace view's configuration
                swSelectedView.ReferencedConfiguration = sTankConfigurationName
                Debug.Print "View " & swSelectedView.name & "'s model'S configuration is changed to " & swSelectedView.ReferencedConfiguration
                If (j = 1) Then
                    swSelectedView.DisplayState = "HALF ROOF HIDE"
                    
                ElseIf (j = 2) Then
                    swSelectedView.DisplayState = "Display State-1"
                    
                    Dim sTankHeightDim As String
                    'sTankHeightDim = "RD1" & "@" & sSelectedViewName
                    'bStatus = swModelDocExt.SelectByID2(sTankHeightDim, "DIMENSION", 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
                    
                    Dim swDisplayDimension As DisplayDimension
                    Dim swDimTankHeight As Dimension
                    Set swDisplayDimension = swSelectedView.GetFirstDisplayDimension
                    
                    While Not swDisplayDimension Is Nothing
                        Set swDimTankHeight = swDisplayDimension.GetDimension
                        If (swDimTankHeight.name = "RD1") Then
                            Dim fOverrideHeight As Double
                            fOverrideHeight = ExtractAndConvertToDouble(sSheetName) / 1000
                            'MsgBox "The extracted Double value is: " & fOverrideHeight
                        
                            swDisplayDimension.SetOverride True, fOverrideHeight
                            
                        End If
                        
                        Set swDisplayDimension = swDisplayDimension.GetNext
                    Wend
                    
                End If
            End If
            
            swActiveModel.ClearSelection
        Next
    Next
    
End Sub

Function ExtractAndConvertToDouble(text As String) As Double
    Dim parts() As String
    Dim numericPart As String
    
    ' Split the text by the hyphen character
    parts = Split(text, "-")
    
    ' Extract the third part which is expected to be the numeric value
    numericPart = parts(2)
    
    ' Convert the extracted part to Double
    ExtractAndConvertToDouble = CDbl(numericPart)
End Function
