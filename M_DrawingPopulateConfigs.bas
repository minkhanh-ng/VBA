Attribute VB_Name = "M_DrawingPopulateConfigs"
Private Declare PtrSafe Function _
    CoRegisterMessageFilter Lib "OLE32.DLL" _
    (ByVal lFilterIn As Long, _
    ByRef lPreviousFilter) As Long

Option Explicit

Dim lMsgFilter                                  As Long
Dim swApp                                       As SldWorks.SldWorks
Dim swModelDocExt                               As SldWorks.ModelDocExtension
Dim swSelectionMgr                              As SldWorks.SelectionMgr

Sub S_DrawingMFG_PopulateTruss()
    Dim swModel                                 As SldWorks.ModelDoc2
    Dim swOriginModel                           As SldWorks.ModelDoc2
    Dim swActiveModel                           As SldWorks.ModelDoc2
    Dim swDraw                                  As SldWorks.DrawingDoc
    Dim swDrawOrigin                            As SldWorks.DrawingDoc

    Dim longErrors                              As Long
    Dim longWarnings                            As Long
    Dim longResponse                            As Long

    Dim bStatus                                 As Boolean

    Dim currentPath                             As String
    Dim swSampleModelPath                       As String
    Dim swModelPath                             As String

    Dim newFolderPath                           As String
    Dim folderPath                              As String

'    Dim usedRange As String
    Dim currentModelParentPath                  As String
    Dim targetPath                              As String
    Dim sCopiedModelPath                        As String

    Dim sSheetNameOfNewCopy                     As String
    Dim sCurrentSheetName                       As String
    Dim sNewSheetName                           As String

    Dim sOriginRange, sOriginConfigRange, sOriginModel, sOriginDrwNo, sOriginDrwNumberRange As String
    Dim sConfigRange, sOriginConfig, sNewConfig, sNewDrwNo, sDrwNumberRange                 As String
    Dim sConfigPartnumberRange                                                              As String
    Dim sCoppiedSuffixRange                                                                 As String
    Dim sCoppiedDrwNumberPrefixRange                                                        As String
    Dim sOriginPartNumberRange                                                              As String
    Dim sToReplaceName                                                                      As String
    Dim sCurrentWorkbookPath                    As String
    Dim wbUsedWb                                As Workbook
    Dim wsUsedWs                                As Worksheet
    Dim Cell, CellOrigin, CellOriginDrw         As Range
    Dim CellSuffix, CellConfig                  As Range
    Dim rngOriginDrwNumberRange                 As Range
    Dim rngCoppiedSuffixRange                   As Range
    Dim rngCoppiedDrwNumberPrefixRange          As Range
    Dim rngConfigRange                          As Range
    Dim rngOriginPartNumberRange                As Range
    Dim rngConfigPartnumberRange                As Range

    Dim longLabelCounter                        As Long
    Dim longLabelCounterBuffer                  As Long

    Dim sSuffix                                 As String
    Dim sPrefix                                 As String

    Dim sPrpDrwTitle, sPrpDrwNo                 As String
    Dim sPrpDrwDescription                      As String

    Dim longLastRow                             As Long
    
    Dim vCreatedDrws()                          As Variant
    Dim vOldNames()                             As Variant
    Dim vToChangeConfigs()                      As Variant
    Dim vToReplaceNames()                       As Variant
    Dim longDrws                                As Long
    Dim i                                       As Long
    
    Dim swFeature                               As SldWorks.Feature
    Dim swSelectedFeature                       As SldWorks.Feature
    Dim swWeldmentCutListFeat                   As SldWorks.WeldmentCutListFeature

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Debug.Print "S_DrawingMFG_PopulateTruss"

'''''''''''''''''''''''''''''' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Remove the message filter before calling Subs (remove OLE waiting warning).

    CoRegisterMessageFilter 0&, lMsgFilter
      
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initiate the parameters

    ' Set the path to your Excel file

    'sCurrentWorkbookPath = sWorkbookPath
    'Workbooks.Open FileName:=sCurrentWorkbookPath
    'Set wbUsedWb = Workbooks.Open(sCurrentWorkbookPath)
    Set wbUsedWb = ThisWorkbook
    Set wsUsedWs = wbUsedWb.Sheets("Sheet1")

    sConfigRange = "AG2:AG12"           'Range of part configurations' name
    Set rngConfigRange = wsUsedWs.Range(sConfigRange)
    
    sConfigPartnumberRange = "AJ2:AJ12" 'Range of part configurations' part number
    Set rngConfigPartnumberRange = wsUsedWs.Range(sConfigPartnumberRange)

    sOriginDrwNumberRange = "AN2:AN7"   'Range of part drawings
    Set rngOriginDrwNumberRange = wsUsedWs.Range(sOriginDrwNumberRange)
    
    sCoppiedSuffixRange = "AO2:AO7"     'Range of part drawings suffixes -BC -IFX etc
    Set rngCoppiedSuffixRange = wsUsedWs.Range(sCoppiedSuffixRange)

    sCoppiedDrwNumberPrefixRange = "AM2:AM12"
    Set rngCoppiedDrwNumberPrefixRange = wsUsedWs.Range(sCoppiedDrwNumberPrefixRange)

    sOriginPartNumberRange = "AL2:AL7"
    Set rngOriginPartNumberRange = wsUsedWs.Range(sOriginPartNumberRange)
    
    sPrefix = "PD"
'    sOriginRange = "AG2"
'    Set CellOrigin = wsUsedWs.Range(sOriginRange)(1, 1)
'    sOriginModel = CellOrigin.Value

'    Set Cell = wsUsedWs.Range(sOriginConfigRange)
'    sOriginConfig = CStr(Cell.Value)

    'Current path
    currentPath = GetLocalWorkbookName(ThisWorkbook.fullName, True)

    longLastRow = rngConfigRange.End(xlDown).Row

    longDrws = 0
    ReDim vCreatedDrws(longDrws)
    ReDim vToChangeConfigs(longDrws)
    ReDim vToReplaceNames(longDrws)
    ReDim vOldNames(longDrws)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copy to new drawings and change names

    Debug.Print "Copy to new drawings and change names"

    For Each CellConfig In rngConfigRange

        If (CellConfig.Row = longLastRow) Then
            Exit For
        End If

        For Each CellOriginDrw In rngOriginDrwNumberRange

            'Copy Drawings
            swSampleModelPath = currentPath & "\MFG DRAWINGS\" & CellOriginDrw & ".SLDDRW" 'origin drw
            targetPath = currentPath & "\MFG DRAWINGS\" 'same as origin
            bStatus = checkFileExists(swSampleModelPath)

                If (bStatus) Then

                    'Use variables in the FileCopy statement
                    Dim xlobj As Object
                    Set xlobj = CreateObject("Scripting.FileSystemObject")

                    sSuffix = rngCoppiedSuffixRange(CellOriginDrw.Row - rngOriginDrwNumberRange.Row + 1, 1)
                    If (sSuffix <> "") Then
                        sSuffix = "-" & sSuffix
                    End If

                    sCopiedModelPath = targetPath & rngCoppiedDrwNumberPrefixRange(CellConfig.Row - rngConfigRange.Row + 1 + 1, 1) & sSuffix & ".SLDDRW"
                    sToReplaceName = rngConfigPartnumberRange(CellConfig.Row - rngConfigRange.Row + 1 + 1, 1)

                    xlobj.CopyFile swSampleModelPath, sCopiedModelPath, True
                    Debug.Print sCopiedModelPath & " created!"

                    'Write results to sheet
                    ReDim Preserve vCreatedDrws(0 To longDrws)
                    vCreatedDrws(longDrws) = sCopiedModelPath

                    ReDim Preserve vToChangeConfigs(0 To longDrws)
                    vToChangeConfigs(longDrws) = rngConfigRange(CellConfig.Row).Value

                    ReDim Preserve vToReplaceNames(0 To longDrws)
                    vToReplaceNames(longDrws) = sToReplaceName

                    ReDim Preserve vOldNames(0 To longDrws)
                    vOldNames(longDrws) = rngConfigPartnumberRange(1, 1)

                    longDrws = longDrws + 1
                Else

                longResponse = MsgBox("The file is Read-Only." & Chr(13) & "Do you want to close the file without Saving?", vbCritical + vbYesNo, "FileOpenRebuildSaveClose")

                End If

        Next CellOriginDrw

        Set xlobj = Nothing

    Next CellConfig

    '' Write the vCreatedDrws values to cells BB1 to BB53 in the first worksheet
    wsUsedWs.Range("BB2:BB" & wsUsedWs.Rows.count).ClearContents
    wsUsedWs.Range("BC2:BC" & wsUsedWs.Rows.count).ClearContents
    wsUsedWs.Range("BD2:BD" & wsUsedWs.Rows.count).ClearContents
    wsUsedWs.Range("BE2:BE" & wsUsedWs.Rows.count).ClearContents

    For i = 0 To UBound(vCreatedDrws)
        wsUsedWs.Range("BB" & i + 1 + 1).Value = vCreatedDrws(i)
        wsUsedWs.Range("BC" & i + 1 + 1).Value = vToChangeConfigs(i)
        wsUsedWs.Range("BD" & i + 1 + 1).Value = vToReplaceNames(i)
        wsUsedWs.Range("BE" & i + 1 + 1).Value = vOldNames(i)
    Next i

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Change sheet name, view name and position / scale of coppied drawings

    Debug.Print "Change sheet name, view name and position / scale of coppied drawings"

    Dim sNewModelPath       As String
    Dim sOldSheetName       As String
    
    '' Initiate
    Set swApp = GetObject(, "SldWorks.Application")
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    swApp.Visible = True

    '' Coppied drawings list
    Dim rngCoppiedDrawings  As Range
    Dim rngToChangeConfigs  As Range
    Dim rngToReplaceNames   As Range
    Dim rngOldNames         As Range

    longLastRow = wsUsedWs.Range("BB" & wsUsedWs.Rows.count).End(xlUp).Row
    Set rngCoppiedDrawings = wsUsedWs.Range("BB2:BB" & longLastRow)
    Set rngToChangeConfigs = wsUsedWs.Range("BC2:BC" & longLastRow)
    Set rngToReplaceNames = wsUsedWs.Range("BD2:BD" & longLastRow)
    Set rngOldNames = wsUsedWs.Range("BE2:BE" & longLastRow)

    '' Get orinignal sheet drawing view posision
'    For Each CellOriginDrw In rngOriginDrwNumberRange
'
'        ''' Open seed drawing, get views properties and stored
'        swSampleModelPath = currentPath & "\MFG DRAWINGS\" & CellOriginDrw & ".SLDDRW" 'origin drw
'        targetPath = currentPath & "\MFG DRAWINGS\" 'same as origin
'        bStatus = checkFileExists(swSampleModelPath)
'
'        If (bStatus) Then
'
'            Set swOriginModel = swapp.OpenDoc6(swSampleModelPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", longErrors, longWarnings)
'            Set swDrawOrigin = swOriginModel
'            sCurrentSheetName = rngOriginPartNumberRange(rngOriginDrwNumberRange.Row - rngConfigRange.Row + 1 + 1, 1) & sSuffix & ".SLDDRW"
'
'            '''' Get views array and position here
'
'
'            '''' Close sample doc
'            swapp.CloseDoc swSampleModelPath
'
'            '''' Change views position and scale
'        Else
'
'            longResponse = MsgBox("The file is Read-Only." & Chr(13) & "Do you want to close the file without Saving?", vbCritical + vbYesNo, "FileOpenRebuildSaveClose")
'
'        End If
'
'    Next CellOriginDrw

    '' Change new drawings' sheet and view name

    For Each Cell In rngCoppiedDrawings

        sNewModelPath = Cell.Value
        sOriginConfig = rngConfigRange(1, 1).Value
        sNewConfig = rngToChangeConfigs(Cell.Row - 1).Value

        ''' Open each children drawing
        Set swModel = swApp.OpenDoc6(sNewModelPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", longErrors, longWarnings)
        Set swModelDocExt = swModel.Extension
        Debug.Print "     Open " & sNewModelPath

        If Not swModel Is Nothing Then
            '''' Set the working directory to the document directory
            swApp.SetCurrentWorkingDirectory (Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\")))

            '''' Activate the loaded document and prompt for rebuild to use getComponents
            Set swModel = swApp.ActivateDoc3(swModel.GetTitle(), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, longErrors)

            Set swDraw = swModel
            Debug.Print ("Error code after document activation: " & longErrors)

            '''' Change sheet name and view name
            sOldSheetName = rngOldNames(Cell.Row - 1).Value
            sSheetNameOfNewCopy = rngToReplaceNames(Cell.Row - 1).Value

            F_MassReplaceDrwSheetName2 sOldSheetName, sSheetNameOfNewCopy ' Change sheets' name
            F_ReplaceDrwViewName2 sOldSheetName, sSheetNameOfNewCopy    ' Change views' name
            
            '''' Get seed views current position array
            Dim rSeedViewWidth(), rSeedViewHeight() As Variant
            Dim rSeedViewPosX(), rSeedViewPosY() As Variant
            S_GetViewPosition swDraw, rSeedViewWidth(), rSeedViewHeight(), rSeedViewPosX(), rSeedViewPosY()
            
            '''' Change view configuration
            S_ChangeViewConfiguration sOriginConfig, sNewConfig
            
            '''' Re-position of views after changing configuration
            S_RePositionViewsByArray swDraw, rSeedViewWidth, rSeedViewHeight, rSeedViewPosX, rSeedViewPosY

            '''' Change drawing properties
            Dim bIsWeldmentTable As Boolean
            bIsWeldmentTable = False
            
            If (InStr(swDraw.GetTitle(), "-BC") > 0 _
            Or InStr(swDraw.GetTitle(), "-JPL") > 0 _
            Or InStr(swDraw.GetTitle(), "-TC") > 0 _
            ) Then
            
                M_WeldPartDrawingProp.Main  'If welding body drawings -> Change properties
            
            ElseIf (InStr(swDraw.GetTitle(), "-IFX") > 0) Then
                
                'Do smthg
                
            Else
            
                bIsWeldmentTable = True
                
            End If
  
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''' If it's weldment drawing then delete old table and insert new table
            If bIsWeldmentTable Then
            
                '''''' Get first feature in FeatureManager design tree
                Set swFeature = swModel.FirstFeature
                
                '''''' If the type of feature is "WeldmentTableFeat" then get the WeldmentCutListFeature object
                
                Do While Not swFeature Is Nothing
            
                    If swFeature.GetTypeName = "WeldmentTableFeat" Then
                        
                        Debug.Print swFeature.name
                        'Set swWeldmentCutListFeat = swFeature.GetSpecificFeature2
                        'bStatus = swModel.Extension.SelectByID2(swFeature.name, swSelectType_e.swSelWELDMENTTABLEFEATS, 0, 0, 0, False, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
                        'Set swSelectedFeature = swSelectionMgr.GetSelectedObject6(1, -1)
                        
                        bStatus = swFeature.Select2(False, 0)
                        
                        ' To delete absorbed features, use enum swDeleteSelectionOptions_e.swDelete_Absorbed
                        ' To delete children features, use enum swDeleteSelectionOptions_e.swDelete_Children
                        ' To keep absorbed features and children features, set longDeleteOption = 0
                        
                        Dim longDeleteOption As Long
                        'longDeleteOption = swDeleteSelectionOptions_e.swDelete_Absorbed
                        'longDeleteOption = swDeleteSelectionOptions_e.swDelete_Children
                        'longDeleteOption = 0
                        longDeleteOption = swDeleteSelectionOptions_e.swDelete_Absorbed + swDeleteSelectionOptions_e.swDelete_Children
                        bStatus = swModel.Extension.DeleteSelection2(longDeleteOption)
                        Debug.Print "Feature deleted? " & bStatus
                                           
                    End If
               
                    ' Get the next feature in the FeatureManager design tree
                    Set swFeature = swFeature.GetNextFeature
               
                Loop
                    
                '''''' Insert new weldment cut list
                    S_MyCutList "C:\Users\khanh.nguyen\OneDrive - xxx\SOLIDWORKS\CutListTableTemplate.sldwldtbt"
                
            End If
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            swModel.ForceRebuild3 True
            swModel.Save2 True

            swApp.CloseDoc sNewModelPath
        End If
    Next Cell

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Restore the message filter after calling Subs (remove OLE waiting warning).

    CoRegisterMessageFilter lMsgFilter, lMsgFilter
    
End Sub


Function F_MassReplaceDrwSheetName2(ByVal sOldText, sNewText As String)
    
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

Sub S_DrawingMFG_PopulateTopAngle()

    'Remove the message filter before calling Subs (remove OLE waiting warning).
    CoRegisterMessageFilter 0&, lMsgFilter

    Dim swModel, swOriginModel As SldWorks.ModelDoc2
    Dim swActiveModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim swDrawOrigin As SldWorks.DrawingDoc
    
    Dim longErrors As Long
    Dim longWarnings As Long
    Dim longResponse As Long
    
    Dim bStatus As Boolean
    
    Dim currentPath As String
    Dim swSampleModelPath As String
    Dim swModelPath As String
   
    Dim newFolderPath As String
    Dim folderPath As String
    
'    Dim usedRange As String
    Dim currentModelParentPath As String
    Dim targetPath As String
    Dim sCopiedModelPath As String
    
    Dim sOriginRange, sOriginConfigRange, sOriginModel, sOriginDrwNo, sOriginDrwNumberRange As String
    Dim sConfigRange, sOriginConfig, sNewConfig, sNewDrwNo, sDrwNumberRange As String
    Dim sSheetNameOfNewCopy As String
    Dim sCurrentSheetName As String
    Dim sNewSheetName As String
    
    Dim sCurrentWorkbookPath As String
    Dim wbUsedWb As Workbook
    Dim wsUsedWs As Worksheet
    Dim Cell, CellOrigin, CellOriginDrw As Range
    Dim rngOriginDrw As Range
    
    Dim longLabelCounter As Long
    Dim longLabelCounterBuffer As Long
    
    Dim sSuffix As String
    
    Dim sPrpDrwTitle, sPrpDrwNo, sPrpDrwDescription As String
    

'////////////////////////////////////////////////////////////////

    ' Set the path to your Excel file
    
    'sCurrentWorkbookPath = sWorkbookPath
    'Workbooks.Open FileName:=sCurrentWorkbookPath
    'Set wbUsedWb = Workbooks.Open(sCurrentWorkbookPath)
    Set wbUsedWb = ThisWorkbook
    Set wsUsedWs = wbUsedWb.Sheets("Sheet1")

'    sConfigRange = "E2:E7"
'    sOriginConfigRange = "E8"
    sConfigRange = "A13:A18"
    sOriginConfigRange = "A13"
    sDrwNumberRange = "AB13:AB18"
    sOriginDrwNumberRange = "AB13"
    sSuffix = " TOP ANGLE"
    
'    usedRange = "A23:A23"
    sOriginRange = "A13"
    
    Set CellOrigin = wsUsedWs.Range(sOriginRange)(1, 1)
    sOriginModel = CellOrigin.Value
    
    Set Cell = wsUsedWs.Range(sOriginConfigRange)
    sOriginConfig = CStr(Cell.Value)
    
    Set CellOriginDrw = wsUsedWs.Range(sOriginDrwNumberRange)
    sOriginDrwNo = CStr(CellOriginDrw.Value)
    Set rngOriginDrw = wsUsedWs.Range(sDrwNumberRange)
     
    currentPath = GetLocalWorkbookName(ThisWorkbook.fullName, True)
    
    Debug.Print "S_DrawingMFG_PopulateTopAngle"
    'Copy Drawings
    swSampleModelPath = currentPath & "\MFG DRAWINGS\" & sOriginDrwNo & ".SLDDRW" 'origin drw
    targetPath = currentPath & "\MFG DRAWINGS\" 'same as origin
    bStatus = checkFileExists(swSampleModelPath)

    If (bStatus) Then

        'Loop through each cell in the range then copy and change names
        'Use variables in the FileCopy statement
        Dim xlobj As Object
        Set xlobj = CreateObject("Scripting.FileSystemObject")

        'object.copyfile,source,destination,file overwrite(True is default)
        For Each Cell In Range(sDrwNumberRange)

            sCopiedModelPath = targetPath & Cell.Value & ".SLDDRW"

            xlobj.CopyFile swSampleModelPath, sCopiedModelPath, True
            Debug.Print sCopiedModelPath & " created!"

        Next Cell
        Set xlobj = Nothing
    Else
        longResponse = MsgBox("The file is Read-Only." & Chr(13) & "Do you want to close the file without Saving?", vbCritical + vbYesNo, "FileOpenRebuildSaveClose")
    End If
    
    '''Initiate
    Set swApp = GetObject(, "SldWorks.Application")
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    swApp.Visible = True
    
    '''Open seed drawing, get views properties and stored
    Set swOriginModel = swApp.OpenDoc6(swSampleModelPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", longErrors, longWarnings)
    Set swDrawOrigin = swOriginModel
    sCurrentSheetName = sOriginModel & sSuffix

    longLabelCounter = 1
    longLabelCounterBuffer = 1

    swApp.CloseDoc swSampleModelPath
    '''Open each children drawing and change it properties
    For Each Cell In Range(sConfigRange)

        sNewConfig = Cell.Value
        swModelPath = targetPath & rngOriginDrw(Cell.Row - rngOriginDrw.Row + 1, 1) & ".SLDDRW"

        Set swModel = swApp.OpenDoc6(swModelPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", longErrors, longWarnings)
'        Set swModelDocExt = swModel.Extension
        Debug.Print "     Open " & swModelPath

        If Not swModel Is Nothing Then
            ' Set the working directory to the document directory
            swApp.SetCurrentWorkingDirectory (Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\")))

            ' Activate the loaded document and prompt for rebuild to use getComponents
            Set swModel = swApp.ActivateDoc3(swModel.GetTitle(), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, longErrors)

            Set swDraw = swModel
            Debug.Print ("Error code after document activation: " & longErrors)

            'Change sheet name and view name
            sSheetNameOfNewCopy = sNewConfig & sSuffix
            sNewSheetName = F_ReplaceDrwSheetName(sCurrentSheetName, sCurrentSheetName, sSheetNameOfNewCopy)

            ' Change view name
            F_MassReplaceDrwSheetName sOriginConfig, sNewConfig
            F_ReplaceDrwViewName2 sOriginConfig, sNewConfig
            S_ChangeViewConfiguration sOriginConfig, sNewConfig
            'S_RePositionViewsBySeedFile2 swDrawOrigin, swDraw


            'Change drawing properties
            sPrpDrwTitle = sNewConfig & " TOP ANGLE"
            'sPrpDrwNo = "PD-TA-" & sNewConfig
            'sPrpDrwDescription = sNewConfig & " TOP ANGLE"

            Edit_Properties "Title", sPrpDrwTitle
            'Edit_Properties "Drawing No", sPrpDrwNo
            'Edit_Properties "Description", sPrpDrwDescription

            swModel.ForceRebuild3 True
            swModel.Save2 True

            swApp.CloseDoc swModelPath
        End If
    Next Cell

    

    ''' Restore the message filter after calling Subs (remove OLE waiting warning).
    CoRegisterMessageFilter lMsgFilter, lMsgFilter
    
End Sub


Sub S_DrawingPopulateConfigs()

    ''' Remove the message filter before calling Subs (remove OLE waiting warning).
    CoRegisterMessageFilter 0&, lMsgFilter

    Dim swModel, swOriginModel As SldWorks.ModelDoc2
    Dim swActiveModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim swDrawOrigin As SldWorks.DrawingDoc
    Dim swModelDocExt As SldWorks.ModelDocExtension
    Dim swSelectionMgr As SldWorks.SelectionMgr
    
    Dim longErrors As Long
    Dim longWarnings As Long
    Dim longResponse As Long
    
    Dim bStatus As Boolean
    
    Dim currentPath As String
    Dim swSampleModelPath As String
    Dim swModelPath As String
   
    Dim newFolderPath As String
    Dim folderPath As String
    
'    Dim usedRange As String
    Dim currentModelParentPath As String
    Dim targetPath As String
    Dim sCopiedModelPath As String
    
    Dim sOriginRange, sOriginConfigRange, sOriginModel As String
    Dim sConfigRange, sOriginConfig, sNewConfig As String
    Dim sSheetNameOfNewCopy As String
    Dim sCurrentSheetName As String
    Dim sNewSheetName As String
    
    Dim sCurrentWorkbookPath As String
    Dim wbUsedWb As Workbook
    Dim wsUsedWs As Worksheet
    Dim Cell, CellOrigin As Range
    
    Dim longLabelCounter As Long
    Dim longLabelCounterBuffer As Long

'////////////////////////////////////////////////////////////////

    ' Set the path to your Excel file
    
    'sCurrentWorkbookPath = sWorkbookPath
    'Workbooks.Open FileName:=sCurrentWorkbookPath
    'Set wbUsedWb = Workbooks.Open(sCurrentWorkbookPath)
    Set wbUsedWb = ThisWorkbook
    Set wsUsedWs = wbUsedWb.Sheets("Sheet1")

    sConfigRange = "E2:E7"
    sOriginConfigRange = "E8"
    
'    usedRange = "A23:A23"
    sOriginRange = "A25"
    
    Set CellOrigin = wsUsedWs.Range(sOriginRange)(1, 1)
    sOriginModel = CellOrigin.Value
    
    Set Cell = wsUsedWs.Range(sOriginConfigRange)
    sOriginConfig = CStr(Cell.Value)
     
    currentPath = GetLocalWorkbookName(ThisWorkbook.fullName, True)

    Debug.Print "S_DrawingPopulateConfigs"
    'Copy Drawings
    targetPath = currentPath & "\" & sOriginModel & "\2D FILES\" 'same as origin
    swSampleModelPath = currentPath & "\" & sOriginModel & "\2D FILES\" & sOriginModel & "-TANK-GA.SLDDRW"
    bStatus = checkFileExists(swSampleModelPath)

    If (bStatus) Then

        'Loop through each cell in the range then copy and change names
        'Use variables in the FileCopy statement
        Dim xlobj As Object
        Set xlobj = CreateObject("Scripting.FileSystemObject")

        'object.copyfile,source,destination,file overwrite(True is default)
        For Each Cell In Range(sConfigRange)

            sCopiedModelPath = targetPath & sOriginModel & "-" & Cell.Value & "-B-TANK-GA.SLDDRW"

            xlobj.CopyFile swSampleModelPath, sCopiedModelPath, True
            Debug.Print sCopiedModelPath & " created!"

        Next Cell
        Set xlobj = Nothing
    Else
        longResponse = MsgBox("The file is Read-Only." & Chr(13) & "Do you want to close the file without Saving?", vbCritical + vbYesNo, "FileOpenRebuildSaveClose")
    End If
    
    '''Initiate
    Set swApp = GetObject(, "SldWorks.Application")
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    swApp.Visible = True
    
    '''Open seed drawing, get views properties and stored
    Set swOriginModel = swApp.OpenDoc6(swSampleModelPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", longErrors, longWarnings)
    Set swDrawOrigin = swOriginModel
    sCurrentSheetName = sOriginModel & "-" & sOriginConfig & "-" & "B"
    
    longLabelCounter = 1
    longLabelCounterBuffer = 1
    
    '''Open each children drawing and change it properties
    For Each Cell In Range(sConfigRange)

        sNewConfig = Cell.Value
        swModelPath = targetPath + sOriginModel + "-" + sNewConfig + "-B" + "-TANK-GA.SLDDRW"
        
        Set swModel = swApp.OpenDoc6(swModelPath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", longErrors, longWarnings)
'        Set swModelDocExt = swModel.Extension
        Debug.Print "    Open " & swModelPath

        If Not swModel Is Nothing Then
            ' Set the working directory to the document directory
            swApp.SetCurrentWorkingDirectory (Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\")))

            ' Activate the loaded document and prompt for rebuild to use getComponents
            Set swModel = swApp.ActivateDoc3(swModel.GetTitle(), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, longErrors)
            Set swDraw = swModel
            Debug.Print ("Error code after document activation: " & longErrors)

            'Change sheet name and view name
            sSheetNameOfNewCopy = sOriginModel & "-" & sNewConfig & "-" & "B"
            sNewSheetName = F_ReplaceDrwSheetName(sCurrentSheetName, sCurrentSheetName, sSheetNameOfNewCopy)

            ' Change view name by sheet name
            S_MassReplaceDrwViewNameBySheetName
            S_ChangeViewConfigurationBySheetName
            S_RePositionViewsBySeedFile2 swDrawOrigin, swDraw
            
            longLabelCounter = longLabelCounterBuffer
            S_ViewSequentialLabelingPreserve longLabelCounter
            longLabelCounterBuffer = longLabelCounter
            
            'swModel.Save2 True

            swModel.ForceRebuild3 True
            swModel.Save2 True

            swApp.CloseDoc swModelPath
        End If
    Next Cell

    ''' Restore the message filter after calling Subs (remove OLE waiting warning).
    CoRegisterMessageFilter lMsgFilter, lMsgFilter
    
End Sub

Sub S_RePositionViewsByArray(swNewDrawingDoc As SldWorks.DrawingDoc, _
                            ByVal rSeedViewWidth As Variant, _
                            ByVal rSeedViewHeight As Variant, _
                            ByVal rSeedViewPosX As Variant, _
                            ByVal rSeedViewPosY As Variant)

    Dim vSeedOutline() As Variant
    Dim vSeedPos() As Variant
    Dim nNumView As Long
    Dim bRet As Boolean
    Dim swModel As SldWorks.ModelDoc2
    
    ''' Set to new drawing file
    '''' Views array including sheets
    Dim vSheets As Variant
    vSheets = swNewDrawingDoc.GetViews()
    
    Dim iIndex As Integer
    
    For iIndex = UBound(vSheets) To 0 Step -1
        
        Dim nextViewIndex As Integer
        nextViewIndex = 0
        
        'Views array of a sheet
        Dim vViews As Variant
        vViews = vSheets(iIndex)
        
        'Sheet is the first of sheet view array
        Dim swSheetView As SldWorks.view
        Set swSheetView = vViews(0)
        
        Dim iNumViewOfSheet As Integer
        iNumViewOfSheet = UBound(vViews)
        
        ReDim vOutline(iNumViewOfSheet)
        ReDim vPos(iNumViewOfSheet)
        
        Dim swSketchPoint As Object
        Dim j As Integer
        
        For j = 1 To UBound(vViews)
        
            Debug.Print "Select view no. " & j & " of sheet no. " & iIndex
            
            Dim swView As SldWorks.view
            Set swView = vViews(j)
            
            Dim viewType As Integer
            viewType = swView.Type
            
            'If viewType <> swDrawingViewTypes_e.swDrawingDetailView And viewType <> swDrawingViewTypes_e.swDrawingSectionView Then
                            
                nextViewIndex = nextViewIndex + 1
                vOutline(j) = swView.GetOutline
                vPos(j) = swView.Position
                Debug.Print "View = " + swView.GetName2
                Debug.Print "  Pos = (" & vPos(j)(0) * 1000# & ", " & vPos(j)(1) * 1000# & ") mm"
                Debug.Print "  Min = (" & vOutline(j)(0) * 1000# & ", " & vOutline(j)(1) * 1000# & ") mm"
                Debug.Print "  Max = (" & vOutline(j)(2) * 1000# & ", " & vOutline(j)(3) * 1000# & ") mm"
                
                Dim rViewWidth As Double
                Dim rViewHeight As Double
                GetViewGeometrySize swView, rViewWidth, rViewHeight
                
                Dim rViewScale As Double
                GetViewScale swView, rViewScale
                
                Debug.Print "  Geometry Size: " & rViewWidth & " x " & rViewHeight
                
                'Calculate and reposition of populated view
                vPos(j)(0) = rSeedViewPosX(j) - (rViewWidth - rSeedViewWidth(j)) / 2 / rViewScale
                vPos(j)(1) = rSeedViewPosY(j) - (rViewHeight - rSeedViewHeight(j)) / 2 / rViewScale
                Debug.Print "  New position X = " & vPos(j)(0); ", Y = " & vPos(j)(0)
            
                swView.Position = vPos(j)
                
               ' End If
                
                
            'End If
            
        Next
        
    Next
    
    Set swModel = swNewDrawingDoc
    swModel.GraphicsRedraw2
    swNewDrawingDoc.EditRebuild
    swModel.Save2 True
    'swApp.CloseDoc swModel.GetPathName
End Sub

Sub S_GetViewPosition(swSeedDrawingDoc As SldWorks.DrawingDoc, _
                        ByRef rSeedViewWidth() As Variant, _
                        ByRef rSeedViewHeight() As Variant, _
                        ByRef rSeedViewPosX() As Variant, _
                        ByRef rSeedViewPosY() As Variant)

    Dim vSeedOutline() As Variant
    Dim vSeedPos() As Variant
    Dim nNumView As Long
    Dim bRet As Boolean
    Dim swModel As SldWorks.ModelDoc2
''' Seed Drawing parameters

    'Views array including sheets
    Dim vSeedSheets As Variant
    
    
    
    vSeedSheets = swSeedDrawingDoc.GetViews()
    
'    Dim rSeedViewWidth(), rSeedViewHeight() As Variant
'    Dim rSeedViewPosX(), rSeedViewPosY() As Variant
    
    Dim i As Integer
    
    'For i = 0 To UBound(vSeedSheets)
    For i = UBound(vSeedSheets) To 0 Step -1
    'Get seed only
    'i = UBound(vSeedSheets)
    
        'Views array of a sheet
        Dim vSeedViews As Variant
        vSeedViews = vSeedSheets(i)
        
        'Sheet is the first of sheet view array
        Dim swSeedSheetView As SldWorks.view
        Set swSeedSheetView = vSeedViews(0)
        
        Dim j As Integer
        
        Dim iNumViewOfSeedSheet As Integer
        iNumViewOfSeedSheet = UBound(vSeedViews)
        
        ReDim vSeedOutline(iNumViewOfSeedSheet)
        ReDim vSeedPos(iNumViewOfSeedSheet)
  
        ReDim Preserve rSeedViewWidth(iNumViewOfSeedSheet), rSeedViewHeight(iNumViewOfSeedSheet)
        ReDim Preserve rSeedViewPosX(iNumViewOfSeedSheet), rSeedViewPosY(iNumViewOfSeedSheet)
        
        For j = 1 To UBound(vSeedViews)
        
            Debug.Print "Select view no. " & j & " of sheet no. " & i
            
            Dim swSeedView As SldWorks.view
            Set swSeedView = vSeedViews(j)
            
            Dim viewType As Integer
            viewType = swSeedView.Type
            
            'If viewType <> swDrawingViewTypes_e.swDrawingDetailView And viewType <> swDrawingViewTypes_e.swDrawingSectionView Then

                vSeedOutline(j) = swSeedView.GetOutline
                vSeedPos(j) = swSeedView.Position
                Debug.Print "View = " + swSeedView.GetName2
                Debug.Print "  Pos = (" & vSeedPos(j)(0) * 1000# & ", " & vSeedPos(j)(1) * 1000# & ") mm"
                Debug.Print "  Min = (" & vSeedOutline(j)(0) * 1000# & ", " & vSeedOutline(j)(1) * 1000# & ") mm"
                Debug.Print "  Max = (" & vSeedOutline(j)(2) * 1000# & ", " & vSeedOutline(j)(3) * 1000# & ") mm"
                
                Dim rViewWidth As Double
                Dim rViewHeight As Double
                GetViewGeometrySize swSeedView, rViewWidth, rViewHeight
                
                Dim rViewScale As Double
                GetViewScale swSeedView, rViewScale
                
                Debug.Print "  Geometry Size: " & rViewWidth & " x " & rViewHeight
                
                'If (i = UBound(vSheets)) Then
                    
                    'Get seed view parameter
                    rSeedViewWidth(j) = rViewWidth
                    rSeedViewHeight(j) = rViewHeight
                    rSeedViewPosX(j) = vSeedPos(j)(0)
                    rSeedViewPosY(j) = vSeedPos(j)(1)
                'End If
                
            'End If
            
        Next
    Next

End Sub


Sub GetViewGeometrySize(view As SldWorks.view, ByRef width As Double, ByRef height As Double)
    
    Dim borderWidth As Double
    borderWidth = GetViewBorderWidth(view)
    
    Dim vOutline As Variant
    vOutline = view.GetOutline()
    
    Dim viewScale As Double
    viewScale = view.ScaleRatio(1) / view.ScaleRatio(0)
    
    width = (vOutline(2) - vOutline(0) - borderWidth * 2) * viewScale
    height = (vOutline(3) - vOutline(1) - borderWidth * 2) * viewScale
    
End Sub

Sub GetViewScale(view As SldWorks.view, ByRef rViewScale As Double)
    
    rViewScale = view.ScaleRatio(1) / view.ScaleRatio(0)
    
End Sub

Function GetViewBorderWidth(view As SldWorks.view) As Double
    
    Const VIEW_BORDER_RATIO = 0.02
    
    Dim width As Double
    Dim height As Double
    
    view.sheet.GetSize width, height
    
    Dim minSize As Double
    
    If width < height Then
        minSize = width
    Else
        minSize = height
    End If
    
    GetViewBorderWidth = minSize * VIEW_BORDER_RATIO
    
End Function


Sub S_ChangeViewConfiguration(ByVal sOldConfig, ByVal sNewConfig As String)
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
                Dim sCurrentViewConfig, sNewViewConfig As String
                sCurrentViewConfig = swSelectedView.ReferencedConfiguration
                Dim sCurrentViewDisplayState As String
                
                If swSelectedView.DisplayState <> "<Default>_Display State 1" Then
                
                    sCurrentViewDisplayState = swSelectedView.DisplayState
                
                End If

                Debug.Print "View " & swSelectedView.name & "'s model's configuration: " & swSelectedView.ReferencedConfiguration & ". State: " & swSelectedView.DisplayState
                
                sNewViewConfig = Replace(sCurrentViewConfig, sOldConfig, sNewConfig)
                swSelectedView.ReferencedConfiguration = sNewViewConfig
                
                swSelectedView.DisplayState = sCurrentViewDisplayState     'Remain choosen display state
                Debug.Print "   Changed to " & swSelectedView.ReferencedConfiguration & ". State: " & swSelectedView.DisplayState
                    
            End If
            
            swActiveModel.ClearSelection
        Next
    Next
    
End Sub


Sub Edit_Properties(sProperties As String, ByVal sPrpValue As String)

'    Set swApp = GetObject(, "SldWorks.Application")
'    If swApp Is Nothing Then
'        Set swApp = CreateObject("SldWorks.Application")
'    End If
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swCustomProperties As SldWorks.CustomPropertyManager
    
    Dim Value_Expression As String
    Dim Evaluated_Value As String
    Dim wasResolved As Boolean
    Dim Islinked As Boolean
    Dim bStatus As Boolean
    
    '---------------------------------------
    
    Set swModel = swApp.ActiveDoc
    Set swCustomProperties = swModel.Extension.CustomPropertyManager("")
    
    ' Print Order Number Property
    Dim Order_Number As Integer
    Order_Number = swCustomProperties.Get6(sProperties, False, Value_Expression, Evaluated_Value, wasResolved, Islinked)
    Debug.Print sProperties & ": " & Order_Number
    
    ' Edit Property
    bStatus = swCustomProperties.IsCustomPropertyEditable(sProperties, "Default")
    If bStatus Then
        bStatus = swCustomProperties.Add3(sProperties, swCustomInfoType_e.swCustomInfoText, sPrpValue, swCustomPropertyAddOption_e.swCustomPropertyReplaceValue)
    Else
        Debug.Print sProperties & " is not editable"
    End If
    
End Sub

Function F_ReplaceDrwViewName2(ByVal sOldText, ByVal sNewText As String)
    
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    
    Set swModel = swApp.ActiveDoc
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
