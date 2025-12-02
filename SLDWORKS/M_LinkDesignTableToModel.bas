Attribute VB_Name = "M_LinkDesignTableToModel"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2

Dim errors As Long
Dim warnings As Long
Dim longStatus As Long

Dim currentPath As String
Dim currentWorkbookPath As String
Dim currentWb As Workbook
Dim currentWs As Worksheet

' Set user preference to not rebuild on activation
Dim bValue     As Boolean
Dim lValue     As Long
Dim nValue     As SwConst.swRebuildOnActivation_e

Dim strResponse As Long
Dim strFileType As Long

Dim arUpdateModels(6) As Variant
Dim arDesignTables(6) As Variant

Dim bStatus As Boolean

Dim arConfigurationNames() As String
Dim longIndx As Long

Sub S_LinkDesignTableToModel(sParentFolderPath, sWorkbookPath, sUsedRange As String)
    Dim i As Long
    
    Dim sCurrentWorkbookPath As String
    Dim wbUsedWb As Workbook
    Dim wsUsedWs As Worksheet
    
    Dim sModelParentPath As String
    Dim arFromFiles(6) As Variant
    Dim arToFiles(6) As Variant
    Dim Cell As Range
    Dim bDone As Boolean
    
    Dim sTempFromPath, sTempToPath As String
    Dim sExcelFilePath As String
    
    Dim sDesignTableParentPath As String
    
    Dim swDesignTable As SldWorks.DesignTable
    Dim swModelPath As String

'////////////////////////////////////////////////////////////////

    ' Set the path to your Excel file
    
    sCurrentWorkbookPath = sWorkbookPath
    'Workbooks.Open FileName:=sCurrentWorkbookPath
    'Set wbUsedWb = Workbooks.Open(sCurrentWorkbookPath)
    Set wbUsedWb = ThisWorkbook
    Set wsUsedWs = wbUsedWb.Sheets("Sheet1")
    
''' Initialize SOLIDWORKS
    'Set swApp = CreateObject("SldWorks.Application")
    Set swApp = GetObject(, "SldWorks.Application")
    'swApp.Visible = True
    
    For Each Cell In wsUsedWs.Range(sUsedRange)
        sModelParentPath = sParentFolderPath & "\" & Cell.Value & "\3D FILES\"
        sDesignTableParentPath = sParentFolderPath & "\" & Cell.Value & "\3D FILES\"
        
        arUpdateModels(0) = sModelParentPath + Cell.Value + "-CUSTOM BLUE ORC PANEL RF PANEL R2.SLDPRT"
        arUpdateModels(1) = sModelParentPath + Cell.Value + "-LAMINATIONS.SLDASM"
        arUpdateModels(2) = sModelParentPath + Cell.Value + "-RINGS.SLDASM"
        arUpdateModels(3) = sModelParentPath + Cell.Value + "-WALL.SLDASM"
        arUpdateModels(4) = sModelParentPath + Cell.Value + "-ROOFSHEETS.SLDPRT"
        arUpdateModels(5) = sModelParentPath + Cell.Value + "-ROOFSHEET-ASSY.SLDASM"
        arUpdateModels(6) = sModelParentPath + Cell.Value + "-TANK.SLDASM"
        
        arDesignTables(0) = sDesignTableParentPath + "DesignTable__ " + Cell.Value + "-CUSTOM BLUE ORC PANEL RF PANEL R2.xlsx"
        arDesignTables(1) = sDesignTableParentPath + "DesignTable__ " + Cell.Value + "-LAMINATIONS.xlsx"
        arDesignTables(2) = sDesignTableParentPath + "DesignTable__ " + Cell.Value + "-RINGS.xlsx"
        arDesignTables(3) = sDesignTableParentPath + "DesignTable__ " + Cell.Value + "-WALL.xlsx"
        arDesignTables(4) = sDesignTableParentPath + "DesignTable__ " + Cell.Value + "-ROOFSHEETS.xlsx"
        arDesignTables(5) = sDesignTableParentPath + "DesignTable__ " + Cell.Value + "-ROOFSHEET-ASSY.xlsx"
        arDesignTables(6) = sDesignTableParentPath + "DesignTable__ " + Cell.Value + "-TANK.xlsx"
            
        'For i = 0 To UBound(arUpdateModels)
            i = 4
            swModelPath = arUpdateModels(i)
            sExcelFilePath = arDesignTables(i)
                
            'Get model type
            If StrComp((UCase$(Right$(swModelPath, 7))), ".SLDPRT", vbTextCompare) = 0 Then
                strFileType = swDocPART
            ElseIf StrComp((UCase$(Right$(swModelPath, 7))), ".SLDASM", vbTextCompare) = 0 Then
                strFileType = swDocASSEMBLY
            ElseIf StrComp((UCase$(Right$(swModelPath, 7))), ".SLDDRW", vbTextCompare) = 0 Then
                strFileType = swDocDRAWING
            End If
            
            Set swModel = swApp.OpenDoc6(swModelPath, strFileType, 0, "", errors, warnings)
            Set swModel = swApp.ActivateDoc2(swModelPath, False, longStatus)
'            Set swModel = swApp.ActivateDoc3(swModel.GetTitle(), False, swRebuildOnActivation_e.swRebuildActiveDoc, errors)
            
            If (swModel Is Nothing) Then
                strResponse = MsgBox("The file could not be found." & Chr(13) & "Routine Ending.", vbCritical, "FileOpenRebuildSaveClose")
            'End
            End If
            
            If (swModel.IsOpenedReadOnly = "False") Then
                bValue = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swRebuildOnActivation, swRebuildOnActivation_e.swDontRebuildActiveDoc)
        '    Debug.Print ("Rebuild user preference set to not rebuild on activation: " & bValue)
    
                nValue = swApp.GetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swRebuildOnActivation)
                'swApp.SetCurrentWorkingDirectory (Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\")))
                   
                If (swModel.GetType <> swDocDRAWING) Then

                    If (i = 4 Or i = 5) Then
                        'swModel.DeleteDesignTable
                        arConfigurationNames = swModel.GetConfigurationNames()
                        For longIndx = 0 To UBound(arConfigurationNames)
                            bStatus = swModel.DeleteConfiguration2(arConfigurationNames(longIndx))
                        Next
                    End If
                                   
                    swModel.InsertFamilyTableOpen (sExcelFilePath)
                    'swModel.InsertFamilyTableNew
                    Sleep 2500
                    
                    Set swDesignTable = swModel.GetDesignTable
                    
                    swDesignTable.EditFeature
                    swDesignTable.SourceType = swDesignTableSourceTypes_e.swDesignTableSourceFromFile
                    swDesignTable.FileName = sExcelFilePath
                    swDesignTable.LinkToFile = False
                    swDesignTable.Updatable = False
                    bStatus = swDesignTable.UpdateFeature()
                    swModel.CloseFamilyTable
                    
                    Debug.Print "boolStatus Table Update " & bStatus
                    

                    
'                    Set swDesignTable = swModel.GetDesignTable
                    
    '                'Shade Part
    '                swModel.ViewDisplayShaded
    '
    '                'Set view
    '                'swModel.ShowNamedView2 "*Isometric", 7
    '                'swModel.ShowNamedView2 "*Trimetric", 8
    '                swModel.ShowNamedView2 "*Dimetric", 9
    '
    '                'Set Feature Manager Splitter Position
    '                swModel.FeatureManagerSplitterPosition = 0.3
                
                End If
                
                'Rebuild File
                'swModel.EditRebuild3 'Stoplight or [Ctrl]+B
                swModel.ForceRebuild '[Ctrl]+Q
                
                'Zoom to extents
                'swModel.ViewZoomtofit2
                
                Sleep 1000
                'Save
                swModel.Save2 True
                                
                'Sleep 1000

                
                swApp.CloseDoc swModelPath
            
            Else
                strResponse = MsgBox("The file is Read-Only." & Chr(13) & "Do you want to close the file without Saving?", vbCritical + vbYesNo, "FileOpenRebuildSaveClose")
            End If
            
            If (strResponse = VbMsgBoxResult.vbYes) Then
                'Close
                swApp.CloseDoc swModelPath
            End If
            
            Debug.Print "   Mapping " & swModelPath & " to " & sExcelFilePath & " success? " & bDone
            Debug.Print " "
        'Next i
            
    Next Cell
   
End Sub
