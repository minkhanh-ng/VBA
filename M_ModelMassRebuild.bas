Attribute VB_Name = "M_ModelMassRebuild"
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swActiveModel As SldWorks.ModelDoc2

Dim bStatus As Boolean
Dim arStatuses As Variant
Dim i As Long

Dim currentPath As String
Dim targetPath As String
Dim swModelPath As String
Dim arUpdateModels(1) As Variant

Dim wb As Workbook
Dim ws As Worksheet

Dim currentWorkbookPath As String
Dim currentWb As Workbook
Dim currentWs As Worksheet

Dim Cell As Range
Dim bDone As Boolean

Dim strResponse As Long
Dim strFileType As Long
Dim longError As Long
Dim longWarning As Long

Dim arConfigurations As Variant
Dim sConfigName As String
Dim start As Single

Dim bShowConfig As Boolean
Dim bRebuild As Boolean

Sub modelMassRebuild()

    Dim i As Long
    Dim j As Long

    ' Set the path to your Excel file
    ' Initialize SOLIDWORKS
    'Set swApp = CreateObject("SldWorks.Application")
    Set swApp = GetObject(, "SldWorks.Application")
    
    currentPath = GetLocalWorkbookName(ThisWorkbook.fullName, True)
    
    currentWorkbookPath = currentPath + "\PopulateMacro.xlsm"
    Set currentWb = Workbooks.Open(currentWorkbookPath)
    Set currentWs = currentWb.Sheets("Sheet1")
    
    For Each Cell In currentWs.Range("A2:A12")
        targetPath = currentPath & "\" & Cell.Value & "\3D FILES\"
        
        arUpdateModels(0) = targetPath + Cell.Value + "-CUSTOM BLUE ORC PANEL RF PANEL R2.SLDPRT"
        'arUpdateModels(1) = targetPath + cell.Value + "-LAMINATIONS.SLDASM"
        'arUpdateModels(2) = targetPath + cell.Value + "-RINGS.SLDASM"
        'arUpdateModels(3) = targetPath + cell.Value + "-WALL.SLDASM"
        arUpdateModels(1) = targetPath + Cell.Value + "-TANK.SLDASM"
        
        For i = 0 To UBound(arUpdateModels)
            swModelPath = arUpdateModels(i)
                
            'Get model type
            If StrComp((UCase$(Right$(swModelPath, 7))), ".SLDPRT", vbTextCompare) = 0 Then
                strFileType = swDocPART
            ElseIf StrComp((UCase$(Right$(swModelPath, 7))), ".SLDASM", vbTextCompare) = 0 Then
                strFileType = swDocASSEMBLY
            ElseIf StrComp((UCase$(Right$(swModelPath, 7))), ".SLDDRW", vbTextCompare) = 0 Then
                strFileType = swDocDRAWING
            End If
            
            Set swModel = swApp.OpenDoc6(swModelPath, strFileType, 0, "", longError, longWarning)
            Set swModel = swApp.ActivateDoc3(swModel.GetTitle(), False, swRebuildOnActivation_e.swDontRebuildActiveDoc, longError)
            
            If (swModel Is Nothing) Then
                strResponse = MsgBox("The file could not be found." & Chr(13) & "Routine Ending.", vbCritical, "FileOpenRebuildSaveClose")
                End
            End If
            
            If (swModel.IsOpenedReadOnly = "False") Then
                'Rebuild all configurations
                arConfigurations = swModel.GetConfigurationNames
                            
                For j = 0 To 1
                    sConfigName = arConfigurations(j)
                    bShowConfig = swModel.ShowConfiguration2(sConfigName)
                    start = Timer
                    bRebuild = swModel.ForceRebuild3(False)
                    Debug.Print "  Configuration = " & sConfigName
                    Debug.Print "    ShowConfig  = " & bShowConfig
                    Debug.Print "    Rebuild     = " & bRebuild
                    Debug.Print "    Time        = " & Timer - start & " seconds"
                Next j
                
                'Rebuild File
                'swModel.EditRebuild3 'Stoplight or [Ctrl]+B
                swModel.ForceRebuild '[Ctrl]+Q
                
                'Zoom to extents
                'swModel.ViewZoomtofit2
                
                'Save
                swModel.Save2 True
                swApp.CloseDoc swModelPath
            Else
                strResponse = MsgBox("The file is Read-Only." & Chr(13) & "Do you want to close the file without Saving?", vbCritical + vbYesNo, "FileOpenRebuildSaveClose")
            End If
            
            If (strResponse = VbMsgBoxResult.vbYes) Then
                'Close
                swApp.CloseDoc swModelPath
            End If
            
            Debug.Print "   Rebuild " & swModelPath & " done"
            Debug.Print " "
        Next i
    Next Cell
End Sub
