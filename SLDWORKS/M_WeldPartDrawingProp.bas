Attribute VB_Name = "M_WeldPartDrawingProp"
Option Explicit

Dim swApp As SldWorks.SldWorks
Dim fs As Object
Dim a As Object

Dim vPartCustomPropNames, vPartCustomPropVals, vPartCustomPropResolvedVals As Variant

Sub Main()
    
    Dim swModel                                         As SldWorks.ModelDoc2
    Dim swView                                          As SldWorks.view
    Dim vComps                                          As Variant
    Dim swComp                                          As SldWorks.Component2
    
    Dim swPartModel                                     As SldWorks.ModelDoc2
    Dim sPartModelPath, sPartActiveConfig               As String
    
    Dim longErrors                                      As Long
    Dim longWarnings                                    As Long
    
    Dim sPrpPartNo                                      As String
    Dim sPrpPartName                                    As String
    Dim sPrpRawPart                                     As String
    Dim sPrpRawMaterial                                 As String
    Dim sPrpPartWeightKG                                As String
    Dim sPrpPartWeightLBS                               As String
    Dim sPrpRawMaterialEqui                             As String
    
    '------------------------------------------------------------------------------------
    
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set a = fs.CreateTextFile("c:\temp\logMWeldPartDrawingProp.txt", True)

    Set swApp = GetObject(, "SldWorks.Application")
    Set swModel = swApp.ActiveDoc
  
    If Not swModel Is Nothing Then
        
        Set swView = swModel.GetFirstView
        Set swView = swView.GetNextView
        
        If Not swView Is Nothing Then
            
            'Get active component body list of Cut_list_properties
            vComps = swView.GetVisibleComponents()
            
            Set swComp = vComps(0)
            sPartModelPath = swComp.GetPathName()
            sPartActiveConfig = swView.ReferencedConfiguration
            Debug.Print sPartModelPath & " " & sPartActiveConfig
            
            Set swPartModel = swApp.OpenDoc6(sPartModelPath, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_LoadModel, sPartActiveConfig, longErrors, longWarnings)
            swPartModel.ShowConfiguration2 sPartActiveConfig
            
            Dim vBodies As Variant
            vBodies = GetBodies(swView)

            Dim swBody As SldWorks.Body2
            Set swBody = vBodies(0)

            Debug.Print swView.name & " - " & swBody.name
            
            S_GetCutListProperties swPartModel, swBody

            swApp.CloseDoc sPartModelPath
            
            'Change drawing properties
            If Not IsEmpty(vPartCustomPropNames) Then
            
                Dim i As Long
                
                For i = LBound(vPartCustomPropNames) To UBound(vPartCustomPropNames)
                    
                    Dim sTempCustomPropName As String
                    sTempCustomPropName = vPartCustomPropNames(i)
                    
                    Select Case sTempCustomPropName
                        Case "SW-Part Number"
                            sPrpPartNo = vPartCustomPropResolvedVals(i)
                        Case "ITEM NAME"
                            sPrpPartName = vPartCustomPropResolvedVals(i)
                        Case "RAW PART"
                            sPrpRawPart = vPartCustomPropResolvedVals(i)
                        Case "RAW MATERIAL"
                            sPrpRawMaterial = vPartCustomPropResolvedVals(i)
                        Case "WEIGHT"
                            sPrpPartWeightKG = vPartCustomPropResolvedVals(i)
                        Case "WEIGHT LBS"
                            sPrpPartWeightLBS = vPartCustomPropResolvedVals(i)
                        Case "RAW EQUI MATERIAL"
                            sPrpRawMaterialEqui = vPartCustomPropResolvedVals(i)
                    End Select
                    
                Next i
                
            End If

            Edit_Properties swModel, "PartNo", sPrpPartNo
            Edit_Properties swModel, "PartName", sPrpPartName
            Edit_Properties swModel, "RawPart", sPrpRawPart
            Edit_Properties swModel, "RawMaterial", sPrpRawMaterial
            Edit_Properties swModel, "PartWeightKG", sPrpPartWeightKG & " KG"
            Edit_Properties swModel, "PartWeightLBS", sPrpPartWeightLBS & " LBS"
            Edit_Properties swModel, "RawMaterialEqui", sPrpRawMaterialEqui
                       
            swModel.ForceRebuild3 True
            
            Debug.Print "Done!"
            
        Else
            MsgBox "Please select view"
        End If
        
    Else
        MsgBox "Please open model"
    End If
    
End Sub

Function GetBodies(view As SldWorks.view) As Variant
    
    If view.IsFlatPatternView() Then
        
        Dim vComps As Variant
        vComps = view.GetVisibleComponents()
        
        'Flat pattern can be only created for a single body (either single body part or select body for multi-body part)
        Dim swComp As SldWorks.Component2
        Set swComp = vComps(0)
        
        Dim vFaces As Variant
        vFaces = view.GetVisibleEntities2(swComp, swViewEntityType_e.swViewEntityType_Face)
        
        Dim swFace As SldWorks.Face2
        Set swFace = vFaces(0)
        
        Dim swBodies(0) As SldWorks.Body2
        Set swBodies(0) = swFace.GetBody()
        
        GetBodies = swBodies
        
    Else
        GetBodies = view.Bodies
    End If
    
End Function

Sub Edit_Properties(swModel As SldWorks.ModelDoc2, sProperties As String, ByVal sPrpValue As String)

'    Set swApp = GetObject(, "SldWorks.Application")
'    If swApp Is Nothing Then
'        Set swApp = CreateObject("SldWorks.Application")
'    End If
    
    Dim swCustomProperties As SldWorks.CustomPropertyManager
    
    Dim Value_Expression As String
    Dim Evaluated_Value As String
    Dim wasResolved As Boolean
    Dim Islinked As Boolean
    Dim bStatus As Boolean
    
    '---------------------------------------
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

Sub S_GetCutListProperties(swModel As SldWorks.ModelDoc2, swBody As SldWorks.Body2)

    Dim swFeat As SldWorks.Feature
'    a.WriteLine "File: " & swModel.GetPathName
    
    Dim ConfigName As String
    ConfigName = swModel.ConfigurationManager.ActiveConfiguration.name
'    a.WriteLine "Active configuration name: " & ConfigName
    Set swFeat = swModel.FirstFeature
    TraverseFeatures swFeat, True, "Root Feature", swBody
'    a.Close
End Sub

Sub GetFeatureCustomProps(thisFeat As SldWorks.Feature)
    Dim CustomPropMgr As SldWorks.CustomPropertyManager
    Set CustomPropMgr = thisFeat.CustomPropertyManager
    Dim vCustomPropNames As Variant
    vCustomPropNames = CustomPropMgr.GetNames
    
    ReDim vPartCustomPropNames(UBound(vCustomPropNames))
    ReDim vPartCustomPropVals(UBound(vCustomPropNames))
    ReDim vPartCustomPropResolvedVals(UBound(vCustomPropNames))
    
    vPartCustomPropNames = vCustomPropNames
    
    If Not IsEmpty(vCustomPropNames) Then
'        a.WriteLine "               Cut-list custom properties:"
        Dim i As Long
        For i = LBound(vCustomPropNames) To UBound(vCustomPropNames)
            Dim CustomPropName As String
            CustomPropName = vCustomPropNames(i)
            Dim CustomPropType As Long
            CustomPropType = CustomPropMgr.GetType2(CustomPropName)
            Dim CustomPropVal As String
            Dim CustomPropResolvedVal As String
            CustomPropMgr.Get2 CustomPropName, CustomPropVal, CustomPropResolvedVal
'            a.WriteLine "                     Name: " & CustomPropName
'            a.WriteLine "                         Value: " & CustomPropVal
'            a.WriteLine "                         Resolved value: " & CustomPropResolvedVal
            
            vPartCustomPropVals(i) = CustomPropVal
            vPartCustomPropResolvedVals(i) = CustomPropResolvedVal
        Next i
    End If
End Sub
Sub DoTheWork(thisFeat As SldWorks.Feature, ParentName As String, swBody As SldWorks.Body2)
    Static InBodyFolder As Boolean
    Static BodyFolderType(5) As String
    Static BeenHere As Boolean
    Dim bAllFeatures As Boolean
    Dim bCutListCustomProps As Boolean
    Dim vSuppressed As Variant
    
    If Not BeenHere Then
        BodyFolderType(0) = "dummy"
        BodyFolderType(1) = "swSolidBodyFolder"
        BodyFolderType(2) = "swSurfaceBodyFolder"
        BodyFolderType(3) = "swBodySubFolder"
        BodyFolderType(4) = "swWeldmentSubFolder"
        BodyFolderType(5) = "swWeldmentCutListFolder"
        InBodyFolder = False
        BeenHere = True
        bAllFeatures = False
        bCutListCustomProps = False
    End If
    
    'Comment out next line to print information for just BodyFolders
    bAllFeatures = True 'True to print information about all features
    'Comment out next line if you do not want cut list's custom properties
    bCutListCustomProps = True
    Dim FeatType As String
    FeatType = thisFeat.GetTypeName
    If (FeatType = "SolidBodyFolder") And (ParentName = "Root Feature") Then
        InBodyFolder = True
    End If
    If (FeatType <> "SolidBodyFolder") And (ParentName = "Root Feature") Then
        InBodyFolder = False
    End If
    'Only consider the CutListFolders that are under SolidBodyFolder
    If (InBodyFolder = False) And (FeatType = "CutListFolder") Then
        'Skip the second occurrence of the CutListFolders during the feature traversal
        Exit Sub
    End If
    
    'Only consider the SubWeldFolder that are under the SolidBodyFolder
    If (InBodyFolder = False) And (FeatType = "SubWeldFolder") Then
        'Skip the second occurrence of the SubWeldFolders during the feature traversal
        Exit Sub
    End If
    Dim IsBodyFolder As Boolean
    If FeatType = "SolidBodyFolder" Or FeatType = "SurfaceBodyFolder" Or FeatType = "CutListFolder" Or FeatType = "SubWeldFolder" Or FeatType = "SubAtomFolder" Then
        IsBodyFolder = True
    Else
        IsBodyFolder = False
    End If
    
    If bAllFeatures And (Not IsBodyFolder) Then
'        a.WriteLine "Feature name: " & thisFeat.Name
'        a.WriteLine "   Feature type: " & FeatType
        vSuppressed = thisFeat.IsSuppressed2(swInConfigurationOpts_e.swThisConfiguration, Nothing)
        If IsEmpty(vSuppressed) Then
'            a.WriteLine "        Suppression failed"
        Else
'            a.WriteLine "        Suppressed"
        End If
    End If
    
    If IsBodyFolder Then
        Dim BodyFolder As SldWorks.BodyFolder
        Set BodyFolder = thisFeat.GetSpecificFeature2
        Dim BodyCount As Long
        BodyCount = BodyFolder.GetBodyCount
        If (FeatType = "CutListFolder") And (BodyCount < 1) Then
            'When BodyCount = 0, this cut list folder is not displayed in the
            'FeatureManager design tree, so skip it
            Exit Sub
        Else
'            a.WriteLine "Feature name: " & thisFeat.Name
            vSuppressed = thisFeat.IsSuppressed2(swInConfigurationOpts_e.swThisConfiguration, Empty)
            If IsEmpty(vSuppressed) Then
'                a.WriteLine "       Suppression failed"
            Else
'                a.WriteLine "       Suppressed"
            End If
        End If
        If Not bAllFeatures Then
'            a.WriteLine "Feature name: " & thisFeat.Name
            vSuppressed = thisFeat.IsSuppressed2(swInConfigurationOpts_e.swThisConfiguration, Empty)
            If IsEmpty(vSuppressed) Then
'                a.WriteLine "       Suppression failed"
            Else
'                a.WriteLine "       Suppressed"
            End If
        End If
        Dim BodyFolderTypeE As Long
        BodyFolderTypeE = BodyFolder.Type
'        a.WriteLine "        Body folder: " & BodyFolderType(BodyFolderTypeE)
'        a.WriteLine "        Body folder type: BodyFolderTypeE"
'        a.WriteLine "        Body count: " & BodyCount
        Dim vBodies As Variant
        vBodies = BodyFolder.GetBodies
        Dim i As Long
        If Not IsEmpty(vBodies) Then
            For i = LBound(vBodies) To UBound(vBodies)
                Dim Body As SldWorks.Body2
                Set Body = vBodies(i)
'                a.WriteLine "           Body name: " & Body.Name
                
                If Body.name = swBody.name Then
                    Dim sCutListName As String
                    sCutListName = thisFeat.name
                    GetFeatureCustomProps thisFeat
                End If
                
            Next i
        End If
    Else
        If bAllFeatures Then
'            a.WriteLine ""
        End If
    End If
    
    If (FeatType = "CutListFolder") Then
        'When BodyCount = 0, this cut list folder is not displayed
        'in the FeatureManager design tree, so skip it
        If BodyCount > 0 Then
            If bCutListCustomProps Then
                'Comment out this line if you do not want to
                'print the cut list folder's custom properties
                'If Body.Name = swBody.Name Then
'                    GetFeatureCustomProps thisFeat
                'End If
                
            End If
        End If
    End If
    
End Sub
Sub TraverseFeatures(thisFeat As SldWorks.Feature, isTopLevel As Boolean, ParentName As String, swBody As SldWorks.Body2)
    Dim curFeat As SldWorks.Feature
    Set curFeat = thisFeat
    While Not curFeat Is Nothing
        DoTheWork curFeat, ParentName, swBody
        Dim subfeat As SldWorks.Feature
        Set subfeat = curFeat.GetFirstSubFeature
        While Not subfeat Is Nothing
            TraverseFeatures subfeat, False, curFeat.name, swBody
            Dim nextSubFeat As SldWorks.Feature
            Set nextSubFeat = subfeat.GetNextSubFeature
            Set subfeat = nextSubFeat
            Set nextSubFeat = Nothing
        Wend
        Set subfeat = Nothing
        Dim nextFeat As SldWorks.Feature
        If isTopLevel Then
            Set nextFeat = curFeat.GetNextFeature
        Else
            Set nextFeat = Nothing
        End If
        Set curFeat = nextFeat
        Set nextFeat = Nothing
    Wend
End Sub


