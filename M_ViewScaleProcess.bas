Attribute VB_Name = "M_ViewScaleProcess"
Public Enum E_ViewScaleProcessOption
    RePositionBaseOnSeedSheet = 0
    ReScaleBaseOnModelSize = 1
    RePositionBaseOnSeedFile = 2
End Enum

Const BASE_VIEWS_ONLY As Boolean = True

Dim swApp As SldWorks.SldWorks

Dim errors, warnings As Long

Sub ViewScaleProcess(eViewScaleProcessOption As E_ViewScaleProcessOption, swSeedDrawingFilePath, swNewDrawingFilePath As String)
    
    Dim scaleMap As Variant
    scaleMap = Array("6.4-8.0;0.05-100;1:100")
    
    Set swApp = GetObject(, "SldWorks.Application")
    'Set swApp = CreateObject("SldWorks.Application")
    
    Dim swDrawingDoc As SldWorks.DrawingDoc
    Dim swSeedDrawingDoc, swNewDrawingDoc As SldWorks.DrawingDoc
    Dim iErrors, iWarnings As Long

try:
    
    On Error GoTo catch
    
        Select Case eViewScaleProcessOption
            
            Case E_ViewScaleProcessOption.RePositionBaseOnSeedSheet
                
                Set swDrawingDoc = swApp.ActiveDoc
                
                If Not swDrawingDoc Is Nothing Then
                    RePositionViewsBySeedSheet swDrawingDoc
                Else
                    Err.Raise vbError, "", "Please open the drawing document"
                End If
              
             Case E_ViewScaleProcessOption.ReScaleBaseOnModelSize
                
                Set swDrawingDoc = swApp.ActiveDoc
                
                If Not swDrawingDoc Is Nothing Then
            '        RescaleViews swDraw, swDraw.GetCurrentSheet(), scaleMap
                Else
                    Err.Raise vbError, "", "Please open the drawing document"
                End If
             
             Case E_ViewScaleProcessOption.RePositionBaseOnSeedFile
                
'                If (swNewDrawingFilePath <> "" And swSeedDrawingFilePath <> "") Then
'
'                    Set swSeedDrawingDoc = swApp.OpenDoc6(swSeedDrawingFilePath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", iErrors, iWarnings)
'                    Set swNewDrawingDoc = swApp.OpenDoc6(swNewDrawingFilePath, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_LoadLightweight, "", iErrors, iWarnings)
'
'                    If Not swDrawingDoc Is Nothing Then
'                        RePositionViewsBySeedFile swSeedDrawingDoc, swNewDrawingDoc
'                    Else
'                        Err.Raise vbError, "", "Please open the drawing document"
'                    End If
         End Select
            
    GoTo finally
    
catch:
    'MsgBox Err.Description & " (" & Err.Number & ")", vbCritical
finally:
    'End
End Sub


Sub RePositionViewsBySeedFile(swSeedDrawingDoc, swNewDrawingDoc As SldWorks.DrawingDoc)

    Dim vSeedOutline() As Variant
    Dim vSeedPos() As Variant
    Dim nNumView As Long
    Dim bRet As Boolean
    
''' Seed Drawing parameters

    'Views array including sheets
    Dim vSeedSheets As Variant
    vSeedSheets = swSeedDrawingDoc.GetViews()
    
    Dim rSeedViewWidth(), rSeedViewHeight() As Variant
    Dim rSeedViewPosX(), rSeedViewPosY() As Variant
    
    Dim i As Integer
    
    'For i = UBound(vSeedSheets) To 0 Step -1
    'Get seed only
    i = UBound(vSeedSheets)
    
        'Views array of a sheet
        Dim vSeedViews As Variant
        vSeedViews = vSeedSheets(i)
        
        'Sheet is the first of sheet view array
        Dim swSeedSheetView As SldWorks.view
        Set swSeedSheetView = vViews(0)
        
        Dim j As Integer
        
        Dim iNumViewOfSeedSheet As Integer
        iNumViewOfSeedSheet = UBound(vViews)
        
        ReDim vOutline(iNumViewOfSeedSheet)
        ReDim vPos(iNumViewOfSeedSheet)
  
        ReDim Preserve rSeedViewWidth(iNumViewOfSeedSheet), rSeedViewHeight(iNumViewOfSeedSheet)
        ReDim Preserve rSeedViewPosX(iNumViewOfSeedSheet), rSeedViewPosY(iNumViewOfSeedSheet)
        
        For j = 1 To UBound(vViews)
        
            Debug.Print "Select view no. " & j & " of sheet no. " & i
            
            Dim swSeedView As SldWorks.view
            Set swSeedView = vSeedViews(j)
            
            Dim viewType As Integer
            viewType = swSeedView.Type
            
            'If viewType <> swDrawingViewTypes_e.swDrawingDetailView And viewType <> swDrawingViewTypes_e.swDrawingSectionView Then

                vSeedOutline(j) = swSeedView.GetOutline
                vSeedPos(j) = swView.Position
                Debug.Print "View = " + swSeedView.GetName2
                Debug.Print "  Pos = (" & vSeedPos(j)(0) * 1000# & ", " & vSeedPos(j)(1) * 1000# & ") mm"
                Debug.Print "  Min = (" & vSeedOutline(j)(0) * 1000# & ", " & vSeedOutline(j)(1) * 1000# & ") mm"
                Debug.Print "  Max = (" & vSeedOutline(j)(2) * 1000# & ", " & vSeedOutline(j)(3) * 1000# & ") mm"
                
                Dim rViewWidth As Double
                Dim rViewHeight As Double
                GetViewGeometrySize swView, rViewWidth, rViewHeight
                
                Dim rViewScale As Double
                GetViewScale swView, rViewScale
                
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

'''Set to new drawing file
    'Views array including sheets
    Dim vSheets As Variant
    vSheets = swNewDrawingDoc.GetViews()
    
    Dim iIndex As Integer
    
    For iIndex = UBound(vSheets) To 0 Step -1
        
        Dim nextViewIndex As Integer
        nextViewIndex = 0
        
        'Views array of a sheet
        Dim vViews As Variant
        vViews = vSheets(i)
        
        'Sheet is the first of sheet view array
        Dim swSheetView As SldWorks.view
        Set swSheetView = vViews(0)
        
        Dim j As Integer
        
        Dim iNumViewOfSheet As Integer
        iNumViewOfSheet = UBound(vViews)
        
        ReDim vOutline(iNumViewOfSheet)
        ReDim vPos(iNumViewOfSheet)
        
        Dim swSketchPoint As Object
        
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
                
                End If
                
                
            'End If
            
        Next
        
    Next
    swModel.GraphicsRedraw2

    swDrawingDoc.EditRebuild
    
End Sub


Sub RePositionViewsBySeedSheet(swDrawingDoc As SldWorks.DrawingDoc)

    Dim vOutline() As Variant
    Dim vPos() As Variant
    Dim nNumView As Long
    Dim bRet As Boolean
    Dim swModel As Object
      
    'Views array including sheets
    Dim vSheets As Variant
    vSheets = swDrawingDoc.GetViews()
    
    Dim rSeedViewWidth(), rSeedViewHeight() As Variant
    Dim rSeedViewPosX(), rSeedViewPosY() As Variant
    
    Dim i As Integer
    
    For i = UBound(vSheets) To 0 Step -1
        
        Dim nextViewIndex As Integer
        nextViewIndex = 0
        
        'Views array of a sheet
        Dim vViews As Variant
        vViews = vSheets(i)
        
        'Sheet is the first of sheet view array
        Dim swSheetView As SldWorks.view
        Set swSheetView = vViews(0)
        
        Dim j As Integer
        
        Dim iNumViewOfSheet As Integer
        iNumViewOfSheet = UBound(vViews)
        
        ReDim vOutline(iNumViewOfSheet)
        ReDim vPos(iNumViewOfSheet)
        
        Dim swSketchPoint As Object
        
        ReDim Preserve rSeedViewWidth(iNumViewOfSheet), rSeedViewHeight(iNumViewOfSheet)
        ReDim Preserve rSeedViewPosX(iNumViewOfSheet), rSeedViewPosY(iNumViewOfSheet)
        
'        Set swModel = swDrawingDoc
'        Set swSketchPoint = swModel.CreatePoint(0.1381875 * 60, 0.13935 * 60, 0#)
        
        For j = 1 To UBound(vViews)
        
            Debug.Print "Select view no. " & j & " of sheet no. " & i
            
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
                
                If (i = UBound(vSheets)) Then
                    
                    'Get seed view parameter
                    rSeedViewWidth(j) = rViewWidth
                    rSeedViewHeight(j) = rViewHeight
                    rSeedViewPosX(j) = vPos(j)(0)
                    rSeedViewPosY(j) = vPos(j)(1)
                Else
                
                    'Calculate and reposition of populated view
                    vPos(j)(0) = rSeedViewPosX(j) - (rViewWidth - rSeedViewWidth(j)) / 2 / rViewScale
                    vPos(j)(1) = rSeedViewPosY(j) - (rViewHeight - rSeedViewHeight(j)) / 2 / rViewScale
                    Debug.Print "  New position X = " & vPos(j)(0); ", Y = " & vPos(j)(0)
                
                    swView.Position = vPos(j)
                
                End If
                
                
            'End If
            
        Next
        
    Next
    swModel.GraphicsRedraw2

    swDrawingDoc.EditRebuild
    
End Sub


Sub RescaleViews(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet, scaleMap As Variant)
    
    Dim vViews As Variant
    vViews = GetSheetViews(draw, sheet)
    
    Dim i As Integer
    
    For i = 0 To UBound(vViews)
        
        Dim swView As SldWorks.view
        Set swView = vViews(i)
        
        Dim width As Double
        Dim height As Double
        GetViewGeometrySize swView, width, height
        
        Debug.Print swView.name & " : " & width & " x " & height
        
        Dim j As Integer
        
        For j = 0 To UBound(scaleMap)
            
            Dim minWidth As Double
            Dim maxWidth As Double
            Dim minHeight As Double
            Dim maxHeight As Double
            Dim viewScale As Variant
            
            ExtractParameters CStr(scaleMap(j)), minWidth, maxWidth, minHeight, maxHeight, viewScale
            
            If width >= minWidth And width <= maxWidth And height >= minHeight And height <= maxHeight Then
                Debug.Print swView.name & " matches " & CStr(scaleMap(j))
                If Not BASE_VIEWS_ONLY Or swView.GetBaseView() Is Nothing Then
                    Debug.Print "Setting scale of " & swView.name & " to " & viewScale(0) & ":" & viewScale(1)
                    swView.ScaleRatio = viewScale
                Else
                    Debug.Print "Skipping " & swView.name & " view as it is not a base view"
                End If
                
            Else
                Debug.Print swView.name & " doesn't match " & CStr(scaleMap(j))
            End If
            
        Next
        
    Next
    
    draw.EditRebuild
    
End Sub

Function GetSheetViews(draw As SldWorks.DrawingDoc, sheet As SldWorks.sheet) As Variant

    Dim vSheets As Variant
    vSheets = draw.GetViews()
    
    Dim i As Integer
    
    For i = 0 To UBound(vSheets)
    
        Dim vViews As Variant
        vViews = vSheets(i)
        
        Dim swSheetView As SldWorks.view
        Set swSheetView = vViews(0)
        
        If UCase(swSheetView.name) = UCase(sheet.GetName()) Then
            
            If UBound(vViews) > 0 Then
                
                Dim swViews() As SldWorks.view
                
                ReDim swViews(UBound(vViews) - 1)
                
                Dim j As Integer
                
                For j = 1 To UBound(vViews)
                    Set swViews(j - 1) = vViews(j)
                Next
                
                GetSheetViews = swViews
                Exit Function
                
            End If
            
        End If
        
    Next
    
End Function

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

Sub ExtractParameters(params As String, ByRef minWidth As Double, ByRef maxWidth As Double, ByRef minHeight As Double, ByRef maxHeight As Double, ByRef viewScale As Variant)

    Dim vParamsData As Variant
    vParamsData = Split(params, ";")
    
    ExtractSizeBounds CStr(vParamsData(0)), minWidth, maxWidth
    ExtractSizeBounds CStr(vParamsData(1)), minHeight, maxHeight
    
    Dim scaleData As Variant
    scaleData = Split(vParamsData(2), ":")
    
    Dim dViewScale(1) As Double
    dViewScale(0) = CDbl(Trim(scaleData(0)))
    dViewScale(1) = CDbl(Trim(scaleData(1)))
    
    viewScale = dViewScale
    
End Sub

Sub ExtractSizeBounds(boundParam As String, ByRef min As Double, ByRef max As Double)
    
    If Trim(boundParam) = "*" Then
        min = 0
        max = 1000000
    Else
        Dim minMax As Variant
        minMax = Split(boundParam, "-")
        min = CDbl(Trim(minMax(0)))
        max = CDbl(Trim(minMax(1)))
    End If
    
End Sub

