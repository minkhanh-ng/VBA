Attribute VB_Name = "M_ViewSequentialLabeling"
Sub S_ViewSequentialLabeling()
    
    Dim bStatus As Boolean
    Dim sUsedRange As String
    
    Dim swApp As SldWorks.SldWorks
    Dim swDraw As SldWorks.DrawingDoc
    Dim swModel As SldWorks.ModelDoc2

    Dim i As Long
    
    Dim sCurrentWorkbookPath As String
    Dim wbUsedWb As Workbook
    Dim wsUsedWs As Worksheet
    Dim Cell As Range
    Dim iCellIndx As Long

'////////////////////////////////////////////////////////////////

    ' Set the path to your Excel file
    
    'sCurrentWorkbookPath = sWorkbookPath
    'Workbooks.Open FileName:=sCurrentWorkbookPath
    'Set wbUsedWb = Workbooks.Open(sCurrentWorkbookPath)
    Set wbUsedWb = ThisWorkbook
    Set wsUsedWs = wbUsedWb.Sheets("Sheet1")
  
'-------------------------------------------------------

    sUsedRange = "K2:K200"
    
 ''' Initialize SOLIDWORKS
    Set swApp = GetObject(, "SldWorks.Application")
 
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    swApp.Visible = True

    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If

    Dim vSheets As Variant
    'Return array of sheets, consist of array of its views
    vSheets = swDraw.GetViews 'vSheets(0)(0) = Sheet1 / vSheets(0)(1) = S1_View1 / vSheets(0)(2) = S1_View2 / vSheets(1)(0) = Sheet2
    
    Dim sNewViewLabel As String
   
    Set Cell = wsUsedWs.Range(sUsedRange)(1, 1)
    sNewViewLabel = Cell.Value
    iCellIndx = 0
    
    For i = 0 To UBound(vSheets)
        Dim nextViewIndex As Integer
        nextViewIndex = 0

        Dim vArrayViewsAndSheets As Variant
        vArrayViewsAndSheets = vSheets(i)

        Dim swSheet As SldWorks.view

        Set swSheet = vArrayViewsAndSheets(0) 'Each view0 is sheet itself
        
        Dim sSheetName As String
        sSheetName = swSheet.GetName2()
        
        Debug.Print "Sheet: " & sSheetName

        Dim j As Integer

        For j = 1 To UBound(vArrayViewsAndSheets)
            
            bStatus = False
            bIsLabelSame = False
            
            Dim swView As SldWorks.view
            Set swView = vArrayViewsAndSheets(j)

            Dim viewType As Integer
            viewType = swView.Type

            Dim swDetailCircle As SldWorks.DetailCircle
            Dim swSectionView As SldWorks.DrSection

            Debug.Print "View: " & swView.name
            
            If viewType = swDrawingViewTypes_e.swDrawingDetailView Then

                Set swDetailCircle = swView.GetDetail
                Debug.Print "Detail view:"
                'Debug.Print "  Selected: " & swDetailCircle.Select(True, Nothing)
                Debug.Print "  Label: " & swDetailCircle.GetLabel
                
                bStatus = swDetailCircle.SetLabel(sNewViewLabel)
                
                If sNewViewLabel = swDetailCircle.GetLabel Then
                    bIsLabelSame = True
                Else
                    bIsLabelSame = False
                End If
                                
                If False = bStatus Then
                    Debug.Print "   Failed to change view label"
                Else
                    Debug.Print "   Change view label to " & sNewViewLabel
                End If
                
            End If
            
            If viewType = swDrawingViewTypes_e.swDrawingSectionView Then

                Set swSectionView = swView.GetSection
                Debug.Print "Section view:"
                Debug.Print "  Label: " & swSectionView.GetLabel
                
                bStatus = swSectionView.SetLabel2(sNewViewLabel)
                  
                If sNewViewLabel = swSectionView.GetLabel Then
                    bIsLabelSame = True
                Else
                    bIsLabelSame = False
                End If
                                
                If False = bStatus Then
                    Debug.Print "   Failed to change view label"
                Else
                    Debug.Print "   Change view label to " & sNewViewLabel
                End If
                
            End If
            
            If bStatus Or bIsLabelSame Then
                iCellIndx = iCellIndx + 1
                sNewViewLabel = Cell.Offset(iCellIndx, 0).Value
            End If
            
        Next

    Next
    
    'Rebuild File
    swModel.EditRebuild3 'Stoplight or [Ctrl]+B
    'swModel.ForceRebuild '[Ctrl]+Q
    
End Sub

Sub S_ViewSequentialLabelingPreserve(ByRef longCounter As Long)
    
    Dim bStatus As Boolean
    Dim sUsedRange As String
    
    Dim swApp As SldWorks.SldWorks
    Dim swDraw As SldWorks.DrawingDoc
    Dim swModel As SldWorks.ModelDoc2

    Dim i As Long
    
    Dim sCurrentWorkbookPath As String
    Dim wbUsedWb As Workbook
    Dim wsUsedWs As Worksheet
    Dim Cell As Range
    Dim iCellIndx As Long

'////////////////////////////////////////////////////////////////

    ' Set the path to your Excel file
    
    'sCurrentWorkbookPath = sWorkbookPath
    'Workbooks.Open FileName:=sCurrentWorkbookPath
    'Set wbUsedWb = Workbooks.Open(sCurrentWorkbookPath)
    Set wbUsedWb = ThisWorkbook
    Set wsUsedWs = wbUsedWb.Sheets("Sheet1")
  
'-------------------------------------------------------

    sUsedRange = "K2:K200"
    
 ''' Initialize SOLIDWORKS
    Set swApp = GetObject(, "SldWorks.Application")
 
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    swApp.Visible = True

    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    
    If swModel Is Nothing Then
        MsgBox "Please open the drawing"
        End
    End If

    Dim vSheets As Variant
    'Return array of sheets, consist of array of its views
    vSheets = swDraw.GetViews 'vSheets(0)(0) = Sheet1 / vSheets(0)(1) = S1_View1 / vSheets(0)(2) = S1_View2 / vSheets(1)(0) = Sheet2
    
    Dim sNewViewLabel As String
   
    Set Cell = wsUsedWs.Range(sUsedRange)(longCounter, 1)
    sNewViewLabel = Cell.Value
    iCellIndx = longCounter
    
    For i = 0 To UBound(vSheets)
        Dim nextViewIndex As Integer
        nextViewIndex = 0

        Dim vArrayViewsAndSheets As Variant
        vArrayViewsAndSheets = vSheets(i)

        Dim swSheet As SldWorks.view

        Set swSheet = vArrayViewsAndSheets(0) 'Each view0 is sheet itself
        
        Dim sSheetName As String
        sSheetName = swSheet.GetName2()
        
        Debug.Print "Sheet: " & sSheetName

        Dim j As Integer

        For j = 1 To UBound(vArrayViewsAndSheets)
            
            bStatus = False
            bIsLabelSame = False
            
            Dim swView As SldWorks.view
            Set swView = vArrayViewsAndSheets(j)

            Dim viewType As Integer
            viewType = swView.Type

            Dim swDetailCircle As SldWorks.DetailCircle
            Dim swSectionView As SldWorks.DrSection

            Debug.Print "View: " & swView.name
            
            If viewType = swDrawingViewTypes_e.swDrawingDetailView Then

                Set swDetailCircle = swView.GetDetail
                Debug.Print "Detail view:"
                'Debug.Print "  Selected: " & swDetailCircle.Select(True, Nothing)
                Debug.Print "  Label: " & swDetailCircle.GetLabel
                
                bStatus = swDetailCircle.SetLabel(sNewViewLabel)
                
                If sNewViewLabel = swDetailCircle.GetLabel Then
                    bIsLabelSame = True
                Else
                    bIsLabelSame = False
                End If
                                
                If False = bStatus Then
                    Debug.Print "   Failed to change view label"
                Else
                    Debug.Print "   Change view label to " & sNewViewLabel
                End If
                
            End If
            
            If viewType = swDrawingViewTypes_e.swDrawingSectionView Then

                Set swSectionView = swView.GetSection
                Debug.Print "Section view:"
                Debug.Print "  Label: " & swSectionView.GetLabel
                
                bStatus = swSectionView.SetLabel2(sNewViewLabel)
                  
                If sNewViewLabel = swSectionView.GetLabel Then
                    bIsLabelSame = True
                Else
                    bIsLabelSame = False
                End If
                                
                If False = bStatus Then
                    Debug.Print "   Failed to change view label"
                Else
                    Debug.Print "   Change view label to " & sNewViewLabel
                End If
                
            End If
            
            If bStatus Or bIsLabelSame Then
                iCellIndx = iCellIndx + 1
                sNewViewLabel = Cell.Offset(iCellIndx, 0).Value
            End If
            
        Next

    Next
    
    longLabelCounter = iCellIndx
    
    'Rebuild File
    swModel.EditRebuild3 'Stoplight or [Ctrl]+B
    'swModel.ForceRebuild '[Ctrl]+Q
    
End Sub

