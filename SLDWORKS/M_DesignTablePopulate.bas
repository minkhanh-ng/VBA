Attribute VB_Name = "M_DesignTablePopulate"
Option Explicit
Dim wb As Workbook
Dim ws As Worksheet
    
Sub S_PopulateDesignTable(sCurrentPath, sWorkbookPath, sFromPath, sUsedRange As String)
    Dim i As Long
    
    Dim sCurrentWorkbookPath As String
    Dim wbUsedWb As Workbook
    Dim wsUsedWs As Worksheet
    
    Dim sTargetPath As String
    Dim arFromFiles(6) As Variant
    Dim arToFiles(6) As Variant
    Dim Cell As Range
    Dim bDone As Boolean
    
    Dim sTempFromPath, sTempToPath As String
    
'////////////////////////////////////////////////////////////////
       
    sCurrentWorkbookPath = sWorkbookPath
    Workbooks.Open FileName:=sCurrentWorkbookPath
    Set wbUsedWb = ActiveWorkbook
    Set wsUsedWs = wbUsedWb.Sheets("Sheet1")
    
    arFromFiles(0) = sFromPath & "\" & "DesignTable__ XXX-40-CUSTOM BLUE ORC PANEL RF PANEL R2.xlsx"
    arFromFiles(1) = sFromPath & "\" & "DesignTable__ XXX-40-LAMINATIONS.xlsx"
    arFromFiles(2) = sFromPath & "\" & "DesignTable__ XXX-40-RINGS.xlsx"
    arFromFiles(3) = sFromPath & "\" & "DesignTable__ XXX-40-WALL.xlsx"
    arFromFiles(4) = sFromPath & "\" & "DesignTable__ XXX-40-ROOFSHEETS.xlsx"
    arFromFiles(5) = sFromPath & "\" & "DesignTable__ XXX-40-ROOFSHEET-ASSY.xlsx"
    arFromFiles(6) = sFromPath & "\" & "DesignTable__ XXX-40-TANK.xlsx"
    
    'Use variables in the FileCopy statement
    Dim xlobj As Object
    Dim sExcelFilePath As String
    
    Set xlobj = CreateObject("Scripting.FileSystemObject")
    'object.copyfile,source,destination,file overright(True is default)
    
    For Each Cell In wsUsedWs.Range(sUsedRange)
    
        sTargetPath = sCurrentPath & "\" & Cell.Value & "\3D FILES"
             
        'For i = 0 To UBound(arFromFiles)
        i = 4
            arToFiles(i) = Replace(arFromFiles(i), sFromPath, sTargetPath)
            arToFiles(i) = Replace(arToFiles(i), "XXX-40", Cell.Value)
            sTempFromPath = arFromFiles(i)
            sTempToPath = arToFiles(i)
            
            xlobj.CopyFile sTempFromPath, sTempToPath, True
            'CopyFolderWithErrorHandling sTempFromPath, sTempToPath
            
            Debug.Print arToFiles(i) & " created!"
        
            'Change excel files' configs as per modes' name
            sExcelFilePath = sTempToPath
            
'            'Delete Unsed rows in ROOFSHEET table
'            If (i = 4 Or i = 5) Then
'                bDone = F_DeleteRowOfText(sExcelFilePath, "Sheet1", "A3", Cell.Value, "<>", 0)
'                Debug.Print "   Delete unused columns in " & sExcelFilePath & " success? " & bDone
'            End If
'
'            bDone = F_ReplaceTextInSheet(sExcelFilePath, "XXX-40", Cell.Value)
'            Debug.Print "   Delete unused columns in " & sExcelFilePath & " success? " & bDone
            
            'Manipulating data of ROOFSHEET DESIGN TABLE
            Select Case i
            Case 4
                S_RoofSheetTableData sExcelFilePath, "XXX-40", Cell.Value
            Case Else
                bDone = F_ReplaceTextInSheet(sExcelFilePath, "XXX-40", Cell.Value)
                
                'Delete Unused rows of XXX-15
                If Cell.Value = "XXX-15" Then
                    Dim col As String
                    
                    bDone = F_DeleteRowOfText(sExcelFilePath, "Sheet1", "A3", "-6400-", "=", 1)
                    bDone = F_DeleteRowOfText(sExcelFilePath, "Sheet1", "A3", "-5700-", "=", 1)
                    bDone = F_DeleteRowOfText(sExcelFilePath, "Sheet1", "A3", "-5000-", "=", 1)
                    bDone = F_DeleteRowOfText(sExcelFilePath, "Sheet1", "A3", "-4300-", "=", 1)
                    bDone = F_DeleteRowOfText(sExcelFilePath, "Sheet1", "A3", "-3600-", "=", 1)
                    bDone = F_DeleteRowOfText(sExcelFilePath, "Sheet1", "A3", "-2900-", "=", 1)
                    
                    Debug.Print "   Delete unused rows in " & sExcelFilePath & " success? " & bDone
                End If
            End Select
            
            Debug.Print "   Edit " & arToFiles(i) & " success? " & bDone
            Debug.Print " "
        'Next i
    
    Next Cell
    
    Set xlobj = Nothing
End Sub

Function F_ReplaceModel(sExcelFilePath, strToChange, strNew As String) As Boolean

    Dim rangeAllFile As Range
    
    Set wb = Workbooks.Open(sExcelFilePath)
    Set ws = wb.Sheets("Sheet1")
    
    ws.Range("A2", Range("A2").End(xlDown).End(xlToRight)).Replace strToChange, strNew
    wb.Save
    wb.Close
    ReplaceModel = True
    
End Function

Function F_ReplaceTextInSheet(sExcelFilePath, oldText, newText As String) As Boolean

    Dim ws As Worksheet
    Dim searchRange As Range
    
    Set wb = Workbooks.Open(sExcelFilePath)
    Set ws = wb.Sheets("Sheet1")
        
    ' Loop through all worksheets in the workbook
    Set searchRange = ws.usedRange
        
    ' Replace the old text with the new text
    searchRange.Replace What:=oldText, Replacement:=newText, LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    wb.Save
    wb.Close
    
End Function

Function F_DeleteRowOfText(sExcelFilePath As String, usedWorksheet As String, startCell, textToDelete As String, sCriteriaType As String, iRowsOffset As Integer) As Boolean
    
    Dim usedRange As Range
    
    Set wb = Workbooks.Open(sExcelFilePath)
    Set ws = wb.Sheets(usedWorksheet)
    
    Set usedRange = ws.Range(startCell, Range(startCell).End(xlDown))
    
    If ws.AutoFilterMode Or ws.FilterMode Then
        ws.AutoFilter.ShowAllData
    End If
    
    With usedRange
        .AutoFilter Field:=1, Criteria1:=sCriteriaType & "*" & textToDelete & "*"
        .Offset(iRowsOffset, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With
    
    ws.AutoFilterMode = False
    F_DeleteRowOfText = True
    
    wb.Save
    wb.Close
    
End Function


Sub S_RoofSheetTableData(sExcelFilePath, sCellValueToDelete, sNewTankModel As String)

    Dim wbDataWB As Workbook
    Dim wsSourceSheet, wsDestinationSheet As Worksheet
    Dim tblTable As ListObject
    Dim Cell As Range
    Dim iPos As Integer
    Dim iNextFreeRow As Long
    
'    sExcelFilePath = "C:\Users\khanh.nguyen\OneDrive - xxx\Desktop\PackNGo\E-XXX\XXX-25\3D FILES\DesignTable__ XXX-25-ROOFSHEETS.xlsx"
'    sCellValueToDelete = "XXX-40"
    
    Set wbDataWB = Workbooks.Open(sExcelFilePath)
    Set wsSourceSheet = wbDataWB.Sheets("SHEET BUILD TABLE")
    Set tblTable = wsSourceSheet.ListObjects("ROOF_SHEETS_ALL_NAME")
    
    Set wsDestinationSheet = wbDataWB.Sheets("Sheet1")
    
    'Clear destination data column before copy and paste
    wsDestinationSheet.Range("A4:A" & wsDestinationSheet.Cells(wsDestinationSheet.Rows.count, "A").End(xlUp).Row).ClearContents
    
    'With Sheets("SHEET BUILD TABLE") 'Sheet with data to check for value
        ' loop column H untill last cell with value (not entire column)
        For Each Cell In tblTable.ListColumns(1).DataBodyRange
            iPos = InStr(1, Cell.Value, sNewTankModel, vbTextCompare)
            If iPos > 0 Then
                iNextFreeRow = wsDestinationSheet.Cells(wsDestinationSheet.Rows.count, "A").End(xlUp).Row + 1
                
                 'get the next empty row to paste data to
                '.Range("A" & Cell.Row & ",B" & Cell.Row & ",C" & Cell.Row & ",F" & Cell.Row & "," & Cell.Address).Copy Destination:=Sheets("Sheet1").Range("A" & NextFreeRow)
                '.Range("AI" & Cell.Row & "," & Cell.Address).Copy Destination:=Sheets("Sheet1").Range("A" & NextFreeRow)
                '.Range("AI" & Cell.Row & "," & Cell.Address).Copy
                'Sheets("Sheet1").Range("A" & NextFreeRow).PasteSpecial Paste:=xlPasteValues
                
                Cell.Copy
                wsDestinationSheet.Range("A" & iNextFreeRow).PasteSpecial Paste:=xlPasteValues
            End If
        Next Cell
        
        For Each Cell In tblTable.ListColumns(1).DataBodyRange
            iPos = InStr(1, Cell.Value, sNewTankModel, vbTextCompare)
            If iPos > 0 Then
                iNextFreeRow = wsDestinationSheet.Cells(wsDestinationSheet.Rows.count, "A").End(xlUp).Row + 1
                
                Cell.Copy
                wsDestinationSheet.Range("A" & iNextFreeRow).PasteSpecial Paste:=xlPasteValues
                wsDestinationSheet.Cells(iNextFreeRow, 1).Value = Cell.Value & "-MIR"
            End If
        Next Cell
    'End With
    
    wbDataWB.Save
    wbDataWB.Close
    
End Sub
