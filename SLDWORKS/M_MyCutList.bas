Attribute VB_Name = "M_MyCutList"
Option Explicit
'Const WeldmentTableTemplate As String = "C:\Users\khanh.nguyen\OneDrive - xxx\SOLIDWORKS\CutListTableTemplate.sldwldtbt"

Sub S_MyCutList(WeldmentTableTemplate As String)
  Dim swApp As SldWorks.SldWorks
  Dim oDrawing As DrawingDoc
  Dim swView As view
  Dim WMTable As SldWorks.WeldmentCutListAnnotation
  Dim sPartActiveConfig As String

  Set swApp = GetObject(, "SldWorks.Application")
  Set oDrawing = swApp.ActiveDoc
  Set swView = oDrawing.GetFirstView
  Set swView = swView.GetNextView

  ' Insert the weldment cut list table
  sPartActiveConfig = swView.ReferencedConfiguration
  Set WMTable = swView.InsertWeldmentTable(False, 0.378, 0.291, swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_TopRight, sPartActiveConfig, WeldmentTableTemplate)

End Sub

