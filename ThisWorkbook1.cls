VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal target As Range)

    Application.ScreenUpdating = False
'    If Target.Cells.Count > 1 Then Exit Sub
'    Application.ScreenUpdating = False
'    'Clear all cell colours
'    Cells.Interior.ColorIndex = 0
'    With Target
'    'Row and column highlightingfor the active cell
'    .EntireRow.Interior.ColorIndex = 50
'    .EntireColumn.Interior.ColorIndex = 50
'    End With
'    Application.ScreenUpdating = True
    
End Sub



