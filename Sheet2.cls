VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Worksheet_SelectionChange(ByVal target As Range)
    
    'Dim CurrVal As Double
    'CurrVal = ActiveCell.VALUE
    'Debug.Print CurrVal
    'SB.LinkedCell = ActiveCell.Address
        
        
'    Set SB = ActiveSheet.Shapes("Scroll Bar 3").ControlFormat
'
''    SB.Value = 0
'    If WorksheetFunction.IsNumber(ActiveCell.Value) And Application.CutCopyMode = False Then
'        If (ActiveCell.Value >= 0) And (ActiveCell.Value <= 1) Then
'            SB.Value = ActiveCell.Value * 100
'        Else
'            SB.Value = 0
'        End If
'    Else
'        SB.Value = 0
'    End If

'    SB.Max = 100
'    SB.SmallChange = 10
'    SB.LargeChange = 20
    


End Sub

Sub ShowNames()
    Dim x&
    With ActiveWorkbook.ActiveSheet
        For x = 1 To .Shapes.count
            Debug.Print .Shapes(x).Name
    Next
    End With
End Sub

