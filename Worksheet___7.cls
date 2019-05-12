VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Worksheet___7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Worksheet_SelectionChange(ByVal target As range)


    'Dim CurrVal As Double
    'CurrVal = ActiveCell.VALUE
    'Debug.Print CurrVal
    On Error Resume Next
    Set SB = ActiveSheet.Shapes("Scroll Bar 1").ControlFormat
'    SB.Max = 100
'    SB.SmallChange = 1
'    SB.LargeChange = 10
'    SB.LinkedCell = ActiveCell.Address
'    SB.Value = 0
    If WorksheetFunction.IsNumber(ActiveCell.Value) And Application.CutCopyMode = False Then
        If (ActiveCell.Value >= 0) And (ActiveCell.Value <= 1) Then
            SB.Value = ActiveCell.Value * 100
        Else
            SB.Value = 0
        End If
    End If

    If Worksheets("ам╤у").range("A2").Value2 = False And Selection.Cells.Count = 1 And Application.CutCopyMode = False Then
        Set isect = Application.Intersect(range(Evaluate("INDIRECT(""$C$5"")")), Selection)
        range("$AG$5").Calculate
        range("$AH$5").Calculate
        If isect Is Nothing Then
            
        Else
                    'Updateby Extendoffice 20160704
            Application.ScreenUpdating = False
            
            Dim rng As range
            Dim cFnd As String
            Dim xTmp As String
            Dim x As Long
            Dim m As Long
            Dim y As Long
            On Error Resume Next
            cFnd = Selection.Value
            y = Len(cFnd)
            For Each rng In range(Evaluate("INDIRECT(""$C$5"")"))
              'Selection.Font.Size = 12
              rng.Font.Size = 12
              rng.Font.Bold = False
              With rng
                m = UBound(Split(rng.Value, cFnd))
                If m > 0 Then
                  xTmp = ""
                  For x = 0 To m - 1
                    xTmp = xTmp & Split(rng.Value, cFnd)(x)
                    .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.Bold = True
                    xTmp = xTmp & cFnd
                  Next
                End If
              End With
            Next rng
            Application.ScreenUpdating = True
        End If
    End If

    

    
End Sub


