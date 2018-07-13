VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub HighlightStrings()
'Updateby Extendoffice 20160704
Application.ScreenUpdating = False
Dim rng As Range
Dim cFnd As String
Dim xTmp As String
Dim x As Long
Dim m As Long
Dim y As Long
On Error Resume Next
cFnd = Selection.Value
y = Len(cFnd)
For Each rng In Range(Evaluate("INDIRECT(""$B$4"")"))
  
  rng.Font.ColorIndex = 1
  With rng
    m = UBound(Split(rng.Value, cFnd))
    If m > 0 Then
      xTmp = ""
      For x = 0 To m - 1
        xTmp = xTmp & Split(rng.Value, cFnd)(x)
        .Characters(start:=Len(xTmp) + 1, Length:=y).Font.ColorIndex = 3
        xTmp = xTmp & cFnd
      Next
    End If
  End With
Next rng
Application.ScreenUpdating = True
End Sub



Private Sub Worksheet_Calculate()
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)

'
''Updateby Extendoffice 20160704
'Application.ScreenUpdating = False
'Dim Rng As Range
'Dim cFnd As String
'Dim xTmp As String
'Dim x As Long
'Dim m As Long
'Dim y As Long
'On Error Resume Next
'cFnd = Selection.Value
'y = Len(cFnd)
'For Each Rng In Range(Evaluate("INDIRECT(""$B$4"")"))
'
'  Rng.Font.Italic = False
'  With Rng
'    m = UBound(Split(Rng.Value, cFnd))
'    If m > 0 Then
'      xTmp = ""
'      For x = 0 To m - 1
'        xTmp = xTmp & Split(Rng.Value, cFnd)(x)
'        .Characters(Start:=Len(xTmp) + 1, Length:=y).Font.Italic = True
'        xTmp = xTmp & cFnd
'      Next
'    End If
'  End With
'Next Rng
'Application.ScreenUpdating = True
End Sub



