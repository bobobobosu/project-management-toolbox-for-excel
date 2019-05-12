VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    range(Evaluate("INDIRECT(""$C$2"")")).Calculate
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As range)
    Dim nowRow As String
    On Error Resume Next
    Worksheets("ам╤у").range("B2").Value2 = samerowsOf(Selection, Evaluate("D5#"))
End Sub
