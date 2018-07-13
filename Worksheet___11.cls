VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Worksheet___11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Activate()
    Range("$F$3").Calculate
    Range(Evaluate("INDIRECT(""$F$3"")")).Select
    ActiveWindow.ScrollRow = Selection.Row
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)

End Sub
