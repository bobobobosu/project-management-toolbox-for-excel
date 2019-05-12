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
    range(Evaluate("INDIRECT(""$C$1"")")).Calculate
End Sub

Private Sub Worksheet_Deactivate()
    range(Evaluate("INDIRECT(""$C$1"")")).Calculate
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As range)

End Sub
