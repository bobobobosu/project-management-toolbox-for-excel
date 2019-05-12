Attribute VB_Name = "Module18"
Sub SetZero()
        Selection.Value2 = 0
     Debug.Print ActiveSheet.CodeName
    If ActiveSheet.CodeName = "Worksheet___7" Then
        range(Evaluate("INDIRECT(""$B$4"")")).Calculate
    End If
End Sub
