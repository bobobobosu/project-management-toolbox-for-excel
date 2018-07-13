Attribute VB_Name = "Module30"
Sub TimeFunc1_Click()
    For Each cell In Selection
        cell.Value = "=" + Evaluate("INDIRECT(""$AF$4"")")
    Next cell
End Sub
Sub TimeFunc2_Click()
    For Each cell In Selection
        Selection.Value = "=" + Evaluate("INDIRECT(""$AF$3"")")
    Next cell
End Sub

