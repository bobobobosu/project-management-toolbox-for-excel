Attribute VB_Name = "Trash"
Function TableArrayFormula()
    TableArrayFormula = returnoriginal(range("Table13[Column2]"))
End Function
Function erwwgsfd(v As Variant)
    erwwgsfd = vv
End Function
Sub PasteNonBlank()
    
    For i = 1 To Selection.Cells.Count
        If Selection.Cells(i).Value2 <> vbNullString Then
            range("表格2[起始百分比]").Cells(i) = Selection.Cells(i).Value2
        End If
        
    Next
End Sub

Sub ConvertFormulaArray()
    For Each cell In Selection
        cell.FormulaArray = cell.formula
    Next
End Sub

Sub deleteFormula()
    For Each cell In Selection
        If cell.HasFormula Then
            cell.formula = ""
        End If
    Next
End Sub
