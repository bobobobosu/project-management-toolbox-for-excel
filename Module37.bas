Attribute VB_Name = "Module37"
Sub SetEndDate()
    Dim copySelection As range
    Set copySelection = Evaluate("表格2[[#This Row], [預計耗時]:[預計耗時]]")
    Set r = Evaluate("表格2[[#This Row], [End Date]:[End Date]]")
    FirstCellAddress = r.address
    If range(FirstCellAddress).NumberFormat = "m/d/yyyy" Or range(FirstCellAddress).NumberFormat = "h:mm:ss;@" Or range(FirstCellAddress).NumberFormat = "m/d/yy h:mm;@" Then
        Dim FirstCell2 As Variant
        FirstCell2 = InputBox("Date Value", "Please Enter Date Value", Format(range(FirstCellAddress).Value2, "m/d/yy"))
        Dim FirstCell3 As Variant
        FirstCell3 = InputBox("Time Value", "Please Enter Time Value", Format(range(FirstCellAddress).Value2, "h:mm:ss;@"))
        
        If FirstCell2 <> vbNullString Then
            For Each selected In copySelection
                Evaluate("趨勢!$D$2").Value2 = dateValue(FirstCell2) + TimeValue(Format(selected.Value2, "h:mm:ss;@"))
                Evaluate("趨勢!$D$2").Copy
                selected.Select
                selected.PasteSpecial Paste:=xlValues, Transpose:=False
            Next selected
        End If
        
        If FirstCell3 <> vbNullString Then
            For Each selected In copySelection
                Evaluate("趨勢!$D$2").Value2 = dateValue(Format(selected.Value2, "m/d/yy")) + TimeValue(FirstCell3)
                Evaluate("趨勢!$D$2").Copy
                selected.Select
                selected.PasteSpecial Paste:=xlValues, Transpose:=False
            Next selected
        End If
        
        copySelection.Select
        Dim startD As Double
        Dim endD As Double
        startD = Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]").Value2
        endD = copySelection.Value2
        copySelection.Value2 = endD - startD
        
    End If
End Sub
