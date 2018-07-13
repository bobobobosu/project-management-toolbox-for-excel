Attribute VB_Name = "Module24"
Sub AutoUpdate2_Click()
    Call AutoCalculate2
End Sub

Private Sub AutoCalculate2()
    If (ActiveSheet.Range("A6").Value) = True Then
        Range(Evaluate("INDIRECT(""$C$6"")")).Calculate
        If Range("ам╤у!K2").Value = 1 Then
            Application.OnTime Now + TimeValue("00:01:00"), "AutoCalculate2"
        End If
    End If
End Sub


Sub toText_Click()
    For Each cell In Selection
        cell.Copy
        cell.PasteSpecial xlPasteValues
    Next cell
    
End Sub


Sub ScrollToNow()
    Call CalculateRange1
    currCol = Selection.Cells(1).Column
    'Range(Evaluate("INDIRECT(""$S$2"")")).Select
    ActiveWindow.ScrollRow = Range(Evaluate("INDIRECT(""$S$2"")")).Row
    'ActiveWindow.ScrollColumn = currCol
End Sub
Sub ScrollToNow2()
    currCol = Selection.Cells(1).Column
    Range(Evaluate("INDIRECT(""$F$2"")")).Select
    ActiveWindow.ScrollRow = Range(Evaluate("INDIRECT(""$F$2"")")).Row
    'ActiveWindow.ScrollColumn = currCol
End Sub
Sub ScrollToEnd()
    currCol = Selection.Cells(1).Column
    Range(Evaluate("INDIRECT(""$T$2"")")).Select
    ActiveWindow.ScrollRow = Range(Evaluate("INDIRECT(""$T$2"")")).Row
    'ActiveWindow.ScrollColumn = currCol
End Sub
Sub ScrollToRow()
Attribute ScrollToRow.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim finStr As String
    currCol = Selection.Cells(1).Column
    Dim rownum As String
    
    rownum = InputBox("GOTO", "Please Enter Value")
    If rownum <> vbNullString Then
        finStr = evals("=" + Replace(Range("$U$2").Value2, "rownum", rownum))
        'Range(finStr).Select
        ActiveWindow.ScrollRow = Range(finStr).Row
        ActiveWindow.ScrollColumn = currCol
        Cells(Range(finStr).Row, currCol).Select
    Else
     Call ScrollToNow
    End If
End Sub

