Attribute VB_Name = "View"
Sub AutoUpdate2_Click()
    Call AutoCalculate2
End Sub

Private Sub AutoCalculate2()
    If (ActiveSheet.range("A6").Value) = True Then
        range(Evaluate("INDIRECT(""$C$6"")")).Calculate
        If range("趨勢!K2").Value = 1 Then
            Application.OnTime Now + TimeValue("00:01:00"), "AutoCalculate2"
        End If
    End If
End Sub


Sub toText_Click()
'    For Each cell In Selection
'        cell.Copy
'        cell.PasteSpecial xlPasteValues
'    Next cell
    Selection.Value2 = Selection.Value2
End Sub

Sub MoveToCurrentRow()
    targetID = range("趨勢!B2")
    Dim targetRange As range
    On Error Resume Next
    If ActiveSheet.name = "交易" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格2[ID]"), Application.Match(targetID, range("表格2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "存取權修正表" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格6866[ID]"), Application.Match(targetID, range("表格6866[ID]"), 0))
    ElseIf ActiveSheet.name = "存取權增減表" Then
        task = use_Structured2R(Application.WorksheetFunction.index(range("表格2[ID]"), Application.Match(targetID, range("表格2[ID]"), 0)), "表格2", "交易物件")
        Set targetRange = Application.WorksheetFunction.index(range("表格62[工作物件]"), Application.Match(task, range("表格62[工作物件]"), 0))
    ElseIf ActiveSheet.name = "價值表" Then
        task = use_Structured2R(Application.WorksheetFunction.index(range("表格2[ID]"), Application.Match(targetID, range("表格2[ID]"), 0)), "表格2", "交易物件")
        Set targetRange = Application.WorksheetFunction.index(range("表格55[工作物件]"), Application.Match(task, range("表格55[工作物件]"), 0))
    End If
    
    If targetRange.Row <> Selection.Row Then
        Cells(targetRange.Row, Selection.Column).Select
         ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
Sub ScrollToNow()
    range("交易!S2").Calculate
    targetID = range("交易!S2")
    Dim targetRange As range
    If ActiveSheet.name = "交易" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格2[ID]"), Application.Match(targetID, range("表格2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "存取權修正表" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格6866[ID]"), Application.Match(targetID, range("表格6866[ID]"), 0))
    End If
    If targetRange.Row <> Selection.Row Then
        Cells(targetRange.Row, Selection.Column).Select
        ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
Sub ScrollToEnd()
    range("交易!T2").Calculate
    targetID = range("交易!T2")
    Dim targetRange As range
    If ActiveSheet.name = "交易" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格2[ID]"), Application.Match(targetID, range("表格2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "存取權修正表" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格6866[ID]"), Application.Match(targetID, range("表格6866[ID]"), 0))
    End If
    If targetRange.Row <> Selection.Row Then
         Cells(targetRange.Row, Selection.Column).Select
         ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
Sub ScrollToID(targetID As Variant)
    Dim targetRange As range
    If ActiveSheet.name = "交易" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格2[ID]"), Application.Match(targetID, range("表格2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "存取權修正表" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格6866[ID]"), Application.Match(targetID, range("表格6866[ID]"), 0))
    End If
    If targetRange.Row <> Selection.Row Then
         Cells(targetRange.Row, Selection.Column).Select
         ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
Sub ScrollToNow2()
    currCol = Selection.Cells(1).Column
    range("交易!F2").Calculate
    ActiveWindow.ScrollRow = range(range("交易!F2").Value2).Row
    'ActiveWindow.ScrollColumn = currCol
End Sub
Sub ScrollToNow3()
    ActiveWindow.ScrollRow = getCurretnActualR().Row
End Sub
Sub ScrollToTop()
    currCol = Selection.Cells(1).Column
    range("交易!T2").Calculate
    ActiveWindow.ScrollRow = 1
    'ActiveWindow.ScrollColumn = currCol
End Sub
Sub ScrollToTop2()
    currCol = Selection.Cells(1).Column
    ActiveWindow.ScrollRow = 1
End Sub
Sub ScrollToRow()
Attribute ScrollToRow.VB_ProcData.VB_Invoke_Func = "r\n14"
    Dim finStr As String
    currCol = Selection.Cells(1).Column
    Dim rownum As String
    rownum = InputBox("GOTO", "Please Enter Index")
    If rownum <> vbNullString Then
        ActiveWindow.ScrollRow = Application.WorksheetFunction.index(range("表格2[編號]"), Application.Match(CInt(rownum), range("表格2[編號]"), 0)).Row
        ActiveWindow.ScrollColumn = currCol
    Else
     Call ScrollToNow
    End If
End Sub
Sub ScrollToInputID()
Attribute ScrollToInputID.VB_ProcData.VB_Invoke_Func = "i\n14"
    targetID = CInt(InputBox("GOTO", "Please Enter ID"))
    Dim targetRange As range
    If ActiveSheet.name = "交易" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格2[ID]"), Application.Match(targetID, range("表格2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "存取權修正表" Then
        Set targetRange = Application.WorksheetFunction.index(range("表格6866[ID]"), Application.Match(targetID, range("表格6866[ID]"), 0))
    End If
    If targetRange.Row <> Selection.Row Then
         Cells(targetRange.Row, Selection.Column).Select
         ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
