Attribute VB_Name = "View"
Sub AutoUpdate2_Click()
    Call AutoCalculate2
End Sub

Private Sub AutoCalculate2()
    If (ActiveSheet.range("A6").Value) = True Then
        range(Evaluate("INDIRECT(""$C$6"")")).Calculate
        If range("�Ͷ�!K2").Value = 1 Then
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
    targetID = range("�Ͷ�!B2")
    Dim targetRange As range
    On Error Resume Next
    If ActiveSheet.name = "���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���2[ID]"), Application.Match(targetID, range("���2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "�s���v�ץ���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���6866[ID]"), Application.Match(targetID, range("���6866[ID]"), 0))
    ElseIf ActiveSheet.name = "�s���v�W���" Then
        task = use_Structured2R(Application.WorksheetFunction.index(range("���2[ID]"), Application.Match(targetID, range("���2[ID]"), 0)), "���2", "�������")
        Set targetRange = Application.WorksheetFunction.index(range("���62[�u�@����]"), Application.Match(task, range("���62[�u�@����]"), 0))
    ElseIf ActiveSheet.name = "���Ȫ�" Then
        task = use_Structured2R(Application.WorksheetFunction.index(range("���2[ID]"), Application.Match(targetID, range("���2[ID]"), 0)), "���2", "�������")
        Set targetRange = Application.WorksheetFunction.index(range("���55[�u�@����]"), Application.Match(task, range("���55[�u�@����]"), 0))
    End If
    
    If targetRange.Row <> Selection.Row Then
        Cells(targetRange.Row, Selection.Column).Select
         ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
Sub ScrollToNow()
    range("���!S2").Calculate
    targetID = range("���!S2")
    Dim targetRange As range
    If ActiveSheet.name = "���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���2[ID]"), Application.Match(targetID, range("���2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "�s���v�ץ���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���6866[ID]"), Application.Match(targetID, range("���6866[ID]"), 0))
    End If
    If targetRange.Row <> Selection.Row Then
        Cells(targetRange.Row, Selection.Column).Select
        ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
Sub ScrollToEnd()
    range("���!T2").Calculate
    targetID = range("���!T2")
    Dim targetRange As range
    If ActiveSheet.name = "���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���2[ID]"), Application.Match(targetID, range("���2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "�s���v�ץ���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���6866[ID]"), Application.Match(targetID, range("���6866[ID]"), 0))
    End If
    If targetRange.Row <> Selection.Row Then
         Cells(targetRange.Row, Selection.Column).Select
         ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
Sub ScrollToID(targetID As Variant)
    Dim targetRange As range
    If ActiveSheet.name = "���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���2[ID]"), Application.Match(targetID, range("���2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "�s���v�ץ���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���6866[ID]"), Application.Match(targetID, range("���6866[ID]"), 0))
    End If
    If targetRange.Row <> Selection.Row Then
         Cells(targetRange.Row, Selection.Column).Select
         ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
Sub ScrollToNow2()
    currCol = Selection.Cells(1).Column
    range("���!F2").Calculate
    ActiveWindow.ScrollRow = range(range("���!F2").Value2).Row
    'ActiveWindow.ScrollColumn = currCol
End Sub
Sub ScrollToNow3()
    ActiveWindow.ScrollRow = getCurretnActualR().Row
End Sub
Sub ScrollToTop()
    currCol = Selection.Cells(1).Column
    range("���!T2").Calculate
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
        ActiveWindow.ScrollRow = Application.WorksheetFunction.index(range("���2[�s��]"), Application.Match(CInt(rownum), range("���2[�s��]"), 0)).Row
        ActiveWindow.ScrollColumn = currCol
    Else
     Call ScrollToNow
    End If
End Sub
Sub ScrollToInputID()
Attribute ScrollToInputID.VB_ProcData.VB_Invoke_Func = "i\n14"
    targetID = CInt(InputBox("GOTO", "Please Enter ID"))
    Dim targetRange As range
    If ActiveSheet.name = "���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���2[ID]"), Application.Match(targetID, range("���2[ID]"), 0))
    ElseIf ActiveSheet.name = "ResourceTimeline" Then
        Set targetRange = Application.WorksheetFunction.index(range("D5#"), Application.Match(targetID, range("D5#"), 0))
    ElseIf ActiveSheet.name = "�s���v�ץ���" Then
        Set targetRange = Application.WorksheetFunction.index(range("���6866[ID]"), Application.Match(targetID, range("���6866[ID]"), 0))
    End If
    If targetRange.Row <> Selection.Row Then
         Cells(targetRange.Row, Selection.Column).Select
         ActiveWindow.ScrollRow = targetRange.Row
    End If
End Sub
