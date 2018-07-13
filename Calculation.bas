Attribute VB_Name = "Calculation"


Sub CalculateRange1()
    Range(Evaluate("INDIRECT(""$C$1"")")).Calculate
    

    
End Sub
Sub CalculateRange2()
    
    Call CalculateRange1
    Range(Evaluate("INDIRECT(""$C$2"")")).Calculate
'    Range(Evaluate("INDIRECT(""$R$1"")")).Calculate
'    Range(Evaluate("INDIRECT(""$R$2"")")).Calculate
    Call CalculateRange1
End Sub
Sub CalculateRange3()
    Range(Evaluate("INDIRECT(""$B$5"")")).Calculate
End Sub
Sub CalculateRange4()
    Call CalculateRange3
    Range(Evaluate("INDIRECT(""$B$6"")")).Calculate
    Call CalculateRange3
End Sub

Sub CalculateRange5()
    Range(Evaluate("INDIRECT(""$AB$1"")")).Calculate
End Sub

Sub CalculateRange7()
    Call generateID
    Call updateID
    Call CalculateRange1
    Range(Range("���!L2").Value2).Calculate
    Call SyncTimeline
    
    Dim mTable As Range
    Set mTable = Range("���2")
    Call llCalculate(getSubRange(2, 1, _
                     mTable.Rows.count, mTable.Columns.count, _
                    mTable))
'    Range(Evaluate("INDIRECT(""$R$1"")")).Calculate
'    Range(Evaluate("INDIRECT(""$R$2"")")).Calculate
    Call CalculateRange1
End Sub
Sub CalculateRange8()
    Dim regex As Object
'    Dim r As Range, rC As Range
'
'    ' cells in column A
'    Set r = Range("A2", Cells(Rows.count, "A").End(xlUp))
'
'    Set regex = CreateObject("VBScript.RegExp")
'    regex.Pattern = " \<.*?\>"
'
'    ' loop through the cells in column A and execute regex replace
'    For Each rC In r
'        If rC.Value <> "" Then rC.Value = regex.Replace(rC.Value, "$1$2-01-$3")
'    Next rC


'
'    Dim myCell As Range
'    Set regex = CreateObject("VBScript.RegExp")
'    regex.Pattern = " \<.*?\>"
'    For Each myCell In Range(Evaluate("INDIRECT(""$W$1"")")).Cells
'        myCell.Value = regex.Replace(myCell.Value, "")
'    Next
    Range(Evaluate("INDIRECT(""$W$1"")")).Calculate
End Sub
Sub SyncTimeline()
    Call CalculateRange9
    Call CalculateRange10
    Call TransferToCores(Range("���2[[�s��]:[Start Date]]"))
    Call TransferToCores(Range("���68[[�s��]:[�������]]"))
    Call TransferToCores(Range("���6866"))
End Sub
Sub CalculateRange9()
    Call updateID
    Range(Range("���!AS2").Value2).Calculate
    Dim ws As Worksheet
    Set ws = Sheets("�s���v�ץ���")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("���6866")
    Set sortcolumn = Range("���6866[�s��]")
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With
End Sub
Sub CalculateRange10()
    Range(Range("���!AT2").Value2).Calculate
    Dim ws As Worksheet
    Set ws = Sheets("�s���v�ɶ���")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("���68")
    Set sortcolumn = Range("���68[�s��]")
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With
End Sub

Sub CalculateRange11()
    Call SyncTimeline
    Call CalculateRange2
    Call SyncTimeline
End Sub

Sub CalculateRange12()
    Range("Table56").Calculate
    Call llCalculate(Range("Table56"))
End Sub
Sub CalculateRange13()
    Range(Range("$C$2").Value2).Calculate
    Range(Range("$C$3").Value2).Calculate
    Call SyncTimeline
    Call llCalculate(Range(Range("$C$2").Value2))
    'Call SyncTimeline
End Sub
Sub CalculateRange14()
    Range(Range("$C$2").Value2).Calculate
    Range(Range("$C$3").Value2).Calculate
    Call SyncTimeline
    Call llCalculate(Range(Range("$C$3").Value2))
    'Call SyncTimeline
End Sub
Sub CalculateTable2ByOrder()
    Application.ScreenUpdating = False
    Dim mselection As Range
    Set mselection = Selection
    Dim toCal As Variant
    toCal = Array("a", "b")
    toCal = Array("���2[[#This Row], [#This Row], [�s��]:[�������]]", _
        "���2[[#This Row], [�i��]:[�i��]]", _
        "���2[[#This Row], [�M�ײֿnSU-MIN]:[���M�ײֿnSU-MIN]]", _
        "���2[[#This Row], [���ݱM��]:[�ɰ�]]", _
        "���2[[#This Row], [SU]:[����Ӯ�]]", _
        "���2[[#This Row], [Location]:[Location]]", _
        "���2[[#This Row], [�_�l�ʤ���]]", _
        "���2[[#This Row], [�w�p�Ӯ�]]", _
        "���2[[#This Row], [�w�p�ʤ���]]", _
        "���2[[#This Row], [��ڦʤ���]:[��گӮ�]]", _
        "���2[[#This Row], [Start Date]:[End Date]]", _
        "���2[[#This Row], [Start Time]:[End Time]]", _
        "���2[[#This Row], [Buffer]:[����]]", _
        "���2[[#This Row], [Dependency]:[note]]", _
        "���2[[#This Row], [�Ѿl�ɶ�]:[�{�b�w�p�i��]]", _
        "���2[[#This Row], [�ܧ����٦�]:[�w�Ӯ�]]", _
        "���2[[#This Row], [�w�`��]:[Subject]]", _
        "���2[[#This Row], [Certainty]]", _
        "���2[[#This Row], [Latitude]:[Longitude]]", _
        "���2[[#This Row], [Location Verify]:[Dependency Verify]]")
    
    Dim r As Range
    For Each cell In mselection
        
        cell.Select
        For Each toCals In toCal
            'Debug.Print toCals
            Set r = Evaluate(toCals)
            If Not (r Is Nothing) Then
            
                On Error Resume Next
                r.Calculate
                On Error Resume Next
                Call llCalculate(r)
            End If
            'DoEvents
        Next toCals
    Next cell
    
    For Each cell In mselection
        cell.Select
        For Each toCals In toCal
            Set r = Evaluate(toCals)
            If Not (r Is Nothing) Then
                r.Calculate
            End If
            'DoEvents
        Next toCals
    Next cell
    Application.ScreenUpdating = True
End Sub

Sub FilterSubject()
    CurrField = Selection.Cells(1).Column - Range(ActiveCell.ListObject.Name).Column + 1
    Dim subject As String
    subject = Selection.Cells(1).Value 'Range(use_Structured(Selection.Cells(1), 6)).Value2
    'Field = Range(use_Structured(Selection.Cells(1), 6)).Column - Range("���2").Column + 1
    ActiveSheet.ListObjects(ActiveCell.ListObject.Name).Range.AutoFilter Field:=CurrField, Criteria1:=subject
End Sub

Sub ClearFilterSubject()
    For Each cell In Selection
        CurrField = cell.Column - Range(ActiveCell.ListObject.Name).Column + 1
        'Field = Range(use_Structured(Selection.Cells(1), 6)).Column - Range("���2").Column + 1
        ActiveSheet.ListObjects(ActiveCell.ListObject.Name).Range.AutoFilter Field:=CurrField
    Next cell
End Sub


Sub FillResourceByTask()
    Dim TaskNames As Range
    Set TaskNames = Application.InputBox(Prompt:="TaskNames", Title:="TaskNames", Type:=8)
    Dim TargetIDs As Range
    Set TargetIDs = Application.InputBox(Prompt:="TargetIDs", Title:="TargetIDs", Type:=8)
    Dim targetWorksheet As Worksheet
    Set targetWorksheet = Worksheets("�s���v�ץ���")
    
    For i = 1 To TaskNames.Cells.count
    
        Dim TitleCell As Range
        Set TitleCell = Range("�s���v�ץ���!3:3").Find(Replace(TaskNames.Cells(i).Value2, "t.", "r."), LookIn:=xlValues)
        Dim RowCell As Range
        Set RowCell = Range("���6866[ID]").Find(TargetIDs.Cells(i).Value2, LookIn:=xlValues)

        Dim target As Range
        Set target = targetWorksheet.Cells(RowCell.Row, TitleCell.Column)
        target.Value2 = target.Value2 - 1
    Next i
    
End Sub
