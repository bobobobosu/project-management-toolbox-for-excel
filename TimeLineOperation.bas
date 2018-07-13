Attribute VB_Name = "TimeLineOperation"
Sub AlignNow_Click()
            Dim DATECell As String
            DATECell = ""
            Dim PrevCell As String
            PrevCell = ""
            Dim NextCell As String
            NextCell = ""
            
            Dim sumVal As Double
            sumVal = 0
            
            Dim DATEVal As Double
            DATEVal = 0
            Dim PrevVal As Double
            PrevVal = 0
            Dim NextVal As Double
            NextVal = 0
            
'            For Each cell In Selection
'                If DATECell = "" Then
'                    DATECell = cell.Address
'
'                ElseIf PrevtCell = "" Then
'                    PrevtCell = cell.Address
'                    'PrevVal = cell.Value2
'                ElseIf NextCell = "" Then
'                    NextCell = cell.Address
'                    'NextVal = cell.Value2
'                End If
'            Next cell
'
            Dim DATECellR As Range
            Set DATECellR = Application.InputBox(Prompt:="Cell to Align", Title:="Select...", Type:=8)
            Dim PrevCellR As Range
            Set PrevCellR = Application.InputBox(Prompt:="Cell Before", Title:="Select...", Type:=8)
            Dim NextCellR As Range
            Set NextCellR = Application.InputBox(Prompt:="Cell After", Title:="Select...", Type:=8)
            
            DATEVal = Range(use_Structured(DATECellR, 4)).Value2
            PrevVal = Range(use_Structured(PrevCellR, 2)).Value2
            NextVal = Range(use_Structured(NextCellR, 2)).Value2
            
'            DATEVal = Range(use_Structured(Range(DATECell), 4)).Value2
'            PrevVal = Range(use_Structured(Range(PrevtCell), 2)).Value2
'            NextVal = Range(use_Structured(Range(NextCell), 2)).Value2
            
            Dim offset As Double
            offset = Now - DATEVal
            
'            Range(PrevtCell).Value = PrevVal + OFFSET
'            Range(NextCell).Value = NextVal - OFFSET
'            Range(DATECell).Value = Now
        
            Range(use_Structured(PrevCellR, 2)).Value = PrevVal + offset
            Range(use_Structured(NextCellR, 2)).Value = NextVal - offset
            Range(use_Structured(DATECellR, 4)).Value = Now
            Call CalculateRange2
End Sub
Sub Aligned_Click()
            Dim DATECell As String
            DATECell = ""
            Dim toDATECell As String
            toDATECell = ""
            Dim PrevCell As String
            PrevCell = ""
            Dim NextCell As String
            NextCell = ""
            
            Dim sumVal As Double
            sumVal = 0
            
            Dim DATEVal As Double
            DATEVal = 0
            Dim toDATEVal As Double
            toDATEVal = 0
            Dim PrevVal As Double
            PrevVal = 0
            Dim NextVal As Double
            NextVal = 0
            
            For Each cell In Selection
                If DATECell = "" Then
                    DATECell = cell.Address
                    'DATEVal = cell.Value2
                ElseIf toDATECell = "" Then
                    toDATECell = cell.Address
                    'toDATEVal = cell.Value2
                ElseIf PrevCell = "" Then
                    PrevCell = cell.Address
                    'PrevVal = cell.Value2
                ElseIf NextCell = "" Then
                    NextCell = cell.Address
                    'NextVal = cell.Value2
                End If
            Next cell
            
            
            
            Dim DATECellR As Range
            Set DATECellR = Application.InputBox(Prompt:="Cell to Align", Title:="Select...", Type:=8)
            Dim PrevCellR As Range
            Set PrevCellR = Application.InputBox(Prompt:="Cell Before", Title:="Select...", Type:=8)
            Dim NextCellR As Range
            Set NextCellR = Application.InputBox(Prompt:="Cell After", Title:="Select...", Type:=8)

            Dim FirstCell3 As Variant
            FirstCell3 = InputBox("Time Value", "Please Enter Time Value", Format(Now(), "m/d/yy h:mm:ss;@"))
            
            If (FirstCell3 <> vbNullString) Then
            
                DATEVal = Range(use_Structured(DATECellR, 4)).Value2
                PrevVal = Range(use_Structured(PrevCellR, 2)).Value2
                NextVal = Range(use_Structured(NextCellR, 2)).Value2
                toDATEVal = dateValue(FirstCell3) + TimeValue(FirstCell3)
                
                  
                
    '            DATEVal = Range(use_Structured(Range(DATECell), 4)).Value2
    '            PrevVal = Range(use_Structured(Range(PrevCell), 2)).Value2
    '            NextVal = Range(use_Structured(Range(NextCell), 2)).Value2
    '            toDATEVal = Range(use_Structured(Range(toDATECell), 4)).Value2
                
                
                Dim offset As Double
                offset = toDATEVal - DATEVal
                
                Range(use_Structured(PrevCellR, 2)).Value = PrevVal + offset
                Range(use_Structured(NextCellR, 2)).Value = NextVal - offset
                Range(use_Structured(DATECellR, 4)).Value = toDATEVal
            
            
            End If

            

End Sub

Sub TaskComplete()
Dim r As Range

Set r = Evaluate("表格2[[#This Row], [Description]:[Description]]")
r.Value2 = r.Value2

'r.Select
'Call toText_Click

'Set r = Evaluate("表格2[[#This Row], [SU-MIN]:[SU-MIN]]")
'r.Value2 = r.Value2
''r.Select
''Call toText_Click
'
'Set r = Evaluate("表格2[[#This Row], [完整耗時]:[完整耗時]]")
'r.Value2 = r.Value2
'''r.Select
'''Call toText_Click

Set r = Evaluate("表格2[[#This Row], [Location]:[Location]]")
r.Value2 = r.Value2
'r.Select
'Call toText_Click


Set r = Evaluate("表格2[[#This Row], [Latitude]:[Latitude]]")
r.Value2 = r.Value2
'r.Select
'Call toText_Click

Set r = Evaluate("表格2[[#This Row], [Longitude]:[Longitude]]")
r.Value2 = r.Value2
'r.Select
'Call toText_Click

Set r = Evaluate("表格2[[#This Row], [實際百分比]:[實際百分比]]")
r.Value2 = r.Value2
'r.Select
'Call toText_Click

Set r = Evaluate("表格2[[#This Row], [起始百分比]:[起始百分比]]")
r.Value2 = r.Value2
'r.Select
'Call toText_Click


Set r = Evaluate("表格2[[#This Row], [SU]:[SU]]")
r.Value2 = r.Value2
'r.Select
'Call toText_Click
Set r = Evaluate("表格2[[#This Row], [時區]:[時區]]")
r.Value2 = r.Value2


End Sub

Sub Completed()
    For Each cell In Selection
        cell.Select
        Call TaskComplete
        Set r = Evaluate("表格2[[#This Row], [實際耗時]:[實際耗時]]")
        r.Value2 = r.Value2
    Next cell
    
    'Call CalculateRange2
End Sub


Sub CompleteNow()
    Call TaskComplete
    Set r = Evaluate("表格2[[#This Row], [實際耗時]:[實際耗時]]")
    r.Value = Now - Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]")
    
    'Call CalculateRange2
End Sub

Sub FitDuration()

'    Dim fromC As String
'    Dim toC As String
'
'    For Each cell In Selection
'        If fromC = "" Then
'            fromC = cell.Address
'        Else
'            toC = cell.Address
'        End If
'    Next cell
'
'    Dim StartDate As Double
'    Dim Duration As Double
'
'    ActiveSheet.Range(toC).Select
'    Set q = Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]")
'    q.Select
'    StartDate = q.Value
'
'    ActiveSheet.Range(fromC).Select
'    Set q = Evaluate("表格2[[#This Row], [預計耗時]:[預計耗時]]")
'    q.Select
'    Duration = q.Value
'
'    Set q = Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]")
'    q.Select
'    q.Value = StartDate - Duration
'
'    Set q = Evaluate("表格2[[#This Row], [預計耗時]:[預計耗時]]")
'    q.Select
'    q.Value = "=[@完整耗時]"
'
'    Call CalculateRange2
    
'    Dim fromC As String
'    Dim toC As String
'    Dim targetC As String
'
'    For Each cell In Selection
'        If fromC = "" Then
'            fromC = cell.Address
'        ElseIf targetC = "" Then
'            targetC = cell.Address
'        Else
'            toC = cell.Address
'        End If
'    Next cell
    For Each cell In Selection
        cell.Select
        Set r = Evaluate("表格2[[#This Row],[Start Date]:[Start Date]]")
        r.Value = "=INDEX([End Date],MATCH([@編號]-1,[編號],0))"
        Set r = Evaluate("表格2[[#This Row], [預計耗時]:[預計耗時]]")
        r.Value = "=INDEX([Start Date],ROW()-ROW(表格2)+2)-INDEX([Start Date],ROW()-ROW(表格2)+1)"
        Set r = Evaluate("表格2[[#This Row], [實際耗時]:[實際耗時]]")
        r.Value = "=[@預計耗時]"
    
        Evaluate("表格2[@]").Calculate
        
        Set r = Evaluate("表格2[[#This Row], [預計耗時]:[預計耗時]]")
        r.Value2 = r.Value2
    
    Next cell
End Sub

Sub Connect2Next()
'If Selection.Cells.count = 2 Then
    Dim mysel As Range
    Set mysel = Selection
    Dim fromC As String
    Dim toC As String
    Dim TimeBetween As Double
    Dim lastRow As Long
    
    For Each cell In mysel
        lastRow = Range(use_Structured(cell, 0)).Value2
    Next cell
    For Each cell In mysel
        cell.Select
        TimeBetween = TimeBetween + Evaluate("表格2[[#This Row], [實際耗時]:[實際耗時]]")
        Set r = Evaluate("表格2[[#This Row], [編號]:[編號]]")
        r.Value2 = lastRow
        If fromC = "" Then
            fromC = cell.Address
        Else
            toC = cell.Address
        End If
    Next cell
    
    Dim nextStart As Double
    Dim Duration As Double

    Range(toC).Select
    Set r = Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]")
    nextStart = r.Value
    Set r = Evaluate("表格2[[#This Row], [實際耗時]:[實際耗時]]")
    Duration = r.Value


    Range(fromC).Select
    
    Set q = Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]")
    q.Value = nextStart - (TimeBetween - Duration)

    Call SortRowNum
    'Call CalculateRange2
'End If
End Sub


Sub FillInCells()
    Dim toMove_time As Range
    Set toMove_time = Application.InputBox(Prompt:="toMove_time", Title:="toMove_time", Type:=8)
    Dim availible_time As Range
    Set availible_time = Application.InputBox(Prompt:="availible_time", Title:="availible_time", Type:=8)
    Call ScrollToEnd
    Dim toPut_loc As Range
    Set toPut_loc = Application.InputBox(Prompt:="toPut_loc", Title:="toPut_loc", Type:=8)
    
    
    Dim toMove_time_c As New Collection
    Dim toMove_job_c As New Collection
    Dim toMove_des_c As New Collection
    Dim toMove_startPercent_c As New Collection
    Dim toMove_endPercent_c As New Collection
    'toMove_time_arr = toMove_time.Value
    For i = toMove_time.Cells.count To 1 Step -1
        toMove_time_c.Add (Range(use_Structured(getItemByIndexInRange(toMove_time, i), 2)).Value2)
        toMove_job_c.Add (Range(use_Structured(getItemByIndexInRange(toMove_time, i), 6)).Value2)
        toMove_des_c.Add (Range(use_Structured(getItemByIndexInRange(toMove_time, i), 9)).Value2)
        toMove_startPercent_c.Add (Range(use_Structured(getItemByIndexInRange(toMove_time, i), 7)).Value2)
        toMove_endPercent_c.Add (Range(use_Structured(getItemByIndexInRange(toMove_time, i), 8)).Value2)
        Debug.Print toMove_time_c.Item(toMove_time_c.count)
    Next i
    

    Dim availible_time_c As New Collection
    Dim availible_start_c As New Collection
    'availible_time_arr = availible_time.Value
    For i = availible_time.Cells.count To 1 Step -1
        availible_time_c.Add (Range(use_Structured(getItemByIndexInRange(availible_time, i), 2)).Value2)
        availible_start_c.Add (Range(use_Structured(getItemByIndexInRange(availible_time, i), 4)).Value2)
        Range(use_Structured(getItemByIndexInRange(availible_time, i), 2)).Value2 = 0
    Next i
    


    Dim toPut_time_c As New Collection
    Dim toPut_start_c As New Collection
    Dim toPut_job_c As New Collection
    Dim toPut_des_c As New Collection
    Dim toPut_startPercent_c As New Collection
    Dim toPut_endPercent_c As New Collection
    
    Debug.Print "ffff"
'    Dim toMove_time_c As New Collection
'    Dim toMove_job_c As New Collection
'    Dim availible_time_c As New Collection
'    Dim availible_start_c As New Collection
'    Dim toPut_time_c As New Collection
'    Dim toPut_start_c As New Collection
'    Dim toPut_job_c As New Collection
    
    
    
    Dim adjustedstartPercent
    If toMove_startPercent_c.count > 0 Then adjustedstartPercent = lastCol(toMove_startPercent_c)
    Do While (toMove_job_c.count > 0) And (availible_time_c.count > 0)
        If lastCol(toMove_time_c) < lastCol(availible_time_c) Then
            toPut_job_c.Add (lastCol(toMove_job_c))
            toPut_des_c.Add (lastCol(toMove_des_c))
            toPut_time_c.Add (lastCol(toMove_time_c))
            toPut_start_c.Add (lastCol(availible_start_c))
            toPut_startPercent_c.Add (adjustedstartPercent)
            toPut_endPercent_c.Add (lastCol(toMove_endPercent_c))
'            Call UpdateCol(availible_time_c, availible_time_c(1) - toMove_time_c(1), 1)
'            Call UpdateCol(availible_start_c, availible_start_c(1) + toMove_time_c(1), 1)
            remainingAvailible = lastCol(availible_time_c) - lastCol(toMove_time_c)
            availible_time_c.Remove (availible_time_c.count)
            availible_time_c.Add (remainingAvailible)
            adjustedStart = lastCol(availible_start_c) + lastCol(toMove_time_c)
            availible_start_c.Remove (availible_start_c.count)
            availible_start_c.Add (adjustedStart)
'            Set availible_time_c.Item(1) = availible_time_c(1) - toMove_time_c(1)
'            Set availible_start_c.Item(1) = availible_start_c(1) + toMove_time_c(1)
            toMove_time_c.Remove (toMove_time_c.count)
            toMove_job_c.Remove (toMove_job_c.count)
            toMove_des_c.Remove (toMove_des_c.count)
            toMove_startPercent_c.Remove (toMove_startPercent_c.count)
            toMove_endPercent_c.Remove (toMove_endPercent_c.count)
            If toMove_startPercent_c.count > 0 Then adjustedstartPercent = lastCol(toMove_startPercent_c)
        ElseIf lastCol(toMove_time_c) = lastCol(availible_time_c) Then
            toPut_job_c.Add (lastCol(toMove_job_c))
            toPut_des_c.Add (lastCol(toMove_des_c))
            toPut_start_c.Add (lastCol(availible_start_c))
            toPut_time_c.Add (lastCol(toMove_time_c))
            toPut_startPercent_c.Add (adjustedstartPercent)
            toPut_endPercent_c.Add (lastCol(toMove_endPercent_c))
            availible_start_c.Remove (availible_start_c.count)
            availible_time_c.Remove (availible_time_c.count)
            toMove_time_c.Remove (toMove_time_c.count)
            toMove_job_c.Remove (toMove_job_c.count)
            toMove_des_c.Remove (toMove_des_c.count)
            toMove_startPercent_c.Remove (toMove_startPercent_c.count)
            toMove_endPercent_c.Remove (toMove_endPercent_c.count)
            If toMove_startPercent_c.count > 0 Then adjustedstartPercent = lastCol(toMove_startPercent_c)
        Else
            toPut_job_c.Add (lastCol(toMove_job_c))
            toPut_des_c.Add (lastCol(toMove_des_c))
            toPut_start_c.Add (lastCol(availible_start_c))
            toPut_time_c.Add (lastCol(availible_time_c))
            toPut_startPercent_c.Add (adjustedstartPercent)
            toPut_endPercent_c.Add (adjustedstartPercent + (lastCol(toMove_endPercent_c) - lastCol(toMove_startPercent_c)) * (lastCol(availible_time_c) / lastCol(toMove_time_c)))
            'Debug.Print "xx"
            'Debug.Print (lastCol(availible_time_c) / lastCol(toMove_time_c))
            adjustedstartPercent = lastCol(toPut_endPercent_c)
'           Call UpdateCol(toMove_time_c, 1, toMove_time_c(1) - availible_time_c(1))
            'Set toMove_time_c.Item(1) = toMove_time_c(1) - availible_time_c(1)
            adjustedtoMovetime = lastCol(toMove_time_c) - lastCol(availible_time_c)
            toMove_time_c.Remove (toMove_time_c.count)
            toMove_time_c.Add (adjustedtoMovetime)
            
            availible_start_c.Remove (availible_start_c.count)
            availible_time_c.Remove (availible_time_c.count)
        End If
        Debug.Print "xx"
        Debug.Print lastCol(toPut_startPercent_c)
        'Debug.Print "xxxxxx"
        'Debug.Print (CStr(toPut_job_c(toPut_job_c.count)) & " " & CStr(toPut_start_c(toPut_start_c.count)) & " " & CStr(toPut_time_c(toPut_time_c.count)))
        If Not ((toMove_job_c.count > 0) And (availible_time_c.count > 0)) Then
            Exit Do
        End If
    Loop
    If availible_time_c.count > 0 Then
        For i = availible_time_c.count To 1 Step -1
            toPut_time_c.Add (availible_time_c(i))
            toPut_start_c.Add (availible_start_c(i))
            toPut_job_c.Add ("時序專案(288)")
            toPut_des_c.Add ("")
        Next i
    End If
    For i = toPut_job_c.count To 1 Step -1
        Range(use_Structured(toPut_loc.offset(i - 1), 2)).Value2 = toPut_time_c(i)
        Range(use_Structured(toPut_loc.offset(i - 1), 4)).Value2 = toPut_start_c(i)
        Range(use_Structured(toPut_loc.offset(i - 1), 6)).Value2 = toPut_job_c(i)
        Range(use_Structured(toPut_loc.offset(i - 1), 9)).Value2 = toPut_des_c(i)
        On Error Resume Next
        Range(use_Structured(toPut_loc.offset(i - 1), 7)).Value2 = toPut_startPercent_c(i)
        On Error Resume Next
        Range(use_Structured(toPut_loc.offset(i - 1), 8)).Value2 = toPut_endPercent_c(i)
    Next i
    'Set availible = Application.InputBox(Prompt:="availible", Title:="availible", Type:=8)
    'Set toPut = Application.InputBox(Prompt:="toPut", Title:="toPut", Type:=8)
        
        
    'select toMove
    
    'select availible
    'select toPut

End Sub
Public Function lastCol(c As Collection) As Variant
    lastCol = c.Item(c.count)
End Function
Sub test()
    Dim mysel As Range
    Set mysel = Selection
    For Each cell In Selection
        cell.Select
        MsgBox cell.Value
    Next cell
    'Range(Evaluate("ADDRESS(ROW(表格2),COLUMN(表格2))")).Select
   ' MsgBox Range(Evaluate("Cell(""address"",表格2[[#This Row], [實際百分比]:[實際百分比]])")).Column
End Sub

Sub pasteExample(icell As Range)
    Dim sel As Range
    Dim selectedC As Range
    For i = 1 To icell.Cells.count
        Set selectedC = icell.Cells(i)
        Set sel = Range(use_Structured(selectedC, 0)) '.Resize(1, Range("表格2").Columns.count)
        Debug.Print AddressEx(Range("表格2").Cells(1).Resize(1, Range("表格2").Columns.count))
        Range("表格2").offset(1, 0).Resize(1, Range("表格2").Columns.count).Copy
        sel.PasteSpecial Paste:=xlPasteFormulas
        sel.Value2 = 10000
    Next i
    
End Sub
Sub pasteExampl2e()
    Call pasteExample(Selection)
End Sub
Sub Consolidate()
    Dim sel As Range
    Set sel = Selection
    Dim selectedC As Range
    Dim selectedC2 As Range
    For i = 1 To sel.Cells.count - 1
        Set selectedC = sel.Cells(i)
        Set selectedC2 = sel.Cells(i + 1)
        If (Range(use_Structured(selectedC, 6)).Value = Range(use_Structured(selectedC2, 6)).Value) And _
            (Range(use_Structured(selectedC, 9)).Value = Range(use_Structured(selectedC2, 9)).Value) And _
            (Range(use_Structured(selectedC, 5)).Value = Range(use_Structured(selectedC2, 4)).Value) Then
            Range(use_Structured(selectedC2, 2)).Value2 = Range(use_Structured(selectedC2, 2)).Value2 + Range(use_Structured(selectedC, 2)).Value2
            Range(use_Structured(selectedC2, 4)).Value2 = Range(use_Structured(selectedC, 4)).Value2
            Range(use_Structured(selectedC2, 7)).Value2 = Range(use_Structured(selectedC, 7)).Value2
            Call pasteExample(selectedC)
            
        End If
        
        
    Next i
    Call SortRowNum
End Sub

Sub test2()
    MsgBox use_Structured(Range("$A$5"), 1)
End Sub

Function use_Structured(cell As Variant, mode As Integer)
    If mode = 0 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [編號]:[編號]]")
        use_Structured = r.Address
    ElseIf mode = 1 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [完整耗時]:[完整耗時]]")
        use_Structured = r.Address
    ElseIf mode = 2 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [預計耗時]:[預計耗時]]")
        use_Structured = r.Address
    ElseIf mode = 3 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [實際耗時]:[實際耗時]]")
        use_Structured = r.Address
    ElseIf mode = 4 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]")
        use_Structured = r.Address
    ElseIf mode = 5 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [End Date]:[End Date]]")
        use_Structured = r.Address
    ElseIf mode = 6 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [交易物件]:[交易物件]]")
        use_Structured = r.Address
    ElseIf mode = 7 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [起始百分比]:[起始百分比]]")
        use_Structured = r.Address
    ElseIf mode = 8 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [預計百分比]:[預計百分比]]")
        use_Structured = r.Address
    ElseIf mode = 9 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [Description]:[Description]]")
        use_Structured = r.Address
    ElseIf mode = 10 Then
        cell.Select
        Set r = Evaluate("表格2[[#This Row], [ID]:[ID]]")
        use_Structured = r.Address
    End If
End Function


Sub SortDate()
Call generateID
Application.CutCopyMode = False
Dim targetID As String
targetID = Evaluate("表格2[[#This Row], [ID]:[ID]]").Value2

Dim ws As Worksheet
Set ws = ActiveSheet
Dim tbl As ListObject
Set tbl = ws.ListObjects("表格2")
Dim sortcolumn As Range
Set sortcolumn = Range("表格2[Start Date]")
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .Header = xlYes
   .Apply
End With

Set r = Evaluate("表格2[編號]")
r.Calculate

Set tbl = ws.ListObjects("表格2")
Set sortcolumn = Range("表格2[編號]")
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .Header = xlYes
   .Apply
End With

Call CalculateRange1

Set r = Evaluate("表格2[編號]")
Evaluate("表格2[編號]").Value2 = "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))+1"
Range("$A$4").Value2 = 1
Set r = Evaluate("表格2[編號]")
r.Calculate



Dim finStr As String
finStr = evals("=" + Replace(Range("$V$2").Value2, "rownum", targetID))
Range(finStr).Select
ActiveWindow.ScrollRow = Selection.Row
ActiveWindow.ScrollColumn = Selection.Column
    
End Sub


Sub SortRowNum()
Call generateID
Application.CutCopyMode = False
Dim targetID As String
targetID = Evaluate("表格2[[#This Row], [ID]:[ID]]").Value2


Dim ws As Worksheet
Set ws = ActiveSheet
Dim tbl As ListObject
Set tbl = ws.ListObjects("表格2")
Set sortcolumn = Range("表格2[編號]")
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .Header = xlYes
   .Apply
End With



Set r = Evaluate("表格2[編號]")
Evaluate("表格2[編號]").Value2 = "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))+1"
Range("$A$4").Value2 = 1
Set r = Evaluate("表格2[編號]")
r.Calculate

Call CalculateRange1

Dim finStr As String
finStr = evals("=" + Replace(Range("$V$2").Value2, "rownum", targetID))
Range(finStr).Select
ActiveWindow.ScrollRow = Selection.Row
ActiveWindow.ScrollColumn = Selection.Column
    
End Sub

Sub SortRowNumBySel()
    Dim mselection As Range

    Set mselection = Selection
    fromCell = use_Structured(mselection.Cells(1), 0)
    toCell = use_Structured(mselection.Cells(2), 0)
    Range(fromCell).Value2 = Range(toCell).Value2
    Call SortRowNum
End Sub

Sub FillPercent()
    totaltime = 0
    lastPercent = 0
    Set mselection = Selection
    For Each cell In mselection
        totaltime = totaltime + Range(use_Structured(cell, 2)).Value2
        lastPercent = Range(use_Structured(cell, 8)).Value2
    Next cell
    
    prevPercent = Range(use_Structured(mselection(1), 7)).Value2
    firstPercent = Range(use_Structured(mselection(1), 7)).Value2
    For Each cell In mselection
        Delta = (lastPercent - firstPercent) * (Range(use_Structured(cell, 2)).Value2 / totaltime)
        Range(use_Structured(cell, 8)).Value2 = prevPercent + Delta
        prevPercent = Range(use_Structured(cell, 8)).Value2
    Next cell
    
End Sub


Sub SwapCells()
    Dim selected As Range
    Set selected = Selection
    If selected.Areas.count > 1 Then
        Dim row1 As Range
        Dim row2 As Range
        Set row1 = selected.Areas(1)
        Set row2 = selected.Areas(2)
        For i = 1 To row1.Cells.count
            tmpval = row1.Cells(i).Formula
            row1.Cells(i).Formula = row2.Cells(i).Formula
            row2.Cells(i).Formula = tmpval
        Next i
        
    
    End If
    
    
End Sub
