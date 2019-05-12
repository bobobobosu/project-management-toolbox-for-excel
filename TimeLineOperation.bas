Attribute VB_Name = "TimeLineOperation"
Sub FillExampleWithArrayOfDict(Optional plan As Variant)
    If IsMissing(plan) Then
        jsonPlan = Selection
    Else
        jsonPlan = plan
    End If
    
    Dim result As Collection
    Set result = JsonConverter.ParseJson(jsonPlan)
    Set FillRow = FirstExample()
    For Each dict In result
        
        Dim key As Variant
        
        For Each key In dict.Keys
            Debug.Print key, dict(key), use_Structured2R(FillRow, "表格2", key).address
            Dim target As range
            use_Structured2R(FillRow, "表格2", key).Value2 = dict(key)
        Next key
        Set FillRow = FillRow.offset(1)
    Next
End Sub

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
            Dim DATECellR As range
            Set DATECellR = Application.InputBox(Prompt:="Cell to Align", title:="Select...", Type:=8)
            Dim PrevCellR As range
            Set PrevCellR = Application.InputBox(Prompt:="Cell Before", title:="Select...", Type:=8)
            Dim NextCellR As range
            Set NextCellR = Application.InputBox(Prompt:="Cell After", title:="Select...", Type:=8)
            
            DATEVal = range(use_Structured(DATECellR, 4)).Value2
            PrevVal = range(use_Structured(PrevCellR, 2)).Value2
            NextVal = range(use_Structured(NextCellR, 2)).Value2
            
'            DATEVal = Range(use_Structured(Range(DATECell), 4)).Value2
'            PrevVal = Range(use_Structured(Range(PrevtCell), 2)).Value2
'            NextVal = Range(use_Structured(Range(NextCell), 2)).Value2
            
            Dim offset As Double
            offset = Now - DATEVal
            
'            Range(PrevtCell).Value = PrevVal + OFFSET
'            Range(NextCell).Value = NextVal - OFFSET
'            Range(DATECell).Value = Now
        
            range(use_Structured(PrevCellR, 2)).Value = PrevVal + offset
            range(use_Structured(NextCellR, 2)).Value = NextVal - offset
            range(use_Structured(DATECellR, 4)).Value = Now
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
                    DATECell = cell.address
                    'DATEVal = cell.Value2
                ElseIf toDATECell = "" Then
                    toDATECell = cell.address
                    'toDATEVal = cell.Value2
                ElseIf PrevCell = "" Then
                    PrevCell = cell.address
                    'PrevVal = cell.Value2
                ElseIf NextCell = "" Then
                    NextCell = cell.address
                    'NextVal = cell.Value2
                End If
            Next cell
            
            
            
            Dim DATECellR As range
            Set DATECellR = Application.InputBox(Prompt:="Cell to Align", title:="Select...", Type:=8)
            Dim PrevCellR As range
            Set PrevCellR = Application.InputBox(Prompt:="Cell Before", title:="Select...", Type:=8)
            Dim NextCellR As range
            Set NextCellR = Application.InputBox(Prompt:="Cell After", title:="Select...", Type:=8)

            Dim FirstCell3 As Variant
            FirstCell3 = InputBox("Time Value", "Please Enter Time Value", Format(Now(), "m/d/yy h:mm:ss;@"))
            
            If (FirstCell3 <> vbNullString) Then
            
                DATEVal = range(use_Structured(DATECellR, 4)).Value2
                PrevVal = range(use_Structured(PrevCellR, 2)).Value2
                NextVal = range(use_Structured(NextCellR, 2)).Value2
                toDATEVal = dateValue(FirstCell3) + TimeValue(FirstCell3)
                
                  
                
    '            DATEVal = Range(use_Structured(Range(DATECell), 4)).Value2
    '            PrevVal = Range(use_Structured(Range(PrevCell), 2)).Value2
    '            NextVal = Range(use_Structured(Range(NextCell), 2)).Value2
    '            toDATEVal = Range(use_Structured(Range(toDATECell), 4)).Value2
                
                
                Dim offset As Double
                offset = toDATEVal - DATEVal
                
                range(use_Structured(PrevCellR, 2)).Value = PrevVal + offset
                range(use_Structured(NextCellR, 2)).Value = NextVal - offset
                range(use_Structured(DATECellR, 4)).Value = toDATEVal
            
            
            End If

            

End Sub



Sub Completed()

    Call TaskComplete(Selection)
    samerowsOf(Selection, range("表格2[實際耗時]")).Value2 = samerowsOf(Selection, range("表格2[實際耗時]")).Value2
End Sub
Sub CalculateNext(Start As Variant)
    Dim cell As range
    Set cell = Start
    Dim tocal As range
    Set tocal = cell
    Set cell = cell.offset(1)
    Do While range(use_Structured(cell, 4)).HasFormula
        Set tocal = Union(cell, tocal)
        Set cell = cell.offset(1)
    Loop
    Call CalculateTable2ByOrder(tocal)
End Sub
Sub CompleteNow()
    Dim cell As range
    Set cell = Selection
    Call TaskComplete(cell)
    range(use_Structured(cell, 3)).Value2 = Now - range(use_Structured(cell, 4)).Value2
    
    Call CalculateNext(cell)
End Sub
Sub StartNow()
    Dim cell As range
    Set cell = Selection
    range(use_Structured(cell, 4)).Value2 = Now
    
    Call CalculateNext(cell)

End Sub
Sub trgey()
    MsgBox AddressEx(samerowsOf(Selection, range("表格2[交易物件]")))
End Sub
Sub TaskComplete(cell As range)
samerowsOf(cell, range("表格2[currResource]")).Value2 = samerowsOf(cell, range("表格2[currResource]")).Value2
samerowsOf(cell, range("表格2[Description]")).Value2 = samerowsOf(cell, range("表格2[Description]")).Value2
samerowsOf(cell, range("表格2[Location]")).Value2 = samerowsOf(cell, range("表格2[Location]")).Value2
samerowsOf(cell, range("表格2[實際百分比]")).Value2 = samerowsOf(cell, range("表格2[實際百分比]")).Value2
samerowsOf(cell, range("表格2[進度]")).Value2 = samerowsOf(cell, range("表格2[進度]")).Value2
samerowsOf(cell, range("表格2[起始百分比]")).Value2 = samerowsOf(cell, range("表格2[起始百分比]")).Value2
samerowsOf(cell, range("表格2[時區]")).Value2 = samerowsOf(cell, range("表格2[時區]")).Value2
samerowsOf(cell, range("表格2[SU]")).Value2 = samerowsOf(cell, range("表格2[SU]")).Value2
samerowsOf(cell, range("表格2[SU-MIN]")).Value2 = samerowsOf(cell, range("表格2[SU-MIN]")).Value2
samerowsOf(cell, range("表格2[完整耗時]")).Value2 = samerowsOf(cell, range("表格2[完整耗時]")).Value2
samerowsOf(cell, range("表格2[剩餘時間]")).Value2 = samerowsOf(cell, range("表格2[剩餘時間]")).Value2
samerowsOf(cell, range("表格2[現在預計進度]")).Value2 = samerowsOf(cell, range("表格2[現在預計進度]")).Value2
samerowsOf(cell, range("表格2[預計百分比]")).Value2 = samerowsOf(cell, range("表格2[預計百分比]")).Value2
samerowsOf(cell, range("表格2[起始百分比]")).Value2 = samerowsOf(cell, range("表格2[起始百分比]")).Value2
samerowsOf(cell, range("表格2[至完成還有]")).Value2 = samerowsOf(cell, range("表格2[至完成還有]")).Value2
samerowsOf(cell, range("表格2[已耗時]")).Value2 = samerowsOf(cell, range("表格2[已耗時]")).Value2
samerowsOf(cell, range("表格2[已節省]")).Value2 = samerowsOf(cell, range("表格2[已節省]")).Value2
samerowsOf(cell, range("表格2[Dist. To Avg]")).Value2 = samerowsOf(cell, range("表格2[Dist. To Avg]")).Value2
samerowsOf(cell, range("表格2[分進度(%/min)]")).Value2 = samerowsOf(cell, range("表格2[分進度(%/min)]")).Value2
samerowsOf(cell, range("表格2[Probability]")).Value2 = samerowsOf(cell, range("表格2[Probability]")).Value2
samerowsOf(cell, range("表格2[執行率]")).Value2 = samerowsOf(cell, range("表格2[執行率]")).Value2
samerowsOf(cell, range("表格2[Subject]")).Value2 = samerowsOf(cell, range("表格2[Subject]")).Value2
samerowsOf(cell, range("表格2[Location Verify]")).Value2 = samerowsOf(cell, range("表格2[Location Verify]")).Value2
samerowsOf(cell, range("表格2[Chain Verify]")).Value2 = samerowsOf(cell, range("表格2[Chain Verify]")).Value2
samerowsOf(cell, range("表格2[Dependency Verify]")).Value2 = samerowsOf(cell, range("表格2[Dependency Verify]")).Value2
samerowsOf(cell, range("表格2[Concurrency]")).Value2 = samerowsOf(cell, range("表格2[Concurrency]")).Value2
samerowsOf(cell, range("表格2[Certainty]")).Value2 = samerowsOf(cell, range("表格2[Certainty]")).Value2
samerowsOf(cell, range("表格2[Buffer]")).Value2 = samerowsOf(cell, range("表格2[Buffer]")).Value2
samerowsOf(cell, range("表格2[Chain Blanks]")).Value2 = samerowsOf(cell, range("表格2[Chain Blanks]")).Value2
samerowsOf(cell, range("表格2[Buffer]")).Value2 = samerowsOf(cell, range("表格2[Buffer]")).Value2
samerowsOf(cell, range("表格2[Dependency]")).Value2 = samerowsOf(cell, range("表格2[Dependency]")).Value2
samerowsOf(cell, range("表格2[Start Date]")).Value2 = samerowsOf(cell, range("表格2[Start Date]")).Value2
samerowsOf(cell, range("表格2[End Date]")).Value2 = samerowsOf(cell, range("表格2[End Date]")).Value2
samerowsOf(cell, range("表格2[Start Time]")).Value2 = samerowsOf(cell, range("表格2[Start Time]")).Value2
samerowsOf(cell, range("表格2[End Time]")).Value2 = samerowsOf(cell, range("表格2[End Time]")).Value2



End Sub
Sub FitDuration()

    For Each cell In Selection
        Dim r As range
        Set r = cell
        Set StartDate = StructureAboveR("End Date", r)
        Set EndDate = StructureBelowR("Start Date", r)
        s = StartDate.address
        f = EndDate.address
        If (range(use_Structured(r, 4)).HasFormula) Then range(use_Structured(r, 4)).Value2 = StartDate
        range(use_Structured(r, 2)).Value2 = EndDate.Value2 - range(use_Structured(r, 4)).Value2
    Next cell
    Call CalculateNext(Selection(1))
End Sub


Sub Connect2Next()
'If Selection.Cells.count = 2 Then
    Dim mysel As range
    Set mysel = Selection
    Dim fromC As String
    Dim toC As String
    Dim TimeBetween As Double
    Dim lastRow As Long
    
    For Each cell In mysel
        lastRow = range(use_Structured(cell, 0)).Value2
    Next cell
    For Each cell In mysel
        cell.Select
        TimeBetween = TimeBetween + Evaluate("表格2[[#This Row], [實際耗時]:[實際耗時]]")
        Set r = Evaluate("表格2[[#This Row], [編號]:[編號]]")
        r.Value2 = lastRow
        If fromC = "" Then
            fromC = cell.address
        Else
            toC = cell.address
        End If
    Next cell
    
    Dim nextStart As Double
    Dim Duration As Double

    range(toC).Select
    Set r = Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]")
    nextStart = r.Value
    Set r = Evaluate("表格2[[#This Row], [實際耗時]:[實際耗時]]")
    Duration = r.Value


    range(fromC).Select
    
    Set q = Evaluate("表格2[[#This Row], [Start Date]:[Start Date]]")
    q.Value = nextStart - (TimeBetween - Duration)

    Call SortRowNum
    'Call CalculateRange2
'End If
End Sub

Sub SelectPlanRange()
    range(range("交易!AM2")).Select
End Sub
Sub SelectC3Range()
    range(range("C3")).Select
End Sub
Sub FillInCells()
    Dim toMove_time As range
    Set toMove_time = Application.InputBox(Prompt:="toMove_time", title:="toMove_time", Type:=8)
    Dim availible_time As range
    Set availible_time = Application.InputBox(Prompt:="availible_time", title:="availible_time", Type:=8)
    'Call ScrollToEnd
    Dim toPut_loc As range
    Set toPut_loc = Application.InputBox(Prompt:="toPut_loc", title:="toPut_loc", Type:=8)
    
    
    Dim toMove_time_c As New Collection
    Dim toMove_job_c As New Collection
    Dim toMove_des_c As New Collection
    Dim toMove_startPercent_c As New Collection
    Dim toMove_endPercent_c As New Collection
    'toMove_time_arr = toMove_time.Value
    For i = toMove_time.Cells.Count To 1 Step -1
        toMove_time_c.Add (range(use_Structured(getItemByIndexInRange(toMove_time, i), 2)).Value2)
        toMove_job_c.Add (range(use_Structured(getItemByIndexInRange(toMove_time, i), 6)).Value2)
        toMove_des_c.Add (range(use_Structured(getItemByIndexInRange(toMove_time, i), 9)).Value2)
        toMove_startPercent_c.Add (range(use_Structured(getItemByIndexInRange(toMove_time, i), 7)).Value2)
        toMove_endPercent_c.Add (range(use_Structured(getItemByIndexInRange(toMove_time, i), 8)).Value2)
        Debug.Print toMove_time_c.Item(toMove_time_c.Count)
    Next i
    

    Dim availible_time_c As New Collection
    Dim availible_start_c As New Collection
    'availible_time_arr = availible_time.Value
    For i = availible_time.Cells.Count To 1 Step -1
        availible_time_c.Add (range(use_Structured(getItemByIndexInRange(availible_time, i), 2)).Value2)
        availible_start_c.Add (range(use_Structured(getItemByIndexInRange(availible_time, i), 4)).Value2)
        range(use_Structured(getItemByIndexInRange(availible_time, i), 2)).Value2 = 0
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
    If toMove_startPercent_c.Count > 0 Then adjustedstartPercent = lastCol(toMove_startPercent_c)
    Do While (toMove_job_c.Count > 0) And (availible_time_c.Count > 0)
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
            availible_time_c.Remove (availible_time_c.Count)
            availible_time_c.Add (remainingAvailible)
            adjustedStart = lastCol(availible_start_c) + lastCol(toMove_time_c)
            availible_start_c.Remove (availible_start_c.Count)
            availible_start_c.Add (adjustedStart)
'            Set availible_time_c.Item(1) = availible_time_c(1) - toMove_time_c(1)
'            Set availible_start_c.Item(1) = availible_start_c(1) + toMove_time_c(1)
            toMove_time_c.Remove (toMove_time_c.Count)
            toMove_job_c.Remove (toMove_job_c.Count)
            toMove_des_c.Remove (toMove_des_c.Count)
            toMove_startPercent_c.Remove (toMove_startPercent_c.Count)
            toMove_endPercent_c.Remove (toMove_endPercent_c.Count)
            If toMove_startPercent_c.Count > 0 Then adjustedstartPercent = lastCol(toMove_startPercent_c)
        ElseIf lastCol(toMove_time_c) = lastCol(availible_time_c) Then
            toPut_job_c.Add (lastCol(toMove_job_c))
            toPut_des_c.Add (lastCol(toMove_des_c))
            toPut_start_c.Add (lastCol(availible_start_c))
            toPut_time_c.Add (lastCol(toMove_time_c))
            toPut_startPercent_c.Add (adjustedstartPercent)
            toPut_endPercent_c.Add (lastCol(toMove_endPercent_c))
            availible_start_c.Remove (availible_start_c.Count)
            availible_time_c.Remove (availible_time_c.Count)
            toMove_time_c.Remove (toMove_time_c.Count)
            toMove_job_c.Remove (toMove_job_c.Count)
            toMove_des_c.Remove (toMove_des_c.Count)
            toMove_startPercent_c.Remove (toMove_startPercent_c.Count)
            toMove_endPercent_c.Remove (toMove_endPercent_c.Count)
            If toMove_startPercent_c.Count > 0 Then adjustedstartPercent = lastCol(toMove_startPercent_c)
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
            toMove_time_c.Remove (toMove_time_c.Count)
            toMove_time_c.Add (adjustedtoMovetime)
            
            availible_start_c.Remove (availible_start_c.Count)
            availible_time_c.Remove (availible_time_c.Count)
        End If
        Debug.Print "xx"
        Debug.Print lastCol(toPut_startPercent_c)
        'Debug.Print "xxxxxx"
        'Debug.Print (CStr(toPut_job_c(toPut_job_c.count)) & " " & CStr(toPut_start_c(toPut_start_c.count)) & " " & CStr(toPut_time_c(toPut_time_c.count)))
        If Not ((toMove_job_c.Count > 0) And (availible_time_c.Count > 0)) Then
            Exit Do
        End If
    Loop
    If availible_time_c.Count > 0 Then
        For i = availible_time_c.Count To 1 Step -1
            toPut_time_c.Add (availible_time_c(i))
            toPut_start_c.Add (availible_start_c(i))
            toPut_job_c.Add ("時序專案(288)")
            toPut_des_c.Add ("")
        Next i
    End If
    For i = toPut_job_c.Count To 1 Step -1
        range(use_Structured(toPut_loc.offset(i - 1), 2)).Value2 = toPut_time_c(i)
        range(use_Structured(toPut_loc.offset(i - 1), 4)).Value2 = toPut_start_c(i)
        range(use_Structured(toPut_loc.offset(i - 1), 6)).Value2 = toPut_job_c(i)
        range(use_Structured(toPut_loc.offset(i - 1), 9)).Value2 = toPut_des_c(i)
        On Error Resume Next
        range(use_Structured(toPut_loc.offset(i - 1), 7)).Value2 = toPut_startPercent_c(i)
        On Error Resume Next
        range(use_Structured(toPut_loc.offset(i - 1), 8)).Value2 = toPut_endPercent_c(i)
    Next i
    'Set availible = Application.InputBox(Prompt:="availible", Title:="availible", Type:=8)
    'Set toPut = Application.InputBox(Prompt:="toPut", Title:="toPut", Type:=8)
        
        
    'select toMove
    
    'select availible
    'select toPut

End Sub
Public Function lastCol(C As Collection) As Variant
    lastCol = C.Item(C.Count)
End Function
Function FirstExample() As range
    Set FirstExample = range("表格2[交易物件]").Find(What:="_Example", LookIn:=xlValues)
End Function
Sub pasteExample(icell As range)
    Dim sel As range
    Dim selectedC As range
    For i = 1 To icell.Cells.Count
        Set selectedC = icell.Cells(i)
        Set sel = range(use_Structured(selectedC, 0)) '.Resize(1, Range("表格2").Columns.count)
        Debug.Print AddressEx(range("表格2").Cells(1).Resize(1, range("表格2").Columns.Count))
        range("表格2").offset(1, 0).Resize(1, range("表格2").Columns.Count).Copy
        sel.PasteSpecial Paste:=xlPasteFormulas
        sel.Value2 = 10000
    Next i
    
End Sub
Sub pasteExampl2e()
    Call pasteExample(Selection)
End Sub
Sub Consolidate()
    Dim sel As range
    Set sel = Selection
    Dim selectedC As range
    Dim selectedC2 As range
    For i = 1 To sel.Cells.Count - 1
        Set selectedC = sel.Cells(i)
        Set selectedC2 = sel.Cells(i + 1)
        If (range(use_Structured(selectedC, 6)).Value = range(use_Structured(selectedC2, 6)).Value) And _
            (range(use_Structured(selectedC, 9)).Value = range(use_Structured(selectedC2, 9)).Value) And _
            (range(use_Structured(selectedC, 5)).Value = range(use_Structured(selectedC2, 4)).Value) Then
            range(use_Structured(selectedC2, 2)).Value2 = range(use_Structured(selectedC2, 2)).Value2 + range(use_Structured(selectedC, 2)).Value2
            range(use_Structured(selectedC2, 4)).Value2 = range(use_Structured(selectedC, 4)).Value2
            range(use_Structured(selectedC2, 7)).Value2 = range(use_Structured(selectedC, 7)).Value2
            Call pasteExample(selectedC)
            
        End If
        
        
    Next i
    Call SortRowNum
End Sub

Sub SortDate()
Call generateID
Application.CutCopyMode = False
Dim targetID As String
targetID = Evaluate("表格2[[#This Row], [ID]:[ID]]").Value2

Dim ws As Worksheet
Set ws = ActiveSheet
Dim tbl As ListObject
Set tbl = ws.ListObjects("表格2")
Dim sortcolumn As range
Set sortcolumn = range("表格2[Start Date]")
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .header = xlYes
   .Apply
End With

Set r = Evaluate("表格2[編號]")
r.Calculate

Set tbl = ws.ListObjects("表格2")
Set sortcolumn = range("表格2[編號]")
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .header = xlYes
   .Apply
End With

Call CalculateRange1

Set r = Evaluate("表格2[編號]")
Evaluate("表格2[編號]").Value2 = "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))+1"
range("$A$4").Value2 = 1
Set r = Evaluate("表格2[編號]")
r.Calculate


Call MoveToCurrentRow
    
End Sub


Sub SortRowNum()
Call ClearId
Application.CutCopyMode = False
'Dim targetID As String
'targetID = Range(use_Structured(Selection, 10)).Value2 'Evaluate("表格2[[#This Row], [ID]:[ID]]").Value2


Dim ws As Worksheet
Set ws = ActiveSheet
Dim tbl As ListObject
Set tbl = ws.ListObjects("表格2")
Set sortcolumn = range("表格2[編號]")
With tbl.Sort
   .SortFields.Clear
   .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
   .header = xlYes
   .Apply
End With



Set r = Evaluate("表格2[編號]")
Evaluate("表格2[編號]").Value2 = "=INDIRECT(ADDRESS(ROW()-1,COLUMN()))+1"
range("$A$4").Value2 = 1
Set r = Evaluate("表格2[編號]")
r.Calculate

Call CalculateRange1

'Dim finStr As String
'finStr = evals("=" + Replace(Range("$V$2").Value2, "rownum", targetID))
'Range(finStr).Select

Call generateID
ActiveWindow.ScrollRow = Selection.Row
ActiveWindow.ScrollColumn = Selection.Column
    
End Sub

Sub SortRowNumBySel()
    Dim mselection As range

    Set mselection = Selection
    fromCell = use_Structured(mselection.Cells(1), 0)
    toCell = use_Structured(mselection.Cells(2), 0)
    range(fromCell).Value2 = range(toCell).Value2
    Call SortRowNum
End Sub

Sub FillPercent()
    totalTime = 0
    lastPercent = 0
    Set mselection = Selection
    For Each cell In mselection
        totalTime = totalTime + range(use_Structured(cell, 2)).Value2
        lastPercent = range(use_Structured(cell, 8)).Value2
    Next cell
    
    prevPercent = range(use_Structured(mselection(1), 7)).Value2
    firstPercent = range(use_Structured(mselection(1), 7)).Value2
    For Each cell In mselection
        Delta = (lastPercent - firstPercent) * (range(use_Structured(cell, 2)).Value2 / totalTime)
        range(use_Structured(cell, 8)).Value2 = prevPercent + Delta
        prevPercent = range(use_Structured(cell, 8)).Value2
    Next cell
    
End Sub


Sub SwapCells()
    Dim selected As range
    Set selected = Selection
    If selected.Areas.Count > 1 Then
        Dim row1 As range
        Dim row2 As range
        Set row1 = selected.Areas(1)
        Set row2 = selected.Areas(2)
        For i = 1 To row1.Cells.Count
            tmpval = row1.Cells(i).formula
            row1.Cells(i).formula = row2.Cells(i).formula
            row2.Cells(i).formula = tmpval
        Next i
        
    
    End If
    
    
End Sub
