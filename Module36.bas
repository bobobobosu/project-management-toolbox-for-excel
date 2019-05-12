Attribute VB_Name = "Module36"
'Variables
Function experienceFile() As String
    experienceFile = ThisWorkbook.path & "\" & "TCdata" & "\" & ReplaceIllegalCharacters(Evaluate("INDEX(表格2[交易物件],MATCH($A$4,表格2[ID],0))"), "_") & ".csv"
End Function
Function taskFile(Optional id As Variant) As String
    range("進度!C4").Calculate
    If IsMissing(id) Then
        taskFile = ThisWorkbook.path & "\" & "TCdata" & "\" & "NowPercent" & "\" & ReplaceIllegalCharacters(taskFileName(), "_") & ".csv"
    Else
        taskFile = ThisWorkbook.path & "\" & "TCdata" & "\" & "NowPercent" & "\" & ReplaceIllegalCharacters(taskFileName(id), "_") & ".csv"
    End If
    
End Function
Function taskFileName(Optional id As Variant)
    If IsMissing(id) Then
        taskFileName = evals("=CONCATENATE(TEXT(INDEX(表格2[Start Date],MATCH($A$4,表格2[ID],0))," & Chr(34) & "yyyymmddhhmm;@" & Chr(34) & ")," & Chr(34) & "_" & Chr(34) & ",INDEX(表格2[交易物件],MATCH($A$4,表格2[ID],0))," & Chr(34) & "_" & Chr(34) & ",INDEX(表格2[ID],MATCH($A$4,表格2[ID],0)))")
    Else
        taskFileName = evals("=CONCATENATE(TEXT(INDEX(表格2[Start Date],MATCH(" + CStr(id) + ",表格2[ID],0))," & Chr(34) & "yyyymmddhhmm;@" & Chr(34) & ")," & Chr(34) & "_" & Chr(34) & ",INDEX(表格2[交易物件],MATCH(" + CStr(id) + ",表格2[ID],0))," & Chr(34) & "_" & Chr(34) & ",INDEX(表格2[ID],MATCH(" + CStr(id) + ",表格2[ID],0)))")
    End If
End Function
Function RangeToString(r As Variant) As Variant
    s = ""
    If r Is Nothing Then
        RangeToString = s
        Exit Function
    End If
    For Each cell In r
        If cell.Column = r.Columns(r.Columns.Count).Column Then
            s = s + CStr(cell.Value2) + ";"
        Else
            s = s + CStr(cell.Value2) + ","
        End If
    Next
    RangeToString = Left(s, Len(s) - 1)
End Function

Function MultiSplit(s As Variant, Optional delim1 As String = ",", Optional delim2 As String = ";") As Variant
    Dim v As Variant, W As Variant, A As Variant
    Dim i As Long, j As Long, m As Long, n As Long
    v = Split(s, delim2)
    m = UBound(v)
    n = UBound(Split(v(0), delim1))
    ReDim A(0 To m, 0 To n)
    For i = 0 To m
        For j = 0 To n
            W = Split(v(i), delim1)
            A(i, j) = Trim(W(j))
        Next j
    Next i
    MultiSplit = A
End Function

Function CellToRange(r As range)
    
End Function
Function getTotalTasks(milestone As Variant) As Variant
    substoneCount = use_Structured2R(milestone, "NowPercent", "MileStone Count").Value2
    taskcount = use_Structured2R(milestone, "NowPercent", "Task Count").Value2
    
    Set pointerStone = milestone
    
    Do While substoneCount > 0
        Set pointerStone = pointerStone.offset(-1)
        taskcount = taskcount + use_Structured2R(pointerStone, "NowPercent", "Task Count").Value2
        If pointerStone <> vbNullString Then substoneCount = substoneCount - 1
    Loop
    getTotalTasks = taskcount
End Function

Function currentMileStone() As Variant
    Set currentMileStone = range(evals("AddressEx(INDEX(NowPercent[Milestone],MATCH(MAX(NowPercent[Actual]),NowPercent[Planned])+1))"))
    Do While currentMileStone.Value2 = vbNullString
        Set currentMileStone = currentMileStone.offset(1)
    Loop
End Function
Function getTimeDelta(StartTime As Variant, EndTime As Variant, coverage As Collection)

End Function

Function MilstoneofTask(task As Variant) As Variant
    Set milestone = use_Structured2R(task, "NowPercent", "Up MileStone")
    
    Set milestonepointerR = range(evals("AddressEx(INDEX(NowPercent[Task Count],MATCH(" + """" + milestone.Value2 + """" + ",NowPercent[Milestone],0)))"))
'    Set milestonepointerR = use_Structured2R(task, "NowPercent", "MileStone")
'
'    found = False
'    Do While milestonepointerR.Row < (Range("NowPercent").Row + Range("NowPercent").Rows.Count)
'        'If use_Structured2R(milestonepointerR, "NowPercent", "Actual").Value2 <> vbNullString And
'        If milestonepointerR.Value2 = milestone.Value2 Then
'            found = True
'            Exit Do
'        End If
'        Set milestonepointerR = milestonepointerR.offset(1)
'    Loop
    If Not milestonepointerR Is Nothing Then Set MilstoneofTask = milestonepointerR Else Set MilstoneofTask = Nothing
End Function

Function MilestoneStart(milestone As Variant) As Variant
    actualMilestone = use_Structured2R(milestone, "NowPercent", "Actual").Value2
    plannedMilestone = use_Structured2R(milestone, "NowPercent", "Planned").Value2
    If IsNumeric(actualMilestone) Then Position = actualMilestone
    If IsNumeric(plannedMilestone) Then Position = plannedMilestone
    MilestoneStartActual = Position - _
                            use_Structured2R(milestone, "NowPercent", "Task Count").Value2
    evfdv = "AddressEx(INDEX(NowPercent[Milestone],MATCH(" + CStr(MilestoneStartActual) + ",NowPercent[Actual],1)))"
    Set MilestoneStart = range(evals("AddressEx(INDEX(NowPercent[Milestone],MATCH(" + CStr(MilestoneStartActual) + ",NowPercent[Planned],1)))"))
End Function

Sub AddRealRecordNowByOne()
    Call AddRealRecord(Now(), 1)
End Sub
Sub AddRealRecord(NowTime As Variant, NowDeltaTask As Variant)
    If NowDeltaTask = 0 Then Exit Sub
    
    currentPosition = getCurrentActual()
    nextPlanned = evals("=MINIFS(NowPercent[Planned],NowPercent[Planned]," + """>""" + "&" + CStr(currentPosition) + ")")
    Do While NowDeltaTask + currentPosition > nextPlanned
        currentPosition = getCurrentActual()
        nextPlanned = evals("=MINIFS(NowPercent[Planned],NowPercent[Planned]," + """>""" + "&" + CStr(currentPosition) + ")")
        Call AddRealRecord(NowTime, nextPlanned - currentPosition)
        NowDeltaTask = NowDeltaTask - (nextPlanned - currentPosition)
    Loop
    
    taskCompleted = NowDeltaTask
    'use_Structured2R(currentMileStone, "NowPercent", "Task Count").Value2 = use_Structured2R(currentMileStone, "NowPercent", "Task Count").Value2 - taskCompleted

    Dim x(9)
    If NowTime = vbNullString Then
        x(0) = range(evals("AddressEx(INDEX(NowPercent[Time],MATCH(" + CStr(getCurrentActual() + taskCompleted) + ",NowPercent[Planned],1)))"))
    Else
         x(0) = NowTime
    End If
    x(1) = "" ' getCurrentActual() + taskCompleted
    x(2) = getCurrentActual() + taskCompleted
    x(3) = ""
    x(4) = ""
    x(5) = use_Structured2R(currentMileStone, "NowPercent", "Task").Value2
    x(6) = taskCompleted
    x(7) = use_Structured2R(currentMileStone, "NowPercent", "Up MileStone").Value2
    x(8) = x(2)
    x(9) = ""
    
    
    
    'Apply To all milestones
    Dim milestones As New Collection
    Dim milestoneData(4)
    milestoneData(1) = x(4)
    milestoneData(2) = x(5)
    milestoneData(3) = x(7)
    milestoneData(4) = x(6)
    milestones.Add milestoneData
    For Each cell In range("NowPercent[Planned]")
        If use_Structured2R(cell, "NowPercent", "Planned").Value2 = x(2) And _
            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "Start" And _
            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "End" Then
            milestoneData(1) = use_Structured2R(cell, "NowPercent", "Milestone").Value2
            milestoneData(2) = use_Structured2R(cell, "NowPercent", "Task").Value2
            milestoneData(3) = use_Structured2R(cell, "NowPercent", "Up MileStone").Value2
            milestoneData(4) = use_Structured2R(cell, "NowPercent", "Task Count").Value2
             milestones.Add milestoneData
        End If
    Next
    For j = milestones.Count To 1 Step -1
        x(4) = milestones.Item(j)(1)
        x(5) = milestones.Item(j)(2)
        x(7) = milestones.Item(j)(3)
        x(6) = milestones.Item(j)(4)
        AddDataRow "NowPercent", x, True
    Next
    
    Call UpdateThisTask(True)
    Call SetAutoCalculate
    Call SortPercent
    
    
End Sub
Sub egvrbtt()
    MsgBox samerowsOf(Selection, range())
End Sub
Sub SetAutoCalculate()
'    ggertg = samerowsOf(getAutoCalculateR(), Range("NowPercent[Time Elapse]"))
    samerowsOf(getAutoCalculateR(), range("NowPercent[Time Elapse]")).Value2 = "=getElapsedTime(INDEX([Time],MATCH(""Start"",[Milestone],0)),[@Time])"
    samerowsOf(getAutoCalculateR(), range("NowPercent[Percentage]")).Value2 = "=getTaskPercentInMilestone([@Time])"
    samerowsOf(getAutoCalculateR(), range("NowPercent[Sort]")).Value2 = "=IF([@Milestone]=""Start"",-1,IF([@Milestone]=""End"",1000,IF([@Planned]="""",[@Actual],[@Planned])))"
    'Range("NowPercent[Sort]").Value2 = "=IF([@Milestone]=""Start"",-1,IF([@Milestone]=""End"",1000,IF([@Planned]="""",[@Actual],[@Planned])))"
End Sub
Sub DeleteFakePlanned()
    For Each cell In range("NowPercent[Actual]")
        If cell.Value2 <> vbNullString And _
        use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "Start" And _
        use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "End" Then _
            use_Structured2R(cell, "NowPercent", "Planned").Value2 = vbNullString
    Next
End Sub

Function samerowsOf(reference As Variant, thisRange As Variant) As range
    Dim result As range
    For Each cell In reference.Cells
            If result Is Nothing Then
                Set result = Worksheets(thisRange.Parent.name).Cells(cell.Row, thisRange.Cells(1).Column)
            End If
            
            Set result = Union(result, Worksheets(thisRange.Parent.name).Cells(cell.Row, thisRange.Cells(1).Column))
    Next

    Set samerowsOf = result
End Function

Function samecolumnsOf(reference As Variant, thisRange As Variant) As range
    wregdbfnhj = thisRange.Parent.name
    Dim result As range
    For Each cell In reference.Cells
            If result Is Nothing Then
                Set result = Worksheets(thisRange.Parent.name).Cells(thisRange.Cells(1).Row, cell.Column)
            End If
            
            Set result = Union(result, Worksheets(thisRange.Parent.name).Cells(thisRange.Cells(1).Row, cell.Column))
    Next

    Set samecolumnsOf = result
End Function


Function GeneratePlan(r As range, Optional mode As Integer = 1)
    result = planAvalible(samerowsOf(r, range("表格2[編號]")), _
                                            samerowsOf(r, range("表格2[交易物件]")), _
                                            samerowsOf(r, range("表格2[Start Date]")), _
                                            samerowsOf(r, range("表格2[End Date]")), _
                                            Evaluate("=ISFORMULA(" + AddressEx(samerowsOf(r, range("表格2[Start Date]"))) + ")"), _
                                            samerowsOf(r, range("表格2[Location]")), _
                                            samerowsOf(r, range("表格2[currResource]")), _
                                            samerowsOf(r, range("表格2[accuResource]")), mode)
    GeneratePlan = result
End Function
Function getAutoCalculateR() As range
    Dim AutoCalculateR As range

    For Each cell In range("NowPercent[Time]")
        If use_Structured2R(cell, "NowPercent", "Time").Value2 <> vbNullString Then

            If AutoCalculateR Is Nothing Then
                Set AutoCalculateR = range("NowPercent").Rows(cell.Row - range("NowPercent").Row + 1)
            Else
                
                Set AutoCalculateR = Union(AutoCalculateR, range("NowPercent").Rows(cell.Row - range("NowPercent").Row + 1))
            End If
        End If
    Next
    waerfsgtgbdrtedf = AutoCalculateR.address
    Set getAutoCalculateR = AutoCalculateR
End Function


Function getPlannedR() As range
    Dim PlannedR As range

    For Each cell In range("NowPercent[Time]")
        If use_Structured2R(cell, "NowPercent", "Time").Value2 <> vbNullString And _
            use_Structured2R(cell, "NowPercent", "Planned").Value2 <> vbNullString And _
            use_Structured2R(cell, "NowPercent", "Actual").Value2 = vbNullString And _
            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "Start" And _
            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "End" Then

            If PlannedR Is Nothing Then
                Set PlannedR = range("NowPercent").Rows(cell.Row - range("NowPercent").Row + 1)
            Else
                
                Set PlannedR = Union(PlannedR, range("NowPercent").Rows(cell.Row - range("NowPercent").Row + 1))
            End If
        End If
    Next
    Set getPlannedR = PlannedR
End Function

Function getActualR() As range
    Dim ActualR As range

    For Each cell In range("NowPercent[Time]")
        If use_Structured2R(cell, "NowPercent", "Time").Value2 <> vbNullString And _
            use_Structured2R(cell, "NowPercent", "Planned").Value2 = vbNullString And _
            use_Structured2R(cell, "NowPercent", "Actual").Value2 <> vbNullString And _
            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "Start" And _
            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "End" Then
            
            If ActualR Is Nothing Then
                Set ActualR = range("NowPercent").Rows(cell.Row - range("NowPercent").Row + 1)
            Else
                
                Set ActualR = Union(ActualR, range("NowPercent").Rows(cell.Row - range("NowPercent").Row + 1))
            End If
        End If
    Next
    Set getActualR = ActualR
End Function

Function ExpandPreset() As range
    Dim preset As New Collection
    Dim presetR As range

    For Each cell In range("NowPercent[Time]")
        If use_Structured2R(cell, "NowPercent", "Time").Value2 = vbNullString And _
            use_Structured2R(cell, "NowPercent", "Planned").Value2 = vbNullString And _
            use_Structured2R(cell, "NowPercent", "Actual").Value2 = vbNullString And _
            use_Structured2R(cell, "NowPercent", "Percentage").Value2 <> vbNullString And _
            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> vbNullString And _
            use_Structured2R(cell, "NowPercent", "Task Count").Value2 <> vbNullString Then
            
            use_Structured2R(cell, "NowPercent", "Sort").Value2 = 9999
            x = use_Structured2R(cell, "NowPercent", "Percentage").Resize(1, 5).Value2
            preset.Add x
            
            If presetR Is Nothing Then
                Set presetR = range("NowPercent").Rows(cell.Row - range("NowPercent").Row + 1)
            Else
                
                Set presetR = Union(presetR, range("NowPercent").Rows(cell.Row - range("NowPercent").Row + 1))
            End If
        End If
        
    Next
    Set ExpandPreset = presetR
End Function

Function SavePreset()
    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    taskName = range(use_Structured(currentRange, 6)).Value2
    Dim targetRange As range
    Set targetRange = range(evals("AddressEx(INDEX(表格55[進度],MATCH(" + """" + taskName + """" + ",表格55[工作物件],0)))"))
    preset = RangeToString(ExpandPreset)
    range(use_Structured(currentRange, 15)).Value2 = preset
    targetRange.Value2 = preset
End Function

Function restorePreset()
    Set existingPreset = ExpandPreset()
    If Not existingPreset Is Nothing Then existingPreset.Value2 = vbNullString
    
    Dim currentRange As range
    Set currentRange = getStructuredByIDR(range("進度!$A$4").Value2, "表格2", "ID")
    preset = range(use_Structured(currentRange, 15)).Value2
    If preset = 0 Or preset = vbNullString Then Exit Function
    presetArr = MultiSplit(preset)
    
    
    For i = LBound(presetArr, 1) To UBound(presetArr, 1)
        Dim x(9)
        For j = 0 To 9
            x(j) = presetArr(i, j)
        Next
        Call AddDataRow("NowPercent", x, False)
    Next
End Function
Function ExpandTasks(milestone As Variant, tasks As Collection, TimePercentDelta As Variant, TimePercentStart As Variant) As Variant
    'Find Substone
    For milestoneRepeat = 1 To use_Structured2R(milestone, "NowPercent", "Task Count").Value2
        taskStartDelta = TimePercentDelta / use_Structured2R(milestone, "NowPercent", "Task Count").Value2
        taskStart = TimePercentStart + _
                            ((milestoneRepeat - 1) / use_Structured2R(milestone, "NowPercent", "Task Count").Value2) * TimePercentDelta
        prevtaskEnd = taskStart
    
        For Each cell In range("NowPercent[Up MileStone]")
            If cell.Value2 = milestone.Value2 And use_Structured2R(cell, "NowPercent", "Time").Value2 = vbNullString Then
                Dim x(9)
                If use_Structured2R(cell, "NowPercent", "Task").Value2 <> vbNullString Then
                    x(0) = taskStart + use_Structured2R(cell, "NowPercent", "Percentage").Value2 * taskStartDelta
                    x(1) = totalTasksExpand(tasks) + use_Structured2R(cell, "NowPercent", "Task Count").Value2
                    x(2) = ""
                    x(3) = use_Structured2R(cell, "NowPercent", "Percentage").Value2
                    x(4) = use_Structured2R(cell, "NowPercent", "Milestone").Value2
                    x(5) = use_Structured2R(cell, "NowPercent", "Task").Value2
                    x(6) = use_Structured2R(cell, "NowPercent", "Task Count").Value2
                    x(7) = use_Structured2R(cell, "NowPercent", "Up MileStone").Value2
                    x(8) = x(1)
                    x(9) = ""
                    tasks.Add x
                    ExpandTasks = ExpandTasks + x(6)
                Else

                   
                    If tasks.Count > 0 Then prevtaskEnd = tasks(tasks.Count)(0) Else prevtaskEnd = 0
                    MilestoneTaskCount = ExpandTasks(use_Structured2R(cell, "NowPercent", "Milestone"), _
                    tasks, _
                    taskStart + use_Structured2R(cell, "NowPercent", "Percentage").Value2 * taskStartDelta - prevtaskEnd, _
                    prevtaskEnd)
                    
                    
                    x(0) = taskStart + use_Structured2R(cell, "NowPercent", "Percentage").Value2 * taskStartDelta
                    x(1) = totalTasksExpand(tasks)
                    x(2) = ""
                    x(3) = use_Structured2R(cell, "NowPercent", "Percentage").Value2
                    x(4) = use_Structured2R(cell, "NowPercent", "Milestone").Value2
                    x(5) = use_Structured2R(cell, "NowPercent", "Task").Value2
                    x(6) = MilestoneTaskCount
                    x(7) = use_Structured2R(cell, "NowPercent", "Up MileStone").Value2
                    x(8) = x(1)
                    x(9) = ""
                    tasks.Add x
                    
                    
                    ExpandTasks = ExpandTasks + MilestoneTaskCount
                End If
            
                
            End If
        Next
    Next
End Function

Function totalTasksExpand(expand As Collection)
    totalTasksExpand = 0
    For Each x In expand
        If x(5) <> vbNullString Then totalTasksExpand = totalTasksExpand + x(6)
    Next
End Function
Sub ExpandNowPercent()
    'Clean
    Set PlannedR = getPlannedR
    If Not PlannedR Is Nothing Then getPlannedR.Value2 = vbNullString
'    For Each cell In Range("NowPercent[Time]")
'        If cell.Value <> vbNullString And _
'            use_Structured2R(cell, "NowPercent", "Actual").Value2 = vbNullString And _
'            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "Start" And _
'            use_Structured2R(cell, "NowPercent", "Milestone").Value2 <> "End" Then
'            cell.Resize(1, Range("NowPercent").Columns.Count).Value2 = vbNullString
'        End If
'    Next
    
    
    Call SortPercent
    Call OverwritePercentageByExperience

    Set endR = range(evals("=AddressEx(INDEX(NowPercent[Milestone],MATCH(""End"",NowPercent[Milestone],0)))"))
    Dim Queue As Collection
    Set Queue = New Collection
    Dim expand As Collection
    Set expand = New Collection
    
    'Clear End & Expand
     use_Structured2R(endR, "NowPercent", "Task Count").Value2 = 1
    Call SavePreset
    getByMilestone("End", "Task Count") = ExpandTasks(endR, expand, 1, 0)
    'Call SortByColumn(expand, 0)
    
    totalTasksCount = totalTasksExpand(expand)
    
    
    StartTime = getByMilestone("Start", "Time")
    EndTime = getByMilestone("End", "Time") - (1 / (60 * 24))
    startPlanned = getByMilestone("Start", "Planned")
    getByMilestone("End", "Planned") = startPlanned + totalTasksCount
    endPlanned = getByMilestone("End", "Planned") ' * (99 / 100)
    accuPlanned = startPlanned
    
    
    'Add Milestone Rows
    For i = 1 To expand.Count
        Dim y(8)
        x = expand.Item(i)
        If x(1) > startPlanned Then
          'Set Time
          x(0) = generateTime(StartTime, EndTime, x(0))
          'X(0) = startTime + deltaTime * X(0)
            
          'Set Data
          expand.Item(i)(1) = x(1) = x(1) + startPlanned
    
          For j = 0 To 8
              y(j) = x(j)
          Next
          Call AddDataRow("NowPercent", y, False)
        End If
    Next
    
    'Add Startpause Rows
    Dim Z(0)
    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    Dim TaskChain() As String
    TaskChain = Split(range(use_Structured(currentRange, 14)).Value2, ",")
    Dim activeRange As New Collection
    For k = LBound(TaskChain) To UBound(TaskChain)
        If TaskChain(k) <> vbNullString Then
            Z(0) = range(evals("AddressEx(INDEX(表格2[Start Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))"))
            activeRange.Add Z
            Z(0) = range(evals("AddressEx(INDEX(表格2[End Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))"))
            activeRange.Add Z
        End If
    Next
    Call SortByColumn(activeRange, 0)
    activeRange.Remove (1)
    activeRange.Remove (activeRange.Count)
    
    For k = 1 To activeRange.Count
            Dim W(8)
            W(0) = activeRange.Item(k)(0)
            W(1) = Round(getPlannedByTimePercent(expand, (getElapsedTime(StartTime, W(0)) / getElapsedTime(StartTime, EndTime))))
            W(2) = ""
            W(3) = ""
            W(4) = ""
            W(5) = ""
            W(6) = ""
            W(7) = ""
            W(8) = W(1)
            Call AddDataRow("NowPercent", W, False)
    Next
    
    
    
    
    Call SetAutoCalculate
    Call SortPercent
    Call ScaleChart8Auto
End Sub
Function generateTime(StartTime As Variant, EndTime As Variant, timepercent As Variant)
    generateTime = getEndTimebyElapsed(StartTime, getElapsedTime(StartTime, EndTime) * timepercent)
End Function
Function getByMilestone(milestone As String, title As String) As range
    Dim milestoneR As range
    Set milestoneR = range(evals("=AddressEx(INDEX(NowPercent[" + title + "],MATCH(""" & milestone & """,NowPercent[Milestone],0)))"))
    Set getByMilestone = use_Structured2R(milestoneR, "NowPercent", title)
End Function
Function getCurrentMilestone() As range
    Set getCurrentMilestone = range(evals("AddressEx(INDEX(NowPercent[Milestone],MATCH(MAX(NowPercent[Actual]),NowPercent[Planned],1)+1))"))
    Do While getCurrentMilestone.Value2 = vbNullString And _
        use_Structured2R(getCurrentMilestone, "NowPercent", "Planned").Value2 = vbNullString
        Set getCurrentMilestone = getCurrentMilestone.offset(1)
    Loop
End Function

Function getCurrentActual()
    getCurrentActual = Application.WorksheetFunction.Max(range("NowPercent[Actual]")) 'evals("MAX(NowPercent[Actual])")
End Function
Function getCurretnActualR() As range
    With Application.WorksheetFunction
        Set getCurretnActualR = .index(range("NowPercent[Actual]"), .Match(getCurrentActual, range("NowPercent[Actual]"), 0))
    End With
End Function
Function getPlannedByTime(time As Variant) As Variant
    If time > Application.WorksheetFunction.Max(range("NowPercent[Time]")) Then
        getPlannedByTime = Application.WorksheetFunction.Max(range("NowPercent[Planned]"), getCurrentActual(), 0)
        Exit Function
    End If
    
    Set prevTimeR = range(evals("AddressEx(INDEX(NowPercent[Time],MATCH(" + CStr(time) + ",NowPercent[Time],1)))"))
    Set nextTimeR = range(evals("AddressEx(INDEX(NowPercent[Time],MATCH(" + CStr(time) + ",NowPercent[Time],1)+1))"))

    Do While use_Structured2R(prevTimeR, "NowPercent", "Planned").Value2 = vbNullString
        Set prevTimeR = prevTimeR.offset(-1)
    Loop

    Do While use_Structured2R(nextTimeR, "NowPercent", "Planned").Value2 = vbNullString And _
            nextTimeR.Row <= range("NowPercent").Cells(range("NowPercent").Count).Row
        Set nextTimeR = nextTimeR.offset(1)
    Loop
    
    prevTime = use_Structured2R(prevTimeR, "NowPercent", "Time")
    nextTime = use_Structured2R(nextTimeR, "NowPercent", "Time")
    prevPlanned = use_Structured2R(prevTimeR, "NowPercent", "Planned")
    nextPlanned = use_Structured2R(nextTimeR, "NowPercent", "Planned")

    getPlannedByTime = prevPlanned + (nextPlanned - prevPlanned) * ((time - prevTime) / (nextTime - prevTime))
End Function
Function getPlannedByTimePercent(expand As Collection, timepercent As Double) As Variant
    getPlannedByTimePercent = Terp(timepercent, getColumnInCollection(expand, 0), getColumnInCollection(expand, 1))
End Function

Function getPrevMilestone() As range
    Set getPrevMilestone = getCurrentMilestone.offset(-1)
    Do While getPrevMilestone.Value2 = vbNullString
        Set getPrevMilestone = getPrevMilestone.offset(1)
    Loop
End Function

Function getStartStone() As range
    Set getStartStone = range(evals("=AddressEx(INDEX(NowPercent[Milestone],MATCH(""Start"",NowPercent[Milestone],0)))"))
End Function

Function getEndStone() As range
    Set getEndStone = range(evals("=AddressEx(INDEX(NowPercent[Milestone],MATCH(""End"",NowPercent[Milestone],0)))"))
End Function
Function getElapsedTime(StartTime As Variant, thisTime As Variant)
    If thisTime = vbNullString Then
        getElapsedTime = ""
        Exit Function
    End If
    
    Dim x(2)
    Dim timewithholes As New Collection
    x(1) = StartTime
    x(2) = thisTime
    timewithholes.Add x

    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    Dim TaskChain() As String
    TaskChain = Split(range(use_Structured(currentRange, 14)).Value2, ",")
    Dim activeRange As New Collection
    For k = LBound(TaskChain) To UBound(TaskChain)
        If TaskChain(k) <> vbNullString Then
            x(1) = range(evals("AddressEx(INDEX(表格2[Start Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))"))
            x(2) = range(evals("AddressEx(INDEX(表格2[End Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))"))
            activeRange.Add x
        End If
    Next
    Call SortByColumn(activeRange, 1)

    getElapsedTime = TimeDurationWithHoles(timewithholes, activeRange)
End Function
Function getEndTimebyElapsed(StartTime As Variant, Elapsed As Variant)
    If Elapsed = vbNullString Then
        getEndTimebyElapsed = ""
        Exit Function
    End If
    
    Dim x(2)
    Dim timewithholes As New Collection
    x(1) = StartTime
    x(2) = thisTime
    timewithholes.Add x

    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    Dim TaskChain() As String
    TaskChain = Split(range(use_Structured(currentRange, 14)).Value2, ",")
    Dim activeRange As New Collection
    For k = LBound(TaskChain) To UBound(TaskChain)
        If TaskChain(k) <> vbNullString Then
            x(1) = range(evals("AddressEx(INDEX(表格2[Start Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))"))
            x(2) = range(evals("AddressEx(INDEX(表格2[End Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))"))
            activeRange.Add x
        End If
    Next
    Call SortByColumn(activeRange, 1)
    
    
    getEndTimebyElapsed = StartTime
    For i = 1 To activeRange.Count
        If activeRange.Item(i)(2) < StartTime Then
        Else
            If activeRange.Item(i)(2) - activeRange.Item(i)(1) < Elapsed Then
                Elapsed = Elapsed - (activeRange.Item(i)(2) - activeRange.Item(i)(1))
                getEndTimebyElapsed = activeRange.Item(i)(2)
            Else
                getEndTimebyElapsed = activeRange.Item(i)(1) + Elapsed
                Exit Function
            End If
            
        End If
    Next

End Function
Function getTaskPercentInMilestone(task As Variant)
    On Error Resume Next

    Set milestone = MilstoneofTask(task)
    If milestone Is Nothing Then
        getTaskPercentInMilestone = ""
        Exit Function
    End If
    
    StartTime = use_Structured2R(MilestoneStart(milestone), "NowPercent", "Time").Value2
    EndTime = use_Structured2R(milestone, "NowPercent", "Time").Value2
    taskTime = use_Structured2R(task, "NowPercent", "Time").Value2
    
    mileStoneElapsedTime = getElapsedTime(StartTime, EndTime)
    taskElapsedTime = getElapsedTime(StartTime, taskTime)
    
    'If (endTime - startTime) > 0 Then getTaskPercentInMilestone = (taskTime - startTime) / (endTime - startTime) Else getTaskPercentInMilestone = -1
    If mileStoneElapsedTime > 0 Then getTaskPercentInMilestone = taskElapsedTime / mileStoneElapsedTime Else getTaskPercentInMilestone = ""
    
    
End Function
Function TimeDurationWithHoles(TimeWithNoHoles As Collection, activeRanges As Collection) As Variant
    TimeDurationWithHoles = 0
    StartTime = TimeWithNoHoles.Item(1)(1)
    EndTime = TimeWithNoHoles.Item(1)(2)
    For Each aRange In activeRanges
        overlapping = Smaller(EndTime, aRange(2)) - bigger(StartTime, aRange(1))
        If overlapping > 0 Then TimeDurationWithHoles = TimeDurationWithHoles + overlapping
    Next
End Function
Function bigger(A As Variant, B As Variant) As Variant
    If A > B Then
        bigger = A
    ElseIf A < B Then
        bigger = B
    Else
        bigger = A
    End If
End Function
Function Smaller(A As Variant, B As Variant) As Variant
    If A > B Then
        Smaller = B
    ElseIf A < B Then
        Smaller = A
    Else
        Smaller = A
    End If
End Function



Sub GenerateNewTable()
    
    Call UpdateThisTask
    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    StartTime = range(use_Structured(currentRange, 4)).Value2
    EndTime = range(use_Structured(currentRange, 4)).Value2 + range(use_Structured(currentRange, 2)).Value2
    startPercent = range(use_Structured(currentRange, 7)).Value2
    endPercent = range(use_Structured(currentRange, 8)).Value2
    NowPercent = range(use_Structured(currentRange, 28)).Value2

    Dim TaskChain() As String
    TaskChain = Split(range(use_Structured(currentRange, 14)).Value2, ",")
    
    target = range(use_Structured(currentRange, 15)).Value2
    id = range(use_Structured(currentRange, 10)).Value2

    range("NowPercent").Value = ""
    


    'Restore if exists
    Dim arrX As Variant
    Dim arrY As Variant
    For k = LBound(TaskChain) To UBound(TaskChain)
        If TaskChain(k) <> vbNullString Then
            StartTime = Smaller(StartTime, range(evals("AddressEx(INDEX(表格2[Start Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))")))
            EndTime = bigger(EndTime, range(evals("AddressEx(INDEX(表格2[End Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))")))
            taskTitle = range(evals("AddressEx(INDEX(表格2[交易物件],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))"))
            taskEndTime = range(evals("AddressEx(INDEX(表格2[End Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))"))
            If taskTitle = range(use_Structured(currentRange, 6)).Value2 Then
                If IsArrayAllocated(arrX) Then
                    arrY = ArrayFromCSVfile(taskFile(TaskChain(k)))
                    If IsArrayAllocated(arrY) Then arrX = Combine(arrX, arrY)
                Else
                    arrX = ArrayFromCSVfile(taskFile(TaskChain(k)))
                End If
            ElseIf range(evals("AddressEx(INDEX(表格2[實際百分比],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))")) = 1 Then
            
                
                Dim Z(0 To 0, 0 To 6)
                Z(0, 0) = taskEndTime
                Z(0, 1) = ""
                Z(0, 2) = 1
                Z(0, 3) = ""
                Z(0, 4) = taskTitle
                Z(0, 5) = taskTitle
                Z(0, 6) = 1


                If IsArrayAllocated(arrX) Then
                    'FindActual
                    actual = 0
                    For j = LBound(arrX, 1) To UBound(arrX, 1)
                        If IsNumeric(arrX(j, 2)) Then
                            If arrX(j, 2) > actual Then actual = arrX(j, 2)
                        End If
                    Next
                    Z(0, 2) = actual + 1
                    arrX = Combine(arrX, Z)
                Else
                    arrX = Z
                End If
            End If
        End If
    Next
    If IsArrayAllocated(arrX) Then Call Array2Range(arrX, range("NowPercent"))
    
    
    'Update Start End
    For k = LBound(TaskChain) To UBound(TaskChain)
        If TaskChain(k) <> vbNullString Then
            StartTime = Smaller(StartTime, range(evals("AddressEx(INDEX(表格2[Start Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))")))
            EndTime = bigger(EndTime, range(evals("AddressEx(INDEX(表格2[End Date],MATCH(" + CStr(TaskChain(k)) + ",表格2[ID],0)))")))
        End If
    Next
    For Each cell In range("NowPercent[MileStone]")
        If cell.Value = "Start" Then
            thisStart = use_Structured2R(cell, "NowPercent", "Time").Value2
            If use_Structured2R(cell, "NowPercent", "Time").Value2 < StartTime And _
                use_Structured2R(cell, "NowPercent", "Actual") <> vbNullString Then StartTime = use_Structured2R(cell, "NowPercent", "Time").Value2
            use_Structured2R(cell, "NowPercent", "Time").Resize(1, range("NowPercent").Columns.Count) = vbNullString
        End If
        If cell.Value = "End" Then
            If use_Structured2R(cell, "NowPercent", "Time").Value2 > EndTime And _
                use_Structured2R(cell, "NowPercent", "Actual") <> vbNullString Then EndTime = use_Structured2R(cell, "NowPercent", "Time").Value2
            use_Structured2R(cell, "NowPercent", "Time").Resize(1, range("NowPercent").Columns.Count) = vbNullString
        End If
    Next
    
    
    Dim x(9)
    x(0) = StartTime
    x(1) = 0
    x(2) = 0
    x(3) = ""
    x(4) = "Start"
    x(5) = ""  'startPercent
    x(6) = 1
    x(7) = 0
    x(8) = -1
    x(9) = "" ' "=getElapsedTime(INDEX([Time],MATCH(""Start"",[Milestone],0)),[@Time])"
     Dim y(9)
    y(0) = EndTime
    y(1) = ""
    y(2) = "" 'target * endPercent '"=MAX([Actual])"
    y(3) = ""
    y(4) = "End"
    y(5) = ""  'endPercent
    y(6) = 1
    y(7) = "" '"=COUNTA([Milestone])-2"
    y(8) = 1000
    y(9) = "" ' "=getElapsedTime(INDEX([Time],MATCH(""Start"",[Milestone],0)),[@Time])"
    AddDataRow "NowPercent", x, False
    AddDataRow "NowPercent", y, False



    



    Call SetAutoCalculate
    Call restorePreset
    Call OverwritePercentageByExperience
    Call SortPercent
    Call UpdateThisTask
    Call ScaleChart8Auto
End Sub

Sub OverwritePercentageByExperience()
'    Dim arrX As Variant
'    fd = experienceFile()
'    arrX = ArrayFromCSVfile(experienceFile())
'    If IsEmpty(arrX) Then Exit Sub
'
'    Dim milestone As dictionary
'    Set milestone = CreateObject("Scripting.Dictionary")
'
'
'
'    Dim tmpMilstones As Collection
'    Dim tmpUpMileStone As Collection
'    Dim tmpPercentage As Collection
'    For Each cell In Range("NowPercent[Time]")
'        If cell.Value2 = vbNullString And use_Structured2R(cell, "NowPercent", "Milestone") <> vbNullString Then
'            milestonandup = use_Structured2R(cell, "NowPercent", "Milestone").Value2 + "_" + _
'                            use_Structured2R(cell, "NowPercent", "Up MileStone").Value2
'            If Not milestone.Exists(milestonandup) Then
'                Set tmpPercentage = New Collection
'                milestone.Add milestonandup, tmpPercentage
'            End If
'        End If
'    Next
'
'    For rownum = UBound(arrX, 1) To 0 Step -1
'        milestonandup = arrX(rownum, 4) + "_" + arrX(rownum, 7)
'        If arrX(rownum, 4) <> "Start" And _
'            arrX(rownum, 4) <> "End" And _
'            arrX(rownum, 4) <> vbNullString And _
'            arrX(rownum, 2) <> vbNullString And _
'            arrX(rownum, 3) >= 0 And _
'            arrX(rownum, 3) <= 1 And _
'            milestone.Exists(milestonandup) Then
'                Set tmpPercentage = milestone(milestonandup)
'                milestone(milestonandup).Add arrX(rownum, 3)
'        End If
'    Next
'
'    For Each cell In Range("NowPercent[Time]")
'        If cell.Value2 = vbNullString And use_Structured2R(cell, "NowPercent", "Milestone") <> vbNullString Then
'            milestonandup = use_Structured2R(cell, "NowPercent", "Milestone").Value2 + "_" + _
'                            use_Structured2R(cell, "NowPercent", "Up MileStone").Value2
'            If milestone.Exists(milestonandup) Then
'                If milestone(milestonandup).Count > 0 Then use_Structured2R(cell, "NowPercent", "Percentage").Value2 = avg_Collection(milestone(milestonandup))
'            End If
'        End If
'    Next
     

End Sub
Sub AddDataRow(tableName As String, values() As Variant, Optional FromTop As Boolean)
    If Not IsMissing(FromTop) Then
        Low = range("NowPercent").Row
        lastRowTable = range("NowPercent").Cells(range("NowPercent").Cells.Count).Row
        For Each ColR In range("NowPercent").Columns
            lastRowNum = ColR.Cells(ColR.Cells.Count).End(xlUp).Row
            If lastRowNum > Low And lastRowTable <> lastRowNum Then
                Low = ColR.Cells(ColR.Cells.Count).End(xlUp).Row
            End If
        Next
        If FromTop = True Then
            Set RangeWithVal = range("NowPercent").Resize(Low - range("NowPercent").Row + 1, range("NowPercent").Columns.Count)
            Set NewRangeWithVal = RangeWithVal.Cells(1).offset(1).Resize(RangeWithVal.Rows.Count, RangeWithVal.Columns.Count)
            NewRangeWithVal.Value2 = RangeWithVal.Value2
            range("NowPercent").Resize(1, range("NowPercent").Columns.Count) = values
            Exit Sub
            
        Else

            
'            lastRowTable = Range("NowPercent").Cells(Range("NowPercent").Cells.Count).Row
'            If low = lastRowTable Then
'                Range("NowPercent").Resize(1, Range("NowPercent").Columns.Count) = values
'            Else
                range("NowPercent").offset(Low - range("NowPercent").Row + 1).Resize(1, range("NowPercent").Columns.Count) = values
''            End If
            
            Exit Sub
        End If
    End If

    Dim sheet As Worksheet
    Dim table As ListObject
    Dim tableRange As range
    Dim col As Integer
    Dim lastRow As range

    Set sheet = ActiveWorkbook.Worksheets("進度")
    Set table = sheet.ListObjects.Item(tableName)
    Set tableRange = table.range

    'First check if the last row is empty; if not, add a row
    If table.ListRows.Count > 0 Then
        Set lastRow = table.ListRows(table.ListRows.Count).range
        For col = 1 To lastRow.Columns.Count
            If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
                table.ListRows.Add
                Exit For
            End If
        Next col
    Else
        table.ListRows.Add
    End If

'    'Iterate through the last row and populate it with the entries from values()
'    Set lastRow = Table.ListRows(Table.ListRows.Count).Range
'    For Col = 1 To lastRow.Columns.Count
'        If Col <= UBound(values) + 1 Then
'            lastRow.Cells(1, Col).Value2 = values(Col - 1)
'        End If
'    Next Col



    Dim currentRowFirstCell As range
    Set currentRowFirstCell = range(tableName).Cells(1, 1)
    
    notblank = True
    Do While notblank
        notblank = False
'        For i = 1 To Range(tableName).Columns.Count
'            On Error Resume Next
'            If currentRowFirstCell.offset(0, i).Value <> vbNullString Then
'                notblank = True
'            End If
'        Next
        
        If WorksheetFunction.CountA(currentRowFirstCell.Resize(1, range(tableName).Columns.Count)) <> 0 Then
            notblank = True
        End If
        
        If notblank = False Then Exit Do
        Set currentRowFirstCell = currentRowFirstCell.offset(1)
    Loop
    
    currentRowFirstCell.Resize(1, tableRange.Columns.Count) = values
'    For Col = 1 To tableRange.Columns.Count
'        currentRowFirstCell.offset(0, Col - 1).Value2 = values(Col - 1)
'    Next
End Sub

Sub SortPercent()
Set sheet = ActiveWorkbook.Worksheets("進度")
Set mTable = sheet.ListObjects("NowPercent")



'Set sortcolumn = Range("NowPercent[Up MileStone]")
'    With mTable.Sort
'       .SortFields.Clear
'       .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
'       .Header = xlYes
'       .Apply
'    End With
'
'Set sortcolumn = Range("NowPercent[Time]")
'    With mTable.Sort
'       .SortFields.Clear
'       .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
'       .Header = xlYes
'       .Apply
'    End With

Set sortcolumn = range("NowPercent[Sort]")
    With mTable.Sort
       .SortFields.Clear
       .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
       .header = xlYes
       .Apply
    End With
    
Call ScaleChart8Axes
End Sub

Sub exportNowPercent()
    Call UpdateThisTask
    Call SavePreset
    Call saveTimeline
    Call ExportToFile
End Sub
Sub saveTimeline()
    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    current = Application.WorksheetFunction.Max(range("NowPercent[Actual]"))
    target = Application.WorksheetFunction.Max(range("NowPercent[Planned]"))
    endPercent = CDbl(range(use_Structured(currentRange, 8)).Value2)
    StartTime = range(use_Structured(currentRange, 4)).Value2
    currentTime = evals("=INDEX(NowPercent[Time],MATCH(MAX(NowPercent[Actual]),NowPercent[Actual],0))")
    range(use_Structured(currentRange, 30)).Calculate
    If target <> 0 Then
        range(use_Structured(currentRange, 12)).Value2 = current / target
        range(use_Structured(currentRange, 3)).Value2 = currentTime - StartTime
    End If

End Sub
Sub ExportToFile()
    Call ExportTable("NowPercent", taskFile)
    Call ExportRange(range("NowPercent"), experienceFile)
End Sub
Function IsArrayAllocated(arr As Variant) As Boolean
        On Error Resume Next
        IsArrayAllocated = IsArray(arr) And _
                           Not IsError(LBound(arr, 1)) And _
                           LBound(arr, 1) <= UBound(arr, 1)
End Function

Sub SetCurrentTask()
    Call exportNowPercent
    Call UpdateThisTask
    id = range("交易!S2")
    range("進度!$A$4").Value2 = id
    Call GenerateNewTable
End Sub

Public Function avg_Collection(collect As Collection) As Double
    Total = 0
    Count = 0
    For Each eaItem In collect
        If IsNumeric(eaItem) Then
            Total = Total + eaItem
            Count = Count + 1
        End If
    Next
    
    If Total * Count = 0 Then avg_Collection = -1 Else avg_Collection = Total / Count
End Function

Sub DeleteTableRows(ByVal tableName As String, KeepFormulas As Boolean)

Set sheet = ActiveWorkbook.Worksheets("進度")
Set table = sheet.ListObjects.Item(tableName)


On Error Resume Next

If Not KeepFormulas Then
    table.DataBodyRange.ClearContents
End If

table.DataBodyRange.Rows.Delete

On Error GoTo 0

End Sub

Sub EnterProgressMinimize()
    
    ActiveWindow.WindowState = xlMinimized
    EnterProgress.Show vbModeless
End Sub

Function CheckBehind() As Integer
    CheckBehind = Round(getPlannedByTime(CDbl(dateValue(Now()) + TimeValue(Now()))) - getCurrentActual())
End Function

Sub EnterProgressMinimizeIfBehind()
    If Round(getPlannedByTime(CDbl(dateValue(Now()) + TimeValue(Now()))) - getCurrentActual()) > 1 Then
        'ActiveWindow.WindowState = xlMinimized
        EnterProgress.Show vbModeless
    End If
End Sub


Sub ExportTable(tableName As String, filename As String)

    Dim WB As Workbook, wbNew As Workbook
    Dim ws As Worksheet, wsNew As Worksheet
    Dim wbNewName As String
    Dim r As range
        Set r = range(tableName)

   Set WB = ThisWorkbook
   Set ws = ActiveSheet

   Set wbNew = Workbooks.Add

   With wbNew
        
        Set wsNew = wbNew.Sheets(1)
        Set r = r.Resize(r.End(xlDown).Row - r.Row + 1)
       r.Copy
       wsNew.range("A1").PasteSpecial Paste:=xlPasteValues
       
        On Error Resume Next
       .SaveAs filename:=filename, _
             FileFormat:=xlCSVMSDOS, CreateBackup:=False
        On Error GoTo 0
   End With
    wbNew.Close savechanges:=False
End Sub


Sub ExportRange(r As range, filename As String)
    Call UpdateThisTask
    
    Dim WB As Workbook, wbNew As Workbook
    Dim ws As Worksheet, wsNew As Worksheet
    Dim wbNewName As String


   Set WB = ThisWorkbook
   Set ws = ActiveSheet

   Set wbNew = Workbooks.Add

   With wbNew
       Set wsNew = wbNew.Sheets(1)
       numRows = r.Rows.Count
        Set r = r.Resize(r.End(xlDown).Row - r.Row + 1)
       r.Copy
       wsNew.range("A1").PasteSpecial Paste:=xlPasteValues
       

        Dim arrX As Variant
        arrX = ArrayFromCSVfile(filename)
        
        If IsArrayAllocated(arrX) Then
            Dim pointer As range
            Set pointer = wsNew.range("A1")
            Set pointer = pointer.offset(r.Rows.Count)
            Array2Range arrX, pointer
        End If

        Application.DisplayAlerts = False
       .SaveAs filename:=filename, _
             FileFormat:=xlCSVMSDOS, CreateBackup:=False
        Application.DisplayAlerts = True
             
   End With
    wbNew.Close savechanges:=False
End Sub

Function ReplaceIllegalCharacters(strIn As String, strChar As String) As String
    Dim strSpecialChars As String
    Dim i As Long
    strSpecialChars = "~""#%&*:<>?{|}/\[]" & Chr(10) & Chr(13)

    For i = 1 To Len(strSpecialChars)
        strIn = Replace(strIn, Mid$(strSpecialChars, i, 1), strChar)
    Next

    ReplaceIllegalCharacters = strIn
End Function

Sub Array2Range(arr, destTL As range)
    'dumps [arr] (1D/2D) onto a sheet where [destTL] is the top-left output cell.
    destTL.Resize(UBound(arr, 1) - LBound(arr, 1) + 1, _
        UBound(arr, 2) - LBound(arr, 2) + 1) = arr
End Sub
















Sub UpdateCurrentPercentByInput()
    EnterProgress.Show vbModeless
    
'    Dim toAdd As Double
'    toAdd = InputBox("Enter your progress", "", Range("進度!$A$4"))
'    UpdateCurrentPercent (toAdd)
End Sub




Sub UpdateThisTask(Optional Fast As Boolean)
    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    If Fast <> True Then
        range(range("交易!$C$1").Value).Calculate
        range("交易!S2").Calculate
        range(use_Structured(currentRange, 4)).Calculate
        range(use_Structured(currentRange, 5)).Calculate
        range(use_Structured(currentRange, 7)).Calculate
        range(use_Structured(currentRange, 8)).Calculate
        range(use_Structured(currentRange, 30)).Calculate
        range(range("交易!$C$1").Value).Calculate
    End If
    range("進度!A4:H4").Calculate
End Sub

Function getNowTarget(Delta As Boolean) As Double
    Call UpdateThisTask(True)
    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    
    endPercent = range(use_Structured(currentRange, 8)).Value
    target = CDbl(range(use_Structured(currentRange, 15)).Value2)
    NowPercent = range(use_Structured(currentRange, 28)).Value2
    current = Application.WorksheetFunction.Max(range("NowPercent[Actual]"))
    
    If Delta = True Then
        getNowTarget = Round(NowPercent * (target / endPercent)) - current
    Else
        getNowTarget = Round(NowPercent * (target / endPercent))
    End If
    
    
End Function

Sub UpdateCurrentPercent(currTarget As Double, mode As Integer, Optional milestone As Variant, Optional NowTime As Variant)
    Call UpdateThisTask
    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    StartTime = Application.WorksheetFunction.Min(range("NowPercent[Time]")) '  Range(use_Structured(currentRange, 4)).Value2
    startPercent = range(use_Structured(currentRange, 7)).Value2
    endPercent = range(use_Structured(currentRange, 8)).Value2
    NowPercent = range(use_Structured(currentRange, 28)).Value2
    target = Application.WorksheetFunction.Max(range("NowPercent[Planned]")) ' CDbl(Range(use_Structured(currentRange, 15)).Value2)
    startTarget = Application.WorksheetFunction.Min(range("NowPercent[Actual]"))
    Dim current As range
    r = evals("AddressEx(INDEX(NowPercent[Actual],MATCH(MAX(NowPercent[Actual]),NowPercent[Actual],0)))")
    Set current = range(r) 'Application.WorksheetFunction.Max(Range("NowPercent[Actual]")))
    currentTime = Sheets("進度").Cells(current.Row, range("NowPercent[Time]").Column).Value2 'evals("=INDEX(NowPercent[Time],MATCH(MAX(NowPercent[Actual]),NowPercent[Actual],0))")
    
    If Not IsMissing(NowTime) Then
        If NowTime > Application.WorksheetFunction.Max(range("NowPercent[Time]")) Then
            projected = StartTime + (target - startTarget) * (NowTime - StartTime) / (currTarget - startTarget)
            range(evals("=AddressEx(INDEX(NowPercent[Time],MATCH(""End"",NowPercent[Milestone],0)))")).Value2 = projected
        End If
        
        
        For Each cell In range("NowPercent[Milestone]")
            If cell.Value2 <> vbNullString Then
                ff = cell.Value2
                PlannedTarget = Sheets("進度").Cells(cell.Row, range("NowPercent[Planned]").Column).Value2
                If PlannedTarget >= currTarget And NowTime > Sheets("進度").Cells(cell.Row, range("NowPercent[Time]").Column).Value2 Then
                    projected = StartTime + (Sheets("進度").Cells(cell.Row, range("NowPercent[Planned]").Column).Value2 - startTarget) * (NowTime - StartTime) / (currTarget - startTarget)
                    If projected >= range(evals("=AddressEx(INDEX(NowPercent[Time],MATCH(""End"",NowPercent[Milestone],0)))")).Value2 Then
                        projected = NowTime + (1 / (60 * 24))
                    End If
                    Sheets("進度").Cells(cell.Row, range("NowPercent[Time]").Column).Value2 = projected 'NowTime + (1 / (60 * 24))
                End If
            End If
        Next
    End If
    EndTime = Application.WorksheetFunction.Max(range("NowPercent[Time]")) 'Range(use_Structured(currentRange, 4)).Value2 + Range(use_Structured(currentRange, 2)).Value2


    
 
 
    If IsMissing(NowTime) Then NowTime = Now()
    Dim x(6)
    If mode = 1 Then 'Add new real entry
    
        x(0) = NowTime
        x(1) = current + (NowTime - currentTime) * ((target - current) / (EndTime - currentTime))
        g = (NowTime - currentTime) * ((target - current) / (EndTime - currentTime))
        If x(1) >= target Then
            x(1) = target
        End If
        x(2) = currTarget
        x(3) = x(2) - x(1)
        x(4) = milestone
        x(5) = ""  'X(1) / (Target / endPercent) 'NowPercent
        x(6) = ""
        AddDataRow "NowPercent", x

    ElseIf mode = 2 Then 'Add milestone by time
    
        x(0) = NowTime
        x(1) = current + (NowTime - currentTime) * ((target - current) / (EndTime - currentTime))
        If x(1) >= target Then
            x(1) = target
        End If
        x(2) = vbNullString
        x(3) = vbNullString
        x(4) = milestone
        x(5) = "" 'X(1) / (Target / endPercent) 'NowPercent
        x(6) = ""
        AddDataRow "NowPercent", x
    
    ElseIf mode = 3 Then 'Add milestone by target
        prevTime = evals("INDEX(NowPercent[Time],MATCH(" & currTarget & ",NowPercent[Planned],1))")
        nextTime = evals("INDEX(NowPercent[Time],MATCH(" & currTarget & ",NowPercent[Planned],1)+1)")
        prevTarget = evals("INDEX(NowPercent[Planned],MATCH(" & currTarget & ",NowPercent[Planned],1))")
        nextTarget = evals("INDEX(NowPercent[Planned],MATCH(" & currTarget & ",NowPercent[Planned],1)+1)")
        x(0) = prevTime + (currTarget - prevTarget) * ((nextTime - prevTime) / (nextTarget - prevTarget))
        x(1) = currTarget
        x(2) = vbNullString
        x(3) = vbNullString
        x(4) = milestone
        x(5) = x(1) / (target / endPercent) 'NowPercent
        x(6) = ""
        AddDataRow "NowPercent", x
    
    
    ElseIf mode = 4 Then
    
    Else
    End If
    
    Call SortPercent
    

    
    If current >= target Then  'Or evals("MAX(NowPercent[Time])") < Now()
        Call SetCurrentTask
        Exit Sub
    End If

    Call SortPercent
    'Worksheets("進度").Calculate
End Sub





Sub CleanPercentTable()

    Call UpdateThisTask
    Dim currentRange As range
    Set currentRange = range(evals("AddressEx(INDEX(表格2[編號],MATCH(進度!$A$4,表格2[ID],0)))"))
    StartTime = range(use_Structured(currentRange, 4)).Value2
    EndTime = range(use_Structured(currentRange, 4)).Value2 + range(use_Structured(currentRange, 2)).Value2
    startPercent = range(use_Structured(currentRange, 7)).Value2
    endPercent = range(use_Structured(currentRange, 8)).Value2
    NowPercent = range(use_Structured(currentRange, 28)).Value2
    target = range(use_Structured(currentRange, 15)).Value2
    id = range(use_Structured(currentRange, 10)).Value2

    range("NowPercent").Value = ""
    
    'Restore if exists
    Dim arrX As Variant
    arrX = ArrayFromCSVfile(taskFile)
    If IsArrayAllocated(arrX) Then Call Array2Range(arrX, range("NowPercent"))
    
    Dim x(7)
    x(0) = StartTime
    x(1) = startPercent * target / endPercent
    x(2) = startPercent * target / endPercent
    x(3) = ""
    x(4) = "Start"
    x(5) = ""  'startPercent
    x(6) = 1
    x(7) = 0
     Dim y(7)
    y(0) = EndTime
    y(1) = target
    y(2) = "" 'target * endPercent '"=MAX([Actual])"
    y(3) = ""
    y(4) = "End"
    y(5) = ""  'endPercent
    y(6) = 1
    y(7) = ""
    AddDataRow "NowPercent", x
    AddDataRow "NowPercent", y
    Call AddMilestones(StartTime, EndTime, startPercent, endPercent, target)
    Call SortPercent
    Call UpdateThisTask
    
End Sub
Sub AddMilestones(StartTime As Variant, EndTime As Variant, startPercent As Variant, endPercent As Variant, target As Variant)
    Dim arrX As Variant
    fd = experienceFile()
    arrX = ArrayFromCSVfile(experienceFile())
    If IsEmpty(arrX) Then Exit Sub
    
    Dim milestone As dictionary
    Set milestone = CreateObject("Scripting.Dictionary")
     
    Dim tmpcollect As Collection
    Dim timepercent As Collection
    Dim targetPercent As Collection
    Dim leftSlope As Collection
    Dim rightSlope As Collection
    
    For rownum = UBound(arrX, 1) To 0 Step -1
        If arrX(rownum, 4) <> "Start" And _
            arrX(rownum, 4) <> "End" And _
            arrX(rownum, 4) <> vbNullString And _
            Not milestone.Exists(arrX(rownum, 4)) Then
                Set tmpcollect = New Collection
                Set timepercent = New Collection
                Set targetPercent = New Collection
                Set leftSlope = New Collection
                Set rightSlope = New Collection
                tmpcollect.Add timepercent '(1)
                tmpcollect.Add targetPercent '(2)
                tmpcollect.Add leftSlope '(3)
                tmpcollect.Add rightSlope '(4)
                milestone.Add arrX(rownum, 4), tmpcollect
        End If
    Next
    
    
    thisStart = 0
    thisStarttarget = 0
    thisStartPercent = 0
    thisEnd = 0
    thisEndtarget = 0
    thisEndPercent = 0
    thisPercent = 0
    For rownum = UBound(arrX, 1) To 0 Step -1
        If arrX(rownum, 4) = "End" Then
            thisEnd = arrX(rownum, 0)
            thisEndtarget = arrX(rownum, 2)
            thisEndPercent = arrX(rownum, 5) 'arrX(rownum, 2) / (arrX(rownum, 1) / arrX(rownum, 5))
            For startPos = rownum To 0 Step -1
                If arrX(startPos, 4) = "Start" Then
                    thisStarttarget = arrX(startPos, 1)
                    thisStart = arrX(startPos, 0)
                    thisStartPercent = arrX(startPos, 5)
                    Exit For
                End If
                If thisEndtarget = vbNullString And arrX(startPos, 2) <> vbNullString Then
                    thisEnd = arrX(startPos, 0)
                    thisEndtarget = arrX(startPos, 2)
                    thisEndPercent = thisEndtarget / (arrX(startPos, 1) / arrX(startPos, 5))
                    rownum = startPos
                End If
            Next
        ElseIf arrX(rownum, 4) = "Start" Or _
                arrX(rownum, 2) = vbNullString Or _
                arrX(rownum, 4) = vbNullString Then
        Else
            Set tmpcollect = milestone(arrX(rownum, 4))
            tmpcollect(1).Add (arrX(rownum, 0) - thisStart) / (thisEnd - thisStart)
            If arrX(rownum, 2) <> vbNullString Then
                thisPercent = arrX(rownum, 2) / (thisEndtarget / thisEndPercent)
                tmpcollect(2).Add thisPercent '(arrX(rownum, 2) - thisStarttarget) / (thisEndtarget - thisStarttarget)
            Else
                thisPercent = thisStartPercent + (thisEndPercent - thisStartPercent) * (arrX(rownum, 0) - thisStart) / (thisEnd - thisStart)
                tmpcollect(2).Add thisPercent
            End If
            tmpcollect(3).Add (thisPercent - thisStartPercent) / (arrX(rownum, 0) - thisStart)
            tmpcollect(4).Add (thisEndPercent - thisPercent) / (thisEnd - arrX(rownum, 0))
            
            
        End If
    Next
    
    For Each key In milestone.Keys
        Set tmpcollect = milestone(key)
        milestonePosition = avg_Collection(tmpcollect(2))
        If milestonePosition <> -1 Then
            Dim y(6)
            'y(0) = startTime + avg_Collection(tmpcollect(1)) * (endTime - startTime)
            y(0) = calTimeBySlope(StartTime, EndTime, startPercent, endPercent, avg_Collection(tmpcollect(2)), avg_Collection(tmpcollect(3)), avg_Collection(tmpcollect(4)))
            y(1) = (avg_Collection(tmpcollect(2)) * target / endPercent)
            y(2) = ""
            y(3) = ""
            y(4) = key
            y(5) = ""
            y(6) = ""
            AddDataRow "NowPercent", y
        End If
    Next
    
End Sub

Function calTimeBySlope(StartTime As Variant, EndTime As Variant, startPercent As Variant, endPercent As Variant, targetPercent As Variant, leftSlope As Variant, rightSlope As Variant) As Variant
    slopeScaleFactor = -1
    'k(S1+S2)/2=Sm
    linearSlope = (endPercent - startPercent) / (EndTime - StartTime)
    slopeScaleFactor = linearSlope * 2 / (leftSlope + rightSlope)
    newRightSlope = slopeScaleFactor * rightSlope
    newLeftSlope = slopeScaleFactor * leftSlope
    ewewqew = (newLeftSlope + newRightSlope) / 2
    
    calTimeBySlopeL = StartTime + (targetPercent - startPercent) / newLeftSlope
    calTimeBySlopeR = EndTime - (endPercent - targetPercent) / newRightSlope
    If calTimeBySlopeL >= StartTime And calTimeBySlopeL <= EndTime Then
        If calTimeBySlopeR >= StartTime And calTimeBySlopeR <= EndTime Then
            calTimeBySlope = (calTimeBySlopeL + calTimeBySlopeR) / 2
        Else
            calTimeBySlope = calTimeBySlopeL
        End If
    Else
        calTimeBySlope = calTimeBySlopeR
    End If
    
End Function

Function Combine(A As Variant, B As Variant, Optional stacked As Boolean = True) As Variant
    'assumes that A and B are 2-dimensional variant arrays
    'if stacked is true then A is placed on top of B
    'in this case the number of rows must be the same,
    'otherwise they are placed side by side A|B
    'in which case the number of columns are the same
    'LBound can be anything but is assumed to be
    'the same for A and B (in both dimensions)
    'False is returned if a clash

    Dim lb As Long, m_A As Long, n_A As Long
    Dim m_B As Long, n_B As Long
    Dim m As Long, n As Long
    Dim i As Long, j As Long, k As Long
    Dim C As Variant

    If TypeName(A) = "Range" Then A = A.Value
    If TypeName(B) = "Range" Then B = B.Value

    lb = LBound(A, 1)
    m_A = UBound(A, 1)
    n_A = UBound(A, 2)
    m_B = UBound(B, 1)
    n_B = UBound(B, 2)

    If stacked Then
        m = m_A + m_B + 1 - lb
        n = n_A
        If n_B <> n Then
            Combine = False
            Exit Function
        End If
    Else
        m = m_A
        If m_B <> m Then
            Combine = False
            Exit Function
        End If
        n = n_A + n_B + 1 - lb
    End If
    ReDim C(lb To m, lb To n)
    For i = lb To m
        For j = lb To n
            If stacked Then
                If i <= m_A Then
                    C(i, j) = A(i, j)
                Else
                    C(i, j) = B(lb + i - m_A - 1, j)
                End If
            Else
                If j <= n_A Then
                    C(i, j) = A(i, j)
                Else
                    C(i, j) = B(i, lb + j - n_A - 1)
                End If
            End If
        Next j
    Next i
    Combine = C
End Function


