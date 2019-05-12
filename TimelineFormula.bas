Attribute VB_Name = "TimelineFormula"
Function getStartDateIgnoreBlank(StartDateR As Variant, EndDateR As Variant, RealTimeR As Variant)
    EndDateArr = EndDateR.Value2
    RealTimeArr = RealTimeR.Value2
    SizeofCol = UBound(EndDateArr) - LBound(EndDateArr) + 1
    prevEndDate = EndDateArr(1, 1)
    'ReDim StartDateArr(1 To SizeofCol, 1 To 1)
    StartDateArr = StartDateR.Value2
    For i = 1 To SizeofCol
        If StartDateR(i).HasFormula Then
            StartDateArr(i, 1) = prevEndDate
        End If
        If RealTimeArr(i, 1) > 0 Then
            prevEndDate = StartDateArr(i, 1) + RealTimeArr(i, 1)
        End If
    Next
    getStartDateIgnoreBlank = StartDateArr
End Function

Function getStartPercent(TaskChain As Variant, TaskIndex As Variant)
    Dim TaskChainCol As Collection
    On Error GoTo Error
    Set TaskChainCol = Str2Collection(TaskChain.Value2, ",")
    
    Dim mostRecentPercent As Double
    mostRecentPercent = 0
    For Each task In TaskChainCol
        With Application.WorksheetFunction
            On Error GoTo Error
            Set tmp = .index(range("表格2[ID]"), .Match(CDbl(task), range("表格2[ID]"), 0))
        End With
        
        If use_Structured2R(tmp, "表格2", "編號").Value2 < _
            TaskIndex And _
            use_Structuroed2R(tmp, "表格2", "實際百分比").Value2 > _
            mostRecentPercent Then
            
            mostRecentPercent = use_Structured2R(tmp, "表格2", "實際百分比").Value2
        End If
    Next
    
    getStartPercent = BiggerThanOneSetZero(mostRecentPercent)
    Exit Function
    
Error:
    getStartPercent = 0
End Function

