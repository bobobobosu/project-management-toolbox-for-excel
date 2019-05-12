Attribute VB_Name = "ChartMacros"
Sub UpdateChart7()
    Worksheets("趨勢").ChartObjects("Chart 7").Activate
    ActiveChart.PlotVisibleOnly = True
    ActiveChart.PlotVisibleOnly = False
    Application.ScreenUpdating = True
    With Worksheets("趨勢").ChartObjects("Chart 7").Chart
        For index = 1 To .SeriesCollection.Count
            For Each pt In .SeriesCollection(index).Points
                pt.HasDataLabel = True
                pt.DataLabel.ShowValue = False
                pt.DataLabel.ShowRange = True
            Next
        Next
    End With
    Set rsfgdhnj = Worksheets("趨勢").ChartObjects("Chart 7").Chart.Axes(xlValue, xlPrimary)
    
    'StartTime = Application.WorksheetFunction.RoundDown(Now())
    'EndTime = Application.WorksheetFunction.RoundUp(Now())
    With Worksheets("趨勢").ChartObjects("Chart 7").Chart.Axes(xlValue, xlPrimary)
        .MinimumScale = Int(Now())
        .MaximumScale = Int(Now()) + 1
    End With
End Sub
Sub UpdateChart8()
    Application.ScreenUpdating = False
    Worksheets("進度").ChartObjects("Chart 8").Activate
    ActiveChart.PlotVisibleOnly = True
    ActiveChart.PlotVisibleOnly = False
    Application.ScreenUpdating = True
    With Worksheets("進度").ChartObjects("Chart 8").Chart
        For index = 1 To .SeriesCollection.Count
            For Each pt In .SeriesCollection(index).Points
                pt.HasDataLabel = False
                pt.DataLabel.ShowValue = False
                pt.DataLabel.ShowRange = False
            Next
        Next
    End With
    
    With Worksheets("進度").ChartObjects("Chart 8").Chart
        For index = 1 To .SeriesCollection.Count
            For Each pt In .SeriesCollection(index).Points
                pt.HasDataLabel = True
                pt.DataLabel.ShowValue = False
                pt.DataLabel.ShowRange = True
            Next
        Next
    End With
    With Worksheets("進度").ChartObjects("Chart 8").Chart.Axes(xlCategory, xlPrimary)
        .MinimumScale = .MinimumScale
        .MaximumScale = .MaximumScale
    End With
End Sub
Sub DynamicChartScale()
    With Application.WorksheetFunction
        Set currentR = use_Structured2R(getCurretnActualR, "NowPercent", "Milestone")
        currentVal = getCurretnActualR.Value2
        
        Set prevR = currentR
        Set nextR = currentR
        
        Dim vals As range
        Set vals = use_Structured2R(currentR, "NowPercent", "Actual")
        Set vals = Union(vals, use_Structured2R(vals, "NowPercent", "Planned"))
        
        prevcnt = 3
        nextcnt = 3
        
        Do While prevcnt > 0 And prevR.Row > range("NowPercent").Row
            Set prevR = prevR.offset(-1)
            g = prevR.Value2
            If prevR.Value2 <> vbNullString Then
                prevcnt = prevcnt - 1
            End If
            Set vals = Union(vals, use_Structured2R(prevR, "NowPercent", "Actual"))
            Set vals = Union(vals, use_Structured2R(prevR, "NowPercent", "Planned"))
        Loop
        Do While nextcnt > 0 And nextR.Row < range("NowPercent").Row + range("NowPercent").Rows.Count - 1
            Set nextR = nextR.offset(1)
            If nextR.Value2 <> vbNullString And use_Structured2R(nextR, "NowPercent", "Planned") > currentVal Then
                nextcnt = nextcnt - 1
            End If
            Set vals = Union(vals, use_Structured2R(nextR, "NowPercent", "Actual"))
            Set vals = Union(vals, use_Structured2R(nextR, "NowPercent", "Planned"))
        Loop
    End With
    With Worksheets("進度").ChartObjects("Chart 8").Chart.Axes(xlCategory, xlPrimary)
        .MinimumScale = Application.WorksheetFunction.Min(samerowsOf(vals, range("NowPercent[Time]"))) ' use_Structured2R(prevR, "NowPercent", "Time")
        .MaximumScale = Application.WorksheetFunction.Max(samerowsOf(vals, range("NowPercent[Time]"))) ' use_Structured2R(nextR, "NowPercent", "Time")
    End With
    With Worksheets("進度").ChartObjects("Chart 8").Chart.Axes(xlValue, xlPrimary)
        .MinimumScale = Application.WorksheetFunction.Min(vals)
        .MaximumScale = Application.WorksheetFunction.Max(vals)
    End With
    
    
End Sub

Sub ScaleChart8Axes()
Dim sel As range
On Error Resume Next
Set sel = Selection
If sel Is Nothing Then
    Call ScaleChart8Auto
    Application.OnTime Now, "UpdateChart8"
    Exit Sub
End If
If sel.Cells.Count = 1 Then
    Call DynamicChartScale
    Application.OnTime Now, "UpdateChart8"
    Exit Sub
End If
With Worksheets("進度").ChartObjects("Chart 8").Chart.Axes(xlCategory, xlPrimary)
.MinimumScale = Application.WorksheetFunction.Min(sel)
.MaximumScale = Application.WorksheetFunction.Max(sel)
End With


If Intersect(Selection, range("NowPercent")).Cells.Count > 0 And Selection.Columns.Count = 1 Then
    With Worksheets("進度").ChartObjects("Chart 8").Chart.Axes(xlValue, xlPrimary)
    .MinimumScale = Application.WorksheetFunction.Min(use_Structured2R(sel.Cells(0), "NowPercent", "Planned").Resize(Selection.Cells.Count, 2))
    .MaximumScale = Application.WorksheetFunction.Max(use_Structured2R(sel.Cells(0), "NowPercent", "Planned").Resize(Selection.Cells.Count, 2))
    '.MajorUnit = Worksheets("進度").Range("G4").Value
    End With
End If
Application.OnTime Now, "UpdateChart8"
End Sub

Sub ScaleChart8Auto()
'    With Worksheets("進度").ChartObjects("Chart 8").Chart.Axes(xlValue)
'    .MinimumScaleIsAuto = True
'    .MaximumScaleIsAuto = True
'    End With
    With Worksheets("進度").ChartObjects("Chart 8").Chart.Axes(xlCategory, xlPrimary)
        .MinimumScale = Application.WorksheetFunction.Min(range("NowPercent[Time]"))
        .MaximumScale = Application.WorksheetFunction.Max(range("NowPercent[Time]"))
    End With
    With Worksheets("進度").ChartObjects("Chart 8").Chart.Axes(xlValue, xlPrimary)
        .MinimumScale = Application.WorksheetFunction.Min(Union(range("NowPercent[Actual]"), range("NowPercent[Planned]")))
        .MaximumScale = Application.WorksheetFunction.Max(Union(range("NowPercent[Actual]"), range("NowPercent[Planned]")))
    '.MajorUnit = Worksheets("進度").Range("G4").Value
    End With
End Sub
