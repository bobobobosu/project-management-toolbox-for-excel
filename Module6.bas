Attribute VB_Name = "Module6"
Function LinearInterpolate(TimeSeries As Variant, DataSeries As Variant, targetTime As Double) As Double
    Dim TimeLower As Range
    Dim DataLower As Range
    Dim TimeUpper As Range
    Dim DataUpper As Range

    initIndex = 1
    On Error Resume Next
    initIndex = Application.WorksheetFunction.Match(targetTime, TimeSeries)


    
    For j = initIndex - 1 To 1 Step -1
        If DataSeries.Cells(j).Value2 <> vbNullString Then
            Set TimeLower = TimeSeries.Cells(j)
            Set DataLower = DataSeries.Cells(j)
            Exit For
        End If
    Next j
    
    For k = initIndex To TimeSeries.Cells.count
        If DataSeries.Cells(k).Value2 <> vbNullString Then
            Set TimeUpper = TimeSeries.Cells(k)
            Set DataUpper = DataSeries.Cells(k)
            Exit For
        End If
    Next k
    
    If Not (TimeLower Is Nothing) And Not (TimeUpper Is Nothing) Then
        x0 = TimeLower.Value2
        y0 = DataLower.Value2
        x1 = TimeUpper.Value2
        y1 = DataUpper.Value2
        x = targetTime
        LinearInterpolate = y0 + (x - x0) * ((y1 - y0) / (x1 - x0))
    ElseIf Not (TimeLower Is Nothing) Then
        LinearInterpolate = DataLower.Value2
    ElseIf Not (TimeUpper Is Nothing) Then
        LinearInterpolate = DataUpper.Value2
    Else
        LinearInterpolate = 0
    End If
    
End Function
Sub testInter()
    MsgBox LinearInterpolate(Range("表格82[a]"), Range("表格82[b]"), 2)
End Sub
'
'Function LinearInterpolate2(TimeSeries As Variant, DataSeries As Variant, targetTime As Double) As Double
'
'    Dim MAX_smaller As Double
'    Dim MIN_bigger As Double
'    Dim t_MAX_smaller As Double
'    Dim t_MIN_bigger As Double
'    Dim resultData As Double
'
'    Dim TimeSeries_arr As Variant
'    TimeSeries_arr = TimeSeries.Value
'    Dim DataSeries_arr As Variant
'    DataSeries_arr = DataSeries.Value
'
'    t_MAX_smaller = TimeSeries_arr(LBound(TimeSeries_arr, 1), 1)
'    t_MIN_bigger = TimeSeries_arr(UBound(TimeSeries_arr, 1), 1)
'    'MAX_smaller = DataSeries_arr(LBound(TimeSeries_arr, 1), 1)
'    'MIN_bigger = DataSeries_arr(UBound(TimeSeries_arr, 1), 1)
'    MAX_smaller = -1
'    MIN_bigger = -1
'
'
'
'
'        For i = 1 To UBound(TimeSeries_arr, 1)
'            'Debug.Print DataSeries_arr(i, 1)
'
'            If DataSeries(i, 1) <> "" Then
'
'                If TimeSeries_arr(i, 1) = targetTime Then
'                    resultData = DataSeries_arr(i, 1)
'                    LinearInterpolate = resultData
'                    Exit Function
'                End If
'                If TimeSeries_arr(i, 1) < targetTime Then
'                    If TimeSeries_arr(i, 1) > t_MAX_smaller Then
'                        t_MAX_smaller = TimeSeries_arr(i, 1)
'                        MAX_smaller = DataSeries(i, 1)
'                    End If
'                End If
'                If TimeSeries_arr(i, 1) > targetTime Then
'                    If TimeSeries_arr(i, 1) < t_MIN_bigger Then
'                        t_MIN_bigger = TimeSeries_arr(i, 1)
'                        MIN_bigger = DataSeries(i, 1)
'                    End If
'                End If
'            End If
'
'
'
'        Next i
'
'
'        If MAX_smaller > 0 And MIN_bigger > 0 Then
'            'Debug.Print "oooooooooooooo"
'            'Debug.Print t_MAX_smaller
'            'Debug.Print t_MIN_bigger
'            'Debug.Print MAX_smaller
'            'Debug.Print MIN_bigger
'            'Debug.Print targetTime
'            LinearInterpolate = ((Abs(t_MAX_smaller - targetTime) * MIN_bigger + Abs(t_MIN_bigger - targetTime) * MAX_smaller)) / ((Abs(t_MAX_smaller - targetTime) + Abs(t_MIN_bigger - targetTime)))
'            'Debug.Print LinearInterpolate
'        ElseIf MAX_smaller <= 0 And MIN_bigger <= 0 Then
'            LinearInterpolate = 0
'        Else
'            'Debug.Print "oooooooooooooo"
'            'Debug.Print t_MAX_smaller
'            'Debug.Print t_MIN_bigger
'            'Debug.Print MAX_smaller
'            'Debug.Print MIN_bigger
'            LinearInterpolate = MAX_smaller + MIN_bigger + 1
'        End If
'
'End Function

'Function LinearInterpolate(rXv As Variant, rYv As Variant, x As Double) As Double
'    Dim rX As Range
'    Dim rY As Range
'    Set rX = rXv
'    Set rY = rYv
'
'     ' linear interpolator / extrapolator
'     ' R is a two-column range containing known x, known y
'    Dim lR As Long, l1 As Long, l2 As Long
'    Dim nR As Long
'     'If x = 1.5 Then Stop
'
'    nR = rX.Rows.count
'    If nR < 2 Then Exit Function
'
'    If x < rX(1) Then ' x < xmin, extrapolate
'        l1 = 1: l2 = 2: GoTo Interp
'
'    ElseIf x > rX(nR) Then ' x > xmax, extrapolate
'        l1 = nR - 1: l2 = nR: GoTo Interp
'
'    Else
'         ' a binary search would be better here
'        For lR = 1 To nR
'            If rX(lR) = x Then ' x is exact from table
'                LinearInterpolate = rY(lR)
'                Exit Function
'
'            ElseIf rX(lR) > x Then ' x is between tabulated values, interpolate
'                l1 = lR: l2 = lR - 1: GoTo Interp
'
'            End If
'        Next
'    End If
'
'Interp:
'    LinearInterpolate = rY(l1) _
'    + (rY(l2) - rY(l1)) _
'    * (x - rX(l1)) _
'    / (rX(l2) - rX(l1))
'
'End Function


