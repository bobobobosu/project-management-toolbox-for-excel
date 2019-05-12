Attribute VB_Name = "Module7"



Function GetSUMIN(target As Variant, fromTime As Double, toTime As Double, TimeSeries As Variant, DataSeries As Variant, CycleMode As Integer, Limit As Integer) As Double
    GetSUMIN = GetSUMIN2(target, fromTime, toTime, 1, TimeSeries, DataSeries, CycleMode, Limit)
End Function



Function GetSUMIN2(target As Variant, fromTime As Double, toTime As Double, PercentOfTime As Double, TimeSeries As Variant, DataSeries As Variant, CycleMode As Integer, Limit As Integer) As Double
    If TimeSeries.Cells(TimeSeries.Cells.Count).Value2 < fromTime Or TimeSeries.Cells(1).Value2 > toTime Then
        GetSUMIN2 = (toTime - fromTime) * 60 * 24 * 100
        Exit Function
    End If
    'CycleMode Definition
    '0: Not Recurring, Limited SU-MIN
    '1: Recurring, Limited SU-MIN, One Time per Day
    '2: Recurring, Limited SU-MIN, One Time per Year
    
    
    Dim targetLimit As Double
    Debug.Print CycleMode
    If Limit = 1 Then
        targetLimit = readAVGSUMIN(target)
        Debug.Print targetLimit
        targetLimit = targetLimit
        targetLimit = targetLimit / Evaluate("=價值表!$A$4")
        targetLimit = targetLimit / (60 * 24)
        Debug.Print targetLimit
    ElseIf Limit = 0 Then
        targetLimit = 10000000000000#
    End If
    Dim TimeSeries_arr As Variant
    TimeSeries_arr = TimeSeries.Value
    Dim DataSeries_arr As Variant
    DataSeries_arr = DataSeries.Value
    Dim Duration As Double
    Duration = (toTime - fromTime) * 24 * 60
    
    If CycleMode = 0 Then
        gqerfg = LinearInterpolate(TimeSeries, DataSeries, toTime)
        kkk = LinearInterpolate(TimeSeries, DataSeries, fromTime)
        GetSUMIN2 = Smaller(targetLimit, PercentOfTime * (LinearInterpolate(TimeSeries, DataSeries, toTime) - LinearInterpolate(TimeSeries, DataSeries, fromTime)))
    ElseIf CycleMode = 1 Then
        '日循環
        If Int(toTime) - Int(fromTime) >= 1 Then
            Dim startPart As Double
            startPart = Smaller(10000000000000#, DataSeries_arr(UBound(DataSeries_arr), 1) - LinearInterpolate(TimeSeries, DataSeries, fromTime - Int(fromTime)))
            Dim endPart As Double
            endPart = Smaller(10000000000000#, LinearInterpolate(TimeSeries, DataSeries, toTime - Int(toTime)) - DataSeries_arr(LBound(DataSeries_arr), 1))
            Dim midPart As Double
            midPart = Smaller(10000000000000#, (DataSeries_arr(UBound(DataSeries_arr), 1) - DataSeries_arr(LBound(DataSeries_arr), 1))) * (Int(toTime) - (Int(fromTime)) - 1)
        
            
            
            GetSUMIN2 = Smaller(targetLimit, PercentOfTime * (startPart + midPart + endPart))
            'GetSUMIN2 = (startPart + midPart + endPart)
            'GetSUMIN2 = 24 * 60 * targetLimit * (toTime - fromTimev
            
            'Debug.Print "ppppp"
            
            'Debug.Print startPart
            'Debug.Print endPart
            'Debug.Print midPart
            'Debug.Print GetSUMIN2
        Else
            Debug.Print "ggggg"
            Debug.Print toTime - Int(toTime)
            Debug.Print fromTime - Int(fromTime)
            Debug.Print "ggggyg"
            Debug.Print LinearInterpolate(TimeSeries, DataSeries, toTime - Int(toTime))
            Debug.Print LinearInterpolate(TimeSeries, DataSeries, fromTime - Int(fromTime))
            Debug.Print (LinearInterpolate(TimeSeries, DataSeries, toTime - Int(toTime)) - LinearInterpolate(TimeSeries, DataSeries, fromTime - Int(fromTime)))
            GetSUMIN2 = Smaller(targetLimit, PercentOfTime * (LinearInterpolate(TimeSeries, DataSeries, toTime - Int(toTime)) - LinearInterpolate(TimeSeries, DataSeries, fromTime - Int(fromTime))))
            'GetSUMIN2 = (LinearInterpolate(TimeSeries, DataSeries, toTime - Int(toTime)) - LinearInterpolate(TimeSeries, DataSeries, fromTime - Int(fromTime)))
        
        End If
    ElseIf CycleMode = 2 Then
        '年循環
        If Year(toTime) - Year(fromTime) >= 1 Then
            Dim startPart2 As Double
            'startPart2 = Smaller(targetLimit, DataSeries_arr(UBound(DataSeries_arr), 1) - LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(fromTime), Day(fromTime)) + fromTime - Int(fromTime)))
            startPart2 = DataSeries_arr(UBound(DataSeries_arr), 1) - LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(fromTime), Day(fromTime)) + fromTime - Int(fromTime))

            Dim endPart2 As Double
            'endPart2 = Smaller(targetLimit, LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(toTime), Day(toTime)) + toTime - Int(toTime)) - DataSeries_arr(LBound(DataSeries_arr), 1))
            endPart2 = LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(toTime), Day(toTime)) + toTime - Int(toTime)) - DataSeries_arr(LBound(DataSeries_arr), 1)
            Dim midPart2 As Double
            'midPart2 = Smaller(targetLimit, (DataSeries_arr(UBound(DataSeries_arr), 1) - DataSeries_arr(LBound(DataSeries_arr), 1))) * (Year(toTime) - Year(fromTime) - 1)
            midPart2 = (DataSeries_arr(UBound(DataSeries_arr), 1) - DataSeries_arr(LBound(DataSeries_arr), 1)) * (Year(toTime) - Year(fromTime) - 1)
            
            'Debug.Print "ppppp"
            'Debug.Print startPart
            'Debug.Print endPart
            
            GetSUMIN2 = Smaller(targetLimit, PercentOfTime * (startPart2 + midPart2 + endPart2))
        Else
            'Debug.Print "ggggg"
            'Debug.Print toTime - Int(toTime)
            'Debug.Print fromTime - Int(fromTime)
            GetSUMIN2 = Smaller(targetLimit, PercentOfTime * (LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(toTime), Day(toTime)) + toTime - Int(toTime)) - LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(fromTime), Day(fromTime)) + fromTime - Int(fromTime))))
            'GetSUMIN2 = (LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(toTime), Day(toTime)) + toTime - Int(toTime)) - LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(fromTime), Day(fromTime)) + fromTime - Int(fromTime)))
        End If
    
        
    End If
        

End Function




Function GetESTTIME(target As Variant, fromTime As Double, Percent As Double, TimeSeries As Variant, DataSeries As Variant, CycleMode As Integer) As Double
    On Error GoTo eh
    GetESTTIME = GetESTTIME2(target, fromTime, Percent, TimeSeries, DataSeries, CycleMode)
Done:
    Exit Function
eh:
    GetESTTIME = 0
    Exit Function
End Function



Function GetESTTIME2(target As Variant, fromTime As Double, Percent As Double, TimeSeries As Variant, DataSeries As Variant, CycleMode As Integer) As Double
    Dim TimeSeries_arr As Variant
    TimeSeries_arr = TimeSeries.Value
    Dim DataSeries_arr As Variant
    DataSeries_arr = DataSeries.Value
    Dim targetLimit As Double
    targetLimit = readAVGSUMIN(target) * Percent
    Dim currEndTime As Double
    currEndTime = fromTime
    
    Dim culmTime As Double
    culmTime = 0
     
    Dim StartTime As Double
        
    If CycleMode = 0 Then
        startData = LinearInterpolate(TimeSeries, DataSeries, fromTime)
        currEndTime = LinearInterpolate(DataSeries, TimeSeries, startData + targetLimit)
        GetESTTIME2 = currEndTime
        Exit Function
    ElseIf CycleMode = 1 Then
        '日循環
        StartTime = fromTime - Int(fromTime)
        startData = LinearInterpolate(TimeSeries, DataSeries, StartTime)
        Do While targetLimit > 0
            If targetLimit < DataSeries_arr(UBound(DataSeries_arr), 1) - startData Then
                currEndTime = currEndTime + (LinearInterpolate(DataSeries, TimeSeries, startData + targetLimit) - StartTime)
                targetLimit = 0
            Else
                currEndTime = currEndTime + TimeSeries_arr(UBound(TimeSeries_arr), 1) - StartTime
                'targetLimit = targetLimit - (DataSeries_arr(UBound(DataSeries_arr), 1) - LinearInterpolate(TimeSeries, DataSeries, startTime))
                OneDay = (DataSeries_arr(UBound(DataSeries_arr), 1) - LinearInterpolate(TimeSeries, DataSeries, StartTime))
                If OneDay <= 0 Then Exit Do
                targetLimit = targetLimit - OneDay
                StartTime = TimeSeries_arr(LBound(TimeSeries_arr), 1)
                startData = DataSeries_arr(LBound(DataSeries_arr), 1)
                
                Debug.Print targetLimit
            End If
        Loop
        GetESTTIME2 = currEndTime - fromTime
        Exit Function
    ElseIf CycleMode = 2 Then
        '年循環
        StartTime = Year(toTime) - Year(fromTime)
        startData = LinearInterpolate(TimeSeries, DataSeries, DateSerial(Year(TimeSeries(1, 1)), Month(fromTime), Day(fromTime)))
        Do While targetLimit > 0
            If targetLimit < DataSeries_arr(UBound(DataSeries_arr), 1) - startData Then
                currEndTime = currEndTime + (LinearInterpolate(DataSeries, TimeSeries, startData + targetLimit) - StartTime)
                targetLimit = 0
            Else
                currEndTime = currEndTime + TimeSeries_arr(UBound(TimeSeries_arr), 1) - StartTime
                OneYear = (DataSeries_arr(UBound(DataSeries_arr), 1) - LinearInterpolate(TimeSeries, DataSeries, StartTime))
                If OneYear <= 0 Then Exit Do
                targetLimit = targetLimit - OneYear
                StartTime = TimeSeries_arr(LBound(TimeSeries_arr), 1)
                startData = DataSeries_arr(LBound(DataSeries_arr), 1)
            End If
        Loop
        GetESTTIME2 = currEndTime - fromTime
        Exit Function
    End If
    
    
    
End Function



Function readAVGSUMIN(target As Variant) As Double
    Dim target_s As String
    target_s = target

    readAVGSUMIN = WorksheetFunction.IfError(Evaluate("=INDEX(表格55[SU-MIN],MATCH( """ & target_s & """,表格55[工作物件],0))"), 0)
End Function








Function Smaller(num1 As Double, num2 As Double)
    'Debug.Print "iiii"
    'Debug.Print num1
    'Debug.Print num2
    If num1 < 0 And num2 >= 0 Then
        Smaller = num2
    ElseIf num1 >= 0 And num2 < 0 Then
        Smaller = num1
    Else
    
    
        If num1 > num2 Then
            'Debug.Print num2
            Smaller = num2
        ElseIf num2 > num1 Then
            Smaller = num1
        Else
            Smaller = num2
        End If
    End If
    
End Function


Function bigger(num1 As Double, num2 As Double)
        If num1 > num2 Then
            'Debug.Print num2
            bigger = num1
        ElseIf num2 > num1 Then
            bigger = num2
        Else
            bigger = num1
        End If
End Function
Function BiggerThanOneSetZero(num As Double)
    If num < 0.999 Then
        BiggerThanOneSetZero = num
    Else
        BiggerThanOneSetZero = 0
    End If
End Function

Public Function EXPLODE_V(texte As String, delimiter As String)
    EXPLODE_V = Application.WorksheetFunction.Transpose(Split(texte, delimiter))
End Function
Public Function EXPLODE_H(texte As String, delimiter As String)
    EXPLODE_H = Split(texte, delimiter)
End Function
