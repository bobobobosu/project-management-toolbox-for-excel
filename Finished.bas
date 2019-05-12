Attribute VB_Name = "Finished"
Function SubstractFinished(target As Variant, fromTime As Variant, toTime As Variant, CycleMode As Variant) As Double

    If CycleMode = 0 Then
        SubstractFinished = 0
    Else
    
        Dim targetRange As range
        Dim StartDateR As range
        Dim EndDateR As range
        Dim SUpercentR As range
        Set targetRange = Evaluate("表格2[交易物件]")
        Set StartDateR = Evaluate("表格2[Start Date]")
        Set EndDateR = Evaluate("表格2[End Date]")
        Set SUpercentR = Evaluate("表格2[Projected SU%]")
        Dim targetArray As Variant
        Dim startDateA As Variant
        Dim endDateA As Variant
        Dim SUpercentA As Variant
        
        targetArray = targetRange.Value2
        startDateA = StartDateR.Value2
        endDateA = EndDateR.Value2
        SUpercentA = SUpercentR.Value2
        
        Dim accuIncome As Double
        accuIncome = 0
        
        
   
        For i = 1 To UBound(targetArray, 1)
        
            If targetArray(i, 1) = target And SUpercentA(i, 1) > 0 And startDateA(i, 1) > 0 And endDateA(i, 1) > 0 Then
                accuIncome = accuIncome + 100 * SUpercentA(i, 1) * getOverLap(startDateA(i, 1), endDateA(i, 1), fromTime, toTime) * 60 * 24
                SubstractFinished = accuIncome
                'Debug.Print SubstractFinished
            End If
            On Error Resume Next
        Next i
        SubstractFinished = accuIncome
        'Debug.Print SubstractFinished
    
    End If
End Function

Function TimeFinished(target As Variant, fromTime As Variant, toTime As Variant, CycleMode As Variant) As Double

    If CycleMode = -1 Then
        TimeFinished = 0
    Else
    
        Dim targetRange As range
        Dim StartDateR As range
        Dim EndDateR As range
        Dim SUpercentR As range
        Set targetRange = Evaluate("表格2[交易物件]")
        Set StartDateR = Evaluate("表格2[Start Date]")
        Set EndDateR = Evaluate("表格2[End Date]")
        Set SUpercentR = Evaluate("表格2[Projected SU%]")
        Dim targetArray As Variant
        Dim startDateA As Variant
        Dim endDateA As Variant
        Dim SUpercentA As Variant
        
        targetArray = targetRange.Value2
        startDateA = StartDateR.Value2
        endDateA = EndDateR.Value2
        SUpercentA = SUpercentR.Value2
        
        Dim accuIncome As Double
        accuIncome = 0
        
        If target <> "時序專案(288)" Then
            'For Others
            For i = 1 To UBound(targetArray, 1)
            
                'If targetArray(i, 1) = Target And SUpercentA(i, 1) > 0 And startDateA(i, 1) > 0 And endDateA(i, 1) > 0 Then
                If targetArray(i, 1) = target And startDateA(i, 1) > 0 And endDateA(i, 1) > 0 Then
                    accuIncome = accuIncome + getOverLap(startDateA(i, 1), endDateA(i, 1), fromTime, toTime) * 60 * 24
                    TimeFinished = accuIncome
                    'Debug.Print SubstractFinished
                End If
                On Error Resume Next
            Next i
            TimeFinished = accuIncome
            'Debug.Print SubstractFinished
            
            
        Else
            'For 時序專案
            Dim nonblank As Double
            nonblank = 0
            For i = 1 To UBound(targetArray, 1)
            
                'If targetArray(i, 1) = Target And SUpercentA(i, 1) > 0 And startDateA(i, 1) > 0 And endDateA(i, 1) > 0 Then
                If targetArray(i, 1) <> target And startDateA(i, 1) > 0 And endDateA(i, 1) > 0 Then
                    nonblank = nonblank + getOverLap(startDateA(i, 1), endDateA(i, 1), fromTime, toTime) * 60 * 24
                End If
                On Error Resume Next
            Next i
            TimeFinished = (toTime - fromTime) * 60 * 24 - nonblank
        End If
        
        
        
    
    End If
End Function


