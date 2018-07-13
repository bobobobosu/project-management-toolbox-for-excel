Attribute VB_Name = "Module12"
Function getOverLap(ByVal fromTime1 As Variant, ByVal toTime1 As Variant, ByVal fromTime2 As Variant, ByVal toTime2 As Variant) As Double

    getOverLap = WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2) - WorksheetFunction.Max(fromTime1, fromTime2))

    
    'getOverLap = Application.Evaluate("=MAX(0,MIN(toTime1,toTime2)-MAX(fromTime1,fromTime2))")
End Function


Function getOverLapSum(ByVal fromTime1 As Variant, ByVal toTime1 As Variant, fromTime2 As Variant, ByVal toTime2 As Variant) As Double
    
    'Dim fromTime2_rng As Range
    'fromTime2_rng = fromTime2
    
    'Dim toTime2_rng As Range
    'toTime2_rng = toTime2

    'fromTime2(1) = 50
    
    
    
    'Dim gg As String
    'gg = fromTime2(0)
    'MsgBox gg
    getOverLapSum = 0
'    For i = 1 To fromTime2.Rows.Count
'        'MsgBox "ii"
'        'MsgBox fromTime2.Cells(i).Value
'        getOverLapSum = getOverLapSum + WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2.Cells(i).Value) - WorksheetFunction.Max(fromTime1, fromTime2.Cells(i).Value))
'    Next i


    For i = 1 To UBound(fromTime2)
        'MsgBox "gg"
        getOverLapSum = getOverLapSum + WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2(i, 1)) - WorksheetFunction.Max(fromTime1, fromTime2(i, 1)))
    Next i
    
    
    'getOverLapSum = WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2) - WorksheetFunction.Max(fromTime1, fromTime2))
    
    
    'getOverLap = Application.Evaluate("=MAX(0,MIN(toTime1,toTime2)-MAX(fromTime1,fromTime2))")
End Function

Function TestArr(ByRef arr()) As String
    MsgBox arr(1, 1)
    
    TestArr = 0
End Function



Sub text()
Dim haha()
haha = Array("Tom", "Mary", "Adam")
TestArr haha()


End Sub

Function testing(ByRef check()) As String
Dim track As Long
For track = LBound(check) To UBound(check)
    check(track) = check(track) & " OMG"
Next
End Function
