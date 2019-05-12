Attribute VB_Name = "Chain"
Function getOverLap(ByVal fromTime1 As Variant, ByVal toTime1 As Variant, ByVal fromTime2 As Variant, ByVal toTime2 As Variant) As Double
    getOverLap = WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2) - WorksheetFunction.Max(fromTime1, fromTime2))
End Function


Function getOverLapSum(ByVal fromTime1 As Variant, ByVal toTime1 As Variant, fromTime2 As Variant, ByVal toTime2 As Variant) As Double
    getOverLapSum = 0
    For i = 1 To UBound(fromTime2)
        getOverLapSum = getOverLapSum + WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2(i, 1)) - WorksheetFunction.Max(fromTime1, fromTime2(i, 1)))
    Next i
End Function

Function getOverLapCellSum(ByVal fromTime1 As Variant, ByVal toTime1 As Variant, fromTime2 As Variant, ByVal toTime2 As Variant) As Double
    getOverLapCellSum = 0
    
    For i = 1 To fromTime2.Cells.Count
        getOverLapCellSum = getOverLapCellSum + WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2.Cells(i)) - WorksheetFunction.Max(fromTime1, fromTime2.Cells(i)))
    Next i
End Function


Function getOverLapCheck(ByVal fromTime1 As Variant, ByVal toTime1 As Variant, fromTime2 As Variant, ByVal toTime2 As Variant) As Double
    OverLapCellSum = 0
    
    For i = 1 To fromTime2.Cells.Count
        OverLapCellSum = OverLapCellSum + WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2.Cells(i)) - WorksheetFunction.Max(fromTime1, fromTime2.Cells(i)))
        If (OverLapCellSum - (toTime1 - fromTime1)) > 0 Then
            getOverLapCheck = 1
            Exit Function
        End If
    Next i
    
    getOverLapCheck = 0

End Function

Function getConcurrentTasks(ByVal fromTime1 As Variant, ByVal toTime1 As Variant, fromTime2 As Variant, ByVal toTime2 As Variant, id As Variant) As String
    concurrent = ""
    
    For i = 1 To fromTime2.Cells.Count
    
        If (WorksheetFunction.Max(0, WorksheetFunction.Min(toTime1, toTime2.Cells(i)) - WorksheetFunction.Max(fromTime1, fromTime2.Cells(i)))) > (1 / 24) / 60 Then
            If id.Cells(i).Value2 <> range(use_Structured(range(Application.Caller.address), 10)).Value2 Then
                concurrent = concurrent + CStr(id.Cells(i).Value2) + ","
            End If
        End If
    Next i
    
    getConcurrentTasks = concurrent

End Function

Function getTaskChain(cell As Variant) As String
    totalRows = range("表格2[編號]").Cells.Count
    Dim Chain As String
    Chain = CStr(range(use_Structured(cell, 10)).Value2) + ","
    startFound = False
    endFound = False

    On Error Resume Next
    If range(use_Structured(cell, 7)).Value2 = 0 Then startFound = True
    On Error Resume Next
    If range(use_Structured(cell, 12)).Value2 >= 0.999 Then endFound = True
    
    Count = cell.Row - range("表格2").Row
    Dim pointer As range
    Set pointer = cell
    Do While pointer.Row > range("表格2").Row And startFound = False
        Set pointer = pointer.offset(-1)
        checkTarget = range(use_Structured(pointer, 6)).text = range(use_Structured(cell, 6)).text
        checkDescription = range(use_Structured(pointer, 9)).text = range(use_Structured(cell, 9)).text
        checkPercent = range(use_Structured(pointer, 12)).Value2 <= range(use_Structured(cell, 7)).Value2
        If checkTarget And checkDescription And Not checkPercent Then
            startFound = True
            Exit Do
        End If
        If checkTarget And checkDescription And checkPercent Then
            Chain = Chain + CStr(range(use_Structured(pointer, 10)).Value2) + ","
            If range(use_Structured(pointer, 7)).Value2 = 0 Then startFound = True
        End If
        Count = Count - 1
    Loop
    
    Count = range("表格2").Rows.Count - (cell.Row - range("表格2").Row + 1)
    Set pointer = cell
    Do While pointer.Row < (range("表格2").Row + range("表格2").Rows.Count - 1) And endFound = False
        Set pointer = pointer.offset(1)
        checkTarget = range(use_Structured(pointer, 6)).text = range(use_Structured(cell, 6)).text
        checkDescription = range(use_Structured(pointer, 9)).text = range(use_Structured(cell, 9)).text
        checkPercent = True 'Range(use_Structured(pointer, 7)).Value2 >= Range(use_Structured(cell, 12)).Value2
        If checkTarget And checkDescription And Not checkPercent Then
            endFound = True
            Exit Do
        End If
        If checkTarget And checkDescription And checkPercent Then
            Chain = Chain + CStr(range(use_Structured(pointer, 10)).Value2) + ","
            On Error Resume Next
            If range(use_Structured(pointer, 12)).Value2 >= 1 Then endFound = True
        End If
        Count = Count - 1
    Loop
    getTaskChain = Chain
End Function
Function getCountBySplit(inputS As String, deli As String)
    WordsList = Split(inputS, deli)
    Count = 0
    For i = 0 To UBound(WordsList)
        If WordsList(i) <> vbNullString Then
            Count = Count + 1
        End If
    Next i
    getCountBySplit = Count
End Function

Function TestArr(ByRef arr()) As String
    MsgBox arr(1, 1)
    
    TestArr = 0
End Function

Sub resize2dayScale(toResize As range)

'Changing the 3rd-25the row Height

    
    Dim selected As range
    'Ignore zero
    For Each cell In toResize
        If range(use_Structured(cell, 3)).Value2 > 0 Then
            If Not selected Is Nothing Then
                Set selected = Union(selected, cell)
            Else
                Set selected = cell
            End If
        End If
    Next
    If Not selected Is Nothing Then
        For Each cell In selected
            Height = (20 * 15.8) * (range(use_Structured(cell, 3)).Value2)
            If Height > 15.8 And Height < (20 * 15.8) Then cell.RowHeight = Height
        Next cell
    End If

End Sub

Function WithinSameDay(selected As range) As range
    Dim sameday As range
    Set sameday = selected
    
    For Each cell In selected
        On Error Resume Next:
        Dim pointer As range
        Set pointer = cell
        Do While pointer.Row > range("表格2").Row And startFound = False
            Set pointer = pointer.offset(-1)
            If dateValue(range(use_Structured(pointer, 5))) = dateValue(range(use_Structured(cell, 4))) Then
                Set sameday = Union(sameday, pointer)
            Else
                startFound = True
            End If
        Loop
    
        Set pointer = cell
        Do While pointer.Row < (range("表格2").Row + range("表格2").Rows.Count - 1) And endFound = False
            Set pointer = pointer.offset(1)
            If dateValue(range(use_Structured(pointer, 4))) = dateValue(range(use_Structured(cell, 4))) Then
                Set sameday = Union(sameday, pointer)
            Else
                endFound = True
            End If
        Loop
    Next
    
    Set WithinSameDay = sameday
End Function

Sub resetdayScale()
range("表格2").RowHeight = 15.8
End Sub
