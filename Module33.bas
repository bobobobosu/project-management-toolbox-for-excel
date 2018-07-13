Attribute VB_Name = "Module33"
Function MaxProduct(firstA As Variant, secondA As Variant)
    Dim maxVal As Variant
    maxVal = firstA(1) * secondA(1)
    For i = LBound(firstA) To UBound(firstA)
        tmp = firstA(i) * secondA(i)
        If tmp > maxVal Then
            maxVal = tmp
        End If
    Next i
    MaxProduct = maxVal
End Function

Function returnBuffer(allR As Range, position As Long, columnNum As Long) As Variant

    Dim resultA() As Variant
    
    Dim totalColumns As Long
    totalColumns = columnNum
    
    Dim totalRows As Long
    totalRows = allR.Rows.count
    
    ReDim resultA(1 To 1, 1 To totalColumns)
    Dim nowColumn As Long
    
    'MsgBox totalColumns
    Dim columnofthistask As Range
    For nowColumn = 1 To totalColumns
        
      ' MsgBox nowColumn
        Set columnofthistask = getSubRange(1, nowColumn, totalRows, nowColumn, allR)
        'MsgBox columnofthistask.Rows.Count
        
        resultA(1, nowColumn) = bufferStart(position, columnofthistask)
    Next nowColumn


    
    returnBuffer = resultA

End Function


Function testthis(allR As Range) As Variant
    Dim gg As Range
    Set gg = getSubRange(1, 1, 3, 1, allR)
    testthis = gg.Rows.count
End Function

Public Function bufferStart(position As Variant, target2 As Range)
    'MsgBox target2.Cells(current).Address
    Dim capacity As Double
    
    Dim currentCell As Range
    Dim numCell As Range
    
    With Application.WorksheetFunction
        Set currentCell = .index(target2, .Match(position, Range("表格68[ID]"), 0))
        Set numCell = .index(Range("表格68[編號]"), .Match(position, Range("表格68[ID]"), 0))
    End With
    
    If currentCell.Value2 < 0 Then
        bufferStart = numCell.Value2
        Exit Function
    End If
    
    capacity = -1 * (currentCell.Value - currentCell.offset(-1).Value)
    If capacity > 0 Then
        A = 0
    End If
    Dim pointer As Range
    Set pointer = currentCell.offset(-1)
    Set numCell = numCell.offset(-1)
    Do While (pointer.Value2 >= capacity)
        On Error GoTo Oops
        Set pointer = pointer.offset(-1)
        Set numCell = numCell.offset(-1)
Oops:
        Exit Do
    Loop
    bufferStart = numCell.Value2
End Function
Public Function bufferStart2(current As Variant, target2 As Range)

If current <> 1 Then
    
    Dim count As Integer
    Dim Counter As Integer
    
    Dim capacity As Double
    
    capacity = -1 * (target2.Cells(current).Value - target2.Cells(current - 1).Value)
    
    Counter = 0
    For Each cell In target2
        Counter = Counter + 1
        
        
        If Counter < current Then
            If cell.Value2 >= capacity Then
            count = count + 1
            Else
                count = 0
            End If
        End If
    Next

    bufferStart = current - count
Else
    bufferStart = 1
End If

End Function

Public Function MakesSound(AnimalName As Range) As Variant
Dim Ansa() As Variant
Dim vData As Variant
Dim j As Long
vData = AnimalName.Value2
ReDim Ansa(1 To UBound(vData), 1 To 1)
For j = 1 To UBound(vData)
    Select Case vData(j, 1)
    Case Is = "Duck"
        Ansa(j, 1) = 1
    Case Is = "Cow"
        Ansa(j, 1) = 2
    Case Is = "Bird"
        Ansa(j, 1) = 3
    Case Is = "Sheep"
        Ansa(j, 1) = "Ba-Ba-Ba!"
    Case Is = "Dog"
        Ansa(j, 1) = "Woof!"
    Case Else
        Ansa(j, 1) = "Eh?"
    End Select
Next j
MakesSound = Ansa
End Function


Public Function genVertArray(paraR As Range, toevalR As Range, evalFunc As String, arraySize As Integer) As Variant
Dim Ansa() As Variant
Dim vData As Variant
Dim j As Long
vData = AnimalName.Value2
ReDim Ansa(1 To UBound(vData), 1 To 1)
For j = 1 To UBound(vData)
    Select Case vData(j, 1)
    Case Is = "Duck"
        Ansa(j, 1) = 1
    Case Is = "Cow"
        Ansa(j, 1) = 2
    Case Is = "Bird"
        Ansa(j, 1) = 3
    Case Is = "Sheep"
        Ansa(j, 1) = "Ba-Ba-Ba!"
    Case Is = "Dog"
        Ansa(j, 1) = "Woof!"
    Case Else
        Ansa(j, 1) = "Eh?"
    End Select
Next j
MakesSound = Ansa
End Function




Function getSubRange(iRow1 As Long, iCol1 As Long, _
                     iRow2 As Long, iCol2 As Long, _
                     sourceRange As Range) As Range
' Returns a sub-range of the source range (non-boolean version of makeSubRange().
' Inputs:
'   iRow1       -  Row and colunn indices in the sourceRange of
'   iCol1          the upper left and lower right corners of
'   iRow2          the requested subrange.
'   iCol2
'   sourceRange - The range from which a sub-range is requested.
'
' Return: Reference to a sub-range of sourceRange bounded by the input row and
'         and column indices.
' Notes: A null range will be returned if the following is not true.
'        1 <= iRow1 <= SourceRange.Rows.count
'        1 <= iRow2 <= SourceRange.Rows.count
'        1 <= iCol1 <= SourceRange.Columns.count
'        1 <= iCol2 <= SourceRange.Columns.count

   Const AM1 = 64 'Ascii value of 'A' = 65; Asc('A') - 1 = 64
   Dim rangeStr As String

'   If (1 <= iRow1) And (iRow1 <= sourceRange.rows.count) And _
'      (1 <= iRow2) And (iRow2 <= sourceRange.rows.count) And _
'      (1 <= iCol1) And (iCol1 <= sourceRange.Columns.count) And _
'      (1 <= iCol2) And (iCol2 <= sourceRange.Columns.count) Then
'      rangeStr = Chr(AM1 + iCol1) & CStr(iRow1) & ":" _
'               & Chr(AM1 + iCol2) & CStr(iRow2)
'      Set getSubRange = sourceRange.Range(rangeStr)
'   Else
'      Set getSubRange = Nothing
'   End If
    If sourceRange.Cells.count > 1 Then
        Set getSubRange = Range(sourceRange.Cells(iRow1, iCol1), sourceRange.Cells(iRow2, iCol2))
    Else
        Set getSubRange = sourceRange
    End If
End Function 'getSubRange()

