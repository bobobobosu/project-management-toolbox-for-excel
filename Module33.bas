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

Function getTitleColumnRange(table As Variant, titleExample As Variant) As range
    titleFormula = titleExample.formula
    Dim tabletitleR As range
    Set tabletitleR = table.offset(-2).Resize(1, table.Columns.Count)
    rightCol = tabletitleR.Column
    leftCol = tabletitleR.Column + tabletitleR.Columns.Count - 1
    For i = 1 To tabletitleR.Cells.Count
        If tabletitleR.Cells(i).Value2 <> vbNullString And tabletitleR.Cells(i).formula = titleFormula And tabletitleR.Cells(i).Column < leftCol Then leftCol = tabletitleR.Cells(i).Column
        If tabletitleR.Cells(i).Value2 <> vbNullString And tabletitleR.Cells(i).formula = titleFormula And tabletitleR.Cells(i).Column > rightCol Then rightCol = tabletitleR.Cells(i).Column
    Next

    Set getTitleColumnRange = tabletitleR.offset(0, leftCol - tabletitleR.Column).Resize(1, rightCol - leftCol + 1)
End Function
Function countTitleColumnRange(table As Variant, titleExample As Variant)
    titleFormula = titleExample.formula
    Dim tabletitleR As range
    Set tabletitleR = table.offset(-2).Resize(1, table.Columns.Count)
    rightCol = tabletitleR.Column
    leftCol = tabletitleR.Column + tabletitleR.Columns.Count - 1
    For i = 1 To tabletitleR.Cells.Count
        If tabletitleR.Cells(i).Value2 <> vbNullString And tabletitleR.Cells(i).formula = titleFormula And tabletitleR.Cells(i).Column < leftCol Then leftCol = tabletitleR.Cells(i).Column
        If tabletitleR.Cells(i).Value2 <> vbNullString And tabletitleR.Cells(i).formula = titleFormula And tabletitleR.Cells(i).Column > rightCol Then rightCol = tabletitleR.Cells(i).Column
    Next
    countTitleColumnRange = rightCol - leftCol + 1
End Function
Function minuschk(data As Variant) As Variant
    Dim dep()
    dep = data
    
    For i = LBound(dep) To UBound(dep)
        If InStr(dep(i, 1), "-") = 0 Then
            dep(i, 1) = 0
        Else
            dep(i, 1) = 1
        End If
    Next
    
    minuschk = dep
End Function
Function arrOR(arr1 As Variant, arr2 As Variant) As Variant
    Dim resultA()
    resultA = arr1
    For i = LBound(arr1) To UBound(arr1)
        If arr2(i, 1) <> 0 Then
            resultA(i, 1) = 1
        End If
    Next
    arrOR = resultA
End Function

Function thisDataRow(table As Variant, titleExample As Variant, thisRowCell As Variant) As Variant
    thisDataRow = Application.Transpose(samecolumnsOf(getTitleColumnRange(table, titleExample), thisRowCell).Value2)
End Function
Function thisDataRowR(table As Variant, titleExample As Variant, thisRowCell As Variant) As Variant
    Set thisDataRowR = samecolumnsOf(getTitleColumnRange(table, titleExample), thisRowCell)
End Function
Function FindBuffer(table As Variant, titleExample As Variant, thisRowCell As Variant)
    nearestRow = -1
    Dim rowdata As range
    Set rowdata = thisDataRowR(table, titleExample, thisRowCell)
    edwfv = AddressEx(rowdata)
    Dim pointer As range
    For Each cell In rowdata
        Set pointer = cell
        dep = pointer.Value2 - pointer.offset(-1).Value2
        If dep < 0 Then
            
            Do While pointer.offset(-1).Value2 + dep >= 0
                edfghjedwfv = AddressEx(pointer)
                Set pointer = pointer.offset(-1)
            Loop
            If pointer.Row > nearestRow Then nearestRow = pointer.Row
        Else
            
        End If
    Next
    If nearestRow <> -1 Then
        FindBuffer = samerowsOf(Cells(nearestRow, 1), thisRowCell).Value2
    Else
        FindBuffer = -1
    End If
End Function

Function getBuffer(resourceCurr As range, resourceAccu As range, index As range)
    startRow = resourceCurr.Row
    Do While startRow - resourceCurr.Row < 200
        If Satisfied(resourceCurr, resourceAccu) Then
            Set resourceCurr = resourceCurr.offset(-1)
            Set resourceAccu = resourceAccu.offset(-1)
            Set index = index.offset(-1)
        Else
            getBuffer = index.Value
            Exit Function
        End If
    Loop
    getBuffer = index.Value
    Exit Function
End Function

Function returnBuffer(allr As range, Position As Long, columnNum As Long) As Variant

    Dim resultA() As Variant
    
    Dim totalColumns As Long
    totalColumns = columnNum
    
    Dim totalRows As Long
    totalRows = allr.Rows.Count
    
    ReDim resultA(1 To 1, 1 To totalColumns)
    Dim nowColumn As Long
    
    'MsgBox totalColumns
    Dim columnofthistask As range
    For nowColumn = 1 To totalColumns
        
      ' MsgBox nowColumn
        Set columnofthistask = getSubRange(1, nowColumn, totalRows, nowColumn, allr)
        'MsgBox columnofthistask.Rows.Count
        
        resultA(1, nowColumn) = bufferStart(Position, columnofthistask)
    Next nowColumn


    
    returnBuffer = resultA

End Function
Sub FindFromBeginning()
Dim Position As Integer
Position = InStr(1, "Excel VBA", "V", vbBinaryCompare)
MsgBox Position
End Sub
Function getManualDependancy(id As Integer) As Variant
    Dim resultA As Variant
    Dim tosearch As range
    With Application.WorksheetFunction
        If .CountIf(range("表格6866[ID]"), id) > 0 Then
            Set tosearch = range(.index(range("表格6866[Description]"), .Match(id, range("表格6866[ID]"), 0)).offset(0, 1), _
            .index(range("表格6866[Description]"), .Match(id, range("表格6866[ID]"), 0)).offset(0, range("表格62").Columns.Count - 6))
            rrr = tosearch.address
            ReDim resultA(1 To 1, 1 To tosearch.Cells.Count)
            For i = 1 To tosearch.Cells.Count
                On Error Resume Next
                fgg = Application.WorksheetFunction.Search("-", CStr(tosearch(i).Value2))
                If err <> 0 Then
                '   String Not found
                    resultA(1, i) = 0
                Else
                    resultA(1, i) = 1
                End If
                On Error GoTo 0
            Next
            getManualDependancy = resultA
        Else
            Exit Function
        End If

    End With
End Function

Function getTaskDependancy(task As String) As Variant
    Dim resultA As Variant
    Dim tosearch As range
    With Application.WorksheetFunction
        If .CountIf(range("表格62[工作物件]"), task) > 0 Then
            Set tosearch = range(.index(range("表格62[Location]"), .Match(task, range("表格62[工作物件]"), 0)).offset(0, 1), _
            .index(range("表格62[Location]"), .Match(task, range("表格62[工作物件]"), 0)).offset(0, range("表格62").Columns.Count - 6))
            
            ReDim resultA(1 To 1, 1 To tosearch.Cells.Count)
            For i = 1 To tosearch.Cells.Count
                On Error Resume Next
                fgg = Application.WorksheetFunction.Search("-", CStr(tosearch(i).Value2))
                If err <> 0 Then
                '   String Not found
                    resultA(1, i) = 0
                Else
                    resultA(1, i) = 1
                End If
                On Error GoTo 0
            Next
            getTaskDependancy = resultA
        Else
            Exit Function
        End If

    End With
End Function


Function testthis(allr As range) As Variant
    Dim gg As range
    Set gg = getSubRange(1, 1, 3, 1, allr)
    testthis = gg.Rows.Count
End Function

Public Function bufferStart(Position As Variant, target2 As range)
    'MsgBox target2.Cells(current).Address
    Dim capacity As Double
    
    Dim currentCell As range
    Dim numCell As range
    
    With Application.WorksheetFunction
        If .CountIf(range("表格68[ID]"), Position) > 0 Then
            Set currentCell = .index(target2, .Match(Position, range("表格68[ID]"), 0))
            Set numCell = .index(range("表格68[編號]"), .Match(Position, range("表格68[ID]"), 0))
        Else
        End If

    End With
    
    If currentCell.Value2 < 0 Then
        bufferStart = numCell.Value2
        Exit Function
    End If
    
    capacity = -1 * (currentCell.Value - currentCell.offset(-1).Value)
    If capacity > 0 Then
        A = 0
    End If
    Dim pointer As range
    Set pointer = currentCell.offset(-1)
    Set numCell = numCell.offset(-1)
    Do While (pointer.Value2 >= capacity) And (pointer.Row > target2.Cells(1).Row)
        'On Error GoTo Oops
        Set pointer = pointer.offset(-1)
        Set numCell = numCell.offset(-1)
'Oops:
'        Exit Do
    Loop
    bufferStart = numCell.offset(1).Value2
End Function
Public Function bufferStart2(current As Variant, target2 As range)

If current <> 1 Then
    
    Dim Count As Integer
    Dim Counter As Integer
    
    Dim capacity As Double
    
    capacity = -1 * (target2.Cells(current).Value - target2.Cells(current - 1).Value)
    
    Counter = 0
    For Each cell In target2
        Counter = Counter + 1
        
        
        If Counter < current Then
            If cell.Value2 >= capacity Then
            Count = Count + 1
            Else
                Count = 0
            End If
        End If
    Next

    bufferStart = current - Count
Else
    bufferStart = 1
End If

End Function

Public Function MakesSound(AnimalName As range) As Variant
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


Public Function genVertArray(paraR As range, toevalR As range, evalFunc As String, arraySize As Integer) As Variant
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


Function getRangeBetween(rng1 As range, rng2 As range) As range
    leftTopRow = WorksheetFunction.Min(rng1.Row, rng2.Row)
    leftTopCol = WorksheetFunction.Min(rng1.Column, rng2.Column)
    rightBotRow = WorksheetFunction.Max(rng1.Row, rng2.Row)
    rightBotCol = WorksheetFunction.Max(rng1.Column, rng2.Column)
    Set getRangeBetween = range(rng1.Worksheet.Cells(leftTopRow, leftTopCol), rng1.Worksheet.Cells(rightBotRow, rightBotCol))
End Function


Function getSubRange(iRow1 As Variant, iCol1 As Variant, _
                     iRow2 As Variant, iCol2 As Variant, _
                     sourceRange As range) As range
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
    If sourceRange.Cells.Count > 1 Then
        Set getSubRange = range(sourceRange.Cells(iRow1, iCol1), sourceRange.Cells(iRow2, iCol2))
    Else
        Set getSubRange = sourceRange
    End If
End Function 'getSubRange()

