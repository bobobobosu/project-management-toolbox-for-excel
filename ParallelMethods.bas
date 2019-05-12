Attribute VB_Name = "ParallelMethods"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

'Save data to the master file
Public Sub SaveRangeToMaster(masterWorkbookName As String, r As range)
    Set oXL = GetObject(masterFileName)
    For Each Item In r.Areas
            Dim caledArray As Variant
            caledArray = Item.formula
            
            On Error Resume Next
            'oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).range(Item.Address)(1) = 0
            If oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).HasArray Then
                 oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula = caledArray
                oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).FormulaArray = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula
            Else
                oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula = caledArray
            End If

    Next Item
    Set oXL = Nothing
End Sub

Public Function ll(toeval As String, val As Variant)
   ll = val
End Function
Public Function getf_ll(toeval As String, val As Variant)
   getf_ll = restoreQuotes(toeval)
End Function
Public Function restorellAsString(inputFormula As String) As String
        Dim v As String
        v = "=getf_" & Right(inputFormula, Len(inputFormula) - 1)
        If Len(v) > 250 Then
            restorellAsString = restoreQuotes(bypass_String255_in_eval(v))
        Else
            restorellAsString = restoreQuotes(evals(v))
        End If
End Function
Function bypass_String255_in_eval(functionS As String) As String
        range("ам╤у!T2").formula = functionS
        bypass_String255_in_eval = range("ам╤у!T2").Value2
End Function
Function CheckRangeIdential(r As range) As Boolean
If Checkll(r) Then
    Dim identical As Boolean
    identical = True
    Dim SameFormula As String
    SameFormula = restorellAsString(r(1).formula)
    For Each cell In r
        Dim scell As range
        Set scell = cell
            If restorellAsString(scell.formula) <> SameFormula Then
                identical = False
                CheckRangeIdential = identical
                Exit Function
            End If
    Next cell
    CheckRangeIdential = identical
Else
    CheckRangeIdential = False
End If
End Function
Function CheckRangeIdential_noll(r As range) As Boolean
    Dim identical As Boolean
    identical = True
    Dim SameFormula As String
    SameFormula = r(1).formula
    For Each cell In r
        If cell.formula <> SameFormula Then
            identical = False
            CheckRangeIdential_noll = identical
            Exit Function
        End If
    Next cell
    CheckRangeIdential_noll = identical
End Function
Function RefreshCalAdd(toRefreshadd As Variant) As String
    Dim toRefresh As range
    Set toRefresh = Range2(toRefreshadd)
    Call RefreshCal(toRefresh)
End Function
Function RefreshCal(toRefresh As range) As String
    Set toRefresh = Filterll(toRefresh)
    If toRefresh Is Nothing Then Exit Function
    Application.Workbooks(1).Sheets(toRefresh.Worksheet.name).Activate
    Call restorell(toRefresh)
    Call convertll(toRefresh)
End Function
Function Filterll(r As range, Optional seqFrom As Long, Optional seqTo As Long = -1, Optional columnMode As Long = 0) As range
    'seqFrom = seqTo =0: No Task
    'seqTo = -1: Filter All
    Dim g As range
    Set g = Nothing
    If Not (r Is Nothing) Then
        If seqTo = -1 Then
            For Each Item In r.Areas
                For Each cell In Item.Cells
                    Dim r2 As range
                    Set r2 = cell
                    If Checkll(r2) Then
                        If (g Is Nothing) Then
                            Set g = r2
                        Else
                            Set g = Union(g, r2)
                        End If
                    End If
                Next cell
            Next Item
            Set Filterll = g
            Exit Function
        ElseIf seqTo > 0 Then
            'Cell mode
            If columnMode = 0 Then
                For i = seqFrom To seqTo
                    Dim nocolR As range
                    Set nocolR = getItemByIndexInRange(r, i)
                    'If Checkll(nocolR) Then
                        If (g Is Nothing) Then
                            Set g = nocolR
                        Else
                            Set g = Union(g, nocolR)
                        End If
                    'End If
                Next i
                Set Filterll = g
            'Column mode
            ElseIf columnMode <> 0 Then
                For j = (seqFrom) To (seqTo)
                
                    If (g Is Nothing) Then
                        Set g = r.Columns(j)
                    Else
                        Set g = Union(g, r.Columns(j))
                    End If
'                    g = r.Columns(j).Rows.count
'                    For i = 1 To r.Columns(j).Rows.count
'                        Dim r3 As Range
'                        Set r3 = r.Cells(i, j)
'                        'If Checkll(r3) Then
'                                If (g Is Nothing) Then
'                                    Set g = r3
'                                Else
'                                    Set g = Union(g, r3)
'                                End If
'                        'End If
'                    Next i
                Next j
                rg = r.address
                Set Filterll = g
            Else
                 Filterll = g
            End If
        End If
    Else
        Set Filterll = g
    End If
End Function
Public Function getItemByIndexInRange(r As range, indexnum As Variant) As range
    index = indexnum
    If r.Areas.Count > 1 Then
        For i = 1 To r.Areas.Count
            If r.Areas(i).Cells.Count < index Then
                index = index - r.Areas(i).Cells.Count
            Else
                Set getItemByIndexInRange = r.Areas(i).Cells(index)
                Exit For
            End If
        Next i
    Else
        Set getItemByIndexInRange = r.Cells(index)
    End If
End Function

Function Checkll(r As range) As Boolean
    Dim Contains As Boolean
    Contains = True
    For Each cell In r.Cells
        If (InStr(cell.formula, "=ll(")) = 0 Then
            Contains = False
            Checkll = Contains
            Exit Function
        End If
    Next cell
    Checkll = Contains
End Function
Public Function convertll(inR As range, Optional customFormula As Variant)
    If (inR Is Nothing) Then Exit Function
    For Each sArea In inR.Areas
        For Each sCol In sArea.Columns
            Dim thisCol As range
            Set thisCol = sCol
            f = IsMissing(customFormula)
            If (Not IsMissing(customFormula) Or (Not Checkll(thisCol.Cells))) And (thisCol.Cells.Count > 1) Then
                    Dim calResultFor() As Variant
                    
                    calResultVal = thisCol.Value2
                    calResultFor = thisCol.formula
                    
                    If Not IsMissing(customFormula) Then
                        For index = LBound(calResultFor) To UBound(calResultFor)
                            calResultFor(index, 1) = customFormula
                        Next
                    End If
                    
                    x = UBound(calResultVal, 1) - LBound(calResultVal, 1) + 1
                    y = UBound(calResultVal, 2) - LBound(calResultVal, 2) + 1
                    
                    For i = 1 To y
                        For j = 1 To x
                            calResultFor(j, i) = "=ll(" & convert2Con((replaceQuotes(calResultFor(j, i)))) & "," + CStr(dynamicCast(WorksheetFunction.IfError(calResultVal(j, i), 0))) + ")"
                        Next j
                    Next i
                    
                If thisCol.HasArray Then
                    If Len(calResultFor(1, 1)) < 255 Then
                        thisCol.formula = calResultFor
                        thisCol.FormulaArray = thisCol.formula
                    End If
                Else
                    Debug.Print thisCol.address
                    Debug.Print inR.address
                    thisCol.formula = calResultFor
                End If
            Else
                For Each colC In sCol.Cells
                    Dim cell As range
                    Set cell = colC
                    
                    calResultVal2 = cell.Value2
                    calResultFor2 = cell.formula
                    If Not IsMissing(customFormula) Then
                            formula = customFormula
                    End If
                    
                    If (Not IsMissing(customFormula)) Or (Not Checkll(cell)) Then
                        If cell.HasArray Then
                            Dim ffff As String
                            ffff = "=ll(" & convert2Con((replaceQuotes(calResultFor2))) & "," + CStr(dynamicCast(WorksheetFunction.IfError(calResultVal2, 0))) + ")"
                            If Len(ffff) < 255 Then
                                cell.formula = ffff
                                cell.FormulaArray = cell.formula
                            End If
                            'Call Set_FormulaArray(cell)
                        Else
                            Dim dd As String
                            dd = "=ll(" & convert2Con((replaceQuotes(calResultFor2))) & "," + CStr(dynamicCast(WorksheetFunction.IfError(calResultVal2, 0))) + ")"
                            'If Len(dd) < 255 Then
                                cell.formula = dd
                            'End If
                        End If
                    End If
                Next colC
            End If
        Next sCol
    Next sArea
End Function
Public Function rangeIdentical(r As range) As Boolean
    For Each cell In r.Areas(1).Cells
            If (cell.FormulaArray = r.Areas(1).Cells(1).FormulaArray) And (cell.HasArray = r.Areas(1).Cells(1).HasArray) Then
            Else
                rangeIdentical = False
                Exit Function
            End If
    Next cell
    rangeIdentical = True
End Function

Public Function restorell(inputR As range)
    If (inputR Is Nothing) Then Exit Function
    Dim cell As range
    Set cell = inputR
    For Each sArea In cell.Areas
        For Each sColumn In sArea.Columns
            Dim r As range
            Set r = Filterll(sColumn.Cells)
            If Not (r Is Nothing) Then
                If CheckRangeIdential(r.Cells) And r.Cells.Count > 1 Then
                    identicalFormula = restorellAsString(r.Cells(1).formula)
                    If r.Cells(1).HasArray Then
                        On Error Resume Next
                        sColumn.formula = identicalFormula
                        On Error Resume Next
                        sColumn.FormulaArray = sColumn.formula
                    Else
                        On Error Resume Next
                        sColumn.formula = identicalFormula
                    End If
                Else
                    For Each icell In r.Cells
                        Dim scell As range
                        Set scell = icell
                        thisFormula = restorellAsString(scell.formula)
                        If scell.HasArray Then
                            On Error Resume Next
                            scell.formula = thisFormula
                            On Error Resume Next
                            scell.FormulaArray = scell.formula
                        Else
                            On Error Resume Next
                            scell.formula = thisFormula
                        End If
                    Next icell
                End If
            End If
        Next sColumn
    Next sArea
End Function
Public Sub Enablell()
        Call convertll(Selection)
End Sub
Public Sub Disablell()
    Call restorell(Selection)
End Sub
Public Sub Setll(inputR As range, formula As String, val As String)
  
End Sub
Public Function cropString(inputstr As String, Optional front As String = "**", Optional endStr As String = "**")
    first = InStr(inputstr, front) + Len(front)
    last = InStrRev(inputstr, endStr) - 1
    cropString = Mid(inputstr, first, (last - first) + 1)
End Function
Public Function bracketString(inputstr As String, Optional front As String = "**", Optional endStr As String = "**")
    bracketString = front + inputstr + endStr
End Function
Public Function replaceQuotes(inputstr As Variant)
    replaceQuotes = Replace(inputstr, Chr(34), "~")
End Function
Public Function restoreQuotes(inputstr As String)
    restoreQuotes = Replace(inputstr, "~", Chr(34))
End Function
Public Function convert2Con(inputString As String, Optional charLimit As Long = 250)
    Dim ReturnString As String
    ReturnString = "CONCATENATE(" & Chr(34) & Chr(34)
    Do Until Len(inputString) < charLimit
        ReturnString = ReturnString & "," & Chr(34) & Left(inputString, charLimit) & Chr(34)
        inputString = Right(inputString, Len(inputString) - charLimit)
    Loop
    ReturnString = ReturnString & "," & Chr(34) & Left(inputString, charLimit) & Chr(34) & ")"
    
    convert2Con = ReturnString
End Function
Function dynamicCast(myVariant As Variant) As Variant
If IsArray(myVariant) Then
    Dim element As Variant
    For Each element In myVariant
        element = Chr(34) & element & Chr(34)
    Next element
    dynamicCast = myVariant
Else
    If VarType(myVariant) = vbString Then
        dynamicCast = Chr(34) & myVariant & Chr(34)
    Else
        dynamicCast = myVariant
    End If
End If
End Function
Function RangeToArray(inputRange As range) As Variant
   Dim InputArray As Variant
   InputArray = inputRange.Value
   RangeToArray = InputArray
End Function


Public Sub SetRangeToMaster(masterWorkbookName As String, sheetName As String, rangeAddress As String, val As Variant)
    Dim oXL As Object
    Set oXL = GetObject(, "Excel.Application")
    oXL.Workbooks(masterWorkbookName).Sheets(sheetName).range(rangeAddress).Value = val
    Set oXL = Nothing
End Sub

'Thread methods
Public Function GetForThreadNr()
    GetForThreadNr = CLng(Mid(ActiveWorkbook.name, InStr(ActiveWorkbook.name, "_") + 1, InStr(ActiveWorkbook.name, ".") - InStr(ActiveWorkbook.name, "_") - 1))
End Function


Public Function restore2Str(inputString As String)
    restore2Str = evals(inputString)
End Function


Public Function ReadMyFunc(rng As String, index As Long) As String

    Dim vaArgs As Variant

    vaArgs = Split(Split(Left(rng, Len(rng) - 1), "ll(")(1), ",")
    'vaArgs = Split(Split(Left(rng.Formula, Len(rng.Formula) - 1), "ll(")(1), ",")

    'MsgBox "First arg is: " & vaArgs(0) & vbNewLine & "Second arg is: " & vaArgs(1)
    
    ReadMyFunc = vaArgs(index)
End Function
Public Sub Set_FormulaArray(r As range)
    Set r = range("E1")
    'r.Formula = range("k5").Formula
    Dim formulaStr As String
    formulaStr = r.formula
    formulaStr = replaceQuotes(formulaStr)
    formulaStr = Right(formulaStr, Len(formulaStr) - 1)
    inputString = formulaStr
    formulaStr = Replace(formulaStr, "=", "")
     r.FormulaArray = "=" & Chr(34) & "BBBBB" & Chr(34)

    Dim ReturnString As String
    ReturnString = ""
    charLimit = 25
    Do Until Len(inputString) <= charLimit
        ReturnString = Left(inputString, charLimit) & Chr(34) & "&" & Chr(34) & "BBBBB"
        r.Replace "BBBBB", ReturnString
        inputString = Right(inputString, Len(inputString) - charLimit)
    Loop
    r.Replace "BBBBB", inputString
'    On Error Resume Next
'    r.Replace Chr(34), ""
'    r.Replace Chr(34) & "&" & Chr(34), ""
    jj = Replace(Replace(Replace(r.formula, Chr(34) & "&" & Chr(34), ""), Chr(34), ""), "~", Chr(34))
    kk = evals(r.formula)
    W = evals(Replace(Replace(Replace(r.formula, Chr(34) & "&" & Chr(34), ""), Chr(34), ""), "~", Chr(34)))
    ee = 1
End Sub


Public Function hhhhh()
    Call convertll(Selection, "=NOW()")
End Function
