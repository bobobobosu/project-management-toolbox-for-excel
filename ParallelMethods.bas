Attribute VB_Name = "ParallelMethods"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If
Public Const masterFileName  As String = "Z:\My Drive\_Storage\Backup\Documents\RAMDISK_loc\Documents\root\Data\TC.xlsb"
'Save data to the master file
Public Sub SaveRangeToMaster(masterWorkbookName As String, r As Range)
    Set oXL = GetObject(masterFileName)
    For Each Item In r.Areas
            Dim caledArray As Variant
            caledArray = Item.Formula
            
            On Error Resume Next
            'oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).range(Item.Address)(1) = 0
            If oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).HasArray Then
                 oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula = caledArray
                oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).FormulaArray = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula
            Else
                oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula = caledArray
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
        Range("ам╤у!T2").Formula = functionS
        bypass_String255_in_eval = Range("ам╤у!T2").Value2
End Function
Function CheckRangeIdential(r As Range) As Boolean
If Checkll(r) Then
    Dim identical As Boolean
    identical = True
    Dim SameFormula As String
    SameFormula = restorellAsString(r(1).Formula)
    For Each cell In r
        Dim scell As Range
        Set scell = cell
            If restorellAsString(scell.Formula) <> SameFormula Then
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
Function CheckRangeIdential_noll(r As Range) As Boolean
    Dim identical As Boolean
    identical = True
    Dim SameFormula As String
    SameFormula = r(1).Formula
    For Each cell In r
        If cell.Formula <> SameFormula Then
            identical = False
            CheckRangeIdential_noll = identical
            Exit Function
        End If
    Next cell
    CheckRangeIdential_noll = identical
End Function
Function RefreshCal(toRefresh As Range) As String
    Set toRefresh = Filterll(toRefresh)
        Call restorell(toRefresh)
        Call convertll(toRefresh)
End Function
Sub wgrsteh()
    Dim ggtrweg As Range
    Set ggtrweg = Range("k2:L1896")
    Debug.Print ggtrweg.Address
    Call RefreshCal(ggtrweg)
End Sub
Function Filterll(r As Range, Optional seqFrom As Long, Optional seqTo As Long = -1, Optional columnMode As Long = 0) As Range
    'seqFrom = seqTo =0: No Task
    'seqTo = -1: Filter All
    Dim g As Range
    Set g = Nothing
    If Not (r Is Nothing) Then
        If seqTo = -1 Then
            For Each Item In r.Areas
                For Each cell In Item.Cells
                    Dim r2 As Range
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
                    Dim nocolR As Range
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
                    For i = 1 To r.Columns(j).Rows.count
                        Dim r3 As Range
                        Set r3 = r.Cells(i, j)
                        'If Checkll(r3) Then
                                If (g Is Nothing) Then
                                    Set g = r3
                                Else
                                    Set g = Union(g, r3)
                                End If
                        'End If
                    Next i
                Next j
                Set Filterll = g
            Else
                 Filterll = g
            End If
        End If
    Else
        Set Filterll = g
    End If
End Function
Public Function getItemByIndexInRange(r As Range, indexnum As Variant) As Range
    index = indexnum
    If r.Areas.count > 1 Then
        For i = 1 To r.Areas.count
            If r.Areas(i).Cells.count < index Then
                index = index - r.Areas(i).Cells.count
            Else
                Set getItemByIndexInRange = r.Areas(i).Cells(index)
                Exit For
            End If
        Next i
    Else
        Set getItemByIndexInRange = r.Cells(index)
    End If
End Function

Function Checkll(r As Range) As Boolean
    Dim contains As Boolean
    contains = True
    For Each cell In r.Cells
        If (InStr(cell.Formula, "=ll(")) = 0 Then
            contains = False
            Checkll = contains
            Exit Function
        End If
    Next cell
    Checkll = contains
End Function
Sub gwertd()
    Call convertll(Range("C6:C28"))
End Sub
Public Function convertll(inR As Range)
    If (inR Is Nothing) Then Exit Function
    For Each sArea In inR.Areas
        For Each sCol In sArea.Columns
            Dim thisCol As Range
            Set thisCol = sCol
            
            If (Not Checkll(thisCol.Cells)) And (thisCol.Cells.count > 1) Then
                    Dim calResultFor() As Variant
                    calResultVal = thisCol.Value2
                    calResultFor = thisCol.Formula
                    x = UBound(calResultVal, 1) - LBound(calResultVal, 1) + 1
                    y = UBound(calResultVal, 2) - LBound(calResultVal, 2) + 1
                    
                    For i = 1 To y
                        For j = 1 To x
                            calResultFor(j, i) = "=ll(" & convert2Con((replaceQuotes(calResultFor(j, i)))) & "," + CStr(dynamicCast(WorksheetFunction.IfError(calResultVal(j, i), 0))) + ")"
                        Next j
                    Next i
                    
                If thisCol.HasArray Then
                    If Len(calResultFor(1, 1)) < 255 Then
                        thisCol.Formula = calResultFor
                        thisCol.FormulaArray = thisCol.Formula
                    End If
                Else
                    Debug.Print thisCol.Address
                    Debug.Print inR.Address
                    thisCol.Formula = calResultFor
                End If
            Else
                For Each colC In sCol.Cells
                    Dim cell As Range
                    Set cell = colC
                    If (Not Checkll(cell)) Then
                        If cell.HasArray Then
                            Dim ffff As String
                            ffff = "=ll(" & convert2Con((replaceQuotes(cell.Formula))) & "," + CStr(dynamicCast(WorksheetFunction.IfError(cell.Value2, 0))) + ")"
                            If Len(ffff) < 255 Then
                                cell.Formula = ffff
                                cell.FormulaArray = cell.Formula
                            End If
                            'Call Set_FormulaArray(cell)
                        Else
                            Dim dd As String
                            dd = "=ll(" & convert2Con((replaceQuotes(cell.Formula))) & "," + CStr(dynamicCast(WorksheetFunction.IfError(cell.Value2, 0))) + ")"
                            'If Len(dd) < 255 Then
                                cell.Formula = dd
                            'End If
                        End If
                    End If
                Next colC
            End If
        Next sCol
    Next sArea
End Function
Public Function rangeIdentical(r As Range) As Boolean
    For Each cell In r.Areas(1).Cells
            If (cell.FormulaArray = r.Areas(1).Cells(1).FormulaArray) And (cell.HasArray = r.Areas(1).Cells(1).HasArray) Then
            Else
                rangeIdentical = False
                Exit Function
            End If
    Next cell
    rangeIdentical = True
End Function
Public Sub qwearfdeg()
    Dim gtgt As Range
    Set gtgt = Range("AM6:AM7")
    gtgt.Formula = "=1"
    gtgt.FormulaArray = gtgt.Formula
End Sub

Public Function restorell(inputR As Range)
    If (inputR Is Nothing) Then Exit Function
    Dim cell As Range
    Set cell = inputR
    For Each sArea In cell.Areas
        For Each sColumn In sArea.Columns
            Dim r As Range
            Set r = Filterll(sColumn.Cells)
            If Not (r Is Nothing) Then
                If CheckRangeIdential(r.Cells) And r.Cells.count > 1 Then
                    identicalFormula = restorellAsString(r.Cells(1).Formula)
                    If r.Cells(1).HasArray Then
                        sColumn.Formula = identicalFormula
                        sColumn.FormulaArray = sColumn.Formula
                    Else
                        sColumn.Formula = identicalFormula
                    End If
                Else
                    For Each icell In r.Cells
                        Dim scell As Range
                        Set scell = icell
                        thisFormula = restorellAsString(scell.Formula)
                        If scell.HasArray Then
                            scell.Formula = thisFormula
                            scell.FormulaArray = scell.Formula
                        Else
                            scell.Formula = thisFormula
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



Public Sub gertghsgbdv()
    Dim ggggg As Range
    Set ggggg = Range("$L$5:$L$2225")
    Call restorell(ggggg)
End Sub

Sub gggggrgewsfd()
    
    MsgBox TypeName(dynamicCast(5))

End Sub
Function RangeToArray(inputRange As Range) As Variant
   Dim inputArray As Variant
   inputArray = inputRange.Value
   RangeToArray = inputArray
End Function


Public Sub SetRangeToMaster(masterWorkbookName As String, sheetName As String, rangeAddress As String, val As Variant)
    Dim oXL As Object
    Set oXL = GetObject(, "Excel.Application")
    oXL.Workbooks(masterWorkbookName).Sheets(sheetName).Range(rangeAddress).Value = val
    Set oXL = Nothing
End Sub

'Thread methods
Public Function GetForThreadNr()
    GetForThreadNr = CLng(Mid(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, "_") + 1, InStr(ActiveWorkbook.Name, ".") - InStr(ActiveWorkbook.Name, "_") - 1))
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
Public Sub Set_FormulaArray(r As Range)
    Set r = Range("E1")
    'r.Formula = range("k5").Formula
    Dim formulaStr As String
    formulaStr = r.Formula
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
    jj = Replace(Replace(Replace(r.Formula, Chr(34) & "&" & Chr(34), ""), Chr(34), ""), "~", Chr(34))
    kk = evals(r.Formula)
    w = evals(Replace(Replace(Replace(r.Formula, Chr(34) & "&" & Chr(34), ""), Chr(34), ""), "~", Chr(34)))
    ee = 1
End Sub
Sub gggewrgr()
       Call Set_FormulaArray(Range("E1"))
End Sub
Sub ggf()
    Dim bbb As String
    bbb = bracketString("hhhh", "**", "**")
    MsgBox bbb
    bbb = cropString(bbb, "**", "**")
    MsgBox bbb
End Sub

Public Sub gtershtsxdfb()
    Dim gggggg As String
    gggggg = "nafweildjfbgaoiuewhnirk"
    convert2Con (gggggg)
End Sub


Sub ggggg()
    Set g = Range("D3:D5").Columns(1)
   MsgBox Range("D3:D5").Columns(1).Rows.count
End Sub
Sub gg()
    Dim returnR As Range
    Set r = Range("A1")
    threadFileName = ActiveWorkbook.path & "\" & "TC.xlsb"
    Set oXL = GetObject(threadFileName)
    oXL.Application.Workbooks(1).Sheets(r.Worksheet.Name).Range(r.Address) = r '.PasteSpecial xlFormulas
    

End Sub


Public Function hhhhh()
    gggg = Now()
    Do Until Now() > (gggg + TimeValue("00:00:10"))
        x = x + 1
    Loop
End Function
