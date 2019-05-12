Attribute VB_Name = "Calculation"


Sub CalculateRange1()
    'Range(Evaluate("INDIRECT(""$C$1"")")).Calculate
    range(range("交易!C1").Value2).Calculate

    
End Sub
Sub CalculateRange2()
    
    Call CalculateRange1
    'Range(Evaluate("INDIRECT(""$C$2"")")).Calculate
    range(range("交易!C2").Value2).Calculate

'    Range(Evaluate("INDIRECT(""$R$1"")")).Calculate
'    Range(Evaluate("INDIRECT(""$R$2"")")).Calculate
    Call CalculateRange1
End Sub
Sub CalculateRange3()
    CalculateList (range("價值表!B5").Value2)
End Sub
Sub CalculateRange4()
    'Call CalculateRange3
    'Range(Evaluate("INDIRECT(""$B$6"")")).Calculate
    
    

    Dim first As Collection
    Set first = Str2Collection(range("AX2").Value2)
    
    Dim columnR As range
    Set columnR = range(first(first.Count))
    On Error Resume Next
    columnR.Calculate
    For Each RowCell In Selection
        For i = 1 To first.Count - 1
            If first(i) <> vbNullString Then
                Dim cell As range
                Set cell = Worksheets("價值表").Cells(RowCell.Row, range(first(i)).Column)
                'MsgBox cell.Address
                Application.StatusBar = "Calculating: " + cell.address
                'On Error Resume Next
                cell.Calculate
                'On Error Resume Next
                Call llCalculate(cell)
            End If
        Next
    Next
    On Error Resume Next
    columnR.Calculate
    
    Call CalculateRange3
    
End Sub

Sub CalculateChart()
    range("表格2[[同步化時間軸]:[趨勢資料軸(理想)]]").Calculate
    Call CalculateRange3
    range("表格55[Curr. % of Time]").Calculate
    range("表格55[Curr. % of 下一日]").Calculate
    range("表格55[Curr. % of 下一月]").Calculate
End Sub

Sub CalculateRange5()
    range(Evaluate("INDIRECT(""$AB$1"")")).Calculate
End Sub

Sub CalculateRange7()
    Call generateID
    Call updateID
    Call CalculateRange1
    range(range("交易!L2").Value2).Calculate
    Call SyncTimeline
    
    Dim mTable As range
    Set mTable = range("表格2")
    Call llCalculate(getSubRange(2, 1, _
                     mTable.Rows.Count, mTable.Columns.Count, _
                    mTable))
'    Range(Evaluate("INDIRECT(""$R$1"")")).Calculate
'    Range(Evaluate("INDIRECT(""$R$2"")")).Calculate
    Call CalculateRange1
End Sub
Function getRangeByNum(col As range, startNum As Variant, endNum As Variant)
    Dim pointer As range
    Dim startR As range
    Dim endR As range
    Set startR = range("表格2").Cells(1).offset(startNum - 1)
    Set endR = range("表格2").Cells(1).offset(endNum - 1)
    Set pointer = startR.Resize(endR.Row - startR.Row)
    getRangeByNum = AddressEx(pointer)
End Function


Sub CalculateRange15()
    Call CalculateRange1
    range("交易!AM2").Calculate
    Call CalculateTable2ByOrder(range(range("交易!AM2")))
    Call CalculateListTimeline
End Sub


Sub CalculateRange8()
    Dim regex As Object
'    Dim r As Range, rC As Range
'
'    ' cells in column A
'    Set r = Range("A2", Cells(Rows.count, "A").End(xlUp))
'
'    Set regex = CreateObject("VBScript.RegExp")
'    regex.Pattern = " \<.*?\>"
'
'    ' loop through the cells in column A and execute regex replace
'    For Each rC In r
'        If rC.Value <> "" Then rC.Value = regex.Replace(rC.Value, "$1$2-01-$3")
'    Next rC


'
'    Dim myCell As Range
'    Set regex = CreateObject("VBScript.RegExp")
'    regex.Pattern = " \<.*?\>"
'    For Each myCell In Range(Evaluate("INDIRECT(""$W$1"")")).Cells
'        myCell.Value = regex.Replace(myCell.Value, "")
'    Next
    range(Evaluate("INDIRECT(""$W$1"")")).Calculate
End Sub
Sub SyncTimeline()
    Call SyncGroupColumn
    Call CalculateRange9
    'Call CalculateRange10
    
    'Set formulas
'    Range(Range("存取權時間表!$E$2")).Value2 = Range(Range("存取權時間表!$E$2")).Value2
'    Call convertll(Range(Range("存取權時間表!$D$1")), "=" + Range("存取權時間表!$E$1").Value2)
    'Range("表格68[Verify]").Value2 = ""

    'Call convertll(Range(Range("存取權時間表!$C$3")), "=" + Range("存取權時間表!$E$3").Value2)
'    Call convertll(Range(Range("存取權時間表!$C$3")), "=" + Range("存取權時間表!$E$3").Value2)
'    Range(Range("存取權時間表!$D$3")).FormulaArray = Range(Range("存取權時間表!$D$3")).formula
    'range("存取權時間表!$C$3").Calculate
    'range("表格68[ResourceCurrentDeli]").Value2 = range("表格68[ResourceCurrentDeli]").Value2
    'Call convertll(range(range("存取權時間表!$C$3")), "=" + range("存取權時間表!$E$3").Value2)

    'Call TransferToCores(Range("表格2[[編號]:[Start Date]]"))
    'Call TransferToCores(Range("表格68"))
    'Call TransferToCores(Range("表格6866"))
    
    'range("趨勢!V2").Value2 = "[Refresh]"
End Sub
Sub SyncAllTable()
    Call CalculateRange9
    Call CalculateRange10
    Call TransferToCores(range("表格2[[編號]:[Start Date]]"))
    Call TransferToCores(range("表格68"))
    Call TransferToCores(range("表格6866"))
End Sub
Sub CalculateRange9()
    Call updateID
    range(range("交易!AS2").Value2).Calculate
    Dim ws As Worksheet
    Set ws = Sheets("存取權修正表")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("表格6866")
    Set sortcolumn = range("表格6866[編號]")
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
       .header = xlYes
       .Apply
    End With
End Sub
Sub CalculateRange10()
    range(range("交易!AT2").Value2).Calculate
    Dim ws As Worksheet
    Set ws = Sheets("存取權時間表")
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("表格68")
    Set sortcolumn = range("表格68[編號]")
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
       .header = xlYes
       .Apply
    End With
End Sub

Sub CalculateRange11()
    Call SyncTimeline
    Call CalculateRange2
    Call SyncTimeline
End Sub

Sub CalculateRange12()
    range("Table56").Calculate
    Call llCalculate(range("Table56"))
End Sub
Sub CalculateRange13()
Call RefreshCal(range("表格68[ResourceCurrentDeli]"))
calculationList = Split(range("存取權時間表!$C$2").Value2, ",")
For i = LBound(calculationList) To UBound(calculationList)
    range(calculationList(i)).Calculate
Next

'    Range(Range("$C$2").Value2).Calculate
'    Range(Range("$C$3").Value2).Calculate
'    Call SyncTimeline
'    Call llCalculate(Range(Range("$C$2").Value2))
'    Call CalculateRange14
End Sub
Sub CalculateListResourceTimeline()
    Call CalculateList(range("$B$2"))
End Sub
Sub CalculateListTimeline()
    Call CalculateList(range("$BL$2"))
End Sub
Function CalculateList(targetAddresses As String)
    calculationList = Split(targetAddresses, ",")
    For i = LBound(calculationList) To UBound(calculationList)
        range(calculationList(i)).Calculate
    Next
    CalculateList = "Calculated at " + Str(Now())
End Function
Function CalculateList2(targetAddresses As String)
    calculationList = Split(targetAddresses, ",")
    For i = LBound(calculationList) To UBound(calculationList)
        Application.OnTime Now, "llCalculate", range(calculationList(i))
    Next
    CalculateList2 = "Calculated at " + Str(Now())
End Function

Sub CalculateRange14()
'    Call TransferToCores(Range("表格68"))
'    Call TransferToCores(Range(Range("存取權時間表!$D$3")))
    range(range("存取權時間表!$C$3").Value2).Calculate
    Call llCalculate(range("表格68[完成]"))
    range("表格2[Dependency Verify]").Calculate
    range("表格2[Buffer]").Calculate
    'Call SyncTimeline
End Sub
Sub CalculateTable2ByOrder(Optional wantedRange As range)
    Debug.Print "Calculation Start"
    If range("趨勢!O2") = 1 Then range("趨勢!O2") = 2

    Application.ScreenUpdating = False
    Dim mselection As range
    If wantedRange Is Nothing Then
        Set mselection = Selection
    Else
        Set mselection = wantedRange
    End If

    'Set first = BuildCalOrder(Range(ActiveSheet.ListObjects("表格2").Range.Rows(mselection(1).Row - Range("表格2").Row + 2).Address))
    orderStr = orderEmptyFilter(range(use_Structured(mselection(1), 38)).Value2)
    Set first = Str2Collection(orderStr)
    
    Set columnR = Str2Collection(first(first.Count), ",")
    On Error Resume Next
    For Each Area In columnR
        vvvv = Now()
        range(Area).Calculate
        Debug.Print samecolumnsOf(range(Area), range("表格2[#Headers]")) + " spend  " + CStr((Now() - vvvv) * 24 * 60 * 60)
    Next
   ' columnR.Calculate
    
    calcount = 2
    Do While calcount > 0
        calcount = calcount - 1
        
        For Each RowCell In mselection
            Dim currentRowCal As Collection
            'Set currentRowCal = BuildCalOrder(Range(ActiveSheet.ListObjects("表格2").Range.Rows(RowCell.Row - Range("表格2").Row + 2).Address))
            orderStr = orderEmptyFilter(range(use_Structured(RowCell, 38)).Value2)
            Set currentRowCal = Str2Collection(orderStr)
            
            For i = 1 To currentRowCal.Count - 1
                
                If currentRowCal(i) <> vbNullString Then
                    Set cell = Cells(RowCell.Row, range(currentRowCal(i)).Column)
                    Application.StatusBar = "Calculating: " + cell.address
                    Dim calCell As range
                    Set calCell = cell
                    On Error Resume Next
                    vvvv = Now()
                    Call llCalculate(calCell)
                    Debug.Print samecolumnsOf(calCell, range("表格2[#Headers]")) + " spend  " + CStr((Now() - vvvv) * 24 * 60 * 60)
                End If
            Next
        Next
    Loop
    
    
    On Error Resume Next
    For Each Area In columnR
        vvvv = Now()
        range(Area).Calculate
        Debug.Print samecolumnsOf(range(Area), range("表格2[#Headers]")) + " spend  " + CStr((Now() - vvvv) * 24 * 60 * 60)
    Next
   ' columnR.Calculate

    Application.ScreenUpdating = True
    
    If range("趨勢!O2") = 2 Then range("趨勢!O2") = 1
    Debug.Print "Calculation End"
End Sub
Function orderEmptyFilter(s) As String
    If s = vbNullString Then
        
        orderEmptyFilter = range("$BK$2").Value2
    Else
        orderEmptyFilter = s
    End If
End Function


Function SubtractRange(rangeA As range, rangeB As Variant) As range
'rangeA is a range to subtract from
'rangeB is the range we want to subtract

 Dim existingRange As range
  Dim resultRange As range
  Set existingRange = rangeA
  Set resultRange = Nothing
  Dim C As range
  For Each C In existingRange
  If Intersect(C, rangeB) Is Nothing Then
    If resultRange Is Nothing Then
      Set resultRange = C
    Else
      Set resultRange = Union(C, resultRange)
    End If
  End If
  Next C
  Set SubtractRange = resultRange
End Function

Function countRightUntilErrorHori(r As range)
    i = 0
    
    Do While (Not IsError(r)) And Not ("N/A" = r.Value2) And Not (r.Value2 = vbNullString)
        i = i + 1
        Set r = r.offset(0, 1)
    Loop
    countRightUntilErrorHori = i
End Function
Function entireRowExcludeSelf() As range
    Set entireRowExcludeSelf = range(ActiveSheet.ListObjects("表格2").range.Rows(range(Application.Caller.address).Row - range("表格2").Row + 2).address)
    Set entireRowExcludeSelf = SubtractRange(entireRowExcludeSelf, Application.Caller)
End Function

Function entireRowExcludeSelfcell(r As range) As range
    Set entireRowExcludeSelfcell = range(ActiveSheet.ListObjects("表格2").range.Rows(r.Row - range("表格2").Row + 2).address)
    Set entireRowExcludeSelfcell = SubtractRange(entireRowExcludeSelfcell, r)
End Function
Function getCalOrderStr(selected As range) As String
    'MsgBox selected.Address
    getCalOrderStr = collection2string(BuildCalOrder(selected))
End Function
Sub generateOrder()
    Dim orderCol As range
    Set orderCol = range("表格2[Order]")
    For Each cell In orderCol
        Dim selected As range
        Set selected = cell
        selected.Value = getCalOrderStr(entireRowExcludeSelfcell(selected))
    Next
End Sub

Sub selectionOrder()
    For Each cell In Selection
        Dim selected As range
        Set selected = cell
        range(use_Structured(cell, 38)).Value = getCalOrderStr(entireRowExcludeSelfcell(selected))
    Next
End Sub

Function Str2Collection(s, Optional dilemeter As String = "|") As Collection
    Dim WrdArray() As String
    WrdArray() = Split(s, dilemeter)
    Dim returnColl As Collection
    Set returnColl = New Collection
    For i = 0 To UBound(WrdArray)
        returnColl.Add WrdArray(i)
    Next
    Set Str2Collection = returnColl
End Function

Function TaskChain2Collection(r) As Collection
    s = range(use_Structured(r, 14)).Value2
    Dim WrdArray() As String
    Dim returnColl As Collection
    Set returnColl = New Collection
    On Error GoTo Error
    WrdArray() = Split(s, ",")
    For i = 0 To UBound(WrdArray)
        returnColl.Add WrdArray(i)
    Next
    Set TaskChain2Collection = returnColl
    Exit Function
    
Error:
    returnColl.Add s
    Set TaskChain2Collection = returnColl
End Function

Function getProjectedDelta(toProject As String)
    Dim f As Collection
    Set f = TaskChain2Collection(range(AddressEx(Application.Caller)))
    
    totalTime = 0
    minPercent = -1
    maxPercent = -1
    
    With Application.WorksheetFunction
    
        For Each task In f
            If (task <> vbNullString) Then
                Dim cell As range
                On Error GoTo err:
                Set cell = .index(range("表格2[實際耗時]"), .Match(val(task), range("表格2[ID]"), 0))
                If (CompleteStatus(cell)) Then
                    
                        
                        totalTime = totalTime + range(use_Structured(cell, 3)).Value
                        
                        If minPercent = -1 Then
                            minPercent = range(use_Structured(cell, 7)).Value
                        Else
                            If range(use_Structured(cell, 7)).Value < minPercent Then
                                minPercent = range(use_Structured(cell, 7)).Value
                            End If
                        End If
                        
                        If maxPercent = -1 Then
                            maxPercent = range(use_Structured(cell, 12)).Value
                        Else
                            If range(use_Structured(cell, 12)).Value > maxPercent Then
                                maxPercent = range(use_Structured(cell, 12)).Value
                            End If
                        End If
                                  
                    
                End If
                
                
            End If
        Next
    End With
    
    On Error GoTo err:
    If (minPercent <> -1 And minPercent <> -1 And totalTime > 0) Then
        If toProject = "Percent" Then
            getProjectedDelta = range(use_Structured(Application.Caller, 2)).Value2 * ((maxPercent - minPercent) / totalTime)
        Else
            getProjectedDelta = totalTime * (range(use_Structured(Application.Caller, 8)).Value2 - range(use_Structured(Application.Caller, 7)).Value2) / ((maxPercent - minPercent))
        End If
    Else
err:
        getProjectedDelta = 0
    End If
    
    

End Function

Function getProjectedDeltaPercent()
    getProjectedDeltaPercent = getProjectedDelta("Percent")
End Function

Function getProjectedDeltaTime()
    getProjectedDeltaTime = getProjectedDelta("Time")

End Function

Function estimatedTime()
    taskEstimate = getProjectedDeltaTime()
End Function
Function collection2string(coll As Collection) As String
    Dim returnS As String
    returnS = ""
    
    For Each ele In coll
        If returnS = "" Then
            returnS = ele
        Else
            returnS = returnS + "|" + ele
        End If
        
    Next
    collection2string = returnS
End Function

Sub ValueTable()
    Dim g As Collection
    Set g = BuildCalOrder(range("A29:AZ29"))
    
    range("AX2").Value2 = collection2string(g)

End Sub

Sub Unifyll()
    Dim selected As range
    Set selected = Selection
    'Unlock ll
    For i = 1 To selected.Cells.Count
        Dim tableTitle As range
        Set tableTitle = getTableTitleR(selected.Cells(i))
        Dim toll As range
        Set toll = StructureCol(tableTitle)
        
        If Checkll(selected.Cells(i)) Then
            Call restorell(toll)
            Call convertll(toll)
        Else
            Call restorell(toll)
        End If
    Next
End Sub

Sub generatePartialOrder()
    range(use_Structured(Selection(1), 38)).Value = getCalOrderStr(CheckPrecedentsRange())
End Sub

Sub setNotInViewll()
    Call restorell(Selection)
    Dim Inview As range
    Set Inview = CheckPrecedentsRange()
    For Each ecell In entireRowExcludeSelfcell(Inview)
        Dim cell As range
        Set cell = ecell
        
        found = False
        For i = 1 To Inview.Cells.Count
            If getTableTitleR(Inview.Cells(i)) = getTableTitleR(cell) Then
                found = True
            End If
        Next
        
        If found = False Then
            Call convertll(cell)
        End If
    Next
End Sub

Function CheckPrecedentsRange() As range
    Dim selected As range
    Set selected = Selection
    Dim Pool As range
    EndofSearch = False
    Set Pool = selected
    Do While EndofSearch = False
        If Pool Is Nothing Then
            startofCount = 0
        Else
            startofCount = Pool.Cells.Count
            MsgBox Pool.Cells.Count
        End If
        
        For Each cell In selected
            precedentsAdd = ""
            Dim noncolOrecedents As range
            Dim Precedents As range
            
            On Error Resume Next
            Set Precedents = cell.DirectPrecedents
            
            For Each Area In Precedents:
                If Area.Rows.Count > 1 Then
                Else
                    If Area.Cells(1).Row = selected.Cells(1).Row Then
                        If noncolOrecedents Is Nothing Then
                            Set noncolOrecedents = Area
                        Else
                            Set noncolOrecedents = Union(noncolOrecedents, Area)
                        End If
                    End If
                End If
            Next

            If Pool Is Nothing Then
                Set Pool = noncolOrecedents
            Else
                Set Pool = Union(Pool, noncolOrecedents)
            End If
            For Each pres In noncolOrecedents:
                Set selected = Union(selected, pres)
            Next
            
        Next
        
        If Pool.Cells.Count - startofCount = 0 Then
            EndofSearch = True
        End If
    Loop
    'Range(use_Structured(selected(1), 38)).Value = AddressEx(Pool)
    Set CheckPrecedentsRange = Pool
End Function
Function BuildCalOrder(selected As range) As Collection
    'Unlock ll
    Dim lled As range
    For i = 1 To selected.Cells.Count
        If Checkll(selected.Cells(i)) Then
            restorell (selected.Cells(i))
            If lled Is Nothing Then
                Set lled = selected.Cells(i)
            Else
                Set lled = Union(lled, selected.Cells(i))
            End If
        End If
    Next



    target = selected.Count
    Dim columnCal As range
    Dim calculated As range
    Dim rngPrecedents As range
    Dim pending As Collection
    Set pending = New Collection
    Dim tocal As Collection
    Set tocal = New Collection
    tocalCount = 0
    endofcal = False
    
    For Each selectedR In selected.Cells
        pending.Add selectedR
    Next
    
    
    Set calculated = range("A1")
    Set columnCal = range("A1")
    Do While target > 0 And endofcal = False
        'check cal diff
        Dim indextodelete As Collection
        Set indextodelete = New Collection
    
        For i = 1 To pending.Count
            Set cell = pending(i)
            
            precedentsAdd = ""
            

            On Error Resume Next
            precedentsAdd = AddressEx(cell.DirectPrecedents)
            'On Error GoTo 0

            
            
            If cell.HasFormula = False Or precedentsAdd = "" Then
                target = target - 1
                Set calculated = Union(calculated, cell)
                indextodelete.Add i
                tocal.Add cell.address
        
            Else
        
                CheckDep = True
                For Each possiblecol In range(precedentsAdd).Areas
                    If possiblecol.Rows.Count > 1 Then
                        Set columnCal = Union(columnCal, possiblecol)
                    End If
                Next
                
                Dim DependencyInRow As range
                Set DependencyInRow = Application.Intersect(range(precedentsAdd), selected)
    
                For Each r In DependencyInRow
                    FoundNotInside = True
                    For Each cal In calculated.Cells
                        If cal.address = r.address Then
                            FoundNotInside = False
                        End If
                    Next

                    If FoundNotInside = True Then
                        CheckDep = False
                    End If
                Next

                
                If CheckDep = True Then
                    target = target - 1
                    Set calculated = Union(calculated, cell)
                    indextodelete.Add i
                    tocal.Add cell.address
                End If
            End If
        Next
        
        
        'check cal diff
        If indextodelete.Count > 0 Then
            For k = indextodelete.Count To 1 Step -1
                pending.Remove indextodelete(k)
            Next
        Else
            endofcal = True
            'add unresolved
            For i = 1 To pending.Count
                Set cell = pending(i)
            Next
        End If
        
    Loop
    
    
    'Add LastCal
    tocal.Add columnCal.address
    
    
    'Relock ll
    For j = 0 To lled.Cells.Count
        Call convertll(lled.Cells(j))
    Next
    
    Set BuildCalOrder = tocal
End Function




Sub FilterSubject()
    CurrField = Selection.Cells(1).Column - range(ActiveCell.ListObject.name).Column + 1

    With ActiveSheet.ListObjects(ActiveCell.ListObject.name).range
        .AutoFilter Field:=CurrField, Criteria1:=Application.WorksheetFunction.Transpose(Selection.Value2), Operator:=xlFilterValues
    End With
    ActiveWindow.ScrollRow = Selection.Row
End Sub

Sub ClearFilterSubject()
    For Each cell In Selection
        CurrField = cell.Column - range(ActiveCell.ListObject.name).Column + 1
        'Field = Range(use_Structured(Selection.Cells(1), 6)).Column - Range("表格2").Column + 1
        ActiveSheet.ListObjects(ActiveCell.ListObject.name).range.AutoFilter Field:=CurrField
    Next cell
     ActiveWindow.ScrollRow = Selection.Row
End Sub


Sub FillResourceByTask()
    Dim TaskNames As range
    Set TaskNames = Application.InputBox(Prompt:="TaskNames", title:="TaskNames", Type:=8)
    Dim TargetIDs As range
    Set TargetIDs = Application.InputBox(Prompt:="TargetIDs", title:="TargetIDs", Type:=8)
    Dim targetWorksheet As Worksheet
    Set targetWorksheet = Worksheets("存取權修正表")
    
    For i = 1 To TaskNames.Cells.Count
    
        Dim TitleCell As range
        Set TitleCell = range("存取權修正表!3:3").Find(Replace(TaskNames.Cells(i).Value2, "t.", "r."), LookIn:=xlValues)
        Dim RowCell As range
        Set RowCell = range("表格6866[ID]").Find(TargetIDs.Cells(i).Value2, LookIn:=xlValues)

        Dim target As range
        Set target = targetWorksheet.Cells(RowCell.Row, TitleCell.Column)
        target.Value2 = target.Value2 - 1
    Next i
    
End Sub
