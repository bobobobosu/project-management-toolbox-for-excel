Attribute VB_Name = "AccessTimeline"
Public Function getTableTitle(thisTable As range)
    Dim CellName As String
    CellName = Application.Caller.address
    Dim thiscell As range
    Set thiscell = range(CellName)
    getTableTitle = thisTable.Cells(1).offset(-1, (thiscell.Column - thisTable.Cells(1).Column))
End Function
Public Function getTableTitleAdd(thisTable As range)
    Dim CellName As String
    CellName = Application.Caller.address
    Dim thiscell As range
    Set thiscell = range(CellName)
    getTableTitleAdd = thisTable.Cells(1).offset(-1, (thiscell.Column - thisTable.Cells(1).Column)).address
End Function

Public Function getTableTitleAbove(thisTable As range)
    Dim CellName As String
    CellName = Application.Caller.address
    Dim thiscell As range
    Set thiscell = range(CellName)
    getTableTitleAbove = thisTable.Cells(1).offset(-2, (thiscell.Column - thisTable.Cells(1).Column))
End Function

Public Function getTableTitleR(cell As range) As range
    Set getTableTitleR = Worksheets("交易").Cells(3, cell.Column)
End Function

Public Function getTablebyColumn(cell As range) As range
    Set getTablebyColumn = cell.Columns(Application.Caller.Column - cell.Column + 1)
End Function
Public Function sameFarthest(pos As range) As range
    changed = False
    Do While changed = False
        If pos.offset(-1).Value2 = pos.Value2 Then
            Set pos = pos.offset(-1)
        Else
            changed = True
        End If
    Loop
    Set sameFarthest = pos
End Function
Public Function StructureAboveIgnoreFlag(title As Variant) As range
    columnAdd = "表格2[" + title.Value2 + "]"
    Dim pointer As range
    Set pointer = Worksheets("交易").Cells(Application.Caller.Row - 1, range(columnAdd).Column)
    
    
    Do While range(use_Structured(pointer, 3)).Value2 = 0 And range(use_Structured(pointer, 2)).Value2 <> vbNullString
        Set pointer = pointer.offset(-1)
    Loop
    ff = pointer.address
    Set StructureAboveIgnoreFlag = pointer
End Function
Public Function StructureBelowIgnoreFlag(title As Variant) As range
    columnAdd = "表格2[" + title.Value2 + "]"
    Dim pointer As range
    Set pointer = Worksheets("交易").Cells(Application.Caller.Row + 1, range(columnAdd).Column)
    
    ff = range(use_Structured(pointer, 3)).address
    Do While range(use_Structured(pointer, 2)).Value2 = 0 And range(use_Structured(pointer, 2)).Value2 <> vbNullString
        Set pointer = pointer.offset(1)
    Loop

    Set StructureBelowIgnoreFlag = pointer
End Function

Public Function StructureAboveIgnoreFlagR(title As Variant, cell As range) As range
    columnAdd = "表格2[" + title + "]"
    Dim pointer As range
    Set pointer = Worksheets("交易").Cells(cell.Row - 1, range(columnAdd).Column)
    
    ff = range(use_Structured(pointer, 3)).address
    Do While range(use_Structured(pointer, 3)).Value2 = 0 And range(use_Structured(pointer, 2)).Value2 <> vbNullString
        Set pointer = pointer.offset(-1)
    Loop

    Set StructureAboveIgnoreFlagR = pointer
End Function

Public Function StructureBelowIgnoreFlagR(title As Variant, cell As range) As range
    columnAdd = "表格2[" + title + "]"
    Dim pointer As range
    Set pointer = Worksheets("交易").Cells(cell.Row + 1, range(columnAdd).Column)
    
    ff = range(use_Structured(pointer, 3)).address
    Do While range(use_Structured(pointer, 2)).Value2 = 0 And range(use_Structured(pointer, 2)).Value2 <> vbNullString
        Set pointer = pointer.offset(1)
    Loop

    Set StructureBelowIgnoreFlagR = pointer
End Function

Public Function StructureAboveR(title As Variant, cell As range) As range
    columnAdd = "表格2[" + title + "]"
    Dim pointer As range
    Set pointer = Worksheets("交易").Cells(cell.Row - 1, range(columnAdd).Column)
    Set StructureAboveR = pointer
End Function
Public Function StructureBelowR(title As Variant, cell As range) As range
    columnAdd = "表格2[" + title + "]"
    Dim pointer As range
    Set pointer = Worksheets("交易").Cells(cell.Row + 1, range(columnAdd).Column)
    Set StructureBelowR = pointer
End Function

Public Function StructureAbove(title As Variant) As range
    columnAdd = "表格2[" + title.Value2 + "]"
    Set StructureAbove = Worksheets("交易").Cells(Application.Caller.Row - 1, range(columnAdd).Column)
End Function
Public Function StructureBelow(title As Variant)
    columnAdd = "表格2[" + title.Value2 + "]"
    StructureBelow = Worksheets("交易").Cells(Application.Caller.Row + 1, range(columnAdd).Column)
End Function
Public Function StructureCol(title As range) As range
    columnAdd = "表格2[" + title.Value2 + "]"
    Set StructureCol = range(columnAdd)
End Function
Public Function cellAbove()
    Dim CellName As String
    CellName = Application.Caller.address
    cellAbove = range(CellName).offset(-1, 0).Value2
End Function
Public Function cellAboveIgnoreBlank(cell As range)
    Dim pointer As range
    Set pointer = range(cell.address)
    Set pointer = pointer.offset(-1, 0)
    Do While pointer = vbNullString
        Set pointer = pointer.offset(-1, 0)
    Loop
    'pointer.offset(1, 0).Value2 = 10
    cellAboveIgnoreBlank = pointer.Value2
End Function
Public Function cellAboveAdd()
    Dim CellName As String
    CellName = Application.Caller.address
    cellAboveAdd = range(CellName).offset(-1, 0).address
End Function

Public Function SyncedTable(thisTable As String, thatTable As String, index As String)

End Function
Public Function CustomFunc1(A As Variant, B2 As Variant, C As Variant, D2 As Variant, E2 As Variant, f As Variant, g As Variant, H As Variant, i2 As Variant, j As Variant, k As Variant, L As Variant)
'Original Function:
'=IFERROR(Evals(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(INDEX(Evals("表格62["&getTableTitle(表格68)&"]"),MATCH([@交易物件],表格62[工作物件],0)),"amt",cellAboveAdd()),"title",getTableTitleAbove(表格68)),"cj",[@交易物件]))*1*[@完成],IFERROR(cellAbove()*1,0))+IFERROR(INDEX(Evals(("表格6866["&getTableTitle(表格68)&"]")),MATCH([@編號],表格6866[編號],0)),0)
'Original Function2:
'=IFERROR(Evals(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(INDEX(Evals(SUBSTITUTE("表格62[title]","title",OFFSET(INDIRECT(ADDRESS(ROW(表格68)-1,COLUMN(表格68))),-1,COLUMN()-COLUMN(表格68)))),MATCH([@交易物件],表格62[工作物件],0)),"amt",ADDRESS(ROW()-1,COLUMN())),"title",INDIRECT(ADDRESS(ROW(表格68)-2,COLUMN()))),"cj",[@交易物件]))*1*[@完成],IFERROR(INDIRECT(ADDRESS(ROW()-1,COLUMN()))*1,0))+IFERROR(INDEX(Evals(CONCATENATE("表格6866","[",INDIRECT(ADDRESS(ROW(表格68)-1,COLUMN()-COLUMN(表格68)+1)),"]")),MATCH([@編號],表格6866[編號],0)),0)
Dim B As String
Dim D As String
Dim E As String
Dim i As String
If TypeName(B2) = "Range" Then
    B = B2.Value2
Else
    B = 0
End If
If TypeName(D2) = "Range" Then
    D = D2.Value2
Else
    D = 0
End If
If TypeName(E2) = "Range" Then
    E = E2.Value2
Else
    E = 0
End If
If TypeName(i2) = "Range" Then
    i = i2.Value2
Else
    i = 0
End If

'Values
'getTableTitle (表格68)    -A
'[@交易物件]    -B
'getTableTitleAbove (表格68)   -C
'[@交易物件]  -D
'[@完成]   -E
'cellAbove() -F
'cellAboveAdd() -G
'getTableTitle(表格68)  -H
'[@編號]   -I

'Ranges
'表格62 [工作物件] -K
'表格6866 [編號]  -L

Dim SubFunc1Str As String
'SubFunc1Str = "=IFERROR(Evals(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(INDEX(Evals(" & Chr(34) & "表格62[" & Chr(34) & "&getTableTitle(表格68)&" & Chr(34) & "]" & Chr(34) & "),MATCH([@交易物件],表格62[工作物件],0))," & Chr(34) & "amt" & Chr(34) & ",cellAboveAdd())," & Chr(34) & "title" & Chr(34) & ",getTableTitleAbove(表格68))," & Chr(34) & "cj" & Chr(34) & ",[@交易物件]))*1*[@完成],IFERROR(cellAbove()*1,0))"
SubFunc1Str = "=IFERROR(Evals(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(INDEX(Evals(" & Chr(34) & "表格62[" & Chr(34) & "&" & Chr(34) & (A) & Chr(34) & "&" & Chr(34) & "]" & Chr(34) & "),MATCH(" & Chr(34) & (D) & Chr(34) & "," & (k) & ",0))," & Chr(34) & "amt" & Chr(34) & "," & Chr(34) & (g) & Chr(34) & ")," & Chr(34) & "title" & Chr(34) & "," & Chr(34) & (H) & Chr(34) & ")," & Chr(34) & "cj" & Chr(34) & "," & Chr(34) & (B) & Chr(34) & "))*1*" & (E) & ",IFERROR(" & Chr(34) & (CStr(f)) & Chr(34) & "*1,0))"

Dim SubFunc2Str As String
'SubFunc2Str = ("=IFERROR(INDEX(Evals(") & ("(") & Chr(34) & ("表格6866[") & Chr(34) & "&" & ("getTableTitle(表格68)") & "&" & Chr(34) & ("]") & Chr(34) & (")),MATCH([@編號],表格6866[編號],0)),0)")
SubFunc2Str = ("=IFERROR(INDEX(Evals(") & ("(") & Chr(34) & ("表格6866[") & Chr(34) & "&" & Chr(34) & (A) & Chr(34) & "&" & Chr(34) & ("]") & Chr(34) & (")),MATCH(") & (CStr(i)) & (",") & (L) & (",0)),0)")

func1 = Evaluate(SubFunc1Str)
func2 = Evaluate(SubFunc1Str)

CustomFunc1 = func1 + func2

End Function
Public Function CustomFunc2()
'Original Function:
'{=MAX(IFERROR((returnBuffer(OFFSET(表格68[[#標題],[a1]],1,0,ROWS(表格68),COLUMNS(表格62)-COLUMNS(表格62[[WBS]:[Location]])),[@編號],COLUMNS(表格62)-COLUMNS(表格62[[WBS]:[Location]]))),0)*--((ISNUMBER(SEARCH("+(-1)",OFFSET(INDEX(表格62[Location],MATCH([@交易物件],表格62[工作物件],0)),0,1,1,COLUMNS(表格62)-COLUMNS(表格62[[WBS]:[Location]])))))))}
End Function


Public Function CompleteStatus(r As range)
    CompleteStatus = (Not range(use_Structured(r, 12)).HasFormula) Or ((Not range(use_Structured(r, 2)).HasFormula) And (Not range(use_Structured(r, 8)).HasFormula))
    
End Function
