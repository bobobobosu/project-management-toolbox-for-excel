Attribute VB_Name = "AccessTimeline"
Public Function getTableTitle(thisTable As Range)
    Dim CellName As String
    CellName = Application.Caller.Address
    Dim thisCell As Range
    Set thisCell = Range(CellName)
    getTableTitle = thisTable.Cells(1).offset(-1, (thisCell.Column - thisTable.Cells(1).Column))
End Function
Public Function getTableTitleAdd(thisTable As Range)
    Dim CellName As String
    CellName = Application.Caller.Address
    Dim thisCell As Range
    Set thisCell = Range(CellName)
    getTableTitleAdd = thisTable.Cells(1).offset(-1, (thisCell.Column - thisTable.Cells(1).Column)).Address
End Function

Public Function getTableTitleAbove(thisTable As Range)
    Dim CellName As String
    CellName = Application.Caller.Address
    Dim thisCell As Range
    Set thisCell = Range(CellName)
    getTableTitleAbove = thisTable.Cells(1).offset(-2, (thisCell.Column - thisTable.Cells(1).Column))
End Function

Public Function cellAbove()
    Dim CellName As String
    CellName = Application.Caller.Address
    cellAbove = Range(CellName).offset(-1, 0).Value2
End Function
Public Function cellAboveAdd()
    Dim CellName As String
    CellName = Application.Caller.Address
    cellAboveAdd = Range(CellName).offset(-1, 0).Address
End Function

Public Function SyncedTable(thisTable As String, thatTable As String, index As String)

End Function
Public Sub ggggggggg()
    Call CustomFunc1("getTableTitle(表格68)", 1, "表格6866 [編號]")
End Sub
Public Function CustomFunc1(A As Variant, B2 As Variant, c As Variant, D2 As Variant, E2 As Variant, f As Variant, g As Variant, H As Variant, i2 As Variant, j As Variant, k As Variant, L As Variant)
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
