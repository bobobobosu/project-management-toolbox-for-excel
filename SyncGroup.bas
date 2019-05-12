Attribute VB_Name = "SyncGroup"
Sub SyncGroupColumn()
    Dim tCol As range
    'Set tCol = Range("表格68").Rows(0).offset(0, 7).Resize(1, Range("表格68").Columns.Count - 7)
    Set tCol = range("ResourceTimeline!E4#")
    Dim mCol As range
    Set mCol = range("存取權修正表!E3#")
    Dim asCol As range
    Set asCol = range("存取權增減表!F3#")
    
    
    
    tCol.EntireColumn.Hidden = False
    mCol.EntireColumn.Hidden = False
    asCol.EntireColumn.Hidden = False
    
    'Ungroup 存取權時間表
    For Each r In tCol.Cells
        Dim cell As range
        Set cell = r
        If cell.EntireColumn.OutlineLevel > 1 Then
            Do While cell.EntireColumn.OutlineLevel > 1
                cell.Ungroup
            Loop
        End If
    Next
    
    'Ungroup 存取權修正表
    For Each r2 In mCol.Cells
        Dim cell2 As range
        Set cell2 = r2
        If cell2.EntireColumn.OutlineLevel > 1 Then
            Do While cell2.EntireColumn.OutlineLevel > 1
                cell2.Ungroup
            Loop
        End If
    Next
    
    'Ungroup 存取權增減表
    For Each r3 In asCol.Cells
        Dim cell3 As range
        Set cell3 = r3
        If cell3.EntireColumn.OutlineLevel > 1 Then
            Do While cell3.EntireColumn.OutlineLevel > 1
                cell3.Ungroup
            Loop
        End If
    Next
     
    'group 存取權時間表 & 存取權增減表 & 存取權修正表
    
    
    Dim favorites As Collection
    Set favorites = New Collection
    
    For Each eachFavorite In range("表格55[最愛]")
        If eachFavorite.Value2 = "*" Then
            favorites.Add range("表格55[工作物件]").Cells(eachFavorite.Row - range("表格55[最愛]").Row + 1).Value2
        End If
    Next
    
    

    For l1 = 1 To tCol.Cells.Count
        If Not HasVal(favorites, tCol.Cells(l1).Value2) Then tCol.Cells(l1).Group
    Next
    For l2 = 1 To mCol.Cells.Count
        If Not HasVal(favorites, mCol.Cells(l2).Value2) Then mCol.Cells(l2).Group
    Next
    For l3 = 1 To asCol.Cells.Count
        If Not HasVal(favorites, asCol.Cells(l3).Value2) Then asCol.Cells(l3).Group
    Next
    


    
    
End Sub

Sub SyncGroupRow()

End Sub
