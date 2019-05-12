Attribute VB_Name = "Serielization"
Function Table2JsonHelper(tableName As String)
    Set table = Worksheets(range(tableName).Parent.name).ListObjects(tableName)
    Call Table2Json(tableName, table.DataBodyRange, table.HeaderRowRange)
    Table2JsonHelper = "Saved at " + Str(Now())
End Function

Function TimelineJsonHelper(tableName As String, Optional TitleOverride)
    Set table = Worksheets(range(tableName).Parent.name).ListObjects(tableName)
    
    If Not IsMissing(TitleOverride) Then
        TimelineJsonHelper = Timeline2Json(tableName, table.DataBodyRange, table.HeaderRowRange, range(TitleOverride))
    Else
        TimelineJsonHelper = Timeline2Json(tableName, table.DataBodyRange, table.HeaderRowRange)
    End If
    TimelineJsonHelper = "Saved at " + Str(Now())
End Function

