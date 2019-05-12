Attribute VB_Name = "RangeTools"
Function SyncCol(thisCol, thatCol)
    SyncCol = thatCol(Application.Caller.Row - thisCol(1).Row + 1)
End Function
