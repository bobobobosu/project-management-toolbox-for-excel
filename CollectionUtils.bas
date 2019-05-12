Attribute VB_Name = "CollectionUtils"
Function string2arr(string1 As String, delimeter As String)
    string2arr = Split(string1, delimeter)
End Function
Function json2arr(jsonS As String)
    Dim result As Collection
   Set result = JsonConverter.ParseJson(jsonS)
   json2arr = CollectionToArray(result)
End Function


Function getInList(listStr As String, deli As String, index As Integer) As Variant

    getInList = Split(listStr, deli)(index)
End Function
Function SortByColumn(coll As Collection, index As Variant)
    Dim tmpColl As New Collection
    Min = None
    MinIndex = None
    Dim sorted As New Collection
    Do While coll.Count > 0
        For i = 1 To coll.Count
            If MinIndex = None Then
                Min = coll.Item(1)(index)
                MinIndex = 1
            End If
            If coll.Item(i)(index) <= Min Then
                Min = coll.Item(i)(index)
                MinIndex = i
            End If
        Next
        tmpColl.Add coll.Item(MinIndex)
        coll.Remove (MinIndex)
        MinIndex = None
    Loop
    
    Set coll = tmpColl

End Function
Function getColumnInCollection(col As Collection, columnIndex As Variant) As Variant
    Dim arr()
    ReDim arr(col.Count - 1)
    
    For i = LBound(arr) To UBound(arr)
        arr(i) = col.Item(i + 1)(columnIndex)
    Next
    
    getColumnInCollection = arr

End Function
Public Function CollectionToArray(myCol As Collection) As Variant

    Dim result  As Variant
    Dim cnt     As Long

    ReDim result(myCol.Count)
    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt
    CollectionToArray = result

End Function
