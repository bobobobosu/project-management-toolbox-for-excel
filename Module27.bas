Attribute VB_Name = "Module27"
Function Range2(addresses As Variant)
    Dim addressesArr() As String
    addressesArr() = Split(addresses, ",")
    Dim returnR As range
    
    For i = LBound(addressesArr) To UBound(addressesArr)
        If returnR Is Nothing Then
            Set returnR = range(addressesArr(i))
        Else
            Set returnR = Union(returnR, range(addressesArr(i)))
        End If
    Next

    Set Range2 = returnR
End Function

Function mergeRange(address As String, r As range)
        If address <> vbNullString Then
            merged = False
            Dim listOfRanges As Collection
            Set listOfRanges = Str2Collection(address)
            For i = 1 To listOfRanges.Count
                thisitem = listOfRanges.Item(i)
                freg = Split(thisitem, "!")(0)
                If Split(thisitem, "!")(0) = Split(AddressEx(r), "!")(0) Then
                    listOfRanges.Add AddressEx(Union(r, Range2(thisitem)))
                    listOfRanges.Remove (i)
                    merged = True
                End If
            Next
            If merged = False Then
                address = address + "|" + AddressEx(r)
            Else
                address = collection2string(listOfRanges)
            End If
        Else
            address = AddressEx(r)
        End If
        mergeRange = address
End Function
