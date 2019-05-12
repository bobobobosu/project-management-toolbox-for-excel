Attribute VB_Name = "Module32"
Public Function randBetweenExcludeRange(lngBottom As Long, lngTop As Long, _
                                        rngExcludeValues As range) As Variant
    Dim C                               As range
    Dim dict                            As Object
    Dim i                               As Long
    Dim blNoItemsAvailable              As Boolean
    Dim lngTest                         As Long
    'some notes on code:
    'It'd probably be a good idea to check for values that are only integers in the range
    'you might be able to sort the already excluded values, choose a number between 1 and
    'the number of remaining available values and then generate that from a full list of
    'values.  (maybe by making the dictionary hold available values only?)
 
    'I'm pretty sure the comment above doesn't make a lot of sense.  If it
    'did, i'd have tried to implement it.
 
    If lngBottom > lngTop Then
        randBetweenExcludeRange = CVErr(xlErrNA)
        Exit Function
    End If
    'get a list of all items in range
    'i = 0
    Set dict = CreateObject("Scripting.dictionary")
    For Each C In rngExcludeValues
        'I should have really only checked for c.values that are longs.
        If IsNumeric(C.Value) Then
            If C.Value >= lngBottom And C.Value <= lngTop Then
                If Not dict.Exists(C.Value) Then
                    dict.Add C.Value, ""
                End If
            End If
        End If
    Next C
 
    'check to make sure that there are values available to use
    If dict.Count >= lngTop - lngBottom + 1 Then
        'initialize error holder to true
        blNoItemsAvailable = True
        For i = lngBottom To lngTop
            If Not dict.Exists(i) Then
                blNoItemsAvailable = False
                Exit For
            End If
        Next i
    End If
    If blNoItemsAvailable Then
        randBetweenExcludeRange = CVErr(xlErrNA)
        Exit Function
    End If
    'this bit could (probably) be made a lot more efficient.  see notes at top
    'of code
 
 
    Do
        lngTest = Int(Rnd() * (lngTop - lngBottom + 1)) + lngBottom
        If Not dict.Exists(lngTest) Then
            randBetweenExcludeRange = lngTest
            Exit Function
        End If
    Loop
End Function


Sub generateID()
    For Each cell In Evaluate("表格2[ID]")
        Dim usedId As range
        Set usedId = Evaluate("表格2[ID]")
        Dim rownum As Long
        rownum = Evaluate("ROWS(表格2[ID])")
        If cell.Value2 = "" Then
                cell.Value = randBetweenExcludeRange(1, rownum, usedId)
        End If
        
    Next cell
    'Range(Evaluate("INDIRECT(""BB2"")")).Value2 = ""
    Call ClearId
End Sub
Sub ClearId()
    range("表格2[ID]").Cells(1).offset(1, 0).Clear
End Sub

Sub updateID_6866()
    Call updateIDtable2("表格6866[ID]")
End Sub

Sub updateID_68()
    Call updateIDtable2("表格68[ID]")
End Sub

Sub updateID()
    Call updateIDtable2("表格6866[ID]")
'    Call updateIDtable2("表格68[ID]")
End Sub
Sub updateIDtable2(tableToUpdate As String)
    Dim usedId As Collection
    Set usedId = New Collection
    Dim existId As Collection
    Set existId = New Collection
    
    Dim existIdA As Variant
    Dim usedIdA As Variant
    existIdA = range("表格2[ID]").Value2
    usedIdA = range(tableToUpdate).Value2
    
    For i = 1 To UBound(existIdA)
        If CStr(existIdA(i, 1)) <> vbNullString Then
            existId.Add existIdA(i, 1), CStr(existIdA(i, 1))
        End If
    Next
    
    For i = 1 To UBound(usedIdA)
        If CStr(usedIdA(i, 1)) <> vbNullString Then
            If Contains(existId, CStr(usedIdA(i, 1))) Then
                existId.Remove CStr(usedIdA(i, 1))
            End If
        End If
    Next
    
    
    For Each cell In Evaluate(tableToUpdate)
        If cell.Value2 = "" And existId.Count > 0 Then
            cell.Value2 = existId(1)
            existId.Remove 1
        End If
    Next cell
    
End Sub
Sub updateIDtable(tableToUpdate As String)
    Dim usedId As range
    Set usedId = Evaluate(tableToUpdate)
    Dim existId As range
    Set existId = Evaluate("表格2[ID]")
    Dim selected As Long
    
    Dim avalibleIds() As Long
    ReDim Preserve avalibleIds(1 To 1) As Long
    avalibleIds(UBound(avalibleIds)) = 0
    
    
    Dim usedIds As Variant
    Dim existIds As Variant
    usedIds = usedId.Value
    existIds = existId.Value

    For Each existIdv In existIds
    
    
        Dim found As Integer
        found = 0
        For Each usedIdv In usedIds
            If usedIdv = existIdv Then
                found = 1
            End If
        Next usedIdv
        
        If found = 0 Then
            ReDim Preserve avalibleIds(1 To UBound(avalibleIds) + 1) As Long
            avalibleIds(UBound(avalibleIds)) = existIdv
        End If
    Next existIdv

    
    Dim Count As Long
    Count = 2

    For Each cell In Evaluate(tableToUpdate)
        If UBound(avalibleIds) >= Count Then
            If cell.Value2 = "" Then
                
                    cell.Value2 = avalibleIds(Count)
                    Count = Count + 1
                    
            End If
        End If
    Next cell

End Sub


Public Function Contains(col As Collection, key As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    Contains = True
    obj = col(key)
    Exit Function
err:

    Contains = False
End Function
