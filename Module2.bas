Attribute VB_Name = "Module2"
Option Explicit



Public Function getMultFromCell(fromRange As Variant, toRange As Variant, from_to As Variant, fromCell As Variant, toCell As Variant, returnMode As Integer) As Double
On Error GoTo Error_handler:
    Dim fromArray As Variant
    fromArray = fromRange.Value
    Dim toArray As Variant
    toArray = toRange.Value
    Dim from_toArray As Variant
    from_toArray = from_to.Value
    

    Dim graph As Variant
    ReDim graph(0)
    
    'Generate Dictionary
    Dim mdictionary As Variant
    ReDim mdictionary(0)
    Call GenerateDict(fromArray, toArray, mdictionary)
    
    'Generate Graph
    Call GenerateGraph(fromArray, toArray, graph, mdictionary)
    
    'get mult
    Dim resultArray As Variant
    ReDim resultArray(0)
    resultArray = GetMultiplicationBetweenNodes(fromArray, toArray, from_toArray, fromCell, toCell, mdictionary, graph)
    
    'return
    'returnMode: 1=min 2=max 3=avg
    If returnMode = 1 Then
        'MAX
        Dim count As Integer, maxVal As Double
        maxVal = resultArray(1)
        For count = 2 To UBound(resultArray)
            If resultArray(count) > maxVal Then
                maxVal = resultArray(count)
            End If
        Next count
        getMultFromCell = maxVal
    ElseIf returnMode = 2 Then
        'min
        Dim Count2 As Integer, minVal As Double
        minVal = resultArray(1)
        For Count2 = 2 To UBound(resultArray)
            If resultArray(Count2) < minVal Then
                minVal = resultArray(Count2)
            End If
        Next Count2
        getMultFromCell = minVal
    ElseIf returnMode = 3 Then
        Dim avg As Double
        avg = 0
        Dim nums As Double
        For nums = LBound(resultArray) + 1 To UBound(resultArray)
            avg = avg + resultArray(nums)
        Next
        If UBound(resultArray) Then
            avg = avg / (UBound(resultArray) - LBound(resultArray))
        End If
        getMultFromCell = avg
    End If

Error_handler:
    'getMultFromCell
    'Debug.Print Join(resultArray, ",")
        
    
End Function


Function DFS(ByVal str_node As String, ByVal end_node As String, ByRef resultArray As Variant, graph As Variant, visited_in As Variant, fromA As Variant, toA As Variant, from_toA As Variant, mdictionary As Variant) As Variant

    Dim visited     As Variant

    visited = visited_in
    
    
    
    Dim cur_node    As String
    Dim child_node  As Variant
    Dim k           As Variant
    cur_node = str_node
    child_node = graph(cur_node)
    ReDim Preserve visited(UBound(visited) + 1)
    visited(UBound(visited)) = str_node
    'Debug.Print Join(visited, ", ")
    'Debug.Print cur_node
        
        
    Dim head As Long
    Dim foot As Long
    head = visited(1)
    foot = visited(UBound(visited))
    'Debug.Print id_to_String(head, mdictionary)
    'Debug.Print id_to_String(foot, mdictionary)
    
    
    Dim multi As Double
    multi = GetMultiplication(visited, mdictionary, fromA, toA, from_toA)
    
    
    Dim end_node_Long As Long
    end_node_Long = string_to_id(end_node, mdictionary)
    If end_node_Long = foot Then
        ReDim Preserve resultArray(UBound(resultArray) + 1)
        resultArray(UBound(resultArray)) = multi
        'Debug.Print '-------------'
        'Debug.Print multi
    End If
    

    'Debug.Print TypeName(child_node)
    Dim childNum As Integer
    childNum = UBound(child_node)
    'Debug.Print childNum
    If TypeName(child_node) <> "Empty" Then
        For Each k In child_node
            
            
            If Not b_value_in_array(k, visited) Then
                Call DFS(k, end_node, resultArray, graph, visited, fromA, toA, from_toA, mdictionary)
            End If
                    
        Next k
    End If

    

    
    
End Function

Public Function b_value_in_array(my_value As Variant, my_array As Variant, Optional b_is_string As Boolean = False) As Boolean

    Dim l_counter   As Long

    If b_is_string Then
        my_array = Split(my_array, ":")
    End If

    For l_counter = LBound(my_array) To UBound(my_array)
        my_array(l_counter) = CStr(my_array(l_counter))
    Next l_counter

    b_value_in_array = Not IsError(Application.Match(CStr(my_value), my_array, 0))
    
End Function

Public Function string_to_id(my_String As String, dictionary As Variant) As Long
    If IsInArray(my_String, dictionary) = -1 Then
        ReDim Preserve dictionary(UBound(dictionary) + 1)
        dictionary(UBound(dictionary)) = my_String
        string_to_id = UBound(dictionary)
    Else
        string_to_id = IsInArray(my_String, dictionary)
    End If
End Function

Public Function id_to_String(id As Long, dictionary As Variant) As String
    id_to_String = dictionary(id)
End Function


Function IsInArray(stringToBeFound As String, arr As Variant) As Long
  Dim i As Long
  ' default return value if value not found in array
  IsInArray = -1

  For i = LBound(arr) To UBound(arr)
    If StrComp(stringToBeFound, arr(i), vbTextCompare) = 0 Then
      IsInArray = i
      Exit For
    End If
  Next i
End Function



Function GenerateDict(fromArray As Variant, toArray As Variant, mdictionary As Variant) As Variant

    Dim element As Variant
    Dim i As Long
    
    'add to dictionary
    For Each element In fromArray
        Dim elef As String
        elef = element
        i = string_to_id(elef, mdictionary)
        'Debug.Print elef
    Next
    
    For Each element In toArray
        Dim elet As String
        elet = element
        i = string_to_id(elet, mdictionary)
    Next

End Function

Function GenerateGraph(fromArray As Variant, toArray As Variant, graph As Variant, mdictionary As Variant) As Variant
    'Generate graph
    Dim tmpdictionary As Variant
    tmpdictionary = mdictionary
    Dim element As Variant

    For Each element In tmpdictionary
        Dim j As Integer
        Dim pendArray As Variant
        ReDim pendArray(0)
        For j = 1 To UBound(fromArray)
            If fromArray(j, 1) = element Then
                ReDim Preserve pendArray(UBound(pendArray) + 1)
                Dim toString As String
                toString = toArray(j, 1)
                pendArray(UBound(pendArray)) = string_to_id(toString, mdictionary)
            End If
        
        Next j
        For j = 1 To UBound(toArray)
            If toArray(j, 1) = element Then
                ReDim Preserve pendArray(UBound(pendArray) + 1)
                Dim fromString As String
                fromString = fromArray(j, 1)
                pendArray(UBound(pendArray)) = string_to_id(fromString, mdictionary)
            End If
       
        Next j
        
        ReDim Preserve graph(UBound(graph) + 1)
        graph(UBound(graph) - 1) = pendArray

    Next


End Function

Function GetMultiplicationBetweenNodes(fromArray As Variant, toArray As Variant, from_to As Variant, ByVal fromNode As String, ByVal toNode As String, mdictionary As Variant, graph As Variant) As Variant
    'get mult
    Dim visited0 As Variant
    ReDim visited0(0)
    
    Dim count As Integer
    count = 0
    Dim resultArray As Variant
    ReDim resultArray(0)

    Call DFS(string_to_id(fromNode, mdictionary), toNode, resultArray, graph, visited0, fromArray, toArray, from_to, mdictionary)
    GetMultiplicationBetweenNodes = resultArray
    
    
    


End Function



Function GetMultiplication(visited As Variant, mdictionary As Variant, fromA As Variant, toA As Variant, from_toA As Variant) As Double
    
    
    Dim multip As Double
    multip = 1
    
    
    
    
     
    Dim cou As Integer
    cou = 0
    For cou = 1 To UBound(visited) - 1
        Dim RowSearchCou As Integer
        RowSearchCou = 1
        For RowSearchCou = 1 To UBound(fromA)
            Dim fromStr As String
            fromStr = fromA(RowSearchCou, 1)
            Dim toStr As String
            toStr = toA(RowSearchCou, 1)
            Dim tomult As Double
            If string_to_id(fromStr, mdictionary) = visited(cou) And string_to_id(toStr, mdictionary) = visited(cou + 1) Then
                tomult = from_toA(RowSearchCou, 1)
                multip = multip * tomult
            End If
            
            If string_to_id(toStr, mdictionary) = visited(cou) And string_to_id(fromStr, mdictionary) = visited(cou + 1) Then
                tomult = from_toA(RowSearchCou, 1)
                
                multip = multip / tomult
            End If
            
            
        
        Next
        

    
        
    Next
    GetMultiplication = multip
    

End Function

Sub WriteArrayToImmediateWindow(arrSubA As Variant)

Dim rowString As String
Dim iSubA As Long
Dim jSubA As Long

rowString = ""

Debug.Print
Debug.Print
Debug.Print "The array is: "
For iSubA = 1 To UBound(arrSubA, 1)
    rowString = arrSubA(iSubA, 1)
    For jSubA = 2 To UBound(arrSubA, 2)
        rowString = rowString & "," & arrSubA(iSubA, jSubA)
    Next jSubA
    Debug.Print rowString
Next iSubA

End Sub

''redim preserve both dimensions for a multidimension array *ONLY
'Public Function ReDimPreserve(aArrayToPreserve, nNewFirstUBound, nNewLastUBound)
'    ReDimPreserve = False
'    'check if its in array first
'    If IsArray(aArrayToPreserve) Then
'        'create new array
'        ReDim aPreservedArray(nNewFirstUBound, nNewLastUBound)
'        'get old lBound/uBound
'        nOldFirstUBound = UBound(aArrayToPreserve, 1)
'        nOldLastUBound = UBound(aArrayToPreserve, 2)
'        'loop through first
'        For nFirst = LBound(aArrayToPreserve, 1) To nNewFirstUBound
'            For nLast = LBound(aArrayToPreserve, 2) To nNewLastUBound
'                'if its in range, then append to new array the same way
'                If nOldFirstUBound >= nFirst And nOldLastUBound >= nLast Then
'                    aPreservedArray(nFirst, nLast) = aArrayToPreserve(nFirst, nLast)
'                End If
'            Next
'        Next
'        'return the array redimmed
'        If IsArray(aPreservedArray) Then ReDimPreserve = aPreservedArray
'    End If
'End Function
'
'
'
