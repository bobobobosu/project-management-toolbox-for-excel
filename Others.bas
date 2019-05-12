Attribute VB_Name = "Others"
Sub ReOpen()
    Application.DisplayAlerts = False
    Workbooks.Open ActiveWorkbook.path & "\" & ActiveWorkbook.name
    Application.DisplayAlerts = True
End Sub

Sub removeDups()
Dim col As range
For Each col In range("A:Z").Columns
    With col
        .RemoveDuplicates Columns:=1, header:=xlYes
    End With
Next col
End Sub
Sub getDependancy()
    
    For Each cell In Selection
        Dim r As range
        Set r = cell
        Call LoopThroughString(r)
    Next
End Sub
Function LoopThroughString(cell As range)
 
Dim Counter As Integer
Dim MyString As String
Dim tmps As String
    Set pointer = cell
    BraOrder = 0
    MyString = cell.Value2

    Dim col As New Collection


    For Counter = 1 To Len(MyString)
        'do something to each character in string
        'here we'll msgbox each character
        
        If Mid(MyString, Counter, 1) = "[" Then
            BraOrder = BraOrder + 1
        ElseIf Mid(MyString, Counter, 1) = "]" Then
            BraOrder = BraOrder - 1
            If BraOrder = 0 Then
            found = False
                For Each icol In col
                    If icol = tmps Then
                        found = True
                    End If
                Next
            
                If found = False Then
                    col.Add tmps
                    Set pointer = pointer.offset(1)
                    pointer.Value2 = tmps
                    tmps = ""
                Else
                    tmps = ""
                End If
            End If
        Else
            If BraOrder > 0 Then
                tmps = tmps + Mid(MyString, Counter, 1)
            End If

        End If
Next
End Function

Sub prefix()
    For Each cell In Selection
        cell.Value2 = "@" + cell.Value2
    Next
End Sub


Function toArray(col As Collection)
  Dim arr() As Variant
  ReDim arr(0 To col.Count - 1) As Variant
  For i = 1 To col.Count
      arr(i - 1) = col(i)
  Next
  toArray = arr
End Function

Function SetAllToNegative()
    For Each cell In Selection
        If cell.Value2 <> vbNullString Then
            If cell.Value2 > 0 Then
                cell.Value2 = cell.Value2 * -1
            End If
        End If
    Next
End Function


