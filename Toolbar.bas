Attribute VB_Name = "Toolbar"
Sub SetOne_Click()
        Selection.Value2 = 1
     Debug.Print ActiveSheet.CodeName
    If ActiveSheet.CodeName = "Worksheet___7" Then
        range(Evaluate("INDIRECT(""$B$4"")")).Calculate
    End If
End Sub
Sub addOne_Click()
    For Each cell In Selection
        cell.Value2 = cell.Value2 + 1
    Next cell
     Debug.Print ActiveSheet.CodeName
    If ActiveSheet.CodeName = "Worksheet___7" Then
        range(Evaluate("INDIRECT(""$B$4"")")).Calculate
    End If
End Sub
Sub minusOne_Click()
    For Each cell In Selection
        cell.Value2 = cell.Value2 - 1
    Next cell
     Debug.Print ActiveSheet.CodeName
    If ActiveSheet.CodeName = "Worksheet___7" Then
        range(Evaluate("INDIRECT(""$B$4"")")).Calculate
    End If
End Sub
Sub plusOneDay()
Attribute plusOneDay.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim offset As String
    offset = InputBox("Input")
    Start = Selection.Cells(1).Value2
    If offset <> vbNullString Then
        For Each cell In Selection
            Start = Start + Evaluate("=" + offset)
            cell.Value2 = Start 'cell.Value2 + Evaluate("=" + offset)
        Next cell
    End If
End Sub
Sub SumToFirst()
    If Selection.Cells.Count > 1 Then
            'Check Numeric
            nonnumeric = 0
            For Each cell In Selection
                If Not IsNumeric(cell.Value2) Then
                    nonnumeric = nonnumeric + 1
                End If
            Next
                            
            Dim FirstCell As String
            FirstCell = ""
            Dim LastCell As String
            LastCell = ""
            If nonnumeric = 0 Then
                Dim sumVal As Double
                sumVal = 0
                For Each cell In Selection
                    If FirstCell = "" Then
                        FirstCell = cell.address
                    End If
                    'MsgBox sumVal
                    sumVal = sumVal + cell.Value2
                    'cell.Value2 = 0
                    
                    LastCell = cell.address
                Next cell
                
                range(FirstCell).Value = sumVal
                
            Else

                Dim sumStr As String
                sumStr = ""
                For Each cell In Selection
                    If cell.Value2 <> vbNellString Then
                    
                        If FirstCell = "" Then
                            FirstCell = cell.address
                            sumStr = cell.text
                            
                        Else
                            sumStr = sumStr & vbLf & cell.text
                            cell.Value2 = ""
                        End If
                        
                        LastCell = cell.address
                    
                    End If
                Next cell
                
                range(FirstCell).Value2 = sumStr
            End If
    End If
End Sub
Sub SumToLast()
    If Selection.Cells.Count > 1 Then
            Dim FirstCell As String
            FirstCell = ""
            Dim LastCell As String
            LastCell = ""
            Dim sumVal As Double
            sumVal = 0
            For Each cell In Selection
                If FirstCell = "" Then
                    FirstCell = cell.address
                End If
                'MsgBox sumVal
                sumVal = sumVal + cell.Value2
                'cell.Value2 = 0
                
                LastCell = cell.address
            Next cell
            
            range(LastCell).Value = sumVal
    End If
End Sub
Sub FirstMinusBelow()
    If Selection.Cells.Count > 1 Then
            Dim FirstCell As String
            FirstCell = ""
            Dim LastCell As String
            LastCell = ""
            Dim sumVal As Double
            sumVal = 0
            Dim firstVal As Double
            firstVal = 0
            For Each cell In Selection
                If FirstCell = "" Then
                    FirstCell = cell.address
                    firstVal = cell.Value2
                Else
                    'MsgBox sumVal
                    sumVal = sumVal + cell.Value2
                    LastCell = cell.address
                    'cell.Value2 = 0
                End If
                
            Next cell
            
            range(FirstCell).Value = firstVal - sumVal
    End If
End Sub
Sub LastMinusBelow()
    If Selection.Cells.Count > 1 Then
            Dim FirstCell As String
            FirstCell = ""
            Dim LastCell As String
            LastCell = ""
            Dim sumVal As Double
            sumVal = 0
            Dim lastVal As Double
            lastVal = 0
            For Each cell In Selection

                    'MsgBox sumVal
                    sumVal = sumVal + cell.Value2
                    LastCell = cell.address
                    lastVal = cell.Value2
                    'cell.Value2 = 0
            Next cell
            
            range(LastCell).Value = lastVal + lastVal - sumVal
    End If
End Sub
Sub Divide()
    For Each cell In Selection
        cell.Value2 = cell.Value2 / range("$B$1").Cells.Value2
    Next cell
End Sub
Sub Divide2()
    For Each cell In Selection
        cell.Value2 = cell.Value2 / range("$AD$1").Cells.Value2
    Next cell
End Sub

Sub littlebitBigger()
    For Each cell In Selection
        cell.Value2 = cell.Value2 + 0.01
    Next cell
End Sub
