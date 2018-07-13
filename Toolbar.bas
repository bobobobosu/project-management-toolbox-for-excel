Attribute VB_Name = "Toolbar"
Sub SetOne_Click()
    For Each cell In Selection
        cell.Value2 = 1
    Next cell
     Debug.Print ActiveSheet.CodeName
    If ActiveSheet.CodeName = "Worksheet___7" Then
        Range(Evaluate("INDIRECT(""$B$4"")")).Calculate
    End If
End Sub
Sub addOne_Click()
    For Each cell In Selection
        cell.Value2 = cell.Value2 + 1
    Next cell
     Debug.Print ActiveSheet.CodeName
    If ActiveSheet.CodeName = "Worksheet___7" Then
        Range(Evaluate("INDIRECT(""$B$4"")")).Calculate
    End If
End Sub

Sub plusOneDay()
Attribute plusOneDay.VB_ProcData.VB_Invoke_Func = "d\n14"
    Dim offset As String
    offset = InputBox("Input")
    If offset <> vbNullString Then
        For Each cell In Selection
            cell.Value2 = cell.Value2 + Evaluate("=" + offset)
        Next cell
    End If
End Sub
Sub SumToFirst()
    If Selection.Cells.count > 1 Then
            Dim FirstCell As String
            FirstCell = ""
            Dim LastCell As String
            LastCell = ""
            Dim sumVal As Double
            sumVal = 0
            For Each cell In Selection
                If FirstCell = "" Then
                    FirstCell = cell.Address
                End If
                'MsgBox sumVal
                sumVal = sumVal + cell.Value2
                'cell.Value2 = 0
                
                LastCell = cell.Address
            Next cell
            
            Range(FirstCell).Value = sumVal
    End If
End Sub
Sub SumToLast()
    If Selection.Cells.count > 1 Then
            Dim FirstCell As String
            FirstCell = ""
            Dim LastCell As String
            LastCell = ""
            Dim sumVal As Double
            sumVal = 0
            For Each cell In Selection
                If FirstCell = "" Then
                    FirstCell = cell.Address
                End If
                'MsgBox sumVal
                sumVal = sumVal + cell.Value2
                'cell.Value2 = 0
                
                LastCell = cell.Address
            Next cell
            
            Range(LastCell).Value = sumVal
    End If
End Sub
Sub FirstMinusBelow()
    If Selection.Cells.count > 1 Then
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
                    FirstCell = cell.Address
                    firstVal = cell.Value2
                Else
                    'MsgBox sumVal
                    sumVal = sumVal + cell.Value2
                    LastCell = cell.Address
                    'cell.Value2 = 0
                End If
                
            Next cell
            
            Range(FirstCell).Value = firstVal - sumVal
    End If
End Sub
Sub LastMinusBelow()
    If Selection.Cells.count > 1 Then
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
                    LastCell = cell.Address
                    lastVal = cell.Value2
                    'cell.Value2 = 0
            Next cell
            
            Range(LastCell).Value = lastVal + lastVal - sumVal
    End If
End Sub
Sub Divide()
    For Each cell In Selection
        cell.Value2 = cell.Value2 / Range("$B$1").Cells.Value2
    Next cell
End Sub
Sub Divide2()
    For Each cell In Selection
        cell.Value2 = cell.Value2 / Range("$AD$1").Cells.Value2
    Next cell
End Sub

Sub littlebitBigger()
    For Each cell In Selection
        cell.Value2 = cell.Value2 + 0.01
    Next cell
End Sub
