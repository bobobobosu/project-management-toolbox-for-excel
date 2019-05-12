Attribute VB_Name = "Module14"
Sub ScrollBar1_Change()
    Set SB = ActiveSheet.Shapes("Scroll Bar 1").ControlFormat
    SB.Max = 100
    SB.SmallChange = 1
    SB.LargeChange = 10
    'ActiveCell.VALUE = SB.VALUE / 100
    For Each cell In Selection
        cell.Value = SB.Value / 100
    Next cell
    
    range(Evaluate("INDIRECT(""$B$4"")")).Calculate
    
End Sub


Sub ScrollBar5_Change()
    Set SB = ActiveSheet.Shapes("Scroll Bar 5").ControlFormat
    SB.Max = 100
    SB.SmallChange = 1
    SB.LargeChange = 10

'    'ActiveCell.VALUE = SB.VALUE / 100
'    Dim first As Boolean
'    first = False
'
'    Dim onehundredVal As Double
'    onehundredVal = Range("趨勢!F2").Value2
'
'
'
'    For Each cell In Selection
'        If first = False Then
'            'SB.Value = 100 * cell.Value2 / onehundredVal
'            cell.Value2 = cell.Value2 + onehundredVal * (SB.Value / 100 - (cell.Value2 / onehundredVal))
'            first = True
'        Else
'            If cell.Value2 = vbNullString Then
'                cell.Value2 = 0
'            End If
'            cell.Value = cell.Value + onehundredVal - onehundredVal * (SB.Value / 100 - (cell.Value2 / onehundredVal)) / (Selection.count - 1)
'        End If
'    Next cell
'

    Dim sum As Double
    sum = 0
    For Each cell In Selection
        sum = sum + cell.Value2
    Next cell
    

   Selection(1).Value = sum * SB.Value / 100
   
   For i = 2 To Selection.Count
    Selection(i).Value = sum * (1 - SB.Value / 100) * (1 / (Selection.Count - 1))
   Next i
   Call updateScrollbar5
End Sub


Sub updateScrollbar5()
    Set SB = ActiveSheet.Shapes("Scroll Bar 5").ControlFormat
    If InRange(Selection, range("表格2[預計耗時]")) Then
        
        Dim sum As Double
        sum = 0
        For Each cell In Selection
            sum = sum + cell.Value2
        Next cell
        If sum > 0 Then
            SB.Value = (Selection(1).Value / sum) * 100
        End If
    End If
End Sub

Function InRange(Range1 As range, Range2 As range) As Boolean
    ' returns True if Range1 is within Range2
    InRange = Not (Application.Intersect(Range1, Range2) Is Nothing)
End Function



Sub testubg()

    MsgBox PrevCell
End Sub
