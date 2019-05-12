VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Paste"
   ClientHeight    =   1085
   ClientLeft      =   84
   ClientTop       =   408
   ClientWidth     =   5160
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Scrollbar_justUpdated As Integer
Public init As Boolean


Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then range("ам╤у!P2").Value = 1
If CheckBox1.Value = False Then range("ам╤у!P2").Value = 0
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True Then range("ам╤у!O2").Value = 1
If CheckBox2.Value = False Then range("ам╤у!O2").Value = 0
End Sub
'
'Private Sub CheckBox3_Click()
'If CheckBox3.Value = True Then Range("ам╤у!S2").Value = 1
'If CheckBox3.Value = False Then Range("ам╤у!S2").Value = 0
'End Sub

Private Sub CommandButton1_Click()
    Call toText_Click
End Sub

Private Sub CommandButton10_Click()
minusOne_Click
End Sub

Private Sub CommandButton11_Click()
SumToFirst
End Sub

Private Sub CommandButton12_Click()
FirstMinusBelow
End Sub

Private Sub CommandButton13_Click()
SumToLast
End Sub

Private Sub CommandButton14_Click()
Call LastMinusBelow
End Sub

Private Sub CommandButton15_Click()
    Call plusOneDay
End Sub

Private Sub CommandButton16_Click()
Call Paste_Input_Click
End Sub

Private Sub CommandButton17_Click()
Call Paste_Value_Click
End Sub



Private Sub CommandButton18_Click()
Call toText_Click
End Sub

Private Sub CommandButton19_Click()
Call SetOne_Click
End Sub

Private Sub CommandButton2_Click()
Call Paste_Formula_Auto_Fast
End Sub

Private Sub CommandButton20_Click()
For Each cell In Selection
        cell.Value2 = cell.Value2 * TextBox1.Value
Next cell
End Sub

Private Sub CommandButton21_Click()
sum = 0
For Each cell In Selection
   sum = sum + cell.Value2
Next cell
For Each cell In Selection
        cell.Value2 = sum / Selection.Count
Next cell
End Sub

Private Sub CommandButton22_Click()
    Call Enablell
End Sub

Private Sub CommandButton23_Click()
    Call Disablell
End Sub

Private Sub CommandButton24_Click()
Selection.SpecialCells(xlCellTypeVisible).Select
End Sub

Private Sub CommandButton25_Click()
 Application.CutCopyMode = False
End Sub

Private Sub CommandButton26_Click()
    If Selection.Cells.Count = 1 Then
        Call CalculateNext(Selection)
    Else
        Call customCalculate
    End If
End Sub

Private Sub CommandButton27_Click()
    Application.CutCopyMode = False
End Sub

Private Sub CommandButton28_Click()
            range("ам╤у!U2").Value = 1
            Call FilterSubject
End Sub

Private Sub CommandButton29_Click()
            range("ам╤у!U2").Value = 0
            Call ClearFilterSubject
End Sub

Private Sub CommandButton3_Click()
Call Paste_Formula_Ignore_Text_Auto
End Sub

Private Sub CommandButton30_Click()
    Selection.SpecialCells(xlCellTypeVisible).Select
End Sub

Private Sub CommandButton31_Click()
    result = GeneratePlan(Selection, InputBox("Mode", , 1))
    Call FillExampleWithArrayOfDict(result)
End Sub

Private Sub CommandButton32_Click()
    Call SwapCells
End Sub

Private Sub CommandButton4_Click()
Call Paste_Value_Transpose_Click
End Sub

Private Sub CommandButton5_Click()
Call Paste_Selection_Vertical
End Sub

Private Sub CommandButton6_Click()
For Each cell In Selection
        cell.Value2 = cell.Value2 / TextBox1.Value
Next cell
End Sub

Private Sub CommandButton7_Click()
Call SetZero
End Sub

Private Sub CommandButton8_Click()
Call TimesMinusOne_Click
End Sub

Private Sub CommandButton9_Click()
Call SetOne_Click
End Sub


Private Sub ScrollBar1_Change()
'MsgBox Scrollbar_justUpdated
If Scrollbar_justUpdated = 1 Then
    Scrollbar_justUpdated = Scrollbar_justUpdated - 1
Else
    If Selection.Count >= 2 Then
    
        ScrollBar1.Max = 100
        ScrollBar1.SmallChange = 1
        ScrollBar1.LargeChange = 10
        
        
        Dim sum As Double
        sum = 0
        For Each cell In Selection
            sum = sum + cell.Value2
        Next cell
        
    
       Selection(1).Value = sum * ScrollBar1.Value / 100
       
       For i = 2 To Selection.Count
        Selection(i).Value = sum * (1 - ScrollBar1.Value / 100) * (1 / (Selection.Count - 1))
       Next i
    End If
End If
End Sub

Private Sub TextBox1_Change()
    range("╔Ф╘Ж!B1") = TextBox1.Value
End Sub


Private Sub ToggleButton1_Click()
    If init = False Then
        If ToggleButton1.Value = True Then
            range("ам╤у!U2").Value = 1
            Call FilterSubject
        Else
            range("ам╤у!U2").Value = 0
            Call ClearFilterSubject
        End If
    End If
End Sub


Private Sub UserForm_Initialize()
    init = True
    CheckBox1.Value = range("ам╤у!P2").Value
    CheckBox2.Value = range("ам╤у!O2").Value
'    CheckBox3.Value = Range("ам╤у!S2").Value
    ToggleButton1 = range("ам╤у!U2").Value
    TextBox1.Value = range("╔Ф╘Ж!B1")
    Scrollbar_justUpdated = 0
    Call updateScrollbar
     init = False
End Sub

Public Sub updateScrollbar()
    Scrollbar_justUpdated = 1
    ScrollBar1.Max = 100
    ScrollBar1.SmallChange = 1
    ScrollBar1.LargeChange = 10
If Selection.Count >= 2 Then
    'scroll bar
    Dim sum As Double
    sum = 0
    For Each cell In Selection
        If IsNumeric(cell.Value) Then
            sum = sum + cell.Value2
        End If
    Next cell
    If sum > 0 Then
        ScrollBar1.Value = (Selection(1).Value / sum) * 100
        'MsgBox ScrollBar1.Value
    End If
End If

End Sub
