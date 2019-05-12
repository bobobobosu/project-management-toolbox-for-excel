Attribute VB_Name = "Module28"

Sub CopyFired()
On Error Resume Next
range("ам╤у!C2").Value2 = AddressEx(Selection)
Selection.Copy
End Sub

Sub CutFired()
On Error Resume Next
range("ам╤у!C2").Value2 = AddressEx(Selection)
Selection.Cut
End Sub

Sub Paste_Formula_Auto_Fast()
    Dim tocopy As range
    Set tocopy = range(range("ам╤у!C2").Value2)
    Dim selected As range
    Set selected = Selection
    
  Dim rng2 As range
  Dim C As range
Dim values() As Variant
  For Each C In selected
    ' Add cells to rng2 if they exceed 10
    If C.HasFormula = False Then
        If C.Value2 <> vbNullString Then
            If (Not Not values) <> 0 Then
                ' Array has been initialized, so you're good to go.
                ReDim Preserve values(UBound(values) + 1)
            Else
                ' Array has NOT been initialized
                ReDim values(0)
            End If
            values(UBound(values)) = C.Value2
            If Not rng2 Is Nothing Then
            ' Add the 2nd, 3rd, 4th etc cell to our new range, rng2
            ' this is the most common outcome so place it first in the IF test (faster coding)
                Set rng2 = Union(rng2, C)
            Else
            ' the first valid cell becomes rng2
                Set rng2 = C
            End If
        End If
    End If
  Next



   
    If tocopy.HasArray Then
        For Each sArea In Selection.Areas
            sArea.formula = tocopy.formula
            On Error Resume Next
            sArea.FormulaArray = sArea.formula
        Next sArea
    Else
        selected.formula = tocopy.formula
    End If
    
    
    Count = 0
    If Not (rng2 Is Nothing) Then
        For Each textCell In rng2.Cells
            textCell.Value2 = values(Count)
            Count = Count + 1
        Next textCell
    End If
End Sub

Sub Paste_Formula_Ignore_Text_Auto()
    Dim tocopy As range
    Set tocopy = range(range("ам╤у!C2").Value2)
    Dim selected As range
    Set selected = Selection
    
    For Each sArea In selected.Areas
        For Each cell In sArea
                formulaTest = False
                blankTest = False
                formulaTest = cell.HasFormula
                On Error Resume Next
                blankTest = (cell.Value2 = vbNullString)
                If formulaTest Or blankTest Then
                    cell.PasteSpecial Paste:=xlPasteFormulas
'                    If tocopy.HasArray Then
'                        cell.Formula = tocopy.Formula
'                        cell.FormulaArray = cell.Formula
'                    Else
'                        cell.Formula = tocopy.Formula
'                    End If
                End If
        Next cell
    Next sArea
End Sub
Sub Paste_Value_Transpose_Click()
    Dim FirstCell As String
    FirstCell = ""
    For Each cell In Selection
                If FirstCell = "" Then
                    FirstCell = cell.address
                    Selection.PasteSpecial Paste:=xlValues, Transpose:=True
                End If
            Next cell
    
End Sub
Sub Paste_Value_Click()
Attribute Paste_Value_Click.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim FirstCell As String
    FirstCell = ""
    For Each cell In Selection
                If FirstCell = "" Then
                    FirstCell = cell.address
                    Selection.PasteSpecial Paste:=xlValues, Transpose:=False
                End If
            Next cell
    
End Sub

Sub Paste_Selection_Vertical()
    Dim LastSelRow As Integer
    Dim LastSelCol As Integer
    Dim LastSel As Variant
    For Each cell In Selection
        LastSelRow = cell.Row
        LastSelCol = cell.Column
        LastSel = cell.address
    Next cell
    Dim mRow As Integer
    mRow = 0
    For Each cell In Selection
        If cell.address <> LastSel Then
            ActiveSheet.Cells(LastSelRow + mRow, LastSelCol).Value = cell.Value
            mRow = mRow + 1
        End If
    Next cell
    
End Sub

Sub ClearMap()
    range("╩Ы╜х╧о!CG1:DL500").Clear
End Sub

Sub SelectMap()
    range("╩Ы╜х╧о!CG1:DL500").Select
End Sub
Sub Clear()
    Dim FirstCell As String
    FirstCell = ""
    For Each cell In range(Evaluate("INDIRECT(""$AA$1"")"))
        cell.Value = ""
    Next cell
    Call CalculateRange5
    
End Sub
Sub Paste_Input_Click()
Attribute Paste_Input_Click.VB_ProcData.VB_Invoke_Func = "t\n14"
'    FormShow UserForm2, ActiveCell
    Dim copySelection As range
    Set copySelection = Selection
    Dim FirstCellAddress As String
    For Each cell In Selection
        FirstCellAddress = cell.address
        Exit For
    Next cell

    If range(FirstCellAddress).NumberFormat = "m/d/yyyy" Or range(FirstCellAddress).NumberFormat = "h:mm:ss;@" Or range(FirstCellAddress).NumberFormat = "m/d/yy h:mm;@" Then
        Dim FirstCell2 As Variant
        'FirstCell2 = InputBox("Date Value", "Please Enter Date Value", Format(Range(FirstCellAddress).Value2, "m/d/yy"))
        UserForm1.Show 'vbModeless
        FirstCell2 = Format(range("ам╤у!I2").Value2, "m/d/yy")
        Dim FirstCell3 As Variant
        FirstCell3 = InputBox("Time Value", "Please Enter Time Value", Format(range(FirstCellAddress).Value2, "h:mm:ss;@"))
        
        If FirstCell2 <> vbNullString Then
            For Each selected In copySelection
                Evaluate("ам╤у!$D$2").Value2 = dateValue(FirstCell2) + TimeValue(Format(selected.Value2, "h:mm:ss;@"))
                Evaluate("ам╤у!$D$2").Copy
                selected.Select
                selected.PasteSpecial Paste:=xlValues, Transpose:=False
            Next selected
'            Evaluate("ам╤у!$D$2").Value2 = dateValue(FirstCell2) + TimeValue(Format(selected.Value2, "h:mm:ss;@"))
'            Evaluate("ам╤у!$D$2").Copy
'            copySelection.PasteSpecial Paste:=xlValues, Transpose:=False
        End If
        
        If FirstCell3 <> vbNullString Then
            For Each selected In copySelection
                Evaluate("ам╤у!$D$2").Value2 = dateValue(Format(selected.Value2, "m/d/yy")) + TimeValue(FirstCell3)
                Evaluate("ам╤у!$D$2").Copy
                selected.Select
                selected.PasteSpecial Paste:=xlValues, Transpose:=False
            Next selected
'                Evaluate("ам╤у!$D$2").Value2 = dateValue(Format(selected.Value2, "m/d/yy")) + TimeValue(FirstCell3)
'                Evaluate("ам╤у!$D$2").Copy
'                copySelection.PasteSpecial Paste:=xlValues, Transpose:=False
        End If
    Else
        
        Dim FirstCell As Variant
        FirstCell = f_ListSearch.Get_Input 'InputBox("Change Value", "Please Enter Value", Range(FirstCellAddress).Value2)
        If FirstCell <> vbNullString Then
            If IsNumeric(FirstCell) Then
                Evaluate("ам╤у!$D$2").Value2 = Evaluate("=" + CStr(FirstCell))
                Evaluate("ам╤у!$D$2").Copy
                Debug.Print FirstCell
            Else
                Evaluate("ам╤у!$D$2").Value2 = FirstCell
                Evaluate("ам╤у!$D$2").Copy
            End If
'            For Each selected In copySelection
'                    'selected.Select
'                    selected.PasteSpecial Paste:=xlValues, Transpose:=False
'            Next selected
            copySelection.PasteSpecial Paste:=xlValues, Transpose:=False
        Else
'            For Each selected In copySelection
'                    selected.Value = vbNullString
'            Next selected
        End If
    End If
    
    
    'Special Processing
    For Each cell In Selection
        If FirstCellAddress = "" Then
            FirstCellAddress = cell.address
        End If
        If range(use_Structured(cell, 5)).address = cell.address Then
            range(use_Structured(cell, 2)).Value2 = range(use_Structured(cell, 5)).Value2 - range(use_Structured(cell, 4)).Value2
         End If
    Next cell

    
    
End Sub



Sub Paste_Formula_Click()
    selected.Cells(1).PasteSpecial xlFormulas
    myFormula = selected.Cells(1).formula
    selected.formula = myFormula
End Sub



Sub Paste_Formula_Ignore_Text_Click()
    Dim myFormula As String
    
    Dim selected As range
    Set selected = Selection
    
    On Error Resume Next
    For Each cell In selected
        If cell.HasArray Then
            If cell.HasFormula = True Or cell.Value2 = "" Then
                        If myFormula = "" Then
                            cell.PasteSpecial xlFormulas
                            myFormula = cell.formula
                        End If
            End If
        End If
    Next cell
    
'    selected.Cells(1).PasteSpecial xlFormulas
'    myFormula = selected.Cells(1).Formula
'    selected.Formula = myFormula
    On Error GoTo 0
End Sub
Sub Paste_FormulaArray_Ignore_Text_Click()
    Dim myFormula As String
    
    Dim selected As range
    Set selected = Selection

    On Error Resume Next
    selected(1).PasteSpecial xlFormulas
    myFormula = cell.formula
    For Each cell In selected
        If cell.HasFormula = True Or cell.Value2 = "" Then
                    myFormula = cell.formula
                    cell.FormulaArray = myFormula
                    Debug.Print cell.Value
        End If

    Next cell


'    selected.Cells(1).PasteSpecial xlFormulas
'    myFormula = selected.Cells(1).Formula
'    selected.FormulaArray = myFormula
    
    On Error GoTo 0
End Sub
Function dateCheck(dateValue As Date) As Boolean
    If dateValue.NumberFormat <> "m/d/yy h:mm;@" Then
=
        dateCheck = True
    End If
End Function

Function test_inters(rng1 As Variant, rng2 As Variant)
    If (rng1.Parent.name = rng2.Parent.name) Then
        Dim ints As range
        Set ints = Application.Intersect(rng1, rng2)
        If (Not (ints Is Nothing)) Then
            ' Do your job
            test_inters = True
            Exit Function
        End If
    End If
    test_inters = False
End Function
