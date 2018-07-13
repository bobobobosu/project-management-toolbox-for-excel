VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Range("ам╤у!K2").Value = 0
End Sub

Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
    
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal target As Range)
    On Error Resume Next
    Range("ам╤у!G2").Value2 = Range("ам╤у!E2").Value2
    On Error Resume Next
    Range("ам╤у!E2").Value2 = AddressEx(Selection(1))
    On Error Resume Next
    Range("ам╤у!H2").Value2 = Range("ам╤у!F2").Value2
    On Error Resume Next
    Range("ам╤у!F2").Value2 = Selection(1).Value2
    On Error Resume Next
    Range("ам╤у!J2").Value2 = AddressEx(Selection)
    Dim mselection As Variant

    If Application.CutCopyMode = False Then
        Call RecalculateSelection
        If Not (Selection Is Nothing) Then
            On Error Resume Next
            If Selection.count = 1 Then
                'UserForm2.Hide
                FormShowFixed UserForm2
            Else
                On Error Resume Next
                UserForm2.updateScrollbar
                FormShow UserForm2, ActiveCell
            End If
        End If
    Else
        On Error Resume Next
        UserForm2.updateScrollbar
        FormShow UserForm2, ActiveCell
    End If
'    If Target.Cells.Count > 1 Then Exit Sub
'    Application.ScreenUpdating = False
'    'Clear all cell colours
'    Cells.Interior.ColorIndex = 0
'    With Target
'    'Row and column highlightingfor the active cell
'    .EntireRow.Interior.ColorIndex = 50
'    .EntireColumn.Interior.ColorIndex = 50
'    End With
'    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
End Sub


Private Sub Workbook_Open()
    'first call appropriate procedure
    Application.ErrorCheckingOptions.BackgroundChecking = False
    Application.OnKey "^c", "CopyFired"
    Application.OnKey "^x", "CutFired"
    Range("ам╤у!K2").Value = 1
    Application.Calculation = xlCalculationManual
    Call CallDoEvent
    If AvalibleCoreNum() > 0 Then
        Range("ам╤у!O2").Value2 = 1
    Else
        Range("ам╤у!O2").Value2 = 0
    End If
    
End Sub

