Attribute VB_Name = "Module1"
'
'Sub solver_PERT_SU_MIN()
''
'' solver_PERT_SU_MIN Macro
''
'
''
'
'    Dim MAX1_or_MIN0 As Integer
'    Dim Table56 As ListObject
'    Set Table56 = ActiveSheet.ListObjects("Table56")
'
'
'    'For Each Element In Evaluate("B18:B20")
'    'Range("$F$2").Value = Element
'    For MAX1_or_MIN0 = 0 To 1
'
'
'    SolverReset
'    SolverOk SetCell:="$H$2", MaxMinVal:=MAX1_or_MIN0 + 1, ValueOf:=0, ByChange:=Evaluate("Table52[Activation]"), _
'        Engine:=2, EngineDesc:="GRG Nonlinear"
'    'SolverAdd CellRef:=Evaluate("Table52[Activation]"), Relation:=5, FormulaText:="binary"
'    'SolverAdd CellRef:=Evaluate("Table56[Flow Detect]"), Relation:=2, FormulaText:=Evaluate("Table56[Flow Settings]")
'
'    SolverAdd CellRef:=Evaluate("Table52[Activation]"), Relation:=1, FormulaText:="1"
'    SolverAdd CellRef:=Evaluate("Table52[Activation]"), Relation:=3, FormulaText:="-1"
'    SolverAdd CellRef:=Evaluate("Table52[Activation]"), Relation:=4, FormulaText:="integer"
'
'    SolverAdd CellRef:="$G$2", Relation:=2, FormulaText:="5"
'    SolverSolve UserFinish:=True
'    SolverFinish KeepFinal:=2, ReportArray:=1
'
'
'
'
'    Range("$H$2").Select
'    Selection.Copy
'
'    Dim cell2copy As Range
'
'    If MAX1_or_MIN0 = 0 Then
'        Set cell2copy = Evaluate("INDEX(Table56,MATCH($F$2,Table56[Elements],0),MATCH(""min_SU_MIN"",Table56[#Headers],0))")
'    Else
'        Set cell2copy = Evaluate("INDEX(Table56,MATCH($F$2,Table56[Elements],0),MATCH(""max_SU_MIN"",Table56[#Headers],0))")
'    End If
'    cell2copy.Select
'
'
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'
'
'
'    Next
'    'Next
'
'
'End Sub
