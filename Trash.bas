Attribute VB_Name = "Trash"
Sub wtrszhfsydg()
    Dim gggggg As Range
    Set gggggg = Range("D6:D14")
    Dim hhh(1 To 4, 1 To 1) As Variant
    hhh(1, 1) = "=1"
     hhh(2, 1) = "=1"
      hhh(3, 1) = "=4"
       hhh(4, 1) = "=1"
        gggggg.Formula = hhh
End Sub
Sub Array2Range()
   Dim directory(1 To 10, 1 To 2) As Variant
   Dim rng As Range, i As Long, j As Long

   For i = 1 To 2
      For j = 1 To 10
         directory(j, i) = "=Now()"
         
      Next j
   Next i

   Set rng = Range("X1:Y10")

   rng.FormulaArray = directory

End Sub
Sub bytg()
MacroFinished ("Hello")
End Sub
Sub gg()
MsgBox ThisWorkbook2.Name
End Sub
