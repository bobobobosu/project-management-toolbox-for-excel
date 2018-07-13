Attribute VB_Name = "Module13"
Sub ScrollBar4_Change()

End Sub
Sub ScrollBar3_Change()
    Set SB = ActiveSheet.Shapes("Scroll Bar 3").ControlFormat
    'ActiveCell.VALUE = SB.VALUE / 100
    For Each cell In Selection
        cell.Value = SB.Value / 100
    Next cell

End Sub
