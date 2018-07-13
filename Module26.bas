Attribute VB_Name = "Module26"
Sub TimesMinusOne_Click()
    For Each cell In Selection
        cell.Value2 = cell.Value2 * -1
    Next cell
End Sub


Sub sTimesMinusOne_Click()
    For Each cell In Selection
        cell.Value2 = "(" + cell.Value2 + ")" + " *(-1)"
    Next cell
End Sub

Sub sPlusOne_Click()
    For Each cell In Selection
        cell.Value2 = "(" + cell.Value2 + ")" + " +(1)"
    Next cell
End Sub


Sub sMinusOne_Click()
    For Each cell In Selection
        cell.Value2 = "(" + cell.Value2 + ")" + " +(-1)"
    Next cell
End Sub


Sub sTimesZeroClick()
    For Each cell In Selection
        cell.Value2 = "(" + cell.Value2 + ")" + " *(0)"
    Next cell
End Sub
