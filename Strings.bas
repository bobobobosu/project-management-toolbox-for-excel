Attribute VB_Name = "Strings"
Function SplitText(s As String, deli As String)
    SplitText = Application.Transpose(Split(s, deli))
End Function

Function SplitTextHor(s As String, deli As String)
    SplitTextHor = Split(s, deli)
End Function

