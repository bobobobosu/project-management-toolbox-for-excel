Attribute VB_Name = "Module5"
Function eval(r As Range) As Variant
    eval = Evaluate(r.Value)
End Function
Function evals(r As String) As Variant
    Debug.Print r
    evals = Evaluate(r)
End Function

