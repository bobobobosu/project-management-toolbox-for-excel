Attribute VB_Name = "Module23"

Sub AutoUpdate_Click()
    'Call AutoCalculate
    
      Call CallDoEvent
    
End Sub

Private Sub AutoCalculate()
    If (Range("A1").Value) = True Then
        Range(Evaluate("INDIRECT(""$N$2"")")).Calculate
        'Call CreateCalendar
        If Range("ам╤у!K2").Value = 1 Then
            
            Application.OnTime Now + TimeValue("00:01:00"), "AutoCalculate"
        End If
    End If
End Sub
