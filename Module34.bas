Attribute VB_Name = "Module34"
Public waitNext As Long
Public Sub CallDoEvent()
    If ActiveSheet.Name = "交易" Then
        'your code here+
        'DoEvents
         'Range(Range("交易!C1").Value2).Calculate
        Range(Range("交易!C1").Value2).Calculate
        Range("交易!K2").Calculate
        Range("交易!I2").Calculate
        Range("交易!M2").Calculate
        Range(Range("交易!C1").Value2).Calculate
         'MsgBox Range("交易!C1").Value2
         'Range(CStr(Range(Range("交易!S2").Value2).Row - 1) + ":" + CStr(Range(Range("交易!S2").Value2).Row + 1)).Calculate
         'Range(Range("交易!C1").Value2).Calculate
         ProgressMessage = CStr(Range("交易!I2").Value2) + " 現進度: " + Format(Range("交易!M2").Value2, "0%")
         'MacroFinished (ProgressMessage)
'        If waitNext < 30 Then
'            waitNext = waitNext + 1
'        Else
'            waitNext = 0
'            MacroFinished (ProgressMessage)
'        End If
         
        Application.StatusBar = ProgressMessage
    End If
    
    If (Range("交易!A1").Value) = True Then
        If Range("趨勢!K2").Value = 1 Then
            Application.OnTime Now + TimeValue("00:00:10"), "CallDoEvent"
        End If
    End If
End Sub

Sub tttt()
         ProgressMessage = "ho"
         MacroFinished ("")
End Sub
