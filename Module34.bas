Attribute VB_Name = "Module34"
Public waitNext As Long
Public Sub CallDoEvent()
    If ActiveSheet.Name = "���" Then
        'your code here+
        'DoEvents
         'Range(Range("���!C1").Value2).Calculate
        Range(Range("���!C1").Value2).Calculate
        Range("���!K2").Calculate
        Range("���!I2").Calculate
        Range("���!M2").Calculate
        Range(Range("���!C1").Value2).Calculate
         'MsgBox Range("���!C1").Value2
         'Range(CStr(Range(Range("���!S2").Value2).Row - 1) + ":" + CStr(Range(Range("���!S2").Value2).Row + 1)).Calculate
         'Range(Range("���!C1").Value2).Calculate
         ProgressMessage = CStr(Range("���!I2").Value2) + " �{�i��: " + Format(Range("���!M2").Value2, "0%")
         'MacroFinished (ProgressMessage)
'        If waitNext < 30 Then
'            waitNext = waitNext + 1
'        Else
'            waitNext = 0
'            MacroFinished (ProgressMessage)
'        End If
         
        Application.StatusBar = ProgressMessage
    End If
    
    If (Range("���!A1").Value) = True Then
        If Range("�Ͷ�!K2").Value = 1 Then
            Application.OnTime Now + TimeValue("00:00:10"), "CallDoEvent"
        End If
    End If
End Sub

Sub tttt()
         ProgressMessage = "ho"
         MacroFinished ("")
End Sub
