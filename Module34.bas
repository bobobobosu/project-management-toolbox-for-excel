Attribute VB_Name = "Module34"
Public waitNext As Long
Public Sub CallDoEvent()
    If ActiveSheet.name = "���" Then
        'your code here+
        'DoEvents
         'Range(Range("���!C1").Value2).Calculate
        range(range("���!C1").Value2).Calculate
        range("���!K2").Calculate
        range("���!I2").Calculate
        range("���!M2").Calculate
        range(range("���!C1").Value2).Calculate
         'MsgBox Range("���!C1").Value2
         'Range(CStr(Range(Range("���!S2").Value2).Row - 1) + ":" + CStr(Range(Range("���!S2").Value2).Row + 1)).Calculate
         'Range(Range("���!C1").Value2).Calculate
         ProgressMessage = CStr(range("���!I2").Value2) + " �{�i��: " + Format(range("���!M2").Value2, "0%")
         'MacroFinished (ProgressMessage)
'        If waitNext < 30 Then
'            waitNext = waitNext + 1
'        Else
'            waitNext = 0
'            MacroFinished (ProgressMessage)
'        End If
         
        Application.StatusBar = ProgressMessage
    End If
    
    If (range("���!A1").Value) = True Then
        If range("�Ͷ�!K2").Value = 1 Then
            Application.OnTime Now + TimeValue("00:00:10"), "CallDoEvent"
        End If
    End If
End Sub
