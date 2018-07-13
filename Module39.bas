Attribute VB_Name = "Module39"
Public Function DINDIRECT(sName As String) As Range
     Dim nName As Name

     On Error Resume Next
          Set nName = ActiveWorkbook.Names(sName)
          Set nName = ActiveSheet.Names(sName)
     On Error GoTo 0

     If Not nName Is Nothing Then
          Set DINDIRECT = nName.RefersToRange
     Else
          DINDIRECT = CVErr(xlErrName)
    End If
End Function


Sub RunForVBAAndWait()
    Dim parallelClass As Parallel
    Set parallelClass = New Parallel
    'The line below will not block macro execution
    Call parallelClass.ParallelAsyncInvoke("RunAsyncVBA", 1, 1000)
    'Do other operations here
    '....
    if parallelClass.IsAsyncRunning then ... 'Check if Async thread is still running
    '....
    'Now let's wait until the thread has finished
    parallelClass.AsyncThreadJoin
End Sub
Sub RunAsyncVBA(workbookName As String, seqFrom As Long, seqTo As Long)
    For i = seqFrom To seqTo
        x = seqFrom / seqTo
    Next i
End Sub
