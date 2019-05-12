VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parallel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author: Tomasz Kacprowicz
'Site:blog.tkacprow.pl
'Variables
Private cThreads As Long
Private parallelKey As String
'Constructor
Private Sub Class_Initialize()
    cThreads = 4
End Sub
'Destructor
Private Sub Class_Terminate()
    DropSemaphore
    DropSemaphores
End Sub
'Setters
Sub SetThreads(threads As Long)
    If threads > 0 Then cThreads = threads
End Sub
'Getters
Function GetThreads()
    GetThreads = cThreads
End Function
'----For Threads----
Function ParallelFor(macroName As String) As String
    'Debug.Print CStr(seqFrom) & " TO " & CStr(seqTo)
    'Variables
    Dim thread As Long
    parallelKey = "CORE_" ' CStr(CInt(Math.Rnd() * 10000)) 'Key for thread joining
    'CreateSemaphores
    'Create For Threads
    
    'Filter Ranges to Calculate
    Dim r As range
    Dim calList As range
    Dim mseqFrom As Long
    Dim mseqTo As Long
    Set calList = getCalculationList()
    Set r = calList ' Filterll(calList, 0, -1, (calList.Columns.count > 1 And calList.Areas.count = 1))
    
    Dim columnMode As Long
    If (r.Columns.Count > 1) Then
        'Column Mode
        seqFrom = 1
        seqTo = r.Columns.Count
        columnMode = 1
    Else
        'Cell Mode
        seqFrom = 1
        seqTo = r.Cells.Count
        columnMode = 0
    End If
    
    
    
    
    
    For i = 1 To 12
        range("ам╤у!L" & CStr((1 + i))).Value2 = ""
    Next i
    If cThreads < (seqTo - seqFrom + 1) Then
        For thread = 1 To cThreads
            mseqFrom = seqFrom + CLng(CDbl(seqTo - seqFrom) * (thread - 1) / cThreads)
            mseqTo = seqFrom + CLng(CDbl(seqTo - seqFrom) * thread / cThreads)
            f = AddressEx(calList)
            Set r = Filterll(calList, mseqFrom, mseqTo, columnMode)
            range("ам╤у!L" & CStr((2 + thread))).Value2 = AddressEx(r)
            ForThread macroName, thread, _
                mseqFrom, _
                mseqTo
        Next thread
    Else
        For thread = seqFrom To seqTo
            mseqFrom = thread
            mseqTo = thread
            Set r = Filterll(calList, mseqFrom, mseqTo, columnMode)
            range("ам╤у!L" & CStr((2 + thread))).Value2 = AddressEx(r)
            ForThread macroName, thread, _
                thread, _
                thread
        Next thread
        
        For thread = seqTo + 1 To cThreads
            ForThread macroName, thread, _
                0, _
                0
        Next thread
    End If





    'Join For Threads
    'ForThreadJoin
End Function
Private Sub CreateSemaphores()
    Dim i As Long
    For i = 1 To cThreads
        Call ActiveWorkbook.names.Add("S" & parallelKey & "_" & i, "=0")
    Next i
End Sub
Private Sub DropSemaphores()
    On Error Resume Next
    Dim i As Long
    For i = 1 To cThreads
        Call ActiveWorkbook.names("S" & parallelKey & "_" & i).Delete
    Next i
End Sub
Private Function CountSemaphores()
    Dim i As Long, C As Long
    C = 0
    For i = 1 To cThreads
        If CLng(Replace(ActiveWorkbook.names("S" & parallelKey & "_" & i).Value, "=", "")) > 0 Then C = C + 1
    Next i
    CountSemaphores = C
End Function
Private Sub ForThread(macroName As String, thread As Long, subSeqFrom As Long, subSeqTo As Long)
    Dim s As String, sFileName As String, wsh As Object, threadFileName As String
    'Save a copy of the Excel workbook
'    threadFileName = ActiveWorkbook.Path & "\" & parallelKey & "_" & thread & ".xlsb"
'    Call ActiveWorkbook.SaveCopyAs(threadFileName)
    
    openedXls = ActiveWorkbook.path & "\" & CStr(thread) + ".xlsb"
    threadFileName = ActiveWorkbook.path & "\" & CStr(thread) + ".xlsb"

    'Save the VBscript
    s = "Set objExcel = GetObject(""" & openedXls & """):"
    s = s & "With objExcel:"
'    s = s & ".Application.Visible = False:"
'    s = s & ".Application.Workbooks(1).Windows(1).Visible = False = xlMinimized:"
'    s = s & ".Application.Workbooks.Open(""" & threadFileName & """):"
    s = s & ".Application.Run """ & thread & ".xlsb!" & macroName & """ , """ & _
        ActiveWorkbook.name & """," & _
        subSeqFrom & "," & _
        subSeqTo & ":"
'    s = s & ".Application.ActiveWorkbook.Close False:"
'    s = s & ".Application.Quit:"
    s = s & "End With:"
'    s = s & "Set oXL = GetObject(""" & Application.ActiveWorkbook.FullName & """):"
'    s = s & "On Error Resume Next" & vbCrLf
'    s = s & "oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value = 1" & vbCrLf
'    s = s & "Do Until CLng(Replace(oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value,""="","""")) = 1" & vbCrLf
'    s = s & "If Err.Number <> 0 Then Exit Do" & vbCrLf
'    s = s & "WScript.Sleep(100)" & vbCrLf
'    s = s & "oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value = 1" & vbCrLf
'    s = s & "Loop" & vbCrLf
'    s = s & "Set oXL = Nothing"
    'Save the VBscript file
    sFileName = ActiveWorkbook.path & "\" & parallelKey & "_" & thread & ".vbs"
    Open sFileName For Output As #1
    Print #1, s
    Close #1
    'Execute the VBscript file asynchronously
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run """" & sFileName & """"
    Set wsh = Nothing
    
End Sub




Private Sub ForThreadJoin()
    Do Until False
        DoEvents
        If CountSemaphores >= cThreads Then
            Sleep 100
            'Kill ActiveWorkbook.Path & "\" & parallelKey & "*.vbs"
            'Kill ActiveWorkbook.Path & "\" & parallelKey & "*.xls"
            'Kill ActiveWorkbook.Path & "\" & "1.xlsb"
            DropSemaphores
            Exit Sub
        End If
        Application.StatusBar = CountSemaphores
        Sleep 10
    Loop
End Sub
'----Async threads----
Sub ParallelAsyncInvoke(macroName As String, Optional arg1, Optional arg2, Optional arg3, Optional arg4, Optional arg5, Optional arg6)
    'Variables
    Dim thread As Long, args As String
    parallelKey = CStr(CInt(Math.Rnd() * 10000)) 'Key for thread joining
    CreateSemaphore
    'Create Thread
    If Not (IsMissing(arg1)) Then args = args & ", " & GetArg(arg1)
    If Not (IsMissing(arg2)) Then args = args & ", " & GetArg(arg2)
    If Not (IsMissing(arg3)) Then args = args & ", " & GetArg(arg3)
    If Not (IsMissing(arg4)) Then args = args & ", " & GetArg(arg4)
    If Not (IsMissing(arg5)) Then args = args & ", " & GetArg(arg5)
    If Not (IsMissing(arg6)) Then args = args & ", " & GetArg(arg6)
    AsyncThread macroName, args
End Sub
Private Sub CreateSemaphore()
    Call ActiveWorkbook.names.Add("I" & parallelKey & "I", "=0")
End Sub
Private Sub DropSemaphore()
    On Error Resume Next
    Call ActiveWorkbook.names("I" & parallelKey & "I").Delete
End Sub
Private Sub AsyncThread(macroName As String, args As String)
    Dim s As String, sFileName As String, wsh As Object, threadFileName As String
    'Save a copy of the Excel workbook
    threadFileName = ActiveWorkbook.path & "\" & parallelKey & ".xls"
    Call ActiveWorkbook.SaveCopyAs(threadFileName)
    'Save the VBscript
    s = "Set objExcel = CreateObject(""Excel.Application""):"
    s = s & "With objExcel:"
    s = s & ".Application.Visible = False:"
    s = s & ".Workbooks.Open(""" & threadFileName & """):"
    s = s & ".Application.Run """ & parallelKey & ".xls!" & macroName & """,""" & ActiveWorkbook.name & """ " & args & ":"
    s = s & ".ActiveWorkbook.Close False:"
    s = s & ".Application.Quit:"
    s = s & "End With:"
    s = s & "Set oXL = GetObject(, ""Excel.Application""):"
    s = s & "On Error Resume Next" & vbCrLf
    s = s & "oXL.Workbooks(""" & ActiveWorkbook.name & """).Names(""I" & parallelKey & "I"").Value = 1" & vbCrLf
    s = s & "Do Until CLng(Replace(oXL.Workbooks(""" & ActiveWorkbook.name & """).Names(""I" & parallelKey & "I"").Value,""="","""")) = 1" & vbCrLf
    s = s & "If Err.Number <> 0 Then Exit Do" & vbCrLf
    s = s & "WScript.Sleep(100)" & vbCrLf
    s = s & "oXL.Workbooks(""" & ActiveWorkbook.name & """).Names(""I" & parallelKey & "I"").Value = 1" & vbCrLf
    s = s & "Loop" & vbCrLf
    s = s & "Set oXL = Nothing" & vbCrLf
    'Save the VBscript file
    sFileName = ActiveWorkbook.path & "\" & parallelKey & ".vbs"
    Open sFileName For Output As #1
    Print #1, s
    Close #1
    'Execute the VBscript file asynchronously
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run """" & sFileName & """"
    Set wsh = Nothing
End Sub


Public Sub AsyncThreadJoin()
    Do Until False
        DoEvents
        If CountSemaphore >= 1 Then
            Sleep 100
            Kill ActiveWorkbook.path & "\" & parallelKey & "*.vbs"
            Kill ActiveWorkbook.path & "\" & parallelKey & "*.xls"
            DropSemaphore
            Exit Sub
        End If
        Sleep 10
    Loop
End Sub
Public Function IsAsyncRunning() As Boolean
    If CountSemaphore < 1 Then
        IsAsyncRunning = True
    Else
        IsAsyncRunning = False
    End If
End Function
Private Function CountSemaphore()
    Dim C As Long
    C = 0
    If CLng(Replace(ActiveWorkbook.names("I" & parallelKey & "I").Value, "=", "")) > 0 Then C = C + 1
    CountSemaphore = C
End Function
Private Function GetArg(val) As String
    If VarType(val) = vbString Then
        GetArg = """" & val & """"
    Else
        GetArg = "" & val
    End If
End Function



Private Sub jythgfjh(macroName As String, thread As Long, subSeqFrom As Long, subSeqTo As Long)
    Dim s As String, sFileName As String, wsh As Object, threadFileName As String
    'Save a copy of the Excel workbook
'    threadFileName = ActiveWorkbook.Path & "\" & parallelKey & "_" & thread & ".xlsb"
'    Call ActiveWorkbook.SaveCopyAs(threadFileName)
    
'    openedXls = ActiveWorkbook.Path & "\" & CStr(thread) + ".xlsb"
'    threadFileName = ActiveWorkbook.Path & "\" & CStr(thread) + ".xlsb"
'
'    'Save the VBscript
'    s = "Set objExcel = GetObject(""" & openedXls & """):"
'    s = s & "With objExcel:"
''    s = s & ".Application.Visible = False:"
''    s = s & ".Application.Workbooks(1).Windows(1).Visible = False = xlMinimized:"
''    s = s & ".Application.Workbooks.Open(""" & threadFileName & """):"
'    s = s & ".Application.Run """ & thread & ".xlsb!" & macroName & """ , """ & _
'        ActiveWorkbook.Name & """," & _
'        subSeqFrom & "," & _
'        subSeqTo & ":"
''    s = s & ".Application.ActiveWorkbook.Close False:"
''    s = s & ".Application.Quit:"
'    s = s & "End With:"
''    s = s & "Set oXL = GetObject(""" & Application.ActiveWorkbook.FullName & """):"
''    s = s & "On Error Resume Next" & vbCrLf
''    s = s & "oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value = 1" & vbCrLf
''    s = s & "Do Until CLng(Replace(oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value,""="","""")) = 1" & vbCrLf
''    s = s & "If Err.Number <> 0 Then Exit Do" & vbCrLf
''    s = s & "WScript.Sleep(100)" & vbCrLf
''    s = s & "oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value = 1" & vbCrLf
''    s = s & "Loop" & vbCrLf
''    s = s & "Set oXL = Nothing"
    'Save the VBscript file
    sFileName = ActiveWorkbook.path & "\" & "ff.vbs"
    Open sFileName For Output As #1
    Print #1, s
    Close #1
    'Execute the VBscript file asynchronously
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run """" & sFileName & """"
    Set wsh = Nothing
    
End Sub
