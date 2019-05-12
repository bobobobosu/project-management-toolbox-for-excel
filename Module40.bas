Attribute VB_Name = "Module40"
Private Sub ttt()
    Dim s As String, sFileName As String, wsh As Object, threadFileName As String
    thread = 1
    parallelKey = 5000
    macroName = "test"
    'Save a copy of the Excel workbook
    threadFileName = ActiveWorkbookpath & "\" & parallelKey & "_" & thread & ".xlsb"
    'threadFileName = ActiveWorkbookpath & "\" & "t.xlsb"
    'threadFileName = ActiveWorkbookpath & "\" & CStr(thread) & ".xlsb"
    threadFileName = ActiveWorkbookpath & "\" & "hh.xlsb"
    'Call ActiveWorkbook.SaveCopyAs(threadFileName)
    openedXls = ActiveWorkbookpath & "\" & CStr(thread) & ".xlsb"
    openedXls = threadFileName
    'Save the VBscript
    s = "Set objExcel = GetObject(""" & openedXls & """):"
    s = s & "With objExcel:"
 '   s = s & ".Application.Visible = False:"
    's = s & ".Application.Workbooks(1).Windows(1).Visible = False = xlMinimized:"
    s = s & ".Application.Workbooks.Open(""" & threadFileName & """):"
    
    s = s & ".Application.Run """ & "hh.xlsb!" & macroName & """ :"
'    s = s & ".Application.Run """ & parallelKey & "_" & thread & ".xlsb!" & macroName & """ , """ & _
'        ActiveWorkbook.Name & """," & _
'        subSeqFrom & "," & _
'        subSeqTo & ":"
'    s = s & ".Application.Run """ & thread & ".xlsb!" & macroName & """ , """ & _
'    ActiveWorkbook.Name & """," & _
'    subSeqFrom & "," & _
'    subSeqTo & ":"
'    s = s & ".Application.Run """ & "t.xlsb!" & macroName & """ , """ & _
'    ActiveWorkbook.Name & """," & _
'    subSeqFrom & "," & _
'    subSeqTo & ":"
    
'    s = s & ".Application.ActiveWorkbook.Close False:"
   ' s = s & ".Application.Quit:"
    s = s & "End With:"
'    s = s & "Set oXL = GetObject(""" & threadFileName & """):"
'    s = s & "On Error Resume Next" & vbCrLf
'    s = s & "oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value = 1" & vbCrLf
'    s = s & "Do Until CLng(Replace(oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value,""="","""")) = 1" & vbCrLf
'    s = s & "If Err.Number <> 0 Then Exit Do" & vbCrLf
'    s = s & "WScript.Sleep(100)" & vbCrLf
'    s = s & "oXL.Application.Workbooks(""" & Application.ActiveWorkbook.Name & """).Names(""S" & parallelKey & "_" & thread & """).Value = 1" & vbCrLf
'    s = s & "Loop" & vbCrLf
'    s = s & "Set oXL = Nothing"
    'Save the VBscript file
    sFileName = ActiveWorkbookpath & "\" & parallelKey & "_" & thread & ".vbs"
    Open sFileName For Output As #1
    Print #1, s
    Close #1
    'Execute the VBscript file asynchronously
    Set wsh = VBA.CreateObject("WScript.Shell")
    wsh.Run """" & sFileName & """"
    Set wsh = Nothing
    
End Sub

