Attribute VB_Name = "Module8"
'Public Const masterFileName  As String = "Z:\My Drive\_Storage\Backup\Documents\RAMDISK_loc\Documents\root\Data\TC.xlsb"
Public Sub RecalculateSelection()

    On Error GoTo ErrorHandler
    If TypeName(Selection) = "Range" Then
        'Detect Calculation Settings
        If Range("ам╤у!K2").Value = 1 Then
            'Application.OnTime Now, "BackgroundCalculate"
            Call BackgroundCalculate
        End If
    End If
    
    Exit Sub
ErrorHandler:
    ' Insert code to handle the error here
    Resume Next
    
End Sub

Public Sub RecalculateActiveSheet()
    ActiveSheet.Calculate
End Sub

Public Function BackgroundCalculate()

    'Default Calculation
    Dim mselection As Range
    Set mselection = Selection
    If Range("ам╤у!P2") = 1 Then
        mselection.Calculate
    End If
    Call llCalculate(mselection)

    
End Function

Public Sub llCalculate(rIN As Range)
    Dim r As Range
    Set r = rIN


    If Range("ам╤у!O2") = 0 Then Exit Sub
    'Isolate firstRow
    Call RefreshCal(r.Rows(1))
    If rIN.Rows.count = 1 Then Exit Sub
    Set r = getSubRange(2, 1, r.Rows.count, r.Columns.count, r)
    'If no core
    If Range("ам╤у!R2").Value2 = 0 Then
        Call RefreshCal(r)
        Exit Sub
    End If
    'Detect ll Stat
    inuse = WorksheetFunction.sum(Range("ам╤у!N" & "2:N" & CStr(2 + Range("ам╤у!R2").Value2)))
        
    'Run ll if idle
    Dim filtered As Range
    Set filtered = Filterll(r)
    'Update List
    If (Not (filtered Is Nothing)) And inuse = 0 Then
        Range("ам╤у!L2").Value2 = AddressEx(filtered)
        Set myr = getRange(Range("ам╤у!L2").Value2)
        If filtered.count > 10 Then
            'Disable Update
            Range("ам╤у!S2") = 0
            Call TestParallelForLoop(filtered)
        Else
            If Not (filtered Is Nothing) Then
                Call RefreshCal(filtered)
            End If
        End If
    End If
End Sub
Sub sertgfds()
    Debug.Print getRange(Range("ам╤у!L2").Value2).Address
End Sub
Public Function getRange(rangeS As String, Optional FileName As String = "") As Range
    Dim r As Range
    Dim remains As String
    remains = rangeS
    If Len(rangeS) > 254 Then
        'get All but Last
        Do Until InStr(remains, ",") = False
            commaposition = InStr(remains, ",")
            ms = Left(remains, commaposition - 1)
    
             If FileName = "" Then
                If (r Is Nothing) Then
                    Set r = Range(ms)
                Else
                    Set r = Union(r, Range(ms))
                End If
            Else
                If (r Is Nothing) Then
                    Set r = GetObject(FileName).Application.Range(ms)
                Else
                    Set r = Union(r, GetObject(FileName).Application.Range(ms))
                End If
            End If
            remains = Right(remains, Len(remains) - commaposition)
        Loop
        'get Last One
         If FileName = "" Then
            If (r Is Nothing) Then
                Set r = Range(remains)
            Else
                Set r = Union(r, Range(remains))
            End If
        Else
            If (r Is Nothing) Then
                Set r = GetObject(FileName).Application.Range(remains)
            Else
                Set r = Union(r, GetObject(FileName).Application.Range(remains))
            End If
        End If
        Set getRange = r
    Else
         If FileName = "" Then
                Set r = Range(remains)
        Else
                Set r = GetObject(FileName).Application.Range(remains)
        End If
    End If
    Set getRange = r
End Function
Public Sub TriggerSync()
    Dim blankR As Range
    Range("ам╤у!Q2").Value2 = AddressEx(Selection)
    Range("ам╤у!L2").Value2 = ""
    Set blankR = Range("A1").Resize(1, Range("ам╤у!R2").Value)
    Range("ам╤у!S2") = 1
    Call TestParallelForLoop(blankR)
End Sub
Public Function TestParallelForLoop(calList As Range)
    Set calList = UnionRange(calList)
    'Clean ll Stat
    Set oXL = GetObject(masterFileName)
    For i = 2 To (2 + Range("ам╤у!R2").Value - 1)
            oXL.Application.Range("ам╤у!N" & i).Value = 1
    Next i
    oXL.Application.StatusBar = CStr(oXL.Application.Range("ам╤у!R2").Value2 - WorksheetFunction.sum(oXL.Application.Range("ам╤у!N" & "2:N" & CStr(2 + oXL.Application.Range("ам╤у!R2").Value2)))) & " in " & CStr(oXL.Application.Range("ам╤у!R2").Value2) & " Completed " & CStr(Now())
    
    'Set ll
    Dim para As Parallel
    Set para = New Parallel
    para.SetThreads Range("ам╤у!R2").Value

    'Run ll
    Call para.ParallelFor("RunForVBA")
    
    'Timer Start
    Call WaitForCompletion(Now())
End Function

Sub RunForVBA(workbookName As String, seqFrom As Long, seqTo As Long)
    Set oXL = GetObject(masterFileName)
    'No Run
    If seqTo = 0 Then
        Do Until oXL.Application.Range("ам╤у!N" & CStr(Replace(ActiveWorkbook.Name, ".xlsb", "") + 1)).Value = 0
            On Error Resume Next
            oXL.Application.Range("ам╤у!N" & CStr(Replace(ActiveWorkbook.Name, ".xlsb", "") + 1)).Value = 0
        Loop
    End If

    
    'Update Tables
    If oXL.Application.Range("ам╤у!S2") = 1 Then
        Call getTable
        oXL.Application.StatusBar = "Tables Updated"
    End If
    On Error Resume Next
    'Filter Ranges to Calculate
    Dim r As Range
    Dim calList As Range
    Set calList = getCalculationList(CLng(Replace(ActiveWorkbook.Name, ".xlsb", "")))
     'Switch Sheet
     If Application.ActiveSheet.Name <> calList.Worksheet.Name Then
        Application.Workbooks(1).Sheets(calList.Worksheet.Name).Activate
    End If
    
    
    Set r = calList 'Filterll(calList, seqFrom, seqTo, (calList.Columns.count > 1 And calList.Areas.count = 1))
    Debug.Print calList.Address
    'Calculate and Save
    If Not (r Is Nothing) And ((seqTo + seqFrom) <> 0) Then
        Call RefreshCal(r)
        'Call ParallelMethods.SaveRangeToMaster(workbookName, r)
    End If
    
    'Done
    Do Until oXL.Application.Range("ам╤у!N" & CStr(Replace(ActiveWorkbook.Name, ".xlsb", "") + 1)).Value = 0
        On Error Resume Next
        oXL.Application.Range("ам╤у!N" & CStr(Replace(ActiveWorkbook.Name, ".xlsb", "") + 1)).Value = 0
    Loop
    On Error Resume Next
    oXL.Application.StatusBar = CStr(oXL.Application.Range("ам╤у!R2").Value2 - WorksheetFunction.sum(oXL.Application.Range("ам╤у!N" & "2:N" & CStr(2 + oXL.Application.Range("ам╤у!R2").Value2)))) & " in " & CStr(oXL.Application.Range("ам╤у!R2").Value2) & " Completed " & CStr(Now())

End Sub

Public Function AvalibleCoreNum()
    threadC = 12
    'Get Instances
    Dim instances As Collection
    Set instances = GetExcelInstances()
    Dim instanceIndex As New Collection
    'For i = instances.count To 1 Step -1
    For i = 1 To instances.count
        For j = 1 To threadC
            threadFileName = ActiveWorkbook.path & "\" & CStr(j) & ".xlsb"
            If threadFileName = instances(i).ActiveWorkbook.FullName Then
                instanceIndex.Add i
            End If
        Next j
    Next i
    AvalibleCoreNum = instanceIndex.count
End Function
Public Sub SyncCores()
threadC = 12
    'Get Instances
    Dim instances As Collection
    Set instances = GetExcelInstances()
    Dim instanceIndex As New Collection

    For j = 1 To threadC
        For i = 1 To instances.count
            threadFileName = ActiveWorkbook.path & "\" & CStr(j) & ".xlsb"
            If threadFileName = instances(i).ActiveWorkbook.FullName Then
                instanceIndex.Add i
                Exit For
            End If
        Next i
    Next j

'Quit Instances
If instanceIndex.count > 0 Then
    For instanceIndexS = 1 To instanceIndex.count
            instances(instanceIndex.Item(instanceIndexS)).Application.DisplayAlerts = False
            instances(instanceIndex.Item(instanceIndexS)).Application.Workbooks(CStr(instanceIndexS) & ".xlsb").Close savechanges:=False
            instances(instanceIndex.Item(instanceIndexS)).Application.DisplayAlerts = True
            instances(instanceIndex.Item(instanceIndexS)).Application.Quit
    Next instanceIndexS
    'Update Statusbar
    Application.StatusBar = "Cores Closed" & " " & CStr(Now())
End If

'Save Master
ActiveWorkbook.Save
For thread = 1 To threadC
    threadFileName = ActiveWorkbook.path & "\" & CStr(thread) & ".xlsb"
    If thread = 1 Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveCopyAs FileName:=threadFileName
        Application.DisplayAlerts = True
    Else
        copythatfile ActiveWorkbook.path & "\" & CStr(1) & ".xlsb", threadFileName
    End If
    
Next thread

'Open Cores
If instanceIndex.count > 0 Then
    'Reopen
    Call ReOpenCores
End If

'Update Statusbar
Application.StatusBar = "Sync Core Complete, Opening Cores" & " " & CStr(Now())
End Sub

Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Sub ReOpenCores()
    'Save the bat
    s = "setlocal" & vbNewLine
    s = s & "cd /d %~dp0" & vbNewLine
    For i = 1 To Range("ам╤у!R2").Value2
        s = s & "start " & Chr(34) & "title" & Chr(34) & " " & Chr(34) & "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" & Chr(34) & " /X " & CStr(i) & ".xlsb" & vbNewLine
    Next i

    'Save the bat file
    sFileName = ActiveWorkbook.path & "\" & "start" & ".bat"
    Open sFileName For Output As #1
    Print #1, s
    Close #1
    Dim retVal
    ChDir ActiveWorkbook.path
    retVal = Shell(ActiveWorkbook.path & "\" & "start.bat")
End Sub


Sub Kill_Excel()
threadC = 12
    'Get Instances
    Dim instances As Collection
    Set instances = GetExcelInstances()
    Dim instanceIndex As New Collection
    'For i = instances.count To 1 Step -1

    For j = 1 To threadC
        For i = 1 To instances.count
            threadFileName = ActiveWorkbook.path & "\" & CStr(j) & ".xlsb"
            If threadFileName = instances(i).ActiveWorkbook.FullName Then
                instanceIndex.Add i
                Exit For
            End If
        Next i
    Next j
    
    'Quit Instances
    If instanceIndex.count > 0 Then
        For instanceIndexS = 1 To instanceIndex.count
                instances(instanceIndex.Item(instanceIndexS)).Application.DisplayAlerts = False
                instances(instanceIndex.Item(instanceIndexS)).Application.ActiveWorkbook.Close savechanges:=False
                instances(instanceIndex.Item(instanceIndexS)).Application.DisplayAlerts = True
                instances(instanceIndex.Item(instanceIndexS)).Application.Quit
                
        Next instanceIndexS
        'Update Statusbar
    End If
    'Status Bar Update
    Application.StatusBar = "Cores Closed" & " " & CStr(Now())
End Sub

Public Function ExistsInCollection(col As Collection, key As Variant) As Boolean
    Dim f As Boolean
    On Error GoTo err
    ExistsInCollection = True
    f = IsObject(col.Item(key))
    Exit Function
err:
    ExistsInCollection = False
End Function
Public Function copythatfile(source As Variant, destination As Variant)
    Dim xlobj As Object
    Set xlobj = CreateObject("Scripting.FileSystemObject")
    xlobj.CopyFile source, destination, True
    Set xlobj = Nothing
End Function
Public Function getCalculationList(Optional core As Long = 0) As Range
    Dim returnR As Range
    If core = 0 Then
        Set returnR = getRange(GetObject(masterFileName).Application.Range("ам╤у!L2").Value2)
    Else
        Set returnR = getRange(GetObject(masterFileName).Application.Range("ам╤у!L" & CStr(2 + core)).Value2)
    End If
    Set returnR = UnionRange(returnR)
    Set getCalculationList = returnR
End Function

Public Function UnionRange(r As Range) As Range
    Dim returnRange As Range
    If Not (r Is Nothing) Then
        For Each cell In r.Areas
'            For Each cell In sArea.Cells
                If (returnRange Is Nothing) Then
                    Set returnRange = cell
                Else
                    Set returnRange = Union(cell, returnRange)
                End If
'            Next cell
            Set UnionRange = returnRange
        Next cell
    Else
        Set UnionRange = returnRange
    End If
End Function
Sub TransferToCores(r As Range)
    Dim mselection As Range
    Dim threadFileName As String
    Dim mAvalibleCoreNum As Double
    mAvalibleCoreNum = Range("ам╤у!R2").Value2
    Set mselection = r
    For i = 1 To mAvalibleCoreNum
        threadFileName = ActiveWorkbook.path & "\" & CStr(i) & ".xlsb"
        Call InstanceSync(threadFileName, mselection)
        Application.StatusBar = "Synchronizing...  " & CStr(CDbl(i) * 100 / mAvalibleCoreNum) & "%"
    Next i
    Application.StatusBar = "Core Synchronized"
End Sub
Sub InstanceSync(threadFileName As String, r As Range)
    Set oXL = GetObject(threadFileName)
        For Each Item In r.Areas
            Dim caledArray As Variant
            caledArray = Item.Formula
            
            On Error Resume Next
            If Item.Cells(1).HasArray Then
                 oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula = caledArray
                oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).FormulaArray = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula
            Else
                oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula = caledArray
            End If

    Next Item
End Sub
Sub InstanceSyncBack(threadFileName As String, r As Range)
    Set oXL = GetObject(threadFileName)
        For Each Item In r.Areas
            Dim caledArray As Variant
            caledArray = Item.Formula
            
            On Error Resume Next
            'oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).range(Item.Address)(1) = 0
            If Item.HasArray Then
                Item.Formula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula
                Item.FormulaArray = Item.Formula
            Else
                Item.Formula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula
            End If
    Next Item
End Sub
Sub SyncTable(r As Range)
    Dim r As Range
    Set r = Range(Range("ам╤у!J2").Value2)
    
    For i = 1 To Range("ам╤у!R2").Value
        threadFileName = ActiveWorkbook.path & "\" & CStr(i) & ".xlsb"
        On Error GoTo wtf
        Set oXL = GetObject(threadFileName)
        oXL.Application.Workbooks(1).Sheets(r.Worksheet.Name).Range(AddressEx(r)).Formula = r.Formula
        Set oXL = Nothing
        
wtf:
    On Error GoTo 0
    
    Next i
End Sub

Sub getTable()
    Set oXL = GetObject(masterFileName)
    Dim toRefresh As Range
    Set toRefresh = Nothing
    'CurrentSheet = oXL.Application.range(oXL.Application.range("ам╤у!Q2").Value2).Worksheet.Name
    CurrentSheet = oXL.Application.ActiveSheet.Name
    'Check Same Sheet
    For i = 2 To 100
        If oXL.Application.Range("ам╤у!$Q$" & i).Value2 <> vbNullString Then
            If (oXL.Application.Range(oXL.Application.Range("ам╤у!$Q$" & i).Value2).Worksheet.Name = CurrentSheet) Then
                If (toRefresh Is Nothing) Then
                    Set toRefresh = Range(oXL.Application.Range("ам╤у!$Q$" & i).Value2)
                Else
                    Set toRefresh = Union(toRefresh, Range(oXL.Application.Range("ам╤у!$Q$" & i).Value2))
                End If
            End If
        End If
    Next i
    
    'Refresh
    If Not (toRefresh Is Nothing) Then
        For Each sArea In toRefresh.Areas
            For Each Item In sArea.Cells
                On Error Resume Next
                If oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).HasArray Then
                    arrayformula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula
                    If Len(arrayformula) < 255 Then
                        Item.FormulaArray = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula
                    End If
                Else
                    Item.Formula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).Range(Item.Address).Formula
                End If
            Next Item
        Next sArea
    End If

        Set oXL = Nothing
End Sub
Public Function AddressEx(Range As Range, Optional blnBuildAddressForNamedRangeValue As Boolean = False) As String
                                         
    Const Seperator As String = ","

    Dim WorksheetName As String
    Dim TheAddress As String
    Dim Areas As Areas
    Dim Area As Range

    'WorksheetName = "'" & Range.Worksheet.Name & "'"
    WorksheetName = Range.Worksheet.Name

    For Each Area In Range.Areas
'           ='Sheet 1'!$H$8:$H$15,'Sheet 1'!$C$12:$J$12
        TheAddress = TheAddress & WorksheetName & "!" & Area.Address(External:=False) & Seperator

    Next Area

    AddressEx = Left(TheAddress, Len(TheAddress) - Len(Seperator))

    If blnBuildAddressForNamedRangeValue Then
        AddressEx = "=" & AddressEx
    End If

End Function


Function subRange(r As Range, startPos As Long, endPos As Long) As Range
    Set subRange = r.Parent.Range(r.Cells(startPos), r.Cells(endPos))
End Function


Public Sub WaitForCompletion(startTime As Double)

    If WorksheetFunction.sum(Range("ам╤у!N" & "2:N" & CStr(2 + Range("ам╤у!R2").Value2))) > 0 Then
        Application.OnTime Now + TimeValue("00:00:01"), "'WaitForCompletion """ & startTime & """'"
    Else
        Call returnToMaster
        Application.StatusBar = "Completed in " & Format(CStr(Now() - startTime), "hh:mm:ss")
    End If
End Sub
Public Sub returnToMaster()
    Dim threadFileName As String
    Corenum = Range("ам╤у!R2").Value2
    For i = 1 To Corenum
        If Range("ам╤у!L" & CStr((2 + i))).Value2 <> "" Then
            threadFileName = ActiveWorkbook.path & "\" & CStr(i) & ".xlsb"
            Call InstanceSyncBack(threadFileName, getRange(Range("ам╤у!L" & CStr((2 + i))).Value2))
            'Call TransferToCores(getRange(Range("ам╤у!L" & CStr((2 + i))).Value2))
            Application.StatusBar = "Retrieving...  " & CStr(CDbl(i) * 100 / CDbl(Corenum)) & "%"
        End If
    Next i
    Application.StatusBar = "Results Retrieved"
    
End Sub


Sub gtertrtweh()
Dim g As Range
Set g = Range("B7:B10")
Debug.Print AddressEx(g)
Set g = Range(AddressEx(g))
Debug.Print AddressEx(g)

End Sub
Sub ertght()
    Dim r As Range
    Set r = Range("AQ20:AR25,AS20:AS24")
    For Each Item In r.Areas
        Set r = Union(r, Item)
    Next Item
    Dim B As Range
    Set B = Range("AT20:AU25,AV20:AV24")
    For Each Item In r.Areas
        Set B = Union(B, Item)
    Next Item
    
    Call RefreshCal(B)
    Debug.Print B.Address

End Sub

Sub gqwertrefvg()
    Dim we As Range
    Set we = Range("F265:F268,F266:F277")
    Set we = UnionRange(we)
    Debug.Print we.Address
End Sub
Sub testvgg()
    Dim calList As Range
    Set calList = Range("Table56")
    If (calList.Columns.count > 1 And calList.Areas.count = 1) Then
        Call RunForVBA("TC.xlsb", 1, calList.Columns.count)
    Else
        Call RunForVBA("TC.xlsb", 1, calList.count)
    End If
    
    
End Sub
Public Sub gg()
thread = 6
threadFileName = ActiveWorkbook.path & "\" & CStr(thread) & ".xlsb"
    Set xlobj = GetObject(threadFileName)
    xlobj.Application.ActiveWorkbook.Close savechanges:=False
End Sub


Sub hfdgsf()
    Dim r As Range
    Set r = Range("╩Ы╜х╕Т╜p!N16:M20")
    MsgBox r.Worksheet.Name
End Sub

Public Sub ggergr()
    getCalculationList().Calculate
End Sub


Public Sub ggggggg()
    Dim returnR As Range
    
    threadFileName = ActiveWorkbook.path & "\" & "TC.xlsb"
    Set oXL = GetObject(threadFileName)
'
'
    Set r = Range("B10:B10")
    Set g = Range("B8:B8")
    'r.Copy
    
    oXL.Application.Workbooks(1).Sheets(r.Worksheet.Name).Range(r.Address).Value2 = r.Value2
    oXL.Application.Workbooks(1).Sheets(r.Worksheet.Name).Range(r.Address).Formula = "=NOW()"  '.PasteSpecial Paste:=xlPasteValues
    
    
    
End Sub
Sub fhwiuefhqporf()
    Set g = Range("E13")
    g.FormulaArray = g.Formula
    MsgBox ""
End Sub



Sub gggggggg()
    Dim gg As Range
    Set gg = Range("A1:B10")
    MsgBox gg.Columns.count
    x = 0

End Sub

Sub gewrsthdg()
    Dim returnR As Range
    Set xlobj = GetObject(threadFileName)
    Set returnR = xlobj.Application.Range("ам╤у!$Q$" & CStr(1))
    MsgBox returnR.Address
End Sub
Sub gqerwg()
Debug.Print ActiveWorkbook.path & "\" & "TC.xlsb"
End Sub
Sub etsghd()
    Debug.Print Application.ActiveSheet.Name
End Sub

'Public Sub SyncCores()
'threadC = 12
'    'Get Instances
'    Dim instances As Collection
'    Set instances = GetExcelInstances()
'    Dim instanceIndex As New Collection
'    'For i = instances.count To 1 Step -1
'
'    For j = 1 To threadC
'        For i = 1 To instances.count
'            threadFileName = ActiveWorkbook.Path & "\" & CStr(j) & ".xlsb"
'            If threadFileName = instances(i).ActiveWorkbook.FullName Then
'                instanceIndex.Add i
'                Exit For
'            End If
'        Next i
'    Next j
'
''Quit Instances
'If instanceIndex.count > 0 Then
'    For instanceIndexS = 1 To instanceIndex.count
''            threadFileName = ActiveWorkbook.Path & "\b" & CStr(instanceIndexS) & ".xlsb"
''            instances(instanceIndex.Item(instanceIndexS)).Application.Workbooks.Open (threadFileName)
'            instances(instanceIndex.Item(instanceIndexS)).Application.DisplayAlerts = False
'
'            instances(instanceIndex.Item(instanceIndexS)).Application.Workbooks(CStr(instanceIndexS) & ".xlsb").Close savechanges:=False
'            instances(instanceIndex.Item(instanceIndexS)).Application.DisplayAlerts = True
'            instances(instanceIndex.Item(instanceIndexS)).Application.Quit
'            'instances(instanceIndex.Item(instanceIndexS)) = Nothing
'
'    Next instanceIndexS
'    'Update Statusbar
'    Application.StatusBar = "Cores Closed" & " " & CStr(Now())
'End If
'
''Save Master
'ActiveWorkbook.Save
'For thread = 1 To threadC
'    threadFileName = ActiveWorkbook.Path & "\" & CStr(thread) & ".xlsb"
'    If thread = 1 Then
'        Application.DisplayAlerts = False
'        ActiveWorkbook.SaveCopyAs FileName:=threadFileName
'        Application.DisplayAlerts = True
'    Else
'        copythatfile ActiveWorkbook.Path & "\" & CStr(1) & ".xlsb", threadFileName
'    End If
'
'Next thread
'
''Open Cores
'If instanceIndex.count > 0 Then
'
''For i = 1 To instanceIndex.count
''    Dim openedXls As String
''    openedXls = ActiveWorkbook.Path & "\b" & CStr(i) & ".xlsb"
''    threadFileName = ActiveWorkbook.Path & "\" & CStr(i) & ".xlsb"
''
''    'Save the VBscript
''    s = "Set objExcel = GetObject(""" & openedXls & """):"
''    s = s & "With objExcel:"
''    s = s & ".Application.Workbooks.Open(""" & threadFileName & """):"
''    s = s & "End With:"
''    'Save the VBscript file
''    sFileName = ActiveWorkbook.Path & "\" & "open" & CStr(i) & ".vbs"
''    Open sFileName For Output As #1
''    Print #1, s
''    Close #1
''    'Execute the VBscript file asynchronously
''    Set wsh = VBA.CreateObject("WScript.Shell")
''    wsh.Run """" & sFileName & """"
''    Set wsh = Nothing
''
''Next i
''    ' Open one by one
'''    For instanceIndexS = 1 To instanceIndex.count
'''        threadFileName = ActiveWorkbook.Path & "\" & CStr(instanceIndexS) & ".xlsb"
'''        instances(instanceIndex.Item(instanceIndexS)).Application.Workbooks.Open (threadFileName)
'''    Next instanceIndexS
'
'    'Reopen
'    Call ReOpenCores
'End If
'
''Update Statusbar
'Application.StatusBar = "Sync Core Complete, Opening Cores" & " " & CStr(Now())
'End Sub
