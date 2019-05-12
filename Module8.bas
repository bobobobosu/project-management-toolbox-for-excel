Attribute VB_Name = "Module8"
'Public Const masterFileName  As String = "Z:\My Drive\_Storage\Backup\Documents\RAMDISK_loc\Documents\root\Data\TC.xlsb"
Public Sub RecalculateSelection()
    On Error GoTo ErrorHandler
    If TypeName(Selection) = "Range" Then
        'Detect Calculation Settings
        If range("ам╤у!K2").Value = 1 Then 'And Selection.Cells.Count < 50 Then
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
    Call llCalculate(Selection)
End Function
Function MySheet()

   MySheet = range("A1").Worksheet.name

End Function
Public Sub customCalculate()
    Dim sel As range
    Set sel = Selection
    Dim table2 As range
    Set table2 = range("╙М╝Ф2")
    If sel.Worksheet.name = table2.Worksheet.name Then
        If Application.Intersect(table2, sel).Count > 0 Then
            Call CalculateRange1
            Call CalculateTable2ByOrder
            Call CalculateRange1
        Else
            Call llf(sel)
        End If
    Else
        Call llf(sel)
    End If
End Sub
Public Function llf(allr As Variant)
    For Each cell1 In allr
        Dim r As range
        Set r = cell1
        On Error Resume Next
        If Not r.DirectPrecedents Is Nothing Then
            For Each cell In r.DirectPrecedents
                Dim r2 As range
                Set r2 = cell
                Set r2 = range(AddressEx(r2))
                r2.Calculate
                Application.OnTime Now, "llCalculate", r2
            Next
        End If
        r.Calculate
        Application.OnTime Now, "llCalculate", r
    Next
    llf = allr.Value2
End Function
Public Sub llCalculate(rIN As range)
    If range("ам╤у!P2") = 1 Then rIN.Calculate
    Dim r As range
    Set r = rIN
    If range("ам╤у!O2") = 0 Then Exit Sub
        
    'Run ll if idle
    Dim filtered As range
    Set filtered = Filterll(r)
    If (Not (filtered Is Nothing)) Then
'        If range("ам╤у!O2") = 1 Then
'            'Isolate firstRow
'            Call RefreshCal(r.Rows(1))
'            If rIN.Rows.Count = 1 Then Exit Sub
'            Set filtered = getSubRange(2, 1, r.Rows.Count, r.Columns.Count, r)
'            If filtered.Count <= 10 Then
'                If Not (filtered Is Nothing) Then
'                    Call RefreshCal(filtered)
'                End If
'            Else
'                If range("ам╤у!L2").Value2 <> vbNullString Then
'                    range("ам╤у!L2").Value2 = range("ам╤у!L2").Value2 + "," + AddressEx(filtered)
'                    range("ам╤у!L2").Value2 = RangeToColumnsAdd(range("ам╤у!L2").Value2)
'                Else
'                    range("ам╤у!L2").Value2 = RangeToColumnsAdd(AddressEx(filtered))
'                End If
'            End If
'        ElseIf range("ам╤у!O2") = 3 Then
        If range("ам╤у!O2") = 1 Then
                If Not (filtered Is Nothing) Then
                    Call RefreshCal(filtered)
                End If
        Else
                If range("ам╤у!L2").Value2 <> vbNullString Then
                    range("ам╤у!L2").Value2 = range("ам╤у!L2").Value2 + "," + AddressEx(filtered)
                    range("ам╤у!L2").Value2 = RangeToColumnsAdd(range("ам╤у!L2").Value2)
                Else
                    range("ам╤у!L2").Value2 = RangeToColumnsAdd(AddressEx(filtered))
                End If
        End If
    End If
End Sub
Function RangeToColumnsAdd(addressStr As Variant)
    Dim r As range
    Set r = Range2(addressStr)
    Set r = UnionRange(r)
    
    Dim result As String
    Dim thisCol As range
    result = ""
    For Each iarea In r.Areas
        For Each icol In iarea.Columns
            Set thisCol = icol
            If result = vbNullString Then
                result = AddressEx(thisCol)
            Else
                result = result + "," + AddressEx(thisCol)
            End If
        Next
    Next
    RangeToColumnsAdd = result
End Function
Public Function getRange(rangeS As String, Optional filename As String = "") As range
    Dim r As range
    Dim remains As String
    remains = rangeS
    If Len(rangeS) > 254 Then
        'get All but Last
        Do Until InStr(remains, ",") = False
            commaposition = InStr(remains, ",")
            ms = Left(remains, commaposition - 1)
    
             If filename = "" Then
                If (r Is Nothing) Then
                    Set r = range(ms)
                Else
                    Set r = Union(r, range(ms))
                End If
            Else
                If (r Is Nothing) Then
                    Set r = GetObject(filename).Application.range(ms)
                Else
                    Set r = Union(r, GetObject(filename).Application.range(ms))
                End If
            End If
            remains = Right(remains, Len(remains) - commaposition)
        Loop
        'get Last One
         If filename = "" Then
            If (r Is Nothing) Then
                Set r = range(remains)
            Else
                Set r = Union(r, range(remains))
            End If
        Else
            If (r Is Nothing) Then
                Set r = GetObject(filename).Application.range(remains)
            Else
                Set r = Union(r, GetObject(filename).Application.range(remains))
            End If
        End If
        Set getRange = r
    Else
         If filename = "" Then
                Set r = range(remains)
        Else
                Set r = GetObject(filename).Application.range(remains)
        End If
    End If
    Set getRange = r
End Function
Public Sub TriggerSync()
    Dim blankR As range
    range("ам╤у!Q2").Value2 = AddressEx(Selection)
    range("ам╤у!L2").Value2 = ""
    Set blankR = range("A1").Resize(1, range("ам╤у!R2").Value)
    range("ам╤у!S2") = 1
    Call TestParallelForLoop(blankR)
End Sub

Public Function TestParallelForLoop(calList As range)
    Set calList = UnionRange(calList)
    'Clean ll Stat
    Set oXL = GetObject(masterFileName)
    For i = 2 To (2 + range("ам╤у!R2").Value - 1)
            oXL.Application.range("ам╤у!N" & i).Value = 1
    Next i
    oXL.Application.StatusBar = CStr(oXL.Application.range("ам╤у!R2").Value2 - WorksheetFunction.sum(oXL.Application.range("ам╤у!N" & "2:N" & CStr(2 + oXL.Application.range("ам╤у!R2").Value2)))) & " in " & CStr(oXL.Application.range("ам╤у!R2").Value2) & " Completed " & CStr(Now())
    
    'Set ll
    Dim para As Parallel
    Set para = New Parallel
    para.SetThreads range("ам╤у!R2").Value

    'Run ll
    Call para.ParallelFor("RunForVBA")
    
    'Timer Start
    Call WaitForCompletion(Now())
End Function

Sub RunForVBA(workbookName As String, seqFrom As Long, seqTo As Long)
    Set oXL = GetObject(masterFileName)
    'No Run
    If seqTo = 0 Then
        Do Until oXL.Application.range("ам╤у!N" & CStr(Replace(ActiveWorkbook.name, ".xlsb", "") + 1)).Value = 0
            On Error Resume Next
            oXL.Application.range("ам╤у!N" & CStr(Replace(ActiveWorkbook.name, ".xlsb", "") + 1)).Value = 0
        Loop
    End If

    
    'Update Tables
    If oXL.Application.range("ам╤у!S2") = 1 Then
        Call getTable
        oXL.Application.StatusBar = "Tables Updated"
    End If
    On Error Resume Next
    'Filter Ranges to Calculate
    Dim r As range
    Dim calList As range
    Set calList = getCalculationList(CLng(Replace(ActiveWorkbook.name, ".xlsb", "")))
     'Switch Sheet
     If Application.ActiveSheet.name <> calList.Worksheet.name Then
        Application.Workbooks(1).Sheets(calList.Worksheet.name).Activate
    End If
    
    
    Set r = calList 'Filterll(calList, seqFrom, seqTo, (calList.Columns.count > 1 And calList.Areas.count = 1))
    Debug.Print calList.address
    'Calculate and Save
    If Not (r Is Nothing) And ((seqTo + seqFrom) <> 0) Then
        Call RefreshCal(r)
        'Call ParallelMethods.SaveRangeToMaster(workbookName, r)
    End If
    
    'Done
    Do Until oXL.Application.range("ам╤у!N" & CStr(Replace(ActiveWorkbook.name, ".xlsb", "") + 1)).Value = 0
        On Error Resume Next
        oXL.Application.range("ам╤у!N" & CStr(Replace(ActiveWorkbook.name, ".xlsb", "") + 1)).Value = 0
    Loop
    On Error Resume Next
    oXL.Application.StatusBar = CStr(oXL.Application.range("ам╤у!R2").Value2 - WorksheetFunction.sum(oXL.Application.range("ам╤у!N" & "2:N" & CStr(2 + oXL.Application.range("ам╤у!R2").Value2)))) & " in " & CStr(oXL.Application.range("ам╤у!R2").Value2) & " Completed " & CStr(Now())

End Sub
Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
Public Function AvalibleCoreNum()
    threadC = 12
    'Get Instances
    Dim instances As Collection
    Set instances = GetExcelInstances()
    Dim instanceIndex As New Collection
    'For i = instances.count To 1 Step -1
    For i = 1 To instances.Count
        For j = 1 To threadC
            threadFileName = ActiveWorkbookpath & "\" & CStr(j) & ".xlsb"
            If GetFilenameFromPath(threadFileName) = GetFilenameFromPath(instances(i).ActiveWorkbook.FullName) Then
                instanceIndex.Add i
            End If
        Next j
    Next i
    AvalibleCoreNum = instanceIndex.Count
End Function
Public Sub SyncCores()
threadC = 12
    'Get Instances
    Dim instances As Collection
    Set instances = GetExcelInstances()
    Dim instanceIndex As New Collection

    For j = 1 To threadC
        For i = 1 To instances.Count
            threadFileName = ActiveWorkbookpath & "\" & CStr(j) & ".xlsb"
            If GetFilenameFromPath(threadFileName) = GetFilenameFromPath(instances(i).ActiveWorkbook.FullName) Then
                instanceIndex.Add i
                Exit For
            End If
        Next i
    Next j

'Quit Instances
If instanceIndex.Count > 0 Then
    For instanceIndexS = 1 To instanceIndex.Count
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
    threadFileName = ActiveWorkbookpath & "\" & CStr(thread) & ".xlsb"
    If thread = 1 Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveCopyAs filename:=threadFileName
        Application.DisplayAlerts = True
    Else
        copythatfile ActiveWorkbookpath & "\" & CStr(1) & ".xlsb", threadFileName
    End If
Next thread

'Open Cores
If instanceIndex.Count > 0 Then
    'Reopen
    Call ReOpenCores
End If

'Update Statusbar
Application.StatusBar = "Sync Core Complete, Opening Cores" & " " & CStr(Now())
End Sub


Sub SaveCores(numCores As Variant)
    ActiveWorkbook.Save
    For thread = 1 To numCores
        threadFileName = ActiveWorkbookpath & "\" & CStr(thread) & ".xlsb"
        If thread = 1 Then
            Application.DisplayAlerts = False
            ActiveWorkbook.SaveCopyAs filename:=threadFileName
            Application.DisplayAlerts = True
        Else
            copythatfile ActiveWorkbookpath & "\" & CStr(1) & ".xlsb", threadFileName
        End If
    Next thread
End Sub


Function IsWorkBookOpen(filename As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open filename For Input Lock Read As #ff
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
    For i = 1 To range("ам╤у!R2").Value2
        s = s & "start " & Chr(34) & "title" & Chr(34) & " " & Chr(34) & "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE" & Chr(34) & " /X " & ActiveWorkbookpath & "\" & CStr(i) & ".xlsb" & vbNewLine
    Next i

    'Save the bat file
    sFileName = ActiveWorkbookpath & "\" & "start" & ".bat"
    Open sFileName For Output As #1
    Print #1, s
    Close #1
    Dim retVal
    ChDir ActiveWorkbookpath
    retVal = Shell(ActiveWorkbookpath & "\" & "start.bat")
End Sub


Sub Kill_Excel()
threadC = 12
    'Get Instances
    Dim instances As Collection
    Set instances = GetExcelInstances()
    Dim instanceIndex As New Collection
    'For i = instances.count To 1 Step -1

    For j = 1 To threadC
        For i = 1 To instances.Count
            threadFileName = ActiveWorkbookpath & "\" & CStr(j) & ".xlsb"
            If GetFilenameFromPath(threadFileName) = GetFilenameFromPath(instances(i).ActiveWorkbook.FullName) Then
                instanceIndex.Add i
                Exit For
            End If
        Next i
    Next j
    
    'Quit Instances
    If instanceIndex.Count > 0 Then
        For instanceIndexS = 1 To instanceIndex.Count
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


Function HasVal(coll As Collection, strKey As String) As Boolean
    For L = 1 To coll.Count
        On Error GoTo err
        If coll(L) = strKey Then
            HasVal = True
            Exit Function
        End If
    Next L
err:
    HasVal = False
End Function


Public Function copythatfile(source As Variant, destination As Variant)
    Dim xlobj As Object
    Set xlobj = CreateObject("Scripting.FileSystemObject")
    xlobj.CopyFile source, destination, True
    Set xlobj = Nothing
End Function
Public Function getCalculationList(Optional core As Long = 0) As range
    Dim returnR As range
    If core = 0 Then
        Set returnR = getRange(GetObject(masterFileName).Application.range("ам╤у!L2").Value2)
    Else
        Set returnR = getRange(GetObject(masterFileName).Application.range("ам╤у!L" & CStr(2 + core)).Value2)
    End If
    Set returnR = UnionRange(returnR)
    Set getCalculationList = returnR
End Function

Public Function UnionRange(r As range) As range
    Dim returnRange As range
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
Sub TransferToCores(r As range)
    Dim mselection As range
    Dim threadFileName As String
    Dim mAvalibleCoreNum As Double
    mAvalibleCoreNum = AvalibleCoreNum()
    If mAvalibleCoreNum < 1 Then
        Exit Sub
    End If
    Set mselection = r
    For i = 1 To mAvalibleCoreNum
        threadFileName = ActiveWorkbookpath & "\" & CStr(i) & ".xlsb"
        Call InstanceSync(threadFileName, mselection)
        Application.StatusBar = "Synchronizing...  " & CStr(CDbl(i) * 100 / mAvalibleCoreNum) & "%"
    Next i
    Application.StatusBar = "Core Synchronized"
End Sub

Sub InstanceSyncUpdateValue(syncaddress As Variant)
    Set oXL = GetObject(ActiveWorkbook.path & "\" & "TC" & ".xlsb")
    Set toSync = Range2(syncaddress)
    For Each Item In toSync.Areas
        Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).Value2 = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).Value2
    Next
End Sub
Sub InstanceSyncUpdateFormula(syncaddress As Variant)
    Set oXL = GetObject(ActiveWorkbook.path & "\" & "TC" & ".xlsb")
    Set toSync = Range2(syncaddress)
    For Each IAreas In toSync.Areas
        Dim Item As range
        Set Item = oXL.Application.Workbooks(1).Sheets(IAreas.Worksheet.name).range(IAreas.address)
        
        On Error Resume Next
        If Item.HasFormula Then
            If Item.HasArray Then
                IAreas.formula = Item.formula
                IAreas.FormulaArray = Item.formula
            Else
                IAreas.formula = Item.formula
            End If
        Else
            IAreas.formula = Item.formula
        End If
    Next
End Sub



Sub addInstanceSync(r As range)
    If range("ам╤у!V2").Value2 = "[Refresh]" Then Exit Sub
    If AddressEx(r) <> AddressEx(range("ам╤у!V2")) And Split(AddressEx(r), "!")(0) <> "ам╤у" Then
        range("ам╤у!V2").Value2 = mergeRange(range("ам╤у!V2").Value2, r)
    End If
End Sub
Sub InstanceSyncBackFormula(syncaddress As Variant)
    Set oXL = GetObject(ActiveWorkbook.path & "\" & "TC" & ".xlsb")
    Set toSync = Range2(syncaddress)
    For Each Item In toSync.Areas
        Dim IAreas As range
        Set IAreas = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address)
        
        On Error Resume Next
        If Item.HasFormula Then
            If Item.HasArray Then
                IAreas.formula = Item.formula
                IAreas.FormulaArray = Item.formula
            Else
                IAreas.formula = Item.formula
            End If
        Else
            IAreas.formula = Item.formula
        End If
    Next
End Sub

Sub InstanceSync(threadFileName As String, r As range)
    Set oXL = GetObject(threadFileName)
    For Each Item In r.Areas
        Dim caledArray As Variant
        caledArray = Item.formula
        
        On Error Resume Next
        If Item.HasFormula Then
            Dim toCheck As range
            Set toCheck = Item
            If Checkll(toCheck) Then
                If Item.HasArray Then
                     oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula = caledArray
                    oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).FormulaArray = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula
                Else
                    oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula = caledArray
                End If
            End If
        Else
            oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula = caledArray
        End If

    Next Item
End Sub
Sub InstanceSyncBack(threadFileName As String, r As range)
    Set oXL = GetObject(threadFileName)
    For Each Item In r.Areas
        Dim caledArray As Variant
        caledArray = Item.formula
        
        On Error Resume Next
        'oXL.Application.Workbooks(1).Sheets(Item.Worksheet.Name).range(Item.Address)(1) = 0
        If Item.HasFormula Then
            Dim toCheck As range
            Set toCheck = Item
            If Checkll(toCheck) Then
                If Item.HasArray Then
                    Item.formula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula
                    Item.FormulaArray = Item.formula
                Else
                    Item.formula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula
                End If
            End If
        Else
            Item.formula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula
        End If
    Next Item
End Sub
Sub SyncTable(r As range)
    Dim r As range
    Set r = range(range("ам╤у!J2").Value2)
    
    For i = 1 To range("ам╤у!R2").Value
        threadFileName = ActiveWorkbookpath & "\" & CStr(i) & ".xlsb"
        On Error GoTo wtf
        Set oXL = GetObject(threadFileName)
        oXL.Application.Workbooks(1).Sheets(r.Worksheet.name).range(AddressEx(r)).formula = r.formula
        Set oXL = Nothing
        
wtf:
    On Error GoTo 0
    
    Next i
End Sub

Sub getTable()
    Set oXL = GetObject(masterFileName)
    Dim toRefresh As range
    Set toRefresh = Nothing
    'CurrentSheet = oXL.Application.range(oXL.Application.range("ам╤у!Q2").Value2).Worksheet.Name
    CurrentSheet = oXL.Application.ActiveSheet.name
    'Check Same Sheet
    For i = 2 To 100
        If oXL.Application.range("ам╤у!$Q$" & i).Value2 <> vbNullString Then
            If (oXL.Application.range(oXL.Application.range("ам╤у!$Q$" & i).Value2).Worksheet.name = CurrentSheet) Then
                If (toRefresh Is Nothing) Then
                    Set toRefresh = range(oXL.Application.range("ам╤у!$Q$" & i).Value2)
                Else
                    Set toRefresh = Union(toRefresh, range(oXL.Application.range("ам╤у!$Q$" & i).Value2))
                End If
            End If
        End If
    Next i
    
    'Refresh
    If Not (toRefresh Is Nothing) Then
        For Each sArea In toRefresh.Areas
            For Each Item In sArea.Cells
                On Error Resume Next
                If oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).HasArray Then
                    arrayformula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula
                    If Len(arrayformula) < 255 Then
                        Item.FormulaArray = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula
                    End If
                Else
                    Item.formula = oXL.Application.Workbooks(1).Sheets(Item.Worksheet.name).range(Item.address).formula
                End If
            Next Item
        Next sArea
    End If

        Set oXL = Nothing
End Sub
Public Function AddressEx(range As range, Optional blnBuildAddressForNamedRangeValue As Boolean = False) As String
                                         
    Const Seperator As String = ","

    Dim WorksheetName As String
    Dim TheAddress As String
    Dim Areas As Areas
    Dim Area As range

    'WorksheetName = "'" & Range.Worksheet.Name & "'"
    WorksheetName = range.Worksheet.name

    For Each Area In range.Areas
'           ='Sheet 1'!$H$8:$H$15,'Sheet 1'!$C$12:$J$12
        TheAddress = TheAddress & WorksheetName & "!" & Area.address(External:=False) & Seperator

    Next Area

    AddressEx = Left(TheAddress, Len(TheAddress) - Len(Seperator))

    If blnBuildAddressForNamedRangeValue Then
        AddressEx = "=" & AddressEx
    End If

End Function


Function subRange(r As range, startPos As Long, endPos As Long) As range
    Set subRange = r.Parent.range(r.Cells(startPos), r.Cells(endPos))
End Function


Public Sub WaitForCompletion(StartTime As Double)

    If WorksheetFunction.sum(range("ам╤у!N" & "2:N" & CStr(2 + range("ам╤у!R2").Value2))) > 0 Then
        Application.OnTime Now + TimeValue("00:00:01"), "'WaitForCompletion """ & StartTime & """'"
    Else
        Call returnToMaster
        Application.StatusBar = "Completed in " & Format(CStr(Now() - StartTime), "hh:mm:ss")
    End If
End Sub
Public Sub returnToMaster()
    Dim threadFileName As String
    Corenum = range("ам╤у!R2").Value2
    For i = 1 To Corenum
        If range("ам╤у!L" & CStr((2 + i))).Value2 <> "" Then
            threadFileName = ActiveWorkbookpath & "\" & CStr(i) & ".xlsb"
            Call InstanceSyncBack(threadFileName, getRange(range("ам╤у!L" & CStr((2 + i))).Value2))
            'Call TransferToCores(getRange(Range("ам╤у!L" & CStr((2 + i))).Value2))
            Application.StatusBar = "Retrieving...  " & CStr(CDbl(i) * 100 / CDbl(Corenum)) & "%"
        End If
    Next i
    Application.StatusBar = "Results Retrieved"
    
End Sub
