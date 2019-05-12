VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public RightClick As Boolean
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    range("ам╤у!K2").Value = 0
    'Call CreateCalendar
    
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Call MoveToCurrentRow
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As range)
    Call addInstanceSync(target)
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal target As range)
    
    Call setPaths
    On Error Resume Next
    If Intersect(Selection, range("ам╤у!1:4")).Count = 0 Then
        On Error Resume Next
        range("ам╤у!G2").Value2 = range("ам╤у!E2").Value2
        On Error Resume Next
        range("ам╤у!E2").Value2 = AddressEx(Selection(1))
        On Error Resume Next
        range("ам╤у!H2").Value2 = range("ам╤у!F2").Value2
        On Error Resume Next
        range("ам╤у!F2").Value2 = Selection(1).Value2
        On Error Resume Next
        range("ам╤у!J2").Value2 = AddressEx(Selection)
    End If
    Dim mselection As Variant

    If Application.CutCopyMode = False Then
        Call RecalculateSelection
        If Not (Selection Is Nothing) Then
            On Error Resume Next
            If Selection.Count = 1 Then
                'UserForm2.Hide
                FormShowFixed UserForm2
            Else
                On Error Resume Next
                UserForm2.updateScrollbar
                FormShow UserForm2, ActiveCell
            End If
        End If
    Else
        On Error Resume Next
        UserForm2.updateScrollbar
        FormShow UserForm2, ActiveCell
    End If

    Application.ScreenUpdating = False
End Sub


Private Sub Workbook_Open()
    Call setPaths
    'first call appropriate procedure
    Application.ErrorCheckingOptions.BackgroundChecking = False
    Application.OnKey "^c", "CopyFired"
    Application.OnKey "^x", "CutFired"
    range("ам╤у!K2").Value = 1
    Application.Calculation = xlCalculationManual
    Call CallDoEvent
    If AvalibleCoreNum() > 0 Then
        range("ам╤у!O2").Value2 = 1
    Else
        range("ам╤у!O2").Value2 = 0
    End If
    
End Sub
Sub setPaths()
    ActiveWorkbookpath = ActiveWorkbook.path '"C:\Users\nicki\_Sync\Devices\root\Documents\root\Data"
    masterFileName = ActiveWorkbook.path & "\TC.xlsb"
    valueMappath = ActiveWorkbook.path & "\data.json"
    calendarSharePath = "C:\_DATA\_Storage\_Sync\Devices\Share\"
    ftpTmpPath = "D:\_Storage\_tmp"
    calendaricsFileName = "calendar.ics"
End Sub



