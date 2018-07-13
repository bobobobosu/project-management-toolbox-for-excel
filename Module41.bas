Attribute VB_Name = "Module41"
#If VBA7 Then
  Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" ( _
    ByVal hWnd As LongPtr, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long

  Private Declare PtrSafe Function FindWindowExA Lib "user32" ( _
    ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
#Else
  Private Declare Function AccessibleObjectFromWindow Lib "oleacc" ( _
    ByVal hWnd As Long, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long

  Private Declare Function FindWindowExA Lib "user32" ( _
    ByVal hwndParent As Long, ByVal hwndChildAfter As Long, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As Long
#End If

Public Function getXL() As Application
  Dim xl As Application
  'MsgBox Application.ActiveWorkbook.FullName
  Dim instances As Collection
  Set instances = GetExcelInstances()
  
  For i = instances.count To 1 Step -1
    If Application.ActiveWorkbook.FullName = instances(i).ActiveWorkbook.FullName Then
        instances.Remove i
    End If
Next i

  For Each xl In instances
        MsgBox "Handle: " & xl.ActiveWorkbook.FullName
  Next
  Set getXL = instances(1)
End Function


Sub gggggg()
    Dim xl As Application
    Set xl = getXL()
    MsgBox xl.ActiveWorkbook.FullName
End Sub

Public Function GetExcelInstances() As Collection
  Dim GUID&(0 To 3), acc As Object, hWnd, hwnd2, hwnd3
  GUID(0) = &H20400
  GUID(1) = &H0
  GUID(2) = &HC0
  GUID(3) = &H46000000

  Set GetExcelInstances = New Collection
  Do
    hWnd = FindWindowExA(0, hWnd, "XLMAIN", vbNullString)
    If hWnd = 0 Then Exit Do
    hwnd2 = FindWindowExA(hWnd, 0, "XLDESK", vbNullString)
    hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", vbNullString)
    If AccessibleObjectFromWindow(hwnd3, &HFFFFFFF0, GUID(0), acc) = 0 Then
      GetExcelInstances.Add acc.Application
    End If
  Loop
End Function

Public Sub ggg()

Dim objXL, WB, strMessage
On Error Resume Next
Set objXL = GetObject(, "Excel.Application")
Set WB = objXL.ActiveWorkbook
On Error GoTo 0
If Not TypeName(objXL) = "Empty" Then
    If Not TypeName(WB) = "Nothing" Then
    strMessage = "Excel Running - " & objXL.ActiveWorkbook.Name & " is active"
    Else
    strMessage = "Excel Running - no workbooks open"
    End If
Else
    strMessage = "Excel NOT Running"
End If
MsgBox strMessage, vbInformation, "Excel Status"""
End Sub

Public Sub gg()

Set objExcel = GetObject("C:\Users\nicki\Downloads\3.xlsb")
Set objWorkbook = objExcel.Application.Workbooks.Open("C:\Users\nicki\Downloads\gg.xlsb")
MsgBox objExcel.FullName

End Sub

Public Sub create()
For thread = 1 To 6
    Dim app As New Excel.Application
    app.Visible = False 'Visible is False by default, so this isn't necessary
    Dim book As Excel.Workbook
    Set book = app.Workbooks.Add(ActiveWorkbook.path & "\" & thread & ".xlsb")
Next thread
End Sub


Public Sub gggg()
    Dim g As String
    g = "=IFERROR(GetESTTIME2(""時序專案(288)"",43205.2789539047,價值時間表!$A$4:$A$243,價值時間表!$G$4:$G$243,0),1582.96285576812)"
    Debug.Print Evaluate(g)
End Sub

