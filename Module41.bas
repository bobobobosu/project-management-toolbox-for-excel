Attribute VB_Name = "Module41"
#If VBA7 Then
  Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" ( _
    ByVal hwnd As LongPtr, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long

  Private Declare PtrSafe Function FindWindowExA Lib "User32" ( _
    ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
#Else
  Private Declare Function AccessibleObjectFromWindow Lib "oleacc" ( _
    ByVal hwnd As Long, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long

  Private Declare Function FindWindowExA Lib "User32" ( _
    ByVal hwndParent As Long, ByVal hwndChildAfter As Long, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As Long
#End If

Public Function getXL() As Application
  Dim xl As Application
  'MsgBox Application.ActiveWorkbook.FullName
  Dim instances As Collection
  Set instances = GetExcelInstances()
  
  For i = instances.Count To 1 Step -1
    If GetFilenameFromPath(Application.ActiveWorkbook.FullName) = GetFilenameFromPath(instances(i).ActiveWorkbook.FullName) Then
        instances.Remove i
    End If
Next i

  For Each xl In instances
        MsgBox "Handle: " & xl.ActiveWorkbook.FullName
  Next
  Set getXL = instances(1)
End Function

Public Function GetExcelInstances() As Collection
  Dim GUID&(0 To 3), acc As Object, hwnd, hwnd2, hwnd3
  GUID(0) = &H20400
  GUID(1) = &H0
  GUID(2) = &HC0
  GUID(3) = &H46000000

  Set GetExcelInstances = New Collection
  Do
    hwnd = FindWindowExA(0, hwnd, "XLMAIN", vbNullString)
    If hwnd = 0 Then Exit Do
    hwnd2 = FindWindowExA(hwnd, 0, "XLDESK", vbNullString)
    hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", vbNullString)
    If AccessibleObjectFromWindow(hwnd3, &HFFFFFFF0, GUID(0), acc) = 0 Then
      GetExcelInstances.Add acc.Application
    End If
  Loop
End Function

Public Sub create()
For thread = 1 To 6
    Dim App As New Excel.Application
    App.Visible = False 'Visible is False by default, so this isn't necessary
    Dim book As Excel.Workbook
    Set book = App.Workbooks.Add(ActiveWorkbook.path & "\" & thread & ".xlsb")
Next thread
End Sub
