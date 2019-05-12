Attribute VB_Name = "Module38"
Option Explicit

Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" ( _
    ByVal hwnd As Long, _
    ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal nIndex As Long) As Long

Const LOGPIXELSX = 88
Const LOGPIXELSY = 90
Const TWIPSPERINCH = 1440

Sub ConvertPixelsToPoints(ByRef x As Single, ByRef y As Single)
    Dim hDC As Long
    Dim retVal As Long
    Dim XPixelsPerInch As Long
    Dim YPixelsPerInch As Long

    hDC = GetDC(0)
    XPixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    YPixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY)
    retVal = ReleaseDC(0, hDC)
    x = x * TWIPSPERINCH / 20 / XPixelsPerInch
    y = y * TWIPSPERINCH / 20 / YPixelsPerInch
End Sub

Sub FormShow(ByVal objForm As Object, ByVal rng As range)
    Dim L As Single, T As Single

    L = ActiveWindow.ActivePane.PointsToScreenPixelsX(rng.Left + 100)
    T = ActiveWindow.ActivePane.PointsToScreenPixelsY(rng.Top + rng.Height)
    ConvertPixelsToPoints L, T

    With objForm
       .StartUpPosition = 0
       .Left = L
       .Top = T
       .Show vbModeless
    End With
    AppActivate Application.caption
End Sub
Sub FormShowFixed(ByVal objForm As Object)
    Dim L As Single, T As Single


    With objForm
       .StartUpPosition = 0
       .Left = Application.Width - .Width - 100
       .Top = 0
       .Show vbModeless
    End With
    AppActivate Application.caption
End Sub
Sub ShowUserForm2()
Attribute ShowUserForm2.VB_ProcData.VB_Invoke_Func = " \n14"
 FormShow UserForm2, Selection.Cells(0)
End Sub

Sub ShowUserForm2_Center()
Attribute ShowUserForm2_Center.VB_ProcData.VB_Invoke_Func = "b\n14"
    Unload UserForm2
    UserForm2.Show vbModeless
End Sub

