Attribute VB_Name = "WindowsNotification"
'***** Paste this code into your standard Module *****
Option Explicit
#If VBA7 Then
    Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Declare PtrSafe Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As LongPtr
    Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    #If Win64 Then
        'Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #End If
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Declare PtrSafe Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As LongPtr
    Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private FHandle As LongPtr
    Private WndProc As LongPtr
#Else
    Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
    Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    Declare Function GetForegroundWindow Lib "user32" () As Long
    Private FHandle As Long
    Private WndProc As Long
#End If
Public myMessage As String

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBL = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBL = &H206
Public Const WM_ACTIVATEAPP = &H1C
 
Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
Public Const NIF_GUID = &H20
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const MAX_TOOLTIP As Integer = 128
Public Const GWL_WNDPROC = (-4)
 
'shell version / NOTIFIYICONDATA struct size constants
Public Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Public Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Public Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
 
Public nfIconData As NOTIFYICONDATA
 
' list the icon types for the balloon message..
Public Const vbNone = 0
Public Const vbInformation = 1
Public Const vbExclamation = 2
Public Const vbCritical = 3

Private Hooking As Boolean
Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Type NOTIFYICONDATA
       cbSize As Long
       hWnd As LongPtr
       uID As Long
       uFlags As Long
       uCallbackMessage As Long
       hIcon As LongPtr
       szTip As String * 128
       dwState As Long
       dwStateMask As Long
       szInfo As String * 256
       uTimeout As Long
       szInfoTitle As String * 64
       dwInfoFlags As Long
       guidItem As GUID
    End Type
#Else
    Type NOTIFYICONDATA
       cbSize As Long
       hWnd As Long
       uID As Long
       uFlags As Long
       uCallbackMessage As Long
       hIcon As Long
       szTip As String * 128
       dwState As Long
       dwStateMask As Long
       szInfo As String * 256
       uTimeout As Long
       szInfoTitle As String * 64
       dwInfoFlags As Long
       guidItem As GUID
    End Type
#End If
Public Sub Unhook()
  If Hooking = True Then
    #If VBA7 Then
      SetWindowLongPtr FHandle, GWL_WNDPROC, WndProc
    #Else
      SetWindowLong FHandle, GWL_WNDPROC, WndProc
    #End If
    Hooking = False
  End If
End Sub

Public Sub RemoveIconFromTray()
Shell_NotifyIcon NIM_DELETE, nfIconData
End Sub

#If VBA7 Then
    Function FindWindowd(ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    FindWindowd = FindWindow(lpClassName, lpWindowName)
    End Function
    
    Function ExtractIcond(ByVal hInst As LongPtr, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As LongPtr
    ExtractIcond = ExtractIcon(hInst, lpszExeFileName, nIconIndex)
    End Function
    
    Public Sub Hook(Lwnd As LongPtr)
      If Hooking = False Then
        FHandle = Lwnd
        WndProc = SetWindowLongPtr(Lwnd, GWL_WNDPROC, AddressOf WindowProc)
        Hooking = True
      End If
    End Sub
    
    Public Function WindowProc(ByVal hw As LongPtr, ByVal uMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
      If Hooking Then
          If lParam = WM_LBUTTONDBL Then
           ufDone.Show 1
           WindowProc = True
        ' Unhook
           Exit Function
          End If
          WindowProc = CallWindowProc(WndProc, hw, uMsg, wParam, lParam)
      End If
    End Function
 
    Public Sub AddIconToTray(MeHwnd As LongPtr, MeIcon As Long, MeIconHandle As LongPtr, Tip As String)
        With nfIconData
          .hWnd = MeHwnd
          .uID = MeIcon
          .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP Or NIF_GUID
          .uCallbackMessage = WM_RBUTTONUP
          .dwState = NIS_SHAREDICON
          .hIcon = MeIconHandle
          .szTip = Tip & Chr$(0)
          .cbSize = NOTIFYICONDATA_V3_SIZE
        End With
        Shell_NotifyIcon NIM_ADD, nfIconData
    End Sub

    Public Sub MacroFinished(myMessagef As String)
        myMessage = myMessagef
        Dim wstate As Long
        Dim hwnd2 As LongPtr
        wstate = Application.WindowState
        hwnd2 = GetForegroundWindow()   'find the current window
        'AppActivate (ThisWorkbook2.Name) 'flash your existing workbook
        ufDone.Show 1                   'so the notification tray shows
        Application.WindowState = wstate
        SetForegroundWindow (hwnd2)     'then, bring the original window back to the front
    End Sub
#Else
    Function FindWindowd(ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    FindWindowd = FindWindow(lpClassName, lpWindowName)
    End Function
    
    Function ExtractIcond(ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
    ExtractIcond = ExtractIcon(hInst, lpszExeFileName, nIconIndex)
    End Function
    
    Public Sub Hook(Lwnd As Long)
    If Hooking = False Then
      FHandle = Lwnd
      WndProc = SetWindowLong(Lwnd, GWL_WNDPROC, AddressOf WindowProc)
      Hooking = True
    End If
    End Sub
    
    Public Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
      If Hooking Then
          If lParam = WM_LBUTTONDBL Then
           ufDone.Show 1
           WindowProc = True
        ' Unhook
           Exit Function
          End If
          WindowProc = CallWindowProc(WndProc, hw, uMsg, wParam, lParam)
      End If
    End Function
    
    Public Sub AddIconToTray(MeHwnd As Long, MeIcon As Long, MeIconHandle As Long, Tip As String)
        With nfIconData
          .hWnd = MeHwnd
          .uID = MeIcon
          .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP Or NIF_GUID
          .uCallbackMessage = WM_RBUTTONUP
          .dwState = NIS_SHAREDICON
          .hIcon = MeIconHandle
          .szTip = Tip & Chr$(0)
          .cbSize = NOTIFYICONDATA_V3_SIZE
        End With
        Shell_NotifyIcon NIM_ADD, nfIconData
    End Sub
    
    Public Sub MacroFinished(myMessagef As String)
        Dim wstate As Long
        Dim hwnd2 As Long
        wstate = Application.WindowState
        hwnd2 = GetForegroundWindow()   'find the current window
        'AppActivate (ThisWorkbook2.Name) 'flash your existing workbook
        ufDone.Show 1                   'so the notification tray shows
        Application.WindowState = wstate
        SetForegroundWindow (hwnd2)     'then, bring the original window back to the front
    End Sub
#End If
 
Public Sub BalloonPopUp_1()
    With nfIconData
        .dwInfoFlags = vbInformation
        .uFlags = NIF_INFO
        .szInfoTitle = ActiveWorkbook.Name & vbNullChar
        .szInfo = myMessage '"Your Macro Has Finished." & vbNullChar
    End With
   
    Shell_NotifyIcon NIM_MODIFY, nfIconData
End Sub
'********************************************
