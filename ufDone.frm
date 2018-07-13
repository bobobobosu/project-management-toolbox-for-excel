VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufDone 
   Caption         =   "UserForm3"
   ClientHeight    =   2863
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   4296
   OleObjectBlob   =   "ufDone.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "ufDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***** Paste this in the code section of a UserForm titled "ufDone" *****
Private Sub UserForm_Initialize()
Me.Hide
End Sub
 
Private Sub UserForm_Activate()
Dim IconPath As String
#If VBA7 Then
  Dim Me_hWnd As LongPtr, Me_Icon As Long, Me_Icon_Handle As LongPtr
#Else
  Dim Me_hWnd As Long, Me_Icon As Long, Me_Icon_Handle As Long
#End If
Me.Hide
RemoveIconFromTray
Unhook
IconPath = Application.path & Application.PathSeparator & "excel.exe"
Me_hWnd = FindWindowd("ThunderDFrame", Me.Caption)
Me_Icon_Handle = ExtractIcond(0, IconPath, 0)
Hook Me_hWnd
AddIconToTray Me_hWnd, 0, Me_Icon_Handle, ""
BalloonPopUp_1
Unload Me
End Sub
'*************************************************************************
