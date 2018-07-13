VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Change Date"
   ClientHeight    =   2028
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   2364
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Range("ÁÍ¶Õ!I2").Value2 = DateClicked
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.MonthView1.Value = Format(Selection(1).Value2, "m/d/yy")
    Range("ÁÍ¶Õ!I2").Value2 = vbNullString
End Sub
