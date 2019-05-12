VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_OpenListSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents AppListSearch As Application
Attribute AppListSearch.VB_VarHelpID = -1

Private Sub Class_Initialize()
    
    Set AppListSearch = Application
    
End Sub


Private Sub AppListSearch_SheetSelectionChange(ByVal Sh As Object, ByVal target As range)
    
Dim lType As Long
    
    On Error Resume Next
        lType = target(1, 1).Validation.Type
    On Error GoTo 0
    
    If lType = 3 Then
        If Not m_Ribbon.IsUserFormLoaded("f_ListSearch") Then
            f_ListSearch.Show
        End If
    End If
    
End Sub






