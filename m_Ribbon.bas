Attribute VB_Name = "m_Ribbon"
Private XLAppListSearch As c_OpenListSearch

Sub Auto_Open()

    Call Set_App

End Sub

Sub Auto_Close()

    Set XLAppListSearch = Nothing

End Sub
Sub Set_App()

    m_Settings.GetOpenOnSelectionChange
    
    If gsOpen Then
        Set XLAppListSearch = New c_OpenListSearch
    Else
        Set XLAppListSearch = Nothing
    End If

End Sub

Sub ShowSearch()
Attribute ShowSearch.VB_ProcData.VB_Invoke_Func = " \n14"
'Assign a keyboard shortcut to this procedure
'to open the form.

    f_ListSearch.Show

End Sub

Sub btnListSearch_onAction(control As IRibbonControl)
    'Callback for ListSearch onAction
    
    f_ListSearch.Show
End Sub

Sub btnListSearchHelp_onAction(control As IRibbonControl)
    'Callback for ListSearchHelp onAction
    
    ThisWorkbook.FollowHyperlink "https://www.excelcampus.com/list-search-help"

End Sub

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
     
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function 'IsUserFormLoaded






