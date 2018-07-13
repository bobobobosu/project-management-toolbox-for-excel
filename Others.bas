Attribute VB_Name = "Others"
Sub ReOpen()
    Application.DisplayAlerts = False
    Workbooks.Open ActiveWorkbook.path & "\" & ActiveWorkbook.Name
    Application.DisplayAlerts = True
End Sub
