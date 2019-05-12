Attribute VB_Name = "xlwings_RunPython"
Sub TCmain()
    RunPython ("import TCmain; TCmain.llServer()")
End Sub

Sub SetInvisible()
    ThisWorkbook.Application.Visible = False
End Sub
Sub SetVisible()
    ThisWorkbook.Application.Visible = True
End Sub
Sub SetTempSheet()
    Worksheets("TEMP").Activate
End Sub
