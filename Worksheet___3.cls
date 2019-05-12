VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Worksheet___3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal target As range)
    u = 0
    If Selection.Cells.Count = 1 And Application.CutCopyMode = False Then
        Set isect = Application.Intersect(range("表格2"), Selection)
        If isect Is Nothing Then
            
        Else
            If Not Application.Intersect(range("表格2[Description]"), Selection) Is Nothing Then
                range("表格2[交易物件]").Interior.ColorIndex = 0
                range("表格2[Description]").Interior.ColorIndex = 0
                Selection.Interior.ColorIndex = 37
                On Error Resume Next
                WordsList = json2arr(use_Structured2R(Selection, "表格2", "Dependency"))
                For i = 0 To UBound(WordsList)
                    If WordsList(i) <> vbNullString Then
                        On Error GoTo Error_handler:
                        WorksheetFunction.index(range("表格2[交易物件]"), WorksheetFunction.Match(val(WordsList(i)), range("表格2[ID]"), 0)).Interior.ColorIndex = 37
                    End If
                Next i
            End If
            If Not Application.Intersect(range("表格2[交易物件]"), Selection) Is Nothing Then
                range("表格2[編號]").Interior.ColorIndex = 0
                On Error Resume Next
                WordsList = Split(range(use_Structured(Selection, 11)).Value2)
                For i = 0 To UBound(WordsList)
                    If WordsList(i) <> vbNullString Then
                        On Error GoTo Error_handler:
                        WorksheetFunction.index(range("表格2[編號]"), WorksheetFunction.Match(val(WordsList(i)), range("表格2[ID]"), 0)).Interior.ColorIndex = 37
                    End If
                Next i
                
                range("表格2[交易物件]").Interior.ColorIndex = 0
                On Error Resume Next
                WordsList = Split(range(use_Structured(Selection, 14)).Value2, ",")
                For i = 0 To UBound(WordsList)
                    If WordsList(i) <> vbNullString Then
                        On Error GoTo Error_handler:
                        WorksheetFunction.index(range("表格2[交易物件]"), WorksheetFunction.Match(val(WordsList(i)), range("表格2[ID]"), 0)).Interior.ColorIndex = 37
                    End If
                Next i
                
                range("表格2[Description]").Interior.ColorIndex = 0
                On Error Resume Next
                WordsList = json2arr(use_Structured2R(Selection, "表格2", "Supply"))
                For i = 0 To UBound(WordsList)
                    If WordsList(i) <> vbNullString Then
                        On Error GoTo Error_handler:
                        WorksheetFunction.index(range("表格2[Description]"), WorksheetFunction.Match(val(WordsList(i)), range("表格2[ID]"), 0)).Interior.ColorIndex = 37
                    End If
                Next i
            End If
        End If
    End If
Error_handler:

    Dim nowRow As String
    On Error Resume Next
    nowRow = Evaluate("表格2[[#This Row], [ID]:[ID]]").Value2
    Worksheets("趨勢").range("B2").Value2 = nowRow
    
    
End Sub

