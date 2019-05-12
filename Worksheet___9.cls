VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Worksheet___9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal target As range)
    Dim nowRow As String
    On Error Resume Next
    nowRow = Evaluate("ªí®æ68[[#This Row], [ID]:[ID]]").Value2
    Worksheets("ÁÍ¶Õ").range("B2").Value2 = nowRow
    
    
'    If Worksheets("ÁÍ¶Õ").Range("A2").Value2 = False And Selection.Cells.Count = 1 And Application.CutCopyMode = False Then
'
'
'
'        Set isect = Application.Intersect(Range(Evaluate("INDIRECT(""$B$2"")")), Selection)
'        If isect Is Nothing Then
'
'        Else
'                    'Updateby Extendoffice 20160704
'            Application.ScreenUpdating = False
'
'            Dim rng As Range
'            Dim cFnd As String
'            Dim xTmp As String
'            Dim x As Long
'            Dim m As Long
'            Dim y As Long
'            On Error Resume Next
'            cFnd = Selection.Value
'            y = Len(cFnd)
'            For Each rng In Range(Evaluate("INDIRECT(""$B$2"")"))
'              'Selection.Font.Size = 12
'              'Rng.Font.Size = 12
'              rng.Font.Bold = False
'              With rng
'                m = UBound(Split(rng.Value, cFnd))
'                If m > 0 Then
'                  xTmp = ""
'                  For x = 0 To m - 1
'                    xTmp = xTmp & Split(rng.Value, cFnd)(x)
'                    .Characters(start:=Len(xTmp) + 1, Length:=y).Font.Bold = True
'                    xTmp = xTmp & cFnd
'                  Next
'                End If
'              End With
'            Next rng
'            Application.ScreenUpdating = True
'        End If
'    End If
End Sub

