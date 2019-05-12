VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnterProgress 
   Caption         =   "Enter Progress"
   ClientHeight    =   4774
   ClientLeft      =   72
   ClientTop       =   300
   ClientWidth     =   6552
   OleObjectBlob   =   "EnterProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EnterProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public deltaTarget As Double
Public absTarget As Double
Public firstTrigger As Boolean
    
Private Sub CommandButton1_Click()
    milestone = ""
    If TextBox3.Value <> vbNullString Then
        Call AddRealRecord(CDbl(dateValue(TextBox3.Value) + TimeValue(TextBox3.Value)), CDbl(TextBox5.Value))
        range("NowPercent").Calculate
    Else
        Call AddRealRecord(TextBox3.Value, CDbl(TextBox5.Value))
        range("NowPercent").Calculate
    End If
   Call UpdateChart
    'Unload Me
End Sub

Private Sub CommandButton2_Click()
    Call AddRealRecordNowByOne
    range("NowPercent").Calculate
    Call UpdateChart
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal data As MSForms.DataObject, ByVal x As Single, ByVal y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub TextBox1_Change()
    On Error Resume Next
    If firstTrigger = True Then Exit Sub
     TextBox5.Value = TextBox1.Value - getCurrentActual()
End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()
    On Error Resume Next
    If firstTrigger = True Then Exit Sub
    On Error Resume Next
    Label7.caption = Labetl7Caption()
    'TextBox5.Value = TextBox4.Value / getMulti2Target()
    Call updateByTime
End Sub
Sub updateByTime()
    Call UpdateThisTask(True)
    Me.caption = range("進度!G4").Value2
    deltaTarget = Round(getPlannedByTime(CDbl(dateValue(TextBox3.Value) + TimeValue(TextBox3.Value))) - getCurrentActual())
    absTarget = Round(getPlannedByTime(CDbl(dateValue(TextBox3.Value) + TimeValue(TextBox3.Value))))
    TextBox1.Value = absTarget
    TextBox5.Value = deltaTarget
End Sub

Private Sub TextBox3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    TextBox3.Value = Format(Now(), "m/d/yyyy h:mm:ss;@")
End Sub

Private Sub TextBox5_Change()
    If firstTrigger = True Then Exit Sub
    On Error Resume Next
    Label7.caption = Labetl7Caption()
    On Error Resume Next
    TextBox1.Value = TextBox5.Value + getCurrentActual()

End Sub
Function Labetl7Caption()
    ThisTaskCount = use_Structured2R(getCurrentMilestone(), "NowPercent", "Task Count").Value2
    Planned = use_Structured2R(getCurrentMilestone(), "NowPercent", "Planned").Value2
    current = getCurrentActual()
    ThisTaskCompleted = Round(ThisTaskCount - Planned + current)
    Labetl7Caption = CStr(ThisTaskCompleted) + "/" + CStr(ThisTaskCount) + _
                            " (" + CStr(Planned - current) + ")"
End Function


Private Sub UserForm_Activate()
    Me.caption = range("進度!G4").Value2
    Call DynamicChartScale
    firstTrigger = True
    deltaTarget = Round(getPlannedByTime(CDbl(dateValue(Now()) + TimeValue(Now()))) - getCurrentActual())
    absTarget = Round(getPlannedByTime(CDbl(dateValue(Now()) + TimeValue(Now()))))
    TextBox1.Value = absTarget
    TextBox3.Value = Format(Now(), "m/d/yyyy h:mm:ss;@")
    TextBox5.Value = deltaTarget
    Label7.caption = Labetl7Caption()
'    Call Insert_Items_To_ListBox1
    TextBox1.SetFocus
    firstTrigger = False
    Call UpdateChart
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub Insert_Items_To_ListBox1()
    ListBox1.Clear
    Dim milestones As range
    Set milestones = range("NowPercent[Milestone]")
    Dim timestamps As range
    Set timestamps = range("NowPercent[Time]")
    
    Count = 0
    For Each cell In milestones
        If cell.Value <> vbNullString Then ListBox1.AddItem cell.Value
        Count = Count + 1
    Next
End Sub

Private Sub UserForm_Initialize()
    
End Sub

Private Sub UpdateChart()
    Dim Fname As String
    Dim MyChart As Chart
    Set MyChart = Worksheets("進度").ChartObjects("Chart 8").Chart
    Fname = ThisWorkbook.path & "\temp1.bmp"
    MyChart.Export filename:=Fname, FilterName:="BMP"
    Fname = ThisWorkbook.path & "\temp1.bmp"
    Me.Image1.Picture = LoadPicture(Fname)
End Sub





