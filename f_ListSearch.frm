VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_ListSearch 
   Caption         =   "List Search - ExcelCampus.com"
   ClientHeight    =   440
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8028
   OleObjectBlob   =   "f_ListSearch.frx":0000
End
Attribute VB_Name = "f_ListSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public gdFORMWIDTH As Double
Public gsLISTSOURCE As String


Private Sub CommandButton_Input_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub CommandButton_Input_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

End Sub

Private Sub UserForm_Initialize()

Dim dWidth As Double

    'Set Initial value
    Me.ComboBox_Search.Value = Selection(1).Value2

    'Check if cell is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell in a worksheet before opening List Search.", vbCritical, "Error Opening List Search"
        End
    End If
    
    'Load combobox with all values
    Call Search_List("")
    
    'Load direction box
    With Me.ComboBox_Direction
        .AddItem "Down"
        .AddItem "Right"
        .AddItem "None"
        .AddItem "Close"
        .AddItem "Paste"
        Application.EnableEvents = False
        Call m_Settings.GetDirection
        .ListIndex = gsDirection
        Application.EnableEvents = True
    End With
    
    'Resize form and combobox
    Me.Width = 230.25
    dWidth = ActiveCell.Width
    If dWidth > 175 Then
        Me.ComboBox_Search.Width = ActiveCell.Width
        Me.Frame_Options.Left = dWidth + 2
        Me.Width = Me.ComboBox_Search.Width + 64 + 17
    End If
    
    gdFORMWIDTH = Me.Width
    
    'Set the state of the toggle buttons from saved registry settings
    Call m_Settings.GetShowMenu
    Me.ToggleButton_Menu.Value = gsMenu
    
    Call m_Settings.GetOpenOnSelectionChange
    Me.ToggleButtonAutoOpen.Value = gsOpen
    
    'Move form near activecell
    Call Move_Form
    
    With ComboBox_Search
        .SelStart = 0
        .SelLength = Len(.text)
    End With
    
End Sub


Private Sub ComboBox_Search_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    If KeyCode = 13 Then 'Enter button pressed, select search field from list box
        Call Input_Value
    End If

End Sub



Private Sub ComboBox_Search_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    With Me.ComboBox_Search
    
        If KeyCode = 16 Then Exit Sub  'Shift is pressed by itself
        
        If KeyCode <> 38 And KeyCode <> 40 And KeyCode <> 13 Then 'NOT Up Arrow, Down Arrow, Enter
        
            Call Search_List(.Value)
            
            If Len(.Value) > 0 Then
                .DropDown
            End If
        
        End If
        
        If KeyCode = 13 Then 'Enter button pressed, select search field from list box
            'Call Input_Value
        End If
        
        If KeyCode = 27 Then 'Esc key pressed, clear input
            If .Value = "" Then
                Unload Me
            Else
                .Value = ""
                .SetFocus
                Call Search_List(.Value)
            End If
           
        End If
    End With

End Sub



Sub Search_List(sSearch As String, Optional sSort As String)
'Purpose: Search pivot fiels for all in-string matches (Instr)
'           Called by: ComboBox_Search_KeyUp

Dim vResults() As Variant
Dim lCount As Long
Dim vArray() As Variant
Dim lResultsCount As Long
Dim sFormula As String
Dim sSheet As String
Dim rCurrent As range
Dim vbAnswer As VbMsgBoxResult
Dim sTable As String
Dim lColumn As Long
Dim rFormula As range
Dim wsSource As Worksheet
Dim sArray() As String

    'Load the array with validation list or range
    
    'Get the sheet name and range for validation source reference
    
    On Error Resume Next
        sFormula = ActiveCell.Validation.formula
    On Error GoTo 0
    
    If sFormula <> "" Then 'Cell has validation list
    
        If Left(sFormula, 1) = "=" Then 'validation formula based
            On Error Resume Next
                Set rFormula = Evaluate(ActiveCell.Validation.Formula1)
                If rFormula Is Nothing Then
                    MsgBox "There is an error with the validation formula.  " & _
                            "Please fix the error in the Data Validation window.", _
                            vbOKOnly, "Error Evaluating Validation Formula"
                    End
                End If
            On Error GoTo 0
            
            'Check if data validation range only contains 1 cell
            If rFormula.Cells.Count = 1 Then
'                ReDim vArray(1 To 1)
'                vArray(1) = rFormula.Value
                vArray = Application.WorksheetFunction.index(Split(rFormula.Value, ";"), 1, 0)
            'Add the range to an array
            Else
              vArray = rFormula.Value
              'Convert to 1D array
              With Application.WorksheetFunction
                  If UBound(vArray, 1) = 1 Then 'Horizontal data validation range
                    vArray = .index(vArray, 1, 0)
                  Else 'Vertical data validation range
                    vArray = .Transpose(.index(vArray, 0, 1))
                  End If
              End With
            End If
            gsLISTSOURCE = "List Type: Validation" & vbNewLine & "List Source: " & rFormula.Parent.name & "!" & rFormula.address
        
        Else 'text based list
            sArray = Split(sFormula, ",")
            'Convert to variant
            With Application.WorksheetFunction
                vArray = .index(sArray, 1, 0)
            End With
        End If
        
    Else 'Find used range in column
        'Determine if activecell is in a Table
        On Error Resume Next
            sTable = ActiveCell.ListObject.name
        On Error GoTo 0
        
        If sTable <> "" Then
            lColumn = ActiveCell.Column - ActiveSheet.ListObjects(sTable).range(, 1).Column + 1
            Set rCurrent = ActiveSheet.ListObjects(sTable).ListColumns(lColumn).DataBodyRange
        Else
            Set rCurrent = Intersect(ActiveCell.CurrentRegion, ActiveCell.EntireColumn)
            sSheet = ""
        End If
        
        'Get uniques
        If rCurrent.Cells.Count = 1 Then
            MsgBox "Please select a cell that is not blank or contains a validation list.", vbOKOnly, "Error Creating List"
            Unload Me
            End
        Else
            vArray = rCurrent.Value
            vArray = UniqueArray(vArray)
            gsLISTSOURCE = "List Type: Range" & vbNewLine & "Range: " & rCurrent.address
            'Convert to 1D array
            With Application.WorksheetFunction
                vArray = .Transpose(.index(vArray, 0, 1))
            End With
        End If
    End If
    
    
    'Search list and add results to array
    ReDim vResults(1 To 2)
    For lCount = LBound(vArray) To UBound(vArray)
        If InStr(1, CStr(vArray(lCount)), sSearch, 1) Then
            lResultsCount = lResultsCount + 1
            ReDim Preserve vResults(1 To lResultsCount)
            vResults(lResultsCount) = vArray(lCount)
        End If
    Next lCount
    
    'Sort the array
    If sSort <> "" Then
        If UBound(vResults) > 2000 Then
            vbAnswer = MsgBox("Lists that contain more than 2,000 items may take additional time to sort.  Do you want to continue?", _
                        vbYesNo, "List Search Sort Warning")
            If vbAnswer = vbYes Then
                vResults = SortArray(vResults, sSort)
            End If
        Else
            vResults = SortArray(vResults, sSort)
        End If
    End If
    
    'Populate the combobox with array
    Me.ComboBox_Search.List() = vResults

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    ' YOUR CODE HERE (Just copy whatever the close button does)

    If CloseMode = vbFormControlMenu Then
        Me.ComboBox_Search.Value = vbNullString
    End If

End Sub

Private Sub CommandButton_Input_Click()
'Change the value of the activecell to the selected value

    
    Call Input_Value

End Sub
Function Get_Input() As String
    Me.Show
    Get_Input = Me.ComboBox_Search.Value
End Function

Sub Input_Value()
Unload Me
Exit Sub

'Input the selected value to the worksheet

Dim lCnt As Long
Dim bExists As Boolean
Dim rActive As range
Dim vbAnswer As VbMsgBoxResult

    'Validate entry
    With Me.ComboBox_Search
        For lCnt = 0 To .ListCount - 1
            If .Value = CStr(.List(lCnt)) Then
                bExists = True
                Exit For
            End If
        Next lCnt
    End With
    
    If bExists Then

        'Input value and select next cell
        If Me.ComboBox_Direction = "Paste" Then
            Unload Me
            ClipBoard_SetData Me.ComboBox_Search.Value
            Application.SendKeys "^v{numlock}"
            Exit Sub
        Else
            
            If Selection.MergeCells = True Then
                Selection.Value = Me.ComboBox_Search.Value
            ElseIf Selection.Cells.Count > 1 Then
                vbAnswer = MsgBox("Mulitple cells are selected.  " & _
                                    "Do you want to fill all selected cells with the input value?", _
                                    vbYesNo, "Fill All Selected Cells")
                If vbAnswer = vbYes Then
                    Selection.Value = Me.ComboBox_Search.Value
                Else
                    ActiveCell.Value = Me.ComboBox_Search.Value
                End If
            Else
                ActiveCell.Value = Me.ComboBox_Search.Value
            End If
            
            Select Case Me.ComboBox_Direction
                Case "Down"
                    Set rActive = ActiveCell
                    Do
                        Set rActive = rActive(2, 1)
                    Loop Until rActive.EntireRow.Hidden = False
                    If gsOpen Then
                        Unload Me
                    End If
                    rActive.Select
                    
                    Call Move_Form
                    Me.ComboBox_Search.Value = ""
                    Call Search_List("")
                
                Case "Right"
                    Set rActive = ActiveCell
                    Do
                        Set rActive = rActive(1, 2)
                    Loop Until rActive.EntireColumn.Hidden = False
                    If gsOpen Then
                        Unload Me
                    End If
                    rActive.Select
                    
                    Call Move_Form
                    Me.ComboBox_Search.Value = ""
                    Call Search_List("")
                
                Case "Close"
                    Unload Me
                    Exit Sub
                    
                Case "None"
                    
            End Select
            
        End If
        
    Else
        MsgBox "The value in the search box does not match a value in the list.  Please select a value from the list.", _
                vbOKOnly, "Validation Error"
    End If

End Sub

Private Sub CommandButton_Clear_Click()
    Me.ComboBox_Search.Value = ""
    Call Search_List(Me.ComboBox_Search.Value)
    Me.ComboBox_Search.SetFocus
End Sub

Private Sub ToggleButton_Menu_Click()

    If Me.ToggleButton_Menu.Value = True Then
        Me.Width = Me.Width + Me.Frame_Options.Width - 66
    Else
        Me.Width = gdFORMWIDTH
    End If
    
    Call m_Settings.SaveShowMenu(Me.ToggleButton_Menu.Value)
    
End Sub


Private Sub ToggleButton_AZ_Click()
    If ToggleButton_AZ.Value = True Then
        Call Search_List(Me.ComboBox_Search.Value, "Asc")
        ToggleButton_ZA.Value = False
        ToggleButton_Orig.Value = False
        Application.EnableEvents = False
            Me.ComboBox_Search.DropDown
        Application.EnableEvents = True
    End If
End Sub

Private Sub ToggleButton_ZA_Click()
    If ToggleButton_ZA.Value = True Then
        Call Search_List(Me.ComboBox_Search.Value, "Desc")
        ToggleButton_AZ.Value = False
        ToggleButton_Orig.Value = False
        Application.EnableEvents = False
            Me.ComboBox_Search.DropDown
        Application.EnableEvents = True
    End If
End Sub

Private Sub ToggleButton_Orig_Click()
    If ToggleButton_Orig.Value = True Then
        Call Search_List(Me.ComboBox_Search.Value)
        ToggleButton_AZ.Value = False
        ToggleButton_ZA.Value = False
        Application.EnableEvents = False
            Me.ComboBox_Search.DropDown
        Application.EnableEvents = True
    End If
End Sub

Private Sub ToggleButtonAutoOpen_Click()
'Store the Auto Open setting
    Call m_Settings.SaveOpenOnSelectionChange(ToggleButtonAutoOpen.Value)
    'Call m_Ribbon.Set_Handler
    Call m_Ribbon.Set_App
End Sub


Sub ComboBox_Direction_Change()
'Store the selected option
    Call m_Settings.SaveDirection(Me.ComboBox_Direction.ListIndex)
End Sub

Private Sub CommandButton_Info_Click()
    MsgBox gsLISTSOURCE & vbNewLine & _
           "List Count: " & Me.ComboBox_Search.ListCount, _
            vbOKOnly, "List Info - List Search - Version 1.1"
End Sub

Private Sub CommandButton_Copy_Click()
'Copy the searchbox list to the clipboard

Dim lCnt As Long
Dim sList As String
Dim sArray() As String

    ReDim sArray(0 To Me.ComboBox_Search.ListCount - 1)
    For lCnt = 0 To Me.ComboBox_Search.ListCount - 1
        sArray(lCnt) = Me.ComboBox_Search.List(lCnt)
    Next lCnt
    
    sList = Join(sArray, vbNewLine)
    ClipBoard_SetData sList
    
    Unload Me
    
    MsgBox "The contents of the drop-down list has been copied to the clipboard.  Select a cell and press Ctrl+V or right-click > Paste to paste the list to the range.", _
            vbOKOnly, "List Copied to Clipboard"
    
End Sub

Sub Move_Form()

Dim dTop As Double
Dim dLeft As Double

        'Center form
'        Me.StartUpPosition = 1
        
        'Move form near activecell
        Me.StartUpPosition = 0
        dTop = ActiveCell.Top - ActiveWindow.VisibleRange.Top + 200
        If dTop > ActiveWindow.Top And dTop < (ActiveWindow.Top + ActiveWindow.Height) Then
            Me.Top = dTop
        Else
            Me.Top = ActiveWindow.Top + (ActiveWindow.Height / 2)
        End If

        dLeft = ActiveCell.Left - ActiveWindow.VisibleRange.Left + 20
        If dLeft > ActiveWindow.Left And dLeft < (ActiveWindow.Left + ActiveWindow.Width) Then
            Me.Left = dLeft
        Else
            Me.Left = ActiveWindow.Left + (ActiveWindow.Width / 2)
        End If
    
    
End Sub

Function SortArray(vArray As Variant, sOrder As String) As Variant
'Sort the array

Dim l1 As Long
Dim l2 As Long
Dim s1 As Variant
Dim s2 As Variant

    If sOrder = "Asc" Then
        For l1 = LBound(vArray) To UBound(vArray)
            For l2 = l1 To UBound(vArray)
                If vArray(l2) < vArray(l1) Then
                    s1 = vArray(l1)
                    s2 = vArray(l2)
                    vArray(l1) = s2
                    vArray(l2) = s1
                End If
            Next l2
        Next l1
    Else
        For l1 = LBound(vArray) To UBound(vArray)
            For l2 = l1 To UBound(vArray)
                If vArray(l2) > vArray(l1) Then
                    s1 = vArray(l1)
                    s2 = vArray(l2)
                    vArray(l1) = s2
                    vArray(l2) = s1
                End If
            Next l2
        Next l1
    End If
    
    SortArray = vArray

End Function

Function UniqueArray(vArray As Variant) As Variant

Dim colUnique As Collection
Dim lCnt As Long
Dim vUnique() As Variant

    Set colUnique = New Collection
    
    On Error Resume Next
        For lCnt = LBound(vArray) To UBound(vArray)
            colUnique.Add vArray(lCnt, 1), CStr(vArray(lCnt, 1))
        Next lCnt
    On Error GoTo 0
    
    ReDim vUnique(1 To colUnique.Count, 1 To 1)

    For lCnt = 1 To colUnique.Count
        vUnique(lCnt, 1) = colUnique.Item(lCnt)
    Next lCnt
    
    UniqueArray = vUnique

End Function





