Attribute VB_Name = "Module19"
Private mIntCutCopyMode As XlCutCopyMode
Private mRngClipboard As range

Public Sub subStoreClipboard()
    On Error GoTo ErrorHandler
    Dim wsActiveSource As Worksheet, wsActiveTarget As Worksheet
    Dim strClipboardRange As String

    mIntCutCopyMode = Application.CutCopyMode

    If Not fctBlnIsExcelClipboard Then Exit Sub


    Application.EnableEvents = False

    'Paste data as link
    Set wsActiveTarget = ActiveSheet
    Set wsActiveSource = ThisWorkbook.ActiveSheet

    With ws_Temp
        .Visible = xlSheetVisible
        .Activate
        .Cells(3, 1).Select
        On Error Resume Next
        .Paste Link:=True
        If err.Number Then
            err.Clear
            GoTo Finalize
        End If
        On Error GoTo ErrorHandler
    End With

    'Extract link from pasted formula and clear range
    With Selection
        strClipboardRange = Mid(.Cells(1, 1).formula, 2)
        If .Rows.Count > 1 Or .Columns.Count > 1 Then
            strClipboardRange = strClipboardRange & ":" & _
                Mid(.Cells(.Rows.Count, .Columns.Count).formula, 2)
        End If
        Set mRngClipboard = range(strClipboardRange)
        .Clear
     End With

Finalize:
    wsActiveSource.Activate
    wsActiveTarget.Parent.Activate
    wsActiveTarget.Activate

    ws_Temp.Visible = xlSheetVeryHidden
    Application.EnableEvents = True

    Exit Sub
ErrorHandler:
    err.Clear
    Resume Finalize
End Sub


Public Sub subRestoreClipboard()
    Select Case mIntCutCopyMode
        Case 0:
        Case xlCopy: mRngClipboard.Copy
        Case xlCut:  mRngClipboard.Cut
    End Select

End Sub

Private Function fctBlnIsExcelClipboard() As Boolean
    Dim var As Variant
    fctBlnIsExcelClipboard = False
    'check if clipboard is in use
    If mIntCutCopyMode = 0 Then Exit Function
    'check if Excel data is in clipboard
    For Each var In Application.ClipboardFormats
        If var = xlClipboardFormatCSV Then
            fctBlnIsExcelClipboard = True
            Exit For
        End If
    Next var
End Function

