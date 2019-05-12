Attribute VB_Name = "References"
Function use_Structured(cell As Variant, mode As Integer)
    If mode = 0 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�s��]").Column))
    ElseIf mode = 1 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[����Ӯ�]]").Column))
    ElseIf mode = 2 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�w�p�Ӯ�]").Column))
    ElseIf mode = 3 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[��گӮ�]").Column))
    ElseIf mode = 4 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Start Date]").Column))
    ElseIf mode = 5 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[End Date]").Column))
    ElseIf mode = 6 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�������]").Column))
    ElseIf mode = 7 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�_�l�ʤ���]").Column))
    ElseIf mode = 8 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�w�p�ʤ���]").Column))
    ElseIf mode = 9 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Description]").Column))
    ElseIf mode = 10 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[ID]").Column))
    ElseIf mode = 11 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Concurrency]").Column))
    ElseIf mode = 12 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[��ڦʤ���]").Column))
    ElseIf mode = 13 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[WBS]").Column))
    ElseIf mode = 14 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Task Chain]").Column))
    ElseIf mode = 15 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�i��]").Column))
    ElseIf mode = 16 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�M�ײֿnSU-MIN]").Column))
    ElseIf mode = 17 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[���M�ײֿnSU-MIN]").Column))
    ElseIf mode = 18 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[���ݱM��]").Column))
    ElseIf mode = 19 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[SU]").Column))
    ElseIf mode = 20 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Location]").Column))
    ElseIf mode = 21 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Start Time]").Column))
    ElseIf mode = 22 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[End Time]").Column))
    ElseIf mode = 23 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Buffer]").Column))
    ElseIf mode = 24 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[����]").Column))
    ElseIf mode = 25 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Dependency]").Column))
    ElseIf mode = 26 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[note]").Column))
    ElseIf mode = 27 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�Ѿl�ɶ�]").Column))
    ElseIf mode = 28 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�{�b�w�p�i��]").Column))
    ElseIf mode = 29 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�ܧ����٦�]").Column))
    ElseIf mode = 30 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�w�Ӯ�]").Column))
    ElseIf mode = 31 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�w�`��]").Column))
    ElseIf mode = 32 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Subject]").Column))
    ElseIf mode = 33 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Certainty]").Column))
    ElseIf mode = 34 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Latitude]").Column))
    ElseIf mode = 35 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Longitude]").Column))
    ElseIf mode = 36 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Location Verify]").Column))
    ElseIf mode = 37 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Dependency Verify]").Column))
    ElseIf mode = 38 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[Order]").Column))
    ElseIf mode = 39 Then
        use_Structured = AddressEx(Sheets("���").Cells(cell.Row, range("���2[�ɰ�]").Column))
    ElseIf mode = 40 Then '�w�p End Time
        use_Structured = range(AddressEx(Sheets("���").Cells(cell.Row, range("���2[Start Date]").Column))).Value + _
                        range(AddressEx(Sheets("���").Cells(cell.Row, range("���2[�w�p�Ӯ�]").Column))).Value
    End If
  
End Function

Function getStructuredByIDR(targetID As Variant, table As String, title As Variant)
    Set targetRange = Application.WorksheetFunction.index(range(table + "[ID]"), Application.Match(targetID, range(table + "[ID]"), 0))
    Set getStructuredByIDR = use_Structured2R(targetRange, table, title)
End Function
Function use_Structured2(cell As Variant, table As String, title As Variant)
    use_Structured2 = AddressEx(cell.Worksheet.Cells(cell.Row, range(table + "[" + title + "]").Column))
End Function
Function use_Structured2R(cell As Variant, table As String, title As Variant)
    Set use_Structured2R = (cell.Worksheet.Cells(cell.Row, range(table + "[" + title + "]").Column))
End Function
