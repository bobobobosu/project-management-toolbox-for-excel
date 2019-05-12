Attribute VB_Name = "References"
Function use_Structured(cell As Variant, mode As Integer)
    If mode = 0 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[編號]").Column))
    ElseIf mode = 1 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[完整耗時]]").Column))
    ElseIf mode = 2 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[預計耗時]").Column))
    ElseIf mode = 3 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[實際耗時]").Column))
    ElseIf mode = 4 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Start Date]").Column))
    ElseIf mode = 5 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[End Date]").Column))
    ElseIf mode = 6 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[交易物件]").Column))
    ElseIf mode = 7 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[起始百分比]").Column))
    ElseIf mode = 8 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[預計百分比]").Column))
    ElseIf mode = 9 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Description]").Column))
    ElseIf mode = 10 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[ID]").Column))
    ElseIf mode = 11 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Concurrency]").Column))
    ElseIf mode = 12 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[實際百分比]").Column))
    ElseIf mode = 13 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[WBS]").Column))
    ElseIf mode = 14 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Task Chain]").Column))
    ElseIf mode = 15 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[進度]").Column))
    ElseIf mode = 16 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[專案累積SU-MIN]").Column))
    ElseIf mode = 17 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[本專案累積SU-MIN]").Column))
    ElseIf mode = 18 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[所屬專案]").Column))
    ElseIf mode = 19 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[SU]").Column))
    ElseIf mode = 20 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Location]").Column))
    ElseIf mode = 21 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Start Time]").Column))
    ElseIf mode = 22 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[End Time]").Column))
    ElseIf mode = 23 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Buffer]").Column))
    ElseIf mode = 24 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[期限]").Column))
    ElseIf mode = 25 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Dependency]").Column))
    ElseIf mode = 26 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[note]").Column))
    ElseIf mode = 27 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[剩餘時間]").Column))
    ElseIf mode = 28 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[現在預計進度]").Column))
    ElseIf mode = 29 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[至完成還有]").Column))
    ElseIf mode = 30 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[已耗時]").Column))
    ElseIf mode = 31 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[已節省]").Column))
    ElseIf mode = 32 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Subject]").Column))
    ElseIf mode = 33 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Certainty]").Column))
    ElseIf mode = 34 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Latitude]").Column))
    ElseIf mode = 35 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Longitude]").Column))
    ElseIf mode = 36 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Location Verify]").Column))
    ElseIf mode = 37 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Dependency Verify]").Column))
    ElseIf mode = 38 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Order]").Column))
    ElseIf mode = 39 Then
        use_Structured = AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[時區]").Column))
    ElseIf mode = 40 Then '預計 End Time
        use_Structured = range(AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[Start Date]").Column))).Value + _
                        range(AddressEx(Sheets("交易").Cells(cell.Row, range("表格2[預計耗時]").Column))).Value
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
