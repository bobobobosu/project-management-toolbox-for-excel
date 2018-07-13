Attribute VB_Name = "Module36"

Sub AddDataRow(tableName As String, values() As Variant)
    Dim sheet As Worksheet
    Dim Table As ListObject
    Dim col As Integer
    Dim lastRow As Range

    Set sheet = ActiveWorkbook.Worksheets("�Ͷ�")
    Set Table = sheet.ListObjects.Item(tableName)

    'First check if the last row is empty; if not, add a row
    If Table.ListRows.count > 0 Then
        Set lastRow = Table.ListRows(Table.ListRows.count).Range
        For col = 1 To lastRow.Columns.count
            If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
                Table.ListRows.Add
                Exit For
            End If
        Next col
    Else
        Table.ListRows.Add
    End If

    'Iterate through the last row and populate it with the entries from values()
    Set lastRow = Table.ListRows(Table.ListRows.count).Range
    For col = 1 To lastRow.Columns.count
        If col <= UBound(values) + 1 Then lastRow.Cells(1, col) = values(col - 1)
    Next col
End Sub

Sub SortPercent()
Set sheet = ActiveWorkbook.Worksheets("�Ͷ�")
Set mTable = sheet.ListObjects("NowPercent")


Set sortcolumn = Range("NowPercent[Time]")
    With mTable.Sort
       .SortFields.Clear
       .SortFields.Add key:=sortcolumn, SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With


End Sub


Sub UpdateCurrentPercent()
Dim toAdd As Double
toAdd = InputBox("Enter your progress", "", Range("�Ͷ�!$A$4"))

        Range(Range("���!C1").Value2).Calculate
        Range("���!K2").Calculate
        Range("���!I2").Calculate
        Range("���!M2").Calculate
        Range(Range("���!C1").Value2).Calculate
If evals("MAX(NowPercent[Time])") < Now() Then
    Call ExportTable("NowPercent", "TCdata" & "\" & "NowPercent" & "\" _
    & Range("�Ͷ�!C4").Value2)
End If

'
Call SortPercent


Dim x(3)
x(0) = Now()
x(1) = evals("INDEX(���2[�i��],MATCH(���!$D$2,���2[Start Date]))*INDEX(���2[�{�b�w�p�i��],MATCH(���!$D$2,���2[Start Date]))")
x(2) = evals("INDEX(NowPercent[Actual],COUNTA(NowPercent[Actual])-1)") + toAdd
x(3) = x(2) - x(1)
AddDataRow "NowPercent", x
Call SortPercent

If evals("MAX(NowPercent[Time])") < Now() Then
    Call CleanPercentTable
End If

Worksheets("�Ͷ�").Calculate
End Sub
Sub exportNowPercent()
    Call ExportTable("NowPercent", "TCdata" & "\" & "NowPercent" & "\" _
    & Range("�Ͷ�!C4").Value2)

End Sub

Sub CleanPercentTable()



'Call DeleteTableRows("NowPercent", True)
Range("NowPercent").Value = ""
Dim x(3)
x(0) = evals("=INDEX((���2[Start Date]),MATCH(���!$D$2,���2[Start Date],1))")
x(1) = evals("=INDEX((���2[�i��]),MATCH(���!$D$2,���2[Start Date],1))*INDEX((���2[�_�l�ʤ���]),MATCH(���!$D$2,���2[Start Date],1))")
x(2) = evals("=INDEX((���2[�i��]),MATCH(���!$D$2,���2[Start Date],1))*INDEX((���2[�_�l�ʤ���]),MATCH(���!$D$2,���2[Start Date],1))")
x(3) = ""
AddDataRow "NowPercent", x
Call SortPercent
Dim y(2)
y(0) = evals("=INDEX((���2[End Date]),MATCH(���!$D$2,���2[Start Date]))")
y(1) = "=Value(INDEX(���2[�i��],MATCH(���!$D$2,���2[Start Date])))"
y(2) = "=Value(INDEX(���2[�i��],MATCH(���!$D$2,���2[Start Date])))"
x(3) = ""
AddDataRow "NowPercent", y
Call SortPercent
Range(Range("���!$C$1").Value).Calculate
Worksheets("�Ͷ�").Calculate
End Sub
Sub DeleteTableRows(ByVal tableName As String, KeepFormulas As Boolean)

Set sheet = ActiveWorkbook.Worksheets("�Ͷ�")
Set Table = sheet.ListObjects.Item(tableName)


On Error Resume Next

If Not KeepFormulas Then
    Table.DataBodyRange.ClearContents
End If

Table.DataBodyRange.Rows.Delete

On Error GoTo 0

End Sub


Sub ExportTable(tableName As String, FileName As String)

    Dim WB As Workbook, wbNew As Workbook
    Dim ws As Worksheet, wsNew As Worksheet
    Dim wbNewName As String


   Set WB = ThisWorkbook2
   Set ws = ActiveSheet

   Set wbNew = Workbooks.Add

   With wbNew
       Set wsNew = wbNew.Sheets("�u�@��1")
       wbNewName = ws.ListObjects.Item(tableName)
       ws.ListObjects(1).Range.Copy
       wsNew.Range("A1").PasteSpecial Paste:=xlPasteAll
       .SaveAs FileName:=WB.path & "\" & FileName & ".csv", _
             FileFormat:=xlCSVMSDOS, CreateBackup:=False
   End With
    wbNew.Close savechanges:=False
End Sub


