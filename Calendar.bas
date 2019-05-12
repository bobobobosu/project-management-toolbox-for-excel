Attribute VB_Name = "Calendar"
Sub PublishCalendar()
    startIndex = Application.WorksheetFunction.index(range("表格2[編號]"), Application.Match(Now() - 60, range("表格2[Start Date]")))
    reault = generateCalendar(startIndex, Evaluate("表格2[#Headers]"), range("表格2"))
End Sub

Public Sub CreateCalendar()
Dim CalendarData As range
Dim N_rows As Integer


' Add a reference to "Microsoft Scripting Runtime"
Dim CurrentFileSystemObject As New FileSystemObject
Dim CurrentTextFile As TextStream

Dim PathString As String
Dim filename As String
Dim FullPath As String

'Columns
range("表格2").Cells(1).Select
Dim WBS As Integer
WBS = range(Evaluate("Cell(""address"",表格2[[#This Row], [編號]:[編號]])")).Column
Dim location As Integer
location = range(Evaluate("Cell(""address"",表格2[[#This Row], [Location]:[Location]])")).Column
'Dim Latitude As Integer
'Latitude = Range(Evaluate("Cell(""address"",表格2[[#This Row], [Latitude]:[Latitude]])")).Column
'Dim Longitude As Integer
'Longitude = Range(Evaluate("Cell(""address"",表格2[[#This Row], [Longitude]:[Longitude]])")).Column
Dim Percent As Integer
Percent = range(Evaluate("Cell(""address"",表格2[[#This Row], [預計百分比]:[預計百分比]])")).Column
Dim StartDate As Integer
StartDate = range(Evaluate("Cell(""address"",表格2[[#This Row], [Start Date]:[Start Date]])")).Column
Dim EndDate As Integer
EndDate = range(Evaluate("Cell(""address"",表格2[[#This Row], [End Date]:[End Date]])")).Column
Dim subject As Integer
subject = range(Evaluate("Cell(""address"",表格2[[#This Row], [Subject]:[Subject]])")).Column
Dim TimeZone As Integer
TimeZone = range(Evaluate("Cell(""address"",表格2[[#This Row], [時區]:[時區]])")).Column


Dim SummaryString As String
Dim LocationString As String
Dim DateStartString As String
Dim DateEndString As String
Dim CurrentDate As Date
Dim CurrentDateEnd As Date
Dim DtstampString As String
Dim UidString As String



    
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2 'Specify stream type - we want To save text/string data.
    fsT.Charset = "utf-8" 'Specify charset For the source text data.
    fsT.Open 'Open the stream And write binary data To the object
    
    
    
    ' It is a assumed the data is arranged in 3 adjacent columns
    ' and uses a named rage "CalendarData"
    '  Column 1: The date of the event
    '  Column 2: The event title string
    '  Column 3: The event location string
    
    Set CalendarData = range("CalendarData")
    N_rows = CalendarData.Rows.Count
    
    filename = "calendar.ics"
    PathString = ActiveWorkbookpath
    PathString = calendarSharePath
    FullPath = PathString + filename
    
    BeginString = "BEGIN:VEVENT"
    EndString = "END:VEVENT"
    
    ' Create the file
    Set CurrentTextFile = CurrentFileSystemObject.CreateTextFile(FullPath)

    
    ' Write header information
    DataString = "BEGIN:VCALENDAR"
    CurrentTextFile.WriteLine (DataString)
    fsT.WriteText DataString
    fsT.WriteText vbCrLf
    
    DataString = "VERSION:2.0"
    CurrentTextFile.WriteLine (DataString)
    fsT.WriteText DataString
    fsT.WriteText vbCrLf

    DataString = "PRODID:-//hacksw/handcal//NONSGML v1.0//EN"
    CurrentTextFile.WriteLine (DataString)
    fsT.WriteText DataString
    fsT.WriteText vbCrLf
    
    ' Write the indivN_rowsidual events
    
    currentPos = getStructuredByIDR(range("交易!S2").Value2, "表格2", "ID").Row - range("表格2").Row
    'For i = N_rows - 549 To N_rows 'N_rows
    For i = currentPos - 100 To N_rows 'N_rows

        CurrentTextFile.WriteLine (BeginString)
        fsT.WriteText BeginString
        fsT.WriteText vbCrLf
        
        
        CurrentDate = CalendarData(i, StartDate) - CalendarData(i, TimeZone) / 24
        CurrentDateEnd = CalendarData(i, EndDate) - CalendarData(i, TimeZone) / 24

        SummaryString = "SUMMARY:" + CStr(CalendarData(i, subject))
        LocationString = "LOCATION:" + CStr(CalendarData(i, location))   'CStr(CalendarData(i, Latitude)) + "," + CStr(CalendarData(i, Longitude))
        
        Dim DescriptionString As String
        DescriptionString = "DESCRIPTION:" ' + "Percentage " + CStr(CalendarData(i, Percent) * 100) + "%"
        
        For j = 1 To 64
            DescriptionString = DescriptionString + CStr(CalendarData(0, j).text) + ": " + "\n" + "     " + CStr(CalendarData(i, j).text) + "\n"
        Next j
        
        
        
        'Application.StatusBar = SummaryString
        
        
        
'        DtstampString = "DTSTAMP;VALUE=DATE:" + Format(CurrentDate, "yyyymmddThhmmss")
'        DateStartString = "DTSTART;VALUE=DATE:" + Format(CurrentDate, "yyyymmddThhmmss")
'        DateEndString = "DTEND;VALUE=DATE:" + Format(CurrentDateEnd, "yyyymmddThhmmss")
        
        DtstampString = "DTSTAMP:" + Format(CurrentDate, "yyyymmddThhmmss") + "Z"
        DateStartString = "DTSTART:" + Format(CurrentDate, "yyyymmddThhmmss") + "Z"
        DateEndString = "DTEND:" + Format(CurrentDateEnd, "yyyymmddThhmmss") + "Z"
        UidString = "UID:" + CStr(CalendarData(i, subject)) + Format(CurrentDate, "yyyymmddhhmmss") + Format(CurrentDateEnd, "yyyymmddhhmmss")
        
        'CurrentTextFile.WriteLine (DtstampString)
        fsT.WriteText DescriptionString
        fsT.WriteText vbCrLf
        
        'CurrentTextFile.WriteLine (DtstampString)
        fsT.WriteText DtstampString
        fsT.WriteText vbCrLf
        
        'CurrentTextFile.WriteLine (UidString)
        fsT.WriteText UidString
        fsT.WriteText vbCrLf
        
        'CurrentTextFile.WriteLine (SummaryString)
        fsT.WriteText SummaryString
        fsT.WriteText vbCrLf
        
        'CurrentTextFile.WriteLine (LocationString)
        fsT.WriteText LocationString
        fsT.WriteText vbCrLf
        
        'CurrentTextFile.WriteLine (DateStartString)
        fsT.WriteText DateStartString
        fsT.WriteText vbCrLf
        
        'CurrentTextFile.WriteLine (DateEndString)
        fsT.WriteText DateEndString
        fsT.WriteText vbCrLf
        
        'CurrentTextFile.WriteLine (EndString)
        fsT.WriteText EndString
        fsT.WriteText vbCrLf
    Next i
    
    ' Write the closing information and close the file
    DataString = "END:VCALENDAR"
    CurrentTextFile.WriteLine (DataString)
    fsT.WriteText DataString
    fsT.WriteText vbCrLf
    
    CurrentTextFile.Close
    
    fsT.SaveToFile FullPath, 2
    

    'Call FtpSend
    
End Sub

Public Sub FtpSend()

Dim vPath As String
Dim vFile As String
Dim vFTPServ As String
Dim fNum As Long

vPath = ftpTmpPath
vFile = calendaricsFileName


'Mounting file command for ftp.exe
fNum = FreeFile()
Open vPath & "\FtpComm.txt" For Output As #fNum
Print #1, "user Public2 P@ssw0rd" ' your login and password"
'Print #1, "cd TargetDir"  'change to dir on server
Print #1, "bin" ' bin or ascii file type to send
Print #1, "put " & vPath & "\" & vFile & " " & "\" & "_Public" & "\" & vFile ' upload local filename to server file
Print #1, "close" ' close connection
Print #1, "quit" ' Quit ftp program
Close


vFTPServ = "nickisverygood.dynu.net"
Shell "ftp -n -i -g -s:" & vPath & "\FtpComm.txt " & vFTPServ, vbNormalNoFocus
'vFTPServ = "192.168.0.152"
'Shell "ftp -n -i -g -s:" & vPath & "\FtpComm.txt " & vFTPServ, vbNormalNoFocus

SetAttr vPath & "\FtpComm.txt", vbNormal

Application.Wait Now + #12:00:01 AM#
'Kill vPath & "\FtpComm.txt"

End Sub


Function hash12(s As String)
' create a 12 character hash from string s

Dim L As Integer, l3 As Integer
Dim s1 As String, s2 As String, s3 As String

L = Len(s)
l3 = Int(L / 3)
s1 = Mid(s, 1, l3)      ' first part
s2 = Mid(s, l3 + 1, l3) ' middle part
s3 = Mid(s, 2 * l3 + 1) ' the rest of the string...

hash12 = hash4(s1) + hash4(s2) + hash4(s3)

End Function

Function hash4(txt)
' copied from the example
Dim x As Long
Dim mask, i, j, nC, crc As Integer
Dim C As String

crc = &HFFFF

For nC = 1 To Len(txt)
    j = Asc(Mid(txt, nC)) ' <<<<<<< new line of code - makes all the difference
    ' instead of j = Val("&H" + Mid(txt, nC, 2))
    crc = crc Xor j
    For j = 1 To 8
        mask = 0
        If crc / 2 <> Int(crc / 2) Then mask = &HA001
        crc = Int(crc / 2) And &H7FFF: crc = crc Xor mask
    Next j
Next nC

C = Hex$(crc)

' <<<<< new section: make sure returned string is always 4 characters long >>>>>
' pad to always have length 4:
While Len(C) < 4
  C = "0" & C
Wend

hash4 = C

End Function
