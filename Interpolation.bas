Attribute VB_Name = "Interpolation"
Function LinearInterpolate(TimeSeries As Variant, DataSeries As Variant, targetTime As Variant) As Double
    Dim TimeLower As range
    Dim DataLower As range
    Dim TimeUpper As range
    Dim DataUpper As range

    initIndex = 1
    On Error Resume Next
    initIndex = Application.WorksheetFunction.Match(targetTime, TimeSeries)


    
    For j = initIndex - 1 To 1 Step -1
        If DataSeries.Cells(j).Value2 <> vbNullString Then
            Set TimeLower = TimeSeries.Cells(j)
            Set DataLower = DataSeries.Cells(j)
            Exit For
        End If
    Next j
    
    For k = initIndex To TimeSeries.Cells.Count
        If DataSeries.Cells(k).Value2 <> vbNullString Then
            Set TimeUpper = TimeSeries.Cells(k)
            Set DataUpper = DataSeries.Cells(k)
            Exit For
        End If
    Next k
    
    If Not (TimeLower Is Nothing) And Not (TimeUpper Is Nothing) Then
        x0 = TimeLower.Value2
        y0 = DataLower.Value2
        X1 = TimeUpper.Value2
        y1 = DataUpper.Value2
        x = targetTime
        LinearInterpolate = y0 + (x - x0) * ((y1 - y0) / (X1 - x0))
    ElseIf Not (TimeLower Is Nothing) Then
        LinearInterpolate = DataLower.Value2
    ElseIf Not (TimeUpper Is Nothing) Then
        LinearInterpolate = DataUpper.Value2
    Else
        LinearInterpolate = 0
    End If
    
End Function
Sub testInter()
    MsgBox LinearInterpolate(range("表格82[a]"), range("表格82[b]"), 2)
End Sub
'
'Function LinearInterpolate2(TimeSeries As Variant, DataSeries As Variant, targetTime As Double) As Double
'
'    Dim MAX_smaller As Double
'    Dim MIN_bigger As Double
'    Dim t_MAX_smaller As Double
'    Dim t_MIN_bigger As Double
'    Dim resultData As Double
'
'    Dim TimeSeries_arr As Variant
'    TimeSeries_arr = TimeSeries.Value
'    Dim DataSeries_arr As Variant
'    DataSeries_arr = DataSeries.Value
'
'    t_MAX_smaller = TimeSeries_arr(LBound(TimeSeries_arr, 1), 1)
'    t_MIN_bigger = TimeSeries_arr(UBound(TimeSeries_arr, 1), 1)
'    'MAX_smaller = DataSeries_arr(LBound(TimeSeries_arr, 1), 1)
'    'MIN_bigger = DataSeries_arr(UBound(TimeSeries_arr, 1), 1)
'    MAX_smaller = -1
'    MIN_bigger = -1
'
'
'
'
'        For i = 1 To UBound(TimeSeries_arr, 1)
'            'Debug.Print DataSeries_arr(i, 1)
'
'            If DataSeries(i, 1) <> "" Then
'
'                If TimeSeries_arr(i, 1) = targetTime Then
'                    resultData = DataSeries_arr(i, 1)
'                    LinearInterpolate = resultData
'                    Exit Function
'                End If
'                If TimeSeries_arr(i, 1) < targetTime Then
'                    If TimeSeries_arr(i, 1) > t_MAX_smaller Then
'                        t_MAX_smaller = TimeSeries_arr(i, 1)
'                        MAX_smaller = DataSeries(i, 1)
'                    End If
'                End If
'                If TimeSeries_arr(i, 1) > targetTime Then
'                    If TimeSeries_arr(i, 1) < t_MIN_bigger Then
'                        t_MIN_bigger = TimeSeries_arr(i, 1)
'                        MIN_bigger = DataSeries(i, 1)
'                    End If
'                End If
'            End If
'
'
'
'        Next i
'
'
'        If MAX_smaller > 0 And MIN_bigger > 0 Then
'            'Debug.Print "oooooooooooooo"
'            'Debug.Print t_MAX_smaller
'            'Debug.Print t_MIN_bigger
'            'Debug.Print MAX_smaller
'            'Debug.Print MIN_bigger
'            'Debug.Print targetTime
'            LinearInterpolate = ((Abs(t_MAX_smaller - targetTime) * MIN_bigger + Abs(t_MIN_bigger - targetTime) * MAX_smaller)) / ((Abs(t_MAX_smaller - targetTime) + Abs(t_MIN_bigger - targetTime)))
'            'Debug.Print LinearInterpolate
'        ElseIf MAX_smaller <= 0 And MIN_bigger <= 0 Then
'            LinearInterpolate = 0
'        Else
'            'Debug.Print "oooooooooooooo"
'            'Debug.Print t_MAX_smaller
'            'Debug.Print t_MIN_bigger
'            'Debug.Print MAX_smaller
'            'Debug.Print MIN_bigger
'            LinearInterpolate = MAX_smaller + MIN_bigger + 1
'        End If
'
'End Function

'Function LinearInterpolate(rXv As Variant, rYv As Variant, x As Double) As Double
'    Dim rX As Range
'    Dim rY As Range
'    Set rX = rXv
'    Set rY = rYv
'
'     ' linear interpolator / extrapolator
'     ' R is a two-column range containing known x, known y
'    Dim lR As Long, l1 As Long, l2 As Long
'    Dim nR As Long
'     'If x = 1.5 Then Stop
'
'    nR = rX.Rows.count
'    If nR < 2 Then Exit Function
'
'    If x < rX(1) Then ' x < xmin, extrapolate
'        l1 = 1: l2 = 2: GoTo Interp
'
'    ElseIf x > rX(nR) Then ' x > xmax, extrapolate
'        l1 = nR - 1: l2 = nR: GoTo Interp
'
'    Else
'         ' a binary search would be better here
'        For lR = 1 To nR
'            If rX(lR) = x Then ' x is exact from table
'                LinearInterpolate = rY(lR)
'                Exit Function
'
'            ElseIf rX(lR) > x Then ' x is between tabulated values, interpolate
'                l1 = lR: l2 = lR - 1: GoTo Interp
'
'            End If
'        Next
'    End If
'
'Interp:
'    LinearInterpolate = rY(l1) _
'    + (rY(l2) - rY(l1)) _
'    * (x - rX(l1)) _
'    / (rX(l2) - rX(l1))
'
'End Function
'\  Returns:    yValue      - Interpolated or extrapolated y value
'\                            corresponding to x value
'\  Called by:  User
'\  Calls:
'\  References:
'\  Originated: 01/01/92 - S. J. Wilson (original C code for EPEC.XLA)
'\  Rev Hist:   10/16/96 - (SRS) Generalized function to handle either
'\                         vertical or horizontal references or arrays.
'\                         Changed name from "vTerp" to "Terp" to reflect
'\                         new capabilities. Removed optional argument
'\                         specifying whether references or arrays are being
'\                         passed (this is now done internally).
'\              06/07/96 - (SRS) Added optional argument specifying whether
'\                         references to columns or a VBA arrays are being
'\                         passed for X and Y column arguments. Also turned
'\                         off error handling in vTerp to allow vTerp to run
'\                         without error messages if input is temporarily bad
'\              05/10/96 - (SRS) Modified to allow 1-D array input for xValues
'\                         and yValues. Previously only 2-D input was allowed
'\                         because a worksheet selection is 2-D by default
'\                         (rows & columns) even if only one column is
'\                         selected.
'\              04/09/96 - (SRS) Added lower case conversion of type arguments
'\                         Added error message for invalid type options
'\              03/28/96 - (SRS) Recoded vTerp in Visual Basic for EPECVB.XLA
'\==========================================================================
Function Terp(xValue As Double, _
              xValues As Variant, _
              yValues As Variant, _
     Optional xType As Variant, _
     Optional yType As Variant) As Variant
    Const xDescending As Integer = -1, _
          xAscending As Integer = 1
    Const xBeforeFirst As Integer = -1, _
          xAfterLast As Integer = 1, _
          xWithin As Integer = 0
    Dim numType As String
    Dim exactMatch As Variant
    Dim firstX As Double, _
        secondX As Double, _
        lastX As Double, _
        x_1 As Double, _
        x_2 As Double, _
        y_1 As Double, _
        y_2 As Double
    Dim sortOrder As Integer, _
        rowBefore As Integer, _
        rowAfter As Integer, _
        numElements As Integer
    On Error Resume Next    '\Don't abort function on an error,
                            '\just skip to next line of code
                            '\and resume execution
    '\Default to linear interpolation
    If (IsMissing(xType)) Then xType = "lin"
    If (IsMissing(yType)) Then yType = "lin"
    xType = LCase(xType) '\Convert to all lower case
    yType = LCase(yType) '\for comparison test
    '\Check to make sure an invalid interpolation type was not entered
    If ((xType <> "lin") And (xType <> "log")) Or _
       ((yType <> "lin") And (yType <> "log")) Then
        Terp = "Invalid option!"
        Exit Function
    End If
    numType = TypeName(xValues)
    If (numType = "Range") Then '\passed a worksheet range
        numElements = Application.Max(xValues.Rows.Count, xValues.Columns.Count)
    ElseIf (Not IsError(Application.Search("()", numType))) Then  '\passed an array
       If (Not IsNumeric(UBound(xValues, 2))) Then '\passed a 1-D row array
            '\Error occurs (subscript out of range in UBound - asking for 2D of 1D array)
            '\Execution resumes at next line of code
            '\(skips to "numElements = UBound(xValues)")
        Else  '\2-D array
            xValues = Application.Transpose(xValues)
            yValues = Application.Transpose(yValues)
        End If
        numElements = UBound(xValues)
    Else
        Terp = "Invalid data format!"
        Exit Function
    End If
    firstX = xValues(1)
    secondX = xValues(2)
    lastX = xValues(numElements)
    '\Determine if x values are increasing or decreasing
    sortOrder = xAscending '\default is increasing x values
    If (secondX < firstX) Then sortOrder = xDescending
    '\Determine if x value is before, after, or within the table x value range
    '\and set interpolation/extrapolation bounds accordingly
    If ((sortOrder = xAscending And (xValue > lastX)) _
      Or (sortOrder = xDescending And (xValue < lastX))) Then   '\x value is after last table value
        rowAfter = numElements                                  '\must extrapolate
        rowBefore = rowAfter - 1
    ElseIf ((sortOrder = xAscending And (xValue < firstX)) _
      Or (sortOrder = xDescending And (xValue > firstX))) Then  '\x value is before first table value
        rowAfter = 2                                            '\must extrapolate
        rowBefore = 1
    Else  '\x value is within table values so check first for an exact match
        exactMatch = Application.Match(xValue, xValues, 0)
        If (Not (Application.IsError(exactMatch))) Then
            Terp = yValues(exactMatch) '\Return exact match
            Exit Function
        Else  '\Not exact match so must interpolate
            If (numType = "Range") Then
                rowBefore = Application.Match(xValue, xValues, sortOrder)
            Else
                rowBefore = Application.Match(xValue, xValues, sortOrder) - 1
            End If
            rowAfter = rowBefore + 1
        End If
    End If
    '\Get bounding x and y values
    x_1 = xValues(rowBefore)
    x_2 = xValues(rowAfter)
    y_1 = yValues(rowBefore)
    y_2 = yValues(rowAfter)
    '\Handle logarithmic interpolation options
    If (xType = "log") Then
        xValue = Application.Log10(xValue)
        x_1 = Application.Log10(x_1)
        x_2 = Application.Log10(x_2)
    End If
    If (yType = "log") Then
        y_1 = Application.Log10(y_1)
        y_2 = Application.Log10(y_2)
    End If
    '\Interpolate/extrapolate
    Terp = y_1 + (y_2 - y_1) * (xValue - x_1) / (x_2 - x_1)
    '\Convert logarithmic y value back to table format
    If (yType = "log") Then Terp = 10 ^ Terp
End Function

