Attribute VB_Name = "Module22"
Option Explicit

Function GetCoordinates(Address As String) As String
    Address = InParens(Address)
    '-----------------------------------------------------------------------------------------------------
    'This function returns the latitude and longitude of a given address using the Google Geocoding API.
    'The function uses the "simplest" form of Google Geocoding API (sending only the address parameter),
    'so, optional parameters such as bounds, key, language, region and components are NOT used.
    'In case of multiple results (for example two cities sharing the same name), the function
    'returns the FIRST OCCURRENCE, so be careful in the input address (tip: use the city name and the
    'postal code if they are available).
    
    'NOTE: As Google points out, the use of the Google Geocoding API is subject to a limit of 2500
    'requests per day, so be careful not to exceed this limit.
    'For more info check: https://developers.google.com/maps/documentation/geocoding
    
    'In order to use this function you must enable the XML, v3.0 library from VBA editor:
    'Go to Tools -> References -> check the Microsoft XML, v3.0.
    
    'Written by:    Christos Samaras
    'Date:          12/06/2014
    'e-mail:        xristos.samaras@gmail.com
    'site:          http://www.myengineeringworld.net
    '-----------------------------------------------------------------------------------------------------
    
    'Declaring the necessary variables. Using 30 at the first two variables because it
    'corresponds to the "Microsoft XML, v3.0" library in VBA (msxml3.dll).
    Dim Request         As New XMLHTTP30
    Dim Results         As New DOMDocument30
    Dim StatusNode      As IXMLDOMNode
    Dim LatitudeNode    As IXMLDOMNode
    Dim LongitudeNode   As IXMLDOMNode
            
    On Error GoTo ErrorHandler
    
    'Create the request based on Google Geocoding API. Parameters (from Google page):
    '- Address: The address that you want to geocode.
    '- Sensor: Indicates whether your application used a sensor to determine the user's location.
    'This parameter is no longer required.
    Request.Open "GET", "http://maps.googleapis.com/maps/api/geocode/xml?" _
    & "&address=" & Address & "&sensor=false", False
            
    'Send the request to the Google server.
    Request.send
    
    'Read the results from the request.
    Results.LoadXML Request.responseText
    
    'Get the status node value.
    Set StatusNode = Results.SelectSingleNode("//status")
    
    'Based on the status node result, proceed accordingly.
    Select Case UCase(StatusNode.text)
    
        Case "OK"   'The API request was successful. At least one geocode was returned.
            
            'Get the latitdue and longitude node values of the first geocode.
            Set LatitudeNode = Results.SelectSingleNode("//result/geometry/location/lat")
            Set LongitudeNode = Results.SelectSingleNode("//result/geometry/location/lng")
            
            'Return the coordinates as string (latitude, longitude).
            GetCoordinates = LatitudeNode.text & ", " & LongitudeNode.text
        
        Case "ZERO_RESULTS"   'The geocode was successful but returned no results.
            GetCoordinates = "The address probably not exists"
            
        Case "OVER_QUERY_LIMIT" 'The requestor has exceeded the limit of 2500 request/day.
            GetCoordinates = "Requestor has exceeded the server limit"
            
        Case "REQUEST_DENIED"   'The API did not complete the request.
            GetCoordinates = "Server denied the request"
            
        Case "INVALID_REQUEST"  'The API request is empty or is malformed.
            GetCoordinates = "Request was empty or malformed"
        
        Case "UNKNOWN_ERROR"    'Indicates that the request could not be processed due to a server error.
            GetCoordinates = "Unknown error"
        
        Case Else   'Just in case...
            GetCoordinates = "Error"
        
    End Select
        
    'In case of error, release the objects.
ErrorHandler:
    Set StatusNode = Nothing
    Set LatitudeNode = Nothing
    Set LongitudeNode = Nothing
    Set Results = Nothing
    Set Request = Nothing
    
End Function

'--------------------------------------------------------------------------
'The next two functions using the GetCoordinates function in order to get
'the latitude and the longitude correspondingly of a given address.
'--------------------------------------------------------------------------

Function GetLatidue(Address As String) As Double

    Dim Coordinates As String
    
    'Get the coordinates for the given address.
    Coordinates = GetCoordinates(Address)
    
    'Return the latitude as number (double).
    If Coordinates <> "" Then
        On Error Resume Next
        GetLatidue = CDbl(Left(Coordinates, WorksheetFunction.Find(",", Coordinates) - 1))
    End If

End Function

Function GetLongitude(Address As String) As Double

    Dim Coordinates As String
    
    'Get the coordinates for the given address.
    Coordinates = GetCoordinates(Address)
    
    'Return the longitude as number (double).
    If Coordinates <> "" Then
        On Error Resume Next
        GetLongitude = CDbl(Right(Coordinates, Len(Coordinates) - WorksheetFunction.Find(",", Coordinates)))
    End If
    
End Function

Function InParens(s As String) As String
    Dim i1 As Long, i2 As Long
    
    i1 = InStr(s, "[")
    If i1 Then
        i2 = InStrRev(s, "]")
        If i2 Then
            InParens = Mid(s, i1 + 1, i2 - i1 - 1)
        End If
    End If
End Function


Sub ggggggggg()
    Debug.Print InParens("gggg[werv]ggggrevwg")
End Sub
