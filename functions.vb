Function GetDuration(start As String, dest As String, key As String)
    Dim firstVal As String, secondVal As String, lastVal As String, dummyURL As String, dummyJson As String
    
    ' Define Google default strings.
    firstVal = "https://maps.googleapis.com/maps/api/distancematrix/json?origins="
    secondVal = "&destinations="
    lastVal = "&mode=driving&language=en&key="
    
    ' Concatonate Google default strings with user parameters.
    Url = firstVal & Replace(start, " ", "+") & secondVal & Replace(dest, " ", "+") & lastVal & key

    ' Get json values from url API as an object.
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", Url, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    
    ' Get the character position of the duration.
    charPos = InStr(objHTTP.responseText, """duration"" : {")
    If charPos = 0 Then GoTo ErrorHandl
    
    ' Check JSON element status.
    Set regEx = CreateObject("VBScript.RegExp"): regEx.Pattern = "status"" : ""([^0-9""]*)""": regEx.Global = False
    Set matches = regEx.Execute(objHTTP.responseText)
    If matches(0).SubMatches(0) <> "OK" Then GoTo ErrorHandl

    ' Replace the string numbers after "duration" with their corresponding numeric values.
    Set regex = CreateObject("VBScript.RegExp"): regex.Pattern = "duration(?:.|\n)*?""value"".*?([0-9]+)": regex.Global = False
    Set matches = regex.Execute(objHTTP.responseText)
    tmpVal = Replace(matches(0).SubMatches(0), ".", Application.International(xlListSeparator))
    
    ' Return value.
    GetDuration = CDbl(tmpVal)/60
    Exit Function
ErrorHandl:
    GetDuration = ""
End Function