Function GetMyPublicIP() As String

    Dim HttpRequest As Object
    
    On Error Resume Next
    'Create the XMLHttpRequest object.
    Set HttpRequest = CreateObject("MSXML2.XMLHTTP")

    'Check if the object was created.
    If Err.Number <> 0 Then
        'Return error message.
        GetMyPublicIP = "Could not create the XMLHttpRequest object!"
        'Release the object and exit.
        Set HttpRequest = Nothing
        Exit Function
    End If
    On Error GoTo 0
    
    'Create the request - no special parameters required.
    HttpRequest.Open "GET", "http://myip.dnsomatic.com", False
    
    'Send the request to the site.
    HttpRequest.Send
        
    'Return the result of the request (the IP string).
    GetMyPublicIP = HttpRequest.ResponseText

End Function

Function GetMyMACAddress() As String

    'Declaring the necessary variables.
    Dim strComputer     As String
    Dim objWMIService   As Object
    Dim colItems        As Object
    Dim objItem         As Object
    Dim myMACAddress    As String
    
    'Set the computer.
    strComputer = "."
    
    'The root\cimv2 namespace is used to access the Win32_NetworkAdapterConfiguration class.
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    
    'A select query is used to get a collection of network adapters that have the property IPEnabled equal to true.
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    'Loop through all the collection of adapters and return the MAC address of the first adapter that has a non-empty IP.
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then myMACAddress = objItem.MACAddress
        Exit For
    Next
    
    'Return the IP string.
    GetMyMACAddress = myMACAddress

End Function

Sub AutoOpen()
'
' AutoOpen Macro
'
'
    Dim IP As String
    Dim MAC As String
    
    IP = GetMyPublicIP()
    MAC = GetMyMACAddress()

End Sub

