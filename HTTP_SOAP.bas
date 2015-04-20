Attribute VB_Name = "Module3"
Option Explicit

'The Microsoft XML library is required in the reference.

'##################################################
'Set the following Constants.
'##################################################
Private Const YOUR_HOST = "www.hogehoge.com"
Private Const YOUR_IF = "interface"

Private Const YOUR_ID = "id"
Private Const YOUR_PASS = "password"


Public Sub Main()

    'Calling Sequense Example
    Call GetContactByID(1234)
    Call GetIncidentByID(5678)
    Call GetAnswerByID(1)
    Call UpdateContactByID(1234, "test1", "test2")

End Sub


'Get Contact data by ID
Public Sub GetContactByID(lngContactID As Long)

    Dim xmlHttp As New MSXML2.xmlHttp
    Dim strURL  As String
    Dim strEnv  As String
    Dim XMLDOC  As New DOMDocument
    
    
    'URL
    strURL = "https://" & YOUR_HOST & "/cgi-bin/" & YOUR_IF & ".cfg/services/soap"
    
    strEnv = vbNullString
    strEnv = strEnv & "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & ">"
    
    'Header
    strEnv = strEnv & GetSOAPHeader()
    
    'Body
    strEnv = strEnv & "<soapenv:Body>"
    strEnv = strEnv & "<ns7:Get xmlns:ns7=" & Chr(34) & "urn:messages.ws.rightnow.com/v1_2" & Chr(34) & ">"
    
    strEnv = strEnv & "<ns7:RNObjects xmlns:ns4=" & Chr(34) & "urn:objects.ws.rightnow.com/v1_2" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & " xsi:type=" & Chr(34) & "ns4:Contact" & Chr(34) & ">"
    strEnv = strEnv & "<ID xmlns=" & Chr(34) & "urn:base.ws.rightnow.com/v1_2" & Chr(34) & " id=" & Chr(34) & CStr(lngContactID) & Chr(34) & " />"
    strEnv = strEnv & "<ns4:Notes />"
    strEnv = strEnv & "</ns7:RNObjects>"
    strEnv = strEnv & "<ns7:ProcessingOptions>"
    strEnv = strEnv & "<ns7:FetchAllNames>false</ns7:FetchAllNames>"
    strEnv = strEnv & "</ns7:ProcessingOptions>"
    
    strEnv = strEnv & "</ns7:Get>"
    strEnv = strEnv & "</soapenv:Body>"
    
    strEnv = strEnv & "</soapenv:Envelope>"


With xmlHttp

    .Open "POST", strURL, False
    .setRequestHeader "Host", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "soapAction", "GetContactByID" ' per the documentation
    
    .send strEnv
    
    XMLDOC.LoadXML .responseText
    XMLDOC.Save ActiveWorkbook.Path & "\WebQueryResult" & Format(Now, "yyyymmddHHMMSS") & ".xml"
    
    
End With

    MsgBox xmlHttp.statusText


End Sub

'Get Incident data by ID
Public Sub GetIncidentByID(lngIncidentID As Long)

    Dim xmlHttp As New MSXML2.xmlHttp
    Dim strURL  As String
    Dim strEnv  As String
    Dim XMLDOC  As New DOMDocument
    
    
    'URL
    strURL = "https://" & YOUR_HOST & "/cgi-bin/" & YOUR_IF & ".cfg/services/soap"
    
    strEnv = vbNullString
    strEnv = strEnv & "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & ">"
    
    'Header
    strEnv = strEnv & GetSOAPHeader()
    
    'Body
    strEnv = strEnv & "<soapenv:Body>"
    strEnv = strEnv & "<ns7:Get xmlns:ns7=" & Chr(34) & "urn:messages.ws.rightnow.com/v1_2" & Chr(34) & ">"
    
    strEnv = strEnv & "<ns7:RNObjects xmlns:ns4=" & Chr(34) & "urn:objects.ws.rightnow.com/v1_2" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & " xsi:type=" & Chr(34) & "ns4:Incident" & Chr(34) & ">"
    strEnv = strEnv & "<ID xmlns=" & Chr(34) & "urn:base.ws.rightnow.com/v1_2" & Chr(34) & " id=" & Chr(34) & CStr(lngIncidentID) & Chr(34) & " />"
    strEnv = strEnv & "<ns4:Severity />"
    strEnv = strEnv & "</ns7:RNObjects>"
    strEnv = strEnv & "<ns7:ProcessingOptions>"
    strEnv = strEnv & "<ns7:FetchAllNames>false</ns7:FetchAllNames>"
    strEnv = strEnv & "</ns7:ProcessingOptions>"
    
    strEnv = strEnv & "</ns7:Get>"
    strEnv = strEnv & "</soapenv:Body>"
    
    strEnv = strEnv & "</soapenv:Envelope>"


With xmlHttp

    .Open "POST", strURL, False
    .setRequestHeader "Host", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "soapAction", "GetIncidentByID" ' per the documentation
    
    .send strEnv
    
    XMLDOC.LoadXML .responseText
    XMLDOC.Save ActiveWorkbook.Path & "\WebQueryResult" & Format(Now, "yyyymmddHHMMSS") & ".xml"
    
    
End With

    MsgBox xmlHttp.statusText


End Sub


'Get Answer data by ID
Public Sub GetAnswerByID(lngAnswerID As Long)

    Dim xmlHttp As New MSXML2.xmlHttp
    Dim strURL  As String
    Dim strEnv  As String
    Dim XMLDOC  As New DOMDocument
    
    
    'URL
    strURL = "https://" & YOUR_HOST & "/cgi-bin/" & YOUR_IF & ".cfg/services/soap"
    
    strEnv = vbNullString
    strEnv = strEnv & "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & ">"
    
    'Header
    strEnv = strEnv & GetSOAPHeader()
    
    strEnv = strEnv & "<soapenv:Body>"
    strEnv = strEnv & "<ns7:Get xmlns:ns7=" & Chr(34) & "urn:messages.ws.rightnow.com/v1_2" & Chr(34) & ">"
    
    strEnv = strEnv & "<ns7:RNObjects xmlns:ns4=" & Chr(34) & "urn:objects.ws.rightnow.com/v1_2" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & " xsi:type=" & Chr(34) & "ns4:Answer" & Chr(34) & ">"
    strEnv = strEnv & "<ID xmlns=" & Chr(34) & "urn:base.ws.rightnow.com/v1_2" & Chr(34) & " id=" & Chr(34) & CStr(lngAnswerID) & Chr(34) & " />"
    strEnv = strEnv & "<ns4:Products />"
    strEnv = strEnv & "</ns7:RNObjects>"
    strEnv = strEnv & "<ns7:ProcessingOptions>"
    strEnv = strEnv & "<ns7:FetchAllNames>false</ns7:FetchAllNames>"
    strEnv = strEnv & "</ns7:ProcessingOptions>"
    strEnv = strEnv & "</ns7:Get>"
    
    strEnv = strEnv & "</soapenv:Body>"
    strEnv = strEnv & "</soapenv:Envelope>"


With xmlHttp

    .Open "POST", strURL, False
    .setRequestHeader "Host", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "soapAction", "GetAnswerByID" ' per the documentation
    
    .send strEnv
    
    XMLDOC.LoadXML .responseText
    XMLDOC.Save ActiveWorkbook.Path & "\WebQueryResult" & Format(Now, "yyyymmddHHMMSS") & ".xml"
    
    
End With

    MsgBox xmlHttp.statusText


End Sub

Public Sub UpdateContactByID(lngContactID As Long, strFirstName As String, strLastName As String)

    Dim xmlHttp As New MSXML2.xmlHttp
    Dim strURL  As String
    Dim strEnv  As String
    Dim XMLDOC  As New DOMDocument
    
    
    'URL
    strURL = "https://" & YOUR_HOST & "/cgi-bin/" & YOUR_IF & ".cfg/services/soap"
    
    strEnv = vbNullString
    strEnv = strEnv & "<soapenv:Envelope xmlns:soapenv=" & Chr(34) & "http://schemas.xmlsoap.org/soap/envelope/" & Chr(34) & ">"
    
    'Header
    strEnv = strEnv & GetSOAPHeader()
    
    'Body
    strEnv = strEnv & "<soapenv:Body>"
    strEnv = strEnv & "<ns7:Update xmlns:ns7=" & Chr(34) & "urn:messages.ws.rightnow.com/v1_2" & Chr(34) & ">"
    
    strEnv = strEnv & "<ns7:RNObjects xmlns:ns4=" & Chr(34) & "urn:objects.ws.rightnow.com/v1_2" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & " xsi:type=" & Chr(34) & "ns4:Contact" & Chr(34) & ">"
    strEnv = strEnv & "<ID xmlns=" & Chr(34) & "urn:base.ws.rightnow.com/v1_2" & Chr(34) & " id=" & Chr(34) & CStr(lngContactID) & Chr(34) & " />"
    strEnv = strEnv & "<ns4:Name>"
    strEnv = strEnv & "<ns4:First>" & strFirstName & "</ns4:First>"
    strEnv = strEnv & "<ns4:Last>" & strLastName & "</ns4:Last>"
    strEnv = strEnv & "</ns4:Name>"
    strEnv = strEnv & "</ns7:RNObjects>"
    strEnv = strEnv & "<ns7:ProcessingOptions>"
    strEnv = strEnv & "<ns7:SuppressExternalEvents>false</ns7:SuppressExternalEvents>"
    strEnv = strEnv & "<ns7:SuppressRules>false</ns7:SuppressRules>"
    strEnv = strEnv & "</ns7:ProcessingOptions>"
    strEnv = strEnv & "</ns7:Update>"
    
    strEnv = strEnv & "</soapenv:Body>"
    strEnv = strEnv & "</soapenv:Envelope>"


With xmlHttp

    .Open "POST", strURL, False
    .setRequestHeader "Host", ""
    .setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    .setRequestHeader "soapAction", "GetContactByID" ' per the documentation
    
    .send strEnv
    
    XMLDOC.LoadXML .responseText
    XMLDOC.Save ActiveWorkbook.Path & "\WebQueryResult" & Format(Now, "yyyymmddHHMMSS") & ".xml"
    
    
End With

    MsgBox xmlHttp.statusText


End Sub


Private Function GetSOAPHeader() As String

    Dim strTmp As String
    
    strTmp = vbNullString
    
    strTmp = strTmp & "<soapenv:Header>"
    strTmp = strTmp & "<ns7:ClientInfoHeader xmlns:ns7=" & Chr(34) & "urn:messages.ws.rightnow.com/v1_2" & Chr(34) & " soapenv:mustUnderstand=" & Chr(34) & "0" & Chr(34) & "> "
    strTmp = strTmp & "<ns7:AppID>Basic Update</ns7:AppID>"
    strTmp = strTmp & "</ns7:ClientInfoHeader>"
    strTmp = strTmp & "<wsse:Security xmlns:wsse=" & Chr(34) & "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" & Chr(34) & " mustUnderstand=" & Chr(34) & "1" & Chr(34) & ">"
    strTmp = strTmp & "<wsse:UsernameToken>"
    
    strTmp = strTmp & "<wsse:Username>" & YOUR_ID & "</wsse:Username>"
    strTmp = strTmp & "<wsse:Password Type=" & Chr(34) & "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText" & Chr(34) & ">" & YOUR_PASS & "</wsse:Password>"
    strTmp = strTmp & "</wsse:UsernameToken>"
    strTmp = strTmp & "</wsse:Security>"
    strTmp = strTmp & "</soapenv:Header>"
    
    GetSOAPHeader = strTmp

End Function

