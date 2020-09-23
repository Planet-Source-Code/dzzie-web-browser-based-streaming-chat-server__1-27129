Attribute VB_Name = "Globals"
Public Type HTTPRequest
     ip As String
     page As String
     method As String
     arg() As String   'all querystring args in key=value format
     qryStr As String
End Type

Public Type user
    ip As String
    index As Integer    'which ws(index) the client is on
    fName As String     'full user name
    pName As String     'parsed username
    says As String      'entrace message
End Type

Global user() As user

'Global CloseAfterSend()
'Global WaitUntilSent()

Global DebugFlag As Boolean
Global ReadyToClose As Boolean
Global ReadyToReturn As Boolean

'Global history(5) As String

Global s As New Strings
Global hdrOK As String

