VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3525
      TabIndex        =   4
      Top             =   2265
      Width           =   465
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   2205
      TabIndex        =   2
      Top             =   390
      Width           =   1875
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1965
   End
   Begin MSWinsockLib.Winsock wServer 
      Index           =   0
      Left            =   795
      Top             =   2220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wListen 
      Left            =   270
      Top             =   2250
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Index | Streaming Clients"
      Height          =   240
      Index           =   1
      Left            =   2205
      TabIndex        =   3
      Top             =   75
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "Index | Request Log"
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    'read the read me :D
    Shell "notepad " & App.path & "\readme.txt", vbNormalFocus
End Sub

Private Sub Form_Load()
    DebugFlag = True 'turns on messages to debug.print
    ReDim user(0)
    wListen.LocalPort = 80
    wListen.Listen
    hdrOK = br("HTTP/1.1 200 OK\nContent-Type: text/html\nConnection: Close\n\n")
    
    ie = "C:\Program Files\Internet Explorer\IEXPLORE.EXE"
    If FileExists(ie) Then
        Shell ie & " http://localhost/", vbNormalFocus
    End If
End Sub

Private Sub wListen_ConnectionRequest(ByVal requestID As Long)
    X = -1
    For i = 1 To wServer.UBound
        If wServer(i).State <> sckConnected And _
           wServer(i).State <> sckConnecting And _
           wServer(i).State <> sckConnectionPending Then
           '------
           X = i
           Exit For
        End If
    Next

    If X < 1 Then X = wServer.UBound + 1: Load wServer(X)
    
    wServer(X).Close
    wServer(X).accept requestID
End Sub

Private Sub wServer_DataArrival(index As Integer, ByVal bytesTotal As Long)
    Dim s As String
    Dim h As HTTPRequest
    
    wServer(index).GetData s, vbString
    
    h = ParseRequest(s, wServer(index).RemoteHostIP)
    'Call DebugHttpHeader(s)
    
    If h.page = Empty Then h.page = "login.html"
    db "Request for " & LCase(h.page) & " Assigned index " & index
    List1.AddItem index & " - " & h.ip & " " & h.page
    
    Select Case LCase(h.page)
        Case "login.html"
            HTTP.ServeFile wServer(index), App.path & "\login.html"
            WaitForSentAndClosed index
        Case "frames.html"
            login = LoginUser(h)
            db "prelogin was " & login & " (-1 means failure)"
            If login <> -1 Then
                HTTP.ServeFrames wServer(index), login
            Else
                HTTP.ServeFile wServer(index), App.path & "\sorry.html"
            End If
            WaitForSentAndClosed index
        Case "banner.html"
            db "Serving up banner!"
            HTTP.ServeBanner2 wServer(index), h
            WaitForSentAndClosed index
            If Len(h.qryStr) > 0 Then Call PostChat(h)
        Case "body.html"
           i = GetUserIndex(h)
           db "Users personal index set to " & i & " (only > 0 is valid)"
           If i > 0 Then
                user(i).index = index
                HTTP.InitalizeBody wServer(index)
                WaitForSendComplete index
                List2.AddItem index & " - " & h.ip & " " & h.arg(0)
           Else
                HTTP.Redirect_ wServer(index), "\login.html"
                WaitForSentAndClosed index
                db "Redirected to login because on invalid user id"
           End If
        Case Else:
            db "Oops couldnt find page " & h.page & " for " & h.ip
            wServer(index).SendData hdrOK & "<html><h1>Opps cant find your page!"
            WaitForSentAndClosed index
    End Select
End Sub

Private Function ParseRequest(X, ip) As HTTPRequest
    Dim h As HTTPRequest
    s.Strng = Trim(X)
    
    h.ip = ip
    h.method = s.SubstringToChar(1, " ")
    
    fsp = s.IndexOf(" ") + 2         'first space
    ssp = s.NextIndexOf              'second space
    h.page = s.Substring(fsp, ssp)   'page request
    
    s.Strng = h.page
    qs = s.IndexOf("?")
       
    If qs > 0 Then
        h.page = s.SubstringToChar(1, "?")
        h.qryStr = s.ToEndOfStr(qs + 1)
        h.arg() = Split(h.qryStr, "&")
    End If
       
    'Debug.Print "method=" & h.method & vbCrLf & _
    '            "page=" & h.page & vbCrLf & _
    '            "args=" & Join(h.arg, ",") & vbCrLf & _
    '            "userA=" & h.uAgent
    
    ParseRequest = h
End Function

Private Function LoginUser(t As HTTPRequest) As Integer
   'only set user(i).index with streaming body.html
   'passes back userindex if successful.. -1 if failed
   'test is based on parsed name so html differences dont matter
   
    fuser = ary.StrFindValFromKey(t.arg, "USER")
    says = ary.StrFindValFromKey(t.arg, "SAYS")
    pName = html.ParseAll(fuser)
    
    For i = 1 To UBound(user)
        If user(i).pName = pName Then LoginUser = -1: Exit Function
    Next
    
    ReDim Preserve user(UBound(user) + 1)
        
    ub = UBound(user)
    With user(ub)
        .ip = t.ip
        .fName = fuser
        .says = says
        .pName = pName
    End With
    
    LoginUser = ub
End Function

Private Sub PostChat(h As HTTPRequest)
    whoto = ary.StrFindValFromKey(h.arg, "WHOTO")
    acton = ary.StrFindValFromKey(h.arg, "ACTION")
    says = ary.StrFindValFromKey(h.arg, "SAYS") & "<br>"
    fuser = ary.StrFindValFromKey(h.arg, "USER")
    time_ = Mid(Time, 1, 8)
    
    n = "<br></a>(" & time_ & ") " & fuser & " "
    n = n & acton & " " & whoto & " : " & says & vbCrLf
    
    'If InStr(acton, "whisper") >= 0 Then 'private post
        'find user index from whoto
        'post to talking user and post to user(i).index
    'else
        n = HTTP.Escape(n)
        db "Chat to post: " & n
        For i = 1 To UBound(user)
            X = user(i).index
            db "testing user " & i & " for valid chat frame with value'" & X & "'"
            If Len(X) > 0 And X > 0 Then
                If wServer(X).State = sckConnected Then
                    wServer(X).SendData n & vbCrLf
                    WaitForSendComplete X
                    db "user " & i & " is valid with index " & X
                Else
                    With user(i)
                        .ip = Empty: .fName = Empty: .pName = Empty: .index = Empty
                    End With
                    wServer(X).Close
                End If
            End If
        Next
    'end if
End Sub




Private Sub wServer_SendComplete(index As Integer)
    db "Index " & index & " Send Complete"
    If Not ReadyToClose Then
        wServer(index).Close
        ReadyToClose = True
        db "Index " & index & " has been closed"
    ElseIf Not ReadyToReturn Then
        ReadyToReturn = True
        db "Index " & index & " ready to return : D"
    End If
End Sub

Function WaitForSentAndClosed(index)
    ReadyToClose = False
    db "Index " & index & " is awaiting close verification"
    While Not ReadyToClose
        DoEvents: DoEvents: DoEvents: DoEvents
    Wend
    db "Index " & index & " Has been verified as closed"
End Function

Function WaitForSendComplete(index)
    ReadyToClose = True
    ReadyToReturn = False
    db "Execution paused until Index " & index & " returns"
    While Not ReadyToReturn
        DoEvents: DoEvents: DoEvents: DoEvents
    Wend
    db "Index " & index & " has returned...execution may proceede"
End Function















'Private Sub wServer_SendComplete(index As Integer)
'    db "Index " & index & " Send Complete"
'    'should probably switch these two arrays over to comma
'    'delimited strings and do a string search, because i cant
'    'skinny these arrays cause it causes weird errors with
'    'this event firing at the remote browsers whim
'
'    If Not AryIsEmpty(CloseAfterSend) Then
'        For i = 0 To UBound(CloseAfterSend)
'            If CloseAfterSend(i) = index Then
'                wServer(index).Close
'                CloseAfterSend(i) = ""
'                db "Index " & index & " Has been closed"
'            End If
'        Next
'     End If
'
'     If Not AryIsEmpty(WaitUntilSent) Then
'        For i = 0 To UBound(WaitUntilSent)
'            If WaitUntilSent(i) = index Then
'                WaitUntilSent(i) = ""
'                db "Index " & index & " Has been completed send..ready to resume"
'            End If
'        Next
'     End If
'End Sub'
'
'Function WaitForSentAndClosed(index)
'    push CloseAfterSend(), index
'    db "Index " & index & " is awaiting verification"
'    i = UBound(CloseAfterSend)
'    While CloseAfterSend(i) <> ""
'        DoEvents: DoEvents: DoEvents: DoEvents
'    Wend
'    db "Index " & index & " Has been verified as closed"
'End Function
'
'Function WaitForSendComplete(index)
'    push WaitUntilSent(), index
'    db "Execution paused until Index " & index & " returns"
'    i = UBound(WaitUntilSent)
'    While WaitUntilSent(i) <> ""
'        DoEvents: DoEvents: DoEvents: DoEvents
'    Wend
'    db "Index " & index & " has returned...execution may proceede"
'End Function

