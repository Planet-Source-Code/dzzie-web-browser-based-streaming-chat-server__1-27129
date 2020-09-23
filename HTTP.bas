Attribute VB_Name = "HTTP"
Sub ServeFile(ws As Winsock, path)
    'If ws.State <> sckConnected Then Exit Sub
    
    If fso.FileExists(path) Then
        db "Serving up " & path & " file exists"
        f = fso.ReadFile(path)
        
        ws.SendData hdrOK & f
        'DebugHttpHeader hdrOK & f
    Else
        db "Oops " & path & " couldnt be found!"
        h = hdrOK & "<html><h1>Resource Not Found"
        ws.SendData h
    End If
End Sub

Function GetUserIndex(h As HTTPRequest)
    fuser = ary.StrFindValFromKey(h.arg, "USER")
    
    For i = 1 To UBound(user)
        If user(i).fName = fuser Then GetUserIndex = i: Exit Function
    Next
    
    'returns zero if user not found
    GetUserIndex = -1
End Function


Function ContentTypeFromPath(path)
    ext = fso.GetExtension(path) 'works for complete paths or just file names
    Select Case ext
        Case ".jpg": X = " image/jpeg \n"
        Case ".gif": X = " image/gif \n"
        Case Else: X = " text/html \n"
        'Case ".txt", ".htm", ".htm": X = " text/html \n"
    End Select
    ContentTypeFromPath = "Content-type: " & X
End Function

Sub Redirect_(ws As Winsock, url)
    h = br("HTTP/1.1 301 Found\nLocation: ___\n\n")
    h = Replace(h, "___", url)
    
    ws.SendData h
    'DebugHttpHeader h
End Sub


Sub ServeBanner2(ws As Winsock, h As HTTPRequest)
    Dim u As String
    fuser = Escape(ary.StrFindValFromKey(h.arg, "USER"))
    
    f = ReadFile(App.path & "\banner.html")
    f = Replace(f, "%USERNAME%", fuser)
    
    t = vbTab & "<option value=""___"">___"
    For i = 1 To UBound(user)
        If user(i).ip <> "" Then
            u = u & Replace(t, "___", user(i).pName) & vbCrLf
        End If
    Next
    f = Replace(f, "%CHATTERS%", u)

    ws.SendData hdrOK & f
    'DebugHttpHeader hdrOK & f
    db "Banner Serve complete"
End Sub

Sub ServeFrames(ws As Winsock, uIndex)
    f = ReadFile(App.path & "\frames.html")
    f = Replace(f, "%USERNAME%", user(uIndex).fName)
    f = Replace(f, "%SAYS%", user(uIndex).says)
    
    ws.SendData hdrOK & f
    'DebugHttpHeader hdrOK & f
End Sub

Sub InitalizeBody(ws As Winsock)
    k = "HTTP/1.1 200 OK" & vbCrLf _
        & "Keep-Alive: timeout 15, Max = 99" & vbCrLf _
        & "Connection: Keep-Alive" & vbCrLf _
        & "Transfer-Encoding: chunked" & vbCrLf _
        & "Content-Type: text/html" & vbCrLf & vbCrLf & "FFFF" _
        & vbCrLf & "<html><body bgcolor=black text=white>"
        
    ws.SendData k
    ws.SendData String(200, " ") & vbCrLf
    'DebugHttpHeader k
End Sub

Sub DebugHttpHeader(header)
    Dim t
    h = Split(header, vbCrLf)
    For i = 0 To UBound(h)
        If h(i) = Empty Then Exit For
        t = t & h(i) & vbCrLf
    Next
    db vbCrLf & String(25, "-") & "HTTP Header Follows" & vbCrLf
    db t & vbCrLf & String(25, "-") & vbCrLf
End Sub

Function Escape(it)
    Dim f(): Dim c()
    n = Replace(it, "+", " ")
    If InStr(n, "%") > 0 Then
        t = Split(n, "%")
        For i = 0 To UBound(t)
            a = Left(t(i), 2)
            b = IsHex(a)
            If b <> Empty Then
                push f(), "%" & a
                push c(), b
            End If
        Next
        For i = 0 To UBound(f)
            n = Replace(n, f(i), c(i))
        Next
    End If
    Escape = n
End Function

Private Function IsHex(it)
    On Error GoTo out
      IsHex = Chr(Int("&H" & it))
    Exit Function
out:  IsHex = Empty
End Function











'------------------------------------------------------------------
'-- still experimenting with best way to efficiently create http --
'--     headers on the fly and the most intutivly..dont need this--
'--     for this project though.                                 --
'------------------------------------------------------------------
'Private Enum rCode
'    OK = 200
'    Redirect = 301
'    Authorize = 401
'    NotFound = 404
'    Forbidden = 403
'    Gone = 410
'End Enum
'
'Private Enum cType
'    KeepAlive = 1
''    Close_ = 2
'    chunked = 3
'End Enum
'
'Private Enum conType
'    html = 1
'    Jpeg = 2
'    Gif = 3
'End Enum
'
'
'Private Function buildHeader(RespCode As rCode, Connection As cType, contentType As conType, Optional data = Empty) As String
''Connection Types: 'Keep-Alive','Chunked','Close'
'ContentType: 'text/html','image/jpeg','image/gif'
'for redirect you must use ResponseCode of 301, or '302 Found'
'404 Not Found, 403 FORBIDDEN , 410 Gone, 401 Authorization Required
'
'    Dim h() As String
'    Select Case RespCode
'       Case 200: X = "200 OK"
'       Case 301: X = "301 Moved Permanently"
'       Case 401: X = "401 Authorization Required"
'       Case 403: X = "403 Forbidden"
'       Case 404: X = "404 Not Found"
'       Case 410: X = "410 Gone"
'    End Select
'
'    push h(), "HTTP/1.1 " & X
'    push h(), "Server: Apache/1.3.11 (Unix)"
'    push h(), "Pragma: no-cache"
'    push h(), "Accept-Ranges: bytes"
'
'    If FileExists(data) Then X = FileLen(data) _
'    Else X = Len(data)
'    push h(), "Content-Length: " & X
'
'    Select Case Connection
'        Case 1: X = "keep-alive"
'        Case 2: X = "Close"
'        Case 3: X = "Chunked"
'    End Select
'    push h(), "Connection: " & X
'
'    Select Case contentType
'        Case 1: X = "text/html"
'        Case 2: X = "image/jpeg"
'        Case 3: X = "image/gif"
'    End Select
'    push h(), "Content-Type: " & X
'
'    If RespCode = Authorize Then
'         push h(), "WWW-Authenticate: Basic realm=""" & data & """"
'    ElseIf RespCode = Redirect Then
'         push h(), "Location: " & data
'    ElseIf data <> Empty Then
'         push h(), vbCrLf & data
'    End If
'
'    buildHeader = Join(h, vbCrLf) & vbCrLf
'End Function

