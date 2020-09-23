Attribute VB_Name = "html"
'cut scripts out of html,
'removes script blocks and
'breaks all inline scripting
'these are kinda old coding so milage will vary..time for
'light overhaul

'Author: dzzie@yahoo.com

Public Function parseScript(info)
  Dim EndOfScript As Integer, scripData As String, trimPage As String
  info = filt(info, "javascript,vbscript,mocha,createobject,activex")
  Script = Split(info, "<script")
  If UBound(Script) = 0 Then parseScript = info: Exit Function _
  Else: trimPage = Script(0)
  For i = 1 To UBound(Script)
    EndOfScript = InStr(1, Script(i), "</script>")
    trimPage = trimPage & Mid(Script(i), EndOfScript + 10, Len(Script(i)))
  Next
  parseScript = trimPage
End Function


'remove all html tags (can be buggered if html
'tag contains quoted > or <
Public Function parseHtml(info) As String
     Dim temp As String, EndOfTag As Integer
     fmat = Replace(info, "&nbsp;", " ")
     cut = Split(fmat, "<")

   For i = 0 To UBound(cut)  'cut at all html start tags
     EndOfTag = InStr(1, cut(i), ">")
        If EndOfTag > 0 Then
          EndOfText = Len(cut(i))
          NL = False
          If Left(cut(i), 2) = "br" Then NL = True
          cut(i) = Mid(cut(i), EndOfTag + 1, EndOfText)
          If NL Then cut(i) = vbCrLf & cut(i)
          If cut(i) = vbCrLf Then cut(i) = ""
        End If
     temp = temp & cut(i)
    Next
    
    parseHtml = temp
End Function

'trims out &amp; type html for text
Public Function parseAnds(info)
  Dim temp As String
  cut = Split(info, "&")
  If UBound(cut) > 0 Then
    For i = 0 To UBound(cut)            'cut at all start tags (&)
      EndOfTag = InStr(1, cut(i), ";")
        If EndOfTag > 0 Then
           EndOfText = Len(cut(i))
           cut(i) = Mid(cut(i), EndOfTag + 1, EndOfText)
        End If
      temp = temp & cut(i)
    Next
   parseAnds = temp
  Else: parseAnds = info
  End If
End Function

Function ParseAll(it) As String
    't = parseAnds(it)
    't = parseScript(t)
    ParseAll = parseHtml(it)
End Function
