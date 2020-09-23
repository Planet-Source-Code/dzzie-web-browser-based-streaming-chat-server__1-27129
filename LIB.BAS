Attribute VB_Name = "lib"
'Info:     commonly needed tests, shortcuts, niceties etc...
'
'License:  you are free to use this library in your personal projects, so
'             long as this header remains inplace. This code cannot be
'             used in any project that is to be sold. This source code
'             can be freely distributed so long as this header reamins
'             intact.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie

'find = comma delimited string of criteria
Function accept(strTest, find) As Boolean
    tmp = Split(find, ",")
    ret = False
    For i = 0 To UBound(tmp)
      X = InStr(1, strTest, tmp(i))
      If X > 0 Then ret = True
    Next
    accept = ret
End Function

'find = comma delimited string of criteria
Function reject(strTest, find) As Boolean
    tmp = Split(find, ",")
    ret = False
    For i = 0 To UBound(tmp)
      X = InStr(1, strTest, tmp(i))
      If X > 0 Then ret = True
    Next
    reject = ret
End Function

'remove = comma delimited string of stuff to remove
Public Function filt(txt, remove As String)
  If Right(txt, 1) = "," Then txt = Mid(txt, 1, Len(txt) - 1)
  tmp = Split(remove, ",")
  For i = 0 To UBound(tmp)
     txt = Replace(txt, tmp(i), "", , , vbTextCompare)
  Next
  filt = txt
End Function

Private Function shave(it)
    it = Trim(it)
    shave = LTrim(it)
End Function

'returns "max_whole_divisions,remainder"
Public Function calcDivs(maxsz As Long, division As Integer) As String
        sz = maxsz
        tmp = 0
        While sz >= division
           sz = sz - division
           tmp = tmp + 1
        Wend
        calcDivs = tmp & "," & sz
End Function

'allows strings similar to c style
Public Function br(it)
 tmp = Replace(it, "\n", vbCrLf)
 tmp = Replace(tmp, "\t", vbTab)
 br = Replace(tmp, "\q", """")
End Function

'takes hex chars -> dec -> returns empty if failed
Function IsHex(it)
    On Error GoTo out
      IsHex = Chr(Int("&H" & it))
    Exit Function
out:  IsHex = Empty
End Function

'adds it to textbox w/ "\n"
Public Sub tx(it, t As TextBox)
  t = t & it & vbCrLf
End Sub

'keep textbox scrolled to last entry
Public Sub scroll(t As TextBox)
  t.SelStart = Len(t)
End Sub

Public Sub activate(txt As TextBox, Optional enable = True)
  If enable Then
    txt.Enabled = True
    txt.BackColor = vbWhite
  Else
    txt.Enabled = False
    txt.BackColor = &H8000000F 'gray
  End If
End Sub

Public Sub db(it)
  If DebugFlag Then Debug.Print it
End Sub


'spits back data as though looking at in hex editor
'good for debugging non printable data transfers
Function hexdump(it)
    my = ""
    For i = 1 To Len(it)
        c = Hex(Asc(Mid(it, i, 1)))
        c = IIf(Len(c) = 1, "0" & c, c)
        my = my & c & " "
        If i Mod 16 = 0 Then
            s = Mid(it, i - 15, 16)
            my = my & "  [" & s & "]" & vbCrLf
        End If
    Next
    If Len(it) < 16 Then hexdump = my & "  [" & it & "]" & vbCrLf _
    Else hexdump = my
End Function











