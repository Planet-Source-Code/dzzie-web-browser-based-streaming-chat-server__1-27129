VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Info:     Macros for VB's built in string parsing routines..gives you
'               more flexibility with less code...this is loosely modeled
'               on a javascript model + other macros i commonly needed...
'               these should make your code easier to read and less error
'               prone.
'
'License:  you are free to use this library in your personal projects, so
'             long as this header remains inplace. This code cannot be
'             used in any project that is to be sold. This source code
'             can be freely distributed so long as this header reamins
'             intact.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie

'these provide a "memory" of sorts so you dont re-enter data every time
Private sString As String    'set the string to parse once and use often
Private sPointer As Integer  'some f(x) track where you are in string
Private sChar As String      'some f(x) also track last character searched for

Public Property Let Strng(it)
    sString = CStr(it)
    sPointer = 1
End Property

Public Property Get Strng()
    Strng = sString
End Property

Public Property Get Pointer() As Integer
    Pointer = sPointer
End Property

Public Property Let Pointer(strPointer As Integer)
    sPointer = CInt(strPointer)
    If sPointer = 0 Or sPointer > leng Then sPointer = 1
End Property

Public Property Get Length() As Integer
    Length = Len(sString)
End Property

Public Function ReadXChars(StartAt, leng) As String
    If StartAt = 0 Then StartAt = 1
    If StartAt + leng > Len(sString) Then leng = Len(sString) - StartAt
    ReadXChars = Mid(sString, StartAt, leng)
End Function

Public Function Substring(X, Y) As String
    If X = 0 Then X = 1
    If Y > Length Then Y = Length
    Substring = Mid(sString, X, Y - X)
End Function

Public Function ToEndOfStr(from)
    If from = 0 Then from = 1
    If from > Length Then
      ToEndOfStr = -1
    Else
      ToEndOfStr = Mid(sString, from, Length)
    End If
End Function

Public Function IndexOf(it, Optional StartAt = 1)
    X = InStr(StartAt, sString, it, vbTextCompare)
    sChar = it
    sPointer = X + 1
    IndexOf = X
End Function

Public Function NextIndexOf(Optional ChangePtrChrTo = Empty)
    'ChangePtrChrTo effectly means search from pointer to this char
    If ChangePtrChrTo <> Empty Then sChar = ChangePtrChrTo
    If sPointer > Length Then NextIndexOf = -1: Exit Function
    X = InStr(sPointer, sString, sChar, vbTextCompare)
    sPointer = X + 1
    NextIndexOf = X
End Function

Function SubstringToChar(startPos, EndChr) As String
    mark = InStr(startPos + 1, sString, EndChr, vbTextCompare)
    If mark < 0 Then SubstringToChar = Empty: Exit Function
    ret = Mid(sString, startPos, mark - startPos)
    SubstringToChar = CStr(ret)
End Function

Public Function SubstringToNext(Optional advPtr As Boolean = False) As String
    X = InStr(sPointer, sString, sChar, vbTextCompare)
    If X > 0 Then
      SubstringToNext = Mid(sString, sPointer, X - sPointer)
      If advPtr Then sPointer = IIf(X + 1 <= Length, X + 1, Length)
    Else
      SubstringToNext = Mid(sString, sPointer, Length)
      If advPtr Then sPointer = Length
    End If
End Function

Public Function CharAt(X) As String
    If X > Length Then CharAt = "-1": Exit Function
    CharAt = Mid(sString, X, 1)
End Function

Public Function GetChar() As String
    If sPointer = 0 Then sPointer = 1
    GetChar = Mid(sString, sPointer, 1)
    If sPointer < Length Then
       sPointer = sPointer + 1
    Else
       sPointer = -1
    End If
End Function

Public Function CharCodeAt(X, Optional inHex As Boolean = False)
    CharCodeAt = IIf(inHex = True, Hex(Asc(CharAt(X))), Asc(CharAt(X)))
End Function

Public Function ReplacePtrChar(ins)
    before = Mid(sString, 1, sPointer - 2)
    after = Mid(sString, sPointer, Length - sPointer + 1)
    sString = before & ins & after
    sPointer = sPointer + Len(ins)
    ReplacePtrChar = sString
End Function

Public Function ReplaceCharAt(xpos, ins)
  On Error GoTo bad
  before = Mid(sString, 1, xpos - 1)
  after = Mid(sString, sPointer, Length - xpos + 1)
  sString = before & ins & after
  sPointer = xpos + Len(ins)
  ReplaceCharAt = sString
Exit Function
bad: ReplaceCharAt = False
End Function

Function ToEndofStrFromChar(StartChar) As String
    sMark = InStr(1, sString, StartChar, vbTextCompare)
    sMark = sMark + Len(StartChar)
    eMark = Length - sMark + 1
    If sMark < 0 Or eMark < 0 Then GoTo failed
    ret = Mid(sString, sMark, eMark)
    ToEndofStrFromChar = CStr(ret)
 Exit Function
failed: ToEndofStrFromChar = Empty: Exit Function
End Function

Function StringToEndMinus(xChars)
    If Length - xChars < 1 Then StringToEndMinus = Empty: Exit Function
    StringToEndMinus = Mid(sString, 1, Length - xChars)
End Function

Function InstrRevtoChar(xChar)
    pos = InStrRev(sString, CStr(xChar), , vbTextCompare)
    If pos < 0 Then InstrRevtoChar = Empty: Exit Function
    InstrRevtoChar = Mid(sString, pos, Length)
End Function

Function SubstringToLastChar(xChar)
    pos = InStrRev(sString, CStr(xChar), , vbTextCompare)
    If pos < 0 Then ISubstringToLastChar = Empty: Exit Function
    SubstringToLastChar = Mid(sString, 1, pos)
End Function
