Attribute VB_Name = "ary"
'Info:     These are macros for dealing with arrays more efficently
'               note where ever there is an argument named ary it must receive
'               an array that has been previosly dimmed as "dim it()"
'               if it is dimmed as string or variant it will fail :0-_
'
'License:  you are free to use this library in your personal projects, so
'               long as this header remains inplace. This code cannot be
'               used in any project that is to be sold. This source code
'               can be freely distributed so long as this header reamins
'               intact.
'
'Author:   dzzie@yahoo.com
'Sight:    http://www.geocities.com/dzzie

Function AryJoin(ary1, ary2)
    On Error GoTo failed
    X = UBound(ary1)

    For i = 0 To UBound(ary2)
        X = X + 1
        ReDim Preserve ary1(X)
        ary1(X) = ary2(i)
    Next
    AryJoin = ary1
    Exit Function
failed:     ReDim ary1(0): AryJoin = ary1
End Function

Function aryize(it As String) As String()
    Dim r() As String
    For i = 1 To Len(it)
        ReDim Preserve r(i)
        r(i) = Mid(it, i, 1)
    Next
    aryize = r()
End Function


Sub skinny(ary, Optional base As Long = 0) 'remove empty elements
    Dim ret() 'return adjustable base array
    c = base
    For i = base To UBound(ary)
        If ary(i) <> "" Then
            ReDim Preserve ret(c)
            ret(c) = ary(i)
            c = c + 1
        End If
    Next
    ary = ret() 'parent ary obj modified and passed back
End Sub

Sub pop(ary, Optional count = 1) 'this modifies parent ary obj
    If count > UBound(ary) Then ReDim ary(0)
    For i = 1 To count
        ReDim Preserve ary(UBound(ary) - 1)
    Next
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Public Function chngArrayBase(t, base)
    Dim ret()
    lbT = LBound(t)
    elem = UBound(t) - LBound(t) + 1
    ReDim ret(base To elem)
    For i = 0 To elem - 1
        ret(base + i) = t(lbT + i)
    Next
    chngArrayBase = ret()
End Function

Public Function slice(ary, lbnd, ubnd)
    If lbnd > ubnd Then slice = "ERROR": Exit Function
    Dim tmp()
    ReDim tmp(ubnd - lbnd)
    For i = 0 To UBound(tmp)
        tmp(i) = ary(lbnd + i)
    Next
    slice = tmp
End Function


Public Function Slice2Str(ary, lbnd, ubnd, Optional joinChr As String = ",")
    If lbnd > ubnd Then Slice2Str = "ERROR": Exit Function
    Dim tmp()
    ReDim tmp(ubnd - lbnd)
    For i = 0 To UBound(tmp)
        tmp(i) = ary(lbnd + i)
    Next
    Slice2Str = Join(tmp, joinChr)
End Function

Function AryIndexExists(ary, index) As Boolean
    On Error GoTo oops
    i = ary(index) '<-non existant index will throw Error
    AryIndexExists = True
    Exit Function
oops:     AryIndexExists = False
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

'---------------------------------------------------------------------
'-- Next 3 for using arryas like collections with key=value format  --
'---------------------------------------------------------------------
Function FindValFromKey(key, ary)
    If AryIsEmpty(ary) Then Exit Function
    
    For i = 0 To UBound(ary)
        s = GetKeyFromIndex(i, ary)
        If s = key Then
            FindValFromKey = GetValueFromIndex(i, ary)
            Exit Function
        End If
    Next
    FindValFromKey = Empty
End Function

'same as above but for arrays dimmed as string()
Function StrFindValFromKey(ary() As String, key)
    For i = 0 To UBound(ary)
        pos = InStr(ary(i), "=")
        If pos > 0 Then
            k = Mid(ary(i), 1, pos - 1)
            v = Mid(ary(i), pos + 1, Len(ary(i)))
            If k = key Then
                StrFindValFromKey = v
                Exit Function
            End If
        End If
    Next
    StrFindValFromKey = Empty
End Function

Function GetKeyFromIndex(index, ary)
    If Not AryIndexExists(ary, index) Then
        s = ary(index)
        GetKeyFromIndex = Mid(s, 1, InStr(s, "=") - 1)
    End If
End Function

Function GetValueFromIndex(index, ary)
    If AryIndexExists(ary, index) Then
        s = ary(index)
        GetValueFromIndex = Mid(s, InStr(s, "=") + 1, Len(s))
    End If
End Function
