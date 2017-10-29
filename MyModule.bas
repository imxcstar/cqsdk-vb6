Attribute VB_Name = "MyModule"
Option Explicit
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (lpString1 As Any, lpString2 As Any) As Long

Public Function BytesToBstr(bytes)
    On Error GoTo CuoWu
    Dim SFCW As Boolean
    Dim Unicode As String
    If IsUTF8(bytes) Then
        Unicode = "UTF-8"
    Else
        Unicode = "GB2312"
    End If
TG:
    Dim objstream As Object
    Set objstream = CreateObject("ADODB.Stream")
    With objstream
        .Type = 1
        .Mode = 3
        .Open
        If SFCW = False Then .Write bytes
        .position = 0
        .Type = 2
        .Charset = Unicode
        BytesToBstr = .ReadText
        .Close
    End With
    Exit Function
CuoWu:
    Unicode = "GB2312"
    SFCW = True
    GoTo TG
End Function

Private Function IsUTF8(bytes) As Boolean
    On Error GoTo CuoWu
    Dim i As Long, AscN As Long, Length As Long
    Length = UBound(bytes) + 1
    
    If Length < 3 Then
        IsUTF8 = False
        Exit Function
    ElseIf bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
        IsUTF8 = True
        Exit Function
    End If
    
    Do While i <= Length - 1
        If bytes(i) < 128 Then
            i = i + 1
            AscN = AscN + 1
        ElseIf (bytes(i) And &HE0) = &HC0 And (bytes(i + 1) And &HC0) = &H80 Then
            i = i + 2
            
        ElseIf i + 2 < Length Then
            If (bytes(i) And &HF0) = &HE0 And (bytes(i + 1) And &HC0) = &H80 And (bytes(i + 2) And &HC0) = &H80 Then
                i = i + 3
            Else
                IsUTF8 = False
                Exit Function
            End If
        Else
            IsUTF8 = False
            Exit Function
        End If
    Loop
    
    If AscN = Length Then
        IsUTF8 = False
    Else
        IsUTF8 = True
    End If
    Exit Function
CuoWu:
    IsUTF8 = False
End Function

Public Function StrTByte(Str As String) As Byte()
    Dim ZC() As Byte
    ZC = StrConv(Str, vbFromUnicode)
    ReDim Preserve ZC(UBound(ZC) + 1)
    ZC(UBound(ZC)) = 0
    StrTByte = ZC
End Function

Public Function pGetStringFromPtr(ByVal lPtr As Long) As String
    Dim Buff() As Byte '声明一个Byte数组
    Dim lPointer As Long '声明一个变量，用于存储指针
    lPointer = lPtr
    ReDim Buff(0 To lstrlen(lPointer) * 2 - 1) As Byte  '分配缓存大小,由于得到的是Unicode，所以乘以2
    lstrcpy Buff(0), ByVal lPointer '复制到缓存Buff中
    pGetStringFromPtr = BytesToBstr(Buff)
End Function
