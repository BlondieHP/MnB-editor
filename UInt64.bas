Attribute VB_Name = "Uint64"
'*************************************************************************
'**模 块 名：Uint64
'**说    明：版权所有2008 - 2009(C)1
'**创 建 人：mafei82394
'**日    期：2008-05-21 08:12:28
'**修 改 人：
'**日    期：
'**描    述：
'**版    本：V1.0.0
'*************************************************************************
Rem==============================================================================
Rem==============================================================================

Public Type Integer64b
        by(0 To 7) As Byte
End Type
Rem==============================================================================
Rem==============================================================================
Rem 存储各进制的数字字符
Public CHex As Variant
Public CDe As Variant
Public COct As Variant
Public CBin As Variant
Rem==============================================================================
Rem==============================================================================
Public Sub Init_Integer64b() '初始化
    CHex = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F")
    CDe = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
    COct = Array("0", "1", "2", "3", "4", "5", "6", "7")
    CBin = Array("0", "1")
End Sub
Rem==============================================================================
Rem==============================================================================
Public Sub let64b(ByRef a As Integer64b, b As Integer64b) '赋值
    Dim TemP As Long
    For n = 0 To 7
        a.by(n) = b.by(n)
    Next n
End Sub
Rem==============================================================================
Rem==============================================================================
Public Function ChkBit8b(var As Byte, bt As Byte) As Boolean '测某位是不是零
    ChkBit8b = Not ((var And (2 ^ bt)) = 0)
End Function
Rem==============================================================================
Public Sub TurnBit(ByRef var As Integer64b, bt As Byte) '把某64位数的某一位反转
    Dim by As Byte, bi As Byte
    by = bt \ 8
    bi = bt Mod 8
    var.by(by) = var.by(by) Xor (2 ^ bi)
End Sub
Rem==============================================================================
Public Function ChkBit64b(a As Integer64b, b As Byte) As Boolean '测某位是不是零(64版)
    Dim c As Byte, D As Byte
    D = b Mod 8
    c = b \ 8
    ChkBit64b = ChkBit8b(a.by(c), D)
End Function
Rem==============================================================================
Rem==============================================================================
Public Function Plus64b(a As Integer64b, b As Integer64b) As Integer64b '64 位加法
    Dim inc As Boolean
    Dim TemP As Integer
    inc = False
    For n = 0 To 7
        TemP = 0 + a.by(n) + b.by(n)
        If inc Then TemP = TemP + 1
        inc = (TemP > 255)
        Plus64b.by(n) = TemP Mod 256
    Next n
End Function
Rem==============================================================================
Rem==============================================================================
Public Sub Plus64b8b(ByRef a As Integer64b, b As Byte) '64位数加8位数,结果直接改原64位数
    Dim TemP As Integer
    Dim inc As Boolean
    inc = False
    For n = 0 To 7
        If n = 0 Then
            TemP = 0 + a.by(0) + b
        Else
            TemP = a.by(n)
        End If
        If inc Then TemP = TemP + 1
        inc = (TemP > 255)
        a.by(n) = TemP Mod 256
    Next n
End Sub
Rem==============================================================================
Rem==============================================================================
Public Function Bu64b(a As Integer64b) As Integer64b '取反加一(求补运算)
    For n = 0 To 7
        Bu64b.by(n) = 255 - a.by(n)
    Next n
    Call Plus64b8b(Bu64b, 1)
End Function
Rem==============================================================================
Rem==============================================================================
Public Sub LeftMv64b(a As Integer64b) '左移
    Dim inc As Boolean
    Dim TemP As Integer
    inc = False
    For n = 0 To 7
        TemP = a.by(n) * 2
        If inc Then TemP = TemP + 1
        inc = (TemP > 255)
        a.by(n) = TemP Mod 256
    Next n
End Sub
Rem==============================================================================
Rem==============================================================================
Public Sub LeftMv64bEx(a As Integer64b, n As Byte)
Dim i As Byte

For i = 1 To n
    LeftMv64b a
Next i
End Sub
Rem==============================================================================
Rem==============================================================================
Public Sub RightMv64b(a As Integer64b) '右移
    Dim Dec As Boolean
    Dim TemP As Integer
    Dec = False
    For n = 0 To 7
        TemP = a.by(7 - n) \ 2
        If Dec Then TemP = TemP + 128
        Dec = a.by(7 - n) Mod 2
        a.by(7 - n) = TemP
    Next n
End Sub
Rem==============================================================================
Rem==============================================================================
Public Sub RightMv64bEx(a As Integer64b, n As Byte)
Dim i As Byte

For i = 1 To n
    RightMv64b a
Next i
End Sub
Rem==============================================================================
Rem==============================================================================
Public Function Multi64b8b(a As Integer64b, b As Byte) As Integer64b '64位整数用的乘法,只能乘以8b正整数
    Dim tmp As Integer64b
    Dim n As Byte
    Call let64b(tmp, a)
    For n = 0 To 7
        If ChkBit8b(b, n) Then Multi64b8b = Plus64b(Multi64b8b, tmp)
        Call LeftMv64b(tmp)
    Next n
End Function
Rem==============================================================================
Rem==============================================================================
Public Function Multi64b(a As Integer64b, b As Integer64b) As Integer64b '64位整数用的乘法
    Dim tmp As Integer64b
    Dim n As Byte
    Call let64b(tmp, a)
    For n = 0 To 63
        If ChkBit64b(b, n) Then Multi64b = Plus64b(Multi64b, tmp)
        Call LeftMv64b(tmp)
    Next n
End Function
Rem==============================================================================
Rem==============================================================================

Public Function Div64b(a As Integer64b, b As Byte) As Integer64b '64位除以8位数
    Dim TemP As Integer64b, TempBcs As Integer64b, Temp1 As Integer64b
    Call let64b(TempBcs, a) '临时被除数
    Temp1.by(0) = 1 '作为权值
    ws8 = CheckWS8(b)
    ws64 = CheckWS64(a)
    TemP.by(0) = b '给临时变量赋值
    If b = 0 Or ws64 < ws8 Then Exit Function
    For n = 1 To ws64 - ws8
        Call LeftMv64b(TemP)
        Call LeftMv64b(Temp1)
    Next n
    For n = 1 To ws64 - ws8
        If Compare64b(TempBcs, TemP) Then
            TempBcs = Minus64b(TempBcs, TemP)
            Div64b = Plus64b(Div64b, Temp1)
        End If

        Call RightMv64b(TemP)
        Call RightMv64b(Temp1)
    Next n
    If Compare64b(TempBcs, TemP) Then
        TempBcs = Minus64b(TempBcs, TemP)
        Div64b = Plus64b(Div64b, Temp1)
    End If
End Function
Rem==============================================================================
Rem==============================================================================
Public Function Mod64b(a As Integer64b, b As Byte) As Byte '64位数模8位数
    Mod64b = Minus64b(a, Multi64b8b(Div64b(a, b), b)).by(0)
End Function
Rem==============================================================================
Rem==============================================================================
Public Function IsZero64b(thenum As Integer64b) As Boolean '测是不是零
    IsZero64b = True
    For n = 0 To 7
        If thenum.by(n) <> 0 Then
           IsZero64b = False
           Exit For
        End If
    Next n
End Function
Rem==============================================================================
Rem==============================================================================
Private Function CheckWS8(a As Byte) As Byte  '测一个8位数实际位数
    Dim n As Byte
    For n = 0 To 7
        If ChkBit8b(a, n) Then CheckWS8 = n + 1
    Next n
End Function
Rem==============================================================================
Rem==============================================================================
Private Function CheckWS64(a As Integer64b) As Byte '测一个64位数实际位数
    Dim n As Byte
    For n = 0 To 63
        If ChkBit64b(a, n) Then CheckWS64 = n + 1
    Next n
End Function
Rem==============================================================================
Rem==============================================================================
Public Function HexStrToI64(a As String) As Integer64b '把十六进制表示的字符串转为64位数
    Dim tS As String
    tS = "0000000000000000"
    tS = Right$(tS & a, 16)
    For n = 0 To 7
        HexStrToI64.by(7 - n) = HexStrToI8(Mid$(tS, n * 2 + 1, 2))
    Next n
End Function
Rem==============================================================================
Rem==============================================================================
Public Function I64toHexStr(a As Integer64b) As String '把64位数转为十六进制字符串
    For n = 0 To 7
        I64toHexStr = I64toHexStr & I8toHexStr(a.by(7 - n))
    Next n
End Function
Rem==============================================================================
Rem==============================================================================
Public Function HexStrToI8(a As String) As Byte '把十六进制表示的字符串转为8位数
    Dim tS As String
    tS = "00"
    tS = Right$(tS & a, 2)
    HexStrToI8 = HexStrToI4(Mid$(tS, 1, 1)) * 16 + HexStrToI4(Mid$(tS, 2, 1))
End Function
Rem==============================================================================
Rem==============================================================================
Public Function HexStrToI4(a As String) As Byte '把十六进制表示的字符串转为4位数
    Dim n As Byte
    For n = 0 To 15
        If UCase$(a) = CHex(n) Then
            HexStrToI4 = n
            Exit For
        End If
    Next n
End Function
Rem==============================================================================
Rem==============================================================================
Public Function I8toHexStr(a As Byte) As String '把8位数转为十六进制字符串
    I8toHexStr = CHex(a \ 16) & CHex(a Mod 16)
End Function

Public Function Minus64b(a As Integer64b, b As Integer64b) As Integer64b '长整数减法
    Minus64b = Plus64b(a, Bu64b(b))
End Function
Rem==============================================================================
Rem==============================================================================
Public Function Compare64b(a As Integer64b, b As Integer64b) As Boolean  '比大小，大于或等于返回true小于则false
    For n = 0 To 7
        If a.by(7 - n) > b.by(7 - n) Then
            Compare64b = True
            Exit Function
        ElseIf a.by(7 - n) < b.by(7 - n) Then
            Compare64b = False
            Exit Function
        End If
    Next n
    Compare64b = True
End Function

Rem==============================================================================
Rem==============================================================================
Public Function And64b(a As Integer64b, b As Integer64b) As Integer64b  'And运算
    For n = 0 To 7
        And64b.by(n) = a.by(n) And b.by(n)  'And_7(a.by(n), b.by(n))
    Next n
End Function

Rem==============================================================================
Rem==============================================================================
Public Function Or64b(a As Integer64b, b As Integer64b) As Integer64b  'Or运算
    For n = 0 To 7
        Or64b.by(n) = a.by(n) Or b.by(n) 'Or_7(a.by(n), b.by(n))
    Next n
End Function

Rem==============================================================================
Rem==============================================================================
Public Function Not64b(a As Integer64b) As Integer64b  'Not运算
    For n = 0 To 7
        Not64b.by(n) = Not a.by(n)

    Next n
End Function

Rem==============================================================================
Rem==============================================================================
Public Function StrToI64(a As String) As Integer64b '把十进字符串转64b长整数
    Dim tS As String
    tS = "00000000000000000000"
    tS = Right$(tS & a, 20)
    For n = 1 To 20
        StrToI64 = Multi64b8b(StrToI64, 10)
        Call Plus64b8b(StrToI64, Val(Mid$(tS, n, 1)))
    Next n
End Function
Rem==============================================================================
Public Function I64toStr(a As Integer64b) As String '把长整数转为十进制字符串
    Dim tI As Integer64b
    Call let64b(tI, a)
    For n = 1 To 20
        I64toStr = CStr(Mod64b(tI, 10)) & I64toStr
        tI = Div64b(tI, 10)
    Next n
End Function
Rem==============================================================================
Public Function I64toStrNZ(a As Integer64b) As String '把长整数转为(前面没有0的十进制)字符串
    Dim b As String
    b = I64toStr(a)
    Do
        If Left$(b, 1) = "0" Then
            b = Right$(b, Len(b) - 1)
        Else
            Exit Do
        End If
    Loop
    If b = "" Then b = "0"
    I64toStrNZ = b
End Function

Rem==============================================================================
Public Function IsEqual64b(a As Integer64b, b As Integer64b) As Boolean '判断是否相等
Dim i As Integer, q As Boolean

q = True
For i = 0 To 7
     If a.by(i) <> b.by(i) Then
        q = False
        Exit For
     End If
Next i

IsEqual64b = q
End Function

Rem==============================================================================
Public Function FixHexStr_64(HexStr As String) As String  '填充16进制字符串
Dim i As Integer, Ub As Integer, TemStr As String
TemStr = HexStr
If Len(HexStr) >= 2 Then
    If Left(HexStr, 2) = "0x" Or Left(HexStr, 2) = "&H" Then
       TemStr = Right(HexStr, Len(HexStr) - 2)
    End If
End If

If Len(TemStr) > 64 Then
    FixHexStr_64 = Left(TemStr, 64)
    Exit Function
End If

If Len(TemStr) < 64 Then
    Ub = 64 - Len(TemStr)
    For i = 1 To Ub
         TemStr = TemStr & "0"
    Next i
End If

FixHexStr_64 = TemStr
End Function

'*************************************************************************
'**函 数 名：BinStrOr
'**输    入：(String)Bin1,(String)Bin2
'**输    出：(String) -
'**功能描述：二进制字符串或运算
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-28 22:20:40
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinStrOr(Bin1 As String, Bin2 As String) As String
Dim L As Long, i As Long, StrZ As String

If Len(Bin1) <> Len(Bin2) Then
    L = Abs(Len(Bin1) - Len(Bin2))
    For i = 1 To L
        StrZ = StrZ & "0"
    Next i
    
    If Len(Bin1) > Len(Bin2) Then
       Bin2 = StrZ & Bin2
    Else
       Bin1 = StrZ & Bin1
    End If

End If

For i = 1 To Len(Bin1)
    If Mid(Bin1, i, 1) = "1" Or Mid(Bin2, i, 1) = "1" Then
          BinStrOr = BinStrOr & "1"
    Else
          BinStrOr = BinStrOr & "0"
    End If
Next i

If BinStrOr = "" Then BinStrOr = "0"

End Function
'*************************************************************************
'**函 数 名：BinStrAnd
'**输    入：(String)Bin1,(String)Bin2
'**输    出：(String) -
'**功能描述：二进制字符串与运算
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-28 22:20:40
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinStrAnd(Bin1 As String, Bin2 As String) As String
Dim L As Long, i As Long, StrZ As String

If Len(Bin1) <> Len(Bin2) Then
    L = Abs(Len(Bin1) - Len(Bin2))
    For i = 1 To L
        StrZ = StrZ & "0"
    Next i
    
    If Len(Bin1) > Len(Bin2) Then
       Bin2 = StrZ & Bin2
    Else
       Bin1 = StrZ & Bin1
    End If

End If

For i = 1 To Len(Bin1)
    If Mid(Bin1, i, 1) = "1" And Mid(Bin2, i, 1) = "1" Then
          BinStrAnd = BinStrAnd & "1"
    Else
          BinStrAnd = BinStrAnd & "0"
    End If
Next i

If BinStrAnd = "" Then BinStrAnd = "0"

End Function
'*************************************************************************
'**函 数 名：BinStrNot
'**输    入：(String)Bin
'**输    出：(String) -
'**功能描述：二进制字符串非运算
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-28 22:20:40
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinStrNot(Bin As String) As String
Dim i As Long

For i = 1 To Len(Bin)
    If Mid(Bin, i, 1) = "1" Then
          BinStrNot = BinStrNot & "0"
    Else
          BinStrNot = BinStrNot & "1"
    End If
Next i

If BinStrNot = "" Then BinStrNot = "0"

End Function
'*************************************************************************
'**函 数 名：I64ToBinStr
'**输    入：(Integer64b)a
'**输    出：(String) -
'**功能描述：把64位长整数转为二进制字符串
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-29 21:46:30
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function I64ToBinStr(a As Integer64b) As String
    Dim tI As Integer64b
    Call let64b(tI, a)
    For n = 1 To 64
        I64ToBinStr = CStr(Mod64b(tI, 2)) & I64ToBinStr
        tI = Div64b(tI, 2)
    Next n
End Function

'*************************************************************************
'**函 数 名：ReplaceBinStr
'**输    入：(String)Bin,(Integer)Bit,(Integer)Rep
'**输    出：(String) -
'**功能描述：替换64位二进制字符串某位
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-29 22:11:20
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function ReplaceBinStr(Bin As String, Bit As Integer, Rep As Integer) As String

ReplaceBinStr = Left(Bin, Len(Bin) - Bit - 1) & CStr(Rep) & Right(Bin, Bit)

End Function

'*************************************************************************
'**函 数 名：BinToHex
'**输    入：(String)Bin
'**输    出：(String) -
'**功能描述：将二进制字符串转换为十六进制字符串
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-29 22:11:20
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinToHex(Bin As String) As String
Dim i As Long, H As String

If Len(Bin) Mod 4 <> 0 Then
      Bin = String(4 - Len(Bin) Mod 4, "0") & Bin
End If

For i = 1 To Len(Bin) Step 4
     Select Case Mid(Bin, i, 4)
        Case "0000"
            H = H & "0"
        Case "0001"
            H = H & "1"
        Case "0010"
            H = H & "2"
        Case "0011"
            H = H & "3"
        Case "0100"
            H = H & "4"
        Case "0101"
            H = H & "5"
        Case "0110"
            H = H & "6"
        Case "0111"
            H = H & "7"
        Case "1000"
            H = H & "8"
        Case "1001"
            H = H & "9"
        Case "1010"
            H = H & "A"
        Case "1011"
            H = H & "B"
        Case "1100"
            H = H & "C"
        Case "1101"
            H = H & "D"
        Case "1110"
            H = H & "E"
        Case "1111"
            H = H & "F"
     End Select
Next i

While Left(H, 1) = "0"
    H = Right(H, Len(H) - 1)
Wend

BinToHex = H
End Function

'*************************************************************************
'**函 数 名：HexToBin
'**输    入：(String)strHex
'**输    出：(String) -
'**功能描述：将十六进制字符串转换为二进制字符串
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 10:23:25
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function HexToBin(strHex As String) As String
Dim i As Long, b As String, k As String * 1

For i = 1 To Len(strHex) Step 1
     k = UCase(Mid(strHex, i, 4))
     Select Case k
        Case "0"
            b = b & "0000"
        Case "1"
            b = b & "0001"
        Case "2"
            b = b & "0010"
        Case "3"
            b = b & "0011"
        Case "4"
            b = b & "0100"
        Case "5"
            b = b & "0101"
        Case "6"
            b = b & "0110"
        Case "7"
            b = b & "0111"
        Case "8"
            b = b & "1000"
        Case "9"
            b = b & "1001"
        Case "A"
            b = b & "1010"
        Case "B"
            b = b & "1011"
        Case "C"
            b = b & "1100"
        Case "D"
            b = b & "1101"
        Case "E"
            b = b & "1110"
        Case "F"
            b = b & "1111"
     End Select
Next i

HexToBin = b
End Function

'*************************************************************************
'**函 数 名：And_7
'**输    入：(Byte)Num_De1,(Byte)Num_De2
'**输    出：(Byte) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 08:41:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function And_7(ByVal Num_De1 As Byte, ByVal Num_De2 As Byte) As Byte
Dim sB(1 To 2) As String * 8, i As Integer, k(1 To 2) As String, j As Integer, res As Integer, TemResult As String
sB(1) = DetoBinString_7(Num_De1)
sB(2) = DetoBinString_7(Num_De2)

For i = 1 To 8
     For j = 1 To 2
          k(j) = Mid(sB(j), i, 1)
     Next j
     
     res = Val(k(1)) And Val(k(2))
     
     TemResult = TemResult & res
Next i

And_7 = BinStringtoDe(TemResult)
End Function
'*************************************************************************
'**函 数 名：Or_7
'**输    入：(Byte)Num_De1,(Byte)Num_De2
'**输    出：(Byte) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 08:41:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function Or_7(ByVal Num_De1 As Byte, ByVal Num_De2 As Byte) As Byte
Dim sB(1 To 2) As String * 8, i As Integer, k(1 To 2) As String, j As Integer, res As Integer, TemResult As String
sB(1) = DetoBinString_7(Num_De1)
sB(2) = DetoBinString_7(Num_De2)

For i = 1 To 8
     For j = 1 To 2
          k(j) = Mid(sB(j), i, 1)
     Next j
     
     res = Val(k(1)) Or Val(k(2))
     
     TemResult = TemResult & res
Next i

Or_7 = BinStringtoDe(TemResult)
End Function

'*************************************************************************
'**函 数 名：Not_7
'**输    入：(Byte)Num_De
'**输    出：(Byte) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 08:41:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function Not_7(ByVal Num_De As Byte) As Byte
Dim sB As String * 8, i As Integer, k As String, res As String, TemResult As String
sB = DetoBinString_7(Num_De)

For i = 1 To 8
          k = Mid(sB, i, 1)
     
     If k = "1" Then
     res = "0"
     Else
     res = "1"
     End If
     
     TemResult = TemResult & res
Next i

Not_7 = BinStringtoDe(TemResult)
End Function

'*************************************************************************
'**函 数 名：DetoBinString_7
'**输    入：(Byte)Num_De
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 08:41:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DetoBinString_7(ByVal Num_De As Byte) As String
Dim i As Integer, Dec As Byte
Dec = Num_De

For i = 0 To 7 Step 1
       If Dec Mod 2 <> 0 Then
         DetoBinString_7 = DetoBinString_7 & "1"
       Else
         DetoBinString_7 = DetoBinString_7 & "0"
       End If
         Dec = Dec \ 2
      DoEvents
Next i
End Function

'*************************************************************************
'**函 数 名：BinStringtoDe_7
'**输    入：(String)StrBin
'**输    出：(byte) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 17:16:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function BinStringtoDe_7(ByVal StrBin As String) As Byte
Dim i As Integer, k As String

If Trim(StrBin) = "" Then
   Exit Function
End If

For i = 1 To 8 Step 1
      k = Mid(StrBin, i, 1)
      
      If k = "1" Then
          BinStringtoDe_7 = BinStringtoDe_7 + 2 ^ (i - 1)
      End If
Next i
End Function

'*************************************************************************
'**函 数 名：AddFlags64b
'**输    入：(Integer64b)Flags,(Integer64b)NewFlags
'**输    出：(Integer64b) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-06 23:02:51
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function AddFlags64b(Flags As Integer64b, NewFlags As Integer64b) As Integer64b
AddFlags64b = Or64b(Flags, NewFlags)
End Function

'*************************************************************************
'**函 数 名：DeleteFlags64b
'**输    入：(Integer64b)Flags,(Integer64b)NewFlags
'**输    出：(Integer64b) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-06 23:03:15
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DeleteFlags64b(Flags As Integer64b, NewFlags As Integer64b) As Integer64b
DeleteFlags64b = And64b(Flags, Not64b(NewFlags))
End Function

'*************************************************************************
'**函 数 名：RemoveUseless0
'**输    入：(Str)Num
'**输    出：(Str) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-2-17 10:35:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function RemoveUseless0(Num As String) As String
Dim i As Integer

For i = 1 To Len(Num)
    If Mid(Num, i, 1) <> "0" Then
        Exit For
    End If
Next i
RemoveUseless0 = Right(Num, Len(Num) - i + 1)

End Function
