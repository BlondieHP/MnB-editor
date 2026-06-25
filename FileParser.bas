Attribute VB_Name = "FileParser"

Option Explicit

Public Pointer As Long '读文件时的指针
Public LinePointer As Long '读文本行时的指针

Public bigNumArray() As Byte
Public MaxPointer As Long '最大指针位置.

Public lngHandle As Long '文件句柄
Public txtLine As String '文本内容

Function GetWord() As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    GetWord = ""
    Dim tmp As Byte
before:
    If Pointer > MaxPointer Then
        Call logErr("FileParser", "GetWord", "OVER_MAX_POINTER", "超出最大文件指针!")
        Exit Function
    End If
    Get lngHandle, Pointer, tmp
    Pointer = Pointer + 1

    If tmp = 10 Or tmp = 13 Or tmp = 32 Then
        If GetWord <> "" Then Exit Function
    Else
        GetWord = GetWord & Chr(tmp)
    End If
    GoTo before
    
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileParser", "GetWord", Err.Number, Err.Description)
End Function

Function GetWordL(Optional bEnd As Boolean) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    GetWordL = ""
    Dim tmp As Byte
    bEnd = False
before:
    If LinePointer > Len(txtLine) Then
        bEnd = True
        Exit Function
    End If
    tmp = Asc(Mid(txtLine, LinePointer, 1))

    If tmp = 10 Or tmp = 13 Or tmp = 32 Then
        If GetWordL <> "" Then Exit Function
    Else
        GetWordL = GetWordL & Chr(tmp)
    End If
    
    LinePointer = LinePointer + 1
    
    GoTo before
    
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileParser", "GetWordL", Err.Number, Err.Description)
End Function

Sub PutWord(TheWord As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim tmp As Byte
    Dim n As Integer
    
    
    For n = 1 To Len(TheWord)
        tmp = Asc(Mid(TheWord, n, 1))
        If tmp = 32 And n = 1 Then
            n = 2
            tmp = Asc(Mid(TheWord, n, 1))
        End If
        Put lngHandle, Pointer, tmp
        Pointer = Pointer + 1
    Next n
    

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileParser", "PutWord", Err.Number, Err.Description)
End Sub

Sub PutSpc()
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim Mspc As Byte

    Mspc = 32
    Put lngHandle, Pointer, Mspc
    Pointer = Pointer + 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileParser", "PutSpc", Err.Number, Err.Description)
End Sub

Sub PutReturn()
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim tmp As Byte
   
    tmp = 13
    Put lngHandle, Pointer, tmp
    Pointer = Pointer + 1
    tmp = 10
    Put #lngHandle, Pointer, tmp
    Pointer = Pointer + 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileParser", "PutReturn", Err.Number, Err.Description)
End Sub

Function GetLine() As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    GetLine = ""
    Dim tmp As Byte
    
   
before:
    Get lngHandle, Pointer, tmp
    Pointer = Pointer + 1
    If Pointer > MaxPointer Then
        Exit Function
    End If
    If tmp = 13 Or tmp = 10 Then
        If GetLine <> "" Then Exit Function
    Else
        GetLine = GetLine & Chr(tmp)
    End If
    GoTo before

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileParser", "GetLine", Err.Number, Err.Description)
End Function

Function GetRealLine() As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    GetRealLine = ""
    Dim tmp As Byte

before:
    Get lngHandle, Pointer, tmp
    Pointer = Pointer + 1
    If Pointer > MaxPointer Then
        Exit Function
    End If
    'If tmp = 13 Or tmp = 10 Then
    If tmp = 13 Then
        Get lngHandle, Pointer, tmp
        Pointer = Pointer + 1
        Exit Function
    Else
        GetRealLine = GetRealLine & Chr(tmp)
    End If
    GoTo before

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileParser", "GetRealLine", Err.Number, Err.Description)
End Function

Function MinusIT(thenum As String) As Long
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim n1, n2 As String
    Dim Num1, Num2 As Long
  
    n1 = "2000000000"
    n2 = Right(thenum, 9)
    Num1 = Val(n1)
    Num2 = Val(n2)
    Num1 = Num1 - 2147483647 - 1
    MinusIT = Num1 + Num2

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileParser", "MinusIT", Err.Number, Err.Description)
End Function
