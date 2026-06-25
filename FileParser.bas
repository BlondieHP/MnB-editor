Attribute VB_Name = "FileParser"

Option Explicit

Public Pointer As Long
Public LinePointer As Long

Public bigNumArray() As Byte
Public MaxPointer As Long

Public lngHandle As Long
Public txtLine As String

' Buffer mode flag: when True, GetWord/GetLine/GetRealLine read from
' bigNumArray (in-memory buffer) instead of byte-by-byte from disk.
' Set by InitBufferRead, cleared by ClearBuffer.
Private gBufferMode As Boolean

'*************************************************************************
' Reads the entire file into bigNumArray() for fast in-memory parsing.
' After calling this, GetWord/GetLine/GetRealLine will read from the
' buffer instead of byte-by-byte from disk (much faster).
'
' Usage:
'   InitBufferRead "C:\path\to\troops.txt"
'   N_Troop = Val(GetWord())   ' reads from buffer at full memory speed
'   ClearBuffer
'*************************************************************************
Public Sub InitBufferRead(ByVal FilePath As String)
    On Error GoTo errorHandle

    Dim hFile As Long
    Dim fileSize As Long

    hFile = FreeFile()
    Open FilePath For Binary Access Read As #hFile
    fileSize = LOF(hFile)
    If fileSize > 0 Then
        ReDim bigNumArray(0 To fileSize - 1)
        Get #hFile, , bigNumArray
    Else
        ReDim bigNumArray(0 To 0)
    End If
    Close #hFile

    Pointer = 1
    MaxPointer = UBound(bigNumArray) + 1
    gBufferMode = True

    Exit Sub

errorHandle:
    gBufferMode = False
    Call logErr("FileParser", "InitBufferRead:[" & FilePath & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
' Releases the file buffer and resets buffer mode.
' After calling this, GetWord/GetLine/GetRealLine revert to
' byte-by-byte disk reads.
'*************************************************************************
Public Sub ClearBuffer()
    gBufferMode = False
    ReDim bigNumArray(0 To 0)
    MaxPointer = 0
    Pointer = 1
End Sub

'*************************************************************************
' Reads the next whitespace-delimited word.
' Uses buffered read from bigNumArray when gBufferMode is True,
' otherwise reads byte-by-byte from lngHandle.
'*************************************************************************
Function GetWord() As String
    On Error GoTo errorHandle

    GetWord = ""
    Dim tmp As Byte

    Do While Pointer <= MaxPointer
        If gBufferMode Then
            tmp = bigNumArray(Pointer - 1)
        Else
            Get lngHandle, Pointer, tmp
        End If
        Pointer = Pointer + 1

        If tmp = 10 Or tmp = 13 Or tmp = 32 Then
            If GetWord <> "" Then Exit Do
        Else
            GetWord = GetWord & Chr(tmp)
        End If
    Loop

    Exit Function

errorHandle:
    Call logErr("FileParser", "GetWord", Err.Number, Err.Description)
End Function

'*************************************************************************
' Reads the next whitespace-delimited word from txtLine string
' at current LinePointer position. Advances LinePointer past the word.
' Sets bEnd = True if end of string reached without finding content.
'*************************************************************************
Function GetWordL(Optional bEnd As Boolean) As String
    On Error GoTo errorHandle

    GetWordL = ""
    Dim tmp As Byte
    bEnd = False

    Do While LinePointer <= Len(txtLine)
        tmp = Asc(Mid(txtLine, LinePointer, 1))

        If tmp = 10 Or tmp = 13 Or tmp = 32 Then
            If GetWordL <> "" Then Exit Do
        Else
            GetWordL = GetWordL & Chr(tmp)
        End If

        LinePointer = LinePointer + 1
    Loop

    If LinePointer > Len(txtLine) Then bEnd = True

    Exit Function

errorHandle:
    Call logErr("FileParser", "GetWordL", Err.Number, Err.Description)
End Function

'*************************************************************************
' Writes a word (string) to the open binary file at current Pointer,
' skipping any leading space character. Advances Pointer.
'*************************************************************************
Sub PutWord(TheWord As String)
    On Error GoTo errorHandle
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

    Exit Sub

errorHandle:
    Call logErr("FileParser", "PutWord", Err.Number, Err.Description)
End Sub

'*************************************************************************
' Writes a single space character (ASCII 32) to the open binary file
' at current Pointer. Advances Pointer by 1.
'*************************************************************************
Sub PutSpc()
    On Error GoTo errorHandle
    Dim Mspc As Byte

    Mspc = 32
    Put lngHandle, Pointer, Mspc
    Pointer = Pointer + 1

    Exit Sub

errorHandle:
    Call logErr("FileParser", "PutSpc", Err.Number, Err.Description)
End Sub

'*************************************************************************
' Writes CR+LF (ASCII 13, 10) line ending to the open binary file
'*************************************************************************
Sub PutReturn()
    On Error GoTo errorHandle
    Dim tmp As Byte

    tmp = 13
    Put lngHandle, Pointer, tmp
    Pointer = Pointer + 1
    tmp = 10
    Put #lngHandle, Pointer, tmp
    Pointer = Pointer + 1

    Exit Sub

errorHandle:
    Call logErr("FileParser", "PutReturn", Err.Number, Err.Description)
End Sub

'*************************************************************************
' Reads until next CR/LF from the open file or buffer.
'*************************************************************************
Function GetLine() As String
    On Error GoTo errorHandle
    GetLine = ""
    Dim tmp As Byte

    Do While Pointer <= MaxPointer
        If gBufferMode Then
            tmp = bigNumArray(Pointer - 1)
        Else
            Get lngHandle, Pointer, tmp
        End If
        Pointer = Pointer + 1
        If Pointer > MaxPointer Then Exit Do
        If tmp = 13 Or tmp = 10 Then
            If GetLine <> "" Then Exit Do
        Else
            GetLine = GetLine & Chr(tmp)
        End If
    Loop

    Exit Function

errorHandle:
    Call logErr("FileParser", "GetLine", Err.Number, Err.Description)
End Function

'*************************************************************************
' Reads a complete line including consuming the CR+LF terminator.
'*************************************************************************
Function GetRealLine() As String
    On Error GoTo errorHandle
    GetRealLine = ""
    Dim tmp As Byte

    Do While Pointer <= MaxPointer
        If gBufferMode Then
            tmp = bigNumArray(Pointer - 1)
        Else
            Get lngHandle, Pointer, tmp
        End If
        Pointer = Pointer + 1
        If Pointer > MaxPointer Then Exit Do
        If tmp = 13 Then
            If gBufferMode Then
                ' consume LF from buffer
                Pointer = Pointer + 1
            Else
                Get lngHandle, Pointer, tmp   ' consume LF
                Pointer = Pointer + 1
            End If
            Exit Do
        Else
            GetRealLine = GetRealLine & Chr(tmp)
        End If
    Loop

    Exit Function

errorHandle:
    Call logErr("FileParser", "GetRealLine", Err.Number, Err.Description)
End Function

'*************************************************************************
' Converts a large number string to its signed 32-bit representation.
'*************************************************************************
Function MinusIT(thenum As String) As Long
    On Error GoTo errorHandle
    Dim n1, n2 As String
    Dim Num1, Num2 As Long

    n1 = "2000000000"
    n2 = Right(thenum, 9)
    Num1 = Val(n1)
    Num2 = Val(n2)
    Num1 = Num1 - 2147483647 - 1
    MinusIT = Num1 + Num2

    Exit Function

errorHandle:
    Call logErr("FileParser", "MinusIT", Err.Number, Err.Description)
End Function
