Attribute VB_Name = "ModMemory"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long

Public Function HiByte(ByVal wParam As Integer) As Byte
    
    'note:   VB4-32   users   should   declare   this   function   As   Integer
      HiByte = (wParam And &HFF00&) \ (&H100)
  
End Function

Public Function HiByteCM(ByVal wParam As Integer) As Byte
    
      CopyMemory HiByteCM, ByVal VarPtr(wParam) + 1, 1
  
End Function


Public Function LoByte(ByVal wParam As Integer) As Byte

    'note:   VB4-32   users   should   declare   this   function   As   Integer
      LoByte = wParam And &HFF&

End Function

Public Function LoByteCM(ByVal wParam As Integer) As Byte
    
      CopyMemory LoByteCM, wParam, 1
  
End Function


Public Function HiWord(wParam As Long) As Integer

      If wParam And &H80000000 Then
            HiWord = (wParam \ 65535) - 1
      Else
            HiWord = wParam \ 65535
      End If

End Function

Public Function HiWordCM(ByVal wParam As Long) As Integer
    
      CopyMemory HiWordCM, ByVal VarPtr(wParam) + 2, 2
  
End Function

Public Function LoWord(wParam As Long) As Integer

      If wParam And &H8000& Then
            LoWord = &H8000& Or (wParam And &H7FFF&)
      Else
            LoWord = wParam And &HFFFF&
      End If

End Function


Public Function LoWordCM(wParam As Long) As Integer

    'using   API
      CopyMemory LoWordCM, wParam, 2
    
End Function

Public Function LShiftWord(ByVal w As Integer, ByVal c As Integer) As Integer

      Dim dw     As Long
      dw = w * (2 ^ c)
      If dw And &H8000& Then
            LShiftWord = CInt(dw And &H7FFF&) Or &H8000&
      Else
            LShiftWord = dw And &HFFFF&
      End If

End Function


Public Function RShiftWord(ByVal w As Integer, ByVal c As Integer) As Integer

      Dim dw     As Long
      If c = 0 Then
            RShiftWord = w
      Else
            dw = w And &HFFFF&
            dw = dw \ (2 ^ c)
            RShiftWord = dw And &HFFFF&
      End If

End Function

Public Function SplitLongAsInteger(ByVal lParam As Long, Optional Site As Long) As Integer

CopyMemory SplitLongAsInteger, ByVal VarPtr(lParam) + Site * 2, 2

End Function

Public Function SplitLongAsByte(ByVal lParam As Long, Optional Site As Long) As Byte

CopyMemory SplitLongAsByte, ByVal VarPtr(lParam) + Site, 1

End Function

Public Function SplitIntegerAsByte(ByVal iParam As Integer, Optional Site As Long) As Byte

CopyMemory SplitIntegerAsByte, ByVal VarPtr(iParam) + Site, 1

End Function

Public Function SplitStringAsByte(ByVal sParam As String, Optional Site As Long) As Byte

CopyMemory SplitStringAsByte, ByVal StrPtr(sParam) + Site, 1

End Function

Public Function MakeInteger(ByVal bHigh As Byte, ByVal bLow As Byte) As Integer

CopyMemory MakeInteger, bLow, 1

CopyMemory ByVal VarPtr(MakeInteger) + 1, bHigh, 1

End Function

Public Function MakeLongFromIntegers(ByVal iHigh As Integer, ByVal iLow As Integer) As Long

CopyMemory MakeLongFromIntegers, iLow, 2

CopyMemory ByVal VarPtr(MakeLongFromIntegers) + 2, iHigh, 2

End Function

Public Function MakeLongFromBytes(Bytes() As Byte) As Long

CopyMemory MakeLongFromBytes, Bytes(0), 1

CopyMemory ByVal VarPtr(MakeLongFromBytes) + 1, Bytes(1), 1

CopyMemory ByVal VarPtr(MakeLongFromBytes) + 2, Bytes(2), 1

CopyMemory ByVal VarPtr(MakeLongFromBytes) + 3, Bytes(3), 1

End Function

Public Function MakeStringFromBytes(Bytes() As Byte) As String
Dim i As Long, strInt As Integer

MakeStringFromBytes = Space((UBound(Bytes) + 1) \ 2)
For i = 0 To ((UBound(Bytes) + 1) \ 2) * 2 - 1
    CopyMemory ByVal StrPtr(MakeStringFromBytes) + i, Bytes(i), 1
Next i

End Function

Public Sub LeftMove(ByVal Address As Long, ByVal Length As Long)
Dim i As Long

For i = 0 To Length - 2
    CopyMemory ByVal Address + i, ByVal Address + i + 1, 1
Next i

CopyMemory Address + Length - 1, 0, 1
End Sub

Public Sub RightMove(ByVal Address As Long, ByVal Length As Long)
Dim i As Long

For i = Length - 2 To 0 Step -1
    CopyMemory ByVal Address + i + 1, ByVal Address + i, 1
Next i

CopyMemory Address, 0, 1
End Sub

Public Function MakeWord(ByVal bHi As Byte, ByVal bLo As Byte) As Integer                       '两个byte转换为一个integer

      If bHi And &H80 Then
            MakeWord = (((bHi And &H7F) * 256) + bLo) Or &H8000&
      Else
            MakeWord = (bHi * 256) + bLo
      End If

End Function


Public Function MakeDWord(wHi As Integer, wLo As Integer) As Long                   '两个integer转换为一个long

      If wHi And &H8000& Then
            MakeDWord = (((wHi And &H7FFF&) * 65536) Or _
                                        (wLo And &HFFFF&)) Or &H80000000
      Else
            MakeDWord = (wHi * 65535) + wLo
      End If

End Function

Public Function MakeDWordFromBytes(b() As Byte) As Long                    '四个byte转换为一个long
Dim i(1) As Integer
      
i(0) = MakeWord(b(1), b(0))
i(1) = MakeWord(b(3), b(2))

MakeDWordFromBytes = MakeDWord(i(1), i(0))
End Function

