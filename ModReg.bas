Attribute VB_Name = "ModReg"
Option Explicit
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
Dim R As Long

Public Sub SaveKey(Hkey As Long, strPath As String)
     '´æ´¢¼ü
     Dim Keyhand&
     R = RegCreateKey(Hkey, strPath, Keyhand&)
     R = RegCloseKey(Keyhand&)
End Sub

Public Function ReadRegString(Hkey As Long, strPath As String, strValue As String) As String
     '»ñÈ¡×Ö·ûÐÍ
     Dim Keyhand As Long
     Dim datatype As Long
     Dim lResult As Long
     Dim strBuf As String
     Dim lDataBufSize As Long
     Dim intZeroPos As Integer, lValueType As Long
     R = RegOpenKey(Hkey, strPath, Keyhand)
     lResult = RegQueryValueEx(Keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
     If lValueType = REG_SZ Then
         strBuf = String(lDataBufSize, " ")
         lResult = RegQueryValueEx(Keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
         If lResult = ERROR_SUCCESS Then
             intZeroPos = InStr(strBuf, Chr$(0))
             If intZeroPos > 0 Then
                 ReadRegString = Left$(strBuf, intZeroPos - 1)
             Else
                 ReadRegString = strBuf
             End If
         End If
     End If
End Function

Public Sub WriteRegString(Hkey As Long, strPath As String, strValue As String, strdata As String)
     '´æ´¢×Ö·ûÐÍ
     Dim Keyhand As Long
     Dim R As Long
     R = RegCreateKey(Hkey, strPath, Keyhand)
     R = RegSetValueEx(Keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
     R = RegCloseKey(Keyhand)
End Sub

Function ReadRegDWord(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
     '»ñÈ¡µ¥×Ö½Ú
     Dim lResult As Long
     Dim lValueType As Long
     Dim lBuf As Long
     Dim lDataBufSize As Long
     Dim R As Long
     Dim Keyhand As Long
     R = RegOpenKey(Hkey, strPath, Keyhand)
     lDataBufSize = 4
     lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
     If lResult = ERROR_SUCCESS Then
         If lValueType = REG_DWORD Then
             ReadRegDWord = lBuf
         End If
     End If
     R = RegCloseKey(Keyhand)
End Function

Function WriteRegDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
     '´æ´¢Ë«×Ö½Ú
     Dim lResult As Long
     Dim Keyhand As Long
     Dim R As Long
     R = RegCreateKey(Hkey, strPath, Keyhand)
     lResult = RegSetValueEx(Keyhand, strValueName, 0&, REG_DWORD, lData, 4)
     R = RegCloseKey(Keyhand)
End Function

Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
     'É¾³ý¼ü
     Dim R As Long
     R = RegDeleteKey(Hkey, strKey)
End Function

Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
     'É¾³ý¼üÖµ
     Dim Keyhand As Long
     R = RegOpenKey(Hkey, strPath, Keyhand)
     R = RegDeleteValue(Keyhand, strValue)
     R = RegCloseKey(Keyhand)
End Function
