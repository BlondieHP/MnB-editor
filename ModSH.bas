Attribute VB_Name = "ModSH"
Option Explicit

Public Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long

Public Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Const MAX_LEN = 200 '字符串最大长度

Public Const DESKTOP = &H0& '桌面

Public Const PROGRAMS = &H2& '程序集

Public Const MYDOCUMENTS = &H5& '我的文档

Public Const MYFAVORITES = &H6& '收藏夹

Public Const STARTUP = &H7& '启动

Public Const RECENT = &H8& '最近打开的文件

Public Const SENDTO = &H9& '发送

Public Const STARTMENU = &HB& '开始菜单

Public Const NETHOOD = &H13& '网上邻居

Public Const FONTS = &H14& '字体

Public Const SHELLNEW = &H15& 'ShellNew

Public Const APPDATA = &H1A& 'Application Data

Public Const PRINTHOOD = &H1B& 'PrintHood

Public Const PAGETMP = &H20& '网页临时文件

Public Const COOKIES = &H21& 'Cookies目录

Public Const HISTORY = &H22& '历史

Public Function GetAppDataPath()
Dim sTmp As String * MAX_LEN '存放结果的固定长度的字符串

Dim pidl As Long '某特殊目录在特殊目录列表中的位置


SHGetSpecialFolderLocation 0, APPDATA, pidl

SHGetPathFromIDList pidl, sTmp

GetAppDataPath = Left(sTmp, InStr(sTmp, Chr(0)) - 1)

End Function
