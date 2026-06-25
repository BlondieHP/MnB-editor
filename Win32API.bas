Attribute VB_Name = "Win32API"
Option Explicit

'API
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long '获取窗体样式
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long '设置窗体样式
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long


Public Type SHFILEOPSTRUCT
    hWnd   As Long
    wFunc   As Long
    pFrom   As String
    pTo   As String
    fFlags   As Integer
    fAborted   As Boolean
    hNameMaps   As Long
    sProgress   As String
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type LVHITTESTINFO
   pt As POINTAPI
   Flags As Long
   iItem As Long
   iSubItem As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetPrivateProfileInt Lib "kernel32" _
    Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal nDefault As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Public ErrorMsg As String

Public Const WM_USER = &H400

Public Const LB_SETHORIZONTALEXTENT = WM_USER + 21
Public Const LB_GETSELITEMS = &H191
Public Const SPI_GETWORKAREA = 48

'虚拟键
Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10
Public Const VK_MENU = &H12


'消息常数
Public Const GWL_STYLE = (-16) '窗体样式
Public Const WS_CAPTION = &HC00000 '带标题窗体
Public Const WS_MAXIMIZEBOX = &H10000 '带最大化按钮
Public Const WS_MINIMIZEBOX = &H20000 '带最小化
Public Const WS_THICKFRAME = &H40000 '窗体可调
Public Const MY_NOT_SIZABLE = 382337024  '窗体不可调
Public Const MY_SIZABLE = 382664704 '窗体不可调
 
Public Const FO_DELETE = &H3

Public Const FOF_ALLOWUNDO = &H40           '   移入回收站
Public Const FOF_CONFIRMMOUSE = &H2           '   删除。不放入回收站
Public Const FOF_NOCONFIRMATION = &H10           '   没有提示

Public Const EM_SETTARGETDEVICE = (WM_USER + 72)

'文本框相关
Public Const EM_GETSEL = &HB0
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB

'ListView
Public Const LVM_FIRST = &H1000
Public Const LVM_HITTEST = LVM_FIRST + 18
Public Const LVM_GETCOUNTPERPAGE = LVM_FIRST + 40
Public Const LVM_SETCOLUMNWIDTH = &H1000 + 30
Public Const LVSCW_AUTOSIZE_USEHEADER = -2


'*************************************************************************
'**函 数 名：ErrorMsg_Initialize
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:43:13
'**修 改 人：
'**日    期：
'**版    本：V0.9
'*************************************************************************
Public Sub ErrorMsg_Initialize()
    Call Class_Initialize
End Sub

'*************************************************************************
'**函 数 名：Class_Initialize
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:43:06
'**修 改 人：
'**日    期：
'**版    本：V0.9
'*************************************************************************
Private Sub Class_Initialize()
    ErrorMsg = vbNullString
End Sub

'*************************************************************************
'**函 数 名：WriteString
'**输    入：iniFileName(String) -
'**        ：Section(String)     -
'**        ：Key(String)         -
'**        ：Value(String)       -
'**输    出：(Boolean) -
'**功能描述：写入
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:42:48
'**修 改 人：
'**日    期：
'**版    本：V0.951.13
'*************************************************************************
Public Function WriteString(iniFileName As String, Section As String, Key As String, Value As String) As Boolean
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim F As Integer
    
    WriteString = False
    ErrorMsg = vbNullString
    If Not FileExists(iniFileName) Then
        'ErrorMsg = "INI file has not been specifyed!"
        'Exit Function
        F = FreeFile
        
        Open iniFileName For Output As #F
        Close #F
    End If
    If WritePrivateProfileString(Section, Key, Value, iniFileName) = 0 Then
        ErrorMsg = "Failed to write to the ini file!"
        Exit Function
    End If
 
    WriteString = True
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Win32API", "WriteString", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：ReadString
'**输    入：iniFileName(String) -
'**        ：Section(String)     -
'**        ：Key(String)         -
'**        ：Size(Long)          -
'**输    出：(String) -
'**功能描述：读出字符串
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:41:43
'**修 改 人：
'**日    期：
'**版    本：V0.9
'*************************************************************************
Public Function ReadString(iniFileName As String, Section As String, Key As String, Optional Size As Long = 255) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim ReturnStr As String
    Dim ReturnLng As Long
    ErrorMsg = vbNullString
    ReadString = vbNullString
    If Not FileExists(iniFileName) Then
        ErrorMsg = "INI file has not been specifyed!"
        Exit Function
    End If
    ReturnStr = Space(Size)
    ReturnLng = GetPrivateProfileString(Section, Key, vbNullString, ReturnStr, Size, iniFileName)
    'ReadString = Left(ReturnStr, ReturnLng)
    ReadString = StripTerFlag(ReturnStr)
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Win32API", "ReadString", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：ReadInt
'**输    入：iniFileName(String) -
'**        ：Section(String)     -
'**        ：Key(String)         -
'**输    出：(Long) -
'**功能描述：读出数值
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:41:50
'**修 改 人：
'**日    期：
'**版    本：V0.9
'*************************************************************************
Public Function ReadInt(iniFileName As String, Section As String, Key As String) As Long
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim ReturnLng As Long
    ReadInt = 0
    ErrorMsg = vbNullString
    If iniFileName = "" Then
        ErrorMsg = "INI file has not been specifyed!"
        Exit Function
    End If
    ReturnLng = GetPrivateProfileInt(Section, Key, 0, iniFileName)
    If ReturnLng = 0 Then
        ReturnLng = GetPrivateProfileInt(Section, Key, 1, iniFileName)
        If ReturnLng = 1 Then
            ErrorMsg = "Can not read the ini file!"
            Exit Function
        End If
    End If
    ReadInt = ReturnLng
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Win32API", "ReadInt", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：ChangeSectionName
'**输    入：iniFileName(String) -
'**        ：oSection(String)     -
'**        ：nSection(String)         -
'**输    出：(Long) -
'**功能描述：更改键名
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-16 14:01:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function ChangeSectionName(iniFileName As String, oSection As String, nSection As String) As Long
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim F As Long
    
    F = FreeFile
    
    Open iniFileName For Input As #F
      
    Close #F
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Win32API", "ReadInt", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：GetMyDocumentDirectory
'**输    入：Null
'**输    出：(Long) -
'**功能描述：读出数值
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:41:50
'**修 改 人：
'**日    期：
'**版    本：V0.9
'*************************************************************************
Public Function GetMyDocumentDirectory() As String
Dim ObjTem As Object

Set ObjTem = CreateObject("Shell.Application")

GetMyDocumentDirectory = ObjTem.NameSpace(5).Self.Path

End Function

'*************************************************************************
'**函 数 名：SwitchSizable
'**输    入：(Form)ObjForm
'**输    出：(Long) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-30 21:25:58
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwitchSizable(ObjForm As Form)
  Dim WinStyle     As Long
  WinStyle = GetWindowLong(ObjForm.hWnd, GWL_STYLE) '取得窗体样式
  SetWindowLong ObjForm.hWnd, GWL_STYLE, WinStyle Xor (WS_MAXIMIZEBOX Or WS_THICKFRAME)  '设置样式 用原来样式 异或 最大化最小化标题栏3样式属性 也就是原来有的就成没有 反之亦然
  
  ObjForm.Hide
  ObjForm.Show
End Sub

'*************************************************************************
'**函 数 名：SetFormStyle
'**输    入：(Form)ObjForm,(Long)Style
'**输    出：(Long) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-30 22:17:15
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SetFormStyle(ObjForm As Form, ByVal Style As Long)
  SetWindowLong ObjForm.hWnd, GWL_STYLE, Style   '设置样式 用原来样式 异或 最大化最小化标题栏3样式属性 也就是原来有的就成没有 反之亦然
  ObjForm.Hide
  ObjForm.Show
End Sub


'*************************************************************************
'**函 数 名：DeleteFolder
'**输    入：(String)sObject
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-01 23:06:09
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub DeleteFolder(sObject As String)
    Dim SHFileOp     As SHFILEOPSTRUCT

    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sObject
        .fFlags = FOF_CONFIRMMOUSE Or FOF_NOCONFIRMATION
    End With
    SHFileOperation SHFileOp
End Sub


'*************************************************************************
'**函 数 名：GetListViewVisibleLines
'**输    入：(ListView)ListView
'**输    出：(Long)-
'**功能描述：获得listview可见行数
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-04-11 23:15:45
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetListViewVisibleLines(ListView As ListView) As Long
    GetListViewVisibleLines = SendMessage(ListView.hWnd, LVM_GETCOUNTPERPAGE, 0, 0)
End Function


'*************************************************************************
'**函 数 名：GetListViewItemIndexUnderMousePointer
'**输    入：(ListView)ListView
'**输    出：(Long)-
'**功能描述：获得listview鼠标指针下的项目index
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-04-11 23:47:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetListViewItemIndexUnderMousePointer(ListView As ListView, X As Long, Y As Long) As Long
Dim lvhti As LVHITTESTINFO

   lvhti.pt.X = X / Screen.TwipsPerPixelX
   lvhti.pt.Y = Y / Screen.TwipsPerPixelY
   GetListViewItemIndexUnderMousePointer = SendMessage(ListView.hWnd, LVM_HITTEST, 0, lvhti) + 1
End Function

'*************************************************************************
'**函 数 名：GetTaskbarHeight
'**输    入：(ListView)ListView
'**输    出：(Long)-
'**功能描述：获得开始菜单栏高度
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-04-14 17:47:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

'*************************************************************************
'**函 数 名：StripTerFlag
'**输    入：(String)Str
'**输    出：(String)-
'**功能描述：获得结束标识符
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-05-26 22:25:52
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function StripTerFlag(Str As String) As String
Dim n As Integer
n = InStr(Str, Chr$(0))
If n > 0 Then
    StripTerFlag = Left(Str, n - 1)
Else
    StripTerFlag = Str
End If
End Function

'*************************************************************************
'**函 数 名：AutoSwitchLine
'**输    入：(Long)RichText , (Boolean)bSwitch
'**输    出：-
'**功能描述：设置 RichTextBox 自动换行
'**全局变量：
'**调用模块：
'**作    者：TechnoFantasy
'**日    期：2005-02-05 23:25:19
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub AutoSwitchLine(ByRef RichText As RichTextBox, ByVal bSwitch As Boolean)
    If bSwitch Then
        '设置 RichTextBox 自动换行
        Call SendMessage(RichText.hWnd, EM_SETTARGETDEVICE, _
        GetDC(RichText.hWnd), RichText.Width / 15)
        If RichText.RightMargin = 0 Then
            RichText.RightMargin = 1
        Else
            RichText.RightMargin = 0
        End If
    Else
        '设置 RichTextBox 不自动换行
        Call SendMessage(RichText.hWnd, EM_SETTARGETDEVICE, 0, 1)
    End If
End Sub
