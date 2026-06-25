Attribute VB_Name = "ModApp"

Option Explicit

Type kiss_type_troopsTree
    p As Integer    '父节点下标.
    n As Integer    '子节点个数.
    c(2) As Integer '子节点下标数组.
    ID As Integer   '子节点ID
    showed As Boolean
End Type

Public vTrpsTree() As kiss_type_troopsTree

Dim FormOldWidth As Long '保存窗体的原始宽度
Dim FormOldHeight As Long    '保存窗体的原始高度

'*************************************************************************
'**函 数 名：appExit
'**输    入：(Boolean)Restart
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-06-14 06:46:03
'**修 改 人：SSgt_Edward
'**日    期：2011-02-06 23:19:50
'**版    本：V1.1321
'*************************************************************************
Sub appExit(Optional Restart As Boolean = False)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    If Restart Then
       Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
    End If
        End
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModApp", "appExit [Restart=" & Restart & "]", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**函 数 名：logErr
'**输    入：ModName(String) -
'**        ：subName(String) -
'**        ：errNum(String)  -
'**        ：errMsg(String)  -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 00:10:31
'**修 改 人：
'**日    期：
'**版    本：V0.951.7
'*************************************************************************
'
'Call logErr("Form1", "cmdCopyForNewTroop_Click", Err.Number, Err.Description)
'
Sub logErr(ModName As String, subName As String, errNum As String, errMsg As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim strMsg As String

    strMsg = ModName & ":" & subName

    strMsg = strMsg & ", Err.Number=" & errNum
    
    strMsg = strMsg & " : " & "Error=" & errMsg

    
    'Form2.labDebugMsg.Caption = strMsg
    
    OutAsDebugTex (strMsg)
    
    SetMouseDefault
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Debug.Print "错误发生时间:"; Format(Now, "YYYY-MM-DD HH:MM:SS")
    Debug.Print "错误 的 类型:"; Err.Number
    Debug.Print "错误 的 信息:"; Err.Description
    Debug.Print "错误函数名称:logErr"
    Debug.Print "错误模块名称:Functions"
    SetMouseDefault
End Sub

'*************************************************************************
'**函 数 名：GetBackupFilename
'**输    入：FileName(String) -
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:24:09
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************
Public Function GetBackupFilename(FileName As String, Path As String) As String
    On Error Resume Next
    Dim Now As Date
    Dim today As Date
    Now = Time
    today = Date
    
    Dim strPath As String
    strPath = Path & "\backup"
    If Not FileExists(strPath) Then
        MkDir strPath
    End If
    
    GetBackupFilename = strPath & "\" & Format(today, "yyyy-mm-dd_") & Format(Now, "hh.mm.ss_") & FileName
        
End Function


'*************************************************************************
'**函 数 名：getRnd
'**输    入：lowerbound(Long) -
'**        ：upperbound(Long) -
'**输    出：(Long) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:24:21
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************
Function getRnd(lowerbound As Long, upperbound As Long) As Long
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Randomize
    
    getRnd = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("public", "getRnd", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：Round
'**输    入：nValue(Double)   -
'**        ：nDigits(Integer) -
'**输    出：(Double) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-06-14 06:45:00
'**修 改 人：
'**日    期：
'**版    本：V0.955.6
'*************************************************************************
Function Round(nValue As Double, nDigits As Integer) As Double
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("public", "Round", Err.Number, Err.Description)
End Function

Sub DebugText(strMsg As String)
    Call OutAsDebugTex(strMsg, "错误")
End Sub

'*************************************************************************
'**函 数 名：OutAsDebugTex
'**输    入：S(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:24:35
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************
Sub OutAsDebugTex(ByVal s As String, Optional Caption As String = "")
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    If Len(s) = 0 Then
        Exit Sub
    End If
    
    DebugForm.Caption = "输出窗口:[" & Caption & "]"
    DebugForm.Show
    
        DebugForm.Text1.Text = DebugForm.Text1.Text & vbCrLf & "---------" & CStr(Now) & "---------" & vbCrLf & s
        
        DebugForm.Text1.SelStart = Len(DebugForm.Text1.Text)
        
        DebugForm.ZOrder
    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("public", "OutAsDebugTex", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：showHelp
'**输    入：strType(String) -
'**        ：id(Integer)     -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-07-12 07:12:59
'**修 改 人：
'**日    期：
'**版    本：V0.960.28
'*************************************************************************
Public Sub showHelp(strType As String, ID As Integer)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    If Len(strType) = 0 Then
        Exit Sub
    End If

    SetMouseWait

    Dim strTmp As String

    Dim i As Integer
    Dim n As Integer
    Dim tmpMsg As String

    strTmp = ReadHelpString("", strType, CStr(ID) & "_count", 250)
    n = Val(strTmp)
    If n > 0 Then
    
        If Len(DebugForm.Text1.Text) > 0 Then
            DebugForm.Text1.Text = DebugForm.Text1.Text & vbCrLf & "---------" & CStr(Now) & vbCrLf
        End If
    
        For i = 0 To n
            strTmp = ReadHelpString("", strType, CStr(ID) & "_" & CStr(i), 250)
            DebugForm.Text1.Text = DebugForm.Text1.Text & vbCrLf & strTmp
        Next

        DebugForm.Caption = "HELP MSG"
        DebugForm.Show
    Else

        MsgBox "HELP:" & strType & "  " & CStr(ID) & " 没有定义!"

    End If

    SetMouseDefault
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("public", "showHelp", Err.Number, Err.Description)

End Sub

'*************************************************************************
'**函 数 名：ReadHelpString
'**输    入：iniFileName(String) -
'**        ：Section(String)     -
'**        ：Key(String)         -
'**        ：Size(Long)          -
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-07-12 07:23:53
'**修 改 人：
'**日    期：
'**版    本：V0.960.28
'*************************************************************************
Public Function ReadHelpString(iniFileName As String, Section As String, Key As String, Size As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim strFileName As String

    If Len(iniFileName) = 0 Then
        strFileName = GetLanguageFileName("")
    Else
        strFileName = iniFileName
    End If
    strFileName = Left$(strFileName, Len(strFileName) - 4) & "_help.ini"

    If Not FileExists(strFileName) Then
        MsgBox "Can not found HELP FILE [ " & iniFileName & " ] !", , "ERROR"
        ReadHelpString = ""

        'init new help.ini
        'write to file
        '初始化要创建的节:
        Dim vInitHelpIniLine(5) As String
        vInitHelpIniLine(0) = "[Help]"
        vInitHelpIniLine(1) = "[Form1_edittabsHelp]"
        vInitHelpIniLine(2) = "[Form2_Help]"
        vInitHelpIniLine(3) = "[Form3_edittabsHelp]"
        vInitHelpIniLine(4) = "[ModTesterForm_help]"
        vInitHelpIniLine(5) = "[PTForm_edittabsHelp]"
        
        '每一节初始化的数量.
        Dim vInitCount(5) As Integer
        vInitCount(0) = 0
        vInitCount(1) = 5
        vInitCount(2) = 0
        vInitCount(3) = 5
        vInitCount(4) = 5
        vInitCount(5) = 5

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        lngHandle = FreeFile()
        Open strFileName For Output As #lngHandle
        For i = 0 To 5
            Print #lngHandle, vInitHelpIniLine(i)
            For j = 0 To vInitCount(i)
                Print #lngHandle, CStr(j) & "_count=5"
                For k = 0 To 5
                    Print #lngHandle, CStr(j) & "_" & CStr(k) & "=" & vInitHelpIniLine(i) & " init help message string " & CStr(j) & "_" & CStr(k)
                Next
                Print #lngHandle, "------" & vbCrLf
            Next
        Next
        Close #lngHandle

        Exit Function

    End If

    ReadHelpString = ReadString(strFileName, Section, Key, Size)

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("public", "ReadHelpString", Err.Number, Err.Description)

End Function



'*************************************************************************
'**函 数 名：TrimPath
'**输    入：sPath(String) -
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 10:03:10
'**修 改 人：
'**日    期：
'**版    本：V0.951.13
'*************************************************************************
Public Function TrimPath(sPath As String) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    'remove path from path & filename
    'returns string AFTER last "\"
    'example:
    'nopath$ = TrimPath("C:\TXTFILES\JUSTFILE.TXT")
    'nopath$ will = "JUSTFILE.TXT"
    
    Dim i As Integer

    For i% = Len(sPath) To 1 Step -1
        If InStr(i%, sPath, "\", 1) = i% Then Exit For
    Next i%

    TrimPath = Right$(sPath, Len(sPath) - i%)

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("public", "TrimPath", Err.Number, Err.Description)
End Function



'*************************************************************************
'**函 数 名：setMouseWait
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-07-03 06:24:36
'**修 改 人：
'**日    期：
'**版    本：V0.960.17
'*************************************************************************
Public Sub SetMouseWait()
    Screen.MousePointer = 11 '是漏斗,vbHourglass)
End Sub

Public Sub SetMouseDefault()
    Screen.MousePointer = 0
End Sub

Public Sub SetMouse(intMode As Integer)
    Screen.MousePointer = intMode
End Sub

Public Sub ResizeInit(FormName As Form)
    Dim Obj As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight

    On Error Resume Next
    For Each Obj In FormName
        Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height & " "
    Next Obj
    On Error GoTo 0
End Sub

'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4) As Double
    Dim i As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
    ScaleX = FormName.ScaleWidth / FormOldWidth    '保存窗体宽度缩放比例
    ScaleY = FormName.ScaleHeight / FormOldHeight  '保存窗体高度缩放比例
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1

        For i = 0 To 4  '读取控件的原始位置与大小
            TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
            If TempPos > 0 Then
                Pos(i) = Mid$(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(i) = 0
            End If

            '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
            Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
        Next i

    Next Obj
    On Error GoTo 0
End Sub


'*************************************************************************
'**函 数 名：isTopNode
'**输    入：node(kiss_type_troopsTree) -
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-07-11 09:23:29
'**修 改 人：
'**日    期：
'**版    本：V0.960.25
'*************************************************************************
Function isTopNode(node As kiss_type_troopsTree) As Boolean
    isTopNode = False
    'If node Is Nothing Then Exit Function

    If node.p = 0 Then
        isTopNode = True
    End If
End Function

'*************************************************************************
'**函 数 名：isLeafNode
'**输    入：node(kiss_type_troopsTree) -
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-07-11 09:23:31
'**修 改 人：
'**日    期：
'**版    本：V0.960.25
'*************************************************************************
Function isLeafNode(node As kiss_type_troopsTree) As Boolean
    isLeafNode = False
    'If node Is Nothing Then Exit Function

    If node.n = 0 Then
        isLeafNode = True
    Else
        If node.c(0) = 0 And node.c(1) = 0 And node.c(2) = 0 Then
            node.n = 0
            isLeafNode = True
        End If
    End If
End Function

'*************************************************************************
'**函 数 名：fnNumberFixedLength
'**输    入：lngNumber(Long)    -原始数字
'**        ：lngLength(Long)    -长度限制
'**        ：strFixChar(String) -长度不足的时候前面补什么字符
'**输    出：(String) -
'**功能描述：
'**        ：eg.fnNumberFixedLength(10,3," ")  =" 10"
'**        ：eg.fnNumberFixedLength(10,3,"0")  ="010"
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2009-02-07 15:57:01
'**修 改 人：
'**日    期：
'**版    本：V1.11.10
'*************************************************************************
Function fnNumberFixedLength(lngNumber As Long, lngLength As Long, strFixChar As String) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim strStr As String
    Dim strTmp As String
    
    strStr = CStr(lngNumber)
    
    If Len(strStr) >= lngLength Then
        If Len(strFixChar) = 0 Then
            strTmp = strStr
        Else
            strTmp = Left(strStr, lngLength)
        End If

    Else
        If Len(strFixChar) = 0 Then
            'default fixChar
            strTmp = strStr
            While Len(strTmp) < lngLength
                strTmp = "0" & strTmp
            Wend
        Else
            strTmp = strStr
            While Len(strTmp) < lngLength
                strTmp = strFixChar & strTmp
            Wend
        End If

    End If
    
    fnNumberFixedLength = strTmp

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("public", "fnNumberFixedLength", Err.Number, Err.Description)
    Resume Next
End Function

'*************************************************************************
'**函 数 名：fnStrFixedLength
'**输    入：strStr(String)     -原始字符串
'**        ：lngLength(Long)    -长度限制
'**输    出：(String) -
'**功能描述：定长字符串
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2009-02-07 15:54:34
'**修 改 人：
'**日    期：
'**版    本：V1.11.10
'*************************************************************************
Function fnStrFixedLength(strStr As String, lngLength As Long, strFixChar As String) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim strTmp As String
    If Len(strStr) >= lngLength Then
        If Len(strFixChar) = 0 Then
            strTmp = strStr
        Else
            strTmp = Left(strStr, lngLength)
        End If

    Else
        If Len(strFixChar) = 0 Then
            'default fixChar
            strTmp = strStr
            While Len(strTmp) < lngLength
                strTmp = strTmp & " "
            Wend
        Else
            strTmp = strStr
            While Len(strTmp) < lngLength
                strTmp = strTmp & strFixChar
            Wend
        End If

    End If
    
    fnStrFixedLength = strTmp

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("public", "fnStrFixedLength", Err.Number, Err.Description)
    Resume Next
End Function
