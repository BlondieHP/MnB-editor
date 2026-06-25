Attribute VB_Name = "MdlLanMgr"

'*************************************************************************
'**模 块 名：MdlLanMgr
'**说    明：DPS4E.com 版权所有2008 - 2009(C)1
'**创 建 人：kevin
'**日    期：2008-05-17 01:11:16
'**修 改 人：SSgt_Edward,Ser_Charles
'**日    期：
'**描    述：
'**版    本：V0.951.7
'*************************************************************************
    Option Explicit
    
    Private Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)

    Private LanguageFileName As String
    
    Public PublicTags(36) As String   'ends_add
    Public PublicEditors(13) As String
    Public PublicEditors_Simplified(13) As String
    Public PublicTools(5) As String
    Public PublicHelp(1) As String
    Public PublicTips(5) As String
    Public PublicSkills() As String
    Public PublicMsgs(167) As String
    Public PublicWizards(31) As String
    Public PublicAbout(37) As String
    
    
'*************************************************************************
'**函 数 名：SelectLanguage
'**输    入：LanName(String) -
'**输    出：无
'**功能描述：选择语言
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 01:10:59
'**修 改 人：SSgt_Edward
'**日    期：2011-02-05 22:35:01
'**版    本：V1.1321
'*************************************************************************
Public Sub SelectLanguage(LanName As String)
    On Error Resume Next
    
LanguageFileName = ""

  With MnBInfo
    ' check empty
    If Len(LanName) = 0 Then
        GoTo Default
    End If
        
    .Language_Edit = LanName
    
    If .Language_Edit <> "cns" Then

      If FileExists(App.Path & "\" & .Language_Edit & ".lan.ini") Then
         LanguageFileName = App.Path & "\" & .Language_Edit & ".lan.ini"
         'TranslateForms
      Else
         GoTo Default
      End If
    Else
      InitPublicWords
    End If
   End With
   
   TranslatePublicWords
   
   Exit Sub
   
Default:
MnBInfo.Language_Edit = "cns"
InitPublicWords
End Sub

'*************************************************************************
'**函 数 名：GetLanguageFileName
'**输    入：无
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:58:03
'**修 改 人：
'**日    期：
'**版    本：V0.951.13
'*************************************************************************
Public Function GetLanguageFileName(LanName As String) As String

    If Len(LanguageFileName) = 0 Then
        SelectLanguage (LanName)
        GetLanguageFileName = LanguageFileName
    Else
        GetLanguageFileName = LanguageFileName
    End If

End Function

Public Function SetLanguage(LanName As String) As String

        SelectLanguage (LanName)
        SetLanguage = LanguageFileName

End Function

'*************************************************************************
'**函 数 名：TranslateStr
'**输    入：sSection(String) -
'**        ：sKey(String)     -
'**        ：sDefVal(String)  -
'**输    出：(String) -
'**功能描述：翻译文字
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 01:10:40
'**修 改 人：SSgt_Edward
'**日    期：2010-02-10 18:13:43
'**版    本：V1.1321
'*************************************************************************
Public Function TranslateStr(sSection As String, sKey As String, sDefVal As String) As String
    Dim sValue As String * 256
    Dim n As Long

    On Error Resume Next

    TranslateStr = sDefVal
    If Not FileExists(LanguageFileName) Then
        Exit Function
    End If
    n = GetPrivateProfileString(sSection, sKey, sDefVal, sValue, 255, LanguageFileName)
    n = InStr(1, sValue, Chr(0), vbTextCompare)
    
    If n > 0 Then
        TranslateStr = Left$(sValue, n - 1)
    End If
End Function

'*************************************************************************
'**函 数 名：TranslateForm
'**输    入：Frm(Form) -
'**输    出：无
'**功能描述：自动转换Form上的一些基本控件（也可自行扩展，目前支持CommandButton,Label,OptionButton,CheckButton）
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 01:09:23
'**修 改 人：SSgt_Edward
'**日    期：2011-02-05 22:49:12
'**版    本：1.1321
'*************************************************************************
Public Sub TranslateForm(frm As Form)
    Dim i As Long, j As Integer

    On Error Resume Next
If LanguageFileName <> "" Then
If FileExists(LanguageFileName) Then
    frm.Caption = TranslateStr(frm.Name, "Caption", frm.Caption)
    
    For i = 0 To frm.Controls.Count - 1
        If (TypeOf frm.Controls(i) Is CommandButton) Or (TypeOf frm.Controls(i) Is Label) Or (TypeOf frm.Controls(i) Is Menu) _
           Or (TypeOf frm.Controls(i) Is OptionButton) Or (TypeOf frm.Controls(i) Is CheckBox) Or (TypeOf frm.Controls(i) Is Frame) Then
            'frm.Controls(i).Caption = TranslateStr(frm.Name, frm.Controls(i).Name, frm.Controls(i).Caption)
            If checkIndex(frm, i) >= 0 Then
                frm.Controls(i).Caption = TranslateStr(frm.Name & "_" & frm.Controls(i).Name, checkIndex(frm, i), frm.Controls(i).Caption)
            Else
                frm.Controls(i).Caption = TranslateStr(frm.Name, frm.Controls(i).Name, frm.Controls(i).Caption)
            End If
        ElseIf (TypeOf frm.Controls(i) Is TabStrip) Then
          For j = 1 To frm.Controls(i).Tabs.Count
            If checkIndex(frm, i) >= 0 Then
                frm.Controls(i).Tabs(j).Caption = TranslateStr(frm.Name & "_" & frm.Controls(i).Name & "_" & checkIndex(frm, i), CStr(j), frm.Controls(i).Tabs(j).Caption)
            Else
                frm.Controls(i).Tabs(j).Caption = TranslateStr(frm.Name & "_" & frm.Controls(i).Name, CStr(j), frm.Controls(i).Tabs(j).Caption)
            End If
          Next j
        End If
    Next i
    
    If LCase(frm.Name) = "frmitems" Then
        For i = 0 To frm.LstFlags.ListCount - 1
            frm.LstFlags.List(i) = TranslateStr(frm.Name & "_LstFlags", CStr(i), frm.LstFlags.List(i))
        Next i
        
        For i = 0 To frm.LstAction.ListCount - 1
            frm.LstAction.List(i) = TranslateStr(frm.Name & "_LstAction", CStr(i), frm.LstAction.List(i))
        Next i
        
        For i = 0 To frm.CBixmesh.ListCount - 1
            frm.CBixmesh.List(i) = TranslateStr(frm.Name & "_CBixmesh", CStr(i), frm.CBixmesh.List(i))
        Next i
    ElseIf LCase(frm.Name) = "frmpsys" Then
    
        For i = 0 To frm.LstFlags.ListCount - 1
            frm.LstFlags.List(i) = TranslateStr(frm.Name & "_LstFlags", CStr(i), frm.LstFlags.List(i))
        Next i
    ElseIf LCase(frm.Name) = "frmmap" Then
    
        For i = 0 To frm.LstFilter(0).ListCount - 1
            frm.LstFilter(0).List(i) = TranslateStr(frm.Name & "_LstFilter", CStr(i), frm.LstFilter(0).List(i))
        Next i
    End If
End If
End If
End Sub



'*************************************************************************
'**函 数 名：InitTranslateForm
'**输    入：Frm(Form)           -
'**        ：iniFileName(String) -
'**输    出：无
'**功能描述：把Form上的一些基本控件的Caption记录到ini文件
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 01:09:11
'**修 改 人：SSgt_Edward
'**日    期：2011-02-05 22:49:12
'**版    本：1.1321
'*************************************************************************
Public Sub InitTranslateForm(frm As Form, iniFileName As String)
    Dim i As Long, j As Integer
    On Error Resume Next

    'Frm.Caption = TranslateStr(Frm.Name, "Caption", Frm.Caption)
    Call WriteString(iniFileName, frm.Name, "Caption", frm.Caption)
    
    For i = 0 To frm.Controls.Count - 1
        If (TypeOf frm.Controls(i) Is CommandButton) Or (TypeOf frm.Controls(i) Is Label) Or (TypeOf frm.Controls(i) Is Menu) _
           Or (TypeOf frm.Controls(i) Is OptionButton) Or (TypeOf frm.Controls(i) Is CheckBox) Or (TypeOf frm.Controls(i) Is Frame) Then
            'Frm.Controls(I).Caption = TranslateStr(Frm.Name, Frm.Controls(I).Name, Frm.Controls(I).Caption)
            If checkIndex(frm, i) >= 0 Then
                Call WriteString(iniFileName, frm.Name & "_" & frm.Controls(i).Name, checkIndex(frm, i), frm.Controls(i).Caption)
            Else
                Call WriteString(iniFileName, frm.Name, frm.Controls(i).Name, frm.Controls(i).Caption)
            End If
        ElseIf (TypeOf frm.Controls(i) Is TabStrip) Then
            For j = 1 To frm.Controls(i).Tabs.Count
               If checkIndex(frm, i) >= 0 Then
                Call WriteString(iniFileName, frm.Name & "_" & frm.Controls(i).Name & "_" & checkIndex(frm, i), CStr(j), frm.Controls(i).Tabs(j).Caption)
               Else
                Call WriteString(iniFileName, frm.Name & "_" & frm.Controls(i).Name, CStr(j), frm.Controls(i).Tabs(j).Caption)
               End If
            Next j
        End If
    Next i
    
    If LCase(frm.Name) = "frmitems" Then
        For i = 0 To frm.LstFlags.ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_LstFlags", CStr(i), frm.LstFlags.List(i))
        Next i
        
        For i = 0 To frm.LstAction.ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_LstAction", CStr(i), frm.LstAction.List(i))
        Next i
        
        For i = 0 To frm.CBixmesh.ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_CBixmesh", CStr(i), frm.CBixmesh.List(i))
        Next i
    ElseIf LCase(frm.Name) = "frmpsys" Then
    
        For i = 0 To frm.LstFlags.ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_LstFlags", CStr(i), frm.LstFlags.List(i))
        Next i
    ElseIf LCase(frm.Name) = "frmmap" Then
        For i = 0 To frm.LstFilter(0).ListCount - 1
            Call WriteString(iniFileName, frm.Name & "_LstFilter", CStr(i), frm.LstFilter(0).List(i))
        Next i
        
    End If
End Sub

'*************************************************************************
'**函 数 名：TranslatePublicWords
'**输    入：iniFileName(String)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2011-02-08 15:06:35
'**版    本：1.1321
'*************************************************************************
Public Sub TranslatePublicWords()
Dim i As Integer, j As Integer

If Not FileExists(LanguageFileName) Then
   Exit Sub
End If
 
For i = 0 To UBound(PublicTags)
    PublicTags(i) = TranslateStr("PublicTags", CStr(i), PublicTags(i))
Next i

For i = 0 To UBound(PublicEditors)
    PublicEditors(i) = TranslateStr("PublicEditors", CStr(i), PublicEditors(i))
Next i

For i = 0 To UBound(PublicEditors_Simplified)
    PublicEditors_Simplified(i) = TranslateStr("PublicEditors_Simplified", CStr(i), PublicEditors_Simplified(i))
Next i

For i = 0 To UBound(PublicTools)
    PublicTools(i) = TranslateStr("PublicTools", CStr(i), PublicTools(i))
Next i

For i = 0 To UBound(PublicHelp)
    PublicHelp(i) = TranslateStr("PublicHelp", CStr(i), PublicHelp(i))
Next i

For i = 0 To UBound(PublicHelp)
    PublicHelp(i) = TranslateStr("PublicHelp", CStr(i), PublicHelp(i))
Next i

For i = 0 To UBound(PublicTips)
    PublicTips(i) = TranslateStr("PublicTips", CStr(i), PublicTips(i))
Next i

For i = 0 To UBound(PublicMsgs)
    PublicMsgs(i) = TranslateStr("PublicMsgs", CStr(i), PublicMsgs(i))
Next i

For i = 0 To UBound(PublicSkills)
    PublicSkills(i) = TranslateStr("PublicSkills", CStr(i), PublicSkills(i))
Next i

For i = 0 To UBound(Operation)
   With Operation(i)
     .Op_CSVname = TranslateStr("Operation" & i, "Name", .Op_CSVname)
    
     For j = 1 To .ParaNum
         .Para(j).Value = TranslateStr("Operation" & i, CStr(j), .Para(j).Value)
     Next j
   End With
Next i

   For i = 0 To UBound(tiOn)
     tiOn(i).Y = TranslateStr("tiOn", CStr(i), tiOn(i).Y)
   Next i

For i = 0 To UBound(PublicWizards)
    PublicWizards(i) = TranslateStr("PublicWizards", CStr(i), PublicWizards(i))
Next i

For i = 0 To UBound(PublicAbout)
    PublicAbout(i) = TranslateStr("PublicAbout", CStr(i), PublicAbout(i))
Next i

End Sub

'*************************************************************************
'**函 数 名：WritePublicWords
'**输    入：
'**输    出：无
'**功能描述：把PublicWords记录到ini文件
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-02-08 15:06:35
'**修 改 人：
'**日    期：
'**版    本：1.1321
'*************************************************************************
Public Sub WritePublicWords()
Dim i As Integer, iniFileName As String, j As Integer

iniFileName = App.Path & "\new.lan.ini"
 
For i = 0 To UBound(PublicTags)
    Call WriteString(iniFileName, "PublicTags", CStr(i), PublicTags(i))
Next i

For i = 0 To UBound(PublicEditors)
    Call WriteString(iniFileName, "PublicEditors", CStr(i), PublicEditors(i))
Next i

For i = 0 To UBound(PublicEditors_Simplified)
    Call WriteString(iniFileName, "PublicEditors_Simplified", CStr(i), PublicEditors_Simplified(i))
Next i

For i = 0 To UBound(PublicTools)
    Call WriteString(iniFileName, "PublicTools", CStr(i), PublicTools(i))
Next i

For i = 0 To UBound(PublicHelp)
    Call WriteString(iniFileName, "PublicHelp", CStr(i), PublicHelp(i))
Next i

For i = 0 To UBound(PublicTips)
    Call WriteString(iniFileName, "PublicTips", CStr(i), PublicTips(i))
Next i

For i = 0 To UBound(PublicMsgs)
    Call WriteString(iniFileName, "PublicMsgs", CStr(i), PublicMsgs(i))
Next i

For i = 0 To UBound(PublicSkills)
    Call WriteString(iniFileName, "PublicSkills", CStr(i), PublicSkills(i))
Next i

For i = 0 To UBound(Operation)
   With Operation(i)
     Call WriteString(iniFileName, "Operation" & i, "Name", .Op_CSVname)
     
     For j = 1 To .ParaNum
         Call WriteString(iniFileName, "Operation" & i, CStr(j), .Para(j).Value)
     Next j
   End With
Next i

For i = 0 To UBound(tiOn)
     Call WriteString(iniFileName, "tiOn", CStr(i), tiOn(i).Y)
Next i

For i = 0 To UBound(PublicWizards)
    Call WriteString(iniFileName, "PublicWizards", CStr(i), PublicWizards(i))
Next i

For i = 0 To UBound(PublicAbout)
    Call WriteString(iniFileName, "PublicAbout", CStr(i), PublicAbout(i))
Next i

End Sub

'*************************************************************************
'**函 数 名：AddSplash
'**输    入：path(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 01:09:07
'**修 改 人：
'**日    期：
'**版    本：V0.951.7
'*************************************************************************
Public Function AddSplash(Path As String)
    On Error Resume Next
    '------------------------------------------------
    'AddSplash(App.path)
    
    AddSplash = App.Path & "\"

End Function


'*************************************************************************
'**函 数 名：FileExists
'**输    入：file(String) -
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 01:08:50
'**修 改 人：
'**日    期：
'**版    本：V0.951.7
'*************************************************************************
Public Function FileExists(file As String) As Boolean
    On Error Resume Next
    If Trim(file) = "" Then
        FileExists = False
        Exit Function
    End If
    If Dir(file, vbNormal + vbReadOnly + vbHidden + vbSystem + vbArchive) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

'*************************************************************************
'**函 数 名：DirExists
'**输    入：file(String) -
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-27 23:17:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function DirExists(Path As String) As Boolean
Dim s As String
    On Error Resume Next
    DirExists = True
    If Trim(Path) = "" Then
        DirExists = False
        Exit Function
    End If
    
    s = Dir(Path, vbNormal + vbHidden + vbSystem + vbDirectory)
    s = Trim(s)
    If s = "" Then
       DirExists = False
       Exit Function
    End If
    
    If Right(s, 1) = "." Then
       DirExists = False
       Exit Function
    End If
    
End Function


'*************************************************************************
'**函 数 名：checkIndex
'**输    入：frm(Form) -
'**        ：i(Long)   -
'**输    出：(Integer) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 01:39:57
'**修 改 人：
'**日    期：
'**版    本：V0.951.9
'*************************************************************************
Private Function checkIndex(frm As Form, i As Long) As Integer
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    checkIndex = CInt(frm.Controls(i).Index)
    

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    checkIndex = -1
End Function



'*************************************************************************
'**函 数 名：showLanMsg
'**输    入：msgIndex(Integer) -
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 08:06:23
'**修 改 人：
'**日    期：
'**版    本：V0.951.9
'*************************************************************************
Public Function showLanMsg(msgIndex As Integer) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    'showLanMsg = Form1.msgConf(msgIndex)
    'showLanMsg = gMsgConf(msgIndex)
    
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    showLanMsg = "#ERROR# NO_MESSAGE! [" & msgIndex & "]"
    Call logErr("MdlLanMgr", "showLanMsg", Err.Number, Err.Description)
End Function

Public Function InitTranslateForms()

'InitTranslateForm frmAbout, App.Path & "\new.lan.ini"
InitTranslateForm frmBackUpManager, App.Path & "\new.lan.ini"
InitTranslateForm frmFactions, App.Path & "\new.lan.ini"
InitTranslateForm frmItems, App.Path & "\new.lan.ini"
InitTranslateForm frmLine, App.Path & "\new.lan.ini"
InitTranslateForm frmMain, App.Path & "\new.lan.ini"
InitTranslateForm frmMap, App.Path & "\new.lan.ini"
InitTranslateForm frmMap_Icons, App.Path & "\new.lan.ini"
InitTranslateForm frmMesh, App.Path & "\new.lan.ini"
InitTranslateForm frmParties, App.Path & "\new.lan.ini"
InitTranslateForm frmParty_Templates, App.Path & "\new.lan.ini"
InitTranslateForm frmPSys, App.Path & "\new.lan.ini"
InitTranslateForm frmSaveAs, App.Path & "\new.lan.ini"
InitTranslateForm frmScenes, App.Path & "\new.lan.ini"
InitTranslateForm frmSelectPath, App.Path & "\new.lan.ini"
InitTranslateForm frmSoundRess, App.Path & "\new.lan.ini"
InitTranslateForm frmSounds, App.Path & "\new.lan.ini"
InitTranslateForm frmTabMat, App.Path & "\new.lan.ini"
InitTranslateForm frmTip, App.Path & "\new.lan.ini"
InitTranslateForm frmTroops, App.Path & "\new.lan.ini"
InitTranslateForm frmTrigger, App.Path & "\new.lan.ini"
InitTranslateForm Welcome, App.Path & "\new.lan.ini"

End Function


'*************************************************************************
'**函 数 名：InitPublicWords
'**输    入：无
'**输    出：无
'**功能描述：初始化PublicTags
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-07 00:22:45
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub InitPublicWords()
Dim n As Integer, Discrabe() As Variant

'ends-add
PublicTags(0) = "保留"
PublicTags(Tag_Register) = "寄存器"
PublicTags(Tag_Variable) = "全局变量"
PublicTags(Tag_String) = "字符串"
PublicTags(Tag_Item) = "物品"
PublicTags(Tag_Troop) = "兵种"
PublicTags(Tag_Faction) = "阵营"
PublicTags(Tag_Quest) = "任务"
PublicTags(Tag_Party_Tpl) = "部队模板"
PublicTags(Tag_Party) = "部队"
PublicTags(Tag_Scene) = "场景"
PublicTags(Tag_Mission_tpl) = "任务模板"
PublicTags(Tag_Menu) = "菜单"
PublicTags(Tag_Script) = "脚本"
PublicTags(Tag_Particle_Sys) = "粒子系统"
PublicTags(Tag_Scene_Prop) = "场景道具"
PublicTags(Tag_Sound) = "声音"
PublicTags(Tag_Local_Variable) = "局部变量"
PublicTags(Tag_Map_Icon) = "大地图图标"
PublicTags(Tag_Skill) = "技能"
PublicTags(Tag_Mesh) = "网格模型"
PublicTags(Tag_Presentation) = "展示"
PublicTags(Tag_Quick_String) = "快速字符串"
PublicTags(Tag_Track) = "曲目"
PublicTags(Tag_Tableau) = "可变材质"
PublicTags(Tag_Animation) = "动画"
PublicTags(Tags_End) = "其他"
PublicTags(Tags_End + 1) = "位置"
PublicTags(Tags_End + 2) = "物品类型"
PublicTags(Tags_End + 3) = "兵种标签"
PublicTags(Tags_End + 4) = "设置"
PublicTags(Tags_End + 5) = "部队AI行为"
PublicTags(Tags_End + 6) = "播放选项"
PublicTags(Tags_End + 7) = "部队标签"
PublicTags(Tags_End + 8) = "访问权限"
PublicTags(Tags_End + 9) = "返回值设定"
PublicTags(Tags_End + 10) = "s"

PublicEditors(0) = "编辑器"
PublicEditors(1) = "兵种"
PublicEditors(2) = "物品"
PublicEditors(3) = "部队"
PublicEditors(4) = "部队模板"
PublicEditors(5) = "阵营"
PublicEditors(6) = "场景"
PublicEditors(7) = "大地图图标"
PublicEditors(8) = "粒子系统"
PublicEditors(9) = "声音"
PublicEditors(10) = "声音资源"
PublicEditors(11) = "可变素材"
PublicEditors(12) = "网格模型"
PublicEditors(13) = "触发器"

PublicEditors_Simplified(0) = "编辑器"
PublicEditors_Simplified(1) = "兵种"
PublicEditors_Simplified(2) = "物品"
PublicEditors_Simplified(3) = "部队"
PublicEditors_Simplified(4) = "模板"
PublicEditors_Simplified(5) = "阵营"
PublicEditors_Simplified(6) = "场景"
PublicEditors_Simplified(7) = "图标"
PublicEditors_Simplified(8) = "粒子系统"
PublicEditors_Simplified(9) = "声音"
PublicEditors_Simplified(10) = "资源"
PublicEditors_Simplified(11) = "可变素材"
PublicEditors_Simplified(12) = "网格模型"
PublicEditors_Simplified(13) = "触发器"
        
PublicTools(0) = "工具"
PublicTools(1) = "卡拉迪亚地图"
PublicTools(2) = "备份管理器"
PublicTools(3) = "物品导入向导"
PublicTools(4) = "操作快编写工具"
PublicTools(5) = "字符串管理工具"
        
PublicHelp(0) = "帮助"
PublicHelp(1) = "关于"

PublicTips(0) = "(无)"
PublicTips(1) = "载入[str0]中,请稍候..."
PublicTips(2) = "创建模组中,请稍候..."
PublicTips(3) = "复制资源中,请稍候..."
PublicTips(4) = "保存中,请稍候..."
PublicTips(5) = "导入中,请稍候..."

PublicMsgs(0) = "提示"
PublicMsgs(1) = "确定套用当前设置?"
PublicMsgs(2) = "是否删除[[str0]]?"
PublicMsgs(3) = "删除[str0]"
PublicMsgs(4) = "[[str0]]是必要成分,无法删除!"
PublicMsgs(5) = "是否参照[[str0]]来创建新[str1]?"
PublicMsgs(6) = "关系"
PublicMsgs(7) = "[[str0]]是必要成分,无法移动!"
PublicMsgs(8) = "移动[str0]"
PublicMsgs(9) = "创建新[str0]"
PublicMsgs(10) = "确定重置当前设置?"
PublicMsgs(11) = "没找到更多匹配项!"
PublicMsgs(12) = "查询"
PublicMsgs(13) = "序列"
PublicMsgs(14) = "名"
PublicMsgs(15) = "确定导出窗体语言?"
PublicMsgs(16) = "窗体语言已成功导出到new.lan.ini中!点击确定后程序将重启。"
PublicMsgs(17) = "确定要创建一个还原点吗?"
PublicMsgs(18) = "[str0]最多只能有[str1]!"
PublicMsgs(19) = "错误!"
PublicMsgs(20) = "创建还原点失败!请确保备份剧本的完整性!"
PublicMsgs(21) = "创建成功!备份文件夹在[[str0]]"
PublicMsgs(22) = "确定退出 【骑马与砍杀:战团】剧本编辑器 吗?"
PublicMsgs(23) = "确认退出"
PublicMsgs(24) = "确定按现在设定保存剧本[[str0]]?"
PublicMsgs(25) = "保存剧本"
PublicMsgs(26) = "点数"
PublicMsgs(27) = "请确认备份剧本[[str0]]后,按“确定”开始保存。"
PublicMsgs(28) = "上一武器的第二种攻击模式"
PublicMsgs(29) = "剧本[[str0]]保存成功!"
PublicMsgs(30) = "输出窗口:[[str0]]"
PublicMsgs(31) = "技能"
PublicMsgs(32) = "选择剧本"
PublicMsgs(33) = "确定要重新选择剧本?"
PublicMsgs(34) = "删除还原点"
PublicMsgs(35) = "确定要删除所选还原点吗?"
PublicMsgs(36) = "创建一个还原点"
PublicMsgs(37) = "确定要创建一个以时间标注的还原点吗?"
PublicMsgs(38) = "参数"
PublicMsgs(39) = "参数必须是整型!"
PublicMsgs(40) = "还原"
PublicMsgs(41) = "确定要还原[[str0]]吗?"
PublicMsgs(42) = "保留"
PublicMsgs(43) = "还原成功!"
PublicMsgs(44) = "还原失败!"
PublicMsgs(45) = "超出兵种拥有物品限制(64)!添加失败!"
PublicMsgs(46) = "数量"
PublicMsgs(47) = "俘虏"
PublicMsgs(48) = "最少"
PublicMsgs(49) = "最多"
PublicMsgs(50) = "剧本:[str0]"
PublicMsgs(51) = "另存为"
PublicMsgs(52) = "特征操作块"
PublicMsgs(53) = "所输入的剧本名不合法!"
PublicMsgs(54) = "无法识别"
PublicMsgs(55) = "坐标"
PublicMsgs(56) = "不能挥砍"
PublicMsgs(57) = "不能穿刺"
PublicMsgs(58) = "请先复制操作!"
PublicMsgs(59) = "请选择要粘贴的操作所在的触发器!"
PublicMsgs(60) = "至少要有一个模型!"
PublicMsgs(61) = "此物品不能编辑"
PublicMsgs(62) = "请选择要创建参数所在的操作!"
PublicMsgs(63) = "这个操作不能新建参数!"
PublicMsgs(64) = "物品所属阵营不存在,将设为无阵营"
PublicMsgs(65) = "触发器"
PublicMsgs(66) = "条件"
PublicMsgs(67) = "执行"
PublicMsgs(68) = "操作"
PublicMsgs(69) = "[str0][str1][str2][str3]中[str4]不存在,将被设为[str4](0)!点击套用按钮后将套用该修正"
PublicMsgs(70) = "操作只能选择!"
PublicMsgs(71) = "不存在该条件!"
PublicMsgs(72) = "请选择要添加的操作所在的触发器!"
PublicMsgs(73) = "参数不合法!"
PublicMsgs(74) = "此项不能移动!"
PublicMsgs(75) = "请选择要移动的操作!"
PublicMsgs(76) = "物品触发器已复制"
PublicMsgs(77) = "操作已复制"
PublicMsgs(78) = "此项不能复制!"
PublicMsgs(79) = "请选择要复制的触发器或操作!"
PublicMsgs(80) = "确定要删除该触发器么?"
PublicMsgs(81) = "确定要删除该操作么?"
PublicMsgs(82) = "至少要有一个操作!"
PublicMsgs(83) = "必选参数不能删除!"
PublicMsgs(84) = "可选参数只能从最后一个删除!"
PublicMsgs(85) = "此项不能删除!"
PublicMsgs(86) = "未选中任何项目!"
PublicMsgs(87) = "请先复制物品触发器!"
PublicMsgs(88) = "值"
PublicMsgs(89) = "驱动器错误"
PublicMsgs(90) = "[[str0]]已存在,创建失败!"
PublicMsgs(91) = "[[str0]]已存在,请更改ID后再套用。"
PublicMsgs(92) = "任意参数"
PublicMsgs(93) = "含有条件操作[str0]个"
PublicMsgs(94) = "含有结果操作[str0]个"
PublicMsgs(95) = "请选择要复制的操作!"
PublicMsgs(96) = "确定要向注册表填入当前信息吗?"
PublicMsgs(97) = "确认注册"
PublicMsgs(98) = "已成功向注册表填入当前信息!"
PublicMsgs(99) = "是否按当前文本导入?"
PublicMsgs(100) = "导入"
PublicMsgs(101) = "导入完成!"
PublicMsgs(102) = "类 型"
PublicMsgs(103) = "价 格"
PublicMsgs(104) = "单/双手武器"
PublicMsgs(105) = "重 量"
PublicMsgs(106) = "挥 砍"
PublicMsgs(107) = "穿 刺"
PublicMsgs(108) = "砍"
PublicMsgs(109) = "刺"
PublicMsgs(110) = "钝"
PublicMsgs(111) = "速 度"
PublicMsgs(112) = "范 围"
PublicMsgs(113) = "难 度"
PublicMsgs(114) = "特 性"
PublicMsgs(115) = "精 度"
PublicMsgs(116) = "弹 速"
PublicMsgs(117) = "可以穿盾"
PublicMsgs(118) = "防 护"
PublicMsgs(119) = "头 防"
PublicMsgs(120) = "操 纵"
PublicMsgs(121) = "冲 锋"
PublicMsgs(122) = "生 命"
PublicMsgs(123) = "身 防"
PublicMsgs(124) = "腿 防"
PublicMsgs(125) = "抗 击"
PublicMsgs(126) = "尺 寸"
PublicMsgs(127) = "强 度"
PublicMsgs(128) = "质 量"
PublicMsgs(129) = "数 量"
PublicMsgs(130) = "信息已复制到剪切板"
PublicMsgs(131) = "等 级"
PublicMsgs(132) = "技 能"
PublicMsgs(133) = "内容"
PublicMsgs(134) = "已将修改内容套用"
PublicMsgs(135) = "腿 防"        '保留
PublicMsgs(136) = "抗 击"        '保留
PublicMsgs(137) = "尺 寸"        '保留
PublicMsgs(138) = "强 度"        '保留
PublicMsgs(139) = "质 量"        '保留
PublicMsgs(140) = "数 量"        '保留
PublicMsgs(141) = "物品信息已复制到剪切板"        '保留
PublicMsgs(142) = "确定导出注册操作?"
PublicMsgs(143) = "操作已成功导出到new.op.ini中!"
PublicMsgs(144) = "请重启程序使操作导入生效!"
PublicMsgs(145) = "确定要重命名备份[str0]?"
PublicMsgs(146) = "备份已成功更名为[str0]!"
PublicMsgs(147) = "前缀"
PublicMsgs(148) = "建立索引中..."
PublicMsgs(149) = "错误的操作号"
PublicMsgs(150) = "空字符串"
PublicMsgs(151) = "代码源不完整"
PublicMsgs(152) = "错误[str0]:[str1]" & vbCrLf & "第[str2]个操作的第[str3]参数出错。"
PublicMsgs(153) = "不是"
PublicMsgs(154) = "是或下项"
PublicMsgs(155) = "操作号:[[str0]]"
PublicMsgs(156) = "<参数[str0]>"
PublicMsgs(157) = "是"
PublicMsgs(158) = "<逻辑符>"
PublicMsgs(159) = "<更多参数...>"
PublicMsgs(160) = "变量"   '保留
PublicMsgs(161) = "不是或下项"
PublicMsgs(162) = "无法找到操作[[str0]]!"
PublicMsgs(163) = "名称"
PublicMsgs(164) = "位置"   '保留
PublicMsgs(165) = "物品类型"   '保留
PublicMsgs(166) = "输入"
PublicMsgs(167) = "您是想新建变量还是将当前变量更名呢？选择[是]将创建新变量,选择[否]将当前变量更名。"

Discrabe = Array("交易", "统御", "俘虏管理", "", "", _
               "", "", "说服力", "工程学", "急救", _
               "手术", "疗伤", "物品管理", "侦察", "向导", _
               "战术", "跟踪", "教练", "", "", _
               "", "", "掠夺", "骑射", "骑术", _
               "跑动", "盾防", "武器掌握", "", "", _
               "", "", "", "强弓", "强掷", _
               "强击", "铁骨", "", "", "", _
               "", "")
               
ReDim PublicSkills(UBound(Discrabe))

For n = 0 To UBound(Discrabe)
      PublicSkills(n) = Discrabe(n)
Next n

PublicWizards(0) = "下一步>(&N)"
PublicWizards(1) = "<上一步(&F)"
PublicWizards(2) = "取消(&C)"
PublicWizards(3) = "完成(&F)"
PublicWizards(4) = "导入成功"
PublicWizards(5) = "导入失败"
PublicWizards(6) = "错误报告"
PublicWizards(7) = "导入物品txt信息(item_kinds1.txt)"
PublicWizards(8) = "选择需导入装备的文件夹"
PublicWizards(9) = "导入物品语言信息(item_kinds.csv)"
PublicWizards(10) = "确定导入物品"
PublicWizards(11) = "导入物品报告(共[str0]项):"
PublicWizards(12) = "导入模型报告(共[str0]项):"
PublicWizards(13) = "导入贴图报告(共[str0]项):"
PublicWizards(14) = "[未找到]"
PublicWizards(15) = "贴图文件"
PublicWizards(16) = "模型文件"
PublicWizards(17) = "无任何txt读入!导入向导将不会产生相应物品,需要您自己在[物品编辑器]里手动添加。是否继续下一步?"
PublicWizards(18) = "无任何语言信息读入!是否继续下一步?"
PublicWizards(19) = "请将需导入的物品的txt行复制进下面的文本框,如果没有,请直接点击[下一步]。"

PublicWizards(20) = "请选择需要导入的装备的文件夹,"
PublicWizards(21) = "确保该文件夹下有Textures(贴图)和Resource(模型)文件夹,"
PublicWizards(22) = "并确认其中分别有DDS、BRF文件。"
PublicWizards(23) = "如果您的这两种文件不是这样存放的,"
PublicWizards(24) = "也请您按此存放以确保导入的正常进行。"
                                    
PublicWizards(25) = "请将需导入的物品的语言信息(可在下方文本框里设置)复制进下面的文本行,"
PublicWizards(26) = "如果没有,请直接点[下一步]。"

PublicWizards(27) = "导入的信息已设置,查看下列设置报告,如有误，点击[上一步]返回修改；确认正确,点击[下一步]启动导入程序。"

PublicWizards(28) = "物品已成功导入,"
PublicWizards(29) = "您可以在[物品编辑器]里随时修改它们。请记得在关闭编辑器前[保存剧本]以使导入生效。"
PublicWizards(30) = "但是由于没有txt行,向导无法创建相应物品。您可以在[物品编辑器]里创建新物品,并把该物品的模型设为刚导入的模型。"

PublicWizards(31) = "导入过程中断!请确认模型文件的正确性,txt文本的正确性,以及语言信息的正确性!点击[上一步]返回修改。"


PublicAbout(0) = "参与人员:"
PublicAbout(1) = "主程序:SSgt_Edward,Ser_Charles"
PublicAbout(2) = "美工:我不是个过客(hp_honey)"

PublicAbout(3) = "各个编辑器、工具负责人:"
PublicAbout(24) = "兵种:SSgt_Edward"
PublicAbout(25) = "物品:Ser_Charles"
PublicAbout(26) = "部队:SSgt_Edward"
PublicAbout(27) = "部队模板:SSgt_Edward"
PublicAbout(28) = "阵营:SSgt_Edward"
PublicAbout(29) = "场景:SSgt_Edward"
PublicAbout(30) = "大地图图标:SSgt_Edward"
PublicAbout(31) = "卡拉迪亚地图:SSgt_Edward"
PublicAbout(32) = "粒子系统:Ser_Charles"
PublicAbout(33) = "可变素材:Ser_Charles"
PublicAbout(34) = "声音:SSgt_Edward"
PublicAbout(35) = "声音资源:SSgt_Edward"
PublicAbout(36) = "备份管理器:SSgt_Edward"
PublicAbout(37) = "物品导入向导:SSgt_Edward"

PublicAbout(4) = "MnBWarband Editor介绍:"
PublicAbout(5) = "[MnBWarband Editor]编辑器是一款用于查看、修改[骑马与砍杀:战团]这款游戏的免费共享软件。"
PublicAbout(6) = "它是由骑砍中文站的[骑马砍杀TXT编辑器](kevin,kiss2003,mafei82394) 经过移植并完善而来(引用了7个模块),"
PublicAbout(7) = "在新增了[阵营]、[场景]、[部队]、[大地图图标]、[粒子系统]、[声音]、[声音资源]、[可变素材]编辑器以及"
PublicAbout(8) = "[卡拉迪亚地图]、[备份管理器]、[物品导入向导]工具的同时,"
PublicAbout(9) = "完善了原有的[兵种]、[物品]、[部队模板]编辑器,使用户能自由地修改自己的模组。"

PublicAbout(10) = "注意:"
PublicAbout(11) = "MnBWarband Editor在战团1.132环境下编写并经过测试。"

PublicAbout(12) = "声明:"
PublicAbout(13) = "在MnBWarband Editor版本稳定之后,会开放源代码(VB)。"

PublicAbout(14) = "特别感谢:"
PublicAbout(15) = "感谢骑砍中文站的kevin,kiss2003,mafei82394开放了[骑马砍杀TXT编辑器]的源代码"
PublicAbout(16) = "感谢骑砍中文站[骑马与砍杀MOD制作教程－中文版]的汉化组为我们带来的MOD教程"
PublicAbout(17) = "感谢百度骑马与砍杀吧和骑砍中文站的汽油们对MnBWarband Editor的关注与支持!"

PublicAbout(17) = "联系我们:"
PublicAbout(18) = "E-mail:ce1992620@yahoo.com.cn"
PublicAbout(19) = "       ce19920620@gmail.com"
PublicAbout(20) = "QQ:1061675615"
PublicAbout(21) = "   825312405"
PublicAbout(22) = "魔球讨论群:74195097"

PublicAbout(23) = "欢迎随时将您遇到的使用问题、Bug告诉我们,帮助我们完善MnBWarband Editor!"
End Sub
