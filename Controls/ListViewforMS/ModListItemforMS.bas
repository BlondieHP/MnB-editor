Attribute VB_Name = "ModListItemforMS"
Option Explicit

Public Const MENU_TEXT_ONLY = 0
Public Const MENU_LIST_ONLY = 1
Public Const MENU_TEXT_AND_LIST = 2

Public Const MENU_MSG_NULL = 0
Public Const MENU_MSG_ACTIVE = 1
Public Const MENU_MSG_DEACTIVE = 2

Public Const MENU_COLOR_DEFAULT = &H6AA40 '暗绿
Public Const MENU_COLOR_MAIN = vbBlue
Public Const MENU_COLOR_PARAM = &HA98C07   '青色
Public Const MENU_COLOR_CONTROL = vbRed           '&H80&         '暗红
Public Const MENU_COLOR_BUTTON = &H4000&        '深绿
Public Const MENU_COLOR_NEG = vbBlack             '&H404040           '灰色
Public Const MENU_COLOR_OPTIONAL = &HC000C0      '深紫

Public Const MENU_KEY_DEL = vbKeyDelete
Public Const MENU_KEY_CLEAR = vbKeyBack
Public Const MENU_KEY_CANCEL = vbKeyEscape
Public Const MENU_KEY_COPY = vbKeyC
Public Const MENU_KEY_PASTE = vbKeyV
Public Const MENU_KEY_OK = 13

Public Const BOARD_EMPTY = 0
Public Const BOARD_OP = 1
Public Const BOARD_PARAM = 2

Public Type Type_MS_Clipboard
  ContentType As Long
  Value As String
  TemType As String
End Type

Public MSBoard As Type_MS_Clipboard

Public Function TagPlus(Tag As String, Index As Integer, Value As Integer) As String
Dim strTem() As String, lngTem As Integer, i As Integer

If Tag = "" Then Exit Function
strTem = Split(Tag, ",")

lngTem = Val(strTem(Index)) + Value
strTem(Index) = CStr(lngTem)

For i = 0 To UBound(strTem)
  TagPlus = TagPlus & strTem(i) & ","
Next i

TagPlus = Left(TagPlus, Len(TagPlus) - 1)
End Function


Public Function GetTagValue(Tag As String, Index As Integer) As Integer
Dim strTem() As String, lngTem As Integer, i As Integer

If Tag = "" Then Exit Function
strTem = Split(Tag, ",")

GetTagValue = Val(strTem(Index))

End Function

Public Function TagAssign(Tag As String, Index As Integer, Value As Integer) As String
Dim strTem() As String, lngTem As Integer, i As Integer

If Tag = "" Then Exit Function
strTem = Split(Tag, ",")

lngTem = Value
strTem(Index) = CStr(lngTem)

For i = 0 To UBound(strTem)
  TagAssign = TagAssign & strTem(i) & ","
Next i

TagAssign = Left(TagAssign, Len(TagAssign) - 1)
End Function

Public Sub CopyListItemProperties(desItem As ListItemforMS, Item As ListItemforMS)
Dim i As Integer

With desItem

  .SubItems.Count = Item.SubItems.Count
  .SubItems.ForeColor = Item.SubItems.ForeColor
  .SubItems.FontBold = Item.SubItems.FontBold
  
  For i = 1 To .SubItems.Count
    .SubItems(i).KeyWord = Item.SubItems(i).KeyWord
    .SubItems(i).Locked = Item.SubItems(i).Locked
    .SubItems(i).Start = Item.SubItems(i).Start
    .SubItems(i).Text = Item.SubItems(i).Text
    .SubItems(i).Value = Item.SubItems(i).Value
    .SubItems(i).Negation = Item.SubItems(i).Negation
    .SubItems(i).ParaType = Item.SubItems(i).ParaType
    .SubItems(i).TemType = Item.SubItems(i).TemType
  Next i
  
  .Indention = Item.Indention
  .Locked = Item.Locked
  .Text = Item.Text
  .Value = Item.Value

End With

End Sub

Public Sub CopyListItem(desItem As ListItemforMS, Item As ListItemforMS, Optional ExchangeIndex As Boolean = False)
Dim i As Integer

With desItem
    If ExchangeIndex Then .Index = Item.Index
    If Item.Initialized Then
      .Initialize .Index, Item.ListView, Item.Label
    Else
     .Initialized = False
    End If
  .SubItems.Initialize Item.ListView, desItem
  .SubItems.SetCount Item.SubItems.Count
  .SubItems.ForeColor = Item.SubItems.ForeColor
  .SubItems.FontBold = Item.SubItems.FontBold
  
  For i = 1 To .SubItems.Count
      .SubItems(i).Index = Item.SubItems(i).Index
      If Item.SubItems(i).Initialized Then
        .SubItems(i).Initialize .Index, Item.SubItems(i).Index, Item.SubItems(i).ListView, Item.SubItems(i).Label, .SubItems
      Else
        .SubItems(i).Initialized = False
      End If

    .SubItems(i).KeyWord = Item.SubItems(i).KeyWord
    .SubItems(i).Locked = Item.SubItems(i).Locked
    .SubItems(i).Start = Item.SubItems(i).Start
    .SubItems(i).Text = Item.SubItems(i).Text
    .SubItems(i).Value = Item.SubItems(i).Value
    .SubItems(i).Negation = Item.SubItems(i).Negation
    .SubItems(i).ParaType = Item.SubItems(i).ParaType
    .SubItems(i).TemType = Item.SubItems(i).TemType
  Next i
  
  .Indention = Item.Indention
  .Locked = Item.Locked
  .Text = Item.Text
  .Value = Item.Value

End With

End Sub

Public Sub CopySubItem(desItem As ListSubItemforMS, Item As ListSubItemforMS, lIndex As Integer, Optional ExchangeIndex As Boolean = False)
Dim i As Integer

With desItem

      If ExchangeIndex Then .Index = Item.Index
      If Item.Initialized Then
        .Initialize lIndex, .Index, Item.ListView, Item.Label, .Parent
      Else
        .Initialized = False
      End If

    .KeyWord = Item.KeyWord
    .Locked = Item.Locked
    .Start = Item.Start
    .Text = Item.Text
    .Value = Item.Value
    .Negation = Item.Negation
    .ParaType = Item.ParaType
    .TemType = Item.TemType
End With

End Sub
'*************************************************************************
'**函 数 名：ShowParam_Type
'**输    入：(Integer)Tag
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2012-02-25 14:46:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function ShowParam_Type(Tag As Integer, Index As Long, Default_Type As String) As String
Dim intTem As Integer
Dim strTem As String
Dim strModule As String

If Tag <> Tag_Register And Tag <> Tag_Variable And Tag <> Tag_Local_Variable Then
    strModule = "[[index]][csvname]"
Else
    strModule = "[csvname]"
End If

If IsNumeric(Default_Type) Then
    intTem = Val(Default_Type)
    strTem = "(" & PublicTags(intTem) & ")"
Else
    If Tag <> Tag_Register And Tag <> Tag_Variable And Tag <> Tag_Local_Variable Then
        strTem = "(" & PublicTags(Tag) & ")"
    Else
        strTem = ""
    End If
End If

ShowParam_Type = ShowParam(Tag, Index, strTem & strModule)
End Function

'*************************************************************************
'**函 数 名：ShowParam
'**输    入：(Integer)Tag
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-07-09 23:22:14
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function ShowParam(Tag As Integer, Index As Long, strModule As String) As String
Dim i As Long
i = Index
Select Case Tag
     Case 0
       ShowParam = ""
     Case Tag_Register
       ShowParam = strModule
       ShowParam = Replace(ShowParam, "[index]", i)
       ShowParam = Replace(ShowParam, "[dbname]", PublicTags(Tag) & i)
       ShowParam = Replace(ShowParam, "[csvname]", PublicTags(Tag) & i)
       ShowParam = Replace(ShowParam, "[csvname_pl]", PublicTags(Tag) & i)
       ShowParam = Replace(ShowParam, "[disname]", PublicTags(Tag) & i)

     Case Tag_Variable
       ShowParam = strModule
       ShowParam = Replace(ShowParam, "[index]", i)
       ShowParam = Replace(ShowParam, "[csvname]", PYTags(Tag) & GetVariableName(i, True))
     Case Tag_String
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", Strs(i).Name)
         ShowParam = Replace(ShowParam, "[csvname]", Strs(i).Name)
         ShowParam = Replace(ShowParam, "[csvname_pl]", Strs(i).Name)
         ShowParam = Replace(ShowParam, "[disname]", Strs(i).Name)   'A-E
     Case Tag_Item
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", itm(i).dbName)
         ShowParam = Replace(ShowParam, "[csvname]", itm(i).csvName)
         ShowParam = Replace(ShowParam, "[csvname_pl]", itm(i).csvName_pl)
         ShowParam = Replace(ShowParam, "[disname]", itm(i).disname)

     Case Tag_Troop
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", Trps(i).strID)
         ShowParam = Replace(ShowParam, "[csvname]", Trps(i).csvName)
         ShowParam = Replace(ShowParam, "[csvname_pl]", Trps(i).csvName_pl)
         ShowParam = Replace(ShowParam, "[disname]", Trps(i).strName)
         ShowParam = Replace(ShowParam, "[disname_pl]", Trps(i).strPtName)

     Case Tag_Faction
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", Factions(i).strID)
         ShowParam = Replace(ShowParam, "[csvname]", Factions(i).csvName)
         ShowParam = Replace(ShowParam, "[csvname_pl]", Factions(i).csvName)
         ShowParam = Replace(ShowParam, "[disname]", Factions(i).strName)
         ShowParam = Replace(ShowParam, "[disname_pl]", Factions(i).strName)

     Case Tag_Quest
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tag_Party_Tpl
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", PTs(i).ptID)
         ShowParam = Replace(ShowParam, "[csvname]", PTs(i).csvName)
         ShowParam = Replace(ShowParam, "[csvname_pl]", PTs(i).csvName)
         ShowParam = Replace(ShowParam, "[disname]", PTs(i).ptName)
         ShowParam = Replace(ShowParam, "[disname_pl]", PTs(i).ptName)

     Case Tag_Party
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", Parties(i).strID)
         ShowParam = Replace(ShowParam, "[csvname]", Parties(i).csvName)
         ShowParam = Replace(ShowParam, "[csvname_pl]", Parties(i).csvName)
         ShowParam = Replace(ShowParam, "[disname]", Parties(i).strName)
         ShowParam = Replace(ShowParam, "[disname_pl]", Parties(i).strName)

     Case Tag_Scene
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", Scenes(i).strID)
         ShowParam = Replace(ShowParam, "[csvname]", Scenes(i).strName)
         ShowParam = Replace(ShowParam, "[csvname_pl]", Scenes(i).strName)
         ShowParam = Replace(ShowParam, "[disname]", Scenes(i).strName)
         ShowParam = Replace(ShowParam, "[disname_pl]", Scenes(i).strName)

     Case Tag_Mission_tpl
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tag_Menu
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tag_Script
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tag_Particle_Sys
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", PSys(i).strID)
         ShowParam = Replace(ShowParam, "[csvname]", PSys(i).strID)
         ShowParam = Replace(ShowParam, "[csvname_pl]", PSys(i).strID)
         ShowParam = Replace(ShowParam, "[disname]", PSys(i).strID)
         ShowParam = Replace(ShowParam, "[disname_pl]", PSys(i).strID)

     Case Tag_Scene_Prop
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tag_Sound
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", Sounds(i).sndName)
         ShowParam = Replace(ShowParam, "[csvname]", Sounds(i).sndName)
         ShowParam = Replace(ShowParam, "[csvname_pl]", Sounds(i).sndName)
         ShowParam = Replace(ShowParam, "[disname]", Sounds(i).sndName)
         ShowParam = Replace(ShowParam, "[disname_pl]", Sounds(i).sndName)

     Case Tag_Local_Variable
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[csvname]", PYTags(Tag) & GetVariableName(i, False))
     Case Tag_Map_Icon
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", MapIcons(i).strID)
         ShowParam = Replace(ShowParam, "[csvname]", MapIcons(i).strID)
         ShowParam = Replace(ShowParam, "[csvname_pl]", MapIcons(i).strID)
         ShowParam = Replace(ShowParam, "[disname]", MapIcons(i).strID)
         ShowParam = Replace(ShowParam, "[disname_pl]", MapIcons(i).strID)

     Case Tag_Skill
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", PublicSkills(i))
         ShowParam = Replace(ShowParam, "[csvname]", PublicSkills(i))
         ShowParam = Replace(ShowParam, "[csvname_pl]", PublicSkills(i))
         ShowParam = Replace(ShowParam, "[disname]", PublicSkills(i))
         ShowParam = Replace(ShowParam, "[disname_pl]", PublicSkills(i))

     Case Tag_Mesh
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", Mesh(i).strID)
         ShowParam = Replace(ShowParam, "[csvname]", Mesh(i).strID)
         ShowParam = Replace(ShowParam, "[csvname_pl]", Mesh(i).strID)
         ShowParam = Replace(ShowParam, "[disname]", Mesh(i).strID)
         ShowParam = Replace(ShowParam, "[disname_pl]", Mesh(i).strID)
     Case Tag_Presentation
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tag_Quick_String
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tag_Track
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tag_Tableau
         ShowParam = strModule
         ShowParam = Replace(ShowParam, "[index]", i)
         ShowParam = Replace(ShowParam, "[dbname]", TabMat(i).strID)
         ShowParam = Replace(ShowParam, "[csvname]", TabMat(i).strID)
         ShowParam = Replace(ShowParam, "[csvname_pl]", TabMat(i).strID)
         ShowParam = Replace(ShowParam, "[disname]", TabMat(i).strID)
         ShowParam = Replace(ShowParam, "[disname_pl]", TabMat(i).strID)

     Case Tag_Animation
         ShowParam = PublicTags(Tag) & i   'A-E
     Case Tags_End
       ShowParam = ""
   End Select
End Function

'*************************************************************************
'**函 数 名：GetVariableName
'**输    入：(Long)Pid
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-07-20 16:11:36
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************
Public Function GetVariableName(Pid As Long, Optional IsGlobal As Boolean = False) As String
Dim tN As String, Chk_Ves As Type_Variable_Name_Check_List, tIndex As String, Index As Integer, strPre As String

If IsGlobal Then
  Chk_Ves = TemGVarNameList
  Index = 1
  strPre = PublicTags(Tag_Variable)
Else
  Chk_Ves = CurVarNameList
  Index = CheckListTrgIdx
  strPre = PublicTags(Tag_Local_Variable)
End If

If UBound(Chk_Ves.Triggers) >= Index Then
  With Chk_Ves.Triggers(Index)
    If Pid <= UBound(.Checks) Then
      If .Checks(Pid) <> "" Then
        GetVariableName = .Checks(Pid)
      Else
        GetVariableName = strPre & Pid
      End If
    Else
      GetVariableName = strPre & Pid
    End If
  End With
Else
  GetVariableName = strPre & Pid
End If
End Function

'*************************************************************************
'**函 数 名：ShowNoTagParam
'**输    入：(String)ParaType
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-07-20 16:23:42
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function ShowNoTagParam(ParaType As String, Pid As String) As String
Dim temStr As String, TemType As String, q As Boolean
Dim i As Integer

q = IsOptionalParam(ParaType, TemType)

Select Case TemType    'ends_add
  Case ""
    temStr = Pid
  Case "0"
    temStr = Pid
  Case "pos"
    temStr = PublicTags(Tags_End + 1) & Pid
  Case "s"
    temStr = PublicTags(Tags_End + 10) & Pid
  Case "itp"
    temStr = Item_Type(Val(Pid)).Y
  Case "tf"
    For i = 0 To UBound(Tf)
        If I64toStrNZ(Tf(i).Value) = Pid Then Exit For
    Next i
    temStr = Tf(i).csvName
  Case "pf"
    For i = 0 To UBound(Pf)
        If I64toStrNZ(Pf(i).Value) = Pid Then Exit For
    Next i
    temStr = Pf(i).csvName
  Case "bs"
    temStr = BoolSwitch(Val(Pid)).Y
  Case "as"
    If Pid = "1" Then
       temStr = AbsSwitch(1).Y
    Else
       temStr = AbsSwitch(0).Y
    End If
  Case "ap"
    temStr = AccessPrivilege(Val(Pid)).Y
  Case "po"
    temStr = PlayOption(Val(Pid)).Y
  Case "ai_bhvr"
    temStr = AI_Bhvr(Val(Pid)).Y
  Case Else
    temStr = Pid
End Select

ShowNoTagParam = temStr
End Function
'*************************************************************************
'**函 数 名：ShowNoTagParamValue
'**输    入：(String)ParaType
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-02-17 17:45:12
'**修 改 人：
'**日    期：
'**版    本：V1.15
'*************************************************************************
Public Function ShowNoTagParamValue(ParaType As String, Pid As String) As String
Dim temStr As String, TemType As String, q As Boolean

q = IsOptionalParam(ParaType, TemType)

Select Case TemType    'ends_add
  Case ""
    temStr = Pid
  Case "0"
    temStr = Pid
  Case "pos"
    temStr = Pid
  Case "s"
    temStr = Pid
  Case "itp"
    temStr = Pid
  Case "tf"
    temStr = I64toStrNZ(Tf(Val(Pid)).Value)
  Case "pf"
    temStr = I64toStrNZ(Pf(Val(Pid)).Value)
  Case "bs"
    temStr = Pid
  Case "ap"
    temStr = Pid
  Case "as"
    temStr = Pid
  Case "po"
    temStr = Pid
  Case "ai_bhvr"
    temStr = Pid
  Case Else
    temStr = Pid
End Select

ShowNoTagParamValue = temStr
End Function
'*************************************************************************
'**函 数 名：IsOptionalParam
'**输    入：(String)ParaType
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-09-04 15:56:56
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function IsOptionalParam(ParaType As String, Optional ParaType_Simplified As String) As Boolean

IsOptionalParam = False
ParaType_Simplified = ParaType
If Len(ParaType) > 1 Then
  If Right(ParaType, 1) = "#" Then
    ParaType_Simplified = Left(ParaType, Len(ParaType) - 1)
    IsOptionalParam = True
  End If
End If

End Function

'*************************************************************************
'**函 数 名：MenuType
'**输    入：(String)ParaType
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-08-03 19:22:24
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function MenuType(ParaType As String) As Long
Dim temStr As String, i As Integer

'ends_add
If ParaType = "itp" Then
  MenuType = 1
  Exit Function
ElseIf ParaType = "tf" Then
  MenuType = 1
  Exit Function
ElseIf ParaType = "pf" Then
  MenuType = 1
  Exit Function
ElseIf ParaType = "bs" Then
  MenuType = 1
  Exit Function
ElseIf ParaType = "ap" Then
  MenuType = 1
  Exit Function
ElseIf ParaType = "as" Then
  MenuType = 1
  Exit Function
ElseIf ParaType = "po" Then
  MenuType = 1
  Exit Function
ElseIf ParaType = "ai_bhvr" Then
  MenuType = 1
  Exit Function
End If

For i = 1 To 26
  If Val(ParaType) = i Then
    MenuType = 1
    Exit Function
  End If
Next i

MenuType = 0
End Function

'*************************************************************************
'**函 数 名：GetNoTagParamInfo
'**输    入：(String)Param
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-08-11 13:01:57
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetNoTagParamInfo(Param As String, ParaType As String, Pid As String) As Boolean
Dim temStr As String, i As Integer

'pos
If LCase(Left(Param, 3)) = "pos" Then
  temStr = Replace(Param, "pos", "", , , vbTextCompare)
  temStr = Replace(temStr, " ", "", , , vbTextCompare)
  
  If IsNumeric(temStr) Then
    ParaType = "pos"
    Pid = temStr
    GetNoTagParamInfo = True
    Exit Function
  End If

's
ElseIf LCase(Left(Param, 1)) = "s" Then
  temStr = Replace(Param, "s", "", , , vbTextCompare)
  temStr = Replace(temStr, " ", "", , , vbTextCompare)
  
  If IsNumeric(temStr) Then
    ParaType = "s"
    Pid = temStr
    GetNoTagParamInfo = True
    Exit Function
  End If
End If

'pos
If LCase(Left(Param, Len(PublicTags(Tags_End + 1)))) = LCase(PublicTags(Tags_End + 1)) Then
  temStr = Replace(Param, PublicTags(Tags_End + 1), "", , , vbTextCompare)
  temStr = Replace(temStr, " ", "", , , vbTextCompare)
  
  If IsNumeric(temStr) Then
    ParaType = "pos"
    Pid = temStr
    GetNoTagParamInfo = True
    Exit Function
  End If
  
's
ElseIf LCase(Left(Param, Len(PublicTags(Tags_End + 10)))) = LCase(PublicTags(Tags_End + 10)) Then
  temStr = Replace(Param, PublicTags(Tags_End + 10), "", , , vbTextCompare)
  temStr = Replace(temStr, " ", "", , , vbTextCompare)
  
  If IsNumeric(temStr) Then
    ParaType = "s"
    Pid = temStr
    GetNoTagParamInfo = True
    Exit Function
  End If
End If

'0
If IsNumeric(Param) Then
  Pid = Trim(Param)
  ParaType = "0"
  GetNoTagParamInfo = True
End If

End Function

