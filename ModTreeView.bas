Attribute VB_Name = "ModTreeView"

'*************************************************************************
'**函 数 名：GetTVKeyInfo
'**输    入：Key As String, Trg As Integer, Op As Integer, Para As Integer
'**输    出：无
'**功能描述：通过TreeView项键值获得触发器有关信息
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-25 13:44:05
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub GetTVKeyInfo(Key As String, Trg As Integer, Op As Integer, Param As Integer)
Dim tArr As Variant, i As Integer, Str As String

For i = 1 To Len(Key)
    If Mid(Key, i, 1) = "(" Then
         Str = Right(Key, Len(Key) - i)
         Str = Replace(Str, ")", "")
    End If
Next i

tArr = Split(Str, ",")

If Str = "" Then
      Trg = 0
      Op = -1
      Param = -1
Else
    Trg = tArr(0)
    If UBound(tArr) > 0 Then
      Op = tArr(1)
      If UBound(tArr) > 1 Then
          Param = tArr(2)
      End If
    End If
End If

End Sub

'*************************************************************************
'**函 数 名：LoadChangeFrmParam
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-25 13:44:05
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadChangeFrmParam(Op As String, Param_No As Integer, Param_Value As String, Combo As ComboBox)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim i As Integer, Index As Integer, tStr As String, Tag_No As Integer, IsPos As Boolean, IsItp As Boolean, ParamID As String           'A_P

Index = GetOpIndex(RemoveOperationNegations(Op))

If Index >= 0 Then
    If Param_No <= UBound(Operation(Index).Para) Then
        tStr = Replace(Operation(Index).Para(Param_No).Para_Type, "#", "")           '通过注册获知参数类型,并去除可选参数标识“#”
    End If
End If

GetParamCodeInfo Param_Value, Tag_No, ParamID    '通过参数值获知参数类型
'Dim tP As Long
'QuickGetParamCodeInfo Param_Value, Tag_no, tP
'ParamID = CStr(tP)

If tStr = "pos" Then       'ends_add
    IsPos = True
    tStr = "0"
ElseIf tStr = "itp" Then
    IsItp = True
    tStr = "0"
ElseIf tStr = "" Then
    tStr = CStr(Tag_No)
End If

Select Case CInt(Val(tStr))
       Case Tag_Script, Tag_Register, Tag_Variable, Tag_Local_Variable, Tag_String, Tag_Quest, Tag_Mission_tpl, Tag_Menu, Tag_Scene_Prop, Tag_Skill, Tag_Presentation, Tag_Quick_String, Tag_Track, Tag_Animation '没做完
            Combo.Clear
            If Tag_No > 0 Then Combo.Text = Tags(Tag_No) & "_" & ParamID
       Case Tag_Item   '物品
            LoadItemCombo Combo
            If Tag_No = Tag_Item Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then       '如果是变量
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Troop     '兵种
            LoadTroopCombo Combo
            If Tag_No = Tag_Troop Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Faction   '阵营
            LoadFactionCombo Combo
            If Tag_No = Tag_Faction Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Party_Tpl '部队模板
            LoadPartyTemplateCombo Combo
            If Tag_No = Tag_Party_Tpl Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Party    '部队
            LoadPartyCombo Combo
            If Tag_No = Tag_Party Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Scene    '场景
            LoadSceneCombo Combo
            If Tag_No = Tag_Scene Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Particle_Sys    '粒子系统
            LoadPSysCombo Combo
            If Tag_No = Tag_Particle_Sys Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Sound    '声音
            LoadSoundCombo Combo
            If Tag_No = Tag_Sound Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Map_Icon    '大地图图标
            LoadMapIconCombo Combo
            If Tag_No = Tag_Map_Icon Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Mesh    '网格模型
            LoadMeshCombo Combo
            If Tag_No = Tag_Mesh Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case Tag_Tableau    '可变材质
            LoadTabMatCombo Combo
            If Tag_No = Tag_Tableau Then
                 Combo.ListIndex = CLng(ParamID)
            ElseIf Tag_No > 0 Then
                 Combo.Text = Tags(Tag_No) & "_" & ParamID
            End If
       Case 0
            Combo.Clear
            If IsPos Then
                LoadPosCombo Combo
                If Tag_No > 0 Then
                   Combo.Text = Tags(Tag_No) & "_" & ParamID
                Else
                   Combo.ListIndex = CInt(Val(ParamID))
                End If
            ElseIf IsItp Then
                LoadItemTypeCombo Combo
                If Tag_No > 0 Then
                   Combo.Text = Tags(Tag_No) & "_" & ParamID
                Else
                   Combo.ListIndex = CInt(Val(ParamID)) - 1
                End If
            ElseIf Tag_No > 0 Then
                LoadParamCombo Tags(Tag_No), Combo
                If Combo.ListCount <= 0 Then
                    Combo.Text = Tags(Tag_No) & "_" & ParamID
                Else
                    Combo.ListIndex = CLng(ParamID)
                End If
            Else
                Combo.Text = Param_Value
            End If
       Case Else
            Combo.Clear
            Combo.Text = Param_Value
End Select

If IsPos Then
       ChangeTag = "Para_pos"
ElseIf IsItp Then
       ChangeTag = "Para_itp"
Else
   If Tag_No = Tag_Register Or Tag_No = Tag_Variable Or Tag_No = Tag_Local_Variable Then
       ChangeTag = "Para_0"
   Else
       ChangeTag = "Para_" & CInt(Val(tStr))
   End If
End If

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("ModOperation", "LoadChangeFrmParam", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadItemCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-08 18:14:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadItemCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_Item - 1
             CB.AddItem "(" & i & ")" & itm(i).dbName
         Next i
End Sub

'*************************************************************************
'**函 数 名：LoadItemTypeCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-18 15:59:51
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadItemTypeCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = LBound(Item_Type) To UBound(Item_Type)
             CB.AddItem Item_Type(i).Y
         Next i
End Sub
'*************************************************************************
'**函 数 名：LoadTroopCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-08 18:14:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadTroopCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_Troop - 1
             CB.AddItem "(" & i & ")" & Trps(i).strID
         Next i
End Sub
'*************************************************************************
'**函 数 名：LoadPartyCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-08 18:15:19
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadPartyCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_Party - 1
             CB.AddItem "(" & i & ")" & Parties(i).strID
         Next i
End Sub
'*************************************************************************
'**函 数 名：LoadPartyTemplateCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-08 18:15:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadPartyTemplateCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_PT - 1
             CB.AddItem "(" & i & ")" & PTs(i).ptID
         Next i
End Sub
'*************************************************************************
'**函 数 名：LoadSceneCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-08 18:16:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadSceneCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_Scene - 1
             CB.AddItem "(" & i & ")" & Scenes(i).strID
         Next i
End Sub
'*************************************************************************
'**函 数 名：LoadMapIconCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-08 18:16:57
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadMapIconCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_MapIcon - 1
             CB.AddItem "(" & i & ")" & MapIcons(i).strID
         Next i
End Sub
'*************************************************************************
'**函 数 名：LoadFactionCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-08 18:17:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadFactionCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_Faction - 1
             CB.AddItem "(" & i & ")" & Factions(i).strID
         Next i
End Sub
'*************************************************************************
'**函 数 名：LoadSoundCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-14 23:21:15
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadSoundCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_Sound - 1
             CB.AddItem "(" & i & ")" & Sounds(i).sndName
         Next i
         
End Sub

'*************************************************************************
'**函 数 名：LoadPosCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-14 23:21:15
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadPosCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_Pos - 1
             CB.AddItem "Pos" & i
         Next i
End Sub

'*************************************************************************
'**函 数 名：LoadPSysCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-14 23:21:15
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadPSysCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_PSys - 1
             CB.AddItem "(" & i & ")" & PSys(i).strID
         Next i
         
End Sub

'*************************************************************************
'**函 数 名：LoadTabMatCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-23 23:22:45
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadTabMatCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_TabMat - 1
             CB.AddItem "(" & i & ")" & TabMat(i).strID
         Next i
         
End Sub

'*************************************************************************
'**函 数 名：LoadMeshCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-2-3 17:21:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadMeshCombo(CB As ComboBox, Optional Clear As Boolean = True)
Dim i As Long
If Clear Then CB.Clear
         For i = 0 To N_Mesh - 1
             CB.AddItem "(" & i & ")" & Mesh(i).strID
         Next i
         
End Sub

'*************************************************************************
'**函 数 名：LoadOpCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-25 16:49:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadOpCombo(CB As ComboBox)
Dim i As Long
CB.Clear
         For i = 0 To UBound(Operation)
             CB.AddItem Operation(i).Op_CSVname
         Next i
         
End Sub

'*************************************************************************
'**函 数 名：GetIndentation
'**输    入：(String)Text
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-2-27 12:48:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetIndentation(Text As String) As Integer
Dim i As Integer

For i = 1 To Len(Text)
     If Mid(Text, i, 1) <> " " Then
        GetIndentation = i - 1
        Exit For
     End If
Next i

End Function

'*************************************************************************
'**函 数 名：GetIndentationStr
'**输    入：(String)Text
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-2-27 12:48:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function GetIndentationStr(Indentation As Integer) As String
Dim i As Integer, tStr As String
tStr = ""
For i = 1 To Indentation
     tStr = tStr & " "
Next i
GetIndentationStr = tStr

End Function
'*************************************************************************
'**函 数 名：CalcOperationIndentation
'**输    入：(Type_Op_Block)OpBlock(),(String)Indentation()
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-2-27 12:48:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub CalcOperationIndentation(OpBlock() As Type_Op_Block, Indentation() As String)
Dim i As Integer, n As Integer, H As Integer, tStr As String, temOp() As Type_Op_Block

ReDim temOp(LBound(OpBlock) To UBound(OpBlock))

For i = LBound(temOp) To UBound(temOp)
    temOp(i) = OpBlock(i)
Next i

ReDim Indentation(LBound(OpBlock) To UBound(OpBlock))

For i = LBound(temOp) To UBound(temOp)
    If temOp(i).Op = CStr(try_end) Then
         For n = i To LBound(temOp) Step -1
              tStr = RemoveOperationNegations(temOp(n).Op)
              If Val(tStr) = try_begin Or Val(tStr) = try_for_range Or Val(tStr) = try_for_range_backwards Or Val(tStr) = try_for_parties Or Val(tStr) = try_for_agents Then
                  For H = n + 1 To i - 1
                      If Val(temOp(H).Op) = else_try Then
                          temOp(H).Op = "0"
                      Else
                          Indentation(H) = Indentation(H) & "    "
                      End If
                  Next H
                  temOp(n).Op = "0"
                  Exit For
              End If
         Next n
    End If
Next i

End Sub

'*************************************************************************
'**函 数 名：SwapTVOperation
'**输    入：(TreeView)TreeView, (Str)Key
'**输    出：(Int)-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-2-17 20:30:29
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub SwapTVOperation(TreeView As TreeView, OpKey1 As String, OpKey2 As String)
Dim i As Integer, n As Integer, H As Integer, i2 As Integer, n2 As Integer, h2 As Integer, j As Integer, tStr As String, oNode As node, OriginChildren1 As Integer, OriginChildren2 As Integer, temText() As String, temText2() As String
Dim tC As Long

With TreeView
   If .Nodes(OpKey1).Children = 0 Then
       ReDim temText(0)
   Else
       ReDim temText(1 To .Nodes(OpKey1).Children)
   End If
   
   If .Nodes(OpKey2).Children = 0 Then
       ReDim temText2(0)
   Else
       ReDim temText2(1 To .Nodes(OpKey2).Children)
   End If

        '--------交换operation节点Text--------
          tStr = .Nodes(OpKey1).Text
          tC = .Nodes(OpKey1).ForeColor
         .Nodes(OpKey1).Text = .Nodes(OpKey2).Text
         .Nodes(OpKey1).ForeColor = .Nodes(OpKey2).ForeColor
         .Nodes(OpKey2).Text = tStr
         .Nodes(OpKey2).ForeColor = tC
        '-------------------------------------
        '-----将OpKey的Text存入temText-----
         If .Nodes(OpKey1).Children > 0 Then
             temText(1) = .Nodes(OpKey1).Child.Text
             Set oNode = .Nodes(OpKey1).Child.Next
             For j = 2 To .Nodes(OpKey1).Children
                temText(j) = oNode.Text
                 If Not (oNode.Next Is Nothing) Then Set oNode = oNode.Next
             Next j
         End If
         
         If .Nodes(OpKey2).Children > 0 Then
             temText2(1) = .Nodes(OpKey2).Child.Text
             Set oNode = .Nodes(OpKey2).Child.Next
             For j = 2 To .Nodes(OpKey2).Children
                temText2(j) = oNode.Text
                 If Not (oNode.Next Is Nothing) Then Set oNode = oNode.Next
             Next j
         End If
        '---------------------------------------------
        '--------根据参数数目调整子节点数目并更改Text---------
         OriginChildren1 = .Nodes(OpKey1).Children
         OriginChildren2 = .Nodes(OpKey2).Children
         
         If .Nodes(OpKey1).Children > .Nodes(OpKey2).Children Then
              GetTVKeyInfo OpKey2, i, n, H
              GetTVKeyInfo OpKey1, i2, n2, h2
              For j = OriginChildren2 + 1 To OriginChildren1
                  .Nodes.Remove .Nodes("Op(" & i2 & "," & n2 & "," & j & ")").Index
                  .Nodes.Add OpKey2, tvwChild, "Op(" & i & "," & n & "," & j & ")", ""
              Next j
         ElseIf .Nodes(OpKey1).Children < .Nodes(OpKey2).Children Then
              GetTVKeyInfo OpKey1, i, n, H
              GetTVKeyInfo OpKey2, i2, n2, h2
              For j = OriginChildren1 + 1 To OriginChildren2
                  .Nodes.Remove .Nodes("Op(" & i2 & "," & n2 & "," & j & ")").Index
                  .Nodes.Add OpKey1, tvwChild, "Op(" & i & "," & n & "," & j & ")", ""
              Next j
         End If
         
              '----------更改Text---------
               If .Nodes(OpKey1).Children > 0 Then
                  .Nodes(OpKey1).Child.Text = temText2(1)
                  Set oNode = .Nodes(OpKey1).Child.Next
                  For j = 2 To .Nodes(OpKey1).Children
                      oNode.Text = temText2(j)
                      If Not (oNode.Next Is Nothing) Then Set oNode = oNode.Next
                  Next j
               End If

               If .Nodes(OpKey2).Children > 0 Then
                  .Nodes(OpKey2).Child.Text = temText(1)
                  Set oNode = .Nodes(OpKey2).Child.Next
                  For j = 2 To .Nodes(OpKey2).Children
                      oNode.Text = temText(j)
                      If Not (oNode.Next Is Nothing) Then Set oNode = oNode.Next
                  Next j
               End If

              '---------------------------
        '-----------------------------------------------------

End With

End Sub
'*************************************************************************
'**函 数 名：SelectTreeViewItem
'**输    入：(TreeView)TreeView,(String)Key
'**输    出：无
'**功能描述：根据Key选中TreeView项目
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-01-26 17:25:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub SelectTreeViewItem(TreeView As TreeView, Key As String)
Dim t As Integer

For t = 1 To TreeView.Nodes.Count
      If TreeView.Nodes(t).Key = Key Then
            TreeView.Nodes(t).Selected = True
            Exit For
      End If
Next t

End Sub

Public Function GetTVLabelHint(Label As String, Optional Value As String = "") As String
Dim i As Integer

For i = 1 To Len(Label)
     If Mid(Label, i, 1) = ":" Then
         GetTVLabelHint = Left(Label, i - 1)
         Value = Right(Label, Len(Label) - i)
         Exit For
     End If
Next i

End Function

'*************************************************************************
'**函 数 名：GlobalCreateOp
'**输    入：(TreeView)TreeView,(Type_Op_Block)OpBlock(),(Str)ParentKey,(Int)Op_Index
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-16 14:34:16
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub GlobalCreateOp(TreeView As TreeView, OpBlock() As Type_Op_Block, OpNum As Long, Optional ParentKey As String = "", Optional Op_ID As String = CStr(Call_Script), Optional IfHint As Boolean = True, Optional IfBold As Boolean, Optional FontColor As Long = vbBlack)
Dim i As Integer, n As Integer, H As Integer, Op_Index As Integer, Hint As String

If IfHint Then
     Hint = PublicMsgs(68) & ":"
Else
     Hint = ""
End If

OpNum = OpNum + 1
'---------------分配空间------------------
If OpNum = 1 Then
      ReDim OpBlock(1 To OpNum)
Else
      ReDim Preserve OpBlock(1 To OpNum)
End If
'-----------------------------------------

GetTVKeyInfo ParentKey, i, n, H

Op_Index = GetOpIndex(RemoveOperationNegations(Op_ID))

With OpBlock(OpNum)
     .Op = Op_ID
     If ParentKey <> "" Then
        TreeView.Nodes.Add ParentKey, tvwChild, "Op(" & i & "," & OpNum & ",0" & ")", Hint & Operation(Op_Index).Op_CSVname
     Else
        TreeView.Nodes.Add , , "Op(" & i & "," & OpNum & ",0" & ")", Hint & Operation(Op_Index).Op_CSVname
     End If
     
     TreeView.Nodes("Op(" & i & "," & OpNum & ",0" & ")").Bold = IfBold
     TreeView.Nodes("Op(" & i & "," & OpNum & ",0" & ")").ForeColor = FontColor
End With

End Sub
'*************************************************************************
'**函 数 名：TVCreateParam
'**输    入：(int)op_index,(int)trg_index, (int)act_index
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-15 15:44:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub TVCreateParam(TreeView As TreeView, OpBlock As Type_Op_Block, ParentKey As String, Op_Index As Integer, Optional Indentation As String = "", Optional IfPrefix As Boolean = False)
Dim j As Integer, IO As Boolean, i As Integer, n As Integer, H As Integer, Prefix As String
    If Len(Indentation) >= 2 Then
         Indentation = Left(Indentation, Len(Indentation) - 2)
    End If

    If Operation(Op_Index).ParaNum > 0 Then
        GetTVKeyInfo ParentKey, i, n, H
        For j = 1 To Operation(Op_Index).ParaNum
          If IfPrefix = True Then
             If j < Operation(Op_Index).ParaNum Then
                 Prefix = Chr(25) & Chr(6)
             Else
                 Prefix = Chr(3) & Chr(6)
             End If
          Else
             Prefix = ""
          End If
          
             If ParentKey <> "" Then
              TreeView.Nodes.Add ParentKey, tvwChild, "Op(" & i & "," & n & "," & j & ")", Indentation & Prefix & GetParaEntity(RemoveOperationNegations(OpBlock.Op), j, OpBlock.Para(j), IO)
             Else
              TreeView.Nodes.Add , , "Op(" & i & "," & n & "," & j & ")", Indentation & Prefix & GetParaEntity(RemoveOperationNegations(OpBlock.Op), j, OpBlock.Para(j), IO)
             End If
        Next j
    End If
End Sub

'*************************************************************************
'**函 数 名：TVLoadParam
'**输    入：TreeView As TreeView, OpBlock As Type_Op_Block, ParentKey As String, Optional Indentation As String = ""
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-30 17:00:00
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub TVLoadParam(TreeView As TreeView, OpBlock As Type_Op_Block, ParentKey As String, Optional Indentation As String = "", Optional IfPrefix As Boolean = False)
Dim H As Integer, i As Integer, n As Integer, Hint As String, Prefix As String

    If Len(Indentation) >= 1 Then
         Indentation = Left(Indentation, Len(Indentation) - 1)
    End If

GetTVKeyInfo ParentKey, i, n, H
With OpBlock
   For H = 1 To .ParaNum
       'FixParam .Para(H).Value, Hint
       
       If IfPrefix = True Then
         If H < .ParaNum Then
            Prefix = Chr(25) & Chr(6)
         Else
            Prefix = Chr(3) & Chr(6)
         End If
       Else
          Prefix = ""
       End If
             
       If Hint <> "" Then
          If i > 0 Then
              MsgBox ActiveString(PublicMsgs(69), PublicMsgs(65), i, PublicMsgs(68), n, Hint), vbOKOnly, PublicMsgs(0)
          Else
              MsgBox ActiveString(PublicMsgs(69), "", "", PublicMsgs(68), n, Hint), vbOKOnly, PublicMsgs(0)
          End If
       End If
       
       If ParentKey <> "" Then
          TreeView.Nodes.Add ParentKey, tvwChild, "Op(" & i & "," & n & "," & H & ")", Indentation & Prefix & GetParaEntity(CLng(Val(RemoveOperationNegations(.Op))), H, .Para(H), , True)
       Else
          TreeView.Nodes.Add , , "Op(" & i & "," & n & "," & H & ")", Indentation & Prefix & GetParaEntity(CLng(Val(RemoveOperationNegations(.Op))), H, .Para(H), , True)
       End If

   Next H
End With

End Sub

'*************************************************************************
'**函 数 名：TVLoadOp
'**输    入：Item As Type_Item, trg_Idx As Integer, tiAct_Idx As Integer, param_Idx As Integer
'**输    出：无
'**功能描述：创建Op节点(不创建参数子节点)
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-30 17:00:00
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub TVLoadOp(TreeView As TreeView, OpBlock As Type_Op_Block, Act_No As Integer, NotStr As String, OrStr As String, Optional ParentKey As String = "", Optional Indentation As String = "", Optional IfHint As Boolean = True, Optional FontColor As Long = vbBlack, Optional IfBold As Boolean = False)
Dim i As Integer, n As Integer, H As Integer, tNotStr As String, tOrStr As String, tStr As String, Hint As String, Negation As Integer, OpID As Long

If IfHint Then
    Hint = PublicMsgs(68) & ":"
Else
    Hint = ""
End If

With OpBlock
   QuickGetOpCodeInfo .Op, Negation, OpID
   tStr = GetOpStrWithoutNeg(OpID)

   Select Case Negation
          Case 0
               tNotStr = ""
               tOrStr = ""
          Case 1
               tNotStr = NotStr & "|"
               tOrStr = ""
          Case 2
               tOrStr = OrStr & "|"
               tNotStr = ""
          Case 3
               tNotStr = NotStr & "|"
               tOrStr = OrStr & "|"
   End Select
   
   GetTVKeyInfo ParentKey, i, n, H

   If ParentKey <> "" Then
        TreeView.Nodes.Add ParentKey, tvwChild, "Op(" & i & "," & Act_No & ",0)", Hint & Indentation & tOrStr & tNotStr & tStr
   Else
        TreeView.Nodes.Add , , "Op(" & i & "," & Act_No & ",0)", Hint & Indentation & tOrStr & tNotStr & tStr
   End If
   
        TreeView.Nodes("Op(" & i & "," & Act_No & ",0)").Bold = IfBold
        TreeView.Nodes("Op(" & i & "," & Act_No & ",0)").ForeColor = FontColor
        
End With

End Sub

'*************************************************************************
'**函 数 名：AttachParamIndentation
'**输    入：(TreeView)TreeView, (Str)OpKey
'**输    出：无
'**功能描述：使参数缩进与操作缩进一致
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-16 11:37:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub AttachParamIndentation(TreeView As TreeView, OpKey As String)
Dim Indentation As String, oNode As node, i As Integer

With TreeView
   Indentation = GetIndentationStr(GetIndentation(.Nodes(OpKey).Text))
   If Len(Indentation) >= 1 Then
        Indentation = Left(Indentation, Len(Indentation) - 1)
   End If
   
   For i = 1 To .Nodes(OpKey).Children
        If i = 1 Then
           .Nodes(OpKey).Child.Text = Indentation & Trim(.Nodes(OpKey).Child.Text)
            If Not (.Nodes(OpKey).Child.Next Is Nothing) Then Set oNode = .Nodes(OpKey).Child.Next
        Else
            oNode.Text = Indentation & Trim(oNode.Text)
            If Not (oNode.Next Is Nothing) Then Set oNode = oNode.Next
        End If
   Next i
   
End With

End Sub

'*************************************************************************
'**函 数 名：ChangeTVParamPrefix
'**输    入：
'**输    出：无
'**功能描述：更改treeview中参数节点的制表符
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-17 23:00:23
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub ChangeTVParamPrefix(ParamText As String, Last As Boolean)
Dim tStr As String

If Last Then
    tStr = Chr(3) & Chr(6)
Else
    tStr = Chr(25) & Chr(6)
End If

If Len(ParamText) >= 2 Then
    If Left(ParamText, 2) = Chr(3) & Chr(6) Or Left(ParamText, 2) = Chr(25) & Chr(6) Then
        ParamText = tStr & Right(ParamText, Len(ParamText) - 2)
    End If
End If

End Sub

'*************************************************************************
'**函 数 名：LoadParamCombo
'**输    入：(Str)Tag
'**输    出：无
'**功能描述：加载参数combo
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-20 11:09:49
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub LoadParamCombo(Tag As String, Combo As ComboBox, Optional Clear As Boolean = True)

Select Case Tag
       Case Tags(Tag_Item)
            LoadItemCombo Combo, Clear
       Case Tags(Tag_Troop)
            LoadTroopCombo Combo, Clear
       Case Tags(Tag_Party)
            LoadPartyCombo Combo, Clear
       Case Tags(Tag_Party_Tpl)
            LoadPartyTemplateCombo Combo, Clear
       Case Tags(Tag_Scene)
            LoadSceneCombo Combo, Clear
       Case Tags(Tag_Map_Icon)
            LoadMapIconCombo Combo, Clear
       Case Tags(Tag_Tableau)
            LoadTabMatCombo Combo, Clear
       Case Tags(Tag_Faction)
            LoadFactionCombo Combo, Clear
       Case Tags(Tag_Particle_Sys)
            LoadPSysCombo Combo, Clear
       Case Tags(Tag_Sound)
            LoadSoundCombo Combo, Clear
       Case Tags(Tag_Mesh)
            LoadMeshCombo Combo, Clear
            
       'ends_add
       Case "pos"
            LoadPosCombo Combo, Clear
       Case "itp"
            LoadItemTypeCombo Combo, Clear
End Select

End Sub


'*************************************************************************
'**函 数 名：ClearComboPreserveText
'**输    入：(ComboBox)Combo
'**输    出：无
'**功能描述：清空Combo列表但是保留Text
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-20 12:18:15
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub ClearComboPreserveText(Combo As ComboBox)
Dim tStr As String, Sel As Integer

     CBText = Combo.Text
     Sel = Combo.SelStart
     Combo.Clear
     Combo.Text = CBText
     Combo.SelStart = Sel
     
End Sub
