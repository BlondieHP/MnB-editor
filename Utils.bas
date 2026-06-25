Attribute VB_Name = "Utils"

Option Explicit

Function showHexId(ID As String) As String
    On Error Resume Next
    showHexId = CStr(Hex(Val(ID)))
End Function


'*************************************************************************
'**函 数 名：PublicInit
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-17 18:42:18
'**修 改 人：SSgt_Edward
'**日    期：2010-11-23 22:40:09
'**版    本：V1.1321
'*************************************************************************

Function GetTroopSkill(TroopID As Integer, ByVal sklid As Integer) As Byte
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim a As Integer64b, s As String, b As String
    
    sklid = sklid - 1
    
    If TroopID = -1 Then TroopID = CurrentTrpID
    If TroopID = -1 Then Exit Function
    
    With Trps(TroopID)
        
        a = StrToI64(.Skills((sklid \ 8) + 1))
        s = I64toHexStr(a)
        s = Mid(s, Len(s) - ((sklid) Mod 8), 1)
    
    End With
    GetTroopSkill = HexStrToI4(s)
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Utils", "GetTroopSkill", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：GetSkill
'**输    入：ByVal sklid(Integer) -
'**输    出：(Byte) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:24:38
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************

Function GetSkill(ByVal sklid As Integer) As Byte
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim a As Integer64b, s As String, b As String
    
    sklid = sklid - 1
    
    If CurrentTrpID = -1 Then Exit Function
    
    With Trps(CurrentTrpID)
        
        a = StrToI64(.Skills((sklid \ 8) + 1))
        s = I64toHexStr(a)
        s = Mid(s, Len(s) - ((sklid) Mod 8), 1)
    
    End With
    GetSkill = HexStrToI4(s)
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Utils", "GetSkill", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：PutSkill
'**输    入：intVal(Integer)         -
'**        ：ByVal sklid(Integer) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:24:41
'**修 改 人：SSgt_Edward
'**日    期：2010-11-24 15:23:09
'**版    本：V1.1321
'*************************************************************************

Sub PutSkill(Trp As Type_Troops, intVal As Integer, ByVal sklid As Integer)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim a As Integer64b, s As String, b As String
    Dim n As Long
    
    Call Init_Integer64b

    With Trp
    sklid = sklid - 1
    a = StrToI64(.Skills((sklid \ 8) + 1))
    s = I64toHexStr(a)
        For n = 0 To 7
            If n <> (sklid Mod 8) Then
                b = Mid(s, Len(s) - n, 1) + b
            Else
                b = CHex(intVal) + b
            End If
        Next n
        .Skills((sklid \ 8) + 1) = I64toStrNZ(HexStrToI64(b))
    End With
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("Utils", "PutSkill", Err.Number, Err.Description)
End Sub


Sub load_KIS_Team()
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim troopsID As Long
    Dim tempStr As String
    Dim n As Integer
    Dim strTmp As String
    Dim intLoop As Integer

    CountDream = 0

    strTmp = ReadString(MnBInfo.ModIniFileName, "KIS_Team", "GuardsManCount", 250)

    If Val(strTmp) > 0 Then
        intLoop = Val(strTmp)
    Else
        intLoop = 99
    End If
    For n = 0 To Val(strTmp)
        strTmp = ReadString(MnBInfo.ModIniFileName, "KIS_Team", "GuardsMan_" & CStr(n), 250)

        If Len(strTmp) = 0 Or strTmp = "0" Or strTmp = "-1" Then
            strTmp = "-1"
        Else
            troopsID = Val(strTmp)
            DreamTeam(n) = troopsID

            strTmp = fnNumberFixedLength(CLng(CountDream), 3, " ")
            tempStr = strTmp & "["

            strTmp = fnNumberFixedLength(troopsID, 3, "")
            tempStr = tempStr & strTmp & "]"

            If Len(Trps(troopsID).csvName) = 0 Then
                tempStr = tempStr & " " & Trps(troopsID).strID & " "
            Else
                tempStr = tempStr & " " & Trps(troopsID).csvName & " "
            End If

            'PTForm.Dream_list.AddItem tempStr
            CountDream = CountDream + 1
        End If
    Next n

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("PTForm", "KIS_Team", Err.Number, Err.Description)
End Sub




'*************************************************************************
'**函 数 名：updateTriggersFile
'**输    入：id(Integer)   - trigger的ID，更新第id个trigger。0=添加。
'**        ：trigger(String) - trigger内容。
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-06-14 09:43:09
'**修 改 人：
'**日    期：
'**版    本：V0.955.6
'*************************************************************************

Public Function updateTriggersFile(ID As Integer, Trigger As String) As Integer
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim strFileName As String
    Dim fileLineNum As Integer
    Dim strTemp As String
    Dim fileBuff As String
    Dim Count As Long
    Dim opFlag As Boolean

    opFlag = False

    updateTriggersFile = -1
    
    If ID < 0 Then Exit Function
    If ID = 0 And Len(Trigger) = 0 Then Exit Function

    Count = 0

    strFileName = MnBInfo.ModName & "\triggers.txt"
    MaxPointer = FileLen(strFileName)
    lngHandle = FreeFile()
    Open strFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    fileLineNum = 0
    Do Until 0
        strTemp = GetRealLine()
        'fix me
        If Len(strTemp) = 0 Then
            'If fileLineNum >= count + 2 Then
            '    strTemp = GetRealLine()
            'End If
        End If
        fileLineNum = fileLineNum + 1
        If Pointer > MaxPointer Then
            If ID = 0 Then
                fileBuff = fileBuff & strTemp & vbCrLf
                'append trigger .
                fileBuff = fileBuff & Trigger
            Else
                If opFlag = False Then
                    fileBuff = fileBuff & Trigger
                Else
                    fileBuff = fileBuff & strTemp
                End If
            End If
            Exit Do
        End If

        If fileLineNum = 2 Then
            'get trigger 数量记录.
            Count = Val(strTemp)
        End If

        ' add mode
        If ID = 0 Then
            If fileLineNum = 2 Then
                '增加 trigger 数量记录.
                Count = Count + 1
                fileBuff = fileBuff & CStr(Count) & vbCrLf
            Else
                fileBuff = fileBuff & strTemp & vbCrLf
            End If
        Else
            'update mode
            If Count > 0 And fileLineNum = 2 + ID Then
                '替换 trigger .
                fileBuff = fileBuff & Trigger & vbCrLf
                opFlag = True
            Else
                fileBuff = fileBuff & strTemp & vbCrLf
            End If

        End If

    Loop
    Close #lngHandle

    'write to file
    lngHandle = FreeFile()
    Open strFileName For Output As #lngHandle
    Print #lngHandle, fileBuff
    Close #lngHandle

    'fix me
    'OutAsDebugTex (fileBuff)

    updateTriggersFile = fileLineNum - 2

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    updateTriggersFile = -1
    Call logErr("Utils", "updateTriggersFile", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：FixPath
'**输    入：path(String) -
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-06-21 06:54:18
'**修 改 人：SSgt_Edward
'**日    期：2010-11-21 23:13:54
'**版    本：V1.132.21
'*************************************************************************

Public Function FixPath(Path As String) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    FixPath = Trim$(Path)
    If Len(Path) = 0 Then
        Exit Function
    End If

    Dim i As Long
    If Right(FixPath, 1) = "\" Then
        FixPath = Left(FixPath, Len(FixPath) - 1)
    End If

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Utils", "FixPath", Err.Number, Err.Description)
End Function


'*************************************************************************
'**函 数 名：get_Module_Version
'**输    入：ModIniFilePath(String) -
'**输    出：(String) -
'**功能描述：
'**全局变量：MnBInfo.ModPath
'**调用模块：
'**作    者：kevin
'**日    期：2008-06-21 06:43:50
'**修 改 人：
'**日    期：
'**版    本：V0.955.23
'*************************************************************************
Private Function get_Module_Version(ModIniFilePath As String) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim TemP As String
    Dim arrTmp() As String
    Dim Key As String
    Dim ii As Integer
    
    Dim strKeyFileName As String
    strKeyFileName = "module.ini"

    Key = "works_with_version_max"
    ii = Len(Key)

    lngHandle = FreeFile()

    Open MnBInfo.ModPath & strKeyFileName For Input As #lngHandle
    Do Until EOF(lngHandle)
        Input #lngHandle, TemP
        If LCase$(Trim$(Left$(TemP, ii))) = Key Then
            arrTmp = Split(TemP, "=")
            If LCase$(Trim$(arrTmp(0))) = Key Then
                get_Module_Version = LCase$(Trim$(arrTmp(1)))
                Exit Function
            End If
        End If
    Loop
    Close #lngHandle

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    'Call logErr("Utils", "get_Module_Version", Err.Number, Err.Description)
    get_Module_Version = "--"
End Function


Public Function computeListID(objListBox As ListBox, intMode As Integer) As Long
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    computeListID = -1
    If objListBox Is Nothing Then Exit Function

    Dim tempStr As String
    Dim arrTmp() As String
    Dim i As Integer

    If intMode = 1 Then
        tempStr = objListBox.Text
        
        arrTmp = Split(tempStr, " ")
        
        For i = 0 To UBound(arrTmp)
            arrTmp(i) = Trim$(arrTmp(i))
            If Len(arrTmp(i)) > 0 Then
                tempStr = arrTmp(i)
                Exit For
            End If
        Next

        computeListID = Val(tempStr)
    Else
        computeListID = objListBox.ListIndex
    End If

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Utils", "computeListID", Err.Number, Err.Description)
End Function


'*************************************************************************
'**函 数 名：ShowItemListByType
'**输    入：objItemListBox(ListBox)     -
'**        ：objItemTypeListBox(ListBox) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-06-29 07:56:16
'**修 改 人：
'**日    期：
'**版    本：V0.960.7
'*************************************************************************

Public Sub ShowItemListByType(objItemListBox As ListBox, objItemTypeListBox As ListBox)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim itmType As Integer
    Dim intShowItemType As Integer
    Dim tmpCount As Long
    Dim Scrollwidth As Long
    Dim strTmp As String
    
    Dim n As Long

    If objItemListBox Is Nothing Then Exit Sub
    If objItemTypeListBox Is Nothing Then Exit Sub

    If objItemTypeListBox.ListIndex = -1 Then
        Exit Sub
    End If

    intShowItemType = objItemTypeListBox.ListIndex

    tmpCount = 0
    ReDim itmID(0)
    objItemListBox.Clear
    Scrollwidth = 0

    For n = 0 To N_Item - 1

        If intShowItemType = 0 Then
            ' show all
            itmType = 0
        Else
            If Val(itm(n).itmType) < 2100000000 Then
                itmType = Val(itm(n).itmType) Mod 256
            Else
                itmType = MinusIT(itm(n).itmType) Mod 256
            End If
        End If

        If itmType = intShowItemType Then
            If Len(itm(n).csvName) > 0 Then
                strTmp = CStr(n) & " " & itm(n).csvName & " " & itm(n).disname
            Else
                strTmp = CStr(n) & " " & itm(n).disname & " " & itm(n).csvName
            End If

            If Len(strTmp) > Scrollwidth Then
                Scrollwidth = Len(strTmp)
            End If

            objItemListBox.AddItem (strTmp)
            itmID(tmpCount) = n
            tmpCount = tmpCount + 1
            ReDim Preserve itmID(tmpCount)
        End If
    Next n

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("Utils", "ShowItemListByType", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**函 数 名：getTXTID
'**输    入：tag_type(Integer) -
'**        ：id(Long)          -
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-07-27 06:48:49
'**修 改 人：
'**日    期：
'**版    本：V0.960.38
'*************************************************************************

Public Function getTXTID(Tag_Type As Integer, ID As Long, Optional strID As String) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    'tag_register        =  1
    'tag_variable        =  2
    'tag_string          =  3
    'tag_item            =  4
    'tag_troop           =  5
    'tag_faction         =  6
    'tag_quest           =  7
    'tag_party_tpl       =  8
    'tag_party           =  9
    'tag_scene           = 10
    'tag_mission_tpl     = 11
    'tag_menu            = 12
    'tag_script          = 13
    'tag_particle_sys    = 14
    'tag_scene_prop      = 15
    'tag_sound           = 16
    'tag_local_variable  = 17
    'tag_map_icon        = 18
    'tag_skill           = 19
    'tag_mesh            = 20
    'tag_presentation    = 21
    'tag_quick_string    = 22
    'tag_track           = 23
    'tag_tableau         = 24
    'tag_animation       = 25
    'tags_end            = 26

    Dim i64b_num1 As Integer64b
    Dim i64b_num2 As Integer64b
    
    Call Init_Integer64b

    i64b_num1 = HexStrToI64(CStr(Hex$(Tag_Type)) & "00000000000000")
    i64b_num2 = StrToI64(CStr(ID))
    i64b_num2 = Plus64b(i64b_num1, i64b_num2)
    
    getTXTID = I64toStrNZ(i64b_num2)
    
    If Not IsMissing(strID) Then
        strID = GetstrID(Tag_Type, CStr(ID))
    End If
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Utils", "getTXTID", Err.Number, Err.Description)

End Function

'*************************************************************************
'**函 数 名：GetstrID
'**输    入：Tag_Type(Integer),Index(String) -
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-03-21 20:23:53
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function GetstrID(Tag_Type As Integer, Index As String) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
  Select Case Tag_Type
    Case Tag_Register
    
    Case Tag_Variable
    
    Case Tag_String
    Case Tag_Item
         GetstrID = itm(Val(Index)).dbName
    Case Tag_Troop
         GetstrID = Trps(Val(Index)).strID
    Case Tag_Faction
         GetstrID = Factions(Val(Index)).strID
    Case Tag_Quest
    Case Tag_Party_Tpl
         GetstrID = PTs(Val(Index)).ptID
    Case Tag_Party
         GetstrID = Parties(Val(Index)).strID
    Case Tag_Scene
         GetstrID = Scenes(Val(Index)).strID
    Case Tag_Mission_tpl
    Case Tag_Menu
    Case Tag_Script
    Case Tag_Particle_Sys
         GetstrID = PSys(Val(Index)).strID
    Case Tag_Scene_Prop
    Case Tag_Sound
         GetstrID = Sounds(Val(Index)).sndName
    Case Tag_Local_Variable
    Case Tag_Map_Icon
         GetstrID = MapIcons(Val(Index)).strID
    Case Tag_Skill
    Case Tag_Mesh
         GetstrID = Mesh(Val(Index)).strID
    Case Tag_Presentation
    Case Tag_Quick_String
    Case Tag_Track
    Case Tag_Tableau
         GetstrID = TabMat(Val(Index)).strID
    Case Tag_Animation
    Case Tags_End
  End Select
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Utils", "GetstrID", Err.Number, Err.Description)

End Function


'*************************************************************************
'**函 数 名：InitWarbandInfo
'**输    入：无
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-21 22:52:06
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function QuickReadFile(ByVal FilePath As String) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim F As Integer, strTem As String
    
    F = FreeFile
    
    Open FilePath For Input As #F
          Line Input #F, strTem
          
          QuickReadFile = strTem
    Close #F
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Utils", "QuickReadFile", Err.Number, Err.Description)
End Function


Public Function IsDirectory(ByVal Path As String) As Boolean
If GetAttr(Path) And vbDirectory Then
      If Right(Path, 1) <> "." Then
          IsDirectory = True
      End If
End If
End Function

'*************************************************************************
'**函 数 名：MnBtoRGBColor
'**输    入：(String)MnBColor
'**输    出：(String) -
'**功能描述：实现骑砍颜色与RGB颜色的转换
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-24 22:58:28
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function MnBtoRGBColor(ByVal MnBColor As String) As String
Dim r As Byte, G As Byte, b As Byte, MnBColorI64 As Integer64b, RGBColorI64 As Integer64b
'分解RGB颜色值
MnBColorI64 = StrToI64(MnBColor)

r = MnBColorI64.by(2)
G = MnBColorI64.by(1)
b = MnBColorI64.by(0)
RGBColorI64.by(0) = r
RGBColorI64.by(1) = G
RGBColorI64.by(2) = b

MnBtoRGBColor = I64toHexStr(RGBColorI64)
MnBtoRGBColor = Right(MnBtoRGBColor, 6)
'RGB-BGR
End Function

'*************************************************************************
'**函 数 名：RGBtoMnBColor
'**输    入：(String)RGBColor
'**输    出：(String) -
'**功能描述：实现骑砍颜色与RGB颜色的转换
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-24 22:58:28
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function RGBtoMnBColor(ByVal RGBColor As String) As String
Dim r As Byte, G As Byte, b As Byte, MnBColorI64 As Integer64b, RGBColorI64 As Integer64b
'分解RGB颜色值
RGBColorI64 = HexStrToI64(RGBColor)

b = RGBColorI64.by(2)
G = RGBColorI64.by(1)
r = RGBColorI64.by(0)
MnBColorI64.by(0) = b
MnBColorI64.by(1) = G
MnBColorI64.by(2) = r

RGBtoMnBColor = I64toStrNZ(MnBColorI64)
'RGB-BGR
End Function

'*************************************************************************
'**函 数 名：FindItem
'**输    入：...
'**输    出：(ListItem) -
'**功能描述：完善ListView查询功能
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-27 21:19:17
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function FindItem(oListView As ListView, Start As Long, Row As String, KeyWord As String, bPartial As Boolean, Compare As VbCompareMethod, Optional bReverse As Boolean = False) As ListItem
If oListView.ListItems.Count > 0 And Start <= oListView.ListItems.Count And Start >= 0 And Trim(Row) <> "" Then
Dim i As Long, n As Long, r() As String, j As Long, Finded As Boolean, SearchInfo(2) As Long

If Not bReverse Then
SearchInfo(0) = Start + 1
SearchInfo(1) = oListView.ListItems.Count
SearchInfo(2) = 1
Else
SearchInfo(0) = Start - 1
SearchInfo(1) = 1
SearchInfo(2) = -1
End If

r = Split(Row, "|")
    For i = SearchInfo(0) To SearchInfo(1) Step SearchInfo(2)
    With oListView.ListItems(i)
    
    Finded = False
            For j = 0 To UBound(r)
               If Val(r(j)) = 0 Then
                  n = InStr(1, .Text, KeyWord, Compare)
                 If bPartial Then
                   If n > 0 Then Finded = True
                 Else
                   If n = 1 And Len(KeyWord) = Len(.Text) Then Finded = True
                 End If
                  If Finded Then
                    Set FindItem = oListView.ListItems(i)
                    Exit Function
                  End If
            
               Else
                  n = InStr(1, .SubItems(Val(r(j))), KeyWord, Compare)
                    If bPartial Then
                       If n > 0 Then Finded = True
                    Else
                       If n = 1 And Len(KeyWord) = Len(.SubItems(Val(r(j)))) Then Finded = True
                    End If
                  If Finded Then
                    Set FindItem = oListView.ListItems(i)
                    Exit Function
                  End If
               End If
            Next j
    End With
    Next i
End If
End Function
'*************************************************************************
'**函 数 名：ClearResource
'**输    入：(Type_ResourceInSound)Resource
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-08 22:58:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub ClearResource(Resource As Type_ResourceInSound)
With Resource
      .ID = -1
      .Unknown = 0
      .strID = ""
End With
End Sub

'*************************************************************************
'**函 数 名：SwapResource
'**输    入：(Type_ResourceInSound)SoundRes1,(Type_ResourceInSound)SoundRes2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-08 22:15:05
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapResource(Resource1 As Type_ResourceInSound, Resource2 As Type_ResourceInSound)
Dim TemResource As Type_ResourceInSound

TemResource = Resource1
Resource1 = Resource2
Resource2 = TemResource
End Sub

'*************************************************************************
'**函 数 名：SwapInventory
'**输    入：(Type_XY_Index)Inventory1,(Type_XY_Index)Inventory2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-19 19:53:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapInventory(Inventory1 As Type_XY_Index, Inventory2 As Type_XY_Index)
Dim TemInventory As Type_XY_Index

TemInventory = Inventory1
Inventory1 = Inventory2
Inventory2 = TemInventory
End Sub


'*************************************************************************
'**函 数 名：ClearStack
'**输    入：(Type_Stacks)Stack
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-08 22:58:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub ClearStack(Stack As Type_Stacks)
With Stack
      .ID = -1
      .Max = 0
      .Min = 0
      .Flags = 0
End With
End Sub

'*************************************************************************
'**函 数 名：SwapStacks
'**输    入：(Type_Stacks)Stack1,(Type_Stacks)Stack2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-08 22:15:05
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapStacks(Stack1 As Type_Stacks, Stack2 As Type_Stacks)
Dim TemStack As Type_Stacks

TemStack = Stack1
Stack1 = Stack2
Stack2 = TemStack
End Sub

'*************************************************************************
'**函 数 名：SwapMapIcons
'**输    入：(Long)MapIcon1,(Long)MapIcon2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 00:46:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapMapIcons(ByVal MapIcon1 As Long, ByVal MapIcon2 As Long)
Dim TemMapIcon As Type_MapIcon

TemMapIcon = MapIcons(MapIcon1)
 MapIcons(MapIcon1) = MapIcons(MapIcon2)
 MapIcons(MapIcon2) = TemMapIcon
 
SwapLong MapIcons(MapIcon1).ID, MapIcons(MapIcon2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapSounds
'**输    入：(Long)Sound1,(Long)Sound2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-03 14:21:40
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapSounds(ByVal Sound1 As Long, ByVal Sound2 As Long)
Dim TemSound As Type_Sound

TemSound = Sounds(Sound1)
 Sounds(Sound1) = Sounds(Sound2)
 Sounds(Sound2) = TemSound
 
SwapLong Sounds(Sound1).ID, Sounds(Sound2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapSoundRess
'**输    入：(Long)SoundRes1,(Long)SoundRes2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-03 14:21:40
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapSoundRess(ByVal SoundRes1 As Long, ByVal SoundRes2 As Long)
Dim TemSoundRes As Type_SoundResource

TemSoundRes = SoundRess(SoundRes1)
 SoundRess(SoundRes1) = SoundRess(SoundRes2)
 SoundRess(SoundRes2) = TemSoundRes
 
SwapLong SoundRess(SoundRes1).ID, SoundRess(SoundRes2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapLong
'**输    入：(Long)Long1,(Long)Long2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-14 23:20:21
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapLong(Long1 As Long, Long2 As Long)
Dim TemLong As Long

TemLong = Long1
 Long1 = Long2
 Long2 = TemLong
End Sub

'*************************************************************************
'**函 数 名：SwapString
'**输    入：(String)String1,(String)String2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-08 10:57:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapString(String1 As String, String2 As String)
Dim TemString As String

TemString = String1
 String1 = String2
 String2 = TemString
End Sub


'*************************************************************************
'**函 数 名：SwapSingle
'**输    入：(Single)Single1,(Single)Single2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-08 10:58:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapSingle(Single1 As Single, Single2 As Single)
Dim TemSingle As Single

TemSingle = Single1
Single1 = Single2
Single2 = TemSingle
End Sub

'*************************************************************************
'**函 数 名：SwapItems
'**输    入：(Long)Item1,(Long)Item2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-14 23:41:42
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapItems(ByVal Item1 As Long, ByVal Item2 As Long)
Dim TemItem As Type_Item

TemItem = itm(Item1)
 itm(Item1) = itm(Item2)
 itm(Item2) = TemItem
 
 SwapLong itm(Item1).ID, itm(Item2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapPSys
'**输    入：(Long)PSys1,(Long)PSys2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-7 14:44:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapPSys(ByVal PSys1 As Long, ByVal PSys2 As Long)
Dim TemPSys As Type_Particle_System

TemPSys = PSys(PSys1)
 PSys(PSys1) = PSys(PSys2)
 PSys(PSys2) = TemPSys
 
 SwapLong PSys(PSys1).ID, PSys(PSys2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapMesh
'**输    入：(Long)Mesh1,(Long)Mesh2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-29 21:09:10
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapMesh(ByVal Mesh1 As Long, ByVal Mesh2 As Long)
Dim TemMesh As Type_Mesh

TemMesh = Mesh(Mesh1)
 Mesh(Mesh1) = Mesh(Mesh2)
 Mesh(Mesh2) = TemMesh
 
 SwapLong Mesh(Mesh1).ID, Mesh(Mesh2).ID
End Sub
'*************************************************************************
'**函 数 名：SwapTabMat
'**输    入：(Long)TabMat1,(Long)TabMat2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-23 18:01:32
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapTabMat(ByVal tabmat1 As Long, ByVal tabmat2 As Long)
Dim Temtabmat As Type_Tableau_Material

Temtabmat = TabMat(tabmat1)
 TabMat(tabmat1) = TabMat(tabmat2)
 TabMat(tabmat2) = Temtabmat
 
 SwapLong TabMat(tabmat1).ID, TabMat(tabmat2).ID
End Sub
'*************************************************************************
'**函 数 名：SwapTimeTrg
'**输    入：(Long)TimeTrg1,(Long)TimeTrg2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-3-4 23:01:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapTimeTrg(ByVal TimeTrg1 As Long, ByVal TimeTrg2 As Long)
Dim TemTimeTrg As Type_Time_Trigger

TemTimeTrg = TimeTrg(TimeTrg1)
 TimeTrg(TimeTrg1) = TimeTrg(TimeTrg2)
 TimeTrg(TimeTrg2) = TemTimeTrg
 
 SwapLong TimeTrg(TimeTrg1).ID, TimeTrg(TimeTrg2).ID
End Sub
'*************************************************************************
'**函 数 名：SwapScenes
'**输    入：(Long)Scene1,(Long)Scene2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-09 23:53:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapScenes(ByVal Scene1 As Long, ByVal Scene2 As Long)
Dim TemScene As Type_Scene

TemScene = Scenes(Scene1)
 Scenes(Scene1) = Scenes(Scene2)
 Scenes(Scene2) = TemScene
 
 SwapLong Scenes(Scene1).ID, Scenes(Scene2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapParties
'**输    入：(Long)Party1,(Long)Party2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-08 23:22:29
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapParties(ByVal Party1 As Long, ByVal Party2 As Long)
Dim TemParty As Type_Party

TemParty = Parties(Party1)
 Parties(Party1) = Parties(Party2)
 Parties(Party2) = TemParty
 
 SwapLong Parties(Party1).ID, Parties(Party2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapPTs
'**输    入：(Long)PT1,(Long)PT2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-07 16:58:44
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapPTs(ByVal PT1 As Long, ByVal PT2 As Long)
Dim TemPT As Type_PT

TemPT = PTs(PT1)
 PTs(PT1) = PTs(PT2)
 PTs(PT2) = TemPT
 
 SwapLong PTs(PT1).ID, PTs(PT2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapTroops
'**输    入：(Long)Troop1,(Long)Troop2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-29 22:31:27
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapTroops(ByVal Troop1 As Long, ByVal Troop2 As Long)
Dim TemTroop As Type_Troops

TemTroop = Trps(Troop1)
 Trps(Troop1) = Trps(Troop2)
 Trps(Troop2) = TemTroop
 
  SwapLong Trps(Troop1).ID, Trps(Troop2).ID
End Sub

'*************************************************************************
'**函 数 名：SwapFactions
'**输    入：(Long)Faction1,(Long)Faction2
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 21:37:01
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapFactions(ByVal Faction1 As Long, ByVal Faction2 As Long)
Dim TemFaction As Type_Faction

TemFaction = Factions(Faction1)
 Factions(Faction1) = Factions(Faction2)
 Factions(Faction2) = TemFaction
 
SwapLong Factions(Faction1).ID, Factions(Faction2).ID
End Sub

'*************************************************************************
'**函 数 名：RedimFactions
'**输    入：(Long)NewCount
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 21:52:31
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub RedimFactions(ByVal NewCount As Long)
Dim i As Long

N_Faction = NewCount
ReDim Preserve Factions(N_Faction - 1)

End Sub


'*************************************************************************
'**函 数 名：SwapListItem
'**输    入：(ListItem)Item1,(ListItem)Item2,(Integer)SubItemCount
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-29 22:33:52
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SwapListItem(ByVal Item1 As ListItem, ByVal Item2 As ListItem, Optional SubItemCount = 0, Optional NoTextSwap As Boolean = False)
Dim strTem As String, i As Integer

If Not NoTextSwap Then
strTem = Item1.Text
Item1.Text = Item2.Text
Item2.Text = strTem
End If

For i = 1 To SubItemCount    'item1.ListSubItems.Count
     strTem = Item1.SubItems(i)
     Item1.SubItems(i) = Item2.SubItems(i)
     Item2.SubItems(i) = strTem

Next i
End Sub

'*************************************************************************
'**函 数 名：OutputItemLine
'**输    入：(Long)Itm_Idx[如为-1则指当前Item]
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-10-49 10:49:26
'**版    本：V1.1321
'*************************************************************************

Public Function CheckEditable(ByVal QueryType As Long, ByVal Index As Long) As Boolean

If QueryType <> EditInfo_ItemsCount Then
    CheckEditable = Index >= MnBInfo.EditInfo(QueryType)
    '更新于2013-06-28, 去除删除保护
    CheckEditable = True
Else
    CheckEditable = (Index >= MnBInfo.EditInfo(QueryType) - 1) And (Index < N_Item - 1)
    '更新于2013-06-28, 去除删除保护
    CheckEditable = Index < N_Item - 1
End If

End Function

'*************************************************************************
'**函 数 名：CheckExist
'**输    入：(Long)QueryType,(Long)Index
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-30 16:30:09
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function CheckExist(ByVal QueryType As Long, ByVal Index As Long) As Boolean

Select Case QueryType     'A_E
       Case EditInfo_TroopsCount
            CheckExist = Index > -1 And Index < N_Troop
       Case EditInfo_FactionsCount
            CheckExist = Index > -1 And Index < N_Faction
       Case EditInfo_ItemsCount
            CheckExist = Index > -1 And Index < N_Item
       Case EditInfo_PartiesCount
            CheckExist = Index > -1 And Index < N_Party
       Case EditInfo_PartyTemplatesCount
            CheckExist = Index > -1 And Index < N_PT
       Case EditInfo_ScenesCount
            CheckExist = Index > -1 And Index < N_Scene
       Case EditInfo_MapIconsCount
            CheckExist = Index > -1 And Index < N_MapIcon
       Case EditInfo_SoundsCount
            CheckExist = Index > -1 And Index < N_Sound
       Case EditInfo_SoundRessCount
            CheckExist = Index > -1 And Index < N_SoundRes
       Case EditInfo_PSysCount
            CheckExist = Index > -1 And Index < N_PSys
       Case EditInfo_TabMatCount
            CheckExist = Index > -1 And Index < N_TabMat
       Case EditInfo_MeshCount
            CheckExist = Index > -1 And Index < N_Mesh
       Case EditInfo_TimeTrgCount
            CheckExist = Index > -1 And Index < N_TimeTrg
       Case EditInfo_StringsCount
            CheckExist = Index > -1 And Index < N_Str
End Select
End Function

'*************************************************************************
'**函 数 名：SaveItemCSVFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-30 22:08:41
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function GetYesNoStr(ByVal lVal As Long) As String
If lVal = 0 Then
   GetYesNoStr = "否"
Else
   GetYesNoStr = "是"
End If
End Function

'*************************************************************************
'**函 数 名：SaveAll
'**输    入： -
'**输    出： -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-22 15:35:26
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub MkDirEx(ByVal FilePath As String)
Dim i As Long, k As String, FirstFolder As Boolean, FolderNow As String

FirstFolder = True
For i = 1 To Len(FilePath)
        k = Mid(FilePath, i, 1)
        If k = "\" Then
           If Not FirstFolder Then
              FolderNow = Left(FilePath, i - 1)
              If Dir(FolderNow, vbDirectory + vbHidden + vbNormal + vbSystem) = "" Then
                  MkDir FolderNow
              End If
           Else
              FirstFolder = False
           End If
        End If
Next i

If Dir(FilePath, vbDirectory + vbHidden + vbNormal + vbSystem) = "" Then
      MkDir FilePath
End If

End Sub

'*************************************************************************
'**函 数 名：IsWeapon
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为武器
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-29 12:22:10
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsWeapon(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_one_handed_wpn Or IT = itp_type_two_handed_wpn _
Or IT = itp_type_polearm Or IT = itp_type_bow Or IT = itp_type_crossbow Or IT = itp_type_thrown _
Or IT = itp_type_pistol Or IT = itp_type_musket Then
   IsWeapon = True
Else
   IsWeapon = False
End If

End Function

'*************************************************************************
'**函 数 名：IsMeleeWeapon
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为近战武器
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-3 23:28:30
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsMeleeWeapon(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_one_handed_wpn Or IT = itp_type_two_handed_wpn _
Or IT = itp_type_polearm Then
   IsMeleeWeapon = True
Else
   IsMeleeWeapon = False
End If

End Function

'*************************************************************************
'**函 数 名：IsRangedWeapon
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为远程武器
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-3 23:29:35
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsRangedWeapon(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_bow Or IT = itp_type_crossbow Or IT = itp_type_thrown Then
   IsRangedWeapon = True
Else
   IsRangedWeapon = False
End If

End Function

'*************************************************************************
'**函 数 名：IsFireArm
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为火器
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-3 23:29:40
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsFireArm(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_pistol Or IT = itp_type_musket Then
   IsFireArm = True
Else
   IsFireArm = False
End If

End Function

'*************************************************************************
'**函 数 名：IsAmmo
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为弹药
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-3 23:29:45
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsAmmo(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_arrows Or IT = itp_type_bolts Or IT = itp_type_bullets Then
   IsAmmo = True
Else
   IsAmmo = False
End If

End Function

'*************************************************************************
'**函 数 名：IsHorse
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为马匹
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-07 08:52:25
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsHorse(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_horse Then
   IsHorse = True
Else
   IsHorse = False
End If

End Function

'*************************************************************************
'**函 数 名：IsShield
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为盾牌
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-07 08:52:25
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsShield(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_shield Then
   IsShield = True
Else
   IsShield = False
End If

End Function

'*************************************************************************
'**函 数 名：IsGood
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为货物
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-08 23:15:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsGood(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_goods Then
   IsGood = True
Else
   IsGood = False
End If

End Function

'*************************************************************************
'**函 数 名：IsFood
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为食物
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-08 23:16:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsFood(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)

If ChkBit64b(tI, itp_food) Then
   IsFood = True
Else
   IsFood = False
End If

End Function

'*************************************************************************
'**函 数 名：IsBook
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为书本
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-08 23:16:00
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsBook(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_book Then
   IsBook = True
Else
   IsBook = False
End If

End Function


'*************************************************************************
'**函 数 名：IsAnimal
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为动物
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-08 23:16:05
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsAnimal(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_animal Then
   IsAnimal = True
Else
   IsAnimal = False
End If

End Function
'*************************************************************************
'**函 数 名：IsArmor
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为护甲类
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-07 14:48:42
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsArmor(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_head_armor Or IT = itp_type_hand_armor Or IT = itp_type_body_armor Or IT = itp_type_foot_armor Then
   IsArmor = True
Else
   IsArmor = False
End If

End Function
'*************************************************************************
'**函 数 名：IsHeadArmor
'**输    入：(Str)ItmType
'**输    出：无
'**功能描述：判断物品是否为头部护甲类
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-21 15:18:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsHeadArmor(itmType As String) As Boolean
Dim IT As Long

IT = GetItmType(itmType)

If IT = itp_type_head_armor Then
   IsHeadArmor = True
Else
   IsHeadArmor = False
End If

End Function

'*************************************************************************
'**函 数 名：GetItmType
'**输    入：(Str)ItmType
'**输    出：(Integer)-
'**功能描述：得到武器类型
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-29 22:24:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function GetItmType(itmType As String) As Integer

If Val(itmType) < 2100000000 Then

    GetItmType = Val(itmType) Mod 256
Else

    GetItmType = MinusIT(itmType) Mod 256
End If

End Function

'*************************************************************************
'**函 数 名：HaveDoubleUsages
'**输    入：(Str)ItmType
'**输    出：(Boolean)-
'**功能描述：判断物品是否有两种使用模式
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-29 22:10:51
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function HaveDoubleUsages(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
HaveDoubleUsages = ChkBit64b(tI, itp_next_item_as_melee)

End Function

'*************************************************************************
'**函 数 名：IsTwoHanded
'**输    入：(Str)ItmType
'**输    出：(Boolean)-
'**功能描述：判断物品是否可以持盾
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-4-12 23:19:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsTwoHanded(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsTwoHanded = ChkBit64b(tI, itp_two_handed)

End Function

'*************************************************************************
'**函 数 名：IsBonusAgainstShield
'**输    入：(Str)ItmType
'**输    出：(Boolean)-
'**功能描述：判断物品是否对盾加成
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-4-13 21:51:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsBonusAgainstShield(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsBonusAgainstShield = ChkBit64b(tI, itp_bonus_against_shield)

End Function

'*************************************************************************
'**函 数 名：IsUnbalanced
'**输    入：(Str)ItmType
'**输    出：(Boolean)-
'**功能描述：判断物品是否平衡
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-4-13 22:31:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsUnbalanced(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsUnbalanced = ChkBit64b(tI, itp_unbalanced)

End Function

'*************************************************************************
'**函 数 名：IsCrushThrough
'**输    入：(Str)ItmType
'**输    出：(Boolean)-
'**功能描述：判断物品是否破格挡
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-4-13 22:33:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsCrushThrough(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsCrushThrough = ChkBit64b(tI, itp_crush_through)

End Function

'*************************************************************************
'**函 数 名：IsCantUseOnHorseBack
'**输    入：(Str)ItmType
'**输    出：(Boolean)-
'**功能描述：判断物品是否不能在马上使用
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-4-13 22:46:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsCantUseOnHorseBack(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsCantUseOnHorseBack = ChkBit64b(tI, itp_cant_use_on_horseback)

End Function
'*************************************************************************
'**函 数 名：IsCantReloadOnHorseBack
'**输    入：(Str)ItmType
'**输    出：(Boolean)-
'**功能描述：判断物品是否不能在马上装填
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-4-13 22:54:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsCantReloadOnHorseBack(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsCantReloadOnHorseBack = ChkBit64b(tI, itp_cant_reload_on_horseback)

End Function

'*************************************************************************
'**函 数 名：IsCanPenetrateShield
'**输    入：(Str)ItmType
'**输    出：(Boolean)-
'**功能描述：判断物品是否能穿盾
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-4-13 22:50:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function IsCanPenetrateShield(itmType As String) As Boolean
Dim tI As Integer64b

tI = StrToI64(itmType)
IsCanPenetrateShield = ChkBit64b(tI, itp_can_penetrate_shield)

End Function
'*************************************************************************
'**函 数 名：GetAttachment
'**输    入：(String)itmType
'**输    出：(integer)-
'**功能描述：得到绑定类型
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-29 22:24:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function GetAttachment(itmType As String) As Integer
Dim tI As Integer64b, tI2 As Integer64b

tI = StrToI64(itmType)
GetAttachment = 0

If ChkBit64b(tI, itp_Attachment_Left_bit) And ChkBit64b(tI, itp_Attachment_Right_bit) And _
ChkBit64b(tI, itp_Attachment_Armature_bit1) And ChkBit64b(tI, itp_Attachment_Armature_bit2) Then
    GetAttachment = 4
    Exit Function
ElseIf ChkBit64b(tI, itp_Attachment_Left_bit) And ChkBit64b(tI, itp_Attachment_Right_bit) Then
    GetAttachment = 3
    Exit Function
ElseIf ChkBit64b(tI, itp_Attachment_Left_bit) Then
    GetAttachment = 1
    Exit Function
ElseIf ChkBit64b(tI, itp_Attachment_Right_bit) Then
    GetAttachment = 2
    Exit Function
End If

End Function

'*************************************************************************
'**函 数 名：GetDamage
'**输    入：(Long)Damage,(Integer)Damage_Type
'**输    出：(Long)-
'**功能描述：得到伤害数值及类型
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-29 22:24:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function GetDamage(Damage As Long, Damage_Type As Integer) As Long

If Damage - 2 ^ 9 >= 0 Then
    Damage_Type = 2
    GetDamage = Damage - 2 ^ 9
ElseIf Damage - 2 ^ 8 >= 0 Then
    Damage_Type = 1
    GetDamage = Damage - 2 ^ 8
Else
    Damage_Type = 0
    GetDamage = Damage
End If

End Function

'*************************************************************************
'**函 数 名：ExDamage
'**输    入：(Long)Damage,(Integer)Damage_Type
'**输    出：(Long)-
'**功能描述：得到TXT中记录形式的伤害数值
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-07 22:10:38
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function ExDamage(Damage As Long, Damage_Type As Integer) As Long

If Damage_Type > 0 Then
ExDamage = Damage + 2 ^ (7 + Damage_Type)
ElseIf Damage_Type = 0 Then
ExDamage = Damage
End If

End Function

'*************************************************************************
'**函 数 名：strIIf
'**输    入：(Boolean)Condition,(Boolean)strTrue,(Boolean)strFalse
'**输    出：(String)-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-08 21:56:24
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function strIIf(Condition As Boolean, strTrue As String, strFalse As String) As String

If Condition Then
   strIIf = strTrue
Else
   strIIf = strFalse
End If
End Function

'*************************************************************************
'**函 数 名：ReadItemCSVLine
'**输    入：(String)Text
'**输    出：(Long)-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-25 13:41:24
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function ReadItemCSVLine(ByVal Text As String) As Long
Dim i As Integer, n As Integer, TemP As String, StrLine() As String, arrTmp() As String, j As Integer

StrLine() = Split(Text, vbCrLf)
    For i = 0 To UBound(StrLine)
        TemP = StrLine(i)
        If Len(Trim$(TemP)) < 1 Then
            TemP = ""
        Else
            arrTmp = Split(TemP, "|")

            For n = 0 To TemItemCount - 1
                If LCase$(TemItems(n).dbName) = LCase$(arrTmp(0)) Then
                    TemItems(n).csvName = arrTmp(1)
                    'Exit For
                    j = j + 1
                End If
            
            
            If Right(arrTmp(0), 3) = "_pl" Then
                     If LCase$(TemItems(n).dbName) = LCase$(Left(arrTmp(0), Len(arrTmp(0)) - 3)) Then
                        TemItems(n).csvName_pl = arrTmp(1)
                        'Exit For
                     End If
            End If
            
            Next n
        End If
    Next i
    
ReadItemCSVLine = j
End Function

'*************************************************************************
'**函 数 名：LTrimEx
'**输    入：(String)Text
'**输    出：(String)-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-25 13:53:30
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function LTrimEx(ByVal Text As String) As String
Dim i As Integer, k As Long, temStr As String, n As Integer

If Len(Text) > 0 Then
For i = 1 To Len(Text)
     k = Asc(Mid(Text, i, 1))
     
     If k <> 13 And k <> 10 And k <> 32 Then
        n = i
     End If
     
Next i

If n > 0 Then
   LTrimEx = Right(Text, Len(Text) - n + 1)
Else
   
End If

End If
End Function


Public Function GetEditorIndex(ByVal Tag As String) As Long
Dim s As String

s = Right(Tag, Len(Tag) - 5)
GetEditorIndex = Val(s)
End Function


Public Function ActiveString(ByVal StrMain As String, ParamArray StrNew() As Variant) As String
Dim sKey As String, i As Integer, strTem As String

strTem = StrMain

For i = 0 To UBound(StrNew)
  sKey = "[str" & i & "]"
  strTem = Replace(strTem, sKey, CStr(StrNew(i)), , , vbTextCompare)
  
Next i

ActiveString = strTem
End Function

'*************************************************************************
'**函 数 名：StructureTroopInventory
'**输    入：(Long)Trp_Idx
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-19 20:05:16
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub StructureTroopInventory(Optional Trp_Idx As Long = -1)
Dim i As Integer, TemList() As Type_XY_Index, TrpVes As Type_Troops

If Trp_Idx <= -1 Then
   TrpVes = CurrentTrp
Else
   TrpVes = Trps(Trp_Idx)
End If

With TrpVes
ReDim TemList(0)
For i = 1 To 64
        If .lstInventory(i).X > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .lstInventory(i)
        End If
Next i

   For i = 1 To 64
     If i <= UBound(TemList) Then
           .lstInventory(i) = TemList(i)
     Else
           .lstInventory(i).X = -1
           .lstInventory(i).strX = ""
           .lstInventory(i).Y = 0
     End If
   Next i

End With

If Trp_Idx <= -1 Then
   CurrentTrp = TrpVes
Else
    Trps(Trp_Idx) = TrpVes
End If

End Sub

'*************************************************************************
'**函 数 名：StructurePartyStacksEx
'**输    入：(Long)Party_Idx
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-03-20 11:59:01
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub StructurePartyStacksEx(ByVal Party_Idx As Long)
Dim i As Integer, TemList() As Type_Stacks

With Parties(Party_Idx)
ReDim TemList(0)
For i = 1 To .StacksCount
        If .Stacks(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Stacks(i)
        End If
Next i

.StacksCount = UBound(TemList)

If .StacksCount > 0 Then
ReDim .Stacks(1 To .StacksCount)
   For i = 1 To .StacksCount
           .Stacks(i) = TemList(i)
   Next i
End If
End With
End Sub

'*************************************************************************
'**函 数 名：StructureItemFactions
'**输    入：(Long)Itm_Idx
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-03-21 17:07:25
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub StructureItemFactions(ByVal Itm_Idx As Long)
Dim i As Integer, TemList() As Type_Chest, ItmVes As Type_Item

If Itm_Idx <= -1 Then
   ItmVes = CurrentItm
Else
   ItmVes = itm(Itm_Idx)
End If

With ItmVes
ReDim TemList(0)
For i = 1 To .FactionCount
        If .Faction(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Faction(i)
        End If
Next i

.FactionCount = UBound(TemList)

If .FactionCount > 0 Then
ReDim .Faction(1 To .FactionCount)
     For i = 1 To .FactionCount
        .Faction(i) = TemList(i)
     Next i
End If
End With

If Itm_Idx <= -1 Then
   CurrentItm = ItmVes
Else
   itm(Itm_Idx) = ItmVes
End If
End Sub

'*************************************************************************
'**函 数 名：StructureSceneChests
'**输    入：(Long)Scene_Idx
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-03-20 22:59:44
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub StructureSceneChests(ByVal Scene_Idx As Long)
Dim i As Integer, TemList() As Type_Chest, SceneVes As Type_Scene

If Scene_Idx <= -1 Then
   SceneVes = CurrentScene
Else
   SceneVes = Scenes(Scene_Idx)
End If

With SceneVes
ReDim TemList(0)
For i = 1 To .ChestCount
        If .Chests(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Chests(i)
        End If
Next i

.ChestCount = UBound(TemList)

If .ChestCount > 0 Then
ReDim .Chests(1 To .ChestCount)
     For i = 1 To .ChestCount
       .Chests(i) = TemList(i)
     Next i
End If
End With

If Scene_Idx <= -1 Then
   CurrentScene = SceneVes
Else
   Scenes(Scene_Idx) = SceneVes
End If
End Sub

'*************************************************************************
'**函 数 名：StructureSoundRes
'**输    入：(Long)Snd_Idx
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-03-20 23:25:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub StructureSoundRes(ByVal Snd_Idx As Long)
Dim i As Integer, TemList() As Type_ResourceInSound, SoundVes As Type_Sound

If Snd_Idx <= -1 Then
   SoundVes = CurrentSound
Else
   SoundVes = Sounds(Snd_Idx)
End If

With SoundVes
ReDim TemList(0)
For i = 1 To .ResourceCount
        If .Resource(i).ID > -1 Then
            ReDim Preserve TemList(UBound(TemList) + 1)
            TemList(UBound(TemList)) = .Resource(i)
        End If
Next i

.ResourceCount = UBound(TemList)

If .ResourceCount > 0 Then
ReDim .Resource(1 To .ResourceCount)
     For i = 1 To .ResourceCount
      .Resource(i) = TemList(i)
     Next i
End If
End With

If Snd_Idx <= -1 Then
   CurrentSound = SoundVes
Else
   Sounds(Snd_Idx) = SoundVes
End If
End Sub


'*************************************************************************
'**函 数 名：StructureFactionRelationShips
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-03-20 13:14:51
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub StructureFactionRelationShips(ByVal Faction_Idx As Long)
Dim i As Long, FacVes As Type_Faction, TemRels() As Type_RelationShip

If Faction_Idx > -1 Then
  FacVes = Factions(Faction_Idx)
Else
  FacVes = CurrentFaction
End If

ReDim TemRels(N_Faction - 1)

With FacVes
     For i = 0 To UBound(.RelationShip)
         .RelationShip(i).ID = GetID(.RelationShip(i).strID, False, "", -1)
         
         If .RelationShip(i).ID > -1 Then
            TemRels(.RelationShip(i).ID) = .RelationShip(i)
         End If
     Next i
     
     ReDim .RelationShip(N_Faction - 1)
     .RelationShip = TemRels
End With


If Faction_Idx > -1 Then
  Factions(Faction_Idx) = FacVes
Else
  CurrentFaction = FacVes
End If

End Sub


Public Sub PurseSceneInTroop(ByVal HexScene As String, SceneID As Long, Scene_strID As String, Entry As Long)
Dim strTem As String, lngTem As Long

       strTem = "0000" & Hex(HexScene)
       lngTem = Val("&H" & Right(strTem, 4))
       SceneID = lngTem
       Scene_strID = Scenes(lngTem).strID
       Entry = Val("&H" & Left(strTem, Len(strTem) - 4))
         
End Sub

'*************************************************************************
'**函 数 名：EnumEntities
'**输    入：(Integer)Tag
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-07-09 23:22:14
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function EnumEntities(Tag As Integer, strModule As String, oStr() As String) As Boolean
Dim i As Long

Select Case Tag
     Case 0
       EnumEntities = False
     Case Tag_Register
       EnumEntities = False
     Case Tag_Variable
       If UBound(TemGVarNameList.Triggers) >= 1 Then
         ReDim oStr(UBound(TemGVarNameList.Triggers(1).Checks))
         For i = 0 To UBound(TemGVarNameList.Triggers(1).Checks)
           With TemGVarNameList.Triggers(1)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .Checks(i))
             oStr(i) = Replace(oStr(i), "[csvname]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[csvname_pl]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[disname]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[disname_pl]", PYTags(Tag) & .Checks(i))
             'oStr(i) = Replace(oStr(i), "[value]", PYTags(Tag) & .Checks(i))
           End With
         Next i
         EnumEntities = True
       Else
         EnumEntities = False
       End If
         
     Case Tag_String
       EnumEntities = False   'A-E
     Case Tag_Item
       ReDim oStr(N_Item - 1)
       For i = 0 To N_Item - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", itm(i).dbName)
         oStr(i) = Replace(oStr(i), "[csvname]", itm(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", itm(i).csvName_pl)
         oStr(i) = Replace(oStr(i), "[disname]", itm(i).disname)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Item, i))
       Next i
       EnumEntities = True
     Case Tag_Troop
       ReDim oStr(N_Troop - 1)
       For i = 0 To N_Troop - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Trps(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Trps(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Trps(i).csvName_pl)
         oStr(i) = Replace(oStr(i), "[disname]", Trps(i).strName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Trps(i).strPtName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Troop, i))
       Next i
       EnumEntities = True
     Case Tag_Faction
       ReDim oStr(N_Faction - 1)
       For i = 0 To N_Faction - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Factions(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Factions(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Factions(i).csvName)
         oStr(i) = Replace(oStr(i), "[disname]", Factions(i).strName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Factions(i).strName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Faction, i))
       Next i
       EnumEntities = True
     Case Tag_Quest
       EnumEntities = False   'A-E
     Case Tag_Party_Tpl
       ReDim oStr(N_PT - 1)
       For i = 0 To N_PT - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", PTs(i).ptID)
         oStr(i) = Replace(oStr(i), "[csvname]", PTs(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", PTs(i).csvName)
         oStr(i) = Replace(oStr(i), "[disname]", PTs(i).ptName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", PTs(i).ptName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Party_Tpl, i))
       Next i
       EnumEntities = True
     Case Tag_Party
       ReDim oStr(N_Party - 1)
       For i = 0 To N_Party - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Parties(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Parties(i).csvName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Parties(i).csvName)
         oStr(i) = Replace(oStr(i), "[disname]", Parties(i).strName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Parties(i).strName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Party, i))
       Next i
       EnumEntities = True
     Case Tag_Scene
       ReDim oStr(N_Scene - 1)
       For i = 0 To N_Scene - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Scenes(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Scenes(i).strName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Scenes(i).strName)
         oStr(i) = Replace(oStr(i), "[disname]", Scenes(i).strName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Scenes(i).strName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Scene, i))
       Next i
       EnumEntities = True
     Case Tag_Mission_tpl
       EnumEntities = False   'A-E
     Case Tag_Menu
       EnumEntities = False   'A-E
     Case Tag_Script
       EnumEntities = False   'A-E
     Case Tag_Particle_Sys
       ReDim oStr(N_PSys - 1)
       For i = 0 To N_PSys - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", PSys(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", PSys(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", PSys(i).strID)
         oStr(i) = Replace(oStr(i), "[disname]", PSys(i).strID)
         oStr(i) = Replace(oStr(i), "[disname_pl]", PSys(i).strID)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Particle_Sys, i))
       Next i
       EnumEntities = True
     Case Tag_Scene_Prop
       EnumEntities = False   'A-E
     Case Tag_Sound
       ReDim oStr(N_Sound - 1)
       For i = 0 To N_Sound - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Sounds(i).sndName)
         oStr(i) = Replace(oStr(i), "[csvname]", Sounds(i).sndName)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Sounds(i).sndName)
         oStr(i) = Replace(oStr(i), "[disname]", Sounds(i).sndName)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Sounds(i).sndName)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Sound, i))
       Next i
       EnumEntities = True
     Case Tag_Local_Variable
       If UBound(CurVarNameList.Triggers) >= CheckListTrgIdx Then
         ReDim oStr(UBound(CurVarNameList.Triggers(CheckListTrgIdx).Checks))
         For i = 0 To UBound(CurVarNameList.Triggers(CheckListTrgIdx).Checks)
           With CurVarNameList.Triggers(CheckListTrgIdx)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .Checks(i))
             oStr(i) = Replace(oStr(i), "[csvname]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[csvname_pl]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[disname]", PYTags(Tag) & .Checks(i))
             oStr(i) = Replace(oStr(i), "[disname_pl]", PYTags(Tag) & .Checks(i))
             'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Local_Variable, i))
           End With
         Next i
         EnumEntities = True
       Else
         EnumEntities = False
       End If
     Case Tag_Map_Icon
       ReDim oStr(N_MapIcon - 1)
       For i = 0 To N_MapIcon - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", MapIcons(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", MapIcons(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", MapIcons(i).strID)
         oStr(i) = Replace(oStr(i), "[disname]", MapIcons(i).strID)
         oStr(i) = Replace(oStr(i), "[disname_pl]", MapIcons(i).strID)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Map_Icon, i))
       Next i
       EnumEntities = True
     Case Tag_Skill
       ReDim oStr(UBound(PublicSkills))
       For i = 0 To UBound(PublicSkills)
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", PublicSkills(i))
         oStr(i) = Replace(oStr(i), "[csvname]", PublicSkills(i))
         oStr(i) = Replace(oStr(i), "[csvname_pl]", PublicSkills(i))
         oStr(i) = Replace(oStr(i), "[disname]", PublicSkills(i))
         oStr(i) = Replace(oStr(i), "[disname_pl]", PublicSkills(i))
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Skill, i))
       Next i
       EnumEntities = True
     Case Tag_Mesh
       ReDim oStr(N_Mesh - 1)
       For i = 0 To N_Mesh - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", Mesh(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", Mesh(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", Mesh(i).strID)
         oStr(i) = Replace(oStr(i), "[disname]", Mesh(i).strID)
         oStr(i) = Replace(oStr(i), "[disname_pl]", Mesh(i).strID)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Mesh, i))
       Next i
       EnumEntities = True
     Case Tag_Presentation
       EnumEntities = False   'A-E
     Case Tag_Quick_String
       EnumEntities = False   'A-E
     Case Tag_Track
       EnumEntities = False   'A-E
     Case Tag_Tableau
       ReDim oStr(N_TabMat - 1)
       For i = 0 To N_TabMat - 1
         oStr(i) = strModule
         oStr(i) = Replace(oStr(i), "[index]", i)
         oStr(i) = Replace(oStr(i), "[dbname]", TabMat(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname]", TabMat(i).strID)
         oStr(i) = Replace(oStr(i), "[csvname_pl]", TabMat(i).strID)
         oStr(i) = Replace(oStr(i), "[disname]", TabMat(i).strID)
         oStr(i) = Replace(oStr(i), "[disname_pl]", TabMat(i).strID)
         'oStr(i) = Replace(oStr(i), "[value]", getTXTID(Tag_Tableau, i))
       Next i
       EnumEntities = True
     Case Tag_Animation
       EnumEntities = False   'A-E
     Case Tags_End
       EnumEntities = False
   End Select
End Function


'*************************************************************************
'**函 数 名：GetEntityName
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

Public Function GetEntityName(Tag As Integer, Index As Long, strModule As String) As String
Dim i As Long
i = Index
Select Case Tag
     Case 0
       GetEntityName = ""
     Case Tag_Register
       GetEntityName = ""
     Case Tag_Variable
       GetEntityName = ""
     Case Tag_String
       GetEntityName = ""   'A-E
     Case Tag_Item
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[dbname]", itm(i).dbName)
         GetEntityName = Replace(GetEntityName, "[csvname]", itm(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", itm(i).csvName_pl)
         GetEntityName = Replace(GetEntityName, "[disname]", itm(i).disname)

     Case Tag_Troop
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Trps(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", Trps(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Trps(i).csvName_pl)
         GetEntityName = Replace(GetEntityName, "[disname]", Trps(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Trps(i).strPtName)

     Case Tag_Faction
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Factions(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", Factions(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Factions(i).csvName)
         GetEntityName = Replace(GetEntityName, "[disname]", Factions(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Factions(i).strName)

     Case Tag_Quest
       GetEntityName = False   'A-E
     Case Tag_Party_Tpl
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", PTs(i).ptID)
         GetEntityName = Replace(GetEntityName, "[csvname]", PTs(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", PTs(i).csvName)
         GetEntityName = Replace(GetEntityName, "[disname]", PTs(i).ptName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", PTs(i).ptName)

     Case Tag_Party
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Parties(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", Parties(i).csvName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Parties(i).csvName)
         GetEntityName = Replace(GetEntityName, "[disname]", Parties(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Parties(i).strName)

     Case Tag_Scene
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Scenes(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", Scenes(i).strName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Scenes(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname]", Scenes(i).strName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Scenes(i).strName)

     Case Tag_Mission_tpl
       GetEntityName = ""   'A-E
     Case Tag_Menu
       GetEntityName = ""   'A-E
     Case Tag_Script
       GetEntityName = ""   'A-E
     Case Tag_Particle_Sys
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", PSys(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", PSys(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", PSys(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname]", PSys(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", PSys(i).strID)

     Case Tag_Scene_Prop
       GetEntityName = ""   'A-E
     Case Tag_Sound
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", Sounds(i).sndName)
         GetEntityName = Replace(GetEntityName, "[csvname]", Sounds(i).sndName)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", Sounds(i).sndName)
         GetEntityName = Replace(GetEntityName, "[disname]", Sounds(i).sndName)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", Sounds(i).sndName)

     Case Tag_Local_Variable
       GetEntityName = ""
     Case Tag_Map_Icon
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", MapIcons(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", MapIcons(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", MapIcons(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname]", MapIcons(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", MapIcons(i).strID)

     Case Tag_Skill
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", PublicSkills(i))
         GetEntityName = Replace(GetEntityName, "[csvname]", PublicSkills(i))
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", PublicSkills(i))
         GetEntityName = Replace(GetEntityName, "[disname]", PublicSkills(i))
         GetEntityName = Replace(GetEntityName, "[disname_pl]", PublicSkills(i))

     Case Tag_Mesh
       GetEntityName = ""   'A-E
     Case Tag_Presentation
       GetEntityName = ""   'A-E
     Case Tag_Quick_String
       GetEntityName = ""   'A-E
     Case Tag_Track
       GetEntityName = ""   'A-E
     Case Tag_Tableau
         GetEntityName = strModule
         GetEntityName = Replace(GetEntityName, "[index]", i)
         GetEntityName = Replace(GetEntityName, "[dbname]", TabMat(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname]", TabMat(i).strID)
         GetEntityName = Replace(GetEntityName, "[csvname_pl]", TabMat(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname]", TabMat(i).strID)
         GetEntityName = Replace(GetEntityName, "[disname_pl]", TabMat(i).strID)

     Case Tag_Animation
       GetEntityName = ""   'A-E
     Case Tags_End
       GetEntityName = ""
   End Select
End Function

'*************************************************************************
'**函 数 名：Max_Int
'**输    入：(Int)a,(Int)b
'**输    出：(Integer) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-07-20 20:49:52
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function Max_Int(a As Integer, b As Integer) As Integer
Max_Int = IIf(a >= b, a, b)
End Function


'*************************************************************************
'**函 数 名：EnumConsts
'**输    入：(String)ParaType
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-07-09 23:22:14
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function EnumConsts(ParaType As String, strModule As String, oStr() As String) As Boolean
Dim i As Long

Select Case ParaType     'ends_add
     Case ""
       EnumConsts = False
     Case "pos"
       EnumConsts = False
     Case "s"
       EnumConsts = False
     Case "itp"
         ReDim oStr(1 To UBound(Item_Type))
         For i = 1 To UBound(Item_Type)
           With Item_Type(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "tf"
         ReDim oStr(0 To UBound(Tf))
         For i = 0 To UBound(Tf)
           With Tf(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .strName)
             oStr(i) = Replace(oStr(i), "[csvname]", .csvName)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .csvName)
             oStr(i) = Replace(oStr(i), "[disname]", .strName)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .strName)
             oStr(i) = Replace(oStr(i), "[value]", I64toStrNZ(.Value))
           End With
         Next i
         EnumConsts = True
     Case "pf"
         ReDim oStr(0 To UBound(Pf))
         For i = 0 To UBound(Pf)
           With Pf(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .strName)
             oStr(i) = Replace(oStr(i), "[csvname]", .csvName)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .csvName)
             oStr(i) = Replace(oStr(i), "[disname]", .strName)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .strName)
             oStr(i) = Replace(oStr(i), "[value]", I64toStrNZ(.Value))
           End With
         Next i
         EnumConsts = True
     Case "bs"
         ReDim oStr(0 To UBound(BoolSwitch))
         For i = 0 To UBound(BoolSwitch)
           With BoolSwitch(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "ap"
         ReDim oStr(0 To UBound(AccessPrivilege))
         For i = 0 To UBound(AccessPrivilege)
           With AccessPrivilege(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "as"
         ReDim oStr(0 To UBound(AbsSwitch))
         For i = 0 To UBound(AbsSwitch)
           With AbsSwitch(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "ai_bhvr"
         ReDim oStr(0 To UBound(AI_Bhvr))
         For i = 0 To UBound(AI_Bhvr)
           With AI_Bhvr(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
     Case "po"
         ReDim oStr(0 To UBound(PlayOption))
         For i = 0 To UBound(PlayOption)
           With PlayOption(i)
             oStr(i) = strModule
             oStr(i) = Replace(oStr(i), "[index]", i)
             oStr(i) = Replace(oStr(i), "[dbname]", .X)
             oStr(i) = Replace(oStr(i), "[csvname]", .Y)
             oStr(i) = Replace(oStr(i), "[csvname_pl]", .Y)
             oStr(i) = Replace(oStr(i), "[disname]", .X)
             oStr(i) = Replace(oStr(i), "[disname_pl]", .X)
             oStr(i) = Replace(oStr(i), "[value]", i)
           End With
         Next i
         EnumConsts = True
    Case Else
      EnumConsts = False
   End Select
End Function
'*************************************************************************
'**函 数 名：CancelTopForms
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-05 11:29:25
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub CancelTopForms()
  If IsLoadString Then
     SetWindowPos frmStrTool.hWnd, -2, 0, 0, 0, 0, 3
  End If
End Sub
'*************************************************************************
'**函 数 名：SetTopForms
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-03-05 11:31:35
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub SetTopForms()
  If IsLoadString Then
     SetWindowPos frmStrTool.hWnd, -1, 0, 0, 0, 0, 3
  End If
End Sub

