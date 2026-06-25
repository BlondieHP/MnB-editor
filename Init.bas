Attribute VB_Name = "Init"

Option Explicit

Sub InitTags()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    Tags(Tag_Register) = "reg"  '�Զ���tag
    Tags(Tag_Variable) = "var"  '�Զ���tag
    Tags(Tag_String) = "str"
    Tags(Tag_Item) = "itm"
    Tags(Tag_Troop) = "trp"
    Tags(Tag_Faction) = "fac"
    Tags(Tag_Quest) = "qst"
    Tags(Tag_Party_Tpl) = "pt"
    Tags(Tag_Party) = "p"
    Tags(Tag_Scene) = "scn"
    Tags(Tag_Mission_tpl) = "mst"
    Tags(Tag_Menu) = "menu"
    Tags(Tag_Script) = "script"   '�Զ���tag
    Tags(Tag_Particle_Sys) = "psys"
    Tags(Tag_Scene_Prop) = "spr"
    Tags(Tag_Sound) = "snd"
    Tags(Tag_Local_Variable) = "lvar"  '�Զ���tag
    Tags(Tag_Map_Icon) = "icon"     '�Զ���tag
    Tags(Tag_Skill) = "skl"
    Tags(Tag_Mesh) = "mesh"
    Tags(Tag_Presentation) = "prsnt"
    Tags(Tag_Quick_String) = "qstr"
    Tags(Tag_Track) = "track"     '�Զ���tag
    Tags(Tag_Tableau) = "tab"
    Tags(Tag_Animation) = "anim"   '�Զ���tag
    Tags(Tags_End) = "end"

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("Init", "InitTags", Err.Number, Err.Description)
End Sub

'*************************************************************************
' InitShortTags - Initialize 64-bit game ID encoding parameters
'
' Each entity type (Troop, Item, Faction, etc.) has a unique ID space in
' the game's 64-bit identifier system. The ShortTags define how entity
' indices are encoded into game IDs via the formula:
'
'   game_id = X + index  (lower 32 bits)
'   game_id_high = Y     (upper 32 bits)
'
' X values: base offsets incremented by 720-721 per tag type
' Y values: high-word markers, starting from ~7927936 for registers
'           and decreasing by ~2072192 per subsequent tag type
'
' These values are used by getTXTID() and GetstrID() to convert between
' internal indices and the game's 64-bit string identifiers.
'
' Author: Ser_Charles
' Date: 2011-02-07
' Version: V1.132.1
'*************************************************************************

Sub InitShortTags()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    ShortTags(Tag_Register).X = 720
    ShortTags(Tag_Variable).X = 1441
    ShortTags(Tag_String).X = 2161
    ShortTags(Tag_Item).X = 2882
    ShortTags(Tag_Troop).X = 3602
    ShortTags(Tag_Faction).X = 4323
    ShortTags(Tag_Quest).X = 5044
    ShortTags(Tag_Party_Tpl).X = 5764
    ShortTags(Tag_Party).X = 6485
    ShortTags(Tag_Scene).X = 7205
    ShortTags(Tag_Mission_tpl).X = 7926
    ShortTags(Tag_Menu).X = 8646
    ShortTags(Tag_Script).X = 9367
    ShortTags(Tag_Particle_Sys).X = 10088
    ShortTags(Tag_Scene_Prop).X = 10808
    ShortTags(Tag_Sound).X = 11529
    ShortTags(Tag_Local_Variable).X = 12249
    ShortTags(Tag_Map_Icon).X = 12970
    ShortTags(Tag_Skill).X = 13690
    ShortTags(Tag_Mesh).X = 14411
    ShortTags(Tag_Presentation).X = 15132
    ShortTags(Tag_Quick_String).X = 15852
    ShortTags(Tag_Track).X = 16573
    ShortTags(Tag_Tableau).X = 17293
    ShortTags(Tag_Animation).X = 18014
    ShortTags(Tags_End).X = 18734

    ShortTags(Tag_Register).Y = 7927936
    ShortTags(Tag_Variable).Y = 5855872
    ShortTags(Tag_String).Y = 3783808
    ShortTags(Tag_Item).Y = 1711744
    ShortTags(Tag_Troop).Y = 9639680
    ShortTags(Tag_Faction).Y = 7567616
    ShortTags(Tag_Quest).Y = 5495552
    ShortTags(Tag_Party_Tpl).Y = 3423488
    ShortTags(Tag_Party).Y = 1351424
    ShortTags(Tag_Scene).Y = 9279360
    ShortTags(Tag_Mission_tpl).Y = 7207296
    ShortTags(Tag_Menu).Y = 5135232
    ShortTags(Tag_Script).Y = 3063168
    ShortTags(Tag_Particle_Sys).Y = 991104
    ShortTags(Tag_Scene_Prop).Y = 8919040
    ShortTags(Tag_Sound).Y = 6846976
    ShortTags(Tag_Local_Variable).Y = 4774912
    ShortTags(Tag_Map_Icon).Y = 2702848
    ShortTags(Tag_Skill).Y = 630784
    ShortTags(Tag_Mesh).Y = 8558720
    ShortTags(Tag_Presentation).Y = 6486656
    ShortTags(Tag_Quick_String).Y = 4414592
    ShortTags(Tag_Track).Y = 2342528
    ShortTags(Tag_Tableau).Y = 270464
    ShortTags(Tag_Animation).Y = 8198400
    ShortTags(Tags_End).Y = 6126336
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("Init", "InitShortshorttags", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**�� �� ����inititmpicker
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-06-14 06:43:27
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.955.6
'*************************************************************************

Sub InitItmPicker()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim iii As Long
    Dim Scrollwidth As Long
    Dim res As Long
    Dim strTmp As String

    Scrollwidth = 0

    'Form1.itmpicker.Clear
    For iii = 0 To N_Item - 1
        If Len(itm(iii).csvName) > 0 Then
            strTmp = CStr(iii) & " " & itm(iii).csvName & " " & itm(iii).disname
        Else
            strTmp = CStr(iii) & " " & itm(iii).disname & " " & itm(iii).csvName
        End If

        If Len(strTmp) > Scrollwidth Then
            Scrollwidth = Len(strTmp)
        End If
        'Form1.itmpicker.AddItem (strTmp)
    Next iii

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("Init", "InitItmPicker", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����initItmTypeList
'**��    �룺��
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:14
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************

Sub initItmTypeList(objListBox As ListBox)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    objListBox.Clear

    Call ItmTypeListAddItem(objListBox, 0, "ȫ��")
    Call ItmTypeListAddItem(objListBox, itp_type_horse, "��ƥ")
    Call ItmTypeListAddItem(objListBox, itp_type_one_handed_wpn, "��������")
    Call ItmTypeListAddItem(objListBox, itp_type_two_handed_wpn, "˫������")
    Call ItmTypeListAddItem(objListBox, itp_type_polearm, "��������")
    Call ItmTypeListAddItem(objListBox, itp_type_arrows, "��")
    Call ItmTypeListAddItem(objListBox, itp_type_bolts, "���")
    Call ItmTypeListAddItem(objListBox, itp_type_shield, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_bow, "��")
    Call ItmTypeListAddItem(objListBox, itp_type_crossbow, "��")
    Call ItmTypeListAddItem(objListBox, itp_type_thrown, "Ͷ������")
    Call ItmTypeListAddItem(objListBox, itp_type_goods, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_head_armor, "ͷ��")
    Call ItmTypeListAddItem(objListBox, itp_type_body_armor, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_foot_armor, "Ь��")
    Call ItmTypeListAddItem(objListBox, itp_type_hand_armor, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_pistol, "��ǹ")
    Call ItmTypeListAddItem(objListBox, itp_type_musket, "��ǹ")
    Call ItmTypeListAddItem(objListBox, itp_type_bullets, "�ӵ�")
    Call ItmTypeListAddItem(objListBox, itp_type_animal, "����")
    Call ItmTypeListAddItem(objListBox, itp_type_book, "�鼮")

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("Init", "initItmTypeList", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����ItmTypeListAddItem
'**��    �룺itp_type_code(Byte)   -
'**        ��itp_type_name(String) -
'**��    ������
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:17
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************

Sub ItmTypeListAddItem(objListBox As ListBox, itp_type_code As Byte, itp_type_name As String)
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    Dim tmpStr As String
    tmpStr = TranslateStr("Form3_typeopt", CStr(itp_type_code), itp_type_name)
    objListBox.AddItem (CStr(itp_type_code) & " " & tmpStr)

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("Init", "ItmTypeListAddItem", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����showHexId
'**��    �룺(String)id
'**��    ����(String)
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�kevin
'**��    �ڣ�2008-05-18 08:24:30
'**�� �� �ˣ�SSgt_Edward
'**��    �ڣ�2010-12-13 11:11:15
'**��    ����V1.1321
'*************************************************************************

Sub PublicInit()
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
    
    CurrentTrpID = -1
    CurPartyTemplateID = -1
    CurrentItmID = -1
    CurrentPartyID = -1
    CurrentSceneID = -1
    CurrentFactionID = -1
    CurrentMapIconID = -1
    CurrentPSysID = -1
    
    ReDim Trps(0)
    ReDim itm(0)
    ReDim Scenes(0)
    ReDim Parties(0)
    ReDim PTs(0)
    ReDim Factions(0)
    ReDim MapIcons(0)
    ReDim PSys(0)
    
    'Dim i As Integer
    
    'For i = 0 To 99
    '    DreamTeam(i) = -1
    'Next i
    
    Init_Integer64b
    
    RunERR.MissCSV = False
    RunERR.MissINI = False
    RunERR.MissMod = False
        
    'frmData.Show
    InitTags
    InitPYTags
    InitShortTags
    InitPy
    
    InitPublicWords
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("Init", "PublicInit", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**�� �� ����GetTroopSkill
'**��    �룺TroopID(integer,-1��Ϊ��ǰ����), ByVal sklid(Integer) -
'**��    ����(Byte) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�Ser_Charles
'**��    �ڣ�2011-04-17 20:06:12
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V0.951.12
'*************************************************************************

Public Function InitWarbandInfo(Optional MBHome As String = "", Optional Version As String = "") As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------

    Dim strTmp As String
    Dim i As Integer, s As Long
    
    InitWarbandInfo = False
    
    DisplayMode = 0
    
    MnBInfo.InfoFinished = False
    MnBInfo.InitFinished = False
    MnBInfo.iniSetting = App.Path & "\" & App.EXEName & ".ini"
    MnBInfo.Language = "en"
    If Trim(MBHome) = "" Then
      strTmp = ReadRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\Mount&Blade Warband", "")
    Else
      strTmp = MBHome
    End If
    
    ' fix me
    strTmp = FixPath(strTmp)
    If Not DirExists(strTmp) Then
         MsgBox "���ĵ�����û�а�װ�������뿳ɱ:ս�š�!", vbCritical, "Error!"
         Exit Function
    End If
With MnBInfo
     .MBHome = strTmp
     .ModPath = strTmp & "\Modules"
     
     If Trim(Version) = "" Then
       .Version = ReadRegString(HKEY_LOCAL_MACHINE, "SOFTWARE\Mount&Blade Warband", "Version")
     Else
       .Version = Version
     End If
     
     strTmp = GetMyDocumentDirectory()
     strTmp = FixPath(strTmp)
     
     
     .MBsaves = strTmp & "\Mount&Blade Warband Savegames"
     .MBsets = strTmp & "\Mount&Blade Warband"
     If FileExists(.MBsets & "\language.txt") Then
            .Language = QuickReadFile(.MBsets & "\language.txt")
     Else
            strTmp = GetAppDataPath
            strTmp = FixPath(strTmp)
            .MBsets = strTmp & "\Mount&Blade Warband"
            .Language = QuickReadFile(.MBsets & "\language.txt")
     End If
     
     .Language = Trim(.Language)
     If .Language = "" Then
         .Language = "en"
     End If
     
     strTmp = Trim(ReadString(.iniSetting, "Settings", "Language", 250))
     SelectLanguage strTmp
     
     LoadOperations
     
     .InitFinished = True
End With
    InitWarbandInfo = True
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Init", "InitWarbandInfo", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����FinishWarbandInfo
'**��    �룺(String)ModName
'**��    ����(Boolean) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-22 22:20:40
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************

Public Function FinishWarbandInfo(ByVal ModName As String) As Boolean
    On Error GoTo errorHandle '�򿪴�������
    '------------------------------------------------
Dim strTem As String, i As Integer
With MnBInfo
     .ModPath = .MBHome & "\Modules\" & ModName
     .ModBackUp = .ModPath & "\BackUp"
     .ModName = ModName
     .ModIniFileName = .ModPath & "\module.ini"
     .iniFileName = .ModPath & "\MnBWarband_Editor.ini"
     .FirstTimeEdit = False
     
     If FileExists(.iniFileName) Then
       
        For i = 0 To UBound(.EditInfo)
           .EditInfo(i) = ReadInt(.iniFileName, "EDITINFO", "Count" & i)
           
           If .EditInfo(i) = 0 Then
             .FirstTimeEdit = True
             Exit For
           End If
        Next i
       
     Else
       .FirstTimeEdit = True
     End If
     
     LoadModINI
     
     'frmData.Show
     .InfoFinished = True
End With
    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("Init", "FinishWarbandInfo", Err.Number, Err.Description)
End Function

'*************************************************************************
'**�� �� ����QuickReadFile
'**��    �룺(String)FilePath
'**��    ����(String) -
'**����������
'**ȫ�ֱ�����
'**����ģ�飺
'**��    �ߣ�SSgt_Edward
'**��    �ڣ�2010-11-22 23:26:26
'**�� �� �ˣ�
'**��    �ڣ�
'**��    ����V1.1321
'*************************************************************************

