Attribute VB_Name = "FileLoader"

Option Explicit

Sub LoadTroopFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim str_tmp1 As String
    Dim str_tmp2 As String
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    MaxPointer = FileLen(tmpFileName)
    
    lngHandle = FreeFile()

    Open tmpFileName For Random Access Read As lngHandle Len = 1
    Pointer = 1
    For n = 0 To 2
        TrpsVersionInform(n) = GetWord()
    Next n
    N_Troop = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_TroopsCount) = N_Troop
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_TroopsCount, CStr(N_Troop)
    End If
    
    ReDim Trps(N_Troop - 1)
    
    DoEvents
    For n = 0 To N_Troop - 1

        With Trps(n)
            .ID = n
            .strID = GetWord()
            .strName = GetWord()
            .strPtName = GetWord()
            .csvName = .strName 'default
            .csvName_pl = .strPtName 'default
            .unknown_warband(1) = GetWord()
            
            .Flags = GetWord()
            .Scene = CLng(Val(GetWord()))
            .reserved = Val(GetWord())
            
            .Faction = GetWord()
            
            .Upgrade1 = GetWord()
            .Upgrade2 = GetWord()
            
            For m = 1 To 64
                str_tmp1 = GetWord()
                str_tmp2 = GetWord()
                .lstInventory(m).X = Val(str_tmp1)
                .lstInventory(m).Y = Val(str_tmp2)
            Next m
            
            .tAttrib.strPoint = Val(GetWord())
            .tAttrib.agiPoint = Val(GetWord())
            .tAttrib.intPoint = Val(GetWord())
            .tAttrib.chaPoint = Val(GetWord())
            .tAttrib.level = Val(GetWord())
            
            .WP.one_handed = GetWord()
            .WP.two_handed = GetWord()
            .WP.polearm = GetWord()
            .WP.archery = GetWord()
            .WP.crossbow = GetWord()
            .WP.throwing = GetWord()
            .WP.firearm = GetWord()
            
            .Skills(1) = Val(GetWord())
            .Skills(2) = Val(GetWord())
            .Skills(3) = Val(GetWord())
            .Skills(4) = Val(GetWord())
            .Skills(5) = Val(GetWord())
            .Skills(6) = Val(GetWord())
            
            .Face(1) = GetWord()
            .Face(2) = GetWord()
            .Face(3) = GetWord()
            .Face(4) = GetWord()
            .Face(5) = GetWord()
            .Face(6) = GetWord()
            .Face(7) = GetWord()
            .Face(8) = GetWord()
            
            .Edit = CheckEditable(EditInfo_TroopsCount, n)
            
            AddIndex .ID, .strID
            
        End With
    Next
    Close lngHandle
    Pointer = 1
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadTroopFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：ReadTroopLine
'**输    入：Text(String),OutputTroop(Type_Troops)
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 16:24:35
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadTroopLine(Text As String, OutputTroop As Type_Troops)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim str_tmp1 As String
    Dim str_tmp2 As String
    Dim n As Long
    Dim m As Integer
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
        With OutputTroop
            .strID = GetWordL()
            .strName = GetWordL()
            .strPtName = GetWordL()
            .unknown_warband(1) = GetWordL()
            
            .Flags = GetWordL()
            
            .Scene = CLng(Val(GetWordL()))
            PurseSceneInTroop .Scene, .SceneID, .Scene_strID, .Entry
            
            .reserved = Val(GetWordL())
            
            .Faction = GetWordL()
            .Faction_strID = Factions(.Faction).strID
            
            .Upgrade1 = GetWordL()
            .Upgrade1_strID = Trps(.Upgrade1).strID
            .Upgrade2 = GetWordL()
            .Upgrade2_strID = Trps(.Upgrade2).strID
            
            For m = 1 To 64
                str_tmp1 = GetWordL()
                str_tmp2 = GetWordL()
                .lstInventory(m).X = Val(str_tmp1)
                If .lstInventory(m).X > -1 Then
                  .lstInventory(m).strX = itm(.lstInventory(m).X).dbName
                Else
                  .lstInventory(m).strX = ""
                End If
                .lstInventory(m).Y = Val(str_tmp2)
            Next m
            
            .tAttrib.strPoint = Val(GetWordL())
            .tAttrib.agiPoint = Val(GetWordL())
            .tAttrib.intPoint = Val(GetWordL())
            .tAttrib.chaPoint = Val(GetWordL())
            .tAttrib.level = Val(GetWordL())
            
            .WP.one_handed = GetWordL()
            .WP.two_handed = GetWordL()
            .WP.polearm = GetWordL()
            .WP.archery = GetWordL()
            .WP.crossbow = GetWordL()
            .WP.throwing = GetWordL()
            .WP.firearm = GetWordL()
            
            .Skills(1) = Val(GetWordL())
            .Skills(2) = Val(GetWordL())
            .Skills(3) = Val(GetWordL())
            .Skills(4) = Val(GetWordL())
            .Skills(5) = Val(GetWordL())
            .Skills(6) = Val(GetWordL())
            
            .Face(1) = GetWordL()
            .Face(2) = GetWordL()
            .Face(3) = GetWordL()
            .Face(4) = GetWordL()
            .Face(5) = GetWordL()
            .Face(6) = GetWordL()
            .Face(7) = GetWordL()
            .Face(8) = GetWordL()
            
        End With
    
    LinePointer = 1
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadTroopLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadFactionFile
'**输    入：filePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-24 23:03:54
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadFactionFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    MaxPointer = FileLen(tmpFileName)
    
    lngHandle = FreeFile()

    Open tmpFileName For Random Access Read As lngHandle Len = 1
    Pointer = 1
    For n = 0 To 2
        FactionVersionInform(n) = GetWord()
    Next n
    N_Faction = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_FactionsCount) = N_Faction
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_FactionsCount, CStr(N_Faction)
    End If
    
    ReDim Factions(N_Faction - 1)
    
    DoEvents
    For n = 0 To N_Faction - 1
         With Factions(n)
               .ID = n
               .strID = GetWord()
               .strName = GetWord()
               .csvName = .strName 'Default
               .Flags = Val(GetWord())
               .lColor = GetWord()
               
               ReDim Factions(n).RelationShip(N_Faction - 1)
                
               For m = 0 To N_Faction - 1
                     .RelationShip(m).Value = Val(GetWord())
               Next m
               .reserved = Val(GetWord())
               
               .Edit = CheckEditable(EditInfo_FactionsCount, n)
               
               AddIndex .ID, .strID
         End With
    Next
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadFactionFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：ReadFactionLine
'**输    入：Text(String),OutputFaction(Type_Faction)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 16:40:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadFactionLine(Text As String, OutputFaction As Type_Faction)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer
    Dim head As String
    
    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
         With OutputFaction
         
               .strID = GetWordL()
               .strName = GetWordL()
               .Flags = Val(GetWordL())
               .lColor = GetWordL()
               
               ReDim Factions(n).RelationShip(N_Faction - 1)
                
               For m = 0 To N_Faction - 1
                     .RelationShip(m).ID = m
                     .RelationShip(m).strID = Factions(m).strID
                     .RelationShip(m).Value = Val(GetWordL())
               Next m
               .reserved = Val(GetWordL())

         End With

    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadFactionLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadSceneFile
'**输    入：filePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-28 23:07:51
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadSceneFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    MaxPointer = FileLen(tmpFileName)
    
    lngHandle = FreeFile()

    Open tmpFileName For Random Access Read As lngHandle Len = 1
    Pointer = 1
    For n = 0 To 2
        SceneVersionInform(n) = GetWord()
    Next n
    N_Scene = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_ScenesCount) = N_Scene
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_ScenesCount, CStr(N_Scene)
    End If
    
    ReDim Scenes(N_Scene - 1)
    
    DoEvents
    For n = 0 To N_Scene - 1
         With Scenes(n)
               .ID = n
               .strID = GetWord()
               .strName = GetWord()
               .Flags = GetWord()
               .MeshName = GetWord()
               .BodyName = GetWord()
               
               .p(0).X = Val(GetWord())
               .p(0).Y = Val(GetWord())
               .p(1).X = Val(GetWord())
               .p(1).Y = Val(GetWord())
               
               .WaterLevel = Val(GetWord())
               .TerrainCode = GetWord()
               
               .AccessCount = Val(GetWord())
               
               If .AccessCount > 0 Then
                 ReDim .Accesses(1 To .AccessCount)
               
                 For m = 1 To .AccessCount
                  .Accesses(m) = GetWord()
                 Next m
               Else
                 ReDim .Accesses(0)
               End If
               
               .ChestCount = Val(GetWord())
               
               If .ChestCount > 0 Then
                 ReDim .Chests(1 To .ChestCount)
               
                 For m = 1 To .ChestCount
                  .Chests(m).ID = Val(GetWord())
                 Next m
               Else
                 ReDim .Chests(0)
               End If
               
               .Outer_Terrain_Type = GetWord()
               
               .Edit = CheckEditable(EditInfo_ScenesCount, n)
               
               AddIndex .ID, .strID
         End With
    Next
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadSceneFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：ReadSceneLine
'**输    入：Text(String),OutputScene(Type_Scene)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 16:59:36
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadSceneLine(Text As String, OutputScene As Type_Scene)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer
    Dim head As String
    
    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
         With OutputScene
               .strID = GetWordL()
               .strName = GetWordL()
               .Flags = GetWordL()
               .MeshName = GetWordL()
               .BodyName = GetWordL()
               
               .p(0).X = Val(GetWordL())
               .p(0).Y = Val(GetWordL())
               .p(1).X = Val(GetWordL())
               .p(1).Y = Val(GetWordL())
               
               .WaterLevel = Val(GetWordL())
               .TerrainCode = GetWordL()
               
               .AccessCount = Val(GetWordL())
               
               If .AccessCount > 0 Then
                 ReDim .Accesses(1 To .AccessCount)
               
                 For m = 1 To .AccessCount
                  .Accesses(m) = GetWordL()
                 Next m
               Else
                 ReDim .Accesses(0)
               End If
               
               .ChestCount = Val(GetWordL())
               
               If .ChestCount > 0 Then
                 ReDim .Chests(1 To .ChestCount)
               
                 For m = 1 To .ChestCount
                  .Chests(m).ID = Val(GetWordL())
                  .Chests(m).strID = Trps(.Chests(m).ID).strID
                 Next m
               Else
                 ReDim .Chests(0)
               End If
               
               .Outer_Terrain_Type = GetWordL()

         End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadSceneLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveTroopFile
'**输    入：filePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:23:45
'**修 改 人：SSgt_Edward
'**日    期：2010-11-25 16:30:41
'**版    本：V1.1321
'*************************************************************************

Sub LoadPTFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
        
    Dim tmp1 As String
    Dim n As Long
    Dim m As Integer
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件:" & tmpFileName
        Exit Sub
    End If
    
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As #lngHandle Len = 1
    Pointer = 1
    MaxPointer = FileLen(tmpFileName)
    For n = 0 To 2
        PTVersionInform(n) = GetWord()
    Next n
    N_PT = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_PartyTemplatesCount) = N_PT
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_PartyTemplatesCount, CStr(N_PT)
    End If
    
    ReDim PTs(N_PT - 1)
    
    DoEvents
    For n = 0 To N_PT - 1
        
        With PTs(n)
            .ID = n
            .ptID = GetWord()
            .ptName = GetWord()
            .Flags = GetWord()
            .Menu = GetWord()
            .Faction = Val(GetWord())
            .Personality = Val(GetWord())
            
            .csvName = .ptName  'default
            
            For m = 1 To 6
                .Stacks(m).ID = Val(GetWord())
                
                If .Stacks(m).ID < 0 Then
                    'For m1 = m To 6
                    '   .stacks(m1).id = -1
                    '  .stacks(m1).Min = -1
                    '  .stacks(m1).Max = -1
                    '  .stacks(m1).flags = ""
                    'Next
                    'm = m + 10
                Else
                    .Stacks(m).Min = Val(GetWord())
                    .Stacks(m).Max = Val(GetWord())
                    .Stacks(m).Flags = Val(GetWord())
                End If
                
            Next
            
            .Edit = CheckEditable(EditInfo_PartyTemplatesCount, n)
            
            AddIndex .ID, .ptID
        End With
        
    Next n
    
    Close #lngHandle
    Pointer = 1

    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadPTFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：ReadPTLine
'**输    入：Text(String),(Type_PT)OutputPT
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 17:02:26
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadPTLine(Text As String, OutputPT As Type_PT)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
        
    Dim tmp1 As String
    Dim n As Long
    Dim m As Integer
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    head = GetWordL()
    
        With OutputPT
            .ptID = GetWordL()
            .ptName = GetWordL()
            .Flags = GetWordL()
            .Menu = GetWordL()
            
            .Faction = Val(GetWordL())
            .Faction_strID = Factions(.Faction).strID
            
            .Personality = Val(GetWordL())

            For m = 1 To 6
                .Stacks(m).ID = Val(GetWordL())
                
                If .Stacks(m).ID < 0 Then
                    'For m1 = m To 6
                    '   .stacks(m1).id = -1
                    '  .stacks(m1).Min = -1
                    '  .stacks(m1).Max = -1
                    '  .stacks(m1).flags = ""
                    'Next
                    'm = m + 10
                Else
                    .Stacks(m).strID = Trps(.Stacks(m).ID).strID
                    .Stacks(m).Min = Val(GetWordL())
                    .Stacks(m).Max = Val(GetWordL())
                    .Stacks(m).Flags = Val(GetWordL())
                End If
                
            Next m
            
        End With

    LinePointer = 1

    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadPTLine", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**函 数 名：SavePartyTemplateCSVFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-07 15:44:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadPartyFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
        
    Dim tmp1 As String
    Dim n As Long
    Dim m As Integer
    Dim tmpFileName As String
    tmpFileName = FilePath
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件:" & tmpFileName
        Exit Sub
    End If
    
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As #lngHandle Len = 1
    Pointer = 1
    MaxPointer = FileLen(tmpFileName)
    For n = 0 To 2
        PartyVersionInform(n) = GetWord()
    Next n
    N_Party = Val(GetWord())
    N_Party2 = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_PartiesCount) = N_Party
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_PartiesCount, CStr(N_Party)
    End If
    
    ReDim Parties(N_Party - 1)
    
    DoEvents
    For n = 0 To N_Party - 1
        
        With Parties(n)
            .UnknownTitle = Val(GetWord())
            .ID = Val(GetWord())
            .id2 = Val(GetWord())
            .strID = GetWord()
            .strName = GetWord()
            .Flags = GetWord()
            .Menu = GetWord()
            .Template = Val(GetWord())
            .Faction = Val(GetWord())
            .Personality(1) = Val(GetWord())
            .Personality(2) = Val(GetWord())
            .AI_Behavior = GetWord()
            .AI_Target = Val(GetWord())
            .reserved = Val(GetWord())
            
            .csvName = .strName  'default
            
            For m = 1 To 3
               .InitPos(m).X = GetWord()
               .InitPos(m).Y = GetWord()
            Next m
            
            .UnknownStr = GetWord()
            
            '成员数据
            .StacksCount = Val(GetWord())
            If .StacksCount > 0 Then
               ReDim .Stacks(1 To .StacksCount)
               
               For m = 1 To .StacksCount
                  .Stacks(m).ID = Val(GetWord())
                  .Stacks(m).Min = Val(GetWord())
                  .Stacks(m).Max = Val(GetWord())
                  .Stacks(m).Flags = Val(GetWord())
               Next m
            Else
               ReDim .Stacks(0)
            End If
            
            .Degree = GetWord()
            
            .Edit = CheckEditable(EditInfo_PartiesCount, n)
            
            AddIndex .ID, .strID
        End With
        
    Next n
    
    Close #lngHandle
    Pointer = 1

    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadPartyFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：ReadPartyLine
'**输    入：Text(String),OutputParty(Type_Party)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 16:44:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadPartyLine(Text As String, OutputParty As Type_Party)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
        
    Dim tmp1 As String
    Dim n As Long
    Dim m As Integer
    Dim head As String
    
    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
        With OutputParty
            .UnknownTitle = Val(GetWordL())
            .ID = Val(GetWordL())
            .id2 = Val(GetWordL())
            .strID = GetWordL()
            .strName = GetWordL()
            .Flags = GetWordL()
            .Menu = GetWordL()
            
            .Template = Val(GetWordL())
            .Template_strID = PTs(.Template).ptID
            
            .Faction = Val(GetWordL())
            .Faction_strID = Factions(.Faction).strID
            
            .Personality(1) = Val(GetWordL())
            .Personality(2) = Val(GetWordL())
            .AI_Behavior = GetWordL()
            
            .AI_Target = Val(GetWordL())
            .AI_Target_strID = Parties(.AI_Target).strID
            
            .reserved = Val(GetWordL())
            
            For m = 1 To 3
               .InitPos(m).X = GetWordL()
               .InitPos(m).Y = GetWordL()
            Next m
            
            .UnknownStr = GetWordL()
            
            '成员数据
            .StacksCount = Val(GetWordL())
            If .StacksCount > 0 Then
               ReDim .Stacks(1 To .StacksCount)
               
               For m = 1 To .StacksCount
                  .Stacks(m).ID = Val(GetWordL())
                  .Stacks(m).strID = Trps(.Stacks(m).ID).strID
                  .Stacks(m).Min = Val(GetWordL())
                  .Stacks(m).Max = Val(GetWordL())
                  .Stacks(m).Flags = Val(GetWordL())
               Next m
            Else
               ReDim .Stacks(0)
            End If
            
            .Degree = GetWordL()

        End With

    LinePointer = 1
    
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadPartyLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadMapIconFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 00:20:26
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadMapIconFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        MapIconVersionInform(n) = GetWord()
    Next n
    N_MapIcon = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_MapIconsCount) = N_MapIcon
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_MapIconsCount, CStr(N_MapIcon)
    End If
    
    ReDim MapIcons(N_MapIcon - 1)

    DoEvents
    For n = 0 To N_MapIcon - 1

        With MapIcons(n)
            .ID = n
            
            .strID = GetWord()
            .Flags = Val(GetWord())
            .MeshName = GetWord()

            .mScale = GetWord()
            .Sound = Val(GetWord())
          For m = 0 To 2
            .Offset(m) = GetWord()
          Next m
          
            '触发器数据
       .TriggerCount = Val(GetWord())
       If .TriggerCount > 0 Then
            ReDim Preserve .Triggers(1 To .TriggerCount)
            For i = 1 To .TriggerCount
                  .Triggers(i).tiOn = Val(GetWord())
                  .Triggers(i).ActNum = Val(GetWord())
                  If .Triggers(i).ActNum > 0 Then
                  ReDim Preserve .Triggers(i).tiAct(1 To .Triggers(i).ActNum)
                  For m = 1 To .Triggers(i).ActNum
                       .Triggers(i).tiAct(m).Op = GetWord()
                       .Triggers(i).tiAct(m).ParaNum = Val(GetWord())
                       If .Triggers(i).tiAct(m).ParaNum > 0 Then
                       ReDim Preserve .Triggers(i).tiAct(m).Para(1 To .Triggers(i).tiAct(m).ParaNum)
                       For H = 1 To .Triggers(i).tiAct(m).ParaNum
                            .Triggers(i).tiAct(m).Para(H).Value = GetWord()
                       Next H
                       End If
                  Next m
                  End If
            Next i
       Else
         ReDim .Triggers(0)
       End If
        
            .Edit = CheckEditable(EditInfo_MapIconsCount, n)
            
            AddIndex .ID, .strID
        End With
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadMapIconFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：ReadMapIconLine
'**输    入：Text(String),OutputMapIcon(Type_MapIcon)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 16:59:36
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadMapIconLine(Text As String, OutputMapIcon As Type_MapIcon)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    Dim head As String
    
    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
         With OutputMapIcon
            .strID = GetWordL()
            .Flags = Val(GetWordL())
            .MeshName = GetWordL()

            .mScale = GetWordL()
            
            .Sound = Val(GetWordL())
            .Sound_sndName = Sounds(.Sound).sndName
            
          For m = 0 To 2
            .Offset(m) = GetWordL()
          Next m
          
            '触发器数据
       .TriggerCount = Val(GetWordL())
       If .TriggerCount > 0 Then
            ReDim Preserve .Triggers(1 To .TriggerCount)
            For i = 1 To .TriggerCount
                  .Triggers(i).tiOn = Val(GetWordL())
                  .Triggers(i).ActNum = Val(GetWordL())
                  If .Triggers(i).ActNum > 0 Then
                  ReDim Preserve .Triggers(i).tiAct(1 To .Triggers(i).ActNum)
                  For m = 1 To .Triggers(i).ActNum
                       .Triggers(i).tiAct(m).Op = GetWordL()
                       .Triggers(i).tiAct(m).ParaNum = Val(GetWordL())
                       If .Triggers(i).tiAct(m).ParaNum > 0 Then
                       ReDim Preserve .Triggers(i).tiAct(m).Para(1 To .Triggers(i).tiAct(m).ParaNum)
                       For H = 1 To .Triggers(i).tiAct(m).ParaNum
                            .Triggers(i).tiAct(m).Para(H).Value = GetWordL()
                            BuildQuote_ParamCode .Triggers(i).tiAct(m).Para(H)
                       Next H
                       End If
                  Next m
                  End If
            Next i
       End If
         End With

    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadMapIconLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadSoundFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 09:06:30
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadSoundFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        SoundVersionInform(n) = GetWord()
    Next n
    
    'SoundResources
    N_SoundRes = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_SoundRessCount) = N_SoundRes
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_SoundRessCount, CStr(N_SoundRes)
    End If
    
    ReDim SoundRess(N_SoundRes - 1)

    DoEvents
    For n = 0 To N_SoundRes - 1

        With SoundRess(n)
            .ID = n
            
            .sndName = GetWord()
            .Flags = GetWord()
               
            .Edit = CheckEditable(EditInfo_SoundRessCount, n)
            
            AddIndex .ID, .sndName
        End With
    Next n
    
    
    'Sounds
    N_Sound = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_SoundsCount) = N_Sound
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_SoundsCount, CStr(N_Sound)
    End If
    
    ReDim Sounds(N_Sound - 1)

    DoEvents
    For n = 0 To N_Sound - 1

        With Sounds(n)
            .ID = n
            
            .sndName = GetWord()
            .Flags = GetWord()
            
            .ResourceCount = Val(GetWord())
            
            If .ResourceCount > 0 Then
                ReDim .Resource(1 To .ResourceCount)
                
                For m = 1 To .ResourceCount
                     .Resource(m).ID = Val(GetWord())
                     .Resource(m).Unknown = Val(GetWord())
                Next m
            Else
                ReDim .Resource(0)
            End If
               
            .Edit = CheckEditable(EditInfo_SoundsCount, n)
            
            AddIndex .ID, .sndName
        End With
    Next n
    
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadSoundFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：ReadSoundLine
'**输    入：Text(String),OutputSound(Type_Sound)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 17:07:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadSoundLine(Text As String, OutputSound As Type_Sound)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()

        With OutputSound
        
            .sndName = GetWordL()
            .Flags = GetWordL()
            
            .ResourceCount = Val(GetWordL())
            
            If .ResourceCount > 0 Then
                ReDim .Resource(1 To .ResourceCount)
                
                For m = 1 To .ResourceCount
                     .Resource(m).ID = Val(GetWordL())
                     .Resource(m).strID = SoundRess(.Resource(m).ID).sndName
                     .Resource(m).Unknown = Val(GetWordL())
                Next m
            Else
                ReDim .Resource(0)
            End If
               
        End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadSoundLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：ReadSoundResLine
'**输    入：Text(String),OutputSoundRes(Type_Sound)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 17:11:36
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadSoundResLine(Text As String, OutputSoundRes As Type_SoundResource)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()

        With OutputSoundRes
            .sndName = GetWordL()
            .Flags = GetWordL()
               
        End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadSoundResLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadItemFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:24:01
'**修 改 人：Ser_Charles
'**日    期：2010-12-08 22:42:16
'**版    本：V1.1321
'*************************************************************************

Sub LoadItemFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        ItmVersionInform(n) = GetWord()
    Next n
    N_Item = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_ItemsCount) = N_Item
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_ItemsCount, CStr(N_Item)
    End If
    
    ReDim itm(N_Item - 1)

    DoEvents
    For n = 0 To N_Item - 1

        With itm(n)
            .ID = n
            .dbName = GetWord()
            .disname = GetWord()
            .texname = GetWord()
            .nmdl = Val(GetWord())
            
            .csvName = .disname
            .csvName_pl = .disname
            If .nmdl > 0 Then
               ReDim .mdlname(1 To .nmdl)
               ReDim .mdl_b(1 To .nmdl)
               
               For m = 1 To .nmdl
                .mdlname(m) = GetWord()
                .mdl_b(m) = GetWord()
               Next m
            Else
               ReDim .mdlname(0)
               ReDim .mdl_b(0)
            End If

            .itmType = GetWord()
            .Action = GetWord()
            .price = Val(GetWord())
            '.prefix = Val(GetWord())
            .Prefix = GetWord()
            '.weight = Val(GetWord())   'v951
            .weight = GetWord()    'v952
            .abundance = Val(GetWord())
            .head_armor = Val(GetWord())
            .body_armor = Val(GetWord())
            .leg_armor = Val(GetWord())
            .difficulty = Val(GetWord())
            .hit_points = Val(GetWord())
            .speed_rating = Val(GetWord())
            .missile_speed = Val(GetWord())
            .weapon_length = Val(GetWord())
            .max_ammo = Val(GetWord())
            .thrust_damage = Val(GetWord())
            .swing_damage = Val(GetWord())
            
            .FactionCount = Val(GetWord())
            If .FactionCount > 0 Then
               ReDim .Faction(1 To .FactionCount)
               
               For m = 1 To .FactionCount
                .Faction(m).ID = GetWord()
               Next m
            Else
               ReDim .Faction(0)
            End If
            
            
            '触发器数据
       .TriggerCount = Val(GetWord())
       If .TriggerCount > 0 Then
            ReDim .Trigger(1 To .TriggerCount)
            For i = 1 To .TriggerCount
                  .Trigger(i).tiOn = Val(GetWord())
                  .Trigger(i).ActNum = Val(GetWord())
                  If .Trigger(i).ActNum > 0 Then
                     ReDim .Trigger(i).tiAct(1 To .Trigger(i).ActNum)
                     For m = 1 To .Trigger(i).ActNum
                          .Trigger(i).tiAct(m).Op = GetWord()
                          .Trigger(i).tiAct(m).ParaNum = Val(GetWord())
                          If .Trigger(i).tiAct(m).ParaNum > 0 Then
                             ReDim .Trigger(i).tiAct(m).Para(1 To .Trigger(i).tiAct(m).ParaNum)
                             For H = 1 To .Trigger(i).tiAct(m).ParaNum
                                  .Trigger(i).tiAct(m).Para(H).Value = GetWord()
                             Next H
                          Else
                             ReDim .Trigger(i).tiAct(m).Para(0)
                          End If
                     Next m
                  Else
                     ReDim .Trigger(i).tiAct(0)
                  End If
            Next i
       ElseIf .TriggerCount <= 0 Then
            ReDim .Trigger(0)
       End If
        
            .Edit = CheckEditable(EditInfo_ItemsCount, n)
            
            AddIndex .ID, .dbName
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadItemFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadGlobalVarFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-03-02 20:05:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadGlobalVarFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim i As Integer, s As String
    Dim F As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    F = FreeFile
    Open tmpFileName For Input As #F
       i = 0
       ReDim gVars(0)
       Do While Not EOF(F)
          Line Input #F, s
          s = Replace(s, "$", "")
          If s <> "" Then
             ReDim Preserve gVars(i)
             gVars(i).VarName = s
             gVars(i).ID = i
             i = i + 1
          End If
       Loop
       N_gVar = i
    Close F

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadGlobalVarFile [n=" & CStr(i) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadQuickStringFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-02-03 22:45:14
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadQuickStringFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    N_qStr = Val(GetWord())
    
    'If MnBInfo.FirstTimeEdit Then
    '    MnBInfo.EditInfo(EditInfo_QuickStringsCount) = N_qStr
    '    WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_QuickStringsCount, CStr(N_qStr)
    'End If
    
    ReDim qStrs(N_qStr - 1)

    DoEvents
    For n = 0 To N_qStr - 1

        With qStrs(n)
            .ID = n
            .Name = GetWord()
            .Str = GetWord()
        
            '.Edit = CheckEditable(EditInfo_QuickStringsCount, n)
            
            'AddIndex .ID, .Name
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadQuickStringFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadStringFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-03-02 22:42:05
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadStringFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim strTem As String
    Dim spcPos As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Input Access Read As lngHandle Len = 1

    Pointer = 1
    
        Line Input #lngHandle, strTem
        StringVersionInform(n) = strTem
        
        Line Input #lngHandle, strTem
        N_Str = Val(strTem)
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_StringsCount) = N_Str
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_StringsCount, CStr(N_Str)
    End If
    
    
    ReDim Strs(N_Str - 1)

    DoEvents
    For n = 0 To N_Str - 1

        Line Input #lngHandle, strTem
        spcPos = InStr(1, strTem, " ")
        If Trim(strTem) <> "" Then
          If spcPos > 0 Then
            With Strs(n)
              .ID = n
              .Name = Left(strTem, spcPos - 1)
              .Str = Right(strTem, Len(strTem) - spcPos)
        
              .Edit = CheckEditable(EditInfo_StringsCount, n)
            
              AddIndex .ID, .Name
            End With
          End If
        End If
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadStringFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：ReadItemLine
'**输    入：Text(String),(Type_Item)OutputItem
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 16:12:42
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadItemLine(Text As String, OutputItem As Type_Item, Optional lPointer As Long = 1)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    Dim head As String
    
    LinePointer = lPointer
    
    If LinePointer = 1 Then
       txtLine = "h " & Text
       head = GetWordL()
    End If
    
        With OutputItem
            .dbName = GetWordL()
            .disname = GetWordL()
            .texname = GetWordL()
            .nmdl = Val(GetWordL())

            If .nmdl > 0 Then
               ReDim .mdlname(1 To .nmdl)
               ReDim .mdl_b(1 To .nmdl)
               
               For m = 1 To .nmdl
                .mdlname(m) = GetWordL()
                .mdl_b(m) = GetWordL()
               Next m
            Else
               ReDim .mdlname(0)
               ReDim .mdl_b(0)
            End If

            .itmType = GetWordL()
            .Action = GetWordL()
            .price = Val(GetWordL())
            '.prefix = Val(GetWordL())
            .Prefix = GetWordL()
            '.weight = Val(GetWordL())   'v951
            .weight = GetWordL()    'v952
            .abundance = Val(GetWordL())
            .head_armor = Val(GetWordL())
            .body_armor = Val(GetWordL())
            .leg_armor = Val(GetWordL())
            .difficulty = Val(GetWordL())
            .hit_points = Val(GetWordL())
            .speed_rating = Val(GetWordL())
            .missile_speed = Val(GetWordL())
            .weapon_length = Val(GetWordL())
            .max_ammo = Val(GetWordL())
            .thrust_damage = Val(GetWordL())
            .swing_damage = Val(GetWordL())
            
            .FactionCount = Val(GetWordL())
            If .FactionCount > 0 Then
               ReDim .Faction(1 To .FactionCount)
               
               For m = 1 To .FactionCount
                .Faction(m).ID = GetWordL()
                .Faction(m).strID = Factions(.Faction(m).ID).strID
               Next m
            Else
               ReDim .Faction(0)
            End If
            
            
            '触发器数据
       .TriggerCount = Val(GetWordL())
       If .TriggerCount > 0 Then
            ReDim .Trigger(1 To .TriggerCount)
            For i = 1 To .TriggerCount
                  .Trigger(i).tiOn = Val(GetWordL())
                  .Trigger(i).ActNum = Val(GetWordL())
                  If .Trigger(i).ActNum > 0 Then
                     ReDim .Trigger(i).tiAct(1 To .Trigger(i).ActNum)
                     For m = 1 To .Trigger(i).ActNum
                          .Trigger(i).tiAct(m).Op = GetWordL()
                          .Trigger(i).tiAct(m).ParaNum = Val(GetWordL())
                          If .Trigger(i).tiAct(m).ParaNum > 0 Then
                             ReDim .Trigger(i).tiAct(m).Para(1 To .Trigger(i).tiAct(m).ParaNum)
                             For H = 1 To .Trigger(i).tiAct(m).ParaNum
                                 .Trigger(i).tiAct(m).Para(H).Value = GetWordL()
                                  BuildQuote_ParamCode .Trigger(i).tiAct(m).Para(H)
                             Next H
                          Else
                             ReDim .Trigger(i).tiAct(m).Para(0)
                          End If
                     Next m
                  Else
                     ReDim .Trigger(i).tiAct(0)
                  End If
            Next i
       ElseIf .TriggerCount <= 0 Then
            ReDim .Trigger(0)
       End If
        
        End With

    lPointer = LinePointer

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadItemLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveItemFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:24:04
'**修 改 人：SSgt_Edward
'**日    期：2010-11-30 22:59:07
'**版    本：V0.951.12
'*************************************************************************

Function LoadIModFile(FileName As String) As Boolean
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    LoadIModFile = False
    
    Dim tmpFileName As String
    tmpFileName = FileName
    
    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Function
    End If
    
    Dim TemP As String, arrTmp() As String, n As Long, m As Integer

    n = 0
    
    Dim arrFileBuff() As String
    Dim i As Long
    m = FreeFile
    ReDim arrTmp(0)
    Open tmpFileName For Input As #m
         Do While Not (EOF(m))
            Input #m, arrTmp(n)
                    For i = 1 To Len(arrTmp(n))
                        If Mid(arrTmp(n), i, 1) = " " Then
                            arrTmp(n) = Left(arrTmp(n), i - 1)
                            Exit For
                        End If
                    Next i
                ReDim Preserve arrTmp(UBound(arrTmp) + 1)
                n = n + 1
         Loop
    Close #m
    
    Do While arrTmp(UBound(arrTmp)) = ""
        ReDim Preserve arrTmp(UBound(arrTmp) - 1)
    Loop
    
    N_IMod = UBound(arrTmp) + 1
    
    ReDim IMod(N_IMod - 1)
    
    For i = 0 To N_IMod - 1
        IMod(i).ID = arrTmp(i)
    Next i
    
    LoadIModFile = True

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadIModFile", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：LoadPSysFile
'**输    入：filePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-09 14:32:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadPSysFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        PSysVersionInform(n) = GetWord()
    Next n
    N_PSys = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_PSysCount, CStr(N_PSys)
        MnBInfo.EditInfo(EditInfo_PSysCount) = N_PSys
    End If
    
    ReDim PSys(N_PSys - 1)

    DoEvents
    For n = 0 To N_PSys - 1

        With PSys(n)
            .ID = n
            .strID = GetWord()
            .Flags = GetWord()
            .Mesh_Name = GetWord()
            .Particles_Num = CLng(Val(GetWord()))
            .Life = Val(GetWord())
            .Damping = Val(GetWord())
            .Gravity = Val(GetWord())
            .Turbulance_SZ = Val(GetWord())
            .Turbulance_Str = Val(GetWord())
            
            .Alphak(0).X = Val(GetWord())
            .Alphak(0).Y = Val(GetWord())
            .Alphak(1).X = Val(GetWord())
            .Alphak(1).Y = Val(GetWord())
            .Redk(0).X = Val(GetWord())
            .Redk(0).Y = Val(GetWord())
            .Redk(1).X = Val(GetWord())
            .Redk(1).Y = Val(GetWord())
            .Greenk(0).X = Val(GetWord())
            .Greenk(0).Y = Val(GetWord())
            .Greenk(1).X = Val(GetWord())
            .Greenk(1).Y = Val(GetWord())
            .Bluek(0).X = Val(GetWord())
            .Bluek(0).Y = Val(GetWord())
            .Bluek(1).X = Val(GetWord())
            .Bluek(1).Y = Val(GetWord())
            .Scalek(0).X = Val(GetWord())
            .Scalek(0).Y = Val(GetWord())
            .Scalek(1).X = Val(GetWord())
            .Scalek(1).Y = Val(GetWord())
            
            .EBSZ(0) = Val(GetWord())
            .EBSZ(1) = Val(GetWord())
            .EBSZ(2) = Val(GetWord())
            .EV(0) = Val(GetWord())
            .EV(1) = Val(GetWord())
            .EV(2) = Val(GetWord())
            .EDR = Val(GetWord())
            .PRS = Val(GetWord())
            .PRD = Val(GetWord())
            
            .Edit = CheckEditable(EditInfo_PSysCount, n)
            
            AddIndex .ID, .strID
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadPSysFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadTabMatFile
'**输    入：filePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-09 14:32:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadTabMatFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim i As Integer
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    'For n = 0 To 2
    '    TabMatVersionInform(n) = GetWord()
    'Next n
    N_TabMat = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_TabMatCount, CStr(N_TabMat)
        MnBInfo.EditInfo(EditInfo_TabMatCount) = N_TabMat
    End If
    
    ReDim TabMat(N_TabMat - 1)

    DoEvents
    For n = 0 To N_TabMat - 1

        With TabMat(n)
            .ID = n
            .strID = GetWord()
            .Flags = GetWord()
            .Sample = GetWord()
            .Width = CLng(Val(GetWord()))
            .Height = CLng(Val(GetWord()))
            .Min.X = CLng(Val(GetWord()))
            .Min.Y = CLng(Val(GetWord()))
            .Max.X = CLng(Val(GetWord()))
            .Max.Y = CLng(Val(GetWord()))
            .OpCount = CLng(Val(GetWord()))
            
            If .OpCount = 0 Then
                 ReDim .OpBlock(0)
            Else
                 ReDim .OpBlock(1 To .OpCount)
                 For i = 1 To .OpCount
                       .OpBlock(i).Op = GetWord()
                       .OpBlock(i).ParaNum = CLng(Val(GetWord))
                       If .OpBlock(i).ParaNum = 0 Then
                            ReDim .OpBlock(i).Para(0)
                       Else
                            ReDim .OpBlock(i).Para(1 To .OpBlock(i).ParaNum)
                            For m = 1 To .OpBlock(i).ParaNum
                                 .OpBlock(i).Para(m).Value = GetWord()
                            Next m
                       End If
                 Next i
            End If
            
            .Edit = CheckEditable(EditInfo_TabMatCount, n)
            
            AddIndex .ID, .strID
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadTabMatFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadMeshFile
'**输    入：filePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-28 19:10:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadMeshFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    'For n = 0 To 2
    '    MeshVersionInform(n) = GetWord()
    'Next n
    N_Mesh = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_MeshCount, CStr(N_Mesh)
        MnBInfo.EditInfo(EditInfo_MeshCount) = N_Mesh
    End If
    
    ReDim Mesh(N_Mesh - 1)

    DoEvents
    For n = 0 To N_Mesh - 1

        With Mesh(n)
            .ID = n
            .strID = GetWord()
            .Flags = GetWord()
            .Resource_Name = GetWord()
            
            .Translation.X = Val(GetWord())
            .Translation.Y = Val(GetWord())
            .Translation.Z = Val(GetWord())

            .Rotation_Angle.X = Val(GetWord())
            .Rotation_Angle.Y = Val(GetWord())
            .Rotation_Angle.Z = Val(GetWord())
            
            .Scale.X = Val(GetWord())
            .Scale.Y = Val(GetWord())
            .Scale.Z = Val(GetWord())
            
            .Edit = CheckEditable(EditInfo_MeshCount, n)
            
            AddIndex .ID, .strID
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadMeshFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：LoadTriggerFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-22 12:59:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub LoadTriggerFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer, i As Integer, H As Integer
    
    Dim tmpFileName As String
    tmpFileName = FilePath

    If FileLen(tmpFileName) = 0 Then
        MsgBox "缺少文件: ( " & tmpFileName & " )"
        Exit Sub
    End If
    
    MaxPointer = FileLen(tmpFileName)
    lngHandle = FreeFile()
    Open tmpFileName For Random Access Read As lngHandle Len = 1

    Pointer = 1
    For n = 0 To 2
        TimeTrgVersionInform(n) = GetWord()
    Next n
    N_TimeTrg = Val(GetWord())
    
    If MnBInfo.FirstTimeEdit Then
        MnBInfo.EditInfo(EditInfo_TimeTrgCount) = N_TimeTrg
        WriteString MnBInfo.iniFileName, "EDITINFO", "Count" & EditInfo_TimeTrgCount, CStr(N_TimeTrg)
    End If
    
    ReDim TimeTrg(N_TimeTrg - 1)

    DoEvents
    For n = 0 To N_TimeTrg - 1

        With TimeTrg(n)
            .ID = n
            .Check_Interval = Val(GetWord())
            .Delay_Interval = Val(GetWord())
            .Rearm_Interval = Val(GetWord())
    
    '-----------------------------条件块------------------------------------------
       .ConditionsCount = CLng(Val(GetWord()))
       If .ConditionsCount > 0 Then
            ReDim .Condition(1 To .ConditionsCount)
            For i = 1 To .ConditionsCount
                       .Condition(i).Op = GetWord()
                       .Condition(i).ParaNum = Val(GetWord())
                       If .Condition(i).ParaNum > 0 Then
                           ReDim .Condition(i).Para(1 To .Condition(i).ParaNum)
                           For H = 1 To .Condition(i).ParaNum
                                 .Condition(i).Para(H).Value = GetWord()
                           Next H
                       Else
                           ReDim .Condition(i).Para(0)
                       End If
            Next i
       ElseIf .ConditionsCount <= 0 Then
            ReDim .Condition(0)
       End If
     '-----------------------------------------------------------------------------
     
    '-----------------------------结果块------------------------------------------
       .ConsequencesCount = CLng(Val(GetWord()))
       If .ConsequencesCount > 0 Then
            ReDim .Consequence(1 To .ConsequencesCount)
            For i = 1 To .ConsequencesCount
                       .Consequence(i).Op = GetWord()
                       .Consequence(i).ParaNum = Val(GetWord())
                       If .Consequence(i).ParaNum > 0 Then
                           ReDim .Consequence(i).Para(1 To .Consequence(i).ParaNum)
                           For H = 1 To .Consequence(i).ParaNum
                                 .Consequence(i).Para(H).Value = GetWord()
                           Next H
                       Else
                           ReDim .Consequence(i).Para(0)
                       End If
            Next i
       ElseIf .ConsequencesCount <= 0 Then
            ReDim .Consequence(0)
       End If
     '-----------------------------------------------------------------------------
     
            .Edit = CheckEditable(EditInfo_TimeTrgCount, n)
            
        End With
    Next n
    Close lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "LoadTriggerFile [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：ReadPSysLine
'**输    入：Text(String),OutputPSys(Type_Particle_System)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-09 17:07:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadPSysLine(Text As String, OutputPSys As Type_Particle_System)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()

        With OutputPSys
        
            .strID = GetWordL()
            .Flags = GetWordL()
            .Mesh_Name = GetWordL()
            .Particles_Num = CLng(Val(GetWordL()))
            .Life = Val(GetWordL())
            .Damping = Val(GetWordL())
            .Gravity = Val(GetWordL())
            .Turbulance_SZ = Val(GetWordL())
            .Turbulance_Str = Val(GetWordL())
            
            .Alphak(0).X = Val(GetWordL())
            .Alphak(0).Y = Val(GetWordL())
            .Alphak(1).X = Val(GetWordL())
            .Alphak(1).Y = Val(GetWordL())
            .Redk(0).X = Val(GetWordL())
            .Redk(0).Y = Val(GetWordL())
            .Redk(1).X = Val(GetWordL())
            .Redk(1).Y = Val(GetWordL())
            .Greenk(0).X = Val(GetWordL())
            .Greenk(0).Y = Val(GetWordL())
            .Greenk(1).X = Val(GetWordL())
            .Greenk(1).Y = Val(GetWordL())
            .Bluek(0).X = Val(GetWordL())
            .Bluek(0).Y = Val(GetWordL())
            .Bluek(1).X = Val(GetWordL())
            .Bluek(1).Y = Val(GetWordL())
            .Scalek(0).X = Val(GetWordL())
            .Scalek(0).Y = Val(GetWordL())
            .Scalek(1).X = Val(GetWordL())
            .Scalek(1).Y = Val(GetWordL())
            
            .EBSZ(0) = Val(GetWordL())
            .EBSZ(1) = Val(GetWordL())
            .EBSZ(2) = Val(GetWordL())
            .EV(0) = Val(GetWordL())
            .EV(1) = Val(GetWordL())
            .EV(2) = Val(GetWordL())
            .EDR = Val(GetWordL())
            .PRS = Val(GetWordL())
            .PRD = Val(GetWordL())
               
        End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadPSysLine", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：readTriggerLine
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-03-13 11:35:36
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub readTriggerLine(Text As String, OutputTrigger As Type_Time_Trigger)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim n As Long
    Dim i As Integer
    Dim m As Integer
    Dim H As Integer
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
    
        With OutputTrigger
            .Check_Interval = Val(GetWordL())
            .Delay_Interval = Val(GetWordL())
            .Rearm_Interval = Val(GetWordL())
    
    '-----------------------------条件块------------------------------------------
       .ConditionsCount = CLng(Val(GetWordL()))
       If .ConditionsCount > 0 Then
            ReDim .Condition(1 To .ConditionsCount)
            For i = 1 To .ConditionsCount
                       .Condition(i).Op = GetWordL()
                       .Condition(i).ParaNum = Val(GetWordL())
                       If .Condition(i).ParaNum > 0 Then
                           ReDim .Condition(i).Para(1 To .Condition(i).ParaNum)
                           For H = 1 To .Condition(i).ParaNum
                                 .Condition(i).Para(H).Value = GetWordL()
                           Next H
                       Else
                           ReDim .Condition(i).Para(0)
                       End If
            Next i
       ElseIf .ConditionsCount <= 0 Then
            ReDim .Condition(0)
       End If
     '-----------------------------------------------------------------------------
     
    '-----------------------------结果块------------------------------------------
       .ConsequencesCount = CLng(Val(GetWordL()))
       If .ConsequencesCount > 0 Then
            ReDim .Consequence(1 To .ConsequencesCount)
            For i = 1 To .ConsequencesCount
                       .Consequence(i).Op = GetWordL()
                       .Consequence(i).ParaNum = Val(GetWordL())
                       If .Consequence(i).ParaNum > 0 Then
                           ReDim .Consequence(i).Para(1 To .Consequence(i).ParaNum)
                           For H = 1 To .Consequence(i).ParaNum
                                 .Consequence(i).Para(H).Value = GetWordL()
                           Next H
                       Else
                           ReDim .Consequence(i).Para(0)
                       End If
            Next i
       ElseIf .ConsequencesCount <= 0 Then
            ReDim .Consequence(0)
       End If
     '-----------------------------------------------------------------------------
            
        End With
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "readTriggerLine [n=" & CStr(n) & "]", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：ReadTabMatLine
'**输    入：Text(String),OutputTabMat(Type_Tableau_Material) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-09 14:32:22
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadTabMatLine(Text As String, OutputTabMat As Type_Tableau_Material)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim i As Integer
    Dim m As Integer
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()
    
        With OutputTabMat
            .strID = GetWordL()
            .Flags = GetWordL()
            .Sample = GetWordL()
            .Width = CLng(Val(GetWordL()))
            .Height = CLng(Val(GetWordL()))
            .Min.X = CLng(Val(GetWordL()))
            .Min.Y = CLng(Val(GetWordL()))
            .Max.X = CLng(Val(GetWordL()))
            .Max.Y = CLng(Val(GetWordL()))
            .OpCount = CLng(Val(GetWordL()))
            
            If .OpCount = 0 Then
                 ReDim .OpBlock(0)
            Else
                 ReDim .OpBlock(1 To .OpCount)
                 For i = 1 To .OpCount
                       .OpBlock(i).Op = GetWordL()
                       .OpBlock(i).ParaNum = CLng(Val(GetWordL))
                       If .OpBlock(i).ParaNum = 0 Then
                            ReDim .OpBlock(i).Para(0)
                       Else
                            ReDim .OpBlock(i).Para(1 To .OpBlock(i).ParaNum)
                            For m = 1 To .OpBlock(i).ParaNum
                                 .OpBlock(i).Para(m).Value = GetWordL()
                                 BuildQuote_ParamCode .OpBlock(i).Para(m)
                            Next m
                       End If
                 Next i
            End If

        End With

    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadTabMatLine", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：ReadMeshLine
'**输    入：Text(String),OutputMesh(Type_Mesh)
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-01-28 21:24:54
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub ReadMeshLine(Text As String, OutputMesh As Type_Mesh)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    
    Dim n As Long
    Dim m As Integer
    
    Dim head As String

    txtLine = "h " & Text

    LinePointer = 1
    
    head = GetWordL()

        With OutputMesh
        
            .strID = GetWordL()
            .Flags = GetWordL()
            .Resource_Name = GetWordL()
            
            .Translation.X = Val(GetWordL())
            .Translation.Y = Val(GetWordL())
            .Translation.Z = Val(GetWordL())
              
            .Rotation_Angle.X = Val(GetWordL())
            .Rotation_Angle.Y = Val(GetWordL())
            .Rotation_Angle.Z = Val(GetWordL())
            
            .Scale.X = Val(GetWordL())
            .Scale.Y = Val(GetWordL())
            .Scale.Z = Val(GetWordL())
        End With
    
    LinePointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileLoader", "ReadMeshLine", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：GetYesNoStr
'**输    入：lVal(Long) -
'**输    出：(String) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-06 16:38:37
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub ReadAll()   'A_E

frmMain.Caption = Replace(frmMain.Caption, ":[ModName]", ":" & MnBInfo.ModName, , , vbTextCompare)

DoEvents

'ModPic.Picture = LoadPicture(MnBInfo.ModPath & "\Main.bmp")

'Load Sounds
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(9))
LoadSoundFile MnBInfo.ModPath & "\sounds.txt"

'Load Items
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(2))
LoadItemFile MnBInfo.ModPath & "\item_kinds1.txt"
LoadItemCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\item_kinds.csv"

'Load Troops
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(1))
LoadTroopFile MnBInfo.ModPath & "\troops.txt"
LoadTroopCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\troops.csv"

'Load Party_Templates
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(4))
LoadPTFile MnBInfo.ModPath & "\party_templates.txt"
LoadPartyTemplateCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\party_templates.csv"

'Load Parties
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(3))
LoadPartyFile MnBInfo.ModPath & "\parties.txt"
LoadPartyCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\parties.csv"

'Load Factions
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(5))
LoadFactionFile MnBInfo.ModPath & "\factions.txt"
LoadFactionCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\factions.csv"

'Load IModifiers
frmTip.ShowTip ActiveString(PublicTips(1), PublicMsgs(147))
LoadIModFile MnBInfo.MBHome & "\Data\item_modifiers.txt"
LoadIModCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\item_modifiers.csv"

'Load Scenes
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(6))
LoadSceneFile MnBInfo.ModPath & "\scenes.txt"

'Load Particle System
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(8))
LoadPSysFile MnBInfo.ModPath & "\particle_systems.txt"

'Load Map Icons
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(7))
LoadMapIconFile MnBInfo.ModPath & "\map_icons.txt"

'Load Tableau Materials
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(11))
LoadTabMatFile MnBInfo.ModPath & "\tableau_materials.txt"

'Load Meshes
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(12))
LoadMeshFile MnBInfo.ModPath & "\meshes.txt"

'Load Triggers
frmTip.ShowTip ActiveString(PublicTips(1), PublicEditors(13))
LoadTriggerFile MnBInfo.ModPath & "\triggers.txt"

'Load Global Variables
'frmTip.ShowTip ActiveString(PublicTips(1), PublicTags(Tag_Variable))
LoadGlobalVarFile MnBInfo.ModPath & "\variables.txt"

'Load Quick Strings
frmTip.ShowTip ActiveString(PublicTips(1), PublicTags(Tag_Quick_String))
LoadQuickStringFile MnBInfo.ModPath & "\quick_strings.txt"
LoadQuickStringCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\quick_strings.csv"

'Load Strings
If IsLoadString Then
   frmTip.ShowTip ActiveString(PublicTips(1), PublicTags(Tag_String))
   LoadStringFile MnBInfo.ModPath & "\strings.txt"
   LoadStringCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\game_strings.csv"
End If

'Load Editors
'InitEditorsListView

'Build Quote
frmTip.ShowTip PublicMsgs(148)
BuildQuote

'Read Variable Name Check List
ReadVarNameCheckLists

frmTip.HideTip

frmMain.mSave.Enabled = True
frmMain.mSaveAs.Enabled = True
frmMain.mBackUp.Enabled = True
frmMain.mEditor.Enabled = True
frmMain.mTool.Enabled = True

End Sub

