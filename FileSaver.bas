Attribute VB_Name = "FileSaver"

Option Explicit

Sub SaveTroopFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim tmpStr As String
    Dim n As Long, i As Long
    Dim m As Integer

    lngHandle = FreeFile()

    Open FilePath For Output As #lngHandle

    tmpStr = ""
    For n = 0 To 1
        tmpStr = tmpStr & TrpsVersionInform(n) & " "
    Next n
    tmpStr = tmpStr & TrpsVersionInform(2)
    Print #lngHandle, tmpStr
    Print #lngHandle, Trim$(Str$(N_Troop)) & " "
    For n = 0 To N_Troop - 1

        With Trps(n)

            'Check Value
            If Left$(Trps(n).strID, 4) <> "trp_" Then
                Trps(n).strID = "trp_" & Trps(n).strID
            End If

            Print #lngHandle, .strID & " " & .strName & " " & .strPtName & " " & .unknown_warband(1) & " " & .Flags & _
                  " " & CStr(.Scene) & " " & CStr(.reserved) & " " & CStr(.Faction) & " " & .Upgrade1 & " " & .Upgrade2
            tmpStr = "  "
            For m = 1 To 64
                tmpStr = tmpStr & CStr(.lstInventory(m).X) & " " & CStr(.lstInventory(m).Y) & " "
            Next m
            Print #lngHandle, tmpStr

            Print #lngHandle, "  " & CStr(.tAttrib.strPoint) & " " & CStr(.tAttrib.agiPoint) & _
                  " " & CStr(.tAttrib.intPoint) & " " & CStr(.tAttrib.chaPoint) & " " & CStr(.tAttrib.level)

            Print #lngHandle, " " & .WP.one_handed & " " & .WP.two_handed & " " & .WP.polearm & _
                  " " & .WP.archery & " " & .WP.crossbow & " " & .WP.throwing & " " & .WP.firearm

            Print #lngHandle, "" & CStr(.Skills(1)) & " " & CStr(.Skills(2)) & " " & CStr(.Skills(3)) & _
                  " " & CStr(.Skills(4)) & " " & CStr(.Skills(5)) & " " & CStr(.Skills(6)) & " "

            Print #lngHandle, "  " & .Face(1) & " " & .Face(2) & " " & .Face(3) & " " & .Face(4) & " " & _
                  .Face(5) & " " & .Face(6) & " " & .Face(7) & " " & .Face(8) & " "

            Print #lngHandle, ""
            
        End With
    Next n
    Close lngHandle

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveTroopFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveTroopCSVFile
'**输    入：filePath(String) -
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

Public Sub SaveTroopCSVFile(FilePath As String)
     On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SaveTroopCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Troop - 1
       If Trps(i).csvName <> Trps(i).strName Then
          strTem = strTem & Trps(i).strID & "|" & Trps(i).csvName & vbCrLf
       End If
       
       If Trps(i).csvName_pl <> Trps(i).strPtName Then
          strTem = strTem & Trps(i).strID & "_pl" & "|" & Trps(i).csvName_pl & vbCrLf
       End If
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("FileSaver", "SaveTroopCSVFile", "INVALID_HANDLE_VALUE", "无法打开文件:[" & FilePath & "]")
    End If
     Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveTroopCSVFile", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**函 数 名：LoadPTFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-05-18 08:23:48
'**修 改 人：SSgt_Edward
'**日    期：2010-11-30 14:52:37
'**版    本：V1.1321
'*************************************************************************

Public Sub SavePartyTemplateCSVFile(FilePath As String)
     On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SavePartyTemplateCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_PT - 1
       If PTs(i).csvName <> PTs(i).ptName Then
          strTem = strTem & PTs(i).ptID & "|" & PTs(i).csvName & vbCrLf
       End If
       
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("FileSaver", "SavePartyTemplateCSVFile", "INVALID_HANDLE_VALUE", "无法打开文件:[" & FilePath & "]")
    End If
     Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SavePartyTemplateCSVFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SavePartyTemplateFile
'**输    入：FilePath(String) -PT文件的路径
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：kevin
'**日    期：2008-06-07 16:49:36
'**修 改 人：SSgt_Edward
'**日    期：2010-12-07 15:22:37
'**版    本：V1.1321
'*************************************************************************

Sub SavePartyTemplateFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim i As Double
    Dim n As Long
    Dim m As Integer

    Dim tmptxtLine As String

    lngHandle = FreeFile()
    'fix me
    Open FilePath For Output As #lngHandle
    tmptxtLine = ""
    For n = 0 To 1
        tmptxtLine = tmptxtLine & (PTVersionInform(n)) & " "
    Next n
    tmptxtLine = tmptxtLine & (PTVersionInform(2))
    Print #lngHandle, tmptxtLine

    Print #lngHandle, (CStr(N_PT))

    For n = 0 To N_PT - 1
           tmptxtLine = OutputPTLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SavePartyTemplateFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SavePartyCSVFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-08 16:37:32
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SavePartyCSVFile(FilePath As String)
     On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SavePartyTemplateCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Party - 1
       If Parties(i).csvName <> Parties(i).strName Then
          strTem = strTem & Parties(i).strID & "|" & Parties(i).csvName & vbCrLf
       End If
       
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("FileSaver", "SavePartyCSVFile", "INVALID_HANDLE_VALUE", "无法打开文件:[" & FilePath & "]")
    End If
     Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SavePartyCSVFile", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**函 数 名：SaveSoundFile
'**输    入：FilePath(String) -Sound文件的路径
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2011-01-03 15:40:38
'**版    本：V1.1321
'*************************************************************************

Sub SaveSoundFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim i As Double
    Dim n As Long
    Dim m As Integer

    Dim tmptxtLine As String

    lngHandle = FreeFile()
    'fix me
    Open FilePath For Output As #lngHandle
    tmptxtLine = ""
    For n = 0 To 1
        tmptxtLine = tmptxtLine & (SoundVersionInform(n)) & " "
    Next n
    tmptxtLine = tmptxtLine & (SoundVersionInform(2))
    Print #lngHandle, tmptxtLine
    
    Print #lngHandle, (CStr(N_SoundRes))

    For n = 0 To N_SoundRes - 1
           tmptxtLine = OutputSoundResLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    
    Print #lngHandle, (CStr(N_Sound))

    For n = 0 To N_Sound - 1
           tmptxtLine = OutputSoundLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveSoundFile", Err.Number, Err.DescriSoundion)
End Sub

'*************************************************************************
'**函 数 名：SavePartyFile
'**输    入：FilePath(String) -Party文件的路径
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-09 13:22:42
'**版    本：V1.1321
'*************************************************************************

Sub SavePartyFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim i As Double
    Dim n As Long
    Dim m As Integer

    Dim tmptxtLine As String

    lngHandle = FreeFile()
    'fix me
    Open FilePath For Output As #lngHandle
    tmptxtLine = ""
    For n = 0 To 1
        tmptxtLine = tmptxtLine & (PartyVersionInform(n)) & " "
    Next n
    tmptxtLine = tmptxtLine & (PartyVersionInform(2))
    Print #lngHandle, tmptxtLine

    Print #lngHandle, (CStr(N_Party)) & " " & (CStr(N_Party2))

    For n = 0 To N_Party - 1
           tmptxtLine = OutputPartyLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SavePartyFile", Err.Number, Err.DescriPartyion)
End Sub

'*************************************************************************
'**函 数 名：SaveSceneFile
'**输    入：FilePath(String) -Scene文件的路径
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-09 13:22:42
'**版    本：V1.1321
'*************************************************************************

Sub SaveSceneFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim i As Double
    Dim n As Long
    Dim m As Integer

    Dim tmptxtLine As String

    lngHandle = FreeFile()
    'fix me
    Open FilePath For Output As #lngHandle
    tmptxtLine = ""
    For n = 0 To 1
        tmptxtLine = tmptxtLine & (SceneVersionInform(n)) & " "
    Next n
    tmptxtLine = tmptxtLine & (SceneVersionInform(2))
    Print #lngHandle, tmptxtLine

    Print #lngHandle, " " & (CStr(N_Scene))

    For n = 0 To N_Scene - 1
           tmptxtLine = OutputSceneLine(n)
           Print #lngHandle, tmptxtLine
    Next n
    Close #lngHandle
    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveSceneFile", Err.Number, Err.DescriSceneion)
End Sub


'*************************************************************************
'**函 数 名：LoadPartyFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-07 23:35:16
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Sub SaveItemFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    For n = 0 To 1
        txtLine = txtLine & ItmVersionInform(n) & " "
    Next n

    txtLine = txtLine & ItmVersionInform(2) & vbCrLf

    txtLine = txtLine & CStr(N_Item)
    Print #lngFileHandle, txtLine

    For n = 0 To N_Item - 1
            txtLine = OutputItemLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveItemFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveTriggerFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-23 14:11:36
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************

Sub SaveTriggerFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    For n = 0 To 1
        txtLine = txtLine & TimeTrgVersionInform(n) & " "
    Next n

    txtLine = txtLine & TimeTrgVersionInform(2) & vbCrLf

    txtLine = txtLine & CStr(N_TimeTrg)
    Print #lngFileHandle, txtLine

    For n = 0 To N_TimeTrg - 1
            txtLine = OutputTimeTriggerLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveTriggerFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveStringFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-03-03 23:39:36
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************

Sub SaveStringFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""
    
    txtLine = txtLine & StringVersionInform(0) & vbCrLf

    txtLine = txtLine & CStr(N_Str)
    Print #lngFileHandle, txtLine

    For n = 0 To N_Str - 1
            'Strs(n).Name = Replace(Strs(n).Name, " ", "_")
            'Strs(n).Str = Replace(Strs(n).Str, " ", "_")
            txtLine = Strs(n).Name & " " & Strs(n).Str
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveStringFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SavePSysFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-01-07 16:26:23
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************

Sub SavePSysFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    For n = 0 To 1
        txtLine = txtLine & PSysVersionInform(n) & " "
    Next n

    txtLine = txtLine & PSysVersionInform(2) & vbCrLf

    txtLine = txtLine & CStr(N_PSys)
    Print #lngFileHandle, txtLine

    For n = 0 To N_PSys - 1
            txtLine = OutputPSysLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SavepsysFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveTabMatFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-01-22 19:46:54
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************

Sub SaveTabMatFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    txtLine = txtLine & CStr(N_TabMat)
    Print #lngFileHandle, txtLine

    For n = 0 To N_TabMat - 1
            txtLine = OutputTabMatLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveTabMatFile", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：SaveMeshFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-01-07 16:26:23
'**修 改 人：
'**日    期：
'**版    本：V0.951.12
'*************************************************************************

Sub SaveMeshFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim n As Long
    Dim m As Integer

    Dim txtLine As String

    Dim lngFileHandle As Integer
   
    lngFileHandle = FreeFile()
    Open FilePath For Output As #lngFileHandle

    txtLine = ""

    'For n = 0 To 1
    '    txtLine = txtLine & MeshVersionInform(n) & " "
    'Next n

    'txtLine = txtLine & MeshVersionInform(2) & vbCrLf

    txtLine = txtLine & CStr(N_Mesh)
    Print #lngFileHandle, txtLine

    For n = 0 To N_Mesh - 1
            txtLine = OutputMeshLine(n)
            Print #lngFileHandle, txtLine
    Next n

    Close #lngFileHandle

    Pointer = 1

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveMeshFile", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：InitTags
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-07 11:05:54
'**修 改 人：
'**日    期：
'**版    本：V1.132.1
'*************************************************************************

Public Function OutputItemLine(ByVal Itm_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim ItemVes As Type_Item, temLine As String, tmpStr As String, m As Integer

If Itm_Idx >= 0 Then
    ItemVes = itm(Itm_Idx)
Else
    ItemVes = CurrentItm
End If

        With ItemVes
        
            'Check Value
            If Left$(.dbName, 4) <> "itm_" Then
                .dbName = "itm_" & .dbName
            End If
            
            .texname = .disname
            temLine = " " & .dbName & " " & .disname & " " & .texname & " " & CStr(.nmdl) & " "

            For m = 1 To .nmdl
                temLine = temLine & " " & .mdlname(m) & " " & .mdl_b(m) & " "
            Next m
            temLine = temLine & " " & .itmType _
                      & " " & .Action _
                      & " " & CStr(.price) _
                      & " " & CStr(.Prefix) _
                      & " " & Format$(.weight, "0.000000") _
                      & " " & CStr(.abundance) _
                      & " " & CStr(.head_armor) _
                      & " " & CStr(.body_armor) _
                      & " " & CStr(.leg_armor) _
                      & " " & CStr(.difficulty) _
                      & " " & CStr(.hit_points) _
                      & " " & CStr(.speed_rating) _
                      & " " & CStr(.missile_speed) _
                      & " " & CStr(.weapon_length) _
                      & " " & CStr(.max_ammo) _
                      & " " & CStr(.thrust_damage) _
                      & " " & CStr(.swing_damage) & vbCrLf

            temLine = temLine & " " & CStr(.FactionCount) & vbCrLf

            For m = 1 To .FactionCount
                temLine = temLine & " " & .Faction(m).ID
            Next m
            
            If .FactionCount > 0 Then
                temLine = temLine & vbCrLf
            End If
            
            temLine = temLine & .TriggerCount
            
            If .TriggerCount > 0 Then
            temLine = temLine & vbCrLf
               For m = 1 To .TriggerCount
                 tmpStr = OutputTriggerLine(.Trigger(m))
                 temLine = temLine & tmpStr
               Next m
            End If

        End With

     OutputItemLine = temLine & vbCrLf

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputItemLine", Err.Number, Err.Description)
End Function
'*************************************************************************
'**函 数 名：OutputTimeTriggerLine
'**输    入：(Long)TimeTrg_Idx[如为-1则指当前TimeTrg]
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-02-23 14:09:19
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputTimeTriggerLine(ByVal TimeTrg_idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim TimeTrgVes As Type_Time_Trigger, temLine As String, tmpStr As String, m As Integer

If TimeTrg_idx >= 0 Then
    TimeTrgVes = TimeTrg(TimeTrg_idx)
Else
    TimeTrgVes = CurrentTimeTrg
End If

        With TimeTrgVes

            temLine = CStr(Format(.Check_Interval, "0.000000")) & " " & CStr(Format(.Delay_Interval, "0.000000")) & " " & CStr(Format(.Rearm_Interval, "0.000000")) & "  "
            
           '----------------------条件块--------------------------
            temLine = temLine & .ConditionsCount & " "
            If .ConditionsCount > 0 Then
               For m = 1 To .ConditionsCount
                 tmpStr = OutputOperationLine(.Condition(m))
                 temLine = temLine & tmpStr
               Next m
            End If
            temLine = temLine & " "
           '------------------------------------------------------
            
           '----------------------结果块--------------------------
            temLine = temLine & .ConsequencesCount & " "
            If .ConsequencesCount > 0 Then
               For m = 1 To .ConsequencesCount
                 tmpStr = OutputOperationLine(.Consequence(m))
                 temLine = temLine & tmpStr
               Next m
            End If
           '------------------------------------------------------
           
        End With

     OutputTimeTriggerLine = temLine & vbCrLf

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputTimeTriggerLine", Err.Number, Err.Description)
End Function
'*************************************************************************
'**函 数 名：OutputPSysLine
'**输    入：(Long)PSys_Idx[如为-1则指当前PSys]
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：Ser_Charles
'**日    期：2011-1-7 15:53:43
'**版    本：V1.1321
'*************************************************************************

Public Function OutputPSysLine(ByVal PSys_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim PSysVes As Type_Particle_System, temLine As String, tmpStr As String, m As Integer

If PSys_Idx >= 0 Then
    PSysVes = PSys(PSys_Idx)
Else
    PSysVes = CurrentPSys
End If

        With PSysVes
        
            'Check Value
            If Left$(.strID, 5) <> "psys_" Then
                .strID = "psys_" & .strID
            End If
            
            temLine = .strID & " " & .Flags & " " & .Mesh_Name & "  " & .Particles_Num & " " & Format(.Life, "0.000000") & " " & Format(.Damping, "0.000000") & " " & Format(.Gravity, "0.000000") & " " & Format(.Turbulance_SZ, "0.000000") & " " & Format(.Turbulance_Str, "0.000000") & " " & vbCrLf

            temLine = temLine & Format(.Alphak(0).X, "0.000000") & " " & Format(.Alphak(0).Y, "0.000000") & "   " & Format(.Alphak(1).X, "0.000000") & " " & Format(.Alphak(1).Y, "0.000000") & vbCrLf _
                      & Format(.Redk(0).X, "0.000000") & " " & Format(.Redk(0).Y, "0.000000") & "   " & Format(.Redk(1).X, "0.000000") & " " & Format(.Redk(1).Y, "0.000000") & vbCrLf _
                      & Format(.Greenk(0).X, "0.000000") & " " & Format(.Greenk(0).Y, "0.000000") & "   " & Format(.Greenk(1).X, "0.000000") & " " & Format(.Greenk(1).Y, "0.000000") & vbCrLf _
                      & Format(.Bluek(0).X, "0.000000") & " " & Format(.Bluek(0).Y, "0.000000") & "   " & Format(.Bluek(1).X, "0.000000") & " " & Format(.Bluek(1).Y, "0.000000") & vbCrLf _
                      & Format(.Scalek(0).X, "0.000000") & " " & Format(.Scalek(0).Y, "0.000000") & "   " & Format(.Scalek(1).X, "0.000000") & " " & Format(.Scalek(1).Y, "0.000000") & vbCrLf _
                      & Format(.EBSZ(0), "0.000000") & " " & Format(.EBSZ(1), "0.000000") & " " & Format(.EBSZ(2), "0.000000") & "   " _
                      & Format(.EV(0), "0.000000") & " " & Format(.EV(1), "0.000000") & " " & Format(.EV(2), "0.000000") & "   " _
                      & Format(.EDR, "0.000000") & " " & vbCrLf _
                      & Format(.PRS, "0.000000") & " " & Format(.PRD, "0.000000") & " "
         End With
         
     OutputPSysLine = temLine

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputPSysLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputTabMatLine
'**输    入：(Long)TabMat_Idx[如为-1则指当前可变素材]
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：Ser_Charles
'**日    期：2011-1-22 19:18:22
'**版    本：V1.1321
'*************************************************************************

Public Function OutputTabMatLine(ByVal TabMat_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim TabMatVes As Type_Tableau_Material, temLine As String, tmpStr As String, m As Integer, i As Integer

If TabMat_Idx >= 0 Then
    TabMatVes = TabMat(TabMat_Idx)
Else
    TabMatVes = CurrentTabMat
End If

        With TabMatVes
        
            'Check Value
            If Left$(.strID, 4) <> "tab_" Then
                .strID = "tab_" & .strID
            End If
            
            temLine = .strID & " " & .Flags & " " & .Sample & " " & .Width & " " & .Height & " " & .Min.X & " " & .Min.Y & " " & .Max.X & " " & .Max.Y & " "

            temLine = temLine & .OpCount & " "
            
            If .OpCount > 0 Then
               For i = 1 To .OpCount
                 'tmpStr = OutputTriggerLine(.trigger(m))
                 temLine = temLine & .OpBlock(i).Op & " " & .OpBlock(i).ParaNum & " "
                 If .OpBlock(i).ParaNum > 0 Then
                       For m = 1 To .OpBlock(i).ParaNum
                            With .OpBlock(i)
                                temLine = temLine & .Para(m).Value & " "
                            End With
                       Next m
                 End If
               Next i
            End If
            
         End With
         
     OutputTabMatLine = temLine

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputTabMatLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputMeshLine
'**输    入：(Long)Mesh_Idx[如为-1则指当前Mesh]
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：Ser_Charles
'**日    期：2011-1-28 19:19:55
'**版    本：V1.1321
'*************************************************************************

Public Function OutputMeshLine(ByVal Mesh_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim MeshVes As Type_Mesh, temLine As String, tmpStr As String, m As Integer

If Mesh_Idx >= 0 Then
    MeshVes = Mesh(Mesh_Idx)
Else
    MeshVes = CurrentMesh
End If

        With MeshVes
        
            'Check Value
            If Left$(.strID, 5) <> "mesh_" Then
                .strID = "mesh_" & .strID
            End If
            
            temLine = .strID & " " & .Flags & " " & .Resource_Name & " "

            temLine = temLine & Format(.Translation.X, "0.000000") & " " & Format(.Translation.Y, "0.000000") & " " & Format(.Translation.Z, "0.000000") & " " _
                              & Format(.Rotation_Angle.X, "0.000000") & " " & Format(.Rotation_Angle.Y, "0.000000") & " " & Format(.Rotation_Angle.Z, "0.000000") & " " _
                              & Format(.Scale.X, "0.000000") & " " & Format(.Scale.Y, "0.000000") & " " & Format(.Scale.Z, "0.000000")
         End With
         
     OutputMeshLine = temLine

    '------------------------------------------------
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputMeshLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputTroopLine
'**输    入：(Long)Trp_Idx[如为-1则指当前Troop]
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-30 11:36:44
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputTroopLine(ByVal Trp_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim TroopVes As Type_Troops, temLine As String, tmpStr As String, m As Integer

If Trp_Idx >= 0 Then
    TroopVes = Trps(Trp_Idx)
Else
    TroopVes = CurrentTrp
End If

With TroopVes

            'Check Value
            If Left$(.strID, 4) <> "trp_" Then
                .strID = "trp_" & .strID
            End If

            temLine = .strID & " " & .strName & " " & .strPtName & " " & .unknown_warband(1) & " " & .Flags & _
                  " " & CStr(.Scene) & " " & CStr(.reserved) & " " & CStr(.Faction) & " " & .Upgrade1 & " " & .Upgrade2 & vbCrLf
                  
            tmpStr = "  "
            For m = 1 To 64
                tmpStr = tmpStr & CStr(.lstInventory(m).X) & " " & CStr(.lstInventory(m).Y) & " "
            Next m
            temLine = temLine & tmpStr & vbCrLf

            temLine = temLine & "  " & CStr(.tAttrib.strPoint) & " " & CStr(.tAttrib.agiPoint) & _
                  " " & CStr(.tAttrib.intPoint) & " " & CStr(.tAttrib.chaPoint) & " " & CStr(.tAttrib.level) & vbCrLf

            temLine = temLine & " " & .WP.one_handed & " " & .WP.two_handed & " " & .WP.polearm & _
                  " " & .WP.archery & " " & .WP.crossbow & " " & .WP.throwing & " " & .WP.firearm & vbCrLf

            temLine = temLine & "" & CStr(.Skills(1)) & " " & CStr(.Skills(2)) & " " & CStr(.Skills(3)) & _
                  " " & CStr(.Skills(4)) & " " & CStr(.Skills(5)) & " " & CStr(.Skills(6)) & " " & vbCrLf

            temLine = temLine & "  " & .Face(1) & " " & .Face(2) & " " & .Face(3) & " " & .Face(4) & " " & _
                  .Face(5) & " " & .Face(6) & " " & .Face(7) & " " & .Face(8) & " " & vbCrLf

            
        End With
        
        OutputTroopLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputTroopLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputFactionLine
'**输    入：(Long)Faction_Idx[如为-1则指当前Faction]
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-30 11:36:44
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputFactionLine(ByVal Faction_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim FactionVes As Type_Faction, temLine As String, tmpStr As String, m As Integer

If Faction_Idx >= 0 Then
    FactionVes = Factions(Faction_Idx)
Else
    FactionVes = CurrentFaction
End If

With FactionVes

            'Check Value
            If Left$(.strID, 4) <> "fac_" Then
                .strID = "fac_" & .strID
            End If

            temLine = .strID & " " & .strName & " " & CStr(.Flags) & " " & CStr(.lColor) & " " & vbCrLf

            tmpStr = " "
            For m = 0 To UBound(.RelationShip)
                tmpStr = tmpStr & Format(.RelationShip(m).Value, "0.000000") & " "
            Next m
            temLine = temLine & tmpStr & vbCrLf

            temLine = temLine & .reserved & " "
            
        End With
        
        OutputFactionLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputFactionLine", Err.Number, Err.Description)
End Function


'*************************************************************************
'**函 数 名：OutputMapIconLine
'**输    入：(Long)MapIcon_Idx[如为-1则指当前MapIcon]
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 11:15:06
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputMapIconLine(ByVal MapIcon_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim MapIconVes As Type_MapIcon, temLine As String, tmpStr As String, m As Integer

If MapIcon_Idx >= 0 Then
    MapIconVes = MapIcons(MapIcon_Idx)
Else
    MapIconVes = CurrentMapIcon
End If

With MapIconVes


            temLine = .strID & " " & CStr(.Flags) & " " & .MeshName & " " & .mScale & " " & .Sound & " "

            For m = 0 To UBound(.Offset)
               temLine = temLine & .Offset(m) & " "
            Next m

            temLine = temLine & .TriggerCount
            
            If .TriggerCount > 0 Then
            temLine = temLine & vbCrLf
               For m = 1 To .TriggerCount
                 tmpStr = OutputTriggerLine(.Triggers(m))
                 temLine = temLine & tmpStr
               Next m
            End If
            
            temLine = temLine & vbCrLf & vbCrLf
        End With
        
        OutputMapIconLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputMapIconLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputOperationLine
'**输    入：(Type_Op_Block)OperationVes
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 11:13:10
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputOperationLine(OperationVes As Type_Op_Block) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim temLine As String, m As Integer

With OperationVes

            temLine = .Op & " " & .ParaNum & " "
         
       If .ParaNum > 0 Then
            For m = 1 To .ParaNum
                temLine = temLine & .Para(m).Value & " "
            Next m
       End If
            
End With
        
        OutputOperationLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputOperationLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputTriggerLine
'**输    入：(Type_Op_Block)TriggerVes
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 11:21:05
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputTriggerLine(TriggerVes As Type_Trigger) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim temLine As String, tmpStr As String, m As Integer

With TriggerVes

            temLine = Format(.tiOn, "0.000000") & "  " & .ActNum & " "
         
       If .ActNum > 0 Then
            For m = 1 To .ActNum
                tmpStr = OutputOperationLine(.tiAct(m))
                temLine = temLine & tmpStr
            Next m
       End If
            
End With
        
        OutputTriggerLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputTriggerLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputPartyLine
'**输    入：(Long)Party_Idx[如为-1则指当前Party]
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-30 11:36:44
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputPartyLine(ByVal Party_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim PartyVes As Type_Party, temLine As String, tmpStr As String, m As Integer

If Party_Idx >= 0 Then
    PartyVes = Parties(Party_Idx)
Else
    PartyVes = CurrentParty
End If

With PartyVes
            'Check Value
            If Left$(.strID, 2) <> "p_" Then
                .strID = "p_" & .strID
            End If
                temLine = " " & (.UnknownTitle) & " " & (.ID) & " " & (.id2) & " " & (.strID) & " " & (.strName) & " " & (.Flags) & " " & (.Menu) & " " & (.Template) & " " & (.Faction) & " "
                
                For m = 1 To 2
                     temLine = temLine & .Personality(m) & " "
                Next m
                
                temLine = temLine & .AI_Behavior & " " & .AI_Target & " " & .reserved & " "
                
                For m = 1 To 3
                     temLine = temLine & .InitPos(m).X & " " & .InitPos(m).Y & " "
                Next m
                
                temLine = temLine & .UnknownStr & " " & .StacksCount & " "
                
                
                If .StacksCount > 0 Then
                     For m = 1 To .StacksCount
                          temLine = temLine & .Stacks(m).ID & " "
                          temLine = temLine & .Stacks(m).Min & " "
                          temLine = temLine & .Stacks(m).Max & " "
                          temLine = temLine & .Stacks(m).Flags & " "
                     Next m
                End If
                
                temLine = temLine & vbCrLf & .Degree
                
            End With
        OutputPartyLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputPartyLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputSceneLine
'**输    入：(Long)Scene_Idx[如为-1则指当前Scene]
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-10 22:20:38
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputSceneLine(ByVal Scene_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim SceneVes As Type_Scene, temLine As String, tmpStr As String, m As Integer

If Scene_Idx >= 0 Then
    SceneVes = Scenes(Scene_Idx)
Else
    SceneVes = CurrentScene
End If

With SceneVes
            'Check Value
            If Left$(.strID, 4) <> "scn_" Then
                .strID = "scn_" & .strID
            End If
                temLine = (.strID) & " " & (.strName) & " " & (.Flags) & " " & (.MeshName) & " " & (.BodyName) & " "
                
                For m = 0 To 1
                     temLine = temLine & Format(.p(m).X, "0.000000") & " " & Format(.p(m).Y, "0.000000") & " "
                Next m
                
                temLine = temLine & Format(.WaterLevel, "0.000000") & " " & .TerrainCode & " " & vbCrLf
                
                temLine = temLine & "  " & .AccessCount & " "
                
                If .AccessCount > 0 Then
                     For m = 1 To .AccessCount
                          temLine = temLine & " " & .Accesses(m) & " "
                     Next m
                End If
                
                temLine = temLine & vbCrLf & "  " & .ChestCount & " "
                
                If .ChestCount > 0 Then
                     For m = 1 To .ChestCount
                          temLine = temLine & " " & .Chests(m).ID & " "
                     Next m
                End If
                temLine = temLine & vbCrLf & " " & .Outer_Terrain_Type & " "
                
            End With
        OutputSceneLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputSceneLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputSoundLine
'**输    入：(Long)Sound_Idx[如为-1则指当前Sound]
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-10 22:20:38
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputSoundLine(ByVal Sound_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim SoundVes As Type_Sound, temLine As String, tmpStr As String, m As Integer

If Sound_Idx >= 0 Then
    SoundVes = Sounds(Sound_Idx)
Else
    SoundVes = CurrentSound
End If

With SoundVes
            'Check Value
            If Left$(.sndName, 4) <> "snd_" Then
                .sndName = "snd_" & .sndName
            End If
                temLine = (.sndName) & " " & (.Flags)
                
                temLine = temLine & " " & .ResourceCount & " "
                
                If .ResourceCount > 0 Then
                     For m = 1 To .ResourceCount
                          temLine = temLine & .Resource(m).ID & " " & .Resource(m).Unknown & " "
                     Next m
                End If
                
            End With
        OutputSoundLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputSoundLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：OutputSoundResLine
'**输    入：(Long)SoundRes_Idx[如为-1则指当前SoundRes]
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-03 15:52:59
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputSoundResLine(ByVal SoundRes_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim SoundResVes As Type_SoundResource, temLine As String, m As Integer

If SoundRes_Idx >= 0 Then
    SoundResVes = SoundRess(SoundRes_Idx)
Else
    SoundResVes = CurrentSoundRes
End If

With SoundResVes

                temLine = " " & (.sndName) & " " & (.Flags)
                
            End With
        OutputSoundResLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputSoundResLine", Err.Number, Err.Description)
End Function


'*************************************************************************
'**函 数 名：OutputPTLine
'**输    入：(Long)PT_Idx[如为-1则指当前PT]
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-30 11:36:44
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Function OutputPTLine(ByVal PT_Idx As Long) As String
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
Dim PTVes As Type_PT, temLine As String, tmpStr As String, m As Integer

If PT_Idx >= 0 Then
    PTVes = PTs(PT_Idx)
Else
    PTVes = CurPartyTemplate
End If

With PTVes
            'Check Value
            If Left$(.ptID, 3) <> "pt_" Then
                .ptID = "pt_" & .ptID
            End If
                temLine = (.ptID) & " " & (.ptName) & " " & (.Flags) & " " & (.Menu) & " " & (.Faction) & " " & (.Personality) & " "

                For m = 1 To 6
                    If .Stacks(m).ID < 0 Then
                        temLine = temLine & ("-1 ")
                    Else
                        temLine = temLine & (.Stacks(m).ID) & " "
                      If .Stacks(m).ID >= 0 Then
                        temLine = temLine & (.Stacks(m).Min) & " "

                        temLine = temLine & (.Stacks(m).Max) & " "
                        
                        temLine = temLine & (.Stacks(m).Flags) & " "
                      End If
                    End If
                Next
                
            End With
        OutputPTLine = temLine
    Exit Function
    '----------------
errorHandle:
    Call logErr("FileSaver", "OutputPTLine", Err.Number, Err.Description)
End Function

'*************************************************************************
'**函 数 名：CheckEditable
'**输    入：(Long)QueryType,(Long)Index
'**输    出：(Boolean)
'**功能描述：
'**全局变量：N_Item
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-30 14:16:59
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function CheckEditable(ByVal QueryType As Long, ByVal Index As Long) As Boolean

Public Sub SaveItemCSVFile(FilePath As String)
     On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SaveItemCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Item - 1
       If itm(i).csvName <> itm(i).disname Then
          strTem = strTem & itm(i).dbName & "|" & itm(i).csvName & vbCrLf
       End If
       
       If itm(i).csvName_pl <> itm(i).disname Then
          strTem = strTem & itm(i).dbName & "_pl" & "|" & itm(i).csvName_pl & vbCrLf
       End If
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("FileSaver", "SaveItemCSVFile", "INVALID_HANDLE_VALUE", "无法打开文件:[" & FilePath & "]")
    End If
    
     Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveItemCSVFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveStringCSVFile
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

Public Sub SaveStringCSVFile(FilePath As String)
     On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SaveStringCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Str - 1
       If Strs(i).CSV <> "" Then
          strTem = strTem & Strs(i).Name & "|" & Strs(i).CSV
          If i <> N_Str - 1 Then
             strTem = strTem & vbCrLf
          End If
       End If
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("FileSaver", "SaveStringCSVFile", "INVALID_HANDLE_VALUE", "无法打开文件:[" & FilePath & "]")
    End If
    
     Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveStringCSVFile", Err.Number, Err.Description)
End Sub
'*************************************************************************
'**函 数 名：SaveFactionCSVFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-30 23:24:31
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SaveFactionCSVFile(FilePath As String)
     On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
    Dim strTem As String, i As Long, q As Boolean
    
    If LCase$(MnBInfo.Language) = "en" Then
         'SaveFactionCSVFile = True
         Exit Sub
    End If
    
    For i = 0 To N_Faction - 1
       If Factions(i).csvName <> Factions(i).strName Then
          strTem = strTem & Factions(i).strID & "|" & Factions(i).csvName & vbCrLf
       End If
       
    Next i
    
    q = UEFSaveTextFile(FilePath, strTem, False, UEF_UTF8, UEF_UTF8)
    
    If q = False Then
        Call logErr("FileSaver", "SaveFactionCSVFile", "INVALID_HANDLE_VALUE", "无法打开文件:[" & FilePath & "]")
    End If
    
     Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveFactionCSVFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveFactionFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-04 23:34:05
'**版    本：V1.1321
'*************************************************************************

Sub SaveFactionFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim tmpStr As String
    Dim n As Long, i As Long
    Dim m As Integer

    lngHandle = FreeFile()

    Open FilePath For Output As #lngHandle

    tmpStr = ""
    For n = 0 To 1
        tmpStr = tmpStr & FactionVersionInform(n) & " "
    Next n
    tmpStr = tmpStr & FactionVersionInform(2)
    Print #lngHandle, tmpStr
    
    Print #lngHandle, Trim$(Str$(N_Faction))
    
    tmpStr = ""
    For n = 0 To N_Faction - 1
            tmpStr = tmpStr & OutputFactionLine(n)
    Next n
    Print #lngHandle, tmpStr
    Close #lngHandle

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveFactionFile", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：SaveMapIconFile
'**输    入：FilePath(String) -
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-11 11:28:20
'**版    本：V1.1321
'*************************************************************************

Sub SaveMapIconFile(FilePath As String)
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------

    Dim tmpStr As String
    Dim n As Long, i As Long
    Dim m As Integer

    lngHandle = FreeFile()

    Open FilePath For Output As #lngHandle

    tmpStr = ""
    For n = 0 To 1
        tmpStr = tmpStr & MapIconVersionInform(n) & " "
    Next n
    tmpStr = tmpStr & MapIconVersionInform(2)
    Print #lngHandle, tmpStr
    
    Print #lngHandle, Trim$(Str$(N_MapIcon))
    
    For n = 0 To N_MapIcon - 1
            tmpStr = OutputMapIconLine(n)
            Print #lngHandle, tmpStr
    Next n
   
    Close #lngHandle

    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("FileSaver", "SaveMapIconFile", Err.Number, Err.Description)
End Sub


'*************************************************************************
'**函 数 名：LoadIModFile
'**输    入：FileName(String) -
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-12-01 23:14:15
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Public Sub SaveAll()           'A_E
    'Save Quote
     SaveQuote
     
    'Save Items
     SaveItemFile MnBInfo.ModPath & "\item_kinds1.txt"
     SaveItemCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\item_kinds.csv"

    'Save Troops
     SaveTroopFile MnBInfo.ModPath & "\troops.txt"
     SaveTroopCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\troops.csv"
     
    'Save Factions
    SaveFactionFile MnBInfo.ModPath & "\factions.txt"
    SaveFactionCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\factions.csv"
    
    'Save PartyTemplates
    SavePartyTemplateFile MnBInfo.ModPath & "\party_templates.txt"
    SavePartyTemplateCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\party_templates.csv"
    
    'Save Parties
    SavePartyFile MnBInfo.ModPath & "\parties.txt"
    SavePartyCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\parties.csv"
    
    'Save Scenes
    SaveSceneFile MnBInfo.ModPath & "\scenes.txt"
    
    'Save MapIcons
    SaveMapIconFile MnBInfo.ModPath & "\map_icons.txt"
    
    'Save Sound & SoundResource
    SaveSoundFile MnBInfo.ModPath & "\sounds.txt"
    
    'Save Particles System
    SavePSysFile MnBInfo.ModPath & "\particle_systems.txt"
    
    'Save Tableau Materials
    SaveTabMatFile MnBInfo.ModPath & "\tableau_materials.txt"
   
    'Save Meshes
    SaveMeshFile MnBInfo.ModPath & "\meshes.txt"
   
   'Save Triggers
   SaveTriggerFile MnBInfo.ModPath & "\triggers.txt"
   
   'Save Strings
   If IsLoadString Then
      SaveStringFile MnBInfo.ModPath & "\strings.txt"
      SaveStringCSVFile MnBInfo.ModPath & "\languages\" & MnBInfo.Language & "\game_strings.csv"
   End If
   
   'Save
   Dim i As Long
   For i = 1 To UBound(VarNameLists)  '0 global
      SaveVarNameCheckList i
   Next i
   
End Sub

'*************************************************************************
'**函 数 名：ReadAll
'**输    入： -
'**输    出： -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-22 15:42:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

