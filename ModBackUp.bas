Attribute VB_Name = "ModBackUp"
Option Explicit

'*************************************************************************
'**函 数 名：SetBackUp
'**输    入：-
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-01 22:22:22 ←_←
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function SetBackUp(Optional FileName As String = "") As Boolean
On Error GoTo EL

Dim LanSource As String, strBackUp As String
Dim strTime As String

     strTime = Format(Now, "yyyy_mm_dd hh_mm_ss")
     
     If FileName = "" Then
       strBackUp = MnBInfo.ModBackUp & "\" & strTime
     Else
       strBackUp = MnBInfo.ModBackUp & "\" & FileName
     End If
     
     If Not DirExists(strBackUp) Then
       MkDirEx strBackUp
     End If
     
     LanSource = MnBInfo.ModPath & "\languages\" & MnBInfo.Language
     'Copy Items
     FileCopy MnBInfo.ModPath & "\item_kinds1.txt", strBackUp & "\item_kinds1.txt"
     If FileExists(LanSource & "\item_kinds.csv") Then
       FileCopy LanSource & "\item_kinds.csv", strBackUp & "\item_kinds.csv"
     End If
     
    'Copy Troops
     FileCopy MnBInfo.ModPath & "\troops.txt", strBackUp & "\troops.txt"
     If FileExists(LanSource & "\troops.csv") Then
        FileCopy LanSource & "\troops.csv", strBackUp & "\troops.csv"
     End If
     
    'Copy Factions
    FileCopy MnBInfo.ModPath & "\factions.txt", strBackUp & "\factions.txt"
    If FileExists(LanSource & "\factions.csv") Then
       FileCopy LanSource & "\factions.csv", strBackUp & "\factions.csv"
    End If
    
    'Copy PartyTemplates
    FileCopy MnBInfo.ModPath & "\party_templates.txt", strBackUp & "\party_templates.txt"
    If FileExists(LanSource & "\party_templates.csv") Then
       FileCopy LanSource & "\party_templates.csv", strBackUp & "\party_templates.csv"
    End If
    
    'Copy Parties
    FileCopy MnBInfo.ModPath & "\parties.txt", strBackUp & "\parties.txt"
    If FileExists(LanSource & "\parties.csv") Then
       FileCopy LanSource & "\parties.csv", strBackUp & "\parties.csv"
    End If
    
    'Copy Scenes
    FileCopy MnBInfo.ModPath & "\scenes.txt", strBackUp & "\scenes.txt"
    
    'Copy MapIcons
    FileCopy MnBInfo.ModPath & "\map_icons.txt", strBackUp & "\map_icons.txt"
    
    'Copy Sounds
    FileCopy MnBInfo.ModPath & "\sounds.txt", strBackUp & "\sounds.txt"
    
    'Copy Particle System
    FileCopy MnBInfo.ModPath & "\particle_systems.txt", strBackUp & "\particle_systems.txt"
    
    'Copy Tableau Materials
    FileCopy MnBInfo.ModPath & "\tableau_materials.txt", strBackUp & "\tableau_materials.txt"
    
    'Copy Meshes
    FileCopy MnBInfo.ModPath & "\meshes.txt", strBackUp & "\meshes.txt"
    
SetBackUp = True

Exit Function

EL:
Call logErr("ModBackUp", "SetBackUp", Err.Number, Err.Description)
End Function


'*************************************************************************
'**函 数 名：RestoreMod
'**输    入：(String)RevTime
'**输    出：(Boolean) -
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-01 22:22:22 ←_←
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function RestoreMod(ByVal RevTime As String) As Boolean
On Error GoTo EL

Dim LanSource As String, strBackUp As String
     
     strBackUp = MnBInfo.ModBackUp & "\" & RevTime
     
     If Not DirExists(strBackUp) Then
       Exit Function
     End If
     
     LanSource = MnBInfo.ModPath & "\languages\" & MnBInfo.Language
     
     If Not DirExists(LanSource) Then
         MkDirEx LanSource
     End If
     
     'Copy Items
     If FileExists(strBackUp & "\item_kinds1.txt") Then
       FileCopy strBackUp & "\item_kinds1.txt", MnBInfo.ModPath & "\item_kinds1.txt"
     End If
     
     If FileExists(strBackUp & "\item_kinds.csv") Then
       FileCopy strBackUp & "\item_kinds.csv", LanSource & "\item_kinds.csv"
     End If
     
    'Copy Troops
     If FileExists(strBackUp & "\troops.txt") Then
       FileCopy strBackUp & "\troops.txt", MnBInfo.ModPath & "\troops.txt"
     End If
     
     If FileExists(strBackUp & "\troops.csv") Then
       FileCopy strBackUp & "\troops.csv", LanSource & "\troops.csv"
     End If
     
    'Copy Factions
     If FileExists(strBackUp & "\factions.txt") Then
       FileCopy strBackUp & "\factions.txt", MnBInfo.ModPath & "\factions.txt"
     End If
     
     If FileExists(strBackUp & "\factions.csv") Then
       FileCopy strBackUp & "\factions.csv", LanSource & "\factions.csv"
     End If
     
    'Copy PartyTemplates
     If FileExists(strBackUp & "\party_templates.txt") Then
       FileCopy strBackUp & "\party_templates.txt", MnBInfo.ModPath & "\party_templates.txt"
     End If
     
     If FileExists(strBackUp & "\party_templates.csv") Then
       FileCopy strBackUp & "\party_templates.csv", LanSource & "\party_templates.csv"
     End If
     
    'Copy Parties
     If FileExists(strBackUp & "\parties.txt") Then
       FileCopy strBackUp & "\parties.txt", MnBInfo.ModPath & "\parties.txt"
     End If
     
     If FileExists(strBackUp & "\parties.csv") Then
       FileCopy strBackUp & "\parties.csv", LanSource & "\parties.csv"
     End If
     
    'Copy Scenes
     If FileExists(strBackUp & "\scenes.txt") Then
       FileCopy strBackUp & "\scenes.txt", MnBInfo.ModPath & "\scenes.txt"
     End If
     
    'Copy MapIcons
     If FileExists(strBackUp & "\map_icons.txt") Then
       FileCopy strBackUp & "\map_icons.txt", MnBInfo.ModPath & "\map_icons.txt"
     End If
     
    'Copy Sounds
     If FileExists(strBackUp & "\sounds.txt") Then
       FileCopy strBackUp & "\sounds.txt", MnBInfo.ModPath & "\sounds.txt"
     End If
     
    'Copy Particle Systems
     If FileExists(strBackUp & "\particle_systems.txt") Then
       FileCopy strBackUp & "\particle_systems.txt", MnBInfo.ModPath & "\particle_systems.txt"
     End If
     
     'Copy Tableau Materials
     If FileExists(strBackUp & "\tableau_materials.txt") Then
       FileCopy strBackUp & "\tableau_materials.txt", MnBInfo.ModPath & "\tableau_materials.txt"
     End If
     
     'Copy Meshes
     If FileExists(strBackUp & "\Meshes.txt") Then
       FileCopy strBackUp & "\Meshes.txt", MnBInfo.ModPath & "\Meshes.txt"
     End If
     
RestoreMod = True

Exit Function

EL:
Call logErr("ModBackUp", "RestoreMod", Err.Number, Err.Description)
End Function
