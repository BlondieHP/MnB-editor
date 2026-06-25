Attribute VB_Name = "ModData"
Option Explicit

Public Const ERROR_NOELEMENT = 35601
Public Const ERROR_REPEATELEMENT = 35602
Public Const ERROR_INVALIDELEMENT = 35603

Public Const SLOT_NULL = 0
Public Const SLOT_NEW_INDEX = 1
Public Const SLOT_VARIABLE_NAME_CHECK_LIST = 2

Public Function AddIndex(ByVal ID As Long, ByVal strID As String) As Boolean
On Error GoTo EL
Dim oItem As ListItem

Set oItem = frmData.LstIndex.ListItems.Add(, "ID_" & strID, ID)

AddIndex = True

Exit Function

EL:
AddIndex = False

If Err.Number = ERROR_REPEATELEMENT Then
   If frmData.LstIndex.ListItems("ID_" & strID).SubItems(1) <> "" Then
       frmData.LstIndex.ListItems("ID_" & strID).Text = ID
       frmData.LstIndex.ListItems("ID_" & strID).SubItems(1) = ""
       frmData.LstIndex.ListItems("ID_" & strID).SubItems(2) = ""
       AddIndex = True
   End If
End If
End Function

Public Function DelIndex(ByVal strID As String, Optional RootOut As Boolean = True) As Boolean
On Error GoTo EL
Dim NewID(1) As String, StartID As String

StartID = "ID_" & strID

NewID(0) = frmData.LstIndex.ListItems(StartID).SubItems(1)

frmData.LstIndex.ListItems.Remove StartID

If RootOut Then
Search:
      If NewID(0) <> "" Then
         NewID(1) = frmData.LstIndex.ListItems(NewID(0)).SubItems(1)
         frmData.LstIndex.ListItems.Remove NewID(0)
         NewID(0) = NewID(1)
         GoTo Search
      End If
End If

DelIndex = True
Exit Function

EL:

If Err.Number = ERROR_NOELEMENT Then
      '已删除
End If

DelIndex = False
End Function

Public Function ChangeStrID(ByVal strID As String, ByVal New_strID As String) As Boolean
On Error GoTo Errline
Dim oItem As ListItem, StartID As String
Dim NewID As String, Index_Now As Long

Index_Now = frmData.LstIndex.ListItems("ID_" & strID).Index

Search:
NewID = frmData.LstIndex.ListItems(Index_Now).SubItems(1)
If NewID = "" Then
   StartID = frmData.LstIndex.ListItems(Index_Now).Key
Else
   Index_Now = frmData.LstIndex.ListItems(NewID).Index
   GoTo Search
End If

Set oItem = frmData.LstIndex.ListItems.Add(, "ID_" & New_strID)

frmData.LstIndex.ListItems("ID_" & New_strID).Text = frmData.LstIndex.ListItems(StartID).Text
frmData.LstIndex.ListItems("ID_" & New_strID).SubItems(2) = frmData.LstIndex.ListItems(StartID).SubItems(2)

frmData.LstIndex.ListItems(StartID).SubItems(1) = "ID_" & New_strID

ChangeStrID = True
Exit Function

Errline:
ChangeStrID = False
End Function

Public Function ChangeID(ByVal strID As String, ByVal ID As Long) As Boolean
On Error GoTo Errline

frmData.LstIndex.ListItems("ID_" & strID).Text = ID

ChangeID = True
Exit Function

Errline:
ChangeID = False
End Function

Public Function SwapID(ByVal strID1 As String, ByVal strID2 As String) As Boolean
On Error GoTo Errline
Dim t(1) As Long, q As Boolean

t(0) = GetID(strID1)
t(1) = GetID(strID2)

q = True
q = q And ChangeID(strID1, t(1))
q = q And ChangeID(strID2, t(0))

SwapID = q
Exit Function

Errline:
SwapID = False
End Function

Public Function GetID(strID As String, Optional ShowError As Boolean = True, Optional DefaultStr As String, Optional DefaultVal As Long = 0) As Long
On Error GoTo EL
Dim NewID As String, Index_Now As Long

Index_Now = frmData.LstIndex.ListItems("ID_" & strID).Index

Search:
NewID = frmData.LstIndex.ListItems(Index_Now).SubItems(1)
If NewID = "" Then
   GetID = CLng(Val(frmData.LstIndex.ListItems(Index_Now).Text))
   Exit Function
Else
   Index_Now = frmData.LstIndex.ListItems(NewID).Index
   GoTo Search
End If

Exit Function

EL:

If Err.Number = ERROR_NOELEMENT Then
    If ShowError Then
        MsgBox "项目:[" & strID & "]不存在,之前已被手动删除。现在将值重置为[" & DefaultVal & "]。", vbExclamation, "Error"
        GetID = DefaultVal
        strID = DefaultStr
    Else
        GetID = DefaultVal
    End If
Else
    logErr "ModData", "GetID", Err.Number, Err.Description
    GetID = DefaultVal
End If

End Function


Public Function GetSlot(strID As String, Slot_No As Long, Optional ShowError As Boolean = True, Optional DefaultStr As String, Optional DefaultVal As Long = 0) As Long
On Error GoTo EL
Dim NewID As String, Index_Now As Long

Index_Now = frmData.LstIndex.ListItems("ID_" & strID).Index

Search:
NewID = frmData.LstIndex.ListItems(Index_Now).SubItems(1)
If NewID = "" Then
   GetSlot = Val(frmData.LstIndex.ListItems(Index_Now).SubItems(Slot_No))
   Exit Function
Else
   Index_Now = frmData.LstIndex.ListItems(NewID).Index
   GoTo Search
End If

Exit Function

EL:

If Err.Number = ERROR_NOELEMENT Then
    If ShowError Then
        MsgBox "项目:[" & strID & "]不存在,之前已被手动删除。现在将值重置为[" & DefaultVal & "]。", vbExclamation, "Error"
        GetSlot = DefaultVal
        strID = DefaultStr
    Else
        GetSlot = DefaultVal
    End If
Else
    logErr "ModData", "GetSlot", Err.Number, Err.Description
    GetSlot = DefaultVal
End If

End Function

Public Function SetSlot(strID As String, Slot_No As Long, Value As String, Optional ShowError As Boolean = True) As Boolean
On Error GoTo EL
Dim NewID As String, Index_Now As Long

SetSlot = True
Index_Now = frmData.LstIndex.ListItems("ID_" & strID).Index

Search:
NewID = frmData.LstIndex.ListItems(Index_Now).SubItems(1)
If NewID = "" Then
   frmData.LstIndex.ListItems(Index_Now).SubItems(Slot_No) = Value
   Exit Function
Else
   Index_Now = frmData.LstIndex.ListItems(NewID).Index
   GoTo Search
End If

Exit Function

EL:

SetSlot = False
If Err.Number = ERROR_NOELEMENT Then
    If ShowError Then
        MsgBox "项目:[" & strID & "]不存在,之前已被手动删除。", vbExclamation, "Error"
    End If
Else
    logErr "ModData", "SetSlot", Err.Number, Err.Description
End If

End Function

Public Function KeytoStrID(ByVal sKey As String) As String
If Len(sKey) > 3 Then
    If UCase(Left(sKey, 3)) = "ID_" Then
       KeytoStrID = Right(sKey, Len(sKey) - 3)
    End If
End If
End Function

Public Sub BuildQuote()
Dim i As Long, j As Integer, n As Integer, strTem As String, lngTem As Long, I64(0) As Integer64b

For i = 0 To N_Troop - 1
    With Trps(i)
         .Faction_strID = Factions(.Faction).strID
         .Upgrade1_strID = Trps(.Upgrade1).strID
         .Upgrade2_strID = Trps(.Upgrade2).strID

         PurseSceneInTroop .Scene, .SceneID, .Scene_strID, .Entry
         
         For j = 1 To 64
            If .lstInventory(j).X > -1 Then
             .lstInventory(j).strX = itm(.lstInventory(j).X).dbName
            End If
         Next j
    End With
Next i

For i = 0 To N_Item - 1
    With itm(i)
          For j = 1 To .FactionCount
              .Faction(j).strID = Factions(.Faction(j).ID).strID
          Next j
          
          For j = 1 To .TriggerCount
             For n = 1 To .Trigger(j).ActNum
                 BuildQuote_Op_Blocks .Trigger(j).tiAct(n)
             Next n
          Next j
    End With
Next i

For i = 0 To N_Faction - 1
    With Factions(i)
         For j = 0 To N_Faction - 1
              .RelationShip(j).ID = j
              .RelationShip(j).strID = Factions(.RelationShip(j).ID).strID
         Next j
    End With
Next i

For i = 0 To N_Party - 1
    With Parties(i)
         .Template_strID = PTs(.Template).ptID
         .Faction_strID = Factions(.Faction).strID
         .AI_Target_strID = Parties(.AI_Target).strID
         
         I64(0) = StrToI64(.Flags)
         .MapIcon_strID = MapIcons(I64(0).by(0)).strID

         For j = 1 To .StacksCount
              .Stacks(j).strID = Trps(.Stacks(j).ID).strID
         Next j
         
    End With
Next i

For i = 0 To N_PT - 1
    With PTs(i)
         .Faction_strID = Factions(.Faction).strID
         
         For j = 1 To 6
             If .Stacks(j).ID > -1 Then
               .Stacks(j).strID = Trps(.Stacks(j).ID).strID
             Else
               .Stacks(j).strID = ""
             End If
         Next j
    End With
Next i

For i = 0 To N_MapIcon - 1
    With MapIcons(i)
         .Sound_sndName = Sounds(.Sound).sndName
         
          For j = 1 To .TriggerCount
             For n = 1 To .Triggers(j).ActNum
                 BuildQuote_Op_Blocks .Triggers(j).tiAct(n)
             Next n
          Next j
    End With
Next i

For i = 0 To N_Sound - 1
    With Sounds(i)
         For j = 1 To .ResourceCount
            .Resource(j).strID = SoundRess(.Resource(j).ID).sndName
         Next j
    End With
Next i

For i = 0 To N_Scene - 1
    With Scenes(i)
         For j = 1 To .ChestCount
            .Chests(j).strID = Trps(.Chests(j).ID).strID
         Next j
    End With
Next i

For i = 0 To N_TabMat - 1
    With TabMat(i)
         For j = 1 To .OpCount
            BuildQuote_Op_Blocks .OpBlock(j)
         Next j
    End With
Next i

For i = 0 To N_TimeTrg - 1
    With TimeTrg(i)
         For j = 1 To .ConditionsCount
            BuildQuote_Op_Blocks .Condition(j)
         Next j
         
         For j = 1 To .ConsequencesCount
            BuildQuote_Op_Blocks .Consequence(j)
         Next j
    End With
Next i

End Sub


Public Sub SaveQuote()
Dim i As Long, j As Integer, n As Integer, strHex(1) As String, strTem As String, lngTem As String, I64(0) As Integer64b, q As Boolean

For i = 0 To N_Troop - 1
    With Trps(i)
         .Faction = GetID(.Faction_strID, False)
         .Upgrade1 = GetID(.Upgrade1_strID, False)
         .Upgrade2 = GetID(.Upgrade2_strID, False)

         .SceneID = GetID(.Scene_strID, False)
         If .SceneID = 0 Then .Entry = 0
         .Scene = Val("&H" & Hex(.Entry) & Right("0000" & Hex(.SceneID), 4))
         
         q = False
         For j = 1 To 64
            If .lstInventory(j).strX <> "" Then
              .lstInventory(j).X = GetID(.lstInventory(j).strX, False, "", -1)
              If .lstInventory(j).X = -1 Then q = True
            Else
              .lstInventory(j).X = -1
            End If
            
            If q Then StructureTroopInventory i
         Next j
    End With
Next i

For i = 0 To N_Item - 1
    With itm(i)
         q = False
         For j = 1 To .FactionCount
             .Faction(j).ID = GetID(.Faction(j).strID, False, "", -1)
             If .Faction(j).ID = -1 Then q = True
         Next j
         
         If q Then StructureItemFactions i
         
         For j = 1 To .TriggerCount
             For n = 1 To .Trigger(j).ActNum
                 SaveQuote_Op_Blocks .Trigger(j).tiAct(n)
             Next n
         Next j
    End With
Next i

For i = 0 To N_Faction - 1
    With Factions(i)
          StructureFactionRelationShips i
    End With
Next i

For i = 0 To N_Party - 1
    q = False
    With Parties(i)
         .Template = GetID(.Template_strID, False)
         .Faction = GetID(.Faction_strID, False)
         .AI_Target = GetID(.AI_Target_strID, False)
         
         I64(0) = StrToI64(.Flags)
         I64(0).by(0) = GetID(.MapIcon_strID, False)
         .Flags = I64toStr(I64(0))
         
         For j = 1 To .StacksCount
              .Stacks(j).ID = GetID(.Stacks(j).strID, False, "", -1)
              
              If .Stacks(j).ID = -1 Then
                 q = True
              End If
         Next j
         
         If q Then StructurePartyStacksEx i
    End With
Next i

For i = 0 To N_PT - 1
    With PTs(i)
         .Faction = GetID(.Faction_strID, False)
         For j = 1 To 6
            If .Stacks(j).strID <> "" Then
              .Stacks(j).ID = GetID(.Stacks(j).strID, False, , -1)
            Else
              .Stacks(j).ID = -1
            End If
         Next j
    End With
Next i

For i = 0 To N_MapIcon - 1
    With MapIcons(i)
         .Sound = GetID(.Sound_sndName, False)
         
          For j = 1 To .TriggerCount
             For n = 1 To .Triggers(j).ActNum
                 SaveQuote_Op_Blocks .Triggers(j).tiAct(n)
             Next n
          Next j
    End With
Next i

For i = 0 To N_Sound - 1
    q = False
    With Sounds(i)
         For j = 1 To .ResourceCount
            .Resource(j).ID = GetID(.Resource(j).strID, False, "", -1)
            If .Resource(j).ID = -1 Then q = True
         Next j
    End With
    
    If q Then
        StructureSoundRes i
    End If
Next i

For i = 0 To N_Scene - 1
    q = False
    With Scenes(i)
         For j = 1 To .ChestCount
            .Chests(j).ID = GetID(.Chests(j).strID, False, "", -1)
            If .Chests(j).ID = -1 Then q = True
         Next j
         
         If q Then
            StructureSceneChests i
         End If
    End With
Next i

For i = 0 To N_TabMat - 1
    With TabMat(i)
         For j = 1 To .OpCount
            SaveQuote_Op_Blocks .OpBlock(j)
         Next j
    End With
Next i

For i = 0 To N_TimeTrg - 1
    With TimeTrg(i)
         For j = 1 To .ConditionsCount
            SaveQuote_Op_Blocks .Condition(j)
         Next j
         
         For j = 1 To .ConsequencesCount
            SaveQuote_Op_Blocks .Consequence(j)
         Next j
    End With
Next i

End Sub

Private Sub SaveQuote_Op_Blocks(Op_Block As Type_Op_Block)
Dim i As Long, strTem As String, lngTem As Long, Tag_No As Integer

     For i = 1 To Op_Block.ParaNum
         With Op_Block.Para(i)
              If .strID <> "" Then
                 Tag_No = QuickGetParamType(.Value)
                 lngTem = GetID(.strID, False, GetstrID(Tag_No, "0"), 0)
                 .Value = getTXTID(Tag_No, lngTem)
              End If
         End With
     Next i
End Sub

Private Sub BuildQuote_Op_Blocks(Op_Block As Type_Op_Block)
Dim i As Long, strTem As String, Tag_No As Integer

     For i = 1 To Op_Block.ParaNum
         With Op_Block.Para(i)
                 QuickGetParamCodeInfo .Value, Tag_No, strTem
                 .strID = GetstrID(Tag_No, strTem)
         End With
     Next i
End Sub

Public Sub BuildQuote_ParamCode(ParamCode As Type_Param)
Dim i As Long, strTem As String, Tag_No As Integer

With ParamCode
   GetParamCodeInfo .Value, Tag_No, strTem
   .strID = GetstrID(Tag_No, strTem)
End With

End Sub

Public Function SectionExist(ByVal Section As String) As Boolean
On Error GoTo EL

Dim strSec() As String, i As Integer, a As String

If Trim(Section) = "" Then Exit Function
strSec() = Split(Section, "\")

For i = 0 To UBound(strSec)
    
Next i

Exit Function

EL:

End Function

Public Function CreateSection(ByVal Section As String) As Long
On Error GoTo EL

Dim strSec() As String, i As Integer

If Trim(Section) = "" Then Exit Function
strSec() = Split(Section, "\")

For i = 0 To UBound(strSec)
  
Next i

Exit Function

EL:

End Function

Public Function RegisterValue(ByVal Section As String, ByVal Value As String) As Long

End Function
