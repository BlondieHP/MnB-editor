VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'============================================================================
' DataRepository.cls - Centralized data access layer
'
' Wraps all global data arrays (Trps, itm, Factions, etc.) behind a clean,
' consistent API. This is the ONLY place that should directly access the
' global arrays from Models.bas. All forms and business logic should go
' through this class.
'
' Benefits:
'   1. Single point of data access - easy to add validation, logging, caching
'   2. Change notification events - UI can react to data changes
'   3. Migration-ready - only this class needs to be rewritten when porting
'      to .NET/Python; the API surface remains the same
'============================================================================

' --- Change notification events ---
Public Event TroopChanged(ByVal Index As Long)
Public Event ItemChanged(ByVal Index As Long)
Public Event FactionChanged(ByVal Index As Long)
Public Event PartyChanged(ByVal Index As Long)
Public Event PartyTemplateChanged(ByVal Index As Long)
Public Event SceneChanged(ByVal Index As Long)
Public Event DataLoaded()
Public Event DataSaved()

'============================================================================
' Troops
'============================================================================
Public Property Get TroopCount() As Long
    TroopCount = N_Troop
End Property

Public Property Get Troop(ByVal Index As Long) As Type_Troops
    Troop = Trps(Index)
End Property

Public Property Let Troop(ByVal Index As Long, ByVal Value As Type_Troops)
    Trps(Index) = Value
    RaiseEvent TroopChanged(Index)
End Property

Public Function AddTroop(Optional ByVal NewTroop As Type_Troops) As Long
    N_Troop = N_Troop + 1
    ReDim Preserve Trps(N_Troop - 1)
    ' Initialize defaults if not provided
    AddTroop = N_Troop - 1
    RaiseEvent TroopChanged(AddTroop)
End Function

'============================================================================
' Items
'============================================================================
Public Property Get ItemCount() As Long
    ItemCount = N_Item
End Property

Public Property Get Item(ByVal Index As Long) As Type_Item
    Item = itm(Index)
End Property

Public Property Let Item(ByVal Index As Long, ByVal Value As Type_Item)
    itm(Index) = Value
    RaiseEvent ItemChanged(Index)
End Property

Public Function AddItem() As Long
    N_Item = N_Item + 1
    ReDim Preserve itm(N_Item - 1)
    AddItem = N_Item - 1
    RaiseEvent ItemChanged(AddItem)
End Function

'============================================================================
' Factions
'============================================================================
Public Property Get FactionCount() As Long
    FactionCount = N_Faction
End Property

Public Property Get Faction(ByVal Index As Long) As Type_Faction
    Faction = Factions(Index)
End Property

Public Property Let Faction(ByVal Index As Long, ByVal Value As Type_Faction)
    Factions(Index) = Value
    RaiseEvent FactionChanged(Index)
End Property

Public Function AddFaction() As Long
    N_Faction = N_Faction + 1
    ReDim Preserve Factions(N_Faction - 1)
    AddFaction = N_Faction - 1
    RaiseEvent FactionChanged(AddFaction)
End Function

'============================================================================
' Parties
'============================================================================
Public Property Get PartyCount() As Long
    PartyCount = N_Party
End Property

Public Property Get Party(ByVal Index As Long) As Type_Party
    Party = Parties(Index)
End Property

Public Property Let Party(ByVal Index As Long, ByVal Value As Type_Party)
    Parties(Index) = Value
    RaiseEvent PartyChanged(Index)
End Property

Public Function AddParty() As Long
    N_Party = N_Party + 1
    ReDim Preserve Parties(N_Party - 1)
    AddParty = N_Party - 1
    RaiseEvent PartyChanged(AddParty)
End Function

'============================================================================
' Party Templates
'============================================================================
Public Property Get PartyTemplateCount() As Long
    PartyTemplateCount = N_PT
End Property

Public Property Get PartyTemplate(ByVal Index As Long) As Type_PT
    PartyTemplate = PTs(Index)
End Property

Public Property Let PartyTemplate(ByVal Index As Long, ByVal Value As Type_PT)
    PTs(Index) = Value
    RaiseEvent PartyTemplateChanged(Index)
End Property

Public Function AddPartyTemplate() As Long
    N_PT = N_PT + 1
    ReDim Preserve PTs(N_PT - 1)
    AddPartyTemplate = N_PT - 1
    RaiseEvent PartyTemplateChanged(AddPartyTemplate)
End Function

'============================================================================
' Scenes
'============================================================================
Public Property Get SceneCount() As Long
    SceneCount = N_Scene
End Property

Public Property Get Scene(ByVal Index As Long) As Type_Scene
    Scene = Scenes(Index)
End Property

Public Property Let Scene(ByVal Index As Long, ByVal Value As Type_Scene)
    Scenes(Index) = Value
    RaiseEvent SceneChanged(Index)
End Property

Public Function AddScene() As Long
    N_Scene = N_Scene + 1
    ReDim Preserve Scenes(N_Scene - 1)
    AddScene = N_Scene - 1
    RaiseEvent SceneChanged(AddScene)
End Function

'============================================================================
' Swap operations (centralized)
'============================================================================
Public Sub SwapTroops(ByVal Index1 As Long, ByVal Index2 As Long)
    Dim tmp As Type_Troops
    tmp = Trps(Index1)
    Trps(Index1) = Trps(Index2)
    Trps(Index2) = tmp
    RaiseEvent TroopChanged(Index1)
    RaiseEvent TroopChanged(Index2)
End Sub

Public Sub SwapItems(ByVal Index1 As Long, ByVal Index2 As Long)
    Dim tmp As Type_Item
    tmp = itm(Index1)
    itm(Index1) = itm(Index2)
    itm(Index2) = tmp
    RaiseEvent ItemChanged(Index1)
    RaiseEvent ItemChanged(Index2)
End Sub

Public Sub SwapFactions(ByVal Index1 As Long, ByVal Index2 As Long)
    Dim tmp As Type_Faction
    tmp = Factions(Index1)
    Factions(Index1) = Factions(Index2)
    Factions(Index2) = tmp
    RaiseEvent FactionChanged(Index1)
    RaiseEvent FactionChanged(Index2)
End Sub

Public Sub SwapParties(ByVal Index1 As Long, ByVal Index2 As Long)
    Dim tmp As Type_Party
    tmp = Parties(Index1)
    Parties(Index1) = Parties(Index2)
    Parties(Index2) = tmp
    RaiseEvent PartyChanged(Index1)
    RaiseEvent PartyChanged(Index2)
End Sub

'============================================================================
' Current selection state (centralized)
'============================================================================
Public CurrentTroopIndex As Long
Public CurrentItemIndex As Long
Public CurrentFactionIndex As Long
Public CurrentPartyIndex As Long
Public CurrentSceneIndex As Long
Public CurrentMapIconIndex As Long
Public CurrentSoundIndex As Long
Public CurrentPSysIndex As Long
Public CurrentTabMatIndex As Long
Public CurrentMeshIndex As Long
Public CurrentTimeTrgIndex As Long
Public CurrentStrIndex As Long

'============================================================================
' MnBInfo access
'============================================================================
Public Property Get ModPath() As String
    ModPath = MnBInfo.ModPath
End Property

Public Property Get ModName() As String
    ModName = MnBInfo.ModName
End Property

Public Property Get Language() As String
    Language = MnBInfo.Language
End Property

Public Property Get IsFirstTimeEdit() As Boolean
    IsFirstTimeEdit = MnBInfo.FirstTimeEdit
End Property

Public Property Get EditInfo(ByVal EditType As Long) As Long
    EditInfo = MnBInfo.EditInfo(EditType)
End Property

Public Property Let EditInfo(ByVal EditType As Long, ByVal Value As Long)
    MnBInfo.EditInfo(EditType) = Value
End Property

'============================================================================
' Bulk load notification
'============================================================================
Public Sub NotifyDataLoaded()
    RaiseEvent DataLoaded
End Sub

Public Sub NotifyDataSaved()
    RaiseEvent DataSaved
End Sub
