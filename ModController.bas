Attribute VB_Name = "ModController"
Option Explicit

'============================================================================
' ModController.bas - Proof of Concept: Layered Architecture Controller
'
' Demonstrates the MVC-like pattern for VB6 forms using DataRepository.
' This module acts as the "Controller" layer between the DataRepository
' (Model) and forms (View).
'
' PATTERN: All form event handlers should delegate to controller methods
' rather than directly manipulating global arrays.
'
' BEFORE (current code in frmFactions):
'   Private Sub CApply_Click()
'       Factions(CurrentFactionID).strName = txtName.Text   ' direct global write
'       Factions(CurrentFactionID).Flags = txtFlags.Text    ' direct global write
'       SaveFactionFile MnBInfo.ModPath & "\factions.txt"   ' direct save call
'   End Sub
'
' AFTER (using controller):
'   Private Sub CApply_Click()
'       FactionController.ApplyCurrentFaction txtName.Text, txtFlags.Text
'       ' Controller handles validation, data update, event firing, and save
'   End Sub
'============================================================================

'============================================================================
' Faction Controller - Example of per-entity controller pattern
'============================================================================

' Load factions from file and populate UI
Public Function LoadFactions(ByRef oListView As ListView, ByRef oRepo As DataRepository) As Boolean
    On Error GoTo errorHandle

    Dim i As Long
    Dim oItem As ListItem

    ' Load data via FileLoader (preserves existing flow)
    LoadFactionFile oRepo.ModPath & "\factions.txt"
    LoadFactionCSVFile oRepo.ModPath & "\languages\" & oRepo.Language & "\factions.csv"

    ' Populate ListView from DataRepository (not global arrays)
    oListView.ListItems.Clear
    For i = 0 To oRepo.FactionCount - 1
        Set oItem = oListView.ListItems.Add(, , CStr(i + 1))
        oItem.SubItems(1) = oRepo.Faction(i).strID
        oItem.SubItems(2) = oRepo.Faction(i).csvName
    Next

    oRepo.NotifyDataLoaded
    LoadFactions = True

    Exit Function

errorHandle:
    Call logErr("ModController", "LoadFactions", Err.Number, Err.Description)
End Function

' Apply edits from UI to data store (replaces direct global array writes)
Public Function ApplyFactionEdit(ByRef oRepo As DataRepository, _
                                  ByVal Index As Long, _
                                  ByVal strName As String, _
                                  ByVal strFlags As String, _
                                  ByVal strColor As String) As Boolean
    On Error GoTo errorHandle

    Dim fac As Type_Faction

    ' Validate inputs before applying
    If Index < 0 Or Index >= oRepo.FactionCount Then
        MsgBox "Invalid faction index: " & CStr(Index)
        Exit Function
    End If

    If Len(Trim$(strName)) = 0 Then
        MsgBox "Faction name cannot be empty"
        Exit Function
    End If

    ' Read current state via DataRepository
    fac = oRepo.Faction(Index)

    ' Apply changes to local copy
    fac.strName = strName
    fac.Flags = Val(strFlags)
    fac.lColor = strColor
    fac.Edit = True

    ' Write back via DataRepository (fires FactionChanged event)
    oRepo.Faction(Index) = fac

    ApplyFactionEdit = True

    Exit Function

errorHandle:
    Call logErr("ModController", "ApplyFactionEdit", Err.Number, Err.Description)
End Function

' Save all faction data to file (replaces direct SaveFactionFile call)
Public Function SaveFactions(ByRef oRepo As DataRepository) As Boolean
    On Error GoTo errorHandle

    SaveFactionFile oRepo.ModPath & "\factions.txt"
    SaveFactionCSVFile oRepo.ModPath & "\languages\" & oRepo.Language & "\factions.csv"

    oRepo.NotifyDataSaved
    SaveFactions = True

    Exit Function

errorHandle:
    Call logErr("ModController", "SaveFactions", Err.Number, Err.Description)
End Function

'============================================================================
' Generic controller pattern - can be adapted for any entity type
'
' For each entity (Troop, Item, Party, Scene, etc.), create equivalent:
'
'   Load<Entity>s(oListView, oRepo)       ' Load from file + populate UI
'   Apply<Entity>Edit(oRepo, idx, ...)    ' Validate + apply changes to repo
'   Save<Entity>s(oRepo)                  ' Persist to file
'   Refresh<Entity>List(oListView, oRepo) ' Re-populate ListView from repo
'
' This pattern:
'   1. Centralizes validation logic
'   2. Makes it easy to add undo/redo (just snapshot the repo)
'   3. Enables unit testing (controller methods take explicit params)
'   4. Simplifies migration to .NET (same API surface, different impl)
'============================================================================

' Generic ListView refresh from DataRepository
Public Sub RefreshListView(ByRef oListView As ListView, ByRef oRepo As DataRepository, ByVal EntityType As Long)
    On Error GoTo errorHandle

    Dim i As Long
    Dim oItem As ListItem

    oListView.ListItems.Clear

    Select Case EntityType
        Case Tag_Troop
            For i = 0 To oRepo.TroopCount - 1
                Set oItem = oListView.ListItems.Add(, , CStr(i + 1))
                oItem.SubItems(1) = oRepo.Troop(i).strID
                oItem.SubItems(2) = oRepo.Troop(i).csvName
            Next

        Case Tag_Item
            For i = 0 To oRepo.ItemCount - 1
                Set oItem = oListView.ListItems.Add(, , CStr(i + 1))
                oItem.SubItems(1) = oRepo.Item(i).dbName
                oItem.SubItems(2) = oRepo.Item(i).csvName
            Next

        Case Tag_Faction
            For i = 0 To oRepo.FactionCount - 1
                Set oItem = oListView.ListItems.Add(, , CStr(i + 1))
                oItem.SubItems(1) = oRepo.Faction(i).strID
                oItem.SubItems(2) = oRepo.Faction(i).csvName
            Next

        Case Tag_Party
            For i = 0 To oRepo.PartyCount - 1
                Set oItem = oListView.ListItems.Add(, , CStr(i + 1))
                oItem.SubItems(1) = oRepo.Party(i).strID
                oItem.SubItems(2) = oRepo.Party(i).csvName
            Next

        Case Tag_Scene
            For i = 0 To oRepo.SceneCount - 1
                Set oItem = oListView.ListItems.Add(, , CStr(i + 1))
                oItem.SubItems(1) = oRepo.Scene(i).strID
                oItem.SubItems(2) = oRepo.Scene(i).strName
            Next
    End Select

    Exit Sub

errorHandle:
    Call logErr("ModController", "RefreshListView", Err.Number, Err.Description)
End Sub

'============================================================================
' HOW TO MIGRATE AN EXISTING FORM TO USE THIS PATTERN:
'
' 1. Add a module-level DataRepository instance:
'      Private WithEvents mRepo As DataRepository
'
' 2. In Form_Load, initialize the repository:
'      Set mRepo = New DataRepository
'      FactionController.LoadFactions ListView1, mRepo
'
' 3. Replace direct global array access:
'      OLD: Factions(CurrentFactionID).strName = txtName.Text
'      NEW: mRepo.Faction(mRepo.CurrentFactionIndex).strName = txtName.Text
'      BETTER: FactionController.ApplyFactionEdit mRepo, idx, name, flags, color
'
' 4. Handle DataRepository events for cross-form updates:
'      Private Sub mRepo_FactionChanged(ByVal Index As Long)
'          ' Refresh UI for changed faction
'      End Sub
'
' 5. Replace direct save calls:
'      OLD: SaveFactionFile MnBInfo.ModPath & "\factions.txt"
'      NEW: FactionController.SaveFactions mRepo
'============================================================================
