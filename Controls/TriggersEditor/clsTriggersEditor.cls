VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTriggersEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Triggers() As Type_Trigger
Private Lct As String
Private strType As String
Public WithEvents ctlEditor As TriggersEditor
Attribute ctlEditor.VB_VarHelpID = -1
Private UpdateOnce As Boolean
Private AE As Boolean

Private Sub Class_Initialize()
ReDim Triggers(0)
UpdateOnce = False
End Sub

Public Sub Initialize(Editor As TriggersEditor)
Set ctlEditor = Editor

ctlEditor.Initialize Me
End Sub

Friend Sub AssignTrigger(iTriggers() As Type_Trigger, Optional AutoEncode As Boolean = True)
Dim i As Integer, TabIndex As Integer

Triggers = iTriggers

ctlEditor.Strip.Tabs.Clear
For i = 1 To UBound(Triggers)
  ctlEditor.Strip.Tabs.Add , , PublicEditors_Simplified(13) & i
Next i

'UpdateCurVarNameCheckList Lct, UBound(Triggers)

If ctlEditor.Strip.Tabs.Count >= 1 Then
  ctlEditor.TabIndex = 1
  ctlEditor.OpEditor.Visible = True
  ctlEditor.SetComboVisible True
  ctlEditor.Vacant = False
Else
  ctlEditor.TabIndex = 0
  ctlEditor.OpEditor.Visible = False
  ctlEditor.Vacant = True
  ctlEditor.SetComboVisible False
  ctlEditor.Strip.Tabs.Add , , PublicTips(0)
End If
CheckListTrgIdx = ctlEditor.TabIndex

AE = AutoEncode
If AutoEncode Then
  'If Not ctlEditor.Strip.SelectedItem Is Nothing Then
  If CheckListTrgIdx > 0 Then
    ctlEditor.CustomActive = False
    TabIndex = ctlEditor.Strip.SelectedItem.Index
    ctlEditor.SettiOn Triggers(TabIndex).tiOn
    ctlEditor.OpEditor.AssignOpBlock Triggers(TabIndex).tiAct(), AutoEncode, , , False
    ctlEditor.CustomActive = True
  End If
End If
End Sub

Friend Sub GetTrigger(oTriggers() As Type_Trigger)
oTriggers = Triggers
End Sub

Public Property Get Location() As String
Location = Lct
End Property

Public Property Let Location(ByVal vNewValue As String)  '先赋值Location,再给Triggers赋值
Dim i As Long
Lct = vNewValue
i = GetVarNameCheckListNo(Lct)

If i > 0 Then
  CurVarNameList = VarNameLists(i)
Else
  UpdateCurVarNameCheckList Lct, 0, True
End If

TemGVarNameList = VarNameLists(0)

End Property

Public Property Get EditorType() As String
EditorType = strType
End Property

Public Property Let EditorType(ByVal vNewValue As String)
ctlEditor.LoadtiOn vNewValue

strType = vNewValue

End Property

Public Sub ctlEditor_MenuClick(Index As Integer, MenuOrder As Integer)
Dim TemBlocks() As Type_Op_Block, i As Long
ReDim TemBlocks(0)
Select Case MenuOrder
  Case 0
    If Not ctlEditor.Vacant Then
      ReDim Preserve Triggers(1 To UBound(Triggers) + 1)
    Else
      ReDim Triggers(1 To 1)
    End If
    'For i = UBound(Triggers) To Index + 2 Step -1
    '  Triggers(i) = Triggers(i - 1)
    'Next i
    
    Triggers(UBound(Triggers)).ActNum = 0
    Triggers(UBound(Triggers)).tiOn = 0
    ReDim Triggers(UBound(Triggers)).tiAct(0)
    
    If ctlEditor.Vacant Then
      ctlEditor.Vacant = False
      ctlEditor.Strip.Tabs.Remove 1
      ctlEditor.OpEditor.Visible = True
    End If
    ctlEditor.Strip.Tabs.Add , , PublicEditors_Simplified(13) & UBound(Triggers)
    ctlEditor.SetComboVisible True
  Case 1
    If Index > 0 Then TriggerBoard = Triggers(Index)
    If Not TriggerCopied Then TriggerCopied = True
  Case 2
    If TriggerCopied Then
    If Not ctlEditor.Vacant Then
      ReDim Preserve Triggers(1 To UBound(Triggers) + 1)
    Else
      ReDim Triggers(1 To 1)
    End If
    'For i = UBound(Triggers) To Index + 2 Step -1
    '  Triggers(i) = Triggers(i - 1)
    'Next i
    
    Triggers(UBound(Triggers)) = TriggerBoard
    If ctlEditor.Vacant Then
      ctlEditor.Vacant = False
      ctlEditor.Strip.Tabs.Remove 1
      ctlEditor.OpEditor.Visible = True
      ctlEditor.SetComboVisible True
    End If
    
    ctlEditor.Strip.Tabs.Add , , PublicEditors_Simplified(13) & UBound(Triggers)
    
    End If
  Case 3
    If Index > 0 Then
      For i = Index To UBound(Triggers) - 1
        Triggers(i) = Triggers(i + 1)
      Next i
      
      UpdateOnce = True
      If UBound(Triggers) > 1 Then
        ReDim Preserve Triggers(1 To UBound(Triggers) - 1)
      Else
        ReDim Triggers(0)
      End If
    
      ctlEditor.Strip.Tabs.Remove UBound(Triggers) + 1
      
      If ctlEditor.Strip.Tabs.Count = 0 Then
        ctlEditor.Vacant = True
        ctlEditor.TabIndex = 0
        ctlEditor.OpEditor.AssignOpBlock TemBlocks, AE, , , False
        ctlEditor.Strip.Tabs.Add , , PublicTips(0)
        ctlEditor.OpEditor.Visible = False
        ctlEditor.SetComboVisible False
      End If
    End If
End Select
End Sub

Public Sub ctlEditor_TabClick(PreIndex As Integer, Index As Integer, IndexOnly As Boolean)
If (PreIndex <> Index And Not IndexOnly) Or UpdateOnce Then
  If PreIndex <= ctlEditor.Strip.Tabs.Count And PreIndex > 0 And Not UpdateOnce Then
    ctlEditor.OpEditor.GetOpBlock Triggers(PreIndex).tiAct
    Triggers(PreIndex).tiOn = ctlEditor.GettiOn
  End If
  CheckListTrgIdx = Index
  ctlEditor.OpEditor.AssignOpBlock Triggers(Index).tiAct, AE, , , False
  ctlEditor.Changed = False
  ctlEditor.CustomActive = False
  ctlEditor.SettiOn Triggers(Index).tiOn
  ctlEditor.CustomActive = True

  If UpdateOnce Then UpdateOnce = False
End If

End Sub


Public Sub ExchangeTrigger(Trg1 As Integer, Trg2 As Integer)
Dim t As Type_Trigger
t = Triggers(Trg1)
Triggers(Trg1) = Triggers(Trg2)
Triggers(Trg2) = t
End Sub

Friend Sub InputTrg(strLocation As String, strEditorType As String, iTriggers() As Type_Trigger)
Location = strLocation
EditorType = strEditorType
AssignTrigger iTriggers, True
End Sub

Friend Sub OutputTrg(oTriggers() As Type_Trigger)
Dim i As Long, j As Long
If ctlEditor.Changed Then
  ctlEditor.OpEditor.GetOpBlock Triggers(ctlEditor.TabIndex).tiAct
  If LBound(Triggers(ctlEditor.TabIndex).tiAct) = 0 Then
     Triggers(ctlEditor.TabIndex).ActNum = 0
     'ReDim Triggers(ctlEditor.TabIndex).tiAct(0)
  Else
     Triggers(ctlEditor.TabIndex).ActNum = UBound(Triggers(ctlEditor.TabIndex).tiAct)
  End If
End If

GetTrigger oTriggers
i = GetVarNameCheckListNo(Lct)

If i > 0 Then
  VarNameLists(i) = CurVarNameList
Else
  j = CreateVarNameCheckList(Lct)
  VarNameLists(j) = CurVarNameList
  LocateVarNameCheckList Lct, j
End If

VarNameLists(0) = TemGVarNameList
End Sub

Public Sub SettiOn(Index As Integer)
On Error GoTo EL

If Index > -1 Then
  Triggers(ctlEditor.TabIndex).tiOn = tiOns(Index).Value
End If

Exit Sub

EL:
Call logErr("clsTriggersEditor", "SettiOn", Err.Number, Err.Description)
End Sub
