VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TriggersEditor 
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6510
   ScaleHeight     =   4485
   ScaleWidth      =   6510
   Begin VB.ComboBox cbTiOn 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   5055
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   140
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "触发条件："
      Top             =   520
      Width           =   975
   End
   Begin VB.PictureBox PicDrag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      ScaleHeight     =   465
      ScaleWidth      =   1425
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
   End
   Begin MnBWarband_Editor.OpBlockEditor OpBlockEditor1 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6135
      _extentx        =   10821
      _extenty        =   5741
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mMenu 
      Caption         =   "触发器菜单"
      Visible         =   0   'False
      Begin VB.Menu mTrigger 
         Caption         =   "新增触发器(&A)"
         Index           =   0
      End
      Begin VB.Menu mTrigger 
         Caption         =   "复制触发器(&C)"
         Index           =   1
      End
      Begin VB.Menu mTrigger 
         Caption         =   "粘贴触发器(&P)"
         Index           =   2
      End
      Begin VB.Menu mTrigger 
         Caption         =   "删除触发器(&D)"
         Index           =   3
      End
   End
End
Attribute VB_Name = "TriggersEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Strip As TabStrip
Public CombotiOn As ComboBox
Public OpEditor As New clsOpBlock
Public TabIndex As Integer, PopupIndex As Integer
Private clsEditor As clsTriggersEditor
Dim DragIndex As Integer, LastX As Single, LastY As Single
Public Changed As Boolean
Private IdxOnly As Boolean
Public Vacant As Boolean
Public CustomActive As Boolean

Public Event TabClick(PreIndex As Integer, Index As Integer, IndexOnly As Boolean)
Public Event MenuClick(Index As Integer, MenuOrder As Integer)

Public Sub Initialize(clsEd As clsTriggersEditor)
  Set Strip = TabStrip1
  Set CombotiOn = cbTiOn
  OpEditor.Attach OpBlockEditor1
  
  Set clsEditor = clsEd
  Changed = False
  CustomActive = True
End Sub


Private Sub cbTiOn_Click()
If CustomActive Then
  clsEditor.SettiOn cbTiOn.ListIndex
End If
End Sub

Private Sub cbTiOn_Scroll()
Call cbTiOn_Click
End Sub

Private Sub mTrigger_Click(Index As Integer)
RaiseEvent MenuClick(PopupIndex, Index)
End Sub

Private Sub OpBlockEditor1_Change(TabIndex As Integer)
Changed = True
End Sub

Private Sub TabStrip1_Click()
  If Vacant Then
    CheckListTrgIdx = 0
    TabIndex = 0
  Else
    CheckListTrgIdx = TabStrip1.SelectedItem.Index
    RaiseEvent TabClick(TabIndex, TabStrip1.SelectedItem.Index, IdxOnly)
    TabIndex = TabStrip1.SelectedItem.Index
  End If

  'clsEditor.ctlEditor_TabClick TabIndex, TabStrip1.SelectedItem.Index
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
  With TabStrip1
    .Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
     
     txtTitle.Move .ClientLeft + 100, .ClientTop + 40
    cbTiOn.Move txtTitle.Left + txtTitle.Width, .ClientTop, .ClientWidth - txtTitle.Width - txtTitle.Left
    OpBlockEditor1.Move .ClientLeft, cbTiOn.Top + cbTiOn.Height + 10, .ClientWidth, .ClientHeight - cbTiOn.Top - 10
  End With
End Sub


Private Sub TabStrip1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If Button = vbLeftButton And Not Vacant Then
  For i = 1 To TabStrip1.Tabs.Count
    With TabStrip1.Tabs(i)
      If X + TabStrip1.Left >= .Left And X + TabStrip1.Left <= .Left + .Width And Y + TabStrip1.Top >= .Top And Y + TabStrip1.Top <= .Top + .Height Then
        DragIndex = i
        LastX = X
        LastY = Y
        RaiseEvent TabClick(TabIndex, i, False)
        TabIndex = i
        Exit For
      End If
    End With
  Next i

End If
End Sub

Private Sub TabStrip1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If Button = vbLeftButton And DragIndex > 0 Then
  With PicDrag
    If .Visible = False Then
      .Height = TabStrip1.Tabs(DragIndex).Height
      .Width = TabStrip1.Tabs(DragIndex).Width
      .Top = TabStrip1.Tabs(DragIndex).Top
      .Left = TabStrip1.Tabs(DragIndex).Left
      .Visible = True
      Label1.Move 0, 0, .ScaleWidth, .ScaleHeight
      Label1.Caption = TabStrip1.Tabs(DragIndex).Caption
    Else
      .Left = TabStrip1.Tabs(DragIndex).Left + X - LastX
      .Top = TabStrip1.Tabs(DragIndex).Top + Y - LastY
    End If
  End With
End If

End Sub

Private Sub TabStrip1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer, l As Single, ToIndex As Integer, Start As Integer, EndI As Integer, StepI As Integer, Center As Single, q As Boolean
If Button = vbLeftButton Then
  If DragIndex > 0 Then
    
    With PicDrag
      If .Visible Then
        If .Top >= TabStrip1.Top And .Top <= TabStrip1.Tabs(1).Top + TabStrip1.Tabs(1).Height Then
          Center = .Left + .Width / 2
          For i = 1 To TabStrip1.Tabs.Count
            l = TabStrip1.Tabs(i).Left + TabStrip1.Tabs(i).Width / 2
          
            If Center <= l Then
              ToIndex = i
              Exit For
            End If
          
            If i = TabStrip1.Tabs.Count And Center > l Then
               ToIndex = i + 1
            End If

          Next i
        
          If DragIndex > ToIndex Then
            Start = DragIndex
            EndI = ToIndex + 1
            StepI = -1
          Else
            Start = DragIndex
            EndI = ToIndex - 2
            StepI = 1
          End If
        
          For i = Start To EndI Step StepI
            'ExchangeTab i, i + StepI
            clsEditor.ExchangeTrigger i, i + StepI
            If Not q Then q = True
          Next i
          
          If q Then
            IdxOnly = True
            TabStrip1.SelectedItem = TabStrip1.Tabs(i)
            IdxOnly = False
          End If
          'CheckListTrgIdx = i
          'RaiseEvent TabClick(TabIndex, TabStrip1.SelectedItem.Index, True)
          'TabIndex = i
        End If
        .Visible = False
      End If
      DragIndex = 0
      
    End With
    
  End If
  
ElseIf Button = vbRightButton Then
  TabStrip1.SetFocus
  If Not Vacant Then
    For i = 0 To 3
      mTrigger(i).Enabled = True
    Next i
    mTrigger(2).Enabled = TriggerCopied
    For i = 1 To TabStrip1.Tabs.Count
      With TabStrip1.Tabs(i)
        If X + TabStrip1.Left >= .Left And X + TabStrip1.Left <= .Left + .Width And Y + TabStrip1.Top >= .Top And Y + TabStrip1.Top <= .Top + .Height Then
          PopupIndex = i
          PopupMenu mMenu
          Exit For
        End If
      End With
    Next i
  Else
    mTrigger(0).Enabled = True
    mTrigger(2).Enabled = TriggerCopied
    mTrigger(1).Enabled = False
    mTrigger(3).Enabled = False
    PopupIndex = 0
    PopupMenu mMenu
  End If
End If
End Sub

Private Sub ExchangeTab(Tab1 As Integer, Tab2 As Integer)
Dim t As String
t = TabStrip1.Tabs(Tab1)
TabStrip1.Tabs(Tab1) = TabStrip1.Tabs(Tab2)
TabStrip1.Tabs(Tab2) = t
End Sub

Public Sub LoadtiOn(EditorType As String)
Dim i As Integer

cbTiOn.Clear
For i = 0 To UBound(tiOns)
  If tiOns(i).Occation = EditorType Or tiOns(i).Occation = "gnl" Then
    cbTiOn.AddItem tiOns(i).csvName & tiOns(i).Tip
    cbTiOn.ItemData(cbTiOn.ListCount - 1) = i
  End If
Next i
End Sub

Public Sub CleartiOn()
cbTiOn.Clear
End Sub

Public Sub SetComboVisible(Visible As Boolean)
cbTiOn.Visible = Visible
txtTitle.Visible = Visible
End Sub

Public Sub SettiOn(tiOn As Double)
Dim i As Integer

For i = 0 To UBound(tiOns)
  If tiOns(i).Value = tiOn Then
    cbTiOn.ListIndex = i
    Exit For
  End If
Next i
End Sub

Public Function GettiOn() As Double
If cbTiOn.ListIndex > -1 Then
   GettiOn = tiOns(cbTiOn.ListIndex).Value
End If
End Function
