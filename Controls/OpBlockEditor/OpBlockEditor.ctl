VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OpBlockEditor 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   ScaleHeight     =   4710
   ScaleWidth      =   7050
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6375
      Begin MnBWarband_Editor.ComboforOp OpCombo 
         Left            =   2040
         Top             =   2760
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin MnBWarband_Editor.ListViewforMS LV 
         Height          =   2775
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4895
      End
      Begin MSComctlLib.ImageList IL 
         Left            =   3240
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.Label cCMD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "创建(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   180
         Index           =   1
         Left            =   4680
         MouseIcon       =   "OpBlockEditor.ctx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label cCMD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "删除(&E)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   0
         Left            =   5520
         MouseIcon       =   "OpBlockEditor.ctx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label cCMD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上移(&U)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   3000
         MouseIcon       =   "OpBlockEditor.ctx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   3120
         Width           =   705
      End
      Begin VB.Label cCMD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下移(&D)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   3840
         MouseIcon       =   "OpBlockEditor.ctx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   3120
         Width           =   705
      End
   End
   Begin MnBWarband_Editor.RichforPY Rich1 
      Height          =   2775
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   4895
   End
   Begin MnBWarband_Editor.RichforTXT Rich2 
      Height          =   2775
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4895
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4895
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "伪代码"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PY代码"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "TXT代码"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "OpBlockEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event Change(TabIndex As Integer)
Public Event TabClick(Previous As Integer, TabIndex As Integer)
Private LastTabIndex As Integer
Public Pseudo As ListViewforMS
Public RichPY As RichforPY
Public RichTXT As RichforTXT
Dim clsOpbEd As clsOpBlock
Dim p_NoChange As Boolean

Private Sub cCMD_Click(Index As Integer)
Dim MosP As POINTAPI
Select Case Index
  Case 0
    If LV.ListIndex > 0 And LV.ListIndex <= LV.ListItems.Count Then
      If LV.SubIndex = 0 Then
        LV.ListItems.Remove LV.ListIndex
        If LV.ListIndex > LV.ListItems.Count Then LV.ListIndex = LV.ListItems.Count
      Else
        
        If LV.ListItems(LV.ListIndex).Value = Call_Script And LV.ListItems(LV.ListIndex).SubItems(LV.SubIndex).Value = "" Then
          '...
          LV.ListItems(LV.ListIndex).SubItems.Remove LV.SubIndex
          If LV.SubIndex > LV.ListItems(LV.ListIndex).SubItems.Count Then
            LV.SubIndex = LV.ListItems(LV.ListIndex).SubItems.Count
          End If
        Else
          LV.ListItems(LV.ListIndex).SubItems(LV.SubIndex).AssignValue ""
        End If
      End If
      clsOpbEd.CalcIndention
      LV.Refresh
      
      RaiseEvent Change(TabStrip1.SelectedItem.Index)
    End If
  Case 1
    GetCursorPos MosP
    OpCombo.ShowMenu MosP.X, MosP.Y
  Case 2
    If LV.ListIndex < LV.ListItems.Count Then
      LV.ListItems.Swap LV.ListIndex, LV.ListIndex + 1
      LV.ListIndex = LV.ListIndex + 1
      clsOpbEd.CalcIndention
      LV.Refresh
      
      RaiseEvent Change(TabStrip1.SelectedItem.Index)
    End If
  Case 3
    If LV.ListIndex > 1 Then
      LV.ListItems.Swap LV.ListIndex, LV.ListIndex - 1
      LV.ListIndex = LV.ListIndex - 1
      clsOpbEd.CalcIndention
      LV.Refresh
      
      RaiseEvent Change(TabStrip1.SelectedItem.Index)
    End If
End Select
End Sub

Private Sub LV_Change()
If Not p_NoChange Then
  RaiseEvent Change(TabStrip1.SelectedItem.Index)
End If
End Sub

Private Sub OpCombo_ItemSelect(strValue As String, strText As String)
Dim i As Integer, Op(0) As String, NextIdx As Integer, oItem As ListItemforMS
Dim OpID As Long, neg As Integer

Op(0) = strValue
NextIdx = IIf(LV.ListIndex = 0, 0, LV.ListIndex + 1)
Set oItem = LV.ListItems.AddOpBlock(NextIdx, Op())
LV.ListIndex = oItem.Index
LV.SubIndex = 0

clsOpbEd.CalcIndention
oItem.Refresh
LV.Refresh True

RaiseEvent Change(TabStrip1.SelectedItem.Index)
End Sub

Private Sub Rich1_Change()
If Not p_NoChange Then
  RaiseEvent Change(TabStrip1.SelectedItem.Index)
End If
End Sub

Private Sub Rich2_Change()
If Not p_NoChange Then
  RaiseEvent Change(TabStrip1.SelectedItem.Index)
End If
End Sub


Private Sub TabStrip1_Click()
Select Case TabStrip1.SelectedItem.Index
  Case 1
    Frame1.Visible = True
    Rich1.Visible = False
    Rich2.Visible = False
    Frame1.ZOrder
  Case 2
    Frame1.Visible = False
    Rich1.Visible = True
    Rich2.Visible = False
    Rich1.ZOrder
  Case 3
    Frame1.Visible = False
    Rich1.Visible = False
    Rich2.Visible = True
    Rich2.ZOrder
End Select

clsOpbEd.OpBEd_TabClick LastTabIndex, TabStrip1.SelectedItem.Index
'RaiseEvent TabClick(LastTabIndex, TabStrip1.SelectedItem.Index)

LastTabIndex = TabStrip1.SelectedItem.Index
End Sub

Private Sub UserControl_Initialize()
LastTabIndex = 1
End Sub

Public Sub Initialize(OpbEd As clsOpBlock)
LV.Initialize LV
Set Pseudo = LV
LV.Spacing = 1.5
Set RichPY = Rich1
Set RichTXT = Rich2

Set clsOpbEd = OpbEd
Rich1.Initialize
Rich2.OperationColor = vbBlue
Rich2.ParamColor = &H4000&
Rich2.CountColor = vbRed

OpCombo.Initialize
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Dim i As Long

TabStrip1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
Frame1.Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
Rich1.Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight
Rich2.Move TabStrip1.ClientLeft, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight

For i = 0 To 3
  cCMD(i).Top = Frame1.Height - cCMD(i).Height * 1.5
Next i

cCMD(0).Left = Frame1.Width - cCMD(0).Width - cCMD(0).Height * 0.5

For i = 1 To 3
  cCMD(i).Left = cCMD(i - 1).Left - cCMD(0).Width - cCMD(i).Height * 0.5
Next i

LV.Move 0, 0, Frame1.Width, Frame1.Height - cCMD(0).Height * 2
End Sub


Public Property Get TextPY() As String
TextPY = Rich1.Text
End Property

Public Property Let TextPY(ByVal vNewValue As String)
Rich1.Text = vNewValue

PropertyChanged TextPY
End Property

Public Property Get TextTXT() As String
TextTXT = Rich2.Text
End Property

Public Property Let TextTXT(ByVal vNewValue As String)
Rich2.DrawText vNewValue

PropertyChanged TextTXT
End Property


Public Property Get SelStart_PY() As Long
SelStart_PY = Rich1.SelStart
End Property

Public Property Let SelStart_PY(ByVal vNewValue As Long)
Rich1.SelStart = vNewValue

PropertyChanged (SelStart_PY)
End Property


Public Property Get SelLength_PY() As Long
SelLength_PY = Rich1.SelLength
End Property

Public Property Let SelLength_PY(ByVal vNewValue As Long)
Rich1.SelLength = vNewValue

PropertyChanged (SelLength_PY)
End Property

Public Property Get SelColor_PY() As Long
SelColor_PY = Rich1.SelColor
End Property

Public Property Let SelColor_PY(ByVal vNewValue As Long)
Rich1.SelColor = vNewValue

PropertyChanged (SelColor_PY)
End Property


Public Property Get SelStart_TXT() As Long
SelStart_TXT = Rich2.SelStart
End Property

Public Property Let SelStart_TXT(ByVal vNewValue As Long)
Rich2.SelStart = vNewValue

PropertyChanged (SelStart_TXT)
End Property

Public Property Get NoChangeEvent() As Boolean
NoChangeEvent = p_NoChange
End Property

Public Property Let NoChangeEvent(ByVal vNewValue As Boolean)
p_NoChange = vNewValue

PropertyChanged (NoChangeEvent)
End Property

Public Property Get SelLength_TXT() As Long
SelLength_TXT = Rich2.SelLength
End Property

Public Property Let SelLength_TXT(ByVal vNewValue As Long)
Rich2.SelLength = vNewValue

PropertyChanged (SelLength_TXT)
End Property

Public Property Get SelColor_TXT() As Long
SelColor_TXT = Rich2.SelColor
End Property

Public Property Let SelColor_TXT(ByVal vNewValue As Long)
Rich2.SelColor = vNewValue

PropertyChanged (SelColor_TXT)
End Property


Public Property Get TabStripIndex() As Long
TabStripIndex = TabStrip1.SelectedItem.Index
End Property

Public Property Let TabStripIndex(ByVal vNewValue As Long)
TabStrip1.Tabs(vNewValue).Selected = True

PropertyChanged (TabStripIndex)
End Property

Public Property Get Visible() As Boolean
Visible = Rich1.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
Rich1.Visible = vNewValue
Rich2.Visible = vNewValue
Frame1.Visible = vNewValue
TabStrip1.Visible = vNewValue

PropertyChanged Visible
End Property


