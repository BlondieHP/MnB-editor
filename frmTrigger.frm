VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrigger 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "触发器编辑器"
   ClientHeight    =   9375
   ClientLeft      =   3105
   ClientTop       =   960
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_13"
   Begin MSComctlLib.ImageList IL1 
      Left            =   960
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrigger.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrigger.frx":039A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrigger.frx":0734
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前触发器(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8760
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton CApply 
      BackColor       =   &H000000FF&
      Caption         =   "套用(&A)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton CReset 
      BackColor       =   &H000000FF&
      Caption         =   "重置(&R)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   6
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton CQuery 
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   3600
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CQuery 
      Caption         =   "↑"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   720
      TabIndex        =   3
      Top             =   100
      Width           =   2775
   End
   Begin MSComctlLib.ListView LstTrigger 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   14420
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmTrigger.frx":0ACE
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps2"
      Height          =   7935
      Index           =   2
      Left            =   4680
      TabIndex        =   25
      Top             =   480
      Width           =   9375
      Begin MnBWarband_Editor.OpBlockEditor OpBlockEditor2 
         Height          =   6975
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   12303
      End
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   7200
         Width           =   9135
         Begin VB.OptionButton OptHex2 
            Caption         =   "十进制"
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptHex2 
            Caption         =   "16进制"
            Height          =   255
            Index           =   1
            Left            =   6840
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptHex2 
            Caption         =   "二进制"
            Height          =   255
            Index           =   2
            Left            =   7920
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtCalc2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   31
            Text            =   "txtCalc"
            Top             =   180
            Width           =   5415
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
      Height          =   7935
      Index           =   0
      Left            =   4680
      TabIndex        =   9
      Top             =   480
      Width           =   9375
      Begin VB.Frame FraFunction 
         Caption         =   "功能识别"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   7575
         Left            =   4200
         TabIndex        =   26
         Top             =   240
         Width           =   4935
         Begin VB.TextBox txtFuncOp 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   5895
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   29
            Text            =   "frmTrigger.frx":513B
            Top             =   1440
            Width           =   4455
         End
         Begin VB.TextBox txtFuncDes 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   255
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "txtFuncDes"
            Top             =   960
            Width           =   4455
         End
         Begin VB.TextBox txtFuncName 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   255
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "txtFuncName"
            Top             =   480
            Width           =   4455
         End
      End
      Begin VB.Frame FraTime 
         Caption         =   "时间"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3855
         Begin VB.TextBox TxtRearmItv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            TabIndex        =   24
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox TxtDelayItv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            TabIndex        =   21
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox TxtChkItv 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            TabIndex        =   19
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "重启时间:"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   23
            Top             =   1730
            Width           =   1020
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "延迟时间:"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   22
            Top             =   1130
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "检测时间:"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   20
            Top             =   510
            Width           =   1020
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps1"
      Height          =   7935
      Index           =   1
      Left            =   4680
      TabIndex        =   10
      Top             =   480
      Width           =   9375
      Begin MnBWarband_Editor.OpBlockEditor OpBlockEditor1 
         Height          =   7095
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   12515
      End
      Begin VB.Frame FraCalc 
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   7200
         Width           =   9135
         Begin VB.TextBox txtCalc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   40
            Text            =   "txtCalc"
            Top             =   180
            Width           =   5535
         End
         Begin VB.OptionButton OptHex 
            Caption         =   "二进制"
            Height          =   255
            Index           =   2
            Left            =   8040
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptHex 
            Caption         =   "16进制"
            Height          =   255
            Index           =   1
            Left            =   6960
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptHex 
            Caption         =   "十进制"
            Height          =   255
            Index           =   0
            Left            =   5880
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Null"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   15
         Index           =   8
         Left            =   2880
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
      End
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   8415
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   14843
      MultiRow        =   -1  'True
      ImageList       =   "IL1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "激发时间及功能识别(&T)"
            Key             =   "Time"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "激发条件(&C)"
            Key             =   "Condition"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "激发结果(&S)"
            Key             =   "Consequence"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label LbTest 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   720
      TabIndex        =   11
      Top             =   9120
      Visible         =   0   'False
      Width           =   570
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
      Index           =   3
      Left            =   3600
      MouseIcon       =   "frmTrigger.frx":5145
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   8760
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
      Index           =   2
      Left            =   2640
      MouseIcon       =   "frmTrigger.frx":544F
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   8760
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "删除(&D)"
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
      Left            =   3600
      MouseIcon       =   "frmTrigger.frx":5759
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   9000
      Width           =   705
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
      Left            =   2640
      MouseIcon       =   "frmTrigger.frx":5A63
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   9000
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "触发器数:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   8760
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   200
      TabIndex        =   2
      Top             =   165
      Width           =   495
   End
End
Attribute VB_Name = "frmTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CustomActive As Boolean
Dim Loading As Boolean
Dim CalcText_De As String
Dim CalcText_De2 As String
Dim OpCnt1 As New clsOpBlock
Dim OpCnt2 As New clsOpBlock

Private Sub InitTriggerListView()
Dim n As Integer
n = 2
LstTrigger.View = lvwReport
LstTrigger.Sorted = False
LstTrigger.ListItems.Clear
LstTrigger.ColumnHeaders.Clear
LstTrigger.SortOrder = lvwAscending
LstTrigger.FullRowSelect = True
LstTrigger.AllowColumnReorder = False
LstTrigger.LabelEdit = lvwManual
LstTrigger.Checkboxes = False
LstTrigger.GridLines = True
LstTrigger.MultiSelect = False
LstTrigger.HideSelection = False

LstTrigger.ColumnHeaders.Add , , PublicMsgs(13), LstTrigger.Width / n / 4
LstTrigger.ColumnHeaders.Add , , PublicEditors(13) & PublicMsgs(14), LstTrigger.Width / n * 1.5

End Sub

'*************************************************************************
'**函 数 名：LoadTriggerList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-26 23:16:41
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadTriggerList()
Dim n As Long, oItem As ListItem, tI As Integer64b, H As Integer

LstTrigger.ListItems.Clear  '清空列表

For n = 0 To UBound(TimeTrg)

  H = n - 1      '上一物品
  If H < 0 Then H = 0

      Set oItem = LstTrigger.ListItems.Add(, "Trigger_" & CStr(n), n)
      
      With oItem
       '.SubItems(1) = Format(TimeTrg(n).Check_Interval, "0.000000") & "," & Format(TimeTrg(n).Delay_Interval, "0.000000") & "," & Format(TimeTrg(n).Rearm_Interval, "0.000000")
       .SubItems(1) = TrgFunc(GetTriggerFunctionIndex(TimeTrg(n).Condition(), TimeTrg(n).Consequence())).Description
      End With

Next n

End Sub




Private Sub CApply_Click()
Dim tIdx As Integer

If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
    
    TimeTrg(CurrentTimeTrgID) = CurrentTimeTrg
     tIdx = GetTriggerFunctionIndex(TimeTrg(CurrentTimeTrgID).Condition(), TimeTrg(CurrentTimeTrgID).Consequence())
    'LstTrigger.ListItems(CurrentTimeTrgID + 1).SubItems(1) = Format(TimeTrg(CurrentTimeTrgID).Check_Interval, "0.000000") & "," & Format(TimeTrg(CurrentTimeTrgID).Delay_Interval, "0.000000") & "," & Format(TimeTrg(CurrentTimeTrgID).Rearm_Interval, "0.000000")
     LstTrigger.ListItems(CurrentTimeTrgID + 1).SubItems(1) = TrgFunc(tIdx).Description
     txtFuncName.Text = "功能名:" & TrgFunc(tIdx).FunctionName
     txtFuncDes.Text = "功能描述:" & TrgFunc(tIdx).Description
End If

End Sub

Private Sub CBChange_Change()

End Sub

Private Sub CBChange_Click()

End Sub

Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If TimeTrg(CurrentTimeTrgID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), TimeTrg(CurrentTimeTrgID).ID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
              
              If CurrentTimeTrgID < N_TimeTrg - 1 Then
                For i = CurrentTimeTrgID To N_TimeTrg - 2 Step 1
                    j = TimeTrg(i).ID
                    TimeTrg(i) = TimeTrg(i + 1)
                    TimeTrg(i).ID = j
                    LstTrigger.ListItems(i + 1).SubItems(1) = LstTrigger.ListItems(i + 2).SubItems(1)
                Next i
                
                ReDim Preserve TimeTrg(N_TimeTrg - 2)
                LstTrigger.ListItems.Remove N_TimeTrg
                N_TimeTrg = N_TimeTrg - 1
                
              Else
                ReDim Preserve TimeTrg(N_TimeTrg - 2)
                LstTrigger.ListItems.Remove N_TimeTrg
                
                N_TimeTrg = N_TimeTrg - 1
                CurrentTimeTrgID = N_TimeTrg - 1
                
              End If
               
               LstTrigger_ItemClick LstTrigger.ListItems(CurrentTimeTrgID + 1)
               LstTrigger.ListItems(CurrentTimeTrgID + 1).Selected = True
               LstTrigger.ListItems(CurrentTimeTrgID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), TimeTrg(CurrentTimeTrgID).ID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), TimeTrg(CurrentTimeTrgID).ID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then

           ReDim Preserve TimeTrg(N_TimeTrg)
           N_TimeTrg = N_TimeTrg + 1
           TimeTrg(N_TimeTrg - 1) = TimeTrg(CurrentTimeTrgID)
           With TimeTrg(N_TimeTrg - 1)
                 .ID = N_TimeTrg - 1
                 .Edit = True
           End With
           
           Set oItem = LstTrigger.ListItems.Add(, "Trigger_" & TimeTrg(N_TimeTrg - 1).ID, TimeTrg(N_TimeTrg - 1).ID)
           
                 With oItem
                    .SubItems(1) = TrgFunc(GetTriggerFunctionIndex(TimeTrg(N_TimeTrg - 1).Condition(), TimeTrg(N_TimeTrg - 1).Consequence())).Description
                    '.SubItems(1) = Format(TimeTrg(N_TimeTrg - 1).Check_Interval, "0.000000") & "," & Format(TimeTrg(N_TimeTrg - 1).Delay_Interval, "0.000000") & "," & Format(TimeTrg(N_TimeTrg - 1).Rearm_Interval, "0.000000")
                 End With
                 
           LstTrigger_ItemClick LstTrigger.ListItems(N_TimeTrg)
           LstTrigger.ListItems(N_TimeTrg).Selected = True
           LstTrigger.ListItems(N_TimeTrg).EnsureVisible
         End If
         
      Case 2
         If CurrentTimeTrgID > 0 Then
           If TimeTrg(CurrentTimeTrgID - 1).Edit And TimeTrg(CurrentTimeTrgID).Edit Then
                SwapTimeTrg CurrentTimeTrgID - 1, CurrentTimeTrgID
                SwapListItem LstTrigger.ListItems(CurrentTimeTrgID), LstTrigger.ListItems(CurrentTimeTrgID + 1), 1, True
                
               LstTrigger_ItemClick LstTrigger.ListItems(CurrentTimeTrgID)
               LstTrigger.ListItems(CurrentTimeTrgID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), TimeTrg(CurrentTimeTrgID - 1).ID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
         End If
      Case 3
        If CurrentTimeTrgID + 1 <= N_TimeTrg - 1 Then
           If TimeTrg(CurrentTimeTrgID).Edit And TimeTrg(CurrentTimeTrgID + 1).Edit Then
                SwapTimeTrg CurrentTimeTrgID, CurrentTimeTrgID + 1
                SwapListItem LstTrigger.ListItems(CurrentTimeTrgID + 1), LstTrigger.ListItems(CurrentTimeTrgID + 2), 1, True
                
                LstTrigger_ItemClick LstTrigger.ListItems(CurrentTimeTrgID + 2)
                LstTrigger.ListItems(CurrentTimeTrgID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), TimeTrg(CurrentTimeTrgID).ID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
        End If
End Select
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstTrigger, LstTrigger.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstTrigger_ItemClick(LstTrigger.SelectedItem)
End If
End Sub

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbExclamation, PublicMsgs(0)) = vbYes Then
LstTrigger_ItemClick LstTrigger.ListItems(CurrentTimeTrgID + 1)
LstTrigger.ListItems(CurrentTimeTrgID + 1).Selected = True
LstTrigger.ListItems(CurrentTimeTrgID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
'Inits
CustomActive = False

InitTriggerListView
LoadTriggerList
InitFrames

LstTrigger.ListItems(1).Selected = True
CurrentTimeTrgID = 0
CurrentTimeTrg = TimeTrg(0)
Call LoadTriggerInfo(CurrentTimeTrg)

TranslateForm Me
InitCMDs
OpCnt1.Attach OpBlockEditor1
OpCnt2.Attach OpBlockEditor2

Label2.Caption = Label2.Caption & N_TimeTrg

Loading = False
CustomActive = True
txtCalc.Text = ""
txtCalc2.Text = ""
OptHex(0).Value = True
OptHex2(0).Value = True
End Sub


Private Sub InitCMDs()

cCMD(2).Left = cCMD(3).Left - cCMD(2).Width - 200
cCMD(1).Left = cCMD(0).Left - cCMD(1).Width - 200

End Sub

Private Sub LstTrigger_Validate(Cancel As Boolean)
If Loading Then Cancel = True
End Sub

Private Sub OpBlockEditor1_Validate(Cancel As Boolean)
With CurrentTimeTrg
   OpCnt1.GetOpBlock .Condition(), True
   If LBound(.Condition) = 0 Then
      .ConditionsCount = 0
   Else
      .ConditionsCount = UBound(.Condition)
   End If
End With
End Sub

Private Sub OpBlockEditor2_Validate(Cancel As Boolean)
With CurrentTimeTrg
   OpCnt2.GetOpBlock .Consequence(), True
   'MsgBox LBound(.Consequence) & " " & UBound(.Consequence)
   If LBound(.Consequence) = 0 Then
      .ConsequencesCount = 0
   Else
      .ConsequencesCount = UBound(.Consequence)
   End If
End With
End Sub


Private Sub OptHex_Click(Index As Integer)
If CustomActive Then
       CustomActive = False
       Select Case Index
           Case 0
                txtCalc.Text = CalcText_De
           Case 1
                txtCalc.Text = RemoveUseless0(I64toHexStr(StrToI64(CalcText_De)))
           Case 2
                txtCalc.Text = RemoveUseless0(I64ToBinStr(StrToI64(CalcText_De)))
       End Select
       CustomActive = True
End If
End Sub

Private Sub OptHex2_Click(Index As Integer)
If CustomActive Then
       CustomActive = False
       Select Case Index
           Case 0
                txtCalc2.Text = CalcText_De2
           Case 1
                txtCalc2.Text = RemoveUseless0(I64toHexStr(StrToI64(CalcText_De2)))
           Case 2
                txtCalc2.Text = RemoveUseless0(I64ToBinStr(StrToI64(CalcText_De2)))
       End Select
       CustomActive = True
End If
End Sub

Private Sub LstTrigger_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim n As Integer

'If Loading = False Then
'操作块Idx清空

CurrentTimeTrgID = Val(Item.Text)
CurrentTimeTrg = TimeTrg(CurrentTimeTrgID)

LoadTriggerInfo CurrentTimeTrg
'End If

End Sub

'*************************************************************************
'**函 数 名：QueryItem
'**输    入：(ListItem)oLV,(Long)Start,(String)QueryString,(Boolean)bReverse
'**输    出：
'**功能描述：进行ListView查询功能
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-27 21:19:17
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Private Function QueryItem(oLV As ListView, ByVal Start As Long, ByVal QueryString As String, Optional bReverse As Boolean = False) As Boolean
Dim oItem As ListItem
  With oLV
    Set oItem = FindItem(oLV, Start, "0|1", QueryString, True, vbTextCompare, bReverse)
       If Not oItem Is Nothing Then
       .ListItems(oItem.Index).Selected = True
       .ListItems(oItem.Index).EnsureVisible
       QueryItem = True
       Else
        MsgBox PublicMsgs(11), vbInformation, PublicMsgs(12)
        QueryItem = False
       End If
  End With
    Set oItem = Nothing
End Function


Private Sub InitFrames()
Dim i As Integer

For i = 0 To FraProps.UBound
    With FraProps(i)
         .BorderStyle = 0
         .Top = Tab1.ClientTop
         .Left = Tab1.ClientLeft
         .Width = Tab1.ClientWidth
         .Height = Tab1.ClientHeight
         .ZOrder
           If i <> 0 Then
            .Visible = False
           End If
    End With
Next i

End Sub

Private Sub Tab1_BeforeClick(Cancel As Integer)
If Loading Then Cancel = 1
End Sub

Private Sub Tab1_Click()
Dim i As Integer, n As Integer

If CustomActive Then
For i = 0 To FraProps.UBound
    With FraProps(i)
         .Visible = i + 1 = Tab1.SelectedItem.Index
         If i = 1 Then
             'InitPropFrames
         End If
    End With
Next i

LoadTriggerInfo CurrentTimeTrg
End If

End Sub


'*************************************************************************
'**函 数 名：LoadTriggerInfo
'**输    入：Trigger As Type_Tableau_Material
'**输    出：无
'**功能描述：载入可变素材信息
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-3 10:13:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadTriggerInfo(Trigger As Type_Time_Trigger)
Dim i As Integer, tStr As String, tI As Integer64b, H As Byte

CustomActive = False

With Trigger
Select Case Tab1.SelectedItem.Index

        Case 1
            '时间
            If .Check_Interval = Val(tiOn_General(0).X) Then
                 TxtChkItv.Text = tiOn_General(0).Y
            Else
                 TxtChkItv.Text = Format(.Check_Interval, "0.000000")
            End If
            
            If .Delay_Interval = Val(tiOn_General(0).X) Then
                 TxtDelayItv.Text = tiOn_General(0).Y
            Else
                 TxtDelayItv.Text = Format(.Delay_Interval, "0.000000")
            End If
           
            If .Rearm_Interval = Val(tiOn_General(0).X) Then
                 TxtRearmItv.Text = tiOn_General(0).Y
            Else
                 TxtRearmItv.Text = Format(.Rearm_Interval, "0.000000")
            End If
            
            '功能识别
            Dim tIdx As Integer
            tIdx = GetTriggerFunctionIndex(.Condition(), .Consequence())
            txtFuncName.Text = "功能名:" & TrgFunc(tIdx).FunctionName
            txtFuncDes.Text = "功能描述:" & TrgFunc(tIdx).Description
            LoadFuncOpText TrgFunc(tIdx)
         Case 2
            '条件块
            OpCnt1.AssignOpBlock .Condition(), , "TimeTrigger_" & .ID, 1
         Case 3
            '结果块
            OpCnt2.AssignOpBlock .Consequence(), , "TimeTrigger_" & .ID, 2
            
End Select
End With

CustomActive = True

End Sub

Private Sub Timer1_Timer()
Dim D As Long, DT As Integer, tI As Integer64b, i As Integer, Index As Integer

End Sub

Public Sub ReLoadInfo()
LoadTriggerInfo CurrentTimeTrg
End Sub

Private Sub txtCalc_Change()

If CustomActive Then
If OptHex(0).Value = True Then
      CalcText_De = I64toStrNZ(StrToI64(txtCalc.Text))
ElseIf OptHex(1).Value = True Then
      CalcText_De = I64toStrNZ(HexStrToI64(txtCalc.Text))
ElseIf OptHex(2).Value = True Then
      CalcText_De = I64toStrNZ(HexStrToI64(BinToHex(txtCalc.Text)))
End If
End If

End Sub

Private Sub txtCalc2_Change()

If CustomActive Then
If OptHex2(0).Value = True Then
      CalcText_De2 = I64toStrNZ(StrToI64(txtCalc2.Text))
ElseIf OptHex2(1).Value = True Then
      CalcText_De2 = I64toStrNZ(HexStrToI64(txtCalc2.Text))
ElseIf OptHex2(2).Value = True Then
      CalcText_De2 = I64toStrNZ(HexStrToI64(BinToHex(txtCalc2.Text)))
End If
End If

End Sub

Private Sub TxtChkItv_Change()
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
If CustomActive Then
    If Trim(UCase(TxtChkItv.Text)) = UCase(tiOn_General(0).Y) Or Trim(UCase(TxtChkItv.Text)) = UCase(tiOn_General(0).Z) Then
         CurrentTimeTrg.Check_Interval = Val(tiOn_General(0).X)
    Else
         CurrentTimeTrg.Check_Interval = Val(TxtChkItv.Text)
    End If
End If
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("frmTrigger", "TxtChkItv_Change", Err.Number, Err.Description)
End Sub

Private Sub TxtDelayItv_Change()
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
If CustomActive Then
    If Trim(UCase(TxtDelayItv.Text)) = UCase(tiOn_General(0).Y) Or Trim(UCase(TxtDelayItv.Text)) = UCase(tiOn_General(0).Z) Then
         CurrentTimeTrg.Delay_Interval = Val(tiOn_General(0).X)
    Else
         CurrentTimeTrg.Delay_Interval = Val(TxtDelayItv.Text)
    End If
End If
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("frmTrigger", "TxtDelayItv_Change", Err.Number, Err.Description)
End Sub



Private Sub TxtRearmItv_Change()
    On Error GoTo errorHandle '打开错误陷阱
    '------------------------------------------------
If CustomActive Then
    If Trim(UCase(TxtRearmItv.Text)) = UCase(tiOn_General(0).Y) Or Trim(UCase(TxtRearmItv.Text)) = UCase(tiOn_General(0).Z) Then
         CurrentTimeTrg.Rearm_Interval = Val(tiOn_General(0).X)
    Else
         CurrentTimeTrg.Rearm_Interval = Val(TxtRearmItv.Text)
    End If
End If
    '------------------------------------------------
    Exit Sub
    '----------------
errorHandle:
    Call logErr("frmTrigger", "TxtRearmItv_Change", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：LoadFuncOpText
'**输    入：
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-3-26 14:21:53
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFuncOpText(TrgFunc As Type_Trigger_Function)
Dim i As Integer, n As Integer, Space As String, Idx As Integer

Space = "     "

With TrgFunc
     txtFuncOp.Text = PublicMsgs(52) & "1:" & vbCrLf & Space
     If UBound(.Opblock1) = 0 Then
          txtFuncOp.Text = txtFuncOp.Text & PublicTips(0) & vbCrLf
     Else
          For i = 1 To UBound(.Opblock1)
              Idx = GetOpIndex(RemoveOperationNegations(.Opblock1(i).Op))
              If Idx >= 0 Then
                  txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(68) & ":" & Operation(Idx).Op_CSVname & vbCrLf
              Else
                  txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(68) & ":" & .Opblock1(i).Op
              End If
              If .Opblock1(i).ParaNum > 0 Then
                  For n = 1 To .Opblock1(i).ParaNum
                    txtFuncOp.Text = txtFuncOp.Text & Space & Space
                    If .Opblock1(i).Para(n).strID <> "" Then
                       txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(38) & ":" & .Opblock1(i).Para(n).strID
                    ElseIf .Opblock1(i).Para(n).Value <> "" Then
                       txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(38) & ":" & .Opblock1(i).Para(n).Value
                    Else
                       txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(38) & ":" & PublicMsgs(92)
                    End If
                    txtFuncOp.Text = txtFuncOp.Text & vbCrLf
                  Next n
              End If
          Next i
     End If

txtFuncOp.Text = txtFuncOp.Text & vbCrLf

     txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(52) & "2:" & vbCrLf & Space
     If UBound(.OpBlock2) = 0 Then
          txtFuncOp.Text = txtFuncOp.Text & PublicTips(0) & vbCrLf
     Else
          For i = 1 To UBound(.OpBlock2)
              Idx = GetOpIndex(RemoveOperationNegations(.OpBlock2(i).Op))
              If Idx >= 0 Then
                  txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(68) & ":" & Operation(Idx).Op_CSVname & vbCrLf
              Else
                  txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(68) & ":" & .OpBlock2(i).Op
              End If
              If .OpBlock2(i).ParaNum > 0 Then
                  For n = 1 To .OpBlock2(i).ParaNum
                    txtFuncOp.Text = txtFuncOp.Text & Space & Space
                    If .OpBlock2(i).Para(n).strID <> "" Then
                       txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(38) & ":" & .OpBlock2(i).Para(n).strID
                    ElseIf .OpBlock2(i).Para(n).Value <> "" Then
                       txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(38) & ":" & .OpBlock2(i).Para(n).Value
                    Else
                       txtFuncOp.Text = txtFuncOp.Text & PublicMsgs(38) & ":" & PublicMsgs(92)
                    End If
                    txtFuncOp.Text = txtFuncOp.Text & vbCrLf
                  Next n
              End If
          Next i
     End If
End With

txtFuncOp.Text = txtFuncOp.Text & vbCrLf

txtFuncOp.Text = txtFuncOp.Text & ActiveString(PublicMsgs(93), CurrentTimeTrg.ConditionsCount) & vbCrLf
txtFuncOp.Text = txtFuncOp.Text & ActiveString(PublicMsgs(94), CurrentTimeTrg.ConsequencesCount) & vbCrLf

End Sub
