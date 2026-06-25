VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTabMat 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "可变素材编辑器"
   ClientHeight    =   9375
   ClientLeft      =   3105
   ClientTop       =   960
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_11"
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTabMat.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTabMat.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前可变素材(&O)"
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
      Left            =   6120
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
      Left            =   5160
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
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   720
      TabIndex        =   3
      Top             =   100
      Width           =   4455
   End
   Begin MSComctlLib.ListView LstTabMat 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   15055
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmTabMat.frx":0934
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps1"
      Height          =   7935
      Index           =   1
      Left            =   6240
      TabIndex        =   10
      Top             =   480
      Width           =   7815
      Begin MnBWarband_Editor.OpBlockEditor OpBlockEditor1 
         Height          =   6975
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   12303
      End
      Begin VB.Frame FraCalc 
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   7200
         Width           =   7575
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
            TabIndex        =   45
            Text            =   "txtCalc"
            Top             =   180
            Width           =   3735
         End
         Begin VB.OptionButton OptHex 
            Caption         =   "二进制"
            Height          =   255
            Index           =   2
            Left            =   6240
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptHex 
            Caption         =   "16进制"
            Height          =   255
            Index           =   1
            Left            =   5160
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptHex 
            Caption         =   "十进制"
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   42
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
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
      Height          =   7935
      Index           =   0
      Left            =   6240
      TabIndex        =   9
      Top             =   480
      Width           =   7815
      Begin VB.Frame FraEmit 
         Caption         =   "属性"
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
         Height          =   4455
         Left            =   480
         TabIndex        =   25
         Top             =   1920
         Width           =   3015
         Begin VB.Frame FraEBox 
            Caption         =   "网格最大值:"
            Height          =   1215
            Index           =   2
            Left            =   240
            TabIndex        =   36
            Top             =   3000
            Width           =   2535
            Begin VB.TextBox TxtMax 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   840
               TabIndex        =   38
               Text            =   "X"
               Top             =   300
               Width           =   1095
            End
            Begin VB.TextBox TxtMax 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   37
               Text            =   "Y"
               Top             =   670
               Width           =   1095
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "X:"
               Height          =   180
               Index           =   2
               Left            =   600
               TabIndex        =   40
               Top             =   360
               Width           =   180
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Y:"
               Height          =   180
               Index           =   2
               Left            =   600
               TabIndex        =   39
               Top             =   720
               Width           =   180
            End
         End
         Begin VB.Frame FraEBox 
            Caption         =   "网格最小值:"
            Height          =   1215
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   1680
            Width           =   2535
            Begin VB.TextBox TxtMin 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   33
               Text            =   "Y"
               Top             =   670
               Width           =   1095
            End
            Begin VB.TextBox TxtMin 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   840
               TabIndex        =   32
               Text            =   "X"
               Top             =   300
               Width           =   1095
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Y:"
               Height          =   180
               Index           =   1
               Left            =   600
               TabIndex        =   35
               Top             =   720
               Width           =   180
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "X:"
               Height          =   180
               Index           =   1
               Left            =   600
               TabIndex        =   34
               Top             =   360
               Width           =   180
            End
         End
         Begin VB.Frame FraEBox 
            Caption         =   "尺寸:"
            Height          =   1215
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   2535
            Begin VB.TextBox TxtEBox 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   840
               TabIndex        =   28
               Text            =   "H"
               Top             =   300
               Width           =   1095
            End
            Begin VB.TextBox TxtEBox 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "黑体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   840
               TabIndex        =   27
               Text            =   "W"
               Top             =   670
               Width           =   1095
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "高:"
               Height          =   180
               Index           =   0
               Left            =   480
               TabIndex        =   30
               Top             =   360
               Width           =   270
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "宽:"
               Height          =   180
               Index           =   0
               Left            =   480
               TabIndex        =   29
               Top             =   720
               Width           =   270
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "基础信息"
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
         Height          =   1335
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   6975
         Begin VB.TextBox TxtSampleName 
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
            TabIndex        =   23
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox TxtTabMatName 
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
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "样本名:"
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
            Left            =   600
            TabIndex        =   24
            Top             =   885
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "可变素材名:"
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
            Left            =   240
            TabIndex        =   22
            Top             =   405
            Width           =   1245
         End
      End
      Begin VB.Frame FraFlags 
         Caption         =   "特性(暂不可用)"
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
         Height          =   4455
         Left            =   3720
         TabIndex        =   18
         Top             =   1920
         Width           =   3735
         Begin VB.ListBox LstFlags 
            Enabled         =   0   'False
            Height          =   3840
            Left            =   170
            Style           =   1  'Checkbox
            TabIndex        =   19
            Top             =   360
            Width           =   3375
         End
      End
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   8415
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   14843
      MultiRow        =   -1  'True
      ImageList       =   "IL1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本信息(&I)"
            Key             =   "Info"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "操作块(&S)"
            Key             =   "PropBag"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label LbTest 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1560
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
      Left            =   3480
      MouseIcon       =   "frmTabMat.frx":4FA1
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   9120
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
      MouseIcon       =   "frmTabMat.frx":52AB
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   9120
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
      Left            =   5160
      MouseIcon       =   "frmTabMat.frx":55B5
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   9120
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
      Left            =   4320
      MouseIcon       =   "frmTabMat.frx":58BF
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   9120
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "可变素材数:"
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
      Top             =   9120
      Width           =   1080
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
Attribute VB_Name = "frmTabMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loading As Boolean
Dim CustomActive As Boolean
Dim CalcText_De As String
Dim OpCnt As New clsOpBlock

Private Sub InitTabMatListView()
Dim n As Integer
n = 2
LstTabMat.View = lvwReport
LstTabMat.Sorted = False
LstTabMat.ListItems.Clear
LstTabMat.ColumnHeaders.Clear
LstTabMat.SortOrder = lvwAscending
LstTabMat.FullRowSelect = True
LstTabMat.AllowColumnReorder = False
LstTabMat.LabelEdit = lvwManual
LstTabMat.Checkboxes = False
LstTabMat.GridLines = True
LstTabMat.MultiSelect = False
LstTabMat.HideSelection = False

LstTabMat.ColumnHeaders.Add , , PublicMsgs(13), LstTabMat.Width / n / 4
LstTabMat.ColumnHeaders.Add , , PublicEditors(11) & PublicMsgs(14), LstTabMat.Width / n * 1.5

End Sub

'*************************************************************************
'**函 数 名：LoadTabMatList
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
Private Sub LoadTabMatList()
Dim n As Long, oItem As ListItem, tI As Integer64b, H As Integer

LstTabMat.ListItems.Clear  '清空列表

For n = 0 To UBound(TabMat)

  H = n - 1      '上一物品
  If H < 0 Then H = 0

      Set oItem = LstTabMat.ListItems.Add(, "TabMat_" & CStr(n), n)
      
      With oItem
       .SubItems(1) = TabMat(n).strID
      End With

Next n

End Sub

Private Sub CApply_Click()
Dim q As Boolean
If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then

    If UCase(TabMat(CurrentTabMatID).strID) <> UCase(CurrentTabMat.strID) Then               '外引
        q = ChangeStrID(TabMat(CurrentTabMatID).strID, CurrentTabMat.strID)
        
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentTabMat.strID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    TabMat(CurrentTabMatID) = CurrentTabMat
    
    LstTabMat.ListItems(CurrentTabMatID + 1).SubItems(1) = TabMat(CurrentTabMatID).strID

End If

End Sub

Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If TabMat(CurrentTabMatID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), TabMat(CurrentTabMatID).strID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
              
              DelIndex TabMat(CurrentTabMatID).strID
              If CurrentTabMatID < N_TabMat - 1 Then
                For i = CurrentTabMatID To N_TabMat - 2 Step 1
                    ChangeID TabMat(i + 1).strID, TabMat(i + 1).ID - 1
                    j = TabMat(i).ID
                    TabMat(i) = TabMat(i + 1)
                    TabMat(i).ID = j
                    LstTabMat.ListItems(i + 1).SubItems(1) = LstTabMat.ListItems(i + 2).SubItems(1)
 
                Next i
                
                ReDim Preserve TabMat(N_TabMat - 2)
                LstTabMat.ListItems.Remove N_TabMat
                N_TabMat = N_TabMat - 1
                
              Else
                ReDim Preserve TabMat(N_TabMat - 2)
                LstTabMat.ListItems.Remove N_TabMat
                
                N_TabMat = N_TabMat - 1
                CurrentTabMatID = N_TabMat - 1
                
              End If
               
               LstTabMat_ItemClick LstTabMat.ListItems(CurrentTabMatID + 1)
               LstTabMat.ListItems(CurrentTabMatID + 1).Selected = True
               LstTabMat.ListItems(CurrentTabMatID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), TabMat(CurrentTabMatID).strID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), TabMat(CurrentTabMatID).strID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_TabMat, TabMat(CurrentTabMatID).strID & "_New") Then
           ReDim Preserve TabMat(N_TabMat)
           N_TabMat = N_TabMat + 1
           TabMat(N_TabMat - 1) = TabMat(CurrentTabMatID)
           With TabMat(N_TabMat - 1)
                 .ID = N_TabMat - 1
                 .strID = .strID & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstTabMat.ListItems.Add(, "TabMat_" & TabMat(N_TabMat - 1).ID, TabMat(N_TabMat - 1).ID)
           
                 With oItem
                    .SubItems(1) = TabMat(N_TabMat - 1).strID
                 End With
                 
           LstTabMat_ItemClick LstTabMat.ListItems(N_TabMat)
           LstTabMat.ListItems(N_TabMat).Selected = True
           LstTabMat.ListItems(N_TabMat).EnsureVisible
           
           Else
           
           MsgBox ActiveString(PublicMsgs(90), TabMat(CurrentTabMatID).strID & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
         
      Case 2
         If CurrentTabMatID > 0 Then
           If TabMat(CurrentTabMatID - 1).Edit And TabMat(CurrentTabMatID).Edit Then
                SwapID TabMat(CurrentTabMatID - 1).strID, TabMat(CurrentTabMatID).strID
                SwapTabMat CurrentTabMatID - 1, CurrentTabMatID
                SwapListItem LstTabMat.ListItems(CurrentTabMatID), LstTabMat.ListItems(CurrentTabMatID + 1), 1, True
                
               LstTabMat_ItemClick LstTabMat.ListItems(CurrentTabMatID)
               LstTabMat.ListItems(CurrentTabMatID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), TabMat(CurrentTabMatID - 1).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
         End If
      Case 3
        If CurrentTabMatID + 1 <= N_TabMat - 1 Then
           If TabMat(CurrentTabMatID).Edit And TabMat(CurrentTabMatID + 1).Edit Then
                SwapID TabMat(CurrentTabMatID).strID, TabMat(CurrentTabMatID + 1).strID
                SwapTabMat CurrentTabMatID, CurrentTabMatID + 1
                SwapListItem LstTabMat.ListItems(CurrentTabMatID + 1), LstTabMat.ListItems(CurrentTabMatID + 2), 1, True
                
                LstTabMat_ItemClick LstTabMat.ListItems(CurrentTabMatID + 2)
                LstTabMat.ListItems(CurrentTabMatID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), TabMat(CurrentTabMatID).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
        End If
End Select
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstTabMat, LstTabMat.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstTabMat_ItemClick(LstTabMat.SelectedItem)
End If
End Sub

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbExclamation, PublicMsgs(0)) = vbYes Then
LstTabMat_ItemClick LstTabMat.ListItems(CurrentTabMatID + 1)
LstTabMat.ListItems(CurrentTabMatID + 1).Selected = True
LstTabMat.ListItems(CurrentTabMatID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
'Inits
CustomActive = False

InitTabMatListView
LoadTabMatList
InitFrames
InitLstFlags

LstTabMat.ListItems(1).Selected = True
CurrentTabMatID = 0
CurrentTabMat = TabMat(0)
Call LoadTabMatInfo(CurrentTabMat)

TranslateForm Me
InitCMDs
OpCnt.Attach OpBlockEditor1

Label2.Caption = Label2.Caption & N_TabMat

Loading = False
CustomActive = True
txtCalc.Text = ""
OptHex(0).Value = True

End Sub

Private Sub LstFlags_ItemCheck(Item As Integer)
Dim i As Integer

If CustomActive Then
     If Item >= 3 And Item <= 5 Then
         CustomActive = False
         For i = 3 To 5
              LstFlags.Selected(i) = False
         Next i
         LstFlags.Selected(Item) = True
         CustomActive = True
     End If
End If

End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

Private Sub LstFlags_LostFocus()
Dim i As Integer, tI As Integer64b

With CurrentTabMat
    'CurrentTabMat.Flags = ""
    For i = 0 To UBound(PSf)
          If LstFlags.Selected(i) Then
             tI = AddFlagsI64(tI, HexStrToI64(PSf(i).X))
          End If
    Next i
    .Flags = I64toStrNZ(tI)
End With

End Sub

Private Sub LstTabMat_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim n As Integer

'操作块Idx清空

CurrentTabMatID = Val(Item.Text)
CurrentTabMat = TabMat(CurrentTabMatID)

LoadTabMatInfo CurrentTabMat

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

Private Sub LstTabMat_Validate(Cancel As Boolean)
If Loading Then Cancel = True
End Sub


Private Sub OpBlockEditor1_Validate(Cancel As Boolean)
OpCnt.GetOpBlock CurrentTabMat.OpBlock(), True
If LBound(CurrentTabMat.OpBlock) = 0 Then
  CurrentTabMat.OpCount = 0
Else
  CurrentTabMat.OpCount = UBound(CurrentTabMat.OpBlock)
End If

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



Private Sub Tab1_BeforeClick(Cancel As Integer)
If laoding Then Cancel = 1
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

LoadTabMatInfo CurrentTabMat
End If

End Sub

Private Sub InitLstFlags()
Dim i As Integer

LstFlags.Clear

End Sub

'*************************************************************************
'**函 数 名：LoadTabMatInfo
'**输    入：TabMat As Type_Tableau_Material
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
Private Sub LoadTabMatInfo(TabMat As Type_Tableau_Material)
Dim i As Integer, tStr As String, tI As Integer64b, H As Byte

CustomActive = False
TxtTabMatName.Text = ""

With TabMat
Select Case Tab1.SelectedItem.Index

        Case 1
            '显示信息
            TxtTabMatName.Text = TabMat.strID
            TxtSampleName.Text = TabMat.Sample
    
            '属性
            TxtEBox(0).Text = .Height
            TxtEBox(1).Text = .Width
            TxtMin(0).Text = .Min.X
            TxtMin(1).Text = .Min.Y
            TxtMax(0).Text = .Max.X
            TxtMax(1).Text = .Max.Y
            
            'Flags
             'For i = 0 To LstFlags.ListCount - 1
             '       LstFlags.Selected(i) = False
             'Next i
             'LoadLstFlags TabMat.Flags

         Case 2
            '操作块
            OpCnt.AssignOpBlock TabMat.OpBlock(), , TabMat.strID, 1

End Select
End With

CustomActive = True

End Sub

'*************************************************************************
'**函 数 名：LoadLstFlags
'**输    入：-(String)Flags
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-3 10:57:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadLstFlags(Flags As String)

End Sub

Private Sub Timer1_Timer()
Dim D As Long, DT As Integer, tI As Integer64b, i As Integer, Index As Integer

LbTest.Caption = ChangeTag

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

Private Sub TxtEBox_LostFocus(Index As Integer)

Select Case Index
     Case 0
         CurrentTabMat.Height = CLng(Val(TxtEBox(Index).Text))
     Case 1
         CurrentTabMat.Width = CLng(Val(TxtEBox(Index).Text))
End Select

End Sub


Private Sub TxtMax_LostFocus(Index As Integer)

Select Case Index
     Case 0
         CurrentTabMat.Max.X = CLng(Val(TxtMax(Index).Text))
     Case 1
         CurrentTabMat.Max.Y = CLng(Val(TxtMax(Index).Text))
End Select

End Sub

Private Sub TxtMin_LostFocus(Index As Integer)

Select Case Index
     Case 0
         CurrentTabMat.Min.X = CLng(Val(TxtMin(Index).Text))
     Case 1
         CurrentTabMat.Min.Y = CLng(Val(TxtMin(Index).Text))
End Select

End Sub

Private Sub TxtSampleName_LostFocus()

TxtSampleName.Text = Replace(TxtSampleName.Text, " ", "_")
CurrentTabMat.Sample = CStr(TxtSampleName.Text)

End Sub

Private Sub TxtTabMatName_LostFocus()

TxtTabMatName.Text = Replace(TxtTabMatName.Text, " ", "_")
CurrentTabMat.strID = CStr(TxtTabMatName.Text)

End Sub


Public Sub ReLoadInfo()
LoadTabMatInfo CurrentTabMat
End Sub
