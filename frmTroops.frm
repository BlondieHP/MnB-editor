VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTroops 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "兵种编辑器"
   ClientHeight    =   9375
   ClientLeft      =   4200
   ClientTop       =   1050
   ClientWidth     =   14355
   ForeColor       =   &H00C00000&
   Icon            =   "frmTroops.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_1"
   Begin VB.Timer Timer_MousePos_Trp 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   600
   End
   Begin VB.Timer Timer_MousePos 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer Timer_KillTip 
      Interval        =   1000
      Left            =   5760
      Top             =   0
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前兵种(&O)"
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
      TabIndex        =   97
      Top             =   8760
      Width           =   2415
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
      TabIndex        =   62
      Top             =   80
      Width           =   375
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
      TabIndex        =   61
      Top             =   80
      Width           =   375
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   600
      TabIndex        =   60
      Top             =   80
      Width           =   4575
   End
   Begin VB.CommandButton CReset 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   35
      Top             =   8760
      Width           =   2175
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
      TabIndex        =   34
      Top             =   8760
      Width           =   2175
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   240
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTroops.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTroops.frx":13CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTroops.frx":1526
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTroops.frx":1680
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTroops.frx":17DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstTroops 
      Height          =   8535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   15055
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "frmTroops.frx":1934
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
      Height          =   5175
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Top             =   600
      Width           =   4335
      Begin VB.Frame Frame1 
         Caption         =   "武器熟练度"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   3615
         Index           =   5
         Left            =   2040
         TabIndex        =   40
         Top             =   3120
         Width           =   1695
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   16
            Left            =   1080
            TabIndex        =   53
            Text            =   "0"
            Top             =   3120
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   15
            Left            =   1080
            TabIndex        =   51
            Text            =   "0"
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   1080
            TabIndex        =   45
            Text            =   "0"
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   1080
            TabIndex        =   44
            Text            =   "0"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   1080
            TabIndex        =   43
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   1080
            TabIndex        =   42
            Text            =   "0"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   1080
            TabIndex        =   41
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "火器:"
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
            Index           =   18
            Left            =   480
            TabIndex        =   54
            Top             =   3195
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "投掷:"
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
            Index           =   17
            Left            =   480
            TabIndex        =   52
            Top             =   2715
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "弩:"
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
            Index           =   16
            Left            =   675
            TabIndex        =   50
            Top             =   2235
            Width           =   300
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "弓箭:"
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
            Index           =   15
            Left            =   480
            TabIndex        =   49
            Top             =   1755
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "长杆:"
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
            Index           =   13
            Left            =   495
            TabIndex        =   48
            Top             =   1275
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "双手:"
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
            Index           =   12
            Left            =   495
            TabIndex        =   47
            Top             =   795
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单手:"
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
            Index           =   10
            Left            =   450
            TabIndex        =   46
            Top             =   315
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "技能"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   4575
         Index           =   2
         Left            =   3840
         TabIndex        =   31
         Top             =   3120
         Width           =   3735
         Begin MSComctlLib.ListView LstSkills 
            Height          =   4095
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   7223
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "属性"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2895
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   1695
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   960
            TabIndex        =   25
            Text            =   "0"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   960
            TabIndex        =   24
            Text            =   "0"
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   960
            TabIndex        =   23
            Text            =   "0"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   960
            TabIndex        =   22
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtPropBag 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   960
            TabIndex        =   21
            Text            =   "0"
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "等级:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   9
            Left            =   330
            TabIndex        =   30
            Top             =   435
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "力量:"
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
            Index           =   8
            Left            =   375
            TabIndex        =   29
            Top             =   915
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "敏捷:"
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
            Index           =   7
            Left            =   375
            TabIndex        =   28
            Top             =   1395
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "智力:"
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
            Index           =   6
            Left            =   360
            TabIndex        =   27
            Top             =   1875
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "魅力:"
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
            Index           =   5
            Left            =   360
            TabIndex        =   26
            Top             =   2355
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "识别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2775
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   7335
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   15
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   4
            Left            =   1680
            TabIndex        =   19
            Top             =   2160
            Width           =   5055
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   3
            Left            =   1680
            TabIndex        =   17
            Top             =   1680
            Width           =   5055
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   13
            Top             =   720
            Width           =   5055
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   11
            Top             =   240
            Width           =   5055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "兵种复数(NOW):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   4
            Left            =   165
            TabIndex        =   18
            Top             =   2235
            Width           =   1410
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "兵种名(NOW):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   3
            Left            =   360
            TabIndex        =   16
            Top             =   1755
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "兵种复数(EN):"
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
            Index           =   2
            Left            =   285
            TabIndex        =   14
            Top             =   1275
            Width           =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "兵种名(EN):"
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
            Index           =   1
            Left            =   480
            TabIndex        =   12
            Top             =   795
            Width           =   1110
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "兵种ID:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   0
            Left            =   840
            TabIndex        =   10
            Top             =   315
            Width           =   705
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps4"
      Height          =   4815
      Index           =   4
      Left            =   9000
      TabIndex        =   6
      Top             =   1560
      Width           =   4695
      Begin VB.Frame Frame1 
         Caption         =   "面容代码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1935
         Index           =   10
         Left            =   240
         TabIndex        =   87
         Top             =   360
         Width           =   7335
         Begin VB.TextBox txtFace 
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   91
            Top             =   1080
            Width           =   5895
         End
         Begin VB.TextBox txtFace 
            Height          =   375
            Index           =   0
            Left            =   1080
            TabIndex        =   90
            Top             =   400
            Width           =   5895
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "面容2:"
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
            Index           =   21
            Left            =   465
            TabIndex        =   89
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "面容1:"
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
            Index           =   19
            Left            =   465
            TabIndex        =   88
            Top             =   480
            Width           =   600
         End
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   495
         Left            =   1320
         Shape           =   4  'Rounded Rectangle
         Top             =   3000
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   3375
         Left            =   1320
         Picture         =   "frmTroops.frx":5FA1
         Top             =   3120
         Width           =   4845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*面部代码请在编辑模式下按Ctrl+E获取。代码位置如图所示:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   23
         Left            =   360
         TabIndex        =   98
         Top             =   2520
         Width           =   5325
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps1"
      Height          =   1455
      Index           =   1
      Left            =   6720
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
      Begin VB.Frame Frame1 
         Caption         =   "识别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1095
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   240
         Width           =   7335
         Begin VB.ComboBox cbFaction 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   240
            Width           =   5535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阵营:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   14
            Left            =   840
            TabIndex        =   37
            Top             =   315
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "初始场景"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2055
         Index           =   7
         Left            =   240
         TabIndex        =   63
         Top             =   4200
         Width           =   7335
         Begin VB.TextBox txtEntry 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   6480
            TabIndex        =   94
            Text            =   "0"
            Top             =   1250
            Width           =   615
         End
         Begin VB.ComboBox cbScenes 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   480
            Width           =   6735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "具体站位:"
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
            Index           =   22
            Left            =   5490
            TabIndex        =   93
            Top             =   1320
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "升级"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2415
         Index           =   4
         Left            =   240
         TabIndex        =   39
         Top             =   1560
         Width           =   7335
         Begin VB.ComboBox CbUpgrade 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1440
            Width           =   5415
         End
         Begin VB.ComboBox CbUpgrade 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   480
            Width           =   5415
         End
         Begin VB.Line Line2 
            Index           =   2
            X1              =   1440
            X2              =   1320
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   1560
            X2              =   1440
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   1560
            X2              =   1440
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line1 
            X1              =   1440
            X2              =   1440
            Y1              =   720
            Y2              =   1680
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "升级成:"
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
            Index           =   11
            Left            =   600
            TabIndex        =   56
            Top             =   1080
            Width           =   690
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps3"
      Height          =   7935
      Index           =   3
      Left            =   6360
      TabIndex        =   7
      Top             =   480
      Width           =   7695
      Begin VB.TextBox txtQuery2 
         Height          =   330
         Left            =   2280
         TabIndex        =   109
         Top             =   240
         Width           =   4575
      End
      Begin VB.CommandButton CQuery2 
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
         Left            =   6840
         TabIndex        =   108
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CQuery2 
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
         Left            =   7200
         TabIndex        =   107
         Top             =   240
         Width           =   375
      End
      Begin MSComctlLib.ListView LstProperties 
         Height          =   2535
         Left            =   120
         TabIndex        =   58
         Top             =   5040
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4471
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView boxItem 
         Height          =   4095
         Left            =   120
         TabIndex        =   101
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7223
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
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
         Left            =   1800
         TabIndex        =   110
         Top             =   315
         Width           =   495
      End
      Begin VB.Label bCMD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "全不选(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   4
         Left            =   4845
         MouseIcon       =   "frmTroops.frx":A8B2
         MousePointer    =   99  'Custom
         TabIndex        =   106
         Top             =   4800
         Width           =   900
      End
      Begin VB.Label bCMD 
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
         Left            =   6840
         MouseIcon       =   "frmTroops.frx":ABBC
         MousePointer    =   99  'Custom
         TabIndex        =   105
         Top             =   7680
         Width           =   705
      End
      Begin VB.Label bCMD 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "↓添加所选物品(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   1
         Left            =   5880
         MouseIcon       =   "frmTroops.frx":AEC6
         MousePointer    =   99  'Custom
         TabIndex        =   104
         Top             =   4800
         Width           =   1680
      End
      Begin VB.Label bCMD 
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
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   2
         Left            =   5160
         MouseIcon       =   "frmTroops.frx":B1D0
         MousePointer    =   99  'Custom
         TabIndex        =   103
         Top             =   7680
         Width           =   705
      End
      Begin VB.Label bCMD 
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
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   3
         Left            =   6000
         MouseIcon       =   "frmTroops.frx":B4DA
         MousePointer    =   99  'Custom
         TabIndex        =   102
         Top             =   7680
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拥有物品:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   25
         Left            =   120
         TabIndex        =   100
         Top             =   4800
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "可选择物品:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   24
         Left            =   120
         TabIndex        =   99
         Top             =   360
         Width           =   1080
      End
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   8415
      Left            =   6120
      TabIndex        =   3
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   14843
      MultiRow        =   -1  'True
      ImageList       =   "IL1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本属性(&P)"
            Key             =   "PropBag"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "关系设定(&S)"
            Key             =   "WorldSets"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "特性(&E)"
            Key             =   "Special"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "物品栏(&I)"
            Key             =   "ItemBag"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "面容(&F)"
            Key             =   "Face"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps2"
      Height          =   7935
      Index           =   2
      Left            =   6240
      TabIndex        =   8
      Top             =   480
      Width           =   7815
      Begin VB.Frame Frame1 
         Caption         =   "其他特性"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2535
         Index           =   9
         Left            =   240
         TabIndex        =   70
         Top             =   4080
         Width           =   7335
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   15
            Left            =   840
            TabIndex        =   86
            Top             =   1920
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   7
            Left            =   3840
            TabIndex        =   78
            Top             =   1560
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   6
            Left            =   3840
            TabIndex        =   77
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   5
            Left            =   3840
            TabIndex        =   76
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   4
            Left            =   3840
            TabIndex        =   75
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   74
            Top             =   1200
            Width           =   2295
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   73
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   72
            Top             =   480
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   71
            Top             =   1560
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "武器特性"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2055
         Index           =   8
         Left            =   240
         TabIndex        =   69
         Top             =   1680
         Width           =   7335
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   14
            Left            =   3840
            TabIndex        =   85
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   13
            Left            =   3840
            TabIndex        =   84
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   12
            Left            =   3840
            TabIndex        =   83
            Top             =   360
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   11
            Left            =   840
            TabIndex        =   82
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   10
            Left            =   840
            TabIndex        =   81
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   80
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   79
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   855
         Index           =   6
         Left            =   240
         TabIndex        =   64
         Top             =   480
         Width           =   7335
         Begin VB.OptionButton OptSex 
            Caption         =   "男"
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
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   68
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptSex 
            Caption         =   "女"
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
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   67
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton OptSex 
            Caption         =   "其他"
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
            Height          =   255
            Index           =   2
            Left            =   4440
            TabIndex        =   66
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cbSkins 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5400
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   300
            Width           =   1095
         End
      End
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
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   3600
      MouseIcon       =   "frmTroops.frx":B7E4
      MousePointer    =   99  'Custom
      TabIndex        =   96
      Top             =   9135
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
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   2
      Left            =   2760
      MouseIcon       =   "frmTroops.frx":BAEE
      MousePointer    =   99  'Custom
      TabIndex        =   95
      Top             =   9135
      Width           =   705
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
      Left            =   120
      TabIndex        =   59
      Top             =   160
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "兵种数:"
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
      TabIndex        =   33
      Top             =   9120
      Width           =   690
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
      ForeColor       =   &H00004000&
      Height          =   180
      Index           =   1
      Left            =   4440
      MouseIcon       =   "frmTroops.frx":BDF8
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   9135
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
      Left            =   5280
      MouseIcon       =   "frmTroops.frx":C102
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   9135
      Width           =   705
   End
End
Attribute VB_Name = "frmTroops"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CustomActive As Boolean
Dim HitLV As ListView
Dim LstItemMPos As POINTAPI

Private Sub bCMD_Click(Index As Integer)
On Error GoTo EL
Dim Index1 As Long, Index2 As Long
Dim i As Integer, q As Boolean

CustomActive = False
If LstProperties.ListItems.Count > 0 Then
Index1 = LstProperties.SelectedItem.Index
Select Case Index
       Case 2
            If Index1 >= 2 Then
                SwapInventory CurrentTrp.lstInventory(Index1), CurrentTrp.lstInventory(Index1 - 1)
                SwapListItem LstProperties.ListItems(Index1), LstProperties.ListItems(Index1 - 1), 2
                LstProperties.ListItems(Index1 - 1).Selected = True
                LstProperties.ListItems(Index1 - 1).EnsureVisible
            End If
       Case 3
            If Index1 <= LstProperties.ListItems.Count - 1 Then
                SwapInventory CurrentTrp.lstInventory(Index1), CurrentTrp.lstInventory(Index1 + 1)
                SwapListItem LstProperties.ListItems(Index1), LstProperties.ListItems(Index1 + 1), 2
                LstProperties.ListItems(Index1 + 1).Selected = True
                LstProperties.ListItems(Index1 + 1).EnsureVisible
            End If
       Case 0
            For i = 1 To LstProperties.ListItems.Count
                If LstProperties.ListItems(i).Checked Then
                   CurrentTrp.lstInventory(i).X = -1
                   CurrentTrp.lstInventory(i).strX = ""
                   CurrentTrp.lstInventory(i).Y = 0
                End If
            Next i
            StructureTroopInventory
            LoadTroopPropertiesList
       Case 1
            For i = 1 To boxItem.ListItems.Count
                If boxItem.ListItems(i).Checked Then
                    q = AddTroopInventory(i - 1)
                    If Not q Then
                       MsgBox PublicMsgs(45), vbCritical, PublicMsgs(19)
                       Exit For
                    End If
                End If
                boxItem.ListItems(i).Checked = False
            Next i
            LoadTroopPropertiesList
       Case 4
            For i = 1 To boxItem.ListItems.Count
                If boxItem.ListItems(i).Checked Then
                    boxItem.ListItems(i).Checked = False
                End If
            Next i
End Select
End If

CustomActive = True

Exit Sub

EL:
  Call logErr("frmTroops", "bCMD_Click(" & Index & ")", Err.Number, Err.Description)
End Sub


Private Sub boxItem_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = vbKeyC Then
     frmInfo.CopyInfotoClipBoard
     frmTip.ShowTip PublicMsgs(130)
     Timer_KillTip.Enabled = True
End If
End Sub

Private Sub boxItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not frmMain.mBanfrmInfo.Checked Then
   LstItemMPos.X = X
   LstItemMPos.Y = Y
   Set HitLV = boxItem
   
   Timer_MousePos.Enabled = True
   
   If Timer_KillTip.Enabled = True Then
      frmTip.HideTip
      Timer_KillTip.Enabled = False
   End If
End If

End Sub

Private Sub CApply_Click()
On Error GoTo EL
Dim q As Boolean

If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
    
    If UCase(Trps(CurrentTrpID).strID) <> UCase(CurrentTrp.strID) Then               '外引
        q = ChangeStrID(Trps(CurrentTrpID).strID, CurrentTrp.strID)
        
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentTrp.strID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    
    CurrentTrp.Scene = Val("&H" & Hex(CurrentTrp.Entry) & Right("0000" & Hex(CurrentTrp.SceneID), 4))
    
    Trps(CurrentTrpID) = CurrentTrp
    LstTroops.ListItems(CurrentTrpID + 1).SubItems(1) = Trps(CurrentTrpID).csvName
    LstTroops.ListItems(CurrentTrpID + 1).SubItems(2) = Trps(CurrentTrpID).strID
    
    CurrentTrp = Trps(CurrentTrpID)
    LoadTroopInfo
End If

Exit Sub

EL:
  Call logErr("frmTroops", "CApply_Click", Err.Number, Err.Description)
End Sub

Private Sub cbFaction_Click()
On Error GoTo EL
If CustomActive Then
CurrentTrp.Faction = cbFaction.ListIndex
CurrentTrp.Faction_strID = Factions(CurrentTrp.Faction).strID     '外引

End If

Exit Sub

EL:
  Call logErr("frmTroops", "cbFaction_Click", Err.Number, Err.Description)
End Sub

Private Sub cbFaction_Scroll()
Call cbFaction_Click
End Sub




Private Sub cbScenes_Click()
On Error GoTo EL
Dim strHex As String
If CustomActive Then

CurrentTrp.SceneID = cbScenes.ListIndex
CurrentTrp.Scene_strID = Scenes(CurrentTrp.SceneID).strID

End If

Exit Sub

EL:
  Call logErr("frmTroops", "cbScenes_Click", Err.Number, Err.Description)
End Sub

Private Sub cbScenes_Scroll()
Call cbScenes_Click
End Sub





Private Sub cbSkins_Click()
On Error GoTo EL
Dim NewFlags As String, tI As Integer64b
If CustomActive Then

       NewFlags = CStr(cbSkins.ListIndex + Val(tf_undead))

       With CurrentTrp
            tI = DeleteFlags64b(StrToI64(.Flags), HexStrToI64(troop_type_mask))
            tI = AddFlags64b(tI, StrToI64(NewFlags))
            
            .Flags = I64toStrNZ(tI)
       End With
End If

Exit Sub

EL:
  Call logErr("frmTroops", "cbSkins_Click", Err.Number, Err.Description)
End Sub

Private Sub cbSkins_Scroll()
Call cbSkins_Click
End Sub

Private Sub CbUpgrade_Click(Index As Integer)
On Error GoTo EL
If CustomActive Then
Select Case Index
     Case 0
     CurrentTrp.Upgrade1 = CbUpgrade(Index).ListIndex
     CurrentTrp.Upgrade1_strID = Trps(CurrentTrp.Upgrade1).strID
     Case 1
     CurrentTrp.Upgrade2 = CbUpgrade(Index).ListIndex
     CurrentTrp.Upgrade2_strID = Trps(CurrentTrp.Upgrade2).strID
End Select
End If

Exit Sub

EL:
  Call logErr("frmTroops", "CbUpgrade_Click", Err.Number, Err.Description)
End Sub

Private Sub CbUpgrade_Scroll(Index As Integer)
Call CbUpgrade_Click(Index)
End Sub

Private Sub cCMD_Click(Index As Integer)
On Error GoTo EL
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If Trps(CurrentTrpID).Edit Then
           
          If MsgBox(ActiveString(PublicMsgs(2), Trps(CurrentTrpID).strID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
                   
              If CurrentTrpID < N_Troop - 1 Then
              DelIndex Trps(CurrentTrpID).strID
                For i = CurrentTrpID To N_Troop - 2 Step 1
                    ChangeID Trps(i + 1).strID, Trps(i + 1).ID - 1
                    j = Trps(i).ID
                    Trps(i) = Trps(i + 1)
                    Trps(i).ID = j
                    LstTroops.ListItems(i + 1).SubItems(1) = LstTroops.ListItems(i + 2).SubItems(1)
                    LstTroops.ListItems(i + 1).SubItems(2) = LstTroops.ListItems(i + 2).SubItems(2)
                Next i
                
                ReDim Preserve Trps(N_Troop - 2)
                LstTroops.ListItems.Remove N_Troop
                N_Troop = N_Troop - 1
                
              Else
                DelIndex Trps(CurrentTrpID).strID
                ReDim Preserve Trps(N_Troop - 2)
                LstTroops.ListItems.Remove N_Troop
                
                N_Troop = N_Troop - 1
                CurrentTrpID = N_Troop - 1
                
              End If
               
               LstTroops_ItemClick LstTroops.ListItems(CurrentTrpID + 1)
               LstTroops.ListItems(CurrentTrpID + 1).Selected = True
               LstTroops.ListItems(CurrentTrpID + 1).EnsureVisible
               
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), Trps(CurrentTrpID).strID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), Trps(CurrentTrpID).strID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
 
           If AddIndex(N_Troop, Trps(CurrentTrpID).strID & "_New") Then
           ReDim Preserve Trps(N_Troop)
           N_Troop = N_Troop + 1
           Trps(N_Troop - 1) = Trps(CurrentTrpID)
           With Trps(N_Troop - 1)
                 .ID = N_Troop - 1
                 .strID = .strID & "_New"
                 .strName = .strName & "_New"
                 .strPtName = .strPtName & "_New"
                 .csvName = .csvName & "_New"
                 .csvName_pl = .csvName_pl & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstTroops.ListItems.Add(, "trp_" & Trps(N_Troop - 1).ID, Trps(N_Troop - 1).ID)
      
                 With oItem
                    .SubItems(1) = Trps(N_Troop - 1).csvName
                    .SubItems(2) = Trps(N_Troop - 1).strID
                 End With
           LstTroops_ItemClick LstTroops.ListItems(N_Troop)
           LstTroops.ListItems(N_Troop).Selected = True
           LstTroops.ListItems(N_Troop).EnsureVisible
           
           Else
           
           MsgBox ActiveString(PublicMsgs(90), Trps(CurrentTrpID).strID & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
         
      Case 2
         If CurrentTrpID > 0 Then
           If Trps(CurrentTrpID - 1).Edit And Trps(CurrentTrpID).Edit Then
           
                SwapID Trps(CurrentTrpID - 1).strID, Trps(CurrentTrpID).strID
                SwapTroops CurrentTrpID - 1, CurrentTrpID
                SwapListItem LstTroops.ListItems(CurrentTrpID), LstTroops.ListItems(CurrentTrpID + 1), 2, True
                
               LstTroops_ItemClick LstTroops.ListItems(CurrentTrpID)
               LstTroops.ListItems(CurrentTrpID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), Trps(CurrentTrpID - 1).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
         End If
      Case 3
        If CurrentTrpID + 1 <= N_Troop - 1 Then
           If Trps(CurrentTrpID).Edit And Trps(CurrentTrpID + 1).Edit Then
           
                SwapID Trps(CurrentTrpID).strID, Trps(CurrentTrpID + 1).strID
                SwapTroops CurrentTrpID, CurrentTrpID + 1
                SwapListItem LstTroops.ListItems(CurrentTrpID + 1), LstTroops.ListItems(CurrentTrpID + 2), 2, True
                
                LstTroops_ItemClick LstTroops.ListItems(CurrentTrpID + 2)
                LstTroops.ListItems(CurrentTrpID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), Trps(CurrentTrpID).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
           
        End If
End Select

InitTroopUpgradeCombo

Exit Sub

EL:
  Call logErr("frmTroops", "cCMD_Click(" & Index & ")", Err.Number, Err.Description)
End Sub

Private Sub chkFlags_Click(Index As Integer)
On Error GoTo EL
Dim tI As Integer64b
If CustomActive Then
If chkFlags(Index).Value = 1 Then
    tI = AddFlags64b(StrToI64(CurrentTrp.Flags), HexStrToI64(chkFlags(Index).Tag))
    CurrentTrp.Flags = I64toStrNZ(tI)
Else
    tI = DeleteFlags64b(StrToI64(CurrentTrp.Flags), HexStrToI64(chkFlags(Index).Tag))
    CurrentTrp.Flags = I64toStrNZ(tI)
End If
End If

Exit Sub

EL:
  Call logErr("frmTroops", "chkFlags_Click(" & Index & ")", Err.Number, Err.Description)
End Sub



Private Sub COutputLine_Click()
On Error GoTo EL
Dim t As String

frmLine.ShowTxtLine Me.Tag, -1

t = ExportTroopPYCode(CurrentTrp, True)
OutAsDebugTex t, PublicEditors_Simplified(1) & ":" & CurrentTrp.strID
Exit Sub

EL:
  Call logErr("frmTroops", "COutputLine_Click", Err.Number, Err.Description)
End Sub

Private Sub CQuery_Click(Index As Integer)
On Error GoTo EL
Dim q As Boolean
q = QueryItem(LstTroops, LstTroops.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstTroops_ItemClick(LstTroops.SelectedItem)
End If

Exit Sub

EL:
  Call logErr("frmTroops", "CQuery_Click(" & Index & ")", Err.Number, Err.Description)
End Sub

Private Sub CQuery2_Click(Index As Integer)
On Error GoTo EL
Dim q As Boolean
q = QueryItem(boxItem, boxItem.SelectedItem.Index, txtQuery2.Text, CBool(Index))

Exit Sub

EL:
  Call logErr("frmTroops", "CQuery_Click2(" & Index & ")", Err.Number, Err.Description)
End Sub



Private Sub CReset_Click()
On Error GoTo EL
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
LstTroops_ItemClick LstTroops.ListItems(CurrentTrpID + 1)
LstTroops.ListItems(CurrentTrpID + 1).Selected = True
LstTroops.ListItems(CurrentTrpID + 1).EnsureVisible
End If

Exit Sub

EL:
  Call logErr("frmTroops", "CReset_Click", Err.Number, Err.Description)
End Sub

Private Sub Form_Load()
On Error GoTo EL

CustomActive = False

InitFrames

InitSkinsCombo
InitTroopUpgradeCombo
InitTroopsListView
InitSkillsListView
InitItemBox
InitPropertiesListView
InitFlagsList

LoadTroopsList
LoadFactionCombo
'LoadPropertiesCombo
LoadItemBox
LoadScenesCombo
CurrentTrpID = 0
CurrentTrp = Trps(CurrentTrpID)
LoadTroopInfo

TranslateForm Me
InitCMDs

Label2.Caption = Label2.Caption & N_Troop
Label1(3).Caption = Replace(Label1(3).Caption, "(NOW):", "(" & UCase$(MnBInfo.Language) & "):", , , vbTextCompare)
Label1(4).Caption = Replace(Label1(4).Caption, "(NOW):", "(" & UCase$(MnBInfo.Language) & "):", , , vbTextCompare)

If LCase$(MnBInfo.Language) = "en" Then
    txtPropBag(3).Enabled = False
    txtPropBag(4).Enabled = False
End If

CustomActive = True

Me.Show

Exit Sub

EL:
  Call logErr("frmTroops", "Form_Load", Err.Number, Err.Description)
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub
'*************************************************************************
'**函 数 名：LoadTroopsList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-23 12:58:23
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadTroopsList()
Dim n As Long, oItem As ListItem

LstTroops.ListItems.Clear
For n = 0 To N_Troop - 1
    'If Trim(trps(n).strId) <> "" Then

      Set oItem = LstTroops.ListItems.Add(, "trp_" & Trps(n).ID, Trps(n).ID)
      
      With oItem
         .SubItems(1) = Trps(n).csvName
         .SubItems(2) = Trps(n).strID
      End With
    'End If
Next n

End Sub

'*************************************************************************
'**函 数 名：LoadFactionCombo
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-24 23:55:41
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFactionCombo()
Dim n As Long, oItem As ListItem

cbFaction.Clear
For n = 0 To UBound(Factions)
      cbFaction.AddItem Factions(n).csvName
Next n

End Sub

Private Sub InitTroopsListView()
Dim n As Integer
n = 3
LstTroops.View = lvwReport
LstTroops.Sorted = False
LstTroops.ListItems.Clear
LstTroops.ColumnHeaders.Clear
LstTroops.SortOrder = lvwAscending
LstTroops.FullRowSelect = True
LstTroops.AllowColumnReorder = False
LstTroops.LabelEdit = lvwManual
LstTroops.Checkboxes = False
LstTroops.GridLines = True
LstTroops.MultiSelect = False
LstTroops.HideSelection = False

LstTroops.ColumnHeaders.Add , , PublicMsgs(13), LstTroops.Width / n / 3.6
LstTroops.ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(14), LstTroops.Width / n * 1.5
LstTroops.ColumnHeaders.Add , , PublicEditors(1) & "ID", LstTroops.Width / n * 1.5

End Sub

Private Sub InitPropertiesListView()
Dim n As Integer
n = 3
With LstProperties
   .View = lvwReport
   .Sorted = False
   .ListItems.Clear
   .ColumnHeaders.Clear
   .SortOrder = lvwAscending
   .FullRowSelect = True
   .AllowColumnReorder = False
   .LabelEdit = lvwManual
   .Checkboxes = True
   .GridLines = True
   .MultiSelect = False
   .HideSelection = False

   .ColumnHeaders.Add , , PublicEditors(2) & PublicMsgs(13), .Width / n / 2
   .ColumnHeaders.Add , , PublicEditors(2) & PublicMsgs(14), .Width / n * 1.2
   .ColumnHeaders.Add , , PublicEditors(2) & "ID", .Width / n * 1.2

End With
End Sub

Private Sub InitItemBox()
Dim n As Integer
n = 3
With boxItem
   .View = lvwReport
   .Sorted = False
   .ListItems.Clear
   .ColumnHeaders.Clear
   .SortOrder = lvwAscending
   .FullRowSelect = True
   .AllowColumnReorder = False
   .LabelEdit = lvwManual
   .Checkboxes = True
   .GridLines = True
   .MultiSelect = False
   .HideSelection = False

   .ColumnHeaders.Add , , PublicEditors(2) & PublicMsgs(13), .Width / n / 2
   .ColumnHeaders.Add , , PublicEditors(2) & PublicMsgs(14), .Width / n * 1.2
   .ColumnHeaders.Add , , PublicEditors(2) & "ID", .Width / n * 1.2

End With
End Sub

Private Sub LoadItemBox()
Dim n As Long, oItem As ListItem

For n = 0 To N_Item - 1
  With itm(n)
     Set oItem = boxItem.ListItems.Add(, , .ID)
     oItem.SubItems(1) = .csvName
     oItem.SubItems(2) = .dbName
  End With
Next n

End Sub

'*************************************************************************
'**函 数 名：InitSkillsListView
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-23 22:10:03
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitSkillsListView()
Dim n As Integer, oItem As ListItem, strTemp As String
n = 2
LstSkills.View = lvwReport
LstSkills.Sorted = False
LstSkills.ListItems.Clear
LstSkills.ColumnHeaders.Clear
LstSkills.SortOrder = lvwAscending
LstSkills.FullRowSelect = True
LstSkills.AllowColumnReorder = False
LstSkills.LabelEdit = lvwAutomatic
LstSkills.Checkboxes = False
LstSkills.GridLines = True
LstSkills.MultiSelect = False
LstSkills.HideSelection = False
LstSkills.ColumnHeaders.Add , , PublicMsgs(26), LstSkills.Width * 0.5
LstSkills.ColumnHeaders.Add , , PublicMsgs(31), LstSkills.Width * 0.5


For n = 0 To UBound(PublicSkills)
      strTemp = IIf(PublicSkills(n) <> "", PublicSkills(n), PublicMsgs(42))
      Set oItem = LstSkills.ListItems.Add(, "skl_" & n, "")
      
      With oItem
         '.SubItems(1) = strTemp
         .ListSubItems.Add , , strTemp
         If strTemp = PublicMsgs(42) Then
          .ForeColor = vbRed
          .ListSubItems.Item(1).ForeColor = vbRed
         End If
          
      End With
Next n

End Sub

Private Sub InitFrames()
Dim i As Integer

For i = 0 To 4
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

If LCase$(MnBInfo.Language) = "en" Then
     txtPropBag(3).Enabled = False
     txtPropBag(4).Enabled = False
End If

End Sub





Private Sub LstProperties_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = vbKeyC Then
     frmInfo.CopyInfotoClipBoard
     frmTip.ShowTip PublicMsgs(130)
     Timer_KillTip.Enabled = True
End If
End Sub

Private Sub LstProperties_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not frmMain.mBanfrmInfo.Checked Then
   LstItemMPos.X = X
   LstItemMPos.Y = Y
   Set HitLV = LstProperties
   Timer_MousePos.Enabled = True
   
   If Timer_KillTip.Enabled = True Then
      frmTip.HideTip
      Timer_KillTip.Enabled = False
   End If
End If

End Sub

Private Sub LstSkills_AfterLabelEdit(Cancel As Integer, NewString As String)
If CustomActive Then
PutSkill CurrentTrp, CInt(Val(NewString)), LstSkills.SelectedItem.Index
End If
End Sub

Private Sub LstSkills_Click()
LstSkills.StartLabelEdit
End Sub


Private Sub LstTroops_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentTrpID = Val(Item.Text)

CurrentTrp = Trps(CurrentTrpID)
LoadTroopInfo
End Sub

Private Sub LstTroops_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = vbKeyC Then
     frmInfo.CopyInfotoClipBoard
     frmTip.ShowTip PublicMsgs(130)
     Timer_KillTip.Enabled = True
End If
End Sub

Private Sub LstTroops_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Not frmMain.mBanfrmInfo.Checked Then
'If Shift = 2 Then
   'LstItems.SetFocus
   LstItemMPos.X = X
   LstItemMPos.Y = Y
   Set HitLV = LstTroops
   Timer_MousePos_Trp.Enabled = True
   
   If Timer_KillTip.Enabled = True Then
      frmTip.HideTip
      Timer_KillTip.Enabled = False
   End If
'End If
End If

End Sub

Private Sub OptSex_Click(Index As Integer)
Dim NewFlags As String, tI As Integer64b

If CustomActive Then
Select Case Index
      Case 2
       NewFlags = CStr(cbSkins.ListIndex + Val(tf_undead))
       cbSkins.Enabled = True
      Case Else
       cbSkins.Enabled = False
       NewFlags = CStr(Index)
End Select
       With CurrentTrp
            tI = DeleteFlags64b(StrToI64(.Flags), HexStrToI64(troop_type_mask))
            tI = AddFlags64b(tI, StrToI64(NewFlags))

            .Flags = I64toStrNZ(tI)
       End With
End If
End Sub

Private Sub Tab1_Click()
Dim i As Integer

For i = 0 To 4
    With FraProps(i)
         .Visible = i + 1 = Tab1.SelectedItem.Index
    End With
Next i

LoadTroopInfo
End Sub

'*************************************************************************
'**函 数 名：LoadTroopInfo
'**输    入：-
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-23 13:04:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadTroopInfo()
On Error GoTo EL

Dim i As Integer, oItem As ListItem, strTem As String, lngTem As Long
CustomActive = False
With CurrentTrp
  Select Case Tab1.SelectedItem.Index
       Case 1
          'Recognize
          txtPropBag(0).Text = .strID
          txtPropBag(1).Text = .strName
          txtPropBag(2).Text = .strPtName
          txtPropBag(3).Text = .csvName
          txtPropBag(4).Text = .csvName_pl
          'Properties
          txtPropBag(5).Text = .tAttrib.level
          txtPropBag(6).Text = .tAttrib.strPoint
          txtPropBag(7).Text = .tAttrib.agiPoint
          txtPropBag(8).Text = .tAttrib.intPoint
          txtPropBag(9).Text = .tAttrib.chaPoint
          'WeaponProfession
          txtPropBag(10).Text = .WP.one_handed
          txtPropBag(11).Text = .WP.two_handed
          txtPropBag(12).Text = .WP.polearm
          txtPropBag(13).Text = .WP.archery
          txtPropBag(14).Text = .WP.crossbow
          txtPropBag(15).Text = .WP.throwing
          txtPropBag(16).Text = .WP.firearm
              For i = 1 To LstSkills.ListItems.Count
                    LstSkills.ListItems(i).Text = GetSkill(i)
              Next i
       Case 2
          '.Faction = IIf(CheckExist(EditInfo_FactionsCount, .Faction), .Faction, 0)
          .Faction = GetID(.Faction_strID, , Factions(0).strID)
          cbFaction.ListIndex = .Faction
          
          '.Upgrade1 = IIf(CheckExist(EditInfo_TroopsCount, .Upgrade1), .Upgrade1, 0)
          .Upgrade1 = GetID(.Upgrade1_strID, , Trps(0).strID)
          CbUpgrade(0).ListIndex = .Upgrade1
          
          '.Upgrade2 = IIf(CheckExist(EditInfo_TroopsCount, .Upgrade2), .Upgrade2, 0)
          .Upgrade2 = GetID(.Upgrade2_strID, , Trps(0).strID)
          CbUpgrade(1).ListIndex = .Upgrade2
          
          .SceneID = GetID(.Scene_strID, , Scenes(0).strID)
          If .SceneID = 0 Then .Entry = 0
          cbScenes.ListIndex = .SceneID
          
          txtEntry.Text = .Entry
       Case 3
          LoadFlagsList
          'Sex
          CheckSex .Flags
       Case 4
        LoadTroopPropertiesList
        
        Case 5
          txtFace(0).Text = "0x"
          txtFace(1).Text = "0x"
          For i = 1 To 4
           txtFace(0).Text = txtFace(0).Text & LCase$(I64toHexStr(StrToI64(.Face(i))))
           txtFace(1).Text = txtFace(1).Text & LCase$(I64toHexStr(StrToI64(.Face(i + 4))))
          Next i
  End Select
End With
CustomActive = True

Exit Sub
EL:
  Call logErr("frmTroops", "LoadTroopInfo", Err.Number, Err.Description)
End Sub

Private Sub CheckSex(ByVal Flags As String)
Dim i As Integer, tI As Integer64b

tI = And64b(StrToI64(Flags), HexStrToI64(troop_type_mask))

cbSkins.Enabled = False

For i = Val("&H" & troop_type_mask) To Val(tf_undead) Step -1
  If tI.by(0) = i Then
     OptSex(2).Value = True
     cbSkins.Enabled = True
     cbSkins.ListIndex = i - Val(tf_undead)
     Exit Sub
  End If
Next i

OptSex(0).Value = True

If tI.by(0) = Val(tf_female) Then
     OptSex(1).Value = True
End If

End Sub



Private Sub Timer_KillTip_Timer()
frmTip.HideTip
Timer_KillTip.Enabled = False
End Sub

Private Sub Timer_MousePos_Timer()
Dim ItemIndex As Long, ReturnValue As Long, MouseOut As Boolean
Dim MPos As POINTAPI
Static PosOld As POINTAPI
Dim LVRect As RECT
'Dim Hwnd As Long

If Not HitLV Is Nothing Then
ReturnValue = GetCursorPos(MPos)   '获取鼠标绝对位置
ReturnValue = GetWindowRect(HitLV.hWnd, LVRect)   '获取listview区域
MouseOut = MPos.X < LVRect.Left Or MPos.X > LVRect.Right - 300 / Screen.TwipsPerPixelX Or MPos.Y < LVRect.Top + 250 / Screen.TwipsPerPixelX Or MPos.Y > LVRect.Bottom - 300 / Screen.TwipsPerPixelY

If MouseOut Then
     UnLoad frmInfo
     PosOld.X = 0
     PosOld.Y = 0
     Timer_MousePos.Enabled = False
Else
  If LstItemMPos.X <> PosOld.X Or LstItemMPos.Y <> PosOld.Y Then
  '------------------------获得listview鼠标所指位置ItemIndex----------------------------------
   'Hwnd = GetActiveWindow()
   'If Hwnd = frmMain.Hwnd Or Hwnd = frmItems.Hwnd Or Hwnd = frmInfo.Hwnd Then
       'LstItems.SetFocus
   'End If
       
   ItemIndex = GetListViewItemIndexUnderMousePointer(HitLV, LstItemMPos.X, LstItemMPos.Y)
   If ItemIndex > 0 Then
      frmInfo.ShowfrmInfo MPos.X * Screen.TwipsPerPixelX + 300, MPos.Y * Screen.TwipsPerPixelY + 300
      frmInfo.LoadItemInfo Val(HitLV.ListItems(ItemIndex).Text)
      frmInfo.ZOrder
      HitLV.SetFocus
   Else
      UnLoad frmInfo
   End If
  '-------------------------------------------------------------------------------------------
   PosOld = LstItemMPos
  End If
End If
End If

End Sub

Private Sub Timer_MousePos_Trp_Timer()
Dim ItemIndex As Long, ReturnValue As Long, MouseOut As Boolean
Dim MPos As POINTAPI
Static PosOld As POINTAPI
Dim LVRect As RECT
'Dim Hwnd As Long

If Not HitLV Is Nothing Then
ReturnValue = GetCursorPos(MPos)   '获取鼠标绝对位置
ReturnValue = GetWindowRect(HitLV.hWnd, LVRect)   '获取listview区域
MouseOut = MPos.X < LVRect.Left Or MPos.X > LVRect.Right - 300 / Screen.TwipsPerPixelX Or MPos.Y < LVRect.Top + 250 / Screen.TwipsPerPixelX Or MPos.Y > LVRect.Bottom - 300 / Screen.TwipsPerPixelY

If MouseOut Then
     UnLoad frmInfo
     PosOld.X = 0
     PosOld.Y = 0
     Timer_MousePos_Trp.Enabled = False
Else
  If LstItemMPos.X <> PosOld.X Or LstItemMPos.Y <> PosOld.Y Then
  '------------------------获得listview鼠标所指位置ItemIndex----------------------------------
   ItemIndex = GetListViewItemIndexUnderMousePointer(HitLV, LstItemMPos.X, LstItemMPos.Y)
   If ItemIndex > 0 Then
      frmInfo.ShowfrmInfo MPos.X * Screen.TwipsPerPixelX + 300, MPos.Y * Screen.TwipsPerPixelY + 300
      frmInfo.LoadTroopInfo Val(HitLV.ListItems(ItemIndex).Text)
      frmInfo.ZOrder
      HitLV.SetFocus
   Else
      UnLoad frmInfo
   End If
  '-------------------------------------------------------------------------------------------
   PosOld = LstItemMPos
  End If
End If
End If

End Sub

Private Sub txtEntry_LostFocus()
If CustomActive Then
    CurrentTrp.Entry = Val(txtEntry.Text)
End If
End Sub


Private Sub txtFace_LostFocus(Index As Integer)
On Error GoTo EL
Dim TemHexStr As String, i As Integer, strUnit As String

If CustomActive Then
TemHexStr = FixHexStr_64(txtFace(Index).Text)
  For i = 1 To 4
     strUnit = Mid(TemHexStr, (i - 1) * 16 + 1, 16)
     CurrentTrp.Face(Index * 4 + i) = I64toStrNZ(HexStrToI64(strUnit))
  Next i
End If

Exit Sub
EL:
  Call logErr("frmTroops", "txtFace_LostFocus(" & Index & ")", Err.Number, Err.Description)
End Sub

Private Sub txtPropBag_LostFocus(Index As Integer)
On Error GoTo EL
With CurrentTrp

If CustomActive Then

If Index < 3 Then
   txtPropBag(Index).Text = Replace(txtPropBag(Index).Text, " ", "_")
End If

Select Case Index
      Case 0
           'Check Value
            If Left$(txtPropBag(Index).Text, 4) <> "trp_" Then
                txtPropBag(Index).Text = "trp_" & txtPropBag(Index).Text
            End If
         .strID = txtPropBag(Index).Text
      Case 1
         .strName = txtPropBag(Index).Text
           If LCase$(MnBInfo.Language) = "en" Then
             .csvName = .strName
           End If
      Case 2
         .strPtName = txtPropBag(Index).Text
           If LCase$(MnBInfo.Language) = "en" Then
             .csvName_pl = .strPtName
           End If
      Case 3
           If LCase$(MnBInfo.Language) <> "en" Then
             .csvName = txtPropBag(Index).Text
           End If
      Case 4
           If LCase$(MnBInfo.Language) <> "en" Then
             .csvName_pl = txtPropBag(Index).Text
           End If
      Case 5
         .tAttrib.level = CInt(Val(txtPropBag(Index).Text))
      Case 6
         .tAttrib.strPoint = CInt(Val(txtPropBag(Index).Text))
      Case 7
         .tAttrib.agiPoint = CInt(Val(txtPropBag(Index).Text))
      Case 8
         .tAttrib.intPoint = CInt(Val(txtPropBag(Index).Text))
      Case 9
         .tAttrib.chaPoint = CInt(Val(txtPropBag(Index).Text))
      Case 10
         .WP.one_handed = CStr(Abs(Int(Val(txtPropBag(Index).Text))))
      Case 11
         .WP.two_handed = CStr(Abs(Int(Val(txtPropBag(Index).Text))))
      Case 12
         .WP.polearm = CStr(Abs(Int(Val(txtPropBag(Index).Text))))
      Case 13
         .WP.archery = CStr(Abs(Int(Val(txtPropBag(Index).Text))))
      Case 14
         .WP.crossbow = CStr(Abs(Int(Val(txtPropBag(Index).Text))))
      Case 15
         .WP.throwing = CStr(Abs(Int(Val(txtPropBag(Index).Text))))
      Case 16
         .WP.firearm = CStr(Abs(Int(Val(txtPropBag(Index).Text))))
End Select

End If
End With

Exit Sub
EL:
  Call logErr("frmTroops", "txtPropBag_LostFocus(" & Index & ")", Err.Number, Err.Description)
End Sub

'*************************************************************************
'**函 数 名：InitTroopUpgradeCombo
'**输    入：
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-25 16:27:42
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitTroopUpgradeCombo()
Dim i As Integer, j As Long

For i = 0 To 1
   CbUpgrade(i).Clear
   
   For j = 0 To N_Troop - 1
       If j <> 0 Then
         CbUpgrade(i).AddItem "(" & j & ")" & Trps(j).csvName & "|" & Trps(j).strID
       Else
         CbUpgrade(i).AddItem PublicTips(0)
       End If
   Next j
Next i
   
End Sub

'*************************************************************************
'**函 数 名：LoadTroopPropertiesList
'**输    入：-
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-26 17:02:05
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadTroopPropertiesList()
Dim i As Integer, oItem As ListItem, q As Boolean
With CurrentTrp
LstProperties.ListItems.Clear
         For i = 1 To 64
            If .lstInventory(i).strX <> "" Then
               .lstInventory(i).X = GetID(.lstInventory(i).strX, True, "", -1)
               If .lstInventory(i).X = -1 Then q = True
            Else
               .lstInventory(i).X = -1
               .lstInventory(i).Y = 0
            End If
         Next i
         
         If q Then StructureTroopInventory -1
         
         For i = 1 To 64
         
            If .lstInventory(i).X > -1 Then
               Set oItem = LstProperties.ListItems.Add(, , itm(.lstInventory(i).X).ID)
               
               With oItem
                     .SubItems(1) = itm(CurrentTrp.lstInventory(i).X).csvName
                     .SubItems(2) = itm(CurrentTrp.lstInventory(i).X).dbName
               End With
            Else

               Set oItem = LstProperties.ListItems.Add(, , PublicTips(0))
            End If
         Next i
       Set LstProperties.SelectedItem = LstProperties.ListItems(1)
       
End With
End Sub


'*************************************************************************
'**函 数 名：InitFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-26 22:56:56
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitFlagsList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("", "", "", "NPC英雄", "无生命的", "只能击晕", "倒下必死", "无法活捉", "骑马", "商队", _
               "随机相貌", "保证穿鞋子", "保证穿盔甲", "保证戴头盔", "保证戴手套", "保证有马", "保证有盾", "保证有远程武器", _
               "不能作为驻兵")
TemArray = Array(tf_male, tf_female, tf_undead, tf_hero, tf_inactive, tf_unkillable, tf_allways_fall_dead, tf_no_capture_alive, tf_mounted, tf_is_merchant, _
               tf_randomize_face, tf_guarantee_boots, tf_guarantee_armor, tf_guarantee_helmet, tf_guarantee_gloves, tf_guarantee_horse, tf_guarantee_shield, tf_guarantee_ranged, _
               tf_unmoveable_in_party_window)
         For i = 3 To UBound(strTemArray)
             chkFlags(i - 3).Caption = strTemArray(i)
             chkFlags(i - 3).Tag = TemArray(i)
         Next i

End Sub

'*************************************************************************
'**函 数 名：LoadFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-26 23:19:10
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFlagsList()
Dim i As Integer, tI(2) As Integer64b, k As String

With CurrentTrp

tI(0) = StrToI64(.Flags)

For i = 0 To chkFlags.UBound
    tI(2) = HexStrToI64(chkFlags(i).Tag)
    tI(1) = And64b(tI(0), tI(2))
       If IsEqual64b(tI(1), tI(2)) Then
          chkFlags(i).Value = 1
       Else
          chkFlags(i).Value = 0
       End If
Next i

End With
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
    Set oItem = FindItem(oLV, Start, "0|1|2", QueryString, True, vbTextCompare, bReverse)
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

Private Sub InitSkinsCombo()
Dim i As Integer
cbSkins.Clear

For i = Val(tf_undead) To Val("&H" & troop_type_mask)
cbSkins.AddItem "Skin" & i
Next i

cbSkins.ListIndex = 0
End Sub

'*************************************************************************
'**函 数 名：LoadScenesCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-29 10:56:35
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadScenesCombo()
Dim i As Long
cbScenes.Clear
         For i = 0 To N_Scene - 1
                  cbScenes.AddItem "(" & i & ")" & Scenes(i).strName & "|" & Scenes(i).strID
         Next i
         
End Sub


'*************************************************************************
'**函 数 名：AddTroopInventory
'**输    入：(Long)Itm_Idx
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-01-19 20:42:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function AddTroopInventory(ByVal Itm_Idx As Long) As Boolean
Dim i As Integer

With CurrentTrp

   For i = 1 To 64
     If .lstInventory(i).X = -1 Then
         .lstInventory(i).X = Itm_Idx
         .lstInventory(i).strX = itm(Itm_Idx).dbName
         .lstInventory(i).Y = 0
            AddTroopInventory = True
            Exit For
     End If
   Next i

End With
End Function


Public Sub ReLoadInfo()
LoadTroopInfo
End Sub
