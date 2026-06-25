VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScenes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "场景编辑器"
   ClientHeight    =   9375
   ClientLeft      =   1755
   ClientTop       =   180
   ClientWidth     =   14355
   Icon            =   "frmScenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_6"
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
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
      Height          =   8055
      Index           =   0
      Left            =   6240
      TabIndex        =   13
      Top             =   480
      Width           =   7815
      Begin VB.Frame Frame1 
         Caption         =   "外景"
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
         TabIndex        =   51
         Top             =   5160
         Width           =   7335
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   4
            Left            =   1800
            TabIndex        =   52
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型:"
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
            Index           =   4
            Left            =   1200
            TabIndex        =   53
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "模型"
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
         Height          =   3015
         Index           =   9
         Left            =   240
         TabIndex        =   37
         Top             =   1920
         Width           =   7335
         Begin VB.Frame Frame1 
            Caption         =   "表面"
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
            Height          =   1575
            Index           =   4
            Left            =   240
            TabIndex        =   54
            Top             =   1200
            Width           =   6855
            Begin VB.TextBox txtLength 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   1485
               TabIndex        =   58
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox txtTerrainCode 
               Height          =   375
               Left            =   1485
               TabIndex        =   55
               Top             =   360
               Width           =   5055
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "地形边长:"
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
               Left            =   480
               TabIndex        =   57
               Top             =   960
               Width           =   885
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "表面代码:"
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
               Left            =   480
               TabIndex        =   56
               Top             =   435
               Width           =   885
            End
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   39
            Top             =   240
            Width           =   5055
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   3
            Left            =   1800
            TabIndex        =   38
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "碰撞模型名:"
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
            Left            =   630
            TabIndex        =   41
            Top             =   795
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "网格名:"
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
            Left            =   975
            TabIndex        =   40
            Top             =   315
            Width           =   690
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
         Height          =   1335
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   7335
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Width           =   5055
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   15
            Top             =   720
            Width           =   5055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "场景ID:"
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
            Left            =   960
            TabIndex        =   18
            Top             =   315
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "场景名(EN):"
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
            Left            =   600
            TabIndex        =   17
            Top             =   795
            Width           =   1110
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps1"
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
      Height          =   3255
      Index           =   1
      Left            =   10440
      TabIndex        =   20
      Top             =   4080
      Width           =   3135
      Begin VB.Frame Frame1 
         Caption         =   "箱子"
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
         Height          =   4815
         Index           =   2
         Left            =   240
         TabIndex        =   42
         Top             =   3120
         Width           =   7335
         Begin VB.ComboBox CbSelectTroop 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   360
            Width           =   6975
         End
         Begin MSComctlLib.ListView LstTroops 
            Height          =   3735
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   6588
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
         Caption         =   "基本设置"
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
         Index           =   6
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   7335
         Begin VB.Frame Frame1 
            Caption         =   "左上角"
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
            Height          =   1455
            Index           =   8
            Left            =   480
            TabIndex        =   28
            Top             =   360
            Width           =   2775
            Begin VB.TextBox txtPos 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   1140
               TabIndex        =   30
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtPos 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1140
               TabIndex        =   29
               Top             =   840
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "纵坐标Y:"
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
               Index           =   14
               Left            =   210
               TabIndex        =   32
               Top             =   915
               Width           =   795
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "横坐标X:"
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
               Left            =   210
               TabIndex        =   31
               Top             =   435
               Width           =   795
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "右下角"
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
            Height          =   1455
            Index           =   7
            Left            =   3840
            TabIndex        =   23
            Top             =   360
            Width           =   2775
            Begin VB.TextBox txtPos 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   1140
               TabIndex        =   25
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox txtPos 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   1140
               TabIndex        =   24
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "横坐标X:"
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
               Left            =   210
               TabIndex        =   27
               Top             =   435
               Width           =   795
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "纵坐标Y:"
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
               Left            =   210
               TabIndex        =   26
               Top             =   915
               Width           =   795
            End
         End
         Begin VB.TextBox txtWaterLevel 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1605
            TabIndex        =   22
            Top             =   2160
            Width           =   5055
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "水面高度:"
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
            Left            =   600
            TabIndex        =   33
            Top             =   2280
            Width           =   885
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps2"
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
      Height          =   5295
      Index           =   2
      Left            =   7320
      TabIndex        =   34
      Top             =   1200
      Width           =   5295
      Begin VB.Frame Frame1 
         Caption         =   "特性"
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
         Height          =   3255
         Index           =   1
         Left            =   480
         TabIndex        =   35
         Top             =   360
         Width           =   6975
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
            Left            =   600
            TabIndex        =   50
            Top             =   2640
            Width           =   4095
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
            Left            =   600
            TabIndex        =   49
            Top             =   2280
            Width           =   3975
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
            Left            =   600
            TabIndex        =   48
            Top             =   1920
            Width           =   4695
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
            Left            =   600
            TabIndex        =   47
            Top             =   1560
            Width           =   4095
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
            Left            =   600
            TabIndex        =   46
            Top             =   1200
            Width           =   4335
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
            Left            =   600
            TabIndex        =   45
            Top             =   480
            Width           =   4095
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
            Left            =   600
            TabIndex        =   36
            Top             =   840
            Width           =   4335
         End
      End
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   1920
      Top             =   1440
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
            Picture         =   "frmScenes.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScenes.frx":0A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScenes.frx":0B7E
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   5
      Top             =   8805
      Width           =   2175
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
      TabIndex        =   4
      Top             =   8805
      Width           =   2175
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   4575
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前场景(&O)"
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
      TabIndex        =   0
      Top             =   8760
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstScenes 
      Height          =   8415
      Left            =   120
      TabIndex        =   6
      Top             =   525
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   14843
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmScenes.frx":0CD8
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   8535
      Left            =   6120
      TabIndex        =   19
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   15055
      MultiRow        =   -1  'True
      ImageList       =   "IL1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本属性(&P)"
            Key             =   "PropBag"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "场景设定(&S)"
            Key             =   "SceneSets"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "特性(&E)"
            Key             =   "Special"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
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
      MouseIcon       =   "frmScenes.frx":5345
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9060
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
      ForeColor       =   &H00004000&
      Height          =   180
      Index           =   1
      Left            =   4440
      MouseIcon       =   "frmScenes.frx":564F
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9060
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "场景数:"
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
      Left            =   135
      TabIndex        =   10
      Top             =   9060
      Width           =   690
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
      TabIndex        =   9
      Top             =   195
      Width           =   495
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
      MouseIcon       =   "frmScenes.frx":5959
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   9060
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
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   3600
      MouseIcon       =   "frmScenes.frx":5C63
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   9060
      Width           =   705
   End
End
Attribute VB_Name = "frmScenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CustomActive As Boolean

Private Sub CApply_Click()
Dim q As Boolean

If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then

    If UCase(Scenes(CurrentSceneID).strID) <> UCase(CurrentScene.strID) Then               '外引
        q = ChangeStrID(Scenes(CurrentSceneID).strID, CurrentScene.strID)
        
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentScene.strID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    
    Scenes(CurrentSceneID) = CurrentScene
    LstScenes.ListItems(CurrentSceneID + 1).SubItems(1) = Scenes(CurrentSceneID).strName
    LstScenes.ListItems(CurrentSceneID + 1).SubItems(2) = Scenes(CurrentSceneID).strID
    
    CurrentScene = Scenes(CurrentSceneID)
    LoadSceneInfo
End If
End Sub

Private Sub CbSelectTroop_Click()
Dim i As Long

If CustomActive Then
i = LstTroops.SelectedItem.Index
  If LstTroops.SelectedItem.Index < LstTroops.ListItems.Count Then
    With CurrentScene
         .Chests(i).ID = CbSelectTroop.ListIndex - 1
         .Chests(i).strID = Trps(.Chests(i).ID).strID
    
         If .Chests(i).ID = -1 Then
            DeleteCurrentChest LstTroops.SelectedItem.Index
            LoadTroopsListView
         Else
             LstTroops.SelectedItem.Text = .Chests(i).ID
             LstTroops.SelectedItem.SubItems(1) = Trps(.Chests(i).ID).csvName
             LstTroops.SelectedItem.SubItems(2) = Trps(.Chests(i).ID).strID
             
         End If
    End With
  Else
    CreateCurrentChest CbSelectTroop.ListIndex - 1
    If CbSelectTroop.ListIndex - 1 >= 0 Then
       LoadTroopsListView
    End If
  End If
  
End If
End Sub

'*************************************************************************
'**函 数 名：DeleteCurrentChest
'**输    入：(Long)ChestIndex
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-08 22:17:57
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub DeleteCurrentChest(ByVal ChestIndex As Long)
Dim i As Long
With CurrentScene
     If .ChestCount > 0 Then
          .ChestCount = .ChestCount - 1
          If .ChestCount = 0 Then
             'ReDim .Chests(0)
             .Chests(.ChestCount + 1).ID = 0
             .Chests(.ChestCount + 1).strID = ""
          Else
              For i = ChestIndex To .ChestCount
                   .Chests(i) = .Chests(i + 1)
              Next i
              
              ReDim Preserve .Chests(1 To .ChestCount)
                '.Chests(.ChestCount + 1).ID = 0
                '.Chests(.ChestCount + 1).strID = ""
          End If
          
     End If
End With
End Sub

'*************************************************************************
'**函 数 名：CreateCurrentChest
'**输    入：(Long)Trp_ID
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-08 22:35:46
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub CreateCurrentChest(ByVal Trp_ID As Long)
Dim i As Long
If Trp_ID > -1 Then
With CurrentScene
     If .ChestCount > 0 Then
          .ChestCount = .ChestCount + 1
          ReDim Preserve .Chests(1 To .ChestCount)
     Else
          .ChestCount = 1
          ReDim .Chests(1 To .ChestCount)
     End If
     
     .Chests(.ChestCount).ID = Trp_ID
     .Chests(.ChestCount).strID = Trps(Trp_ID).strID
End With
End If
End Sub

Private Sub CbSelectTroop_Scroll()
Call CbSelectTroop_Click
End Sub

Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If Scenes(CurrentSceneID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), Scenes(CurrentSceneID).strID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
              
              DelIndex Scenes(CurrentSceneID).strID
              If CurrentSceneID < N_Scene - 1 Then
                For i = CurrentSceneID To N_Scene - 2 Step 1
                
                    ChangeID Scenes(i + 1).strID, Scenes(i + 1).ID - 1
                    j = Scenes(i).ID
                    Scenes(i) = Scenes(i + 1)
                    Scenes(i).ID = j
                    LstScenes.ListItems(i + 1).SubItems(1) = LstScenes.ListItems(i + 2).SubItems(1)
                    LstScenes.ListItems(i + 1).SubItems(2) = LstScenes.ListItems(i + 2).SubItems(2)
                Next i
                
                ReDim Preserve Scenes(N_Scene - 2)
                LstScenes.ListItems.Remove N_Scene
                N_Scene = N_Scene - 1
                
              Else
                ReDim Preserve Scenes(N_Scene - 2)
                LstScenes.ListItems.Remove N_Scene
                
                N_Scene = N_Scene - 1
                CurrentSceneID = N_Scene - 1
                
              End If
               
               LstScenes_ItemClick LstScenes.ListItems(CurrentSceneID + 1)
               LstScenes.ListItems(CurrentSceneID + 1).Selected = True
               LstScenes.ListItems(CurrentSceneID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), Scenes(CurrentSceneID).strID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), Scenes(CurrentSceneID).strID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_Scene, Scenes(CurrentSceneID).strID & "_New") Then
           ReDim Preserve Scenes(N_Scene)
           N_Scene = N_Scene + 1
           Scenes(N_Scene - 1) = Scenes(CurrentSceneID)
           With Scenes(N_Scene - 1)
                 .ID = N_Scene - 1
                 .strID = .strID & "_New"
                 .strName = .strName & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstScenes.ListItems.Add(, "scn_" & Scenes(N_Scene - 1).ID, Scenes(N_Scene - 1).ID)
      
                 With oItem
                    .SubItems(1) = Scenes(N_Scene - 1).strName
                    .SubItems(2) = Scenes(N_Scene - 1).strID
                 End With
           LstScenes_ItemClick LstScenes.ListItems(N_Scene)
           LstScenes.ListItems(N_Scene).Selected = True
           LstScenes.ListItems(N_Scene).EnsureVisible
           
           Else
             MsgBox ActiveString(PublicMsgs(90), Scenes(CurrentSceneID).strID & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
         
      Case 2
         If CurrentSceneID > 0 Then
           If Scenes(CurrentSceneID - 1).Edit And Scenes(CurrentSceneID).Edit Then
           
                SwapID Scenes(CurrentSceneID - 1).strID, Scenes(CurrentSceneID).strID
                SwapScenes CurrentSceneID - 1, CurrentSceneID
                SwapListItem LstScenes.ListItems(CurrentSceneID), LstScenes.ListItems(CurrentSceneID + 1), 2, True
                
               LstScenes_ItemClick LstScenes.ListItems(CurrentSceneID)
               LstScenes.ListItems(CurrentSceneID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), Scenes(CurrentSceneID - 1).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
        
         End If
      Case 3
        If CurrentSceneID + 1 <= N_Scene - 1 Then
           If Scenes(CurrentSceneID).Edit And Scenes(CurrentSceneID + 1).Edit Then
           
                SwapID Scenes(CurrentSceneID).strID, Scenes(CurrentSceneID + 1).strID
                SwapScenes CurrentSceneID, CurrentSceneID + 1
                SwapListItem LstScenes.ListItems(CurrentSceneID + 1), LstScenes.ListItems(CurrentSceneID + 2), 2, True
                
                LstScenes_ItemClick LstScenes.ListItems(CurrentSceneID + 2)
                LstScenes.ListItems(CurrentSceneID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), Scenes(CurrentSceneID).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
        End If
End Select
End Sub


Private Sub chkFlags_Click(Index As Integer)
Dim a As Integer64b, TemFlags As Long, tI(2) As Integer64b

If CustomActive Then

With CurrentScene
    tI(1) = StrToI64(.Flags)
    tI(2) = HexStrToI64(chkFlags(Index).Tag)
    If chkFlags(Index).Value = 0 Then
      tI(0) = DeleteFlags64b(tI(1), tI(2))
    Else
      tI(0) = AddFlags64b(tI(1), tI(2))
    End If
    .Flags = I64toStrNZ(tI(0))
End With

End If
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstScenes, LstScenes.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstScenes_ItemClick(LstScenes.SelectedItem)
End If
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

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
LstScenes_ItemClick LstScenes.ListItems(CurrentSceneID + 1)
LstScenes.ListItems(CurrentSceneID + 1).Selected = True
LstScenes.ListItems(CurrentSceneID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
CustomActive = False

InitFrames
InitTroopsListView
InitScenesListView
LoadScenesList
InitFlagsList

LoadSelectTroopCombo

CurrentSceneID = 0
CurrentScene = Scenes(CurrentSceneID)
LoadSceneInfo

TranslateForm Me
InitCMDs

Label2.Caption = Label2.Caption & N_Scene

CustomActive = True
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.ubound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

Private Sub InitScenesListView()
Dim n As Integer
n = 3
LstScenes.View = lvwReport
LstScenes.Sorted = False
LstScenes.ListItems.Clear
LstScenes.ColumnHeaders.Clear
LstScenes.SortOrder = lvwAscending
LstScenes.FullRowSelect = True
LstScenes.AllowColumnReorder = False
LstScenes.LabelEdit = lvwManual
LstScenes.Checkboxes = False
LstScenes.GridLines = True
LstScenes.MultiSelect = False
LstScenes.HideSelection = False

LstScenes.ColumnHeaders.Add , , PublicMsgs(13), LstScenes.Width / n / 3.6
LstScenes.ColumnHeaders.Add , , PublicEditors(6) & PublicMsgs(14), LstScenes.Width / n * 1.5
LstScenes.ColumnHeaders.Add , , PublicEditors(6) & "ID", LstScenes.Width / n * 1.5


End Sub


Private Sub InitTroopsListView()
Dim n As Integer
n = 3
With LstTroops
   .View = lvwReport
   .Sorted = False
   .ListItems.Clear
   .ColumnHeaders.Clear
   .SortOrder = lvwAscending
   .FullRowSelect = True
   .AllowColumnReorder = False
   .LabelEdit = lvwManual
   .Checkboxes = False
   .GridLines = True
   .MultiSelect = False
   .HideSelection = False

   .ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(13), .Width / n / 3
   .ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(14), .Width / n * 2
   .ColumnHeaders.Add , , PublicEditors(1) & "ID", .Width / n * 2

End With

For n = 1 To 6
LstTroops.ListItems.Add , , PublicTips(0)
Next n

End Sub

Private Sub LoadTroopsListView()
Dim i As Integer, oItem As ListItem, q As Boolean

CustomActive = False
LstTroops.ListItems.Clear
With CurrentScene

   If .ChestCount > 0 Then
     For i = 1 To .ChestCount
        '.Chests(i).ID = IIf(CheckExist(EditInfo_TroopsCount, .Chests(i)), .Chests(i), 0)
        .Chests(i).ID = GetID(.Chests(i).strID, True, "", -1)
        If .Chests(i).ID > -1 Then
           Set oItem = LstTroops.ListItems.Add(, , .Chests(i).ID)
                oItem.SubItems(1) = Trps(.Chests(i).ID).csvName
                oItem.SubItems(2) = Trps(.Chests(i).ID).strID
        Else
            q = True
        End If
     Next i
     LstTroops.ListItems.Add , , PublicTips(0)

     LstTroops.ListItems(1).Selected = True
     LstTroops_ItemClick LstTroops.ListItems(1)
   Else
     LstTroops.ListItems.Add , , PublicTips(0)
     CbSelectTroop.ListIndex = 0
   End If
   
   If q Then
      StructureSceneChests -1
   End If
End With

CustomActive = True
End Sub



'*************************************************************************
'**函 数 名：LoadScenesList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-02 15:02:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadScenesList()
Dim n As Long, oItem As ListItem

LstScenes.ListItems.Clear
For n = 0 To N_Scene - 1

      Set oItem = LstScenes.ListItems.Add(, "scn_" & Scenes(n).ID, Scenes(n).ID)
      
      With oItem
         .SubItems(1) = Scenes(n).strName
         .SubItems(2) = Scenes(n).strID
      End With

Next n

End Sub


'*************************************************************************
'**函 数 名：LoadSceneInfo
'**输    入：-
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-06 10:58:11
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadSceneInfo()
Dim i As Integer, oItem As ListItem, strTem As String, lngTem As Long, tTI As Type_TerrainInfo
CustomActive = False
With CurrentScene
  Select Case Tab1.SelectedItem.Index
       Case 1
          'Recognize
          txtPropBag(0).Text = .strID
          txtPropBag(1).Text = .strName
          
          txtPropBag(2).Text = .MeshName
          txtPropBag(3).Text = .BodyName
          txtPropBag(4).Text = .Outer_Terrain_Type
          
          txtTerrainCode.Text = .TerrainCode
          tTI = PurseTerrainCode(.TerrainCode)
          txtLength = tTI.Length
       Case 2
          txtPos(0).Text = .p(0).X
          txtPos(1).Text = .p(0).Y
          txtPos(2).Text = .p(1).X
          txtPos(3).Text = .p(1).Y
          txtWaterLevel.Text = .WaterLevel
          
          LoadTroopsListView

       Case 3
          LoadFlagsList

  End Select
End With
CustomActive = True
End Sub

'*************************************************************************
'**函 数 名：InitFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-03 22:22:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitFlagsList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("室内场景", "强制使用天空贴图(无视是否室内场景)", "通过地形生成器产生地形", "随机产生地形生成器的密钥", "随机产生入口", "禁止马进入", "水浑浊")
TemArray = Array(sf_indoors, sf_force_skybox, sf_generate, sf_randomize, sf_auto_entry_points, sf_no_horses, sf_muddy_water)
         For i = 0 To UBound(strTemArray)
             chkFlags(i).Caption = strTemArray(i)
             chkFlags(i).Tag = TemArray(i)
         Next i

End Sub







Private Sub LstScenes_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentSceneID = Val(Item.Text)

CurrentScene = Scenes(CurrentSceneID)
LoadSceneInfo
End Sub


Private Sub LstTroops_ItemClick(ByVal Item As MSComctlLib.ListItem)
With CurrentScene
 If Item.Index < LstTroops.ListItems.Count Then
   CbSelectTroop.ListIndex = .Chests(Item.Index).ID + 1
 Else
   CbSelectTroop.ListIndex = 0
 End If
End With
End Sub

Private Sub Tab1_Click()
Dim i As Integer

For i = 0 To FraProps.ubound
    With FraProps(i)
         .Visible = i + 1 = Tab1.SelectedItem.Index
    End With
Next i

LoadSceneInfo
End Sub




Private Sub txtLength_LostFocus()
Dim tTI As Type_TerrainInfo
If CustomActive Then
     With tTI
          .Code = txtTerrainCode.Text
          .Length = Val(txtLength.Text)
     End With
     
     txtTerrainCode.Text = CalcTerrainCode(tTI)
     
     Call txtTerrainCode_LostFocus
End If
End Sub

Private Sub txtPos_Change(Index As Integer)
If CustomActive Then
 With CurrentScene
   Select Case Index
         Case 0
            .p(0).X = txtPos(Index).Text
         Case 1
            .p(0).Y = txtPos(Index).Text
         Case 2
            .p(1).X = txtPos(Index).Text
         Case 3
            .p(1).Y = txtPos(Index).Text
   End Select
 End With
End If

End Sub

Private Sub txtPropBag_LostFocus(Index As Integer)
If CustomActive Then
With CurrentScene
If Index < 4 Then
   txtPropBag(Index).Text = Replace(txtPropBag(Index).Text, " ", "_")
End If
Select Case Index
      Case 0
          'Check Value
            If Left$(txtPropBag(Index).Text, 4) <> "scn_" Then
               txtPropBag(Index).Text = "scn_" & txtPropBag(Index).Text
            End If
         .strID = txtPropBag(Index).Text
      Case 1
         .strName = txtPropBag(Index).Text
      Case 2
         .MeshName = txtPropBag(Index).Text
      Case 3
         .BodyName = txtPropBag(Index).Text
      Case 4
         .Outer_Terrain_Type = txtPropBag(Index).Text
End Select

End With
End If
End Sub



'*************************************************************************
'**函 数 名：LoadFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-09 22:58:53
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFlagsList()
Dim i As Integer, fI64 As Integer64b, fI64_NOW As Integer64b, resI64 As Integer64b

With CurrentScene

fI64_NOW = StrToI64(.Flags)
For i = 0 To chkFlags.ubound
               fI64 = HexStrToI64(chkFlags(i).Tag)
               resI64 = And64b(fI64_NOW, fI64)
               
               If IsEqual64b(resI64, fI64) Then
                  chkFlags(i).Value = 1
               Else
                  chkFlags(i).Value = 0
               End If
Next i

End With
End Sub

Private Sub InitFrames()
Dim i As Integer

For i = 0 To FraProps.ubound
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

Private Sub txtTerrainCode_LostFocus()
Dim TemHexStr As String, i As Integer, HexStr_40 As String, tTI As Type_TerrainInfo

If CustomActive Then
HexStr_40 = "0000000000000000000000000000000000000000"
TemHexStr = txtTerrainCode.Text

If Len(TemHexStr) >= 2 Then
    If Left(TemHexStr, 2) = "0x" Then
         TemHexStr = Right(TemHexStr, Len(TemHexStr) - 2)
    End If
End If
TemHexStr = Right(HexStr_40 & TemHexStr, 40)
TemHexStr = "0x" & TemHexStr

     CurrentScene.TerrainCode = TemHexStr
     
tTI = PurseTerrainCode(CurrentScene.TerrainCode)
          txtLength = tTI.Length
End If
End Sub

Private Sub LoadScenesListView()
Dim i As Integer, oItem As ListItem

CustomActive = False
LstTroops.ListItems.Clear
With CurrentScene

   If .ChestCount > 0 Then
     For i = 1 To .ChestCount
           Set oItem = LstTroops.ListItems.Add(, , .Chests(i).ID)
                oItem.SubItems(1) = Trps(.Chests(i).ID).csvName
                oItem.SubItems(2) = Trps(.Chests(i).ID).strID

     Next i
     LstTroops.ListItems.Add , , PublicTips(0)

     LstTroops.ListItems(1).Selected = True
     LstTroops_ItemClick LstTroops.ListItems(1)
   Else
     LstTroops.ListItems.Add , , PublicTips(0)
     CbSelectTroop.ListIndex = 0

   End If
End With

CustomActive = True
End Sub

'*************************************************************************
'**函 数 名：LoadSelectTroopCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-06 11:36:35
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadSelectTroopCombo()
Dim i As Long
CbSelectTroop.Clear
CbSelectTroop.AddItem PublicTips(0)
CbSelectTroop.ItemData(0) = -1
         For i = 0 To N_Troop - 1
                  CbSelectTroop.AddItem "(" & i & ")" & Trps(i).csvName & "|" & Trps(i).strID
                  CbSelectTroop.ItemData(i + 1) = i
         Next i
         

End Sub

Private Sub txtWaterLevel_Change()
If CustomActive Then
 With CurrentScene
     .WaterLevel = Val(txtWaterLevel.Text)
 End With
End If
End Sub

Public Sub ReLoadInfo()
LoadSceneInfo
End Sub

'*************************************************************************
'**函 数 名：PurseTerrainCode
'**输    入：(String)TerrainCode
'**输    出：(Type_TerrainInfo)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-03-26 11:19:24
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Function PurseTerrainCode(ByVal TerrainCode As String) As Type_TerrainInfo
Dim i As Integer, strPart(5) As String, tI64(0) As Integer64b, strTem As String

If Left(TerrainCode, 2) = "0x" Then
   TerrainCode = Right(TerrainCode, Len(TerrainCode) - 2)
End If

TerrainCode = Right("000000000000000000000000000000000000000000000000" & TerrainCode, 48)
For i = 0 To 5
    strPart(i) = Mid(TerrainCode, i * 8 + 1, 8)
Next i

tI64(0) = HexStrToI64(strPart(2))
tI64(0) = Div64b(tI64(0), 25)
tI64(0) = Div64b(tI64(0), 41)
strTem = I64toStr(tI64(0))

With PurseTerrainCode
     .Code = TerrainCode
     .Length = Val(strTem)
End With
End Function

'*************************************************************************
'**函 数 名：CalcTerrainCode
'**输    入：(Type_TerrainInfo)TerrainInfo
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-03-26 11:19:24
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Function CalcTerrainCode(TerrainInfo As Type_TerrainInfo) As String
Dim i As Integer, strPart(5) As String, tI64(0) As Integer64b, strTem As String

With TerrainInfo
     
If Left(.Code, 2) = "0x" Then
   .Code = Right(.Code, Len(.Code) - 2)
End If

.Code = Right("000000000000000000000000000000000000000000000000" & .Code, 48)
For i = 0 To 5
    strPart(i) = Mid(.Code, i * 8 + 1, 8)
Next i

tI64(0) = StrToI64(CStr(.Length))
tI64(0) = Multi64b8b(tI64(0), 25)
tI64(0) = Multi64b8b(tI64(0), 41)
strPart(2) = I64toHexStr(tI64(0))
strPart(2) = Right("00000000" & strPart(2), 8)

For i = 0 To 5
    CalcTerrainCode = CalcTerrainCode & strPart(i)
Next i

CalcTerrainCode = "0x" & LCase(CalcTerrainCode)
End With
End Function
