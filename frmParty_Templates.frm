VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParty_Templates 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "部队模板编辑器"
   ClientHeight    =   9375
   ClientLeft      =   4200
   ClientTop       =   1050
   ClientWidth     =   14355
   ForeColor       =   &H00C00000&
   Icon            =   "frmParty_Templates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_4"
   Begin VB.Frame FraProps 
      Caption         =   "FraProps3"
      Height          =   7815
      Index           =   3
      Left            =   6240
      TabIndex        =   6
      Top             =   600
      Width           =   7095
      Begin VB.Frame Frame1 
         Caption         =   "成员"
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
         Height          =   7575
         Index           =   1
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   7575
         Begin MSComctlLib.ListView LstTroops 
            Height          =   3375
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   7335
            _ExtentX        =   12938
            _ExtentY        =   5953
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.CheckBox chkSlave 
            Caption         =   "俘虏"
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
            Left            =   6240
            TabIndex        =   57
            Top             =   7100
            Width           =   855
         End
         Begin VB.TextBox txtMin 
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
            Left            =   825
            TabIndex        =   54
            Top             =   7080
            Width           =   1335
         End
         Begin VB.TextBox txtMax 
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
            Left            =   2880
            TabIndex        =   53
            Top             =   7080
            Width           =   1335
         End
         Begin VB.ComboBox CbSelectTroop 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2910
            Left            =   120
            Style           =   1  'Simple Combo
            TabIndex        =   52
            Text            =   "CbSelectTroop"
            Top             =   3960
            Width           =   7335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "可加入兵种:"
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
            Index           =   23
            Left            =   120
            TabIndex        =   73
            Top             =   3720
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "最少:"
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
            Left            =   240
            TabIndex        =   56
            Top             =   7155
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "最多:"
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
            Index           =   9
            Left            =   2280
            TabIndex        =   55
            Top             =   7155
            Width           =   495
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps2"
      Height          =   3375
      Index           =   2
      Left            =   7680
      TabIndex        =   7
      Top             =   5040
      Width           =   3615
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
         Height          =   3255
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   3720
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
            Index           =   16
            Left            =   3600
            TabIndex        =   71
            Top             =   2640
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
            Index           =   15
            Left            =   3600
            TabIndex        =   70
            Top             =   2280
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
            Index           =   14
            Left            =   3600
            TabIndex        =   69
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
            Index           =   13
            Left            =   3600
            TabIndex        =   68
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
            Index           =   8
            Left            =   960
            TabIndex        =   67
            Top             =   2280
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
            Index           =   12
            Left            =   3600
            TabIndex        =   35
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
            Index           =   7
            Left            =   960
            TabIndex        =   34
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
            Index           =   6
            Left            =   960
            TabIndex        =   33
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
            Index           =   11
            Left            =   3600
            TabIndex        =   32
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
            Index           =   10
            Left            =   3600
            TabIndex        =   31
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
            Left            =   960
            TabIndex        =   30
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
            Left            =   960
            TabIndex        =   29
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
            Left            =   960
            TabIndex        =   28
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
            Index           =   9
            Left            =   960
            TabIndex        =   27
            Top             =   2640
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "携带"
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
         Index           =   8
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   7335
         Begin VB.TextBox txtItemsCount 
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
            Left            =   1440
            TabIndex        =   48
            Top             =   405
            Width           =   5535
         End
         Begin VB.TextBox txtMoney 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   1470
            TabIndex        =   47
            Top             =   885
            Width           =   5535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "物品数:"
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
            Left            =   645
            TabIndex        =   50
            Top             =   480
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "携带金额:"
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
            Left            =   495
            TabIndex        =   49
            Top             =   960
            Width           =   885
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
         Height          =   1575
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   7335
         Begin VB.OptionButton OptLabelSize 
            Caption         =   "大"
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
            Left            =   3240
            TabIndex        =   65
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton OptLabelSize 
            Caption         =   "中"
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
            Left            =   2280
            TabIndex        =   64
            Top             =   960
            Width           =   615
         End
         Begin VB.OptionButton OptLabelSize 
            Caption         =   "小"
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
            Left            =   1320
            TabIndex        =   63
            Top             =   960
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.ComboBox cbIcons 
            Height          =   300
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   405
            Width           =   5295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "标签:"
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
            Left            =   720
            TabIndex        =   66
            Top             =   975
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "图标:"
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
            Left            =   720
            TabIndex        =   46
            Top             =   480
            Width           =   495
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
      Height          =   3735
      Index           =   0
      Left            =   6720
      TabIndex        =   4
      Top             =   600
      Width           =   5295
      Begin VB.Frame Frame1 
         Caption         =   "触发"
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
         Index           =   2
         Left            =   240
         TabIndex        =   59
         Top             =   3360
         Width           =   7335
         Begin VB.TextBox txtMenu 
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
            Left            =   1080
            TabIndex        =   60
            Top             =   360
            Width           =   5775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*此处可能已废除，推荐设为0"
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
            Index           =   11
            Left            =   480
            TabIndex        =   62
            Top             =   960
            Width           =   2550
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "菜单:"
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
            TabIndex        =   61
            Top             =   480
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
         TabIndex        =   8
         Top             =   360
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1920
            Width           =   5295
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   2
            Left            =   1680
            TabIndex        =   13
            Top             =   1440
            Width           =   5295
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   12
            Top             =   960
            Width           =   5295
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   10
            Top             =   480
            Width           =   5295
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
            Left            =   1080
            TabIndex        =   40
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "模板名(NOW):"
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
            TabIndex        =   14
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "模板名(EN):"
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
            TabIndex        =   11
            Top             =   1035
            Width           =   1110
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "部队模板ID:"
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
            Left            =   450
            TabIndex        =   9
            Top             =   555
            Width           =   1095
         End
      End
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps1"
      Height          =   2535
      Index           =   1
      Left            =   8040
      TabIndex        =   5
      Top             =   3960
      Width           =   3615
      Begin VB.Frame Frame1 
         Caption         =   "个性"
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
         TabIndex        =   18
         Top             =   360
         Width           =   7095
         Begin VB.CheckBox chkBanditness 
            Caption         =   "土匪强盗"
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
            Height          =   255
            Left            =   960
            TabIndex        =   45
            Top             =   1800
            Width           =   1815
         End
         Begin VB.ComboBox cbAggressiveness 
            Height          =   300
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1125
            Width           =   2175
         End
         Begin VB.ComboBox cbCourage 
            Height          =   300
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   525
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*勇气8表示中立"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   180
            Index           =   13
            Left            =   4080
            TabIndex        =   72
            Top             =   570
            Width           =   1380
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "侵略性:"
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
            Left            =   885
            TabIndex        =   42
            Top             =   1170
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "勇气:"
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
            Left            =   1080
            TabIndex        =   41
            Top             =   570
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前模板(&O)"
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
      TabIndex        =   38
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
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   80
      Width           =   375
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   600
      TabIndex        =   20
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   8760
      Width           =   2175
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   720
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParty_Templates.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParty_Templates.frx":0464
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParty_Templates.frx":05BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParty_Templates.frx":0718
            Key             =   ""
         EndProperty
      EndProperty
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
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本属性(&P)"
            Key             =   "PropBag"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "行为(&A)"
            Key             =   "Actions"
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
            Caption         =   "成员(&M)"
            Key             =   "Member"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstPTs 
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
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmParty_Templates.frx":0872
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
      MouseIcon       =   "frmParty_Templates.frx":4EDF
      MousePointer    =   99  'Custom
      TabIndex        =   37
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
      MouseIcon       =   "frmParty_Templates.frx":51E9
      MousePointer    =   99  'Custom
      TabIndex        =   36
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
      TabIndex        =   19
      Top             =   160
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "模板数:"
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
      TabIndex        =   15
      Top             =   9135
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
      MouseIcon       =   "frmParty_Templates.frx":54F3
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
      MouseIcon       =   "frmParty_Templates.frx":57FD
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   9135
      Width           =   705
   End
End
Attribute VB_Name = "frmParty_Templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CustomActive As Boolean

Private Sub CApply_Click()
Dim q As Boolean
If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then

    If UCase(PTs(CurPartyTemplateID).ptID) <> UCase(CurPartyTemplate.ptID) Then             '外引
        q = ChangeStrID(PTs(CurPartyTemplateID).ptID, CurPartyTemplate.ptID)
        
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurPartyTemplate.ptID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    PTs(CurPartyTemplateID) = CurPartyTemplate
    LstPTs.ListItems(CurPartyTemplateID + 1).SubItems(1) = PTs(CurPartyTemplateID).csvName
    LstPTs.ListItems(CurPartyTemplateID + 1).SubItems(2) = PTs(CurPartyTemplateID).ptID
    
    CurPartyTemplate = PTs(CurPartyTemplateID)
    LoadPTInfo
End If

End Sub



Private Sub cbAggressiveness_Click()
If CustomActive Then
   SetPersonalities
End If
End Sub

Private Sub cbAggressiveness_Scroll()
Call cbAggressiveness_Click
End Sub

Private Sub cbCourage_Click()
If CustomActive Then
   SetPersonalities
End If
End Sub

Private Sub cbCourage_Scroll()
Call cbCourage_Click
End Sub

Private Sub cbFaction_Click()
If CustomActive Then
   CurPartyTemplate.Faction = cbFaction.ListIndex
   CurPartyTemplate.Faction_strID = Factions(CurPartyTemplate.Faction).strID
End If
End Sub

Private Sub cbFaction_Scroll()
If CustomActive Then
CurPartyTemplate.Faction = cbFaction.ListIndex
End If
End Sub

Private Sub cbIcons_Click()
Dim fI64_NOW As Integer64b

If CustomActive Then
fI64_NOW = StrToI64(CurPartyTemplate.Flags)
fI64_NOW.by(0) = cbIcons.ListIndex

CurPartyTemplate.Flags = I64toStrNZ(fI64_NOW)
End If
End Sub

Private Sub cbIcons_Scroll()
Call cbIcons_Click
End Sub

Private Sub CbSelectTroop_Click()
If CustomActive Then
    With CurPartyTemplate.Stacks(LstTroops.SelectedItem.Index)
         .ID = CbSelectTroop.ListIndex - 1
    
         If .ID = -1 Then
             .Min = 0
             .Max = 0
             .Flags = 0
             .strID = ""
             
             SetTroopExist False
             
             LstTroops.SelectedItem.Text = PublicTips(0)
             LstTroops.SelectedItem.SubItems(1) = ""
             LstTroops.SelectedItem.SubItems(2) = ""
             LstTroops.SelectedItem.SubItems(3) = ""
             LstTroops.SelectedItem.SubItems(4) = ""
             LstTroops.SelectedItem.SubItems(5) = ""
         Else
             LstTroops.SelectedItem.Text = .ID
             LstTroops.SelectedItem.SubItems(1) = Trps(.ID).csvName
             LstTroops.SelectedItem.SubItems(2) = Trps(.ID).strID
             .strID = Trps(.ID).strID
             
             SetTroopExist True
         End If
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
           If PTs(CurPartyTemplateID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), PTs(CurPartyTemplateID).ptID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then

              If CurPartyTemplateID < N_PT - 1 Then
                DelIndex PTs(CurPartyTemplateID).ptID
                For i = CurPartyTemplateID To N_PT - 2 Step 1
                    ChangeID PTs(i + 1).ptID, PTs(i + 1).ID - 1
                    j = PTs(i).ID
                    PTs(i) = PTs(i + 1)
                    PTs(i).ID = j
                    LstPTs.ListItems(i + 1).SubItems(1) = LstPTs.ListItems(i + 2).SubItems(1)
                    LstPTs.ListItems(i + 1).SubItems(2) = LstPTs.ListItems(i + 2).SubItems(2)
                Next i
                
                ReDim Preserve PTs(N_PT - 2)
                LstPTs.ListItems.Remove N_PT
                N_PT = N_PT - 1
                
              Else
                DelIndex PTs(CurPartyTemplateID).ptID
                ReDim Preserve PTs(N_PT - 2)
                LstPTs.ListItems.Remove N_PT
                
                N_PT = N_PT - 1
                CurPartyTemplateID = N_PT - 1
                
              End If
               
               LstPts_ItemClick LstPTs.ListItems(CurPartyTemplateID + 1)
               LstPTs.ListItems(CurPartyTemplateID + 1).Selected = True
               LstPTs.ListItems(CurPartyTemplateID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), PTs(CurPartyTemplateID).ptID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), PTs(CurPartyTemplateID).ptID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_PT, PTs(CurPartyTemplateID).ptID & "_New") Then
           ReDim Preserve PTs(N_PT)
           N_PT = N_PT + 1
           PTs(N_PT - 1) = PTs(CurPartyTemplateID)
           With PTs(N_PT - 1)
                 .ID = N_PT - 1
                 .ptID = .ptID & "_New"
                 .ptName = .ptName & "_New"
                 .csvName = .csvName & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstPTs.ListItems.Add(, "pt_" & PTs(N_PT - 1).ID, PTs(N_PT - 1).ID)
      
                 With oItem
                    .SubItems(1) = PTs(N_PT - 1).csvName
                    .SubItems(2) = PTs(N_PT - 1).ptID
                 End With
           LstPts_ItemClick LstPTs.ListItems(N_PT)
           LstPTs.ListItems(N_PT).Selected = True
           LstPTs.ListItems(N_PT).EnsureVisible
           
           Else
           
           MsgBox ActiveString(PublicMsgs(90), PTs(CurPartyTemplateID).ptID & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
         
      Case 2
         If CurPartyTemplateID > 0 Then
           If PTs(CurPartyTemplateID - 1).Edit And PTs(CurPartyTemplateID).Edit Then
                SwapID PTs(CurPartyTemplateID - 1).ptID, PTs(CurPartyTemplateID).ptID
                SwapPTs CurPartyTemplateID - 1, CurPartyTemplateID
                SwapListItem LstPTs.ListItems(CurPartyTemplateID), LstPTs.ListItems(CurPartyTemplateID + 1), 2, True
                
               LstPts_ItemClick LstPTs.ListItems(CurPartyTemplateID)
               LstPTs.ListItems(CurPartyTemplateID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), PTs(CurPartyTemplateID - 1).ptID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
         End If
      Case 3
        If CurPartyTemplateID + 1 <= N_PT - 1 Then
           If PTs(CurPartyTemplateID).Edit And PTs(CurPartyTemplateID + 1).Edit Then
                SwapID PTs(CurPartyTemplateID).ptID, PTs(CurPartyTemplateID + 1).ptID
                SwapPTs CurPartyTemplateID, CurPartyTemplateID + 1
                SwapListItem LstPTs.ListItems(CurPartyTemplateID + 1), LstPTs.ListItems(CurPartyTemplateID + 2), 2, True
                
                LstPts_ItemClick LstPTs.ListItems(CurPartyTemplateID + 2)
                LstPTs.ListItems(CurPartyTemplateID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), PTs(CurPartyTemplateID).ptID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
           
        End If
End Select
End Sub

Private Sub chkBanditness_Click()
If CustomActive Then
   SetPersonalities
End If
End Sub

Private Sub chkFlags_Click(Index As Integer)
Dim a As Integer64b, TemFlags As Long, tI(2) As Integer64b

If CustomActive Then

With CurPartyTemplate
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



Private Sub chkSlave_Click()
If CustomActive Then
    LstTroops.SelectedItem.SubItems(5) = GetYesNoStr(chkSlave.Value)
    CurPartyTemplate.Stacks(LstTroops.SelectedItem.Index).Flags = chkSlave.Value
End If
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstPTs, LstPTs.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstPts_ItemClick(LstPTs.SelectedItem)
End If
End Sub

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
LstPts_ItemClick LstPTs.ListItems(CurPartyTemplateID + 1)
LstPTs.ListItems(CurPartyTemplateID + 1).Selected = True
LstPTs.ListItems(CurPartyTemplateID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
CustomActive = False

InitFrames

InitIconsCombo
InitPTsListView
InitTroopsListView
InitFlagsList

LoadPTsList
LoadFactionCombo
LoadSelectTroopCombo
LoadPersonalityCombo

CurPartyTemplateID = 0
CurPartyTemplate = PTs(CurPartyTemplateID)
LoadPTInfo

TranslateForm Me
InitCMDs

Label2.Caption = Label2.Caption & N_PT
Label1(3).Caption = Replace(Label1(3).Caption, "(NOW):", "(" & UCase$(MnBInfo.Language) & "):", , , vbTextCompare)

If LCase$(MnBInfo.Language) = "en" Then
    txtPropBag(2).Enabled = False
End If

CustomActive = True
Me.Show
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

'*************************************************************************
'**函 数 名：LoadPTsList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-06 10:48:30
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadPTsList()
Dim n As Long, oItem As ListItem

LstPTs.ListItems.Clear
For n = 0 To N_PT - 1

      Set oItem = LstPTs.ListItems.Add(, "pt_" & PTs(n).ID, PTs(n).ID)
      
      With oItem
         .SubItems(1) = PTs(n).csvName
         .SubItems(2) = PTs(n).ptID
      End With
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

Private Sub InitPTsListView()
Dim n As Integer
n = 3
LstPTs.View = lvwReport
LstPTs.Sorted = False
LstPTs.ListItems.Clear
LstPTs.ColumnHeaders.Clear
LstPTs.SortOrder = lvwAscending
LstPTs.FullRowSelect = True
LstPTs.AllowColumnReorder = False
LstPTs.LabelEdit = lvwManual
LstPTs.Checkboxes = False
LstPTs.GridLines = True
LstPTs.MultiSelect = False
LstPTs.HideSelection = False

LstPTs.ColumnHeaders.Add , , PublicMsgs(13), LstPTs.Width / n / 3.6
LstPTs.ColumnHeaders.Add , , PublicEditors_Simplified(4) & PublicMsgs(14), LstPTs.Width / n * 1.5
LstPTs.ColumnHeaders.Add , , PublicEditors_Simplified(4) & "ID", LstPTs.Width / n * 1.5


End Sub

Private Sub InitTroopsListView()
Dim n As Integer
n = 6
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

   .ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(13), .Width / n / 1.8
   .ColumnHeaders.Add , , PublicEditors(1) & PublicMsgs(14), .Width / n * 1.8
   .ColumnHeaders.Add , , PublicEditors(1) & "ID", .Width / n * 1.8
   .ColumnHeaders.Add , , PublicMsgs(48), .Width / n / 1.5
   .ColumnHeaders.Add , , PublicMsgs(49), .Width / n / 1.5
   .ColumnHeaders.Add , , PublicMsgs(47), .Width / n / 1.5

End With

For n = 1 To 6
LstTroops.ListItems.Add , , PublicTips(0)
Next n

End Sub

Private Sub LoadTroopsListView()
Dim i As Integer, oItem As ListItem

CustomActive = False
With CurPartyTemplate
     For i = 1 To 6
        If .Stacks(i).strID <> "" Then
           .Stacks(i).ID = GetID(.Stacks(i).strID, True, "", -1)
        Else
           .Stacks(i).ID = -1
        End If
        
        If .Stacks(i).ID >= 0 Then
                LstTroops.ListItems(i).Text = .Stacks(i).ID
                LstTroops.ListItems(i).SubItems(1) = Trps(.Stacks(i).ID).csvName
                LstTroops.ListItems(i).SubItems(2) = Trps(.Stacks(i).ID).strID
                LstTroops.ListItems(i).SubItems(3) = .Stacks(i).Min
                LstTroops.ListItems(i).SubItems(4) = .Stacks(i).Max
                LstTroops.ListItems(i).SubItems(5) = GetYesNoStr(.Stacks(i).Flags)
                
        Else
                LstTroops.ListItems(i).Text = PublicTips(0)
                LstTroops.ListItems(i).SubItems(1) = ""
                LstTroops.ListItems(i).SubItems(2) = ""
                LstTroops.ListItems(i).SubItems(3) = ""
                LstTroops.ListItems(i).SubItems(4) = ""
                LstTroops.ListItems(i).SubItems(5) = ""
        End If
           
     Next i
     
     LstTroops.ListItems(1).Selected = True
     LstTroops_ItemClick LstTroops.ListItems(1)
End With

CustomActive = True
End Sub



Private Sub InitFrames()
Dim i As Integer

For i = 0 To 3
    With FraProps(i)
         .BorderStyle = 0
         .Top = Tab1.ClientTop
         .Left = Tab1.ClientLeft
         .Width = Tab1.ClientWidth
         .Height = Tab1.ClientHeight
           If i <> 0 Then
            .Visible = False
           End If
    End With
Next i

If LCase$(MnBInfo.Language) = "en" Then
     txtPropBag(2).Enabled = False
End If

End Sub

Private Sub LstPts_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurPartyTemplateID = Val(Item.Text)

CurPartyTemplate = PTs(CurPartyTemplateID)
LoadPTInfo
End Sub



Private Sub LstTroops_ItemClick(ByVal Item As MSComctlLib.ListItem)
With CurPartyTemplate
  If .Stacks(Item.Index).ID > -2 And .Stacks(Item.Index).ID < N_Troop Then
    CbSelectTroop.ListIndex = .Stacks(Item.Index).ID + 1
  Else
    CbSelectTroop.ListIndex = 0
    .Stacks(Item.Index).ID = -1
  End If
   
   If .Stacks(Item.Index).ID >= 0 Then
       txtMin.Text = .Stacks(Item.Index).Min
       txtMax.Text = .Stacks(Item.Index).Max
       chkSlave.Value = .Stacks(Item.Index).Flags
       
       txtMax.Enabled = True
       txtMin.Enabled = True
       chkSlave.Enabled = True
   End If
End With
End Sub

Private Sub OptLabelSize_Click(Index As Integer)
Dim a As Integer64b, TemFlags As Long, tI(3) As Integer64b

If CustomActive Then
With CurPartyTemplate
    tI(1) = StrToI64(.Flags)     '原模板Flags
    tI(2) = HexStrToI64(pf_label_mask)   '所有模板类型Flags
    
    tI(3) = HexStrToI64(OptLabelSize(Index).Tag)     '要添加的模板类型Flags
    
    tI(0) = DeleteFlags64b(tI(1), tI(2))       '清空所有模板类型Flags
    tI(0) = AddFlags64b(tI(0), tI(3))             '添加模板类型Flags
    .Flags = I64toStrNZ(tI(0))
End With
End If
End Sub

Private Sub Tab1_Click()
Dim i As Integer

For i = 0 To 3
    With FraProps(i)
         .Visible = i + 1 = Tab1.SelectedItem.Index
    End With
Next i

LoadPTInfo
End Sub

'*************************************************************************
'**函 数 名：LoadPTInfo
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
Private Sub LoadPTInfo()
Dim i As Integer, oItem As ListItem, strTem As String, lngTem As Long
CustomActive = False
With CurPartyTemplate
  Select Case Tab1.SelectedItem.Index
       Case 1
          'Recognize
          txtPropBag(0).Text = .ptID
          txtPropBag(1).Text = .ptName
          txtPropBag(2).Text = .csvName
          
          '.Faction = IIf(CheckExist(EditInfo_FactionsCount, .Faction), .Faction, 0)
          .Faction = GetID(.Faction_strID, , Factions(0).strID)
          cbFaction.ListIndex = .Faction
          
          txtMenu.Text = .Menu
       Case 2
          LoadPersonalities
       Case 3
          LoadFlagsList
       Case 4
          LoadTroopsListView
        
  End Select
End With
CustomActive = True
End Sub


Private Sub txtItemsCount_LostFocus()
Dim fI64_NOW As Integer64b
If CustomActive Then

fI64_NOW = StrToI64(CurPartyTemplate.Flags)
fI64_NOW.by(pf_carry_goods_bits \ 8) = Int(Val(txtItemsCount.Text))

CurPartyTemplate.Flags = I64toStrNZ(fI64_NOW)

End If
End Sub


Private Sub txtMax_Change()
If CustomActive Then
    LstTroops.SelectedItem.SubItems(4) = txtMax.Text
    CurPartyTemplate.Stacks(LstTroops.SelectedItem.Index).Max = Val(txtMax.Text)
End If
End Sub

Private Sub txtMenu_Change()
If CustomActive Then
    CurPartyTemplate.Menu = Val(txtMenu.Text)
End If
End Sub

Private Sub txtMin_Change()
If CustomActive Then
    LstTroops.SelectedItem.SubItems(3) = txtMin.Text
    CurPartyTemplate.Stacks(LstTroops.SelectedItem.Index).Min = Val(txtMin.Text)
End If
End Sub

Private Sub txtMoney_LostFocus()
Dim fI64_NOW As Integer64b

If CustomActive Then

fI64_NOW = StrToI64(CurPartyTemplate.Flags)
fI64_NOW.by(pf_carry_gold_bits \ 8) = Int(Val(txtMoney.Text)) / pf_carry_gold_multiplier

CurPartyTemplate.Flags = I64toStrNZ(fI64_NOW)

End If
End Sub

Private Sub txtPropBag_LostFocus(Index As Integer)
If CustomActive Then

With CurPartyTemplate

If Index < 2 Then
   txtPropBag(Index).Text = Replace(txtPropBag(Index).Text, " ", "_")
End If

Select Case Index
      Case 0
            'Check Value
            If Left$(txtPropBag(Index).Text, 3) <> "pt_" Then
                txtPropBag(Index).Text = "pt_" & txtPropBag(Index).Text
            End If
         .ptID = txtPropBag(Index).Text
      Case 1
         .ptName = txtPropBag(Index).Text
           If LCase$(MnBInfo.Language) = "en" Then
             .csvName = .ptName
           End If
      Case 2
           If LCase$(MnBInfo.Language) <> "en" Then
             .csvName = txtPropBag(Index).Text
           End If
End Select

End With

End If
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
                  CbSelectTroop.AddItem i & " " & Trps(i).csvName & "|" & Trps(i).strID
                  CbSelectTroop.ItemData(i + 1) = i
         Next i
         

End Sub

'*************************************************************************
'**函 数 名：InitFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-05 22:52:41
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitFlagsList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("不可用", "船", "静态", "小标签", "中标签", "大标签", "总是可见", "行为默认", "在城镇中自动去除", "任务特设", _
               "无标签", "人数有限", "藏身处", "显示阵营", "不可见", "不攻击平民", "平民")
TemArray = Array(pf_disabled, pf_is_ship, pf_is_static, pf_label_small, pf_label_medium, pf_label_large, pf_always_visible, pf_default_behavior, pf_auto_remove_in_town, pf_quest_party, _
               pf_no_label, pf_limit_members, pf_hide_defenders, pf_show_faction, pf_is_hidden, pf_dont_attack_civilians, pf_civilian)
         For i = 0 To UBound(strTemArray)
            If i < 3 Or i > 5 Then
               chkFlags(i).Caption = strTemArray(i)
               chkFlags(i).Tag = TemArray(i)
               
               If i = 14 Then
                  chkFlags(i).Enabled = False
               End If
            Else
               OptLabelSize(i - 3).Tag = TemArray(i)
            End If
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
'**日    期：2010-12-06 21:49:44
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFlagsList()
Dim i As Integer, fI64 As Integer64b, fI64_NOW As Integer64b, resI64 As Integer64b, Money As Long, Items As Long, NeedCheckSmallLabel As Boolean

With CurPartyTemplate

fI64_NOW = StrToI64(.Flags)
For i = 0 To chkFlags.UBound
            If i < 3 Or i > 5 Then
               fI64 = HexStrToI64(chkFlags(i).Tag)
               resI64 = And64b(fI64_NOW, fI64)
               
               If IsEqual64b(resI64, fI64) Then
                  chkFlags(i).Value = 1
               Else
                  chkFlags(i).Value = 0
               End If
            Else
               
            End If
Next i

NeedCheckSmallLabel = True

For i = 4 To 5
       fI64 = HexStrToI64(OptLabelSize(i - 3).Tag)
       resI64 = And64b(fI64_NOW, fI64)
               
       If IsEqual64b(resI64, fI64) Then
           OptLabelSize(i - 3).Value = True
           NeedCheckSmallLabel = False
       Else
           OptLabelSize(i - 3).Value = False
           
       End If
Next i
OptLabelSize(0).Value = NeedCheckSmallLabel


Items = fI64_NOW.by(pf_carry_goods_bits \ 8)
Money = fI64_NOW.by(pf_carry_gold_bits \ 8)

txtMoney.Text = Money * pf_carry_gold_multiplier
txtItemsCount.Text = Items

If fI64_NOW.by(0) > -1 And fI64_NOW.by(0) < N_MapIcon Then
   cbIcons.ListIndex = fI64_NOW.by(0)
Else
   cbIcons.ListIndex = 0
   fI64_NOW.by(0) = 0
   .Flags = I64toStrNZ(fI64_NOW)
End If
End With

End Sub

'*************************************************************************
'**函 数 名：LoadPersonalities
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-06 17:22:53
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadPersonalities()
Dim Cour As Long, Aggr As Long, IsBandit As Long

With CurPartyTemplate
      Cour = .Personality Mod 16
      Aggr = (.Personality \ 16) Mod 16
      IsBandit = .Personality \ (16 * 16)
      
      cbCourage.ListIndex = Cour
      cbAggressiveness.ListIndex = Aggr
      chkBanditness.Value = IsBandit
End With
End Sub

'*************************************************************************
'**函 数 名：SetPersonalities
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-07 14:31:31
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub SetPersonalities()
Dim Cour As Long, Aggr As Long, IsBandit As Long

With CurPartyTemplate
      Cour = cbCourage.ListIndex
      Aggr = cbAggressiveness.ListIndex
      IsBandit = chkBanditness.Value
      
      .Personality = Cour + Aggr * 16 + IsBandit * 16 * 16
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

Private Sub InitIconsCombo()
Dim i As Integer
CustomActive = False

cbIcons.Clear
For i = 0 To N_MapIcon - 1
      cbIcons.AddItem MapIcons(i).strID
Next i

cbIcons.ListIndex = 0
CustomActive = True
End Sub


'*************************************************************************
'**函 数 名：LoadPersonalityCombo
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-06 11:36:01
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadPersonalityCombo()
Dim i As Long
CustomActive = False
cbAggressiveness.Clear
cbCourage.Clear
         For i = 0 To 15
                cbAggressiveness.AddItem i
                cbAggressiveness.ItemData(i) = i * 16
                cbCourage.AddItem i
                cbAggressiveness.ItemData(i) = i
         Next i
CustomActive = True
End Sub

Private Sub SetTroopExist(ByVal Switch As Boolean)
CustomActive = False

If Not Switch Then
  txtMin.Text = ""
  txtMax.Text = ""
  chkSlave.Value = 0
End If

txtMin.Enabled = Switch
txtMax.Enabled = Switch
chkSlave.Enabled = Switch

CustomActive = True
End Sub

Public Sub ReLoadInfo()
LoadPTInfo
End Sub
