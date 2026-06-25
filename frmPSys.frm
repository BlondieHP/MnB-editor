VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPSys 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "粒子系统编辑器"
   ClientHeight    =   9375
   ClientLeft      =   3105
   ClientTop       =   960
   ClientWidth     =   14355
   Icon            =   "frmPSys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_8"
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
            Picture         =   "frmPSys.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPSys.frx":0C64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前粒子系统(&O)"
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
      TabIndex        =   16
      Top             =   8760
      Width           =   2415
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
   Begin MSComctlLib.ListView LstPSys 
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
      Picture         =   "frmPSys.frx":0FFE
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps1"
      Height          =   7935
      Index           =   1
      Left            =   6240
      TabIndex        =   10
      Top             =   480
      Width           =   7815
      Begin VB.Frame Frame3 
         Caption         =   "旋转"
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
         Height          =   1455
         Left            =   3960
         TabIndex        =   105
         Top             =   3240
         Width           =   3735
         Begin VB.TextBox txtRD 
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
            Left            =   1560
            TabIndex        =   109
            Text            =   "RD"
            Top             =   880
            Width           =   1455
         End
         Begin VB.TextBox txtRV 
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
            Left            =   1560
            TabIndex        =   107
            Text            =   "RV"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "旋转阻尼:"
            Height          =   180
            Left            =   600
            TabIndex        =   108
            Top             =   960
            Width           =   810
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "旋转速度:"
            Height          =   180
            Left            =   600
            TabIndex        =   106
            Top             =   435
            Width           =   810
         End
      End
      Begin VB.Frame FraEmit 
         Caption         =   "喷射"
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
         Height          =   2775
         Left            =   3960
         TabIndex        =   88
         Top             =   240
         Width           =   3735
         Begin VB.Frame FraESpd 
            Caption         =   "喷射初速度:"
            Height          =   1575
            Left            =   1920
            TabIndex        =   98
            Top             =   360
            Width           =   1575
            Begin VB.TextBox TxtESpd 
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
               Left            =   480
               TabIndex        =   101
               Text            =   "X"
               Top             =   300
               Width           =   855
            End
            Begin VB.TextBox TxtESpd 
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
               Left            =   480
               TabIndex        =   100
               Text            =   "Y"
               Top             =   670
               Width           =   855
            End
            Begin VB.TextBox TxtESpd 
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
               Index           =   2
               Left            =   480
               TabIndex        =   99
               Text            =   "Z"
               Top             =   1040
               Width           =   855
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "X:"
               Height          =   180
               Left            =   240
               TabIndex        =   104
               Top             =   360
               Width           =   180
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "Z:"
               Height          =   180
               Left            =   240
               TabIndex        =   103
               Top             =   1080
               Width           =   180
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Y:"
               Height          =   180
               Left            =   240
               TabIndex        =   102
               Top             =   720
               Width           =   180
            End
         End
         Begin VB.Frame FraEBox 
            Caption         =   "喷射区域:"
            Height          =   1575
            Left            =   240
            TabIndex        =   91
            Top             =   360
            Width           =   1575
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
               Index           =   2
               Left            =   480
               TabIndex        =   97
               Text            =   "Z"
               Top             =   1040
               Width           =   855
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
               Left            =   480
               TabIndex        =   96
               Text            =   "Y"
               Top             =   670
               Width           =   855
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
               Index           =   0
               Left            =   480
               TabIndex        =   95
               Text            =   "X"
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Y:"
               Height          =   180
               Left            =   240
               TabIndex        =   94
               Top             =   720
               Width           =   180
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Z:"
               Height          =   180
               Left            =   240
               TabIndex        =   93
               Top             =   1080
               Width           =   180
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "X:"
               Height          =   180
               Left            =   240
               TabIndex        =   92
               Top             =   360
               Width           =   180
            End
         End
         Begin VB.TextBox txtERan 
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
            Left            =   1920
            TabIndex        =   90
            Text            =   "ERan"
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "喷射路径随机性:"
            Height          =   180
            Left            =   480
            TabIndex        =   89
            Top             =   2230
            Width           =   1350
         End
      End
      Begin VB.Frame FraKey 
         Caption         =   "变化"
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
         Height          =   7095
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   3495
         Begin VB.Frame FraScale 
            Caption         =   "尺寸键2"
            Height          =   1095
            Index           =   1
            Left            =   1800
            TabIndex        =   65
            Top             =   5640
            Width           =   1575
            Begin VB.TextBox ScaleMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   67
               Text            =   "Mag1"
               Top             =   680
               Width           =   615
            End
            Begin VB.TextBox ScaleTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   66
               Text            =   "Time1"
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   69
               Top             =   720
               Width           =   450
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   68
               Top             =   360
               Width           =   450
            End
         End
         Begin VB.Frame FraScale 
            Caption         =   "尺寸键1"
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   60
            Top             =   5640
            Width           =   1575
            Begin VB.TextBox ScaleTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   62
               Text            =   "Time1"
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox ScaleMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   61
               Text            =   "Mag1"
               Top             =   680
               Width           =   615
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   64
               Top             =   360
               Width           =   450
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   63
               Top             =   720
               Width           =   450
            End
         End
         Begin VB.Frame FraBlue 
            Caption         =   "蓝键2"
            Height          =   1095
            Index           =   1
            Left            =   1800
            TabIndex        =   55
            Top             =   4320
            Width           =   1575
            Begin VB.TextBox BlueTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   57
               Text            =   "Time1"
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox BlueMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   56
               Text            =   "Mag1"
               Top             =   680
               Width           =   615
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   59
               Top             =   360
               Width           =   450
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   58
               Top             =   720
               Width           =   450
            End
         End
         Begin VB.Frame FraBlue 
            Caption         =   "蓝键1"
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   4320
            Width           =   1575
            Begin VB.TextBox BlueMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   52
               Text            =   "Mag1"
               Top             =   680
               Width           =   615
            End
            Begin VB.TextBox BlueTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   51
               Text            =   "Time1"
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   54
               Top             =   720
               Width           =   450
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   53
               Top             =   360
               Width           =   450
            End
         End
         Begin VB.Frame FraGreen 
            Caption         =   "绿键2"
            Height          =   1095
            Index           =   1
            Left            =   1800
            TabIndex        =   45
            Top             =   3000
            Width           =   1575
            Begin VB.TextBox GreenMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   47
               Text            =   "Mag1"
               Top             =   680
               Width           =   615
            End
            Begin VB.TextBox GreenTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   46
               Text            =   "Time1"
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   49
               Top             =   720
               Width           =   450
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   48
               Top             =   360
               Width           =   450
            End
         End
         Begin VB.Frame FraGreen 
            Caption         =   "绿键1"
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   3000
            Width           =   1575
            Begin VB.TextBox GreenTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   42
               Text            =   "Time1"
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox GreenMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   41
               Text            =   "Mag1"
               Top             =   680
               Width           =   615
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   44
               Top             =   360
               Width           =   450
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   43
               Top             =   720
               Width           =   450
            End
         End
         Begin VB.Frame FraRed 
            Caption         =   "红键2"
            Height          =   1095
            Index           =   1
            Left            =   1800
            TabIndex        =   35
            Top             =   1680
            Width           =   1575
            Begin VB.TextBox RedTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   37
               Text            =   "Time2"
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox RedMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   36
               Text            =   "Mag2"
               Top             =   680
               Width           =   615
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   39
               Top             =   360
               Width           =   450
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   38
               Top             =   720
               Width           =   450
            End
         End
         Begin VB.Frame FraRed 
            Caption         =   "红键1"
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   1575
            Begin VB.TextBox RedMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   32
               Text            =   "Mag1"
               Top             =   680
               Width           =   615
            End
            Begin VB.TextBox RedTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   31
               Text            =   "Time1"
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   34
               Top             =   720
               Width           =   450
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   33
               Top             =   360
               Width           =   450
            End
         End
         Begin VB.Frame FraAlpha 
            Caption         =   "阿尔法键1"
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1575
            Begin VB.TextBox AlphaTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   27
               Text            =   "Time1"
               Top             =   300
               Width           =   615
            End
            Begin VB.TextBox AlphaMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   0
               Left            =   720
               TabIndex        =   26
               Text            =   "Mag1"
               Top             =   680
               Width           =   615
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   29
               Top             =   360
               Width           =   450
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   28
               Top             =   720
               Width           =   450
            End
         End
         Begin VB.Frame FraAlpha 
            Caption         =   "阿尔法键2"
            Height          =   1095
            Index           =   1
            Left            =   1800
            TabIndex        =   20
            Top             =   360
            Width           =   1575
            Begin VB.TextBox AlphaMag 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   22
               Text            =   "Mag2"
               Top             =   680
               Width           =   615
            End
            Begin VB.TextBox AlphaTime 
               Alignment       =   1  'Right Justify
               Height          =   270
               Index           =   1
               Left            =   720
               TabIndex        =   21
               Text            =   "Time2"
               Top             =   300
               Width           =   615
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "大小:"
               Height          =   180
               Left            =   240
               TabIndex        =   24
               Top             =   720
               Width           =   450
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "时间:"
               Height          =   180
               Left            =   240
               TabIndex        =   23
               Top             =   360
               Width           =   450
            End
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
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
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
            Caption         =   "效果(&S)"
            Key             =   "PropBag"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
      Height          =   7935
      Index           =   0
      Left            =   6240
      TabIndex        =   9
      Top             =   480
      Width           =   7815
      Begin VB.Frame FraAttr 
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
         Height          =   2775
         Left            =   360
         TabIndex        =   75
         Top             =   1920
         Width           =   3135
         Begin VB.TextBox txtDamp 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1680
            TabIndex        =   87
            Text            =   "Damp"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtG 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1680
            TabIndex        =   86
            Text            =   "G"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtTSZ 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1680
            TabIndex        =   85
            Text            =   "TSZ"
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtTStr 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1680
            TabIndex        =   84
            Text            =   "TStr"
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txtLife 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1680
            TabIndex        =   83
            Text            =   "Life"
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtPNum 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1680
            TabIndex        =   82
            Text            =   "PNum"
            Top             =   435
            Width           =   735
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "震荡强度:"
            Height          =   180
            Left            =   720
            TabIndex        =   81
            Top             =   2325
            Width           =   810
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "震荡大小:"
            Height          =   180
            Left            =   720
            TabIndex        =   80
            Top             =   1965
            Width           =   810
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "重力:"
            Height          =   180
            Left            =   1080
            TabIndex        =   79
            Top             =   1605
            Width           =   450
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "阻尼:"
            Height          =   180
            Left            =   1080
            TabIndex        =   78
            Top             =   1245
            Width           =   450
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "粒子寿命:"
            Height          =   180
            Left            =   720
            TabIndex        =   77
            Top             =   885
            Width           =   810
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "粒子数量:"
            Height          =   180
            Left            =   720
            TabIndex        =   76
            Top             =   480
            Width           =   810
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
         Left            =   360
         TabIndex        =   70
         Top             =   360
         Width           =   6975
         Begin VB.TextBox TxtMeshName 
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
            TabIndex        =   73
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox TxtPSysName 
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
            TabIndex        =   71
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "模型名:"
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
            TabIndex        =   74
            Top             =   885
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "粒子系统名:"
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
            TabIndex        =   72
            Top             =   405
            Width           =   1245
         End
      End
      Begin VB.Frame FraFlags 
         Caption         =   "特性"
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
         Height          =   2775
         Left            =   3840
         TabIndex        =   17
         Top             =   1920
         Width           =   3495
         Begin VB.ListBox LstFlags 
            Height          =   2160
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   18
            Top             =   360
            Width           =   3015
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
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   3
      Left            =   3480
      MouseIcon       =   "frmPSys.frx":566B
      MousePointer    =   99  'Custom
      TabIndex        =   15
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
      MouseIcon       =   "frmPSys.frx":5975
      MousePointer    =   99  'Custom
      TabIndex        =   14
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
      MouseIcon       =   "frmPSys.frx":5C7F
      MousePointer    =   99  'Custom
      TabIndex        =   13
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
      MouseIcon       =   "frmPSys.frx":5F89
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9120
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "粒子系统数:"
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
Attribute VB_Name = "frmPSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CustomActive As Boolean

Private Sub InitPSysListView()
Dim n As Integer
n = 2
LstPSys.View = lvwReport
LstPSys.Sorted = False
LstPSys.ListItems.Clear
LstPSys.ColumnHeaders.Clear
LstPSys.SortOrder = lvwAscending
LstPSys.FullRowSelect = True
LstPSys.AllowColumnReorder = False
LstPSys.LabelEdit = lvwManual
LstPSys.Checkboxes = False
LstPSys.GridLines = True
LstPSys.MultiSelect = False
LstPSys.HideSelection = False

LstPSys.ColumnHeaders.Add , , PublicMsgs(13), LstPSys.Width / n / 4
LstPSys.ColumnHeaders.Add , , PublicEditors(8) & PublicMsgs(14), LstPSys.Width / n * 1.5

End Sub

'*************************************************************************
'**函 数 名：LoadPSysList
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
Private Sub LoadPSysList()
Dim n As Long, oItem As ListItem, tI As Integer64b, H As Integer

LstPSys.ListItems.Clear  '清空列表

For n = 0 To UBound(PSys)

  H = n - 1      '上一物品
  If H < 0 Then H = 0

      Set oItem = LstPSys.ListItems.Add(, "psys_" & CStr(n), n)
      
      With oItem
       .SubItems(1) = PSys(n).strID
      End With

Next n

End Sub

Private Sub AlphaMag_LostFocus(Index As Integer)
CurrentPSys.Alphak(Index).Y = Round(Val(AlphaMag(Index).Text), 6)
End Sub

Private Sub AlphaTime_LostFocus(Index As Integer)
CurrentPSys.Alphak(Index).X = Round(Val(AlphaTime(Index).Text), 6)
End Sub

Private Sub RedMag_LostFocus(Index As Integer)
CurrentPSys.Redk(Index).Y = Round(Val(RedMag(Index).Text), 6)
End Sub

Private Sub RedTime_LostFocus(Index As Integer)
CurrentPSys.Redk(Index).X = Round(Val(RedTime(Index).Text), 6)
End Sub

Private Sub GreenMag_LostFocus(Index As Integer)
CurrentPSys.Greenk(Index).Y = Round(Val(GreenMag(Index).Text), 6)
End Sub

Private Sub GreenTime_LostFocus(Index As Integer)
CurrentPSys.Greenk(Index).X = Round(Val(GreenTime(Index).Text), 6)
End Sub

Private Sub BlueMag_LostFocus(Index As Integer)
CurrentPSys.Bluek(Index).Y = Round(Val(BlueMag(Index).Text), 6)
End Sub

Private Sub BlueTime_LostFocus(Index As Integer)
CurrentPSys.Bluek(Index).X = Round(Val(BlueTime(Index).Text), 6)
End Sub

Private Sub ScaleMag_LostFocus(Index As Integer)
CurrentPSys.Scalek(Index).Y = Round(Val(ScaleMag(Index).Text), 6)
End Sub

Private Sub ScaleTime_LostFocus(Index As Integer)
CurrentPSys.Scalek(Index).X = Round(Val(ScaleTime(Index).Text), 6)
End Sub

Private Sub CApply_Click()
Dim q As Boolean

If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then

    If UCase(PSys(CurrentPSysID).strID) <> UCase(CurrentPSys.strID) Then               '外引
        q = ChangeStrID(PSys(CurrentPSysID).strID, CurrentPSys.strID)
        
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentPSys.strID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    PSys(CurrentPSysID) = CurrentPSys
    
    LstPSys.ListItems(CurrentPSysID + 1).SubItems(1) = PSys(CurrentPSysID).strID

End If

End Sub

Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If PSys(CurrentPSysID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), PSys(CurrentPSysID).strID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then

              If CurrentPSysID < N_PSys - 1 Then
                DelIndex PSys(CurrentPSysID).strID
                For i = CurrentPSysID To N_PSys - 2 Step 1
                    ChangeID PSys(i + 1).strID, PSys(i + 1).ID - 1
                    j = PSys(i).ID
                    PSys(i) = PSys(i + 1)
                    PSys(i).ID = j
                    LstPSys.ListItems(i + 1).SubItems(1) = LstPSys.ListItems(i + 2).SubItems(1)
 
                Next i
                
                ReDim Preserve PSys(N_PSys - 2)
                LstPSys.ListItems.Remove N_PSys
                N_PSys = N_PSys - 1
                
              Else
                DelIndex PSys(CurrentPSysID).strID
                ReDim Preserve PSys(N_PSys - 2)
                LstPSys.ListItems.Remove N_PSys
                
                N_PSys = N_PSys - 1
                CurrentPSysID = N_PSys - 1
                
              End If
               
               LstPSys_ItemClick LstPSys.ListItems(CurrentPSysID + 1)
               LstPSys.ListItems(CurrentPSysID + 1).Selected = True
               LstPSys.ListItems(CurrentPSysID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), PSys(CurrentPSysID).strID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), PSys(CurrentPSysID).strID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_PSys, PSys(CurrentPSysID).strID & "_New") Then
           ReDim Preserve PSys(N_PSys)
           N_PSys = N_PSys + 1
           PSys(N_PSys - 1) = PSys(CurrentPSysID)
           With PSys(N_PSys - 1)
                 .ID = N_PSys - 1
                 .strID = .strID & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstPSys.ListItems.Add(, "psys_" & PSys(N_PSys - 1).ID, PSys(N_PSys - 1).ID)
           
                 With oItem
                    .SubItems(1) = PSys(N_PSys - 1).strID
                 End With
                 
           LstPSys_ItemClick LstPSys.ListItems(N_PSys)
           LstPSys.ListItems(N_PSys).Selected = True
           LstPSys.ListItems(N_PSys).EnsureVisible
           Else
           
           MsgBox ActiveString(PublicMsgs(90), PSys(CurrentPSysID).strID & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
         
      Case 2
         If CurrentPSysID > 0 Then
           If PSys(CurrentPSysID - 1).Edit And PSys(CurrentPSysID).Edit Then
                SwapID PSys(CurrentPSysID - 1).strID, PSys(CurrentPSysID).strID
                SwapPSys CurrentPSysID - 1, CurrentPSysID
                SwapListItem LstPSys.ListItems(CurrentPSysID), LstPSys.ListItems(CurrentPSysID + 1), 1, True
                
               LstPSys_ItemClick LstPSys.ListItems(CurrentPSysID)
               LstPSys.ListItems(CurrentPSysID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), PSys(CurrentPSysID - 1).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
         End If
      Case 3
        If CurrentPSysID + 1 <= N_PSys - 1 Then
           If PSys(CurrentPSysID).Edit And PSys(CurrentPSysID + 1).Edit Then
                SwapID PSys(CurrentPSysID).strID, PSys(CurrentPSysID + 1).strID
                SwapPSys CurrentPSysID, CurrentPSysID + 1
                SwapListItem LstPSys.ListItems(CurrentPSysID + 1), LstPSys.ListItems(CurrentPSysID + 2), 1, True
                
                LstPSys_ItemClick LstPSys.ListItems(CurrentPSysID + 2)
                LstPSys.ListItems(CurrentPSysID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), PSys(CurrentPSysID).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
        End If
End Select
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstPSys, LstPSys.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstPSys_ItemClick(LstPSys.SelectedItem)
End If
End Sub

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbExclamation, PublicMsgs(0)) = vbYes Then
LstPSys_ItemClick LstPSys.ListItems(CurrentPSysID + 1)
LstPSys.ListItems(CurrentPSysID + 1).Selected = True
LstPSys.ListItems(CurrentPSysID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
'Inits
CustomActive = False

InitPSysListView
LoadPSysList
InitFrames
InitLstFlags

LstPSys.ListItems(1).Selected = True
CurrentPSysID = 0
CurrentPSys = PSys(0)
Call LoadPSysInfo(CurrentPSys)

TranslateForm Me
InitCMDs

Label2.Caption = Label2.Caption & N_PSys

CustomActive = True
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

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

Private Sub LstFlags_LostFocus()
Dim i As Integer, tI As Integer64b

With CurrentPSys
    'CurrentPSys.Flags = ""
    For i = 0 To UBound(PSf)
          If LstFlags.Selected(i) Then
             tI = AddFlagsI64(tI, HexStrToI64(PSf(i).X))
          End If
    Next i
    .Flags = I64toStrNZ(tI)
End With

End Sub

Private Sub LstPSys_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim n As Integer

CurrentPSysID = Val(Item.Text)
CurrentPSys = PSys(CurrentPSysID)


LoadPSysInfo CurrentPSys

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

'If LCase$(MnBInfo.Language) = "en" Then
'     Text1(2).Enabled = False
'End If

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

LoadPSysInfo CurrentPSys
End If

End Sub

Private Sub InitLstFlags()
Dim i As Integer

LstFlags.Clear
For i = 0 To UBound(PSf)
    LstFlags.AddItem PSf(i).Z
Next i

End Sub

'*************************************************************************
'**函 数 名：LoadPSysInfo
'**输    入：PSys As Type_Particle_System
'**输    出：无
'**功能描述：载入粒子系统信息
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-3 10:13:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadPSysInfo(PSys As Type_Particle_System)
Dim i As Integer, tStr As String, tI As Integer64b, H As Byte

CustomActive = False
TxtPSysName.Text = ""

Select Case Tab1.SelectedItem.Index

Case 1
    '显示信息
    TxtPSysName.Text = PSys.strID
    TxtMeshName.Text = PSys.Mesh_Name
    
    'Flags
    For i = 0 To LstFlags.ListCount - 1
         LstFlags.Selected(i) = False
    Next i
    LoadLstFlags PSys.Flags

    '属性
    txtPNum.Text = PSys.Particles_Num
    txtLife.Text = Format(PSys.Life, "0.00")
    txtDamp.Text = Format(PSys.Damping, "0.00")
    txtG.Text = Format(PSys.Gravity, "0.00")
    txtTSZ.Text = Format(PSys.Turbulance_SZ, "0.00")
    txtTStr.Text = Format(PSys.Turbulance_Str, "0.00")
Case 2
    'Keys
    For i = 0 To 1
        AlphaTime(i).Text = Format(PSys.Alphak(i).X, "0.00")
        AlphaMag(i).Text = Format(PSys.Alphak(i).Y, "0.00")
        
        RedTime(i).Text = Format(PSys.Redk(i).X, "0.00")
        RedMag(i).Text = Format(PSys.Redk(i).Y, "0.00")
        
        GreenTime(i).Text = Format(PSys.Greenk(i).X, "0.00")
        GreenMag(i).Text = Format(PSys.Greenk(i).Y, "0.00")
        
        BlueTime(i).Text = Format(PSys.Bluek(i).X, "0.00")
        BlueMag(i).Text = Format(PSys.Bluek(i).Y, "0.00")
        
        ScaleTime(i).Text = Format(PSys.Scalek(i).X, "0.00")
        ScaleMag(i).Text = Format(PSys.Scalek(i).Y, "0.00")
    Next i
    
    For i = 0 To 2
        TxtEBox(i).Text = Format(PSys.EBSZ(i), "0.00")
        TxtESpd(i).Text = Format(PSys.EV(i), "0.00")
    Next i
        txtERan.Text = Format(PSys.EDR, "0.00")
        
        txtRV.Text = Format(PSys.PRS, "0.00")
        txtRD.Text = Format(PSys.PRD, "0.00")
        
End Select

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
Dim tI As Integer64b, i As Byte, tI2 As Integer64b, n As Integer
         tI = StrToI64(Flags)
    With LstFlags
        For i = 0 To LstFlags.ListCount - 1
           If i >= 3 And i <= 5 Then
              tI2 = And64b(tI, HexStrToI64(psf_billboard_drop))
              For n = 3 To 5
                  If I64toStrNZ(tI2) = I64toStrNZ(HexStrToI64(PSf(n).X)) Then
                      LstFlags.Selected(n) = True
                  End If
              Next n
           Else
              tI2 = And64b(tI, HexStrToI64(PSf(i).X))
              If I64toStrNZ(tI2) = I64toStrNZ(HexStrToI64(PSf(i).X)) Then
                   LstFlags.Selected(i) = True
              End If
           End If
        Next i
    End With
End Sub

Private Sub txtDamp_LostFocus()
CurrentPSys.Damping = Round(Val(txtDamp.Text), 6)
End Sub

Private Sub TxtEBox_LostFocus(Index As Integer)
CurrentPSys.EBSZ(Index) = Round(Val(TxtEBox(Index).Text), 6)
End Sub

Private Sub txtERan_LostFocus()
CurrentPSys.EDR = Round(Val(txtERan.Text), 6)
End Sub

Private Sub TxtESpd_LostFocus(Index As Integer)
CurrentPSys.EV(Index) = Round(Val(TxtESpd(Index).Text), 6)
End Sub

Private Sub txtG_LostFocus()
CurrentPSys.Gravity = Round(Val(txtG.Text), 6)
End Sub

Private Sub txtLife_LostFocus()
CurrentPSys.Life = Round(Val(txtLife.Text), 6)
End Sub

Private Sub TxtMeshName_LostFocus()

TxtMeshName.Text = Replace(TxtMeshName.Text, " ", "_")
CurrentPSys.Mesh_Name = CStr(TxtMeshName.Text)

End Sub

Private Sub txtPNum_LostFocus()
CurrentPSys.Particles_Num = CLng(Val(txtPNum.Text))
End Sub

Private Sub TxtPSysName_LostFocus()

TxtPSysName.Text = Replace(TxtPSysName.Text, " ", "_")
CurrentPSys.strID = CStr(TxtPSysName.Text)

End Sub

Private Sub txtRD_LostFocus()
CurrentPSys.PRD = Round(Val(txtRD.Text), 6)

End Sub

Private Sub txtRV_LostFocus()
CurrentPSys.PRS = Round(Val(txtRV.Text), 6)
End Sub

Private Sub txtTStr_LostFocus()
CurrentPSys.Turbulance_Str = Round(Val(txtTStr.Text), 6)

End Sub

Private Sub txtTSZ_LostFocus()
CurrentPSys.Turbulance_SZ = Round(Val(txtTSZ.Text), 6)
End Sub

Public Sub ReLoadInfo()
LoadPSysInfo CurrentPSys
End Sub
