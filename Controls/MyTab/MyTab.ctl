VERSION 5.00
Begin VB.UserControl MyTab 
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   ScaleHeight     =   5925
   ScaleWidth      =   6555
   Begin VB.PictureBox PicTab 
      BackColor       =   &H00FFC0C0&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7755
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.PictureBox PicTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   7065
         TabIndex        =   1
         Top             =   0
         Width           =   7095
         Begin VB.PictureBox PicButton 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   0
            ScaleHeight     =   465
            ScaleWidth      =   1905
            TabIndex        =   2
            Top             =   0
            Width           =   1935
            Begin VB.Image TabImage 
               Height          =   240
               Index           =   0
               Left            =   120
               Picture         =   "MyTab.ctx":0000
               Top             =   120
               Width           =   240
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "基本属性(&P)"
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
               Height          =   255
               Left            =   480
               TabIndex        =   3
               Top             =   140
               Width           =   1095
            End
         End
      End
   End
End
Attribute VB_Name = "MyTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

