VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   8295
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1200
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   240
      ScaleHeight     =   2265
      ScaleWidth      =   2985
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim R As Long, G As Long, B As Long, LC As Long
'分解RGB颜色值
LC = Val(Text1.Text)
B = (LC Mod 256)
R = (Int(LC \ 65536))
G = ((LC - (R * 65536) - B) \ 256)

Picture1.BackColor = RGB(R, G, B)
'RGB-BGR
End Sub

