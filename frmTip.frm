VERSION 5.00
Begin VB.Form frmTip 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "提示"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTip.frx":0000
   ScaleHeight     =   1005
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   540
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmTip.frx":20C52
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowTip(ByVal Text As String, Optional isTopMost As Boolean = False)

Me.Show
Label1.Caption = Text
isShowTip = True
Me.ZOrder
If isTopMost Then
  SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
End If
End Sub

Public Sub HideTip()
isShowTip = False
UnLoad Me
End Sub

Private Sub Form_Deactivate()
Me.ZOrder
End Sub

Private Sub Form_Load()

TranslateForm Me

Me.Show

End Sub

