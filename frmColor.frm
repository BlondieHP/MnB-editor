VERSION 5.00
Begin VB.Form frmColor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "调色板"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5460
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtColor 
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtColor 
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtColor 
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label CStart 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   315
      Left            =   3480
      MouseIcon       =   "frmColor.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2040
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "蓝:"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   1500
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "绿:"
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
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   900
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "红:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   300
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public InitColor As Long

Private Sub CStart_Click()

frmFactions.PicColor.BackColor = Shape1.FillColor
frmFactions.txtPropBag(3).Text = Hex(Shape1.FillColor)
CurrentFaction.lColor = RGBtoMnBColor(frmFactions.txtPropBag(3).Text)

Unload Me
End Sub

Private Sub CStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CStart.ForeColor = &H80FF&
End Sub

Private Sub Form_Load()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CStart.ForeColor = &H4080&
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmFactions.Enabled = True
End Sub

Private Sub txtColor_Change(Index As Integer)
Dim R As Long, G As Long, B As Long

R = CLng(Val(txtColor(0).Text))
G = CLng(Val(txtColor(1).Text))
B = CLng(Val(txtColor(2).Text))

If R > 255 Then R = 255
If G > 255 Then G = 255
If B > 255 Then B = 255

If R < 0 Then R = 0
If G < 0 Then G = 0
If B < 0 Then B = 0

txtColor(0).Text = R
txtColor(1).Text = G
txtColor(2).Text = B

Shape1.FillColor = RGB(R, G, B)
End Sub
