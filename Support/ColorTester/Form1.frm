VERSION 5.00
Begin VB.Form c 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ColorTester"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7710
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "S H O W (&S)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Invert 
      Caption         =   "Invert(&I)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Text            =   "255,255,255"
      Top             =   840
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3915
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "RGB:"
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   4
      Top             =   840
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "HEX:"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'On Error GoTo ErrLine
Dim tC() As String, R As Long, G As Long, B As Long
If Option1(0).Value Then
tC = Split(Text1(0).Text, ",")
Picture1.BackColor = RGB(Val(tC(0)), Val(tC(1)), Val(tC(2)))
Text1(1).Text = "&H" & Hex(Val(Picture1.BackColor))
Else
Picture1.BackColor = Val(Text1(1).Text)
Text1(1).Text = "&H" & Hex(Val(Text1(1).Text))
'分解RGB颜色值
R = (Picture1.BackColor Mod 256)
B = (Int(Picture1.BackColor \ 65536))
G = ((Picture1.BackColor - (B * 65536) - R) \ 256)
Text1(0).Text = R & "," & G & "," & B
End If
Exit Sub
ErrLine:
MsgBox "Error!"
End Sub

Private Sub Invert_Click()
Dim tC() As String, i As Integer
tC = Split(Text1(0).Text, ",")

For i = 0 To 2
     tC(i) = 255 - Val(tC(i))
Next i
Text1(0).Text = tC(0) & "," & tC(1) & "," & tC(2)

Picture1.BackColor = RGB(Val(tC(0)), Val(tC(1)), Val(tC(2)))
Text1(1).Text = "&H" & Hex(Val(Picture1.BackColor))
End Sub
