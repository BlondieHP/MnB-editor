VERSION 5.00
Begin VB.Form frmSaveAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "另存为"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "frmSaveAs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1200
      Width           =   5775
   End
   Begin VB.Label CStart 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "保存(&S)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmSaveAs.frx":1272
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2040
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "剧本名:"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   3
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Image Banner 
      Height          =   2955
      Left            =   0
      Picture         =   "frmSaveAs.frx":157C
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Banner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CStart.ForeColor = &HFFFF&
End Sub

Private Sub CStart_Click()
On Error GoTo EL

If MsgBox(ActiveString(PublicMsgs(24), MnBInfo.ModName), vbInformation + vbYesNo, PublicMsgs(25)) = vbYes Then

     
If Dir(MnBInfo.MBHome & "\Modules\" & Combo1.Text, vbDirectory) = "" Then
     MkDir MnBInfo.MBHome & "\Modules\" & Combo1.Text
     
     frmTip.ShowTip PublicTips(2)
     If LCase(MnBInfo.Language) <> "en" Then
          MkDirEx MnBInfo.MBHome & "\Modules\" & Combo1.Text & "\languages\" & LCase(MnBInfo.Language)
     End If
     
     'frmTip.ShowTip PublicTips(3)
     DoEvents
     'FileCopy MnBInfo.MBHome & "\Modules\" & MnBInfo.ModName & "\main.bmp", MnBInfo.MBHome & "\Modules\" & Combo1.Text & "\main.bmp"
     frmTip.HideTip
Else

    MsgBox ActiveString(PublicMsgs(27), MnBInfo.ModName), vbInformation, PublicMsgs(25)
End If

FinishWarbandInfo Combo1.Text
   
    'CancelTopForms
    frmTip.ShowTip PublicTips(4), True
    DoEvents
    SaveAll
    frmTip.HideTip
    'SetTopForms
    MsgBox ActiveString(PublicMsgs(29), MnBInfo.ModName), vbInformation, PublicMsgs(25)


frmMain.Caption = ActiveString(PublicMsgs(50), MnBInfo.ModName)

UnLoad Me

End If

Exit Sub

EL:
MsgBox PublicMsgs(53), vbCritical, PublicMsgs(51)
End Sub

Private Sub CStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CStart.ForeColor = &H80FF&
End Sub

Private Sub Form_Load()

LoadModules

TranslateForm Me

Me.Show

Me.Width = Banner.Width + 60
Me.Height = Banner.Height + 510 - 60
End Sub

Private Sub LoadModules()
Dim s As String

Combo1.Clear

With MnBInfo
    s = Dir(.MBHome & "\Modules\", vbDirectory)
    Do While s <> ""
        If IsDirectory(.MBHome & "\Modules\" & s) Then   '排除目录首
           Combo1.AddItem s
        End If
        s = Dir
    Loop

End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True
End Sub
