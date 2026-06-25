VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSelectPath 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择战团安装目录"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmSelectPath.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6735
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicButton 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   6705
      TabIndex        =   13
      Top             =   5010
      Width           =   6735
      Begin MSComDlg.CommonDialog CD1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.PictureBox bLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   1800
         MouseIcon       =   "frmSelectPath.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmSelectPath.frx":0BD4
         ScaleHeight     =   465
         ScaleWidth      =   1785
         TabIndex        =   16
         Top             =   0
         Width           =   1815
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "编辑器设置(&E)"
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
            Index           =   1
            Left            =   270
            TabIndex        =   17
            Top             =   120
            Width           =   1305
         End
      End
      Begin VB.PictureBox bLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   0
         MouseIcon       =   "frmSelectPath.frx":1190
         MousePointer    =   99  'Custom
         Picture         =   "frmSelectPath.frx":149A
         ScaleHeight     =   465
         ScaleWidth      =   1785
         TabIndex        =   14
         Top             =   0
         Width           =   1815
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "战团设置(&W)"
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
            Index           =   0
            Left            =   360
            TabIndex        =   15
            Top             =   120
            Width           =   1125
         End
      End
      Begin VB.Image Banner 
         Height          =   480
         Index           =   1
         Left            =   0
         Picture         =   "frmSelectPath.frx":1A56
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6720
      End
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   1
      Left            =   0
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   18
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame1 
         BackColor       =   &H80000005&
         Caption         =   "可选载入项"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   29
         Top             =   2760
         Width           =   6255
         Begin VB.CheckBox Check1 
            BackColor       =   &H80000005&
            Caption         =   "字符串"
            Height          =   375
            Left            =   600
            TabIndex        =   30
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.ComboBox cbOp 
         Height          =   300
         Left            =   2160
         TabIndex        =   25
         Text            =   "cbOp"
         Top             =   1590
         Width           =   3975
      End
      Begin VB.ComboBox txtLang_E 
         Height          =   300
         Left            =   2160
         TabIndex        =   23
         Text            =   "txtLang_E"
         Top             =   480
         Width           =   3975
      End
      Begin VB.FileListBox File1 
         Height          =   1350
         Left            =   6120
         Pattern         =   "*.lan.ini"
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "操作块配置:"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   10
         Left            =   360
         TabIndex        =   28
         Top             =   1560
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*语言名为程序目录下的lan.ini文件的文件名"
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
         Index           =   9
         Left            =   360
         TabIndex        =   27
         Top             =   2070
         Width           =   3960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(如:cns.lan.ini则填入cns)"
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
         Index           =   8
         Left            =   360
         TabIndex        =   26
         Top             =   2310
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(如:cns.lan.ini则填入cns)"
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
         Index           =   6
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   2565
      End
      Begin VB.Label COK 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   1
         Left            =   4200
         MouseIcon       =   "frmSelectPath.frx":60B3
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   4200
         Width           =   1770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*语言名为程序目录下的lan.ini文件的文件名"
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
         Index           =   7
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   3960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编辑器语言:"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   19
         Top             =   450
         Width           =   1740
      End
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Index           =   0
      Left            =   0
      ScaleHeight     =   5025
      ScaleWidth      =   6705
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   3240
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtVersion 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtLanguage 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "安装路径:"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   285
         Width           =   1425
      End
      Begin VB.Label COK 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   0
         Left            =   3840
         MouseIcon       =   "frmSelectPath.frx":63BD
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   3840
         Width           =   1770
      End
      Begin VB.Label CReg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "添加到注册表(&R)"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   3480
         MouseIcon       =   "frmSelectPath.frx":66C7
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   3360
         Width           =   2400
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "版本:"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   9
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Index           =   1
         Left            =   3240
         TabIndex        =   8
         Top             =   2640
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
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
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Top             =   2880
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "语言:"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   795
      End
      Begin VB.Image Banner 
         Height          =   5040
         Index           =   0
         Left            =   0
         Picture         =   "frmSelectPath.frx":69D1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6720
      End
   End
End
Attribute VB_Name = "frmSelectPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bLabel_Click(Index As Integer)
ShowPicBox Index
End Sub

Private Sub Check1_Click()
IsLoadString = Check1.Value
End Sub

Private Sub COK_Click(Index As Integer)
InitWarbandInfo txtPath.Text, txtVersion.Text

'可选载入项设置保存
WriteString MnBInfo.iniSetting, "Settings", "IsLoadString", CStr(IsLoadString)

If MnBInfo.Language <> txtLanguage.Text Then
  MnBInfo.Language = txtLanguage.Text

  MnBInfo.Language = Trim(MnBInfo.Language)
  If MnBInfo.Language = "" Then
     MnBInfo.Language = "en"
  End If
End If

If MnBInfo.Language_Edit <> txtLang_E.Text Then
  SelectLanguage txtLang_E.Text

  If txtLang_E.Text = "cns" Then
    MsgBox "您选用的中文(简体)为程序内置语言。若有不正确的语言显示,重启编辑器即可!", vbInformation, PublicMsgs(0)
  End If

  TranslateForm Welcome
  WriteString MnBInfo.iniSetting, "Settings", "Language", MnBInfo.Language_Edit
End If

If MnBInfo.Op_Set <> cbOp.Text Then
  MnBInfo.Op_Set = cbOp.Text
  WriteString MnBInfo.iniSetting, "Settings", "Op_Set", MnBInfo.Op_Set
  MsgBox PublicMsgs(144), vbInformation, PublicMsgs(0)
End If

Welcome.DisplayWarbandInfo
UnLoad Me
End Sub

Private Sub CReg_Click()
If MsgBox(PublicMsgs(96), vbYesNo + vbExclamation, PublicMsgs(97)) = vbYes Then
   WriteRegString HKEY_LOCAL_MACHINE, "SOFTWARE\Mount&Blade Warband", "", txtPath.Text
   WriteRegString HKEY_LOCAL_MACHINE, "SOFTWARE\Mount&Blade Warband", "Version", txtVersion.Text
   
   MsgBox PublicMsgs(98), vbInformation, PublicMsgs(97)
End If
End Sub


Private Sub Dir1_Click()
txtPath.Text = FixPath(Dir1.List(Dir1.ListIndex))
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
Me.ZOrder
End Sub

Private Sub Form_Deactivate()
Me.ZOrder
End Sub

Private Sub Form_Load()

Label1(1).Caption = "*请去除版本号中的" & Chr(34) & "." & Chr(34) & ","
Label1(2).Caption = "例如1.134,则改为1134。"

Frame1.BackColor = RGB(196, 158, 111)
Check1.BackColor = Frame1.BackColor
If IsLoadString Then
   Check1.Value = 1
Else
   Check1.Value = 0
End If

InitEditLanguages
InitOperationSets

If MnBInfo.InitFinished Then

txtVersion.Text = MnBInfo.Version
txtPath.Text = MnBInfo.MBHome
txtLanguage.Text = MnBInfo.Language

Drive1.Drive = Left(MnBInfo.MBHome, 2)
Dir1.Path = MnBInfo.MBHome

txtLang_E.Text = MnBInfo.Language_Edit
cbOp.Text = MnBInfo.Op_Set
End If

TranslateForm Me

ShowPicBox 0

End Sub

Private Sub Form_Resize()
With PicBox(0)
    Banner(0).Move 0, 0, .ScaleWidth, .ScaleHeight
End With

With PicButton
    Banner(1).Move 0, 0, .ScaleWidth, .ScaleHeight
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Welcome.Enabled = True
End Sub

Private Sub ShowPicBox(ByVal Index As Integer)
Dim i As Integer

For i = PicBox.LBound To PicBox.UBound
    With PicBox(i)
      If i = Index Then
         Set Banner(0).Container = PicBox(i)
         PicBox(i).Visible = True
      Else
         PicBox(i).Visible = False
      End If
    End With
Next i
End Sub

Private Sub Label2_Click(Index As Integer)
Call bLabel_Click(Index)
End Sub

Private Sub InitEditLanguages()
Dim i As Integer, s As String

File1.Path = App.Path
File1.Pattern = "*.lan.ini"
txtLang_E.Clear

txtLang_E.AddItem "cns"
For i = 0 To File1.ListCount - 1
    s = Left(File1.List(i), Len(File1.List(i)) - 8)
    If LCase(s) <> "cns" Then
       txtLang_E.AddItem s
    End If
Next i

End Sub

Private Sub InitOperationSets()
Dim i As Integer, s As String

File1.Path = App.Path
File1.Pattern = "*.op.ini"
cbOp.Clear

cbOp.AddItem "default"
For i = 0 To File1.ListCount - 1
    s = Left(File1.List(i), Len(File1.List(i)) - 7)
    cbOp.AddItem s
Next i

End Sub


