VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "剧本:[ModName]"
   ClientHeight    =   9015
   ClientLeft      =   2490
   ClientTop       =   435
   ClientWidth     =   13395
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   13395
   StartUpPosition =   2  '屏幕中心
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   5535
      Left            =   11880
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   8400
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.PictureBox LstButtons 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1035
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   7560
      Width           =   2895
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备份剧本(&B)"
         Height          =   180
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmMain.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Tag             =   "file_3"
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "保存剧本(&V)"
         Height          =   180
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmMain.frx":0BD4
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Tag             =   "file_2"
         Top             =   720
         Width           =   990
      End
      Begin VB.Image ImgButton 
         Height          =   480
         Index           =   1
         Left            =   1800
         MouseIcon       =   "frmMain.frx":0EDE
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":11E8
         Tag             =   "file_3"
         Top             =   120
         Width           =   480
      End
      Begin VB.Image ImgButton 
         Height          =   480
         Index           =   0
         Left            =   480
         MouseIcon       =   "frmMain.frx":1AB2
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1DBC
         Tag             =   "file_2"
         Top             =   120
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   2040
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2686
            Key             =   "edit_1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F60
            Key             =   "edit_2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":383A
            Key             =   "edit_3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B54
            Key             =   "edit_4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E6E
            Key             =   "edit_5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4748
            Key             =   "edit_6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5022
            Key             =   "tool_1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":58FC
            Key             =   "edit_7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":620B
            Key             =   "tool_0"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AE5
            Key             =   "edit_0"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73BF
            Key             =   "about_0"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C99
            Key             =   "about_1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8573
            Key             =   "file_0"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E4D
            Key             =   "file_1"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9727
            Key             =   "file_3"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A001
            Key             =   "file_2"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8DB
            Key             =   "tool_2"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B1B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BA8F
            Key             =   "edit_9"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C369
            Key             =   "edit_8"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CC43
            Key             =   "edit_10"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF5D
            Key             =   "edit_11"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D837
            Key             =   "tool_3"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E111
            Key             =   "edit_12"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E9EB
            Key             =   "edit_13"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F2C5
            Key             =   "tool_4"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB9F
            Key             =   "tool_5"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TvwMenu 
      Height          =   7335
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   12938
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "IL1"
      Appearance      =   1
   End
   Begin VB.PictureBox PicMain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8175
      Left            =   3240
      Picture         =   "frmMain.frx":10479
      ScaleHeight     =   8115
      ScaleWidth      =   9735
      TabIndex        =   0
      Top             =   240
      Width           =   9795
   End
   Begin VB.Image ImgBack 
      Height          =   9015
      Index           =   1
      Left            =   13080
      Picture         =   "frmMain.frx":23B26
      Stretch         =   -1  'True
      Top             =   0
      Width           =   330
   End
   Begin VB.Image ImgBack 
      Height          =   8415
      Index           =   8
      Left            =   3120
      Picture         =   "frmMain.frx":33EA0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   255
   End
   Begin VB.Image ImgBack 
      Height          =   375
      Index           =   7
      Left            =   13080
      Picture         =   "frmMain.frx":3BE3E
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   375
   End
   Begin VB.Image ImgBack 
      Height          =   255
      Index           =   6
      Left            =   13080
      Picture         =   "frmMain.frx":3C7B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image ImgBack 
      Height          =   375
      Index           =   5
      Left            =   0
      Picture         =   "frmMain.frx":3D0B2
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   255
   End
   Begin VB.Image ImgBack 
      Height          =   255
      Index           =   4
      Left            =   0
      Picture         =   "frmMain.frx":3DA24
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Image ImgBack 
      Height          =   380
      Index           =   3
      Left            =   0
      Picture         =   "frmMain.frx":3E21E
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   13335
   End
   Begin VB.Image ImgBack 
      Height          =   9015
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":51A60
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Image ImgBack 
      Height          =   255
      Index           =   2
      Left            =   0
      Picture         =   "frmMain.frx":61182
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
   Begin VB.Menu mFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mSelectMod 
         Caption         =   "选择剧本(&M)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mSave 
         Caption         =   "保存剧本(&S)"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mBackUp 
         Caption         =   "备份剧本(&B)"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "另存为剧本(&A)"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mDash 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mQuit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mWindows 
      Caption         =   "窗口(&W)"
      Begin VB.Menu mDisplayMode 
         Caption         =   "可调窗体模式(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mBanfrmInfo 
         Caption         =   "禁用项目信息窗口(&I)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mAdvance 
         Caption         =   "高级功能(&A)"
         Begin VB.Menu mInitTranslateForms 
            Caption         =   "导出窗体语言(&O)"
         End
         Begin VB.Menu mOutputOperation 
            Caption         =   "导出注册操作(&R)"
         End
      End
   End
   Begin VB.Menu mEditor 
      Caption         =   "编辑器(&E)"
      Enabled         =   0   'False
      Begin VB.Menu mEditors 
         Caption         =   "编辑器"
         Index           =   0
      End
   End
   Begin VB.Menu mTool 
      Caption         =   "工具(&T)"
      Enabled         =   0   'False
      Begin VB.Menu mTools 
         Caption         =   "工具"
         Index           =   0
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mHelpTheme 
         Caption         =   "帮助主题(&H)"
      End
      Begin VB.Menu mTutorials 
         Caption         =   "修改教程(&T)"
         Begin VB.Menu mIndex 
            Caption         =   "【魔球教程索引贴】"
         End
         Begin VB.Menu mTutorial 
            Caption         =   "用魔球轻松修改出火剑"
            Index           =   0
         End
         Begin VB.Menu mTutorial 
            Caption         =   "用魔球导入一把火枪"
            Index           =   1
         End
         Begin VB.Menu mTutorial 
            Caption         =   "通过魔球粒子系统编辑器制作蓝色火剑"
            Index           =   2
         End
         Begin VB.Menu mTutorial 
            Caption         =   "用魔球制作火箭"
            Index           =   3
         End
         Begin VB.Menu mTutorial 
            Caption         =   "用魔球制作一件纹章甲"
            Index           =   4
         End
      End
      Begin VB.Menu mDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "关于MnBWarband Editor(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Display As Form
Public HaveForm As Boolean, ForceQuit As Boolean

Private Sub Form_Load()
On Error GoTo errorHandle

ForceQuit = False

InitMenu

TranslateForm Me

frmMain.mSave.Enabled = False
frmMain.mSaveAs.Enabled = False
frmMain.mBackUp.Enabled = False
frmMain.mEditor.Enabled = False
frmMain.mTool.Enabled = False
mDisplayMode.Checked = CBool(DisplayMode)
mBanfrmInfo.Checked = CBool(Val(ReadString(MnBInfo.iniSetting, "Settings", "BanfrmInfo", 250)))

ReadAll

HaveForm = False

InitFormSize

Exit Sub
errorHandle:
    Call logErr("frmMain", "Form_Load", Err.Number, Err.Description)
End Sub

Private Sub InitFormSize()
On Error GoTo errorHandle

Dim i As Integer

With Me
  
  If DisplayMode = 0 Then
    SetFormStyle Me, MY_NOT_SIZABLE
    TvwMenu.Move 240, 240, TvwMenu.Width, 9375 - 1095
    LstButtons.Move 240, TvwMenu.Top + TvwMenu.Height, TvwMenu.Width, 1095
    PicMain.Move TvwMenu.Left + TvwMenu.Width + 105, 240, 14355, 9375

    Me.Width = PicMain.Left + PicMain.Width + 330 + 60 * 2
    Me.Height = PicMain.Top + PicMain.Height + 800 + 330  '+60
    VScroll.Visible = False
    HScroll.Visible = False
  ElseIf DisplayMode = 1 Then
    SetFormStyle Me, MY_SIZABLE
    
    TvwMenu.Move 240, 240, TvwMenu.Width, .ScaleHeight - 330 - 240 - 1095
    LstButtons.Move 240, TvwMenu.Top + TvwMenu.Height, TvwMenu.Width, 1095
    PicMain.Move TvwMenu.Left + TvwMenu.Width + 105, 240, .ScaleWidth - PicMain.Left - 330 - VScroll.Width, (.ScaleHeight - 330 - 240 - HScroll.Height)
    VScroll.Move PicMain.Left + PicMain.Width, PicMain.Top, VScroll.Width, PicMain.Height
    HScroll.Move PicMain.Left, PicMain.Top + PicMain.Height, PicMain.Width, HScroll.Height
    VScroll.Visible = True
    HScroll.Visible = True
  End If
  
    ImgBack(0).Move 0, 0, 240, .ScaleHeight
    ImgBack(1).Move .ScaleWidth - 330, 0, 330, .ScaleHeight
    ImgBack(2).Move 0, 0, .ScaleWidth, 240
    ImgBack(3).Move 0, .ScaleHeight - 330, .ScaleWidth, 330
    ImgBack(4).Move 0, 0, 240, 240
    ImgBack(5).Move 0, .ScaleHeight - 330, 240, 330
    ImgBack(6).Move .ScaleWidth - 330, 0, 330, 240
    ImgBack(7).Move .ScaleWidth - 330, .ScaleHeight - 330, 330, 330
    ImgBack(8).Move TvwMenu.Left + TvwMenu.Width, 0, ImgBack(8).Width, Me.ScaleHeight
    
    For i = 0 To 7
       ImgBack(i).ZOrder
    Next i
    
    If DisplayMode = 1 Then
       VScroll.ZOrder
       HScroll.ZOrder
    End If
    
End With

Exit Sub
errorHandle:
    Call logErr("frmMain", "InitFormSize", Err.Number, Err.Description)
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim i As Integer

If DisplayMode = 1 Then
  With Me
    If Me.Height < 2000 Then Me.Height = 2000
    If Me.Width < 2000 Then Me.Width = 2000
    
    TvwMenu.Move 240, 240, TvwMenu.Width, .ScaleHeight - 330 - 240 - 1095
    LstButtons.Move 240, TvwMenu.Top + TvwMenu.Height, TvwMenu.Width, 1095
    PicMain.Move TvwMenu.Left + TvwMenu.Width + 105, 240, .ScaleWidth - PicMain.Left - 330 - VScroll.Width, (.ScaleHeight - 330 - 240 - HScroll.Height)
    VScroll.Move PicMain.Left + PicMain.Width, PicMain.Top, VScroll.Width, PicMain.Height
    HScroll.Move PicMain.Left, PicMain.Top + PicMain.Height, PicMain.Width, HScroll.Height
    
    If HaveForm Then
      Display.Move 0, 0
      ResetScroll
    End If
    
    ImgBack(0).Move 0, 0, 240, .ScaleHeight
    ImgBack(1).Move .ScaleWidth - 330, 0, 330, .ScaleHeight
    ImgBack(2).Move 0, 0, .ScaleWidth, 240
    ImgBack(3).Move 0, .ScaleHeight - 330, .ScaleWidth, 330
    ImgBack(4).Move 0, 0, 240, 240
    ImgBack(5).Move 0, .ScaleHeight - 330, 240, 330
    ImgBack(6).Move .ScaleWidth - 330, 0, 330, 240
    ImgBack(7).Move .ScaleWidth - 330, .ScaleHeight - 330, 330, 330
    ImgBack(8).Move TvwMenu.Left + TvwMenu.Width, 0, ImgBack(8).Width, Me.ScaleHeight
    
    For i = 0 To 7
       ImgBack(i).ZOrder
    Next i
    
    If DisplayMode = 1 Then
       VScroll.ZOrder
       HScroll.ZOrder
    End If
  End With
End If
End Sub

Private Sub ImgButton_Click(Index As Integer)

If Index = 0 Then
   ShowEditor ImgButton(Index).Tag
ElseIf Index = 1 Then
   Call mBackUp_Click
End If

End Sub

Private Sub Label1_Click(Index As Integer)
Call ImgButton_Click(Index)
End Sub


Private Sub mBackUp_Click()
Dim s As Long, q As Boolean
   s = MsgBox(PublicMsgs(17), vbExclamation + vbYesNo + vbDefaultButton2, PublicMsgs(36))

If s = vbYes Then
    q = SetBackUp()
    
    If Not q Then
        MsgBox PublicMsgs(20), vbCritical, PublicMsgs(19)
    Else
        MsgBox ActiveString(PublicMsgs(21), MnBInfo.ModBackUp), vbInformation, PublicMsgs(36)
    End If
End If
End Sub

Private Sub mBanfrmInfo_Click()
Dim t As Integer
   mBanfrmInfo.Checked = Not mBanfrmInfo.Checked
   
   If mBanfrmInfo.Checked Then
      t = 1
   Else
      t = 0
   End If
   
   WriteString MnBInfo.iniSetting, "Settings", "BanfrmInfo", CStr(t)
End Sub

Private Sub mDisplayMode_Click()
mDisplayMode.Checked = Not mDisplayMode.Checked

If mDisplayMode.Checked Then
   DisplayMode = 1
Else
   DisplayMode = 0
End If

InitFormSize

If HaveForm Then
    Display.Move 0, 0
    
    If DisplayMode = 1 Then
     ResetScroll
    End If
End If
'SaveSetting "MnBWarband Editor", "Settings", "DisplayMode", CStr(DisplayMode)
WriteString MnBInfo.iniSetting, "Settings", "DisplayMode", CStr(DisplayMode)
End Sub

Private Sub mHelpTheme_Click()
Dim s As Long
s = ShellExecute(ByVal 0&, vbNullString, "http://bbs.mountblade.com.cn/viewthread.php?tid=168411&extra=page%3D1", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub mIndex_Click()
Dim s As Long, Address As String

Address = "http://bbs.mountblade.com.cn/viewthread.php?tid=173465&extra=page%3D1"

s = ShellExecute(ByVal 0&, vbNullString, Address, vbNullString, vbNullString, vbNormalFocus)

End Sub

Private Sub mInitTranslateForms_Click()
On Error GoTo errorHandle

If MsgBox(PublicMsgs(15), vbInformation + vbYesNo, PublicMsgs(0)) = vbYes Then
   InitTranslateForms
   WritePublicWords
   MsgBox PublicMsgs(16), vbInformation, PublicMsgs(0)
   appExit True
End If

Exit Sub
errorHandle:
    Call logErr("frmMain", "mInitTranslateForms_Click", Err.Number, Err.Description)
End Sub

Private Sub mOutputOperation_Click()
On Error GoTo errorHandle

If MsgBox(PublicMsgs(142), vbInformation + vbYesNo, PublicMsgs(0)) = vbYes Then
   frmTip.ShowTip PublicTips(4)
   OutputOperations App.Path & "\new.op.ini"
   MsgBox PublicMsgs(143), vbInformation, PublicMsgs(0)
   frmTip.HideTip
End If

Exit Sub
errorHandle:
    Call logErr("frmMain", "mOutputOperation_Click", Err.Number, Err.Description)
End Sub

Private Sub mTutorial_Click(Index As Integer)
Dim s As Long, Address As String

Select Case Index
      Case 0
           Address = "http://bbs.mountblade.com.cn/viewthread.php?tid=169246&extra="
      Case 1
           Address = "http://bbs.mountblade.com.cn/viewthread.php?tid=169467&extra="
      Case 2
           Address = "http://bbs.mountblade.com.cn/viewthread.php?tid=170811&extra="
      Case 3
           Address = "http://bbs.mountblade.com.cn/viewthread.php?tid=173207&extra="
      Case 4
           Address = "http://bbs.mountblade.com.cn/viewthread.php?tid=177082&extra="
End Select

s = ShellExecute(ByVal 0&, vbNullString, Address, vbNullString, vbNullString, vbNormalFocus)


End Sub





Private Sub TvwMenu_NodeClick(ByVal node As MSComctlLib.node)
    ShowEditor node.Key
End Sub



Private Sub InitMenu()
On Error GoTo errorHandle

Dim i As Integer

With TvwMenu      '-A_E
    .Nodes.Clear

    .Nodes.Add , , "mEditors", PublicEditors(0), "edit_0"
        .Nodes.Add "mEditors", tvwChild, "edit_1", PublicEditors(1), "edit_1"
        .Nodes.Add "mEditors", tvwChild, "edit_2", PublicEditors(2), "edit_2"
        .Nodes.Add "mEditors", tvwChild, "edit_3", PublicEditors(3), "edit_3"
        .Nodes.Add "mEditors", tvwChild, "edit_4", PublicEditors(4), "edit_4"
        .Nodes.Add "mEditors", tvwChild, "edit_5", PublicEditors(5), "edit_5"
        .Nodes.Add "mEditors", tvwChild, "edit_6", PublicEditors(6), "edit_6"
        .Nodes.Add "mEditors", tvwChild, "edit_7", PublicEditors(7), "edit_7"
        .Nodes.Add "mEditors", tvwChild, "edit_8", PublicEditors(8), "edit_8"
        .Nodes.Add "mEditors", tvwChild, "edit_9", PublicEditors(9), "edit_9"
        .Nodes.Add "mEditors", tvwChild, "edit_10", PublicEditors(10), "edit_10"
        .Nodes.Add "mEditors", tvwChild, "edit_11", PublicEditors(11), "edit_11"
        .Nodes.Add "mEditors", tvwChild, "edit_12", PublicEditors(12), "edit_12"
        .Nodes.Add "mEditors", tvwChild, "edit_13", PublicEditors(13), "edit_13"
        
    .Nodes.Add , , "mTools", PublicTools(0), "tool_0"
        .Nodes.Add "mTools", tvwChild, "tool_1", PublicTools(1), "tool_1"
        .Nodes.Add "mTools", tvwChild, "tool_2", PublicTools(2), "tool_2"
        .Nodes.Add "mTools", tvwChild, "tool_3", PublicTools(3), "tool_3"
        '.Nodes.Add "mTools", tvwChild, "tool_4", PublicTools(4), "tool_4"
        If IsLoadString Then
           .Nodes.Add "mTools", tvwChild, "tool_5", PublicTools(5), "tool_5"
        End If
    .Nodes.Add , , "mHelps", PublicHelp(0), "about_0"
        .Nodes.Add "mHelps", tvwChild, "about_1", PublicHelp(1), "about_1"
End With

'Load Menus  -A_E

For i = 1 To 12
   If i > 0 Then
      Load mEditors(i)
      
      With mEditors(i)
        .Caption = PublicEditors(i)
        .Visible = True
        .Tag = "edit_" & i
      End With

    End If
Next i

mEditors(0).Visible = False

For i = 1 To 5
   If i > 0 Then
      Load mTools(i)
      
      With mTools(i)
        .Caption = PublicTools(i)
        .Visible = True
        .Tag = "tool_" & i
      End With

    End If
Next i

mTools(0).Visible = False

Exit Sub
errorHandle:
    Call logErr("frmMain", "InitMenu", Err.Number, Err.Description)
End Sub

Private Sub ShowEditor(ByVal Tag As String)
On Error GoTo errorHandle

If HaveForm Then
  If Left(Tag, 5) <> "tool_" And Left(Tag, 5) <> "file_" Then
     UnLoad Display
  End If
End If

Select Case Tag         '-A_E
      Case "file_1"
          Call mSelectMod_Click
          Exit Sub
      Case "file_2"
          Call mSave_Click
          Exit Sub
      Case "file_3"
          Call mSaveAs_Click
          Exit Sub
      Case "edit_1"
          Set Display = frmTroops
      Case "edit_2"
          Set Display = frmItems
      Case "edit_3"
          Set Display = frmParties
      Case "edit_4"
          Set Display = frmParty_Templates
      Case "edit_5"
          Set Display = frmFactions
      Case "edit_6"
          Set Display = frmScenes
      Case "edit_7"
          Set Display = frmMap_Icons
      Case "edit_8"
          Set Display = frmPSys
      Case "edit_9"
          Set Display = frmSounds
      Case "edit_10"
          Set Display = frmSoundRess
      Case "edit_11"
          Set Display = frmTabMat
      Case "edit_12"
          Set Display = frmMesh
      Case "edit_13"
          Set Display = frmTrigger
          
      Case "tool_1"
          frmMap.Show
          Exit Sub
      Case "tool_2"
          frmBackUpManager.Show
          Exit Sub
      Case "tool_3"
          frmWizard.wTag = Tag_Item
          frmWizard.Show
          Exit Sub
      Case "tool_4"
          frmCoder.Show
          Exit Sub
      Case "tool_5"
          frmStrTool.Show
          Exit Sub
      Case "about_1"
          Set Display = frmAbout
      Case Else
           Exit Sub
End Select

SetParent Display.hWnd, PicMain.hWnd

Display.Show

Display.Move 0, 0

HaveForm = True

If DisplayMode = 1 Then
   ResetScroll
End If

Exit Sub
errorHandle:
    Call logErr("frmMain", "ShowEditor", Err.Number, Err.Description)
End Sub

Private Sub mAbout_Click()
ShowEditor "about_1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errorHandle

If Not ForceQuit Then
If MsgBox(PublicMsgs(22), vbYesNo + vbExclamation, PublicMsgs(23)) = vbNo Then
    Cancel = True
Else
    End
End If

If HaveForm Then UnLoad Display
UnLoad frmMap

End If

Exit Sub
errorHandle:
    Call logErr("frmMain", "Form_Unload", Err.Number, Err.Description)
End Sub

Private Sub mQuit_Click()
UnLoad Me
End Sub

Public Sub mSave_Click()
On Error GoTo errorHandle

If MsgBox(ActiveString(PublicMsgs(24), MnBInfo.ModName), vbInformation + vbYesNo, PublicMsgs(25)) = vbYes Then
    MsgBox ActiveString(PublicMsgs(27), MnBInfo.ModName), vbInformation, PublicMsgs(25)
    
    'CancelTopForms
    frmTip.ShowTip PublicTips(4), True
    DoEvents
    
    SaveAll
    frmTip.HideTip
    'SetTopForms
        
    MsgBox ActiveString(PublicMsgs(29), MnBInfo.ModName), vbInformation, PublicMsgs(25)
End If

Exit Sub
errorHandle:
    Call logErr("frmMain", "mSave_Click", Err.Number, Err.Description)
End Sub

Private Sub mSaveAs_Click()
On Error GoTo errorHandle

frmSaveAs.Show
frmMain.Enabled = False

Exit Sub
errorHandle:
    Call logErr("frmMain", "mSaveAs_Click", Err.Number, Err.Description)
End Sub

Private Sub mSelectMod_Click()
On Error GoTo errorHandle


If MsgBox(PublicMsgs(33), vbExclamation + vbYesNo + vbDefaultButton2, PublicMsgs(32)) = vbYes Then
If HaveForm Then UnLoad frmFactions

'Unload frmMap
'Welcome.Show

'ForceQuit = True
'Unload Me

appExit True

End If

Exit Sub
errorHandle:
    Call logErr("frmMain", "mSelectMod_Click", Err.Number, Err.Description)
End Sub

Private Sub mEditors_Click(Index As Integer)
ShowEditor mEditors(Index).Tag
End Sub


Private Sub mTools_Click(Index As Integer)
ShowEditor mTools(Index).Tag
End Sub

Private Sub ResetScroll()
Dim DH As Long, dw As Long
DH = Display.ScaleHeight ' / Screen.TwipsPerPixelY
dw = Display.ScaleWidth ' / Screen.TwipsPerPixelX

If DH > PicMain.ScaleHeight Then
     With VScroll
         .Enabled = True
         .Min = 0
         .Max = DH - PicMain.ScaleHeight
         .LargeChange = PicMain.ScaleHeight '((.Max - .Min) / 2)
         .SmallChange = PicMain.ScaleHeight / 2 '(.Max - .Min) / 6
         .Value = 0
     End With
Else
VScroll.Enabled = False

End If

If dw > PicMain.ScaleWidth Then
     With HScroll
         .Enabled = True
         .Min = 0
         .Max = dw - PicMain.ScaleWidth
         .LargeChange = PicMain.ScaleWidth '((.Max - .Min) / 2)
         .SmallChange = PicMain.ScaleWidth / 2 '(.Max - .Min) / 6
         .Value = 0
     End With
Else
HScroll.Enabled = False
End If
End Sub

Private Sub VScroll_Change()
Display.Move 0 - HScroll.Value, 0 - VScroll.Value
End Sub

Private Sub VScroll_Scroll()
Display.Move 0 - HScroll.Value, 0 - VScroll.Value
End Sub


Private Sub HScroll_Change()
Display.Move 0 - HScroll.Value, 0 - VScroll.Value
End Sub

Private Sub HScroll_Scroll()
Display.Move 0 - HScroll.Value, 0 - VScroll.Value
End Sub
