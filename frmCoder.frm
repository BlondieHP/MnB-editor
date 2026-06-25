VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCoder 
   Caption         =   "操作快编写工具"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10995
   Icon            =   "frmCoder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   StartUpPosition =   1  '所有者中心
   Tag             =   "tool_4"
   Begin MnBWarband_Editor.TriggersEditor TriggersEditor1 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10575
      _ExtentX        =   13785
      _ExtentY        =   8281
   End
   Begin MSComctlLib.ListView lstTips 
      Height          =   1695
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2990
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.PictureBox PicTip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Menu mFile 
      Caption         =   "文件(&F)"
   End
   Begin VB.Menu mEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mCode 
         Caption         =   "编码(&C)"
      End
   End
End
Attribute VB_Name = "frmCoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CustomActive As Boolean
Dim TrgCnt As New clsTriggersEditor

Private Sub Form_Load()
Dim n As ListItemforMS, q As Boolean

Me.Show
TrgCnt.Initialize TriggersEditor1
TrgCnt.InputTrg "itm_heraldic_mail_with_surcoat_for_tableau", "itm", itm(598).Trigger()

End Sub

Private Sub Form_Resize()
On Error Resume Next

TriggersEditor1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub mFile_Click()
TrgCnt.OutputTrg itm(598).Trigger()

End Sub

