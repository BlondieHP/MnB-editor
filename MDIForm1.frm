VERSION 5.00
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "骑马与砍杀:战团 剧本编辑器"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10575
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
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
      WindowList      =   -1  'True
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
      Begin VB.Menu mDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "关于MnBWarband Editor(&A)"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub MDIForm_Load()
Dim i As Integer

Me.Show

Welcome.Show

'Load Menus

For i = 1 To 7
   If i > 0 Then
      Load mEditors(i)
      
      With mEditors(i)
        Select Case i
            Case 1
               .Caption = "兵种(&T)"
            Case 2
               .Caption = "物品(&I)"
            Case 3
               .Caption = "部队(&P)"
            Case 4
               .Caption = "部队模板(&M)"
            Case 5
               .Caption = "阵营(&F)"
            Case 6
               .Caption = "场景(&S)"
            Case 7
               .Caption = "大地图图标(&C)"
        End Select
        
        .Visible = True
        .Tag = "edit_" & i
      End With

    End If
Next i

mEditors(0).Visible = False

For i = 1 To 1
   If i > 0 Then
      Load mTools(i)
      
      With mTools(i)
        Select Case i
            Case 1
               .Caption = "卡拉迪亚地图(&C)"
        End Select
        
        .Visible = True
        .Tag = "tool_" & i
      End With

    End If
Next i

mTools(0).Visible = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("确定退出 【骑马与砍杀:战团】剧本编辑器 吗?", vbYesNo + vbExclamation, "确认退出") = vbNo Then
    Cancel = True
End If
End Sub

Private Sub mEditors_Click(Index As Integer)
Select Case mEditors(Index).Tag
      Case "edit_1"
          frmTroops.Show
      Case "edit_2"
          frmItems.Show
      Case "edit_3"
          frmParties.Show
      Case "edit_4"
          frmParty_Templates.Show
      Case "edit_5"
          frmFactions.Show
      Case "edit_6"
          frmScenes.Show
      Case "edit_7"
          frmMap_Icons.Show
End Select
End Sub

Private Sub mQuit_Click()
Unload Me
End Sub

Public Sub mSave_Click()

If MsgBox("确定按现在设定保存剧本[" & MnBInfo.ModName & "]?", vbInformation + vbYesNo, "保存剧本") = vbYes Then
    MsgBox "请确认备份剧本[" & MnBInfo.ModName & "]后,按“确定”开始保存。", vbInformation, "保存剧本"
    
    frmTip.ShowTip "保存中,请稍后..."
    DoEvents
    
    SaveAll
    frmTip.HideTip
        
    MsgBox "剧本[" & MnBInfo.ModName & "]保存成功!", vbInformation, "保存剧本"
End If
End Sub

Private Sub mSaveAs_Click()
frmSaveAs.Show
End Sub

Private Sub mSelectMod_Click()

Unload frmFactions
Unload frmTroops
Unload frmParties
Unload frmParty_Templates
Unload frmItems
Unload frmScenes

Unload frmMap

Unload frmModule
Unload frmSaveAs

Welcome.Show
End Sub



Private Sub mTools_Click(Index As Integer)
Select Case mTools(Index).Tag
      Case "tool_1"
          frmMap.Show

End Select
End Sub
