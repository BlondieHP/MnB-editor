VERSION 5.00
Begin VB.Form frmBackUpManager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "备份管理器"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   Icon            =   "frmBackUpManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8970
   StartUpPosition =   1  '所有者中心
   Tag             =   "tool_2"
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   3960
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   1140
      Left            =   3120
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1170
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox LstFiles 
      Height          =   3480
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.ListBox LstBacks 
      Height          =   5100
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "更名(&N)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   1920
      MouseIcon       =   "frmBackUpManager.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*留空即会以时间作为文件名"
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
      Left            =   4440
      TabIndex        =   12
      Top             =   4440
      Width           =   2445
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备份名:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   2
      Left            =   4440
      TabIndex        =   11
      Top             =   4080
      Width           =   690
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "反选(&I)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   2
      Left            =   1080
      MouseIcon       =   "frmBackUpManager.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "删除(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   3600
      MouseIcon       =   "frmBackUpManager.frx":0EDE
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "创建(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   180
      Index           =   1
      Left            =   2760
      MouseIcon       =   "frmBackUpManager.frx":11E8
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label COK 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "还原(&R)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   5640
      MouseIcon       =   "frmBackUpManager.frx":14F2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4920
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "备份的文件:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "还原点:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   5895
      Left            =   -120
      Picture         =   "frmBackUpManager.frx":17FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9255
   End
End
Attribute VB_Name = "frmBackUpManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cCMD_Click(Index As Integer)
On Error GoTo EL
Dim s As Long, q As Boolean, i As Integer

If Index = 0 Then
s = MsgBox(PublicMsgs(35), vbYesNo + vbDefaultButton2 + vbExclamation, PublicMsgs(34))

If s = vbYes Then
    For i = LstBacks.ListCount - 1 To 0 Step -1
        If LstBacks.Selected(i) Then
            DeleteFolder MnBInfo.ModBackUp & "\" & LstBacks.List(i)
        End If
    Next i
    
    InitBackUpList
End If

ElseIf Index = 1 Then
s = MsgBox(PublicMsgs(37), vbExclamation + vbYesNo + vbDefaultButton2, PublicMsgs(36))

If s = vbYes Then
    q = SetBackUp(txtName.Text)
    
    If Not q Then
        MsgBox PublicMsgs(20), vbCritical, PublicMsgs(19)
    Else
        InitBackUpList
        MsgBox ActiveString(PublicMsgs(21), MnBInfo.ModBackUp), vbInformation, PublicMsgs(36)
    End If
End If

ElseIf Index = 2 Then
    For i = 0 To LstBacks.ListCount - 1
        LstBacks.Selected(i) = Not LstBacks.Selected(i)
    Next i

ElseIf Index = 3 Then
    s = MsgBox(ActiveString(PublicMsgs(145), "[" & LstBacks.List(LstBacks.ListIndex) & "]"), vbExclamation + vbYesNo, PublicMsgs(0))
    If s = vbNo Then Exit Sub
    Name MnBInfo.ModBackUp & "\" & LstBacks.List(LstBacks.ListIndex) As MnBInfo.ModBackUp & "\" & txtName.Text
    InitBackUpList
    
    MsgBox ActiveString(PublicMsgs(146), "[" & txtName.Text & "]"), vbInformation, PublicMsgs(0)
End If

Exit Sub

EL:
    Call logErr("frmBackUpManager", "cCMD_Click(" & Index & ")", Err.Number, Err.Description)
End Sub

Private Sub COK_Click()
On Error GoTo EL
Dim s As Long, q As Boolean

If LstBacks.ListIndex > -1 Then
s = MsgBox(ActiveString(PublicMsgs(41), LstBacks.List(LstBacks.ListIndex)), vbExclamation + vbYesNo + vbDefaultButton2, PublicMsgs(40))
If s = vbYes Then
    q = RestoreMod(LstBacks.List(LstBacks.ListIndex))
    
    If Not q Then
        MsgBox PublicMsgs(44), vbCritical, PublicMsgs(40)     '失败
    Else
        MsgBox PublicMsgs(43), vbInformation, PublicMsgs(40)   '成功
        PublicInit
        ReadAll
    End If
End If

End If

Exit Sub

EL:
    Call logErr("frmBackUpManager", "COK_Click", Err.Number, Err.Description)
End Sub

Private Sub Form_Load()
Dim i As Integer

Image1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight

InitBackUpList

TranslateForm Me

InitCMDs

Me.Show
Call LstBacks_Click
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

Private Sub LoadFileList(ByVal RevTime As String)
On Error GoTo EL

Dim i As Integer
LstFiles.Clear

If DirExists(MnBInfo.ModBackUp & "\" & RevTime) Then

File1.Path = MnBInfo.ModBackUp & "\" & RevTime
File1.Pattern = "*.txt;*.csv"
File1.Refresh
With LstFiles
     For i = 0 To File1.ListCount - 1
          .AddItem File1.List(i)
     Next i
End With


End If

Exit Sub

EL:
    Call logErr("frmBackUpManager", "LoadFileList", Err.Number, Err.Description)
End Sub

Private Sub LstBacks_Click()
On Error GoTo EL
If LstBacks.ListIndex > -1 Then
    LoadFileList LstBacks.List(LstBacks.ListIndex)
    txtName.Text = LstBacks.List(LstBacks.ListIndex)
End If

Exit Sub

EL:
    Call logErr("frmBackUpManager", "LstBacks_Click", Err.Number, Err.Description)
End Sub

Private Sub InitBackUpList()
On Error GoTo EL

Dim i As Integer, strTem As String, n As Integer
LstBacks.Clear

If Not DirExists(MnBInfo.ModBackUp) Then
    MkDirEx MnBInfo.ModBackUp
End If

Dir1.Path = MnBInfo.ModBackUp
Dir1.Refresh
For i = 0 To Dir1.ListCount - 1
    n = InStrRev(Dir1.List(i), "\")
    
    If n > 0 Then
       strTem = Right(Dir1.List(i), Len(Dir1.List(i)) - n)
    Else
       strTem = Dir1.List(i)
    End If
    
    LstBacks.AddItem strTem
Next i

Exit Sub

EL:
    Call logErr("frmBackUpManager", "InitBackUpList", Err.Number, Err.Description)
End Sub
