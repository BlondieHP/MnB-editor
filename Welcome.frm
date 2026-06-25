VERSION 5.00
Begin VB.Form Welcome 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "骑马与砍杀:战团 剧本编辑器"
   ClientHeight    =   10410
   ClientLeft      =   5775
   ClientTop       =   2625
   ClientWidth     =   13530
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   13530
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkDisplayMode 
      BackColor       =   &H0018385A&
      Caption         =   "Check1"
      Height          =   255
      Left            =   600
      MouseIcon       =   "Welcome.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   9600
      Width           =   255
   End
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
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6360
      Width           =   4815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "启动可调窗体模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   840
      MouseIcon       =   "Welcome.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   9660
      Width           =   1560
   End
   Begin VB.Label CSelectPath 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置(&E)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   2505
      MouseIcon       =   "Welcome.frx":0EDE
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   8400
      Width           =   1350
   End
   Begin VB.Label CStart 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开始编辑(&S)"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   495
      Left            =   1800
      MouseIcon       =   "Welcome.frx":11E8
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   7560
      Width           =   2760
   End
   Begin VB.Image ExitButton 
      Height          =   480
      Left            =   12560
      MouseIcon       =   "Welcome.frx":14F2
      MousePointer    =   99  'Custom
      Picture         =   "Welcome.frx":17FC
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "选择剧本:"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   3
      Left            =   720
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "战团版本:[Version]"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   2
      Left            =   9540
      TabIndex        =   0
      Top             =   9600
      Width           =   3450
   End
   Begin VB.Image Banner 
      Height          =   10440
      Left            =   0
      Picture         =   "Welcome.frx":24C6
      Top             =   0
      Width           =   13530
   End
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MosP As tPoint

Private Sub Banner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  MosP = SetPoint(X, Y)
  Banner.MousePointer = 5
End If
End Sub

Private Sub Banner_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errorHandle
Dim dP As tPoint

CStart.ForeColor = &H4080&
CSelectPath.ForeColor = &H40C0&

If Button = vbLeftButton Then
   With dP
        .X = X - MosP.X
        .Y = Y - MosP.Y
   End With
   
   Me.Move Me.Left + dP.X, Me.Top + dP.Y
   
End If

Exit Sub
errorHandle:
    Call logErr("Welcome", "Banner_MouseMove", Err.Number, Err.Description)
End Sub

Private Sub Banner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Banner.MousePointer = 0
End Sub

Private Sub chkDisplayMode_Click()
On Error GoTo errorHandle

DisplayMode = chkDisplayMode.Value
'SaveSetting "MnBWarband Editor", "Settings", "DisplayMode", CStr(DisplayMode)
WriteString MnBInfo.iniSetting, "Settings", "DisplayMode", CStr(DisplayMode)

Exit Sub

errorHandle:
    Call logErr("Welcome", "chkDisplayMode_Click", Err.Number, Err.Description)
End Sub

Private Sub CStart_Click()
On Error GoTo errorHandle

If Combo1.Text <> "" Then
FinishWarbandInfo Combo1.Text
DisplayMode = chkDisplayMode.Value

frmMain.Show

UnLoad Me
End If

Exit Sub
errorHandle:
    Call logErr("Welcome", "CStart_Click", Err.Number, Err.Description)
End Sub

Private Sub CStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errorHandle

CStart.ForeColor = &H80FF&

Exit Sub
errorHandle:
    Call logErr("Welcome", "CStart_MouseMove", Err.Number, Err.Description)
End Sub

Private Sub CSelectPath_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errorHandle

CSelectPath.ForeColor = &H80FF&
Exit Sub
errorHandle:
    Call logErr("Welcome", "CSelectPath_MouseMove", Err.Number, Err.Description)
End Sub

Private Sub ExitButton_Click()
End
End Sub

Private Sub Form_Load()
On Error GoTo errorHandle
Dim q As Boolean, strTem As String

'strTem = GetSetting("MnBWarband Editor", "Settings", "DisplayMode", "0")

PublicInit
q = InitWarbandInfo

TranslatePublicWords
TranslateForm Me

If Not q Then
   Call CSelectPath_Click
End If

If q Then
    DisplayWarbandInfo
End If

strTem = ReadString(MnBInfo.iniSetting, "Settings", "DisplayMode", 250)

DisplayMode = Val(strTem)

If DisplayMode = 0 Or DisplayMode = 1 Then
chkDisplayMode.Value = DisplayMode
End If

strTem = ReadString(MnBInfo.iniSetting, "Settings", "IsLoadString", 250)

If UCase(strTem) = "TRUE" Then
  IsLoadString = True
Else
  IsLoadString = False
End If
'frmSelectPath.Check1.Value = IsLoadString

Me.Move Me.Left, Me.Top, Banner.Width, Banner.Height

Me.Show

Exit Sub
errorHandle:
    Call logErr("Welcome", "Form_Load", Err.Number, Err.Description)
End Sub

Public Sub DisplayWarbandInfo()
On Error GoTo errorHandle

Dim strTem As String

strTem = IIf(Len(MnBInfo.Version) > 0, ":" & Format(Val(MnBInfo.Version) / 1000, "0.000"), PublicMsgs(54))
Label1(2).Caption = Replace(Label1(2).Caption, ":[Version]", strTem, , , vbTextCompare)

LoadModules

Exit Sub

errorHandle:
    Call logErr("Welcome", "DisplayWarbandInfo", Err.Number, Err.Description)
End Sub

Private Sub LoadModules()
On Error GoTo errorHandle
Dim s As String

Combo1.Clear

With MnBInfo
    s = Dir(.ModPath & "\", vbDirectory)
    Do While s <> ""
        If IsDirectory(.ModPath & "\" & s) Then   '排除目录首
           Combo1.AddItem s
        End If
        s = Dir
    Loop

End With

Exit Sub

errorHandle:
    Call logErr("Welcome", "LoadModules", Err.Number, Err.Description)
End Sub


Private Sub CSelectPath_Click()
On Error GoTo errorHandle
Me.Enabled = False

frmSelectPath.Show

Exit Sub

errorHandle:
    Call logErr("Welcome", "CSelectPath_Click", Err.Number, Err.Description)
End Sub

Private Sub Label2_Click()
If chkDisplayMode.Value = 1 Then
   chkDisplayMode.Value = 0
ElseIf chkDisplayMode.Value = 0 Then
   chkDisplayMode.Value = 1
End If

Call chkDisplayMode_Click
End Sub
