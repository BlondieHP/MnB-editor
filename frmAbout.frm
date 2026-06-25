VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "关于MnBWarband Editor"
   ClientHeight    =   9375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14355
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Tag             =   "about_1"
   Begin VB.TextBox txtAbout 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Index           =   1
      Left            =   6360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmAbout.frx":08CA
      Top             =   1440
      Width           =   7695
   End
   Begin VB.TextBox txtAbout 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   4095
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmAbout.frx":08D6
      Top             =   5040
      Width           =   5775
   End
   Begin VB.Image Banner 
      Height          =   1260
      Index           =   1
      Left            =   6240
      MouseIcon       =   "frmAbout.frx":08E2
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":0BEC
      Top             =   7920
      Width           =   3750
   End
   Begin VB.Image Banner 
      Height          =   1125
      Index           =   0
      Left            =   10320
      MouseIcon       =   "frmAbout.frx":2286
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":2590
      Top             =   8040
      Width           =   3750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.0.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   11280
      TabIndex        =   3
      Top             =   720
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   3270
      Left            =   600
      Picture         =   "frmAbout.frx":5152
      Top             =   1440
      Width           =   4995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MnBWarband Editor"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   42
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1185
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   7980
   End
   Begin VB.Image ImgBanner 
      Height          =   480
      Left            =   2040
      Picture         =   "frmAbout.frx":FF58
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Banner_Click(Index As Integer)
Dim s As Long, Address As String

On Error GoTo Errline
Select Case Index
     Case 0
        'Shell "explorer " & Chr(34) & "http://bbs.mountblade.com.cn/" & Chr(34)
        Address = "http://bbs.mountblade.com.cn/"
     Case 1
        'Shell "explorer " & Chr(34) & "http://tieba.baidu.com/f?kw=%C6%EF%C2%ED%D3%EB%BF%B3%C9%B1" & Chr(34)
        Address = "http://tieba.baidu.com/f?kw=%C6%EF%C2%ED%D3%EB%BF%B3%C9%B1"
End Select
s = ShellExecute(ByVal 0&, vbNullString, Address, vbNullString, vbNullString, vbNormalFocus)

If s = 0 Then
     Call logErr("frmAbout", "Banner_Click", "INVALID_LINK", "链接无效")
End If

Exit Sub

Errline:
    Call logErr("frmAbout", "Banner_Click", Err.Number, Err.Description)
End Sub

Private Sub Form_Load()
txtAbout(0).Text = WriteAboutString(0)
txtAbout(1).Text = WriteAboutString(1)
Label2.Caption = App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf
Me.Show
End Sub


'*************************************************************************
'**函 数 名：WriteAboutString
'**输    入：(Integer)Index
'**输    出：(String)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-13 21:52:42
'**版    本：V1.1321
'*************************************************************************
Public Function WriteAboutString(ByVal Index As Integer) As String
Dim strTem As String, i As Integer

With App

Select Case Index

Case 0
strTem = strTem & PublicAbout(0) & vbCrLf & vbCrLf
strTem = strTem & "   " & PublicAbout(1) & vbCrLf
strTem = strTem & "   " & PublicAbout(2) & vbCrLf & vbCrLf

strTem = strTem & PublicAbout(3) & vbCrLf & vbCrLf

For i = 24 To UBound(PublicAbout)
strTem = strTem & "   " & PublicAbout(i) & vbCrLf
Next i

strTem = strTem & vbCrLf

Case 1
strTem = strTem & PublicAbout(4) & vbCrLf & vbCrLf
strTem = strTem & "    " & PublicAbout(5)
strTem = strTem & PublicAbout(6) & PublicAbout(7)
strTem = strTem & PublicAbout(8) & PublicAbout(9) & vbCrLf & vbCrLf

strTem = strTem & PublicAbout(10) & vbCrLf & vbCrLf
strTem = strTem & "    " & PublicAbout(11) & vbCrLf & vbCrLf

strTem = strTem & PublicAbout(12) & vbCrLf & vbCrLf
strTem = strTem & "    " & PublicAbout(13) & vbCrLf & vbCrLf

strTem = strTem & PublicAbout(14) & vbCrLf & vbCrLf
strTem = strTem & "    " & PublicAbout(15) & vbCrLf
strTem = strTem & "    " & PublicAbout(16) & vbCrLf
strTem = strTem & "    " & PublicAbout(17) & vbCrLf & vbCrLf

strTem = strTem & PublicAbout(17) & vbCrLf & vbCrLf
strTem = strTem & PublicAbout(18) & vbCrLf
strTem = strTem & PublicAbout(19) & vbCrLf & vbCrLf
strTem = strTem & PublicAbout(20) & vbCrLf
strTem = strTem & PublicAbout(21) & vbCrLf & vbCrLf
strTem = strTem & PublicAbout(22) & vbCrLf & vbCrLf

strTem = strTem & PublicAbout(23)
End Select

End With

WriteAboutString = strTem
End Function

