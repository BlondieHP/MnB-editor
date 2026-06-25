VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MnBWarband Editor控件注册机"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6495
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "退出(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton CReg 
      BackColor       =   &H000000FF&
      Caption         =   "注册(&R)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox LstCtls 
      Height          =   3840
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注册信息:"
      Height          =   180
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   4080
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "注册信息:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const TYPE_DLL = 0
Const TYPE_OCX = 1

Private Sub InitLstCtls()
With LstCtls
     .Clear
     
     .AddItem "MSSTDFMT.DLL"
       .ItemData(.ListCount - 1) = TYPE_DLL
       .Selected(.ListCount - 1) = Not CheckReg(.List(.ListCount - 1))
     .AddItem "COMDLG32.OCX"
       .ItemData(.ListCount - 1) = TYPE_OCX
       .Selected(.ListCount - 1) = Not CheckReg(.List(.ListCount - 1))
     .AddItem "MSCOMCTL.OCX"
       .ItemData(.ListCount - 1) = TYPE_OCX
       .Selected(.ListCount - 1) = Not CheckReg(.List(.ListCount - 1))
End With

LoadCtlInfo "MSSTDFMT.DLL"
End Sub

Private Sub CExit_Click()
End
End Sub

Private Sub CReg_Click()
On Error Resume Next

Dim i As Integer, F As Integer, Data() As Byte, TemFile As String

For i = 0 To LstCtls.ListCount - 1
     If LstCtls.Selected(i) Then
         F = FreeFile
         Data = LoadResData(LstCtls.List(i), GetType(LstCtls.ItemData(i)))
         
         TemFile = Environ("SystemRoot") & "\system32\" & LstCtls.List(i)
         Open TemFile For Binary As #F
             Put #F, , Data
         Close #F
         
         Shell "regsvr32 " & Chr(34) & TemFile & Chr(34)
     End If
Next i
End Sub

Private Sub Form_Load()

InitLstCtls

End Sub

Private Sub LstCtls_Click()
If LstCtls.ListIndex > -1 Then
   LoadCtlInfo LstCtls.List(LstCtls.ListIndex)
End If
End Sub

Private Sub LoadCtlInfo(ByVal CtlName As String)
Dim s As String
s = Dir(Environ("SystemRoot") & "\system32\" & CtlName)

If Trim(s) = "" Or Trim(s) = "." Or Trim(s) = ".." Then
    Label1(1).Caption = "控件未注册"
Else
    Label1(1).Caption = "控件可能已注册"
End If

End Sub

Private Function CheckReg(ByVal CtlName As String) As Boolean
Dim s As String
s = Dir(Environ("SystemRoot") & "\system32\" & CtlName)

If Trim(s) = "" Or Trim(s) = "." Or Trim(s) = ".." Then
    CheckReg = False
Else
    CheckReg = True
End If
End Function

Private Function GetType(ByVal lngType As Long) As String
Select Case lngType
     Case TYPE_DLL
         GetType = "DLL"
     Case TYPE_OCX
         GetType = "OCX"
End Select
End Function
