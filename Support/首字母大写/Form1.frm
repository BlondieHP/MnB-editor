VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "首字母大写"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5805
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "清空(C)"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "转换(C)"
      Default         =   -1  'True
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   960
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "分隔符:"
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   200
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim strVar() As String, i As Integer, k As String, strTem As String

strVar() = Split(Text2.Text, Text1.Text)

For i = 0 To UBound(strVar)
    If Len(strVar(i)) > 0 Then
       k = UCase(Left(strVar(i), 1))
       strVar(i) = k & LCase(Right(strVar(i), Len(strVar(i)) - 1))
       strTem = strTem & strVar(i) & Text1.Text
    End If
Next i

If Len(strTem) > 0 Then
   strTem = Left(strTem, Len(strTem) - 1)
End If

Text2.Text = strTem
End Sub

Private Sub Command2_Click()
Text2.Text = ""
End Sub

