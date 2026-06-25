VERSION 5.00
Begin VB.Form DebugForm 
   Caption         =   "输出窗口"
   ClientHeight    =   5340
   ClientLeft      =   2940
   ClientTop       =   1215
   ClientWidth     =   9510
   Icon            =   "Debugging.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   5340
   ScaleWidth      =   9510
   StartUpPosition =   1  '所有者中心
   Tag             =   "debg_1"
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "DebugForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

