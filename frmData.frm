VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmData 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "索引"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.TreeView TvRegister 
      Height          =   5295
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9340
      _Version        =   393217
      Style           =   5
      Appearance      =   1
   End
   Begin MSComctlLib.ListView LstOp 
      Height          =   5775
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10186
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LstIndex 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10186
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Hide

InitIndexListView
InitLstOp
InitRegister

End Sub
Private Sub InitLstOp()

With LstOp
      .View = lvwReport
      .Sorted = False
      .ListItems.Clear
      .ColumnHeaders.Clear
      .SortOrder = lvwAscending
      .FullRowSelect = True
      .AllowColumnReorder = False
      .LabelEdit = lvwManual
      .Checkboxes = False
      .GridLines = True
      .MultiSelect = False
      .HideSelection = False

      .ColumnHeaders.Add , , "注册序列"
End With

End Sub

Private Sub InitIndexListView()

With LstIndex
      .View = lvwReport
      .Sorted = False
      .ListItems.Clear
      .ColumnHeaders.Clear
      .SortOrder = lvwAscending
      .FullRowSelect = True
      .AllowColumnReorder = False
      .LabelEdit = lvwManual
      .Checkboxes = False
      .GridLines = True
      .MultiSelect = False
      .HideSelection = False

      .ColumnHeaders.Add , , "索引"
      .ColumnHeaders.Add , , "相关索引"
      .ColumnHeaders.Add , , "临时变量名对照表"
End With

End Sub


Private Sub InitRegister()

With TvRegister
      .Nodes.Clear
      .Nodes.Add , , , "Register"
      
End With

End Sub

Private Sub Form_Resize()

With Me
     LstIndex.Move 0, 0, .ScaleWidth, .ScaleHeight
End With

End Sub

