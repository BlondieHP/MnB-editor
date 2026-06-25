VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IDSysExam"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9255
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton bstrID 
      Caption         =   "CHANGE STRID"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   6360
      Width           =   2535
   End
   Begin VB.CommandButton bID 
      Caption         =   "CHANGE ID"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   6360
      Width           =   2535
   End
   Begin VB.TextBox txtID 
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton bDel 
      Caption         =   "DEl"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton bAdd 
      Caption         =   "ADD"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox txtstrID 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5880
      Width           =   7695
   End
   Begin MSComctlLib.ListView LstIndex 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
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
Option Explicit

Private Sub bAdd_Click()
Me.Caption = AddIndex(Val(txtID.Text), txtstrID.Text)
End Sub

Private Sub bDel_Click()
Me.Caption = DelIndex(KeytoStrID(LstIndex.SelectedItem.Key), True)
End Sub

Private Sub bID_Click()
Me.Caption = ChangeID(KeytoStrID(LstIndex.SelectedItem.Key), Val(txtID.Text))
End Sub

Private Sub bstrID_Click()
Me.Caption = ChangeStrID(KeytoStrID(LstIndex.SelectedItem.Key), txtstrID.Text)
End Sub

Private Sub Form_Load()

InitIndexListView

End Sub


Private Sub InitIndexListView()
Dim n As Integer
n = 3
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

      .ColumnHeaders.Add , , "索引", .Width / 3
      .ColumnHeaders.Add , , "相关索引", .Width / 2
End With

End Sub

Private Sub LstIndex_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtstrID.Text = KeytoStrID(Item.Key)
txtID.Text = GetID(txtstrID.Text)
End Sub

Private Sub LstIndex_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
LstIndex.ToolTipText = LstIndex.SelectedItem.Key
End Sub
