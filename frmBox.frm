VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择框"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Tag             =   "debg_3"
   Begin VB.CommandButton CCancel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "取消(&C)"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton COK 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定(&O)"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   2295
   End
   Begin MSComctlLib.ListView LstBox 
      Height          =   6135
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10821
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton CQuery 
      Caption         =   "↑"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton CQuery 
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "查询:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   195
      Width           =   495
   End
End
Attribute VB_Name = "frmBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

InitListBox

Select Case QuestTag
       Case Tag_Troop
            FillBox_Troops
       Case Tag_Faction
            FillBox_Factions
End Select

End Sub

Private Sub InitListBox()
Dim n As Integer
n = 3

With LstFactions
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

      .ColumnHeaders.Add , , "序列", .Width / n / 3.6
      .ColumnHeaders.Add , , "识别符", .Width / n * 1.5
      .ColumnHeaders.Add , , "名称(" & MnBInfo.Language & ")", .Width / n * 1.5
End With

End Sub

Private Sub FillBox_Troops()
Dim i As Long
Dim oItem As ListItem

For i = 0 To N_Troop - 1
   With Trps(i)
     Set oItem = LstBox.ListItems.Add(, , .ID)
         oItem.SubItems(1) = .strID
         oItem.SubItems(2) = .csvName
   End With
Next i

End Sub

Private Sub FillBox_Factions()
Dim i As Long
Dim oItem As ListItem

For i = 0 To N_Faction - 1
   With Factions(i)
     Set oItem = LstBox.ListItems.Add(, , .ID)
         oItem.SubItems(1) = .strID
         oItem.SubItems(2) = .csvName
   End With
Next i

End Sub
