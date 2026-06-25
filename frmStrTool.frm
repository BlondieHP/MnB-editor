VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStrTool 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "字符串管理工具"
   ClientHeight    =   8715
   ClientLeft      =   3360
   ClientTop       =   1395
   ClientWidth     =   12360
   Icon            =   "frmStrTool.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12360
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   5450
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "语言信息"
      Height          =   3735
      Left            =   5760
      TabIndex        =   16
      Top             =   4320
      Width           =   6375
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Height          =   3135
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "frmStrTool.frx":0E42
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "字符串内容"
      Height          =   3375
      Left            =   5760
      TabIndex        =   14
      Top             =   720
      Width           =   6375
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Index           =   1
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Text            =   "frmStrTool.frx":0E48
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmStrTool.frx":0E4E
      Top             =   240
      Width           =   5895
   End
   Begin VB.CommandButton CReset 
      BackColor       =   &H000000FF&
      Caption         =   "重置(&R)"
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
      Left            =   7680
      TabIndex        =   11
      Top             =   8160
      Width           =   2175
   End
   Begin VB.CommandButton CApply 
      BackColor       =   &H000000FF&
      Caption         =   "套用(&A)"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8160
      Width           =   2175
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   3975
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
      Left            =   5010
      TabIndex        =   1
      Top             =   135
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
      Left            =   4650
      TabIndex        =   0
      Top             =   135
      Width           =   375
   End
   Begin MSComctlLib.ListView LstStrs 
      Height          =   7695
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   13573
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmStrTool.frx":0E56
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   13
      Top             =   240
      Width           =   375
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
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "字符串数:"
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
      TabIndex        =   8
      Top             =   8415
      Width           =   885
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
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   1
      Left            =   3840
      MouseIcon       =   "frmStrTool.frx":54C3
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   8415
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
      Left            =   4680
      MouseIcon       =   "frmStrTool.frx":57CD
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   8415
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "上移(&U)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   2
      Left            =   2160
      MouseIcon       =   "frmStrTool.frx":5AD7
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   8415
      Width           =   705
   End
   Begin VB.Label cCMD 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "下移(&D)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   3
      Left            =   3000
      MouseIcon       =   "frmStrTool.frx":5DE1
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   8415
      Width           =   705
   End
End
Attribute VB_Name = "frmStrTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OriginFormWidth As Long
Dim CustomActive As Boolean

Private Sub CApply_Click()
If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
   'frmTip.ShowTip PublicMsgs(134)
   'Timer1.Enabled = True
   Strs(CurrentStrID) = CurrentStr
   LstStrs.ListItems(CurrentStrID + 1).SubItems(1) = CurrentStr.Name
   If LCase$(MnBInfo.Language) = "en" Then
       LstStrs.ListItems(CurrentStrID + 1).SubItems(2) = CurrentStr.Str
   Else
      LstStrs.ListItems(CurrentStrID + 1).SubItems(2) = CurrentStr.CSV
   End If
End If
End Sub

Private Sub Command1_Click()

With Command1
   If .Caption = "<" Then
      Me.Width = .Left + .Width + 100
      .Caption = ">"
   Else
      Me.Width = OriginFormWidth
      .Caption = "<"
   End If
End With

End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstStrs, LstStrs.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstStrs_ItemClick(LstStrs.SelectedItem)
End If
End Sub

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
   LstStrs_ItemClick LstStrs.ListItems(CurrentStrID + 1)
   LstStrs.ListItems(CurrentStrID + 1).Selected = True
   LstStrs.ListItems(CurrentStrID + 1).EnsureVisible
End If
End Sub

Private Sub Form_LostFocus()
'If Timer1.Enabled Or isShowTip Then
'   SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3
'End If
End Sub
Private Sub Form_Deactivate()
  'If Not Timer1.Enabled And Not isShowTip Then
  '     SetWindowPos Me.hWnd, 0, 0, 0, 0, 0, 3
  'End If
  'If Timer1.Enabled Or isShowTip Then
      'SetWindowPos Me.hWnd, -2, 0, 0, 0, 0, 3
  'End If
End Sub

Private Sub Form_Load()
OriginFormWidth = Me.Width
CustomActive = True

If LCase$(MnBInfo.Language) = "en" Then
     Frame1.Height = Frame2.Top + Frame2.Height - Frame1.Top
     Dim D As Long
     D = 100
     Text1(1).Height = Frame1.Height - Text1(1).Top - D
     Frame2.Visible = False
End If

InitItemsListView
Label2.Caption = Label2.Caption & N_Str
LoadLstStrs

LstStrs.ListItems(1).Selected = True
CurrentStrID = 0
CurrentStr = Strs(0)
Call LoadString(CurrentStr)

SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, 3
'CancelTopForms
End Sub

Private Sub InitItemsListView()
Dim n As Integer
n = 3

With LstStrs
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

   .ColumnHeaders.Add , , PublicMsgs(13), .Width / n / 3.5
   .ColumnHeaders.Add , , PublicTags(Tag_String) & PublicMsgs(14), .Width / n * 1
   .ColumnHeaders.Add , , PublicTags(Tag_String) & PublicMsgs(133), .Width / n * 2
End With

End Sub

'*************************************************************************
'**函 数 名：LoadLstStrs
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-03-03 14:26:51
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadLstStrs()
Dim n As Long, oItem As ListItem

LstStrs.ListItems.Clear  '清空列表

For n = 0 To UBound(Strs)
    Set oItem = LstStrs.ListItems.Add(, "str_" & CStr(n), Strs(n).ID)
    With oItem
       .SubItems(1) = Strs(n).Name
       If LCase$(MnBInfo.Language) = "en" Then
          .SubItems(2) = Strs(n).Str
       Else
          .SubItems(2) = Strs(n).CSV
       End If
    End With
Next n

End Sub

'*************************************************************************
'**函 数 名：QueryItem
'**输    入：(ListItem)oLV,(Long)Start,(String)QueryString,(Boolean)bReverse
'**输    出：
'**功能描述：进行ListView查询功能
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-11-27 21:19:17
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************

Private Function QueryItem(oLV As ListView, ByVal Start As Long, ByVal QueryString As String, Optional bReverse As Boolean = False) As Boolean
Dim oItem As ListItem
  With oLV
    Set oItem = FindItem(oLV, Start, "0|1|2", QueryString, True, vbTextCompare, bReverse)
       If Not oItem Is Nothing Then
       .ListItems(oItem.Index).Selected = True
       .ListItems(oItem.Index).EnsureVisible
       QueryItem = True
       Else
        MsgBox PublicMsgs(11), vbInformation, PublicMsgs(12)
        QueryItem = False
       End If
  End With
    Set oItem = Nothing
End Function


Private Sub LstStrs_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim n As Integer

CurrentStrID = Val(Item.Text)
CurrentStr = Strs(CurrentStrID)

LoadString CurrentStr

End Sub

'*************************************************************************
'**函 数 名：LoadString
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2012-03-03 14:55:34
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadString(Str As Type_String)

With Str
   Text1(0).Text = .Name
   Text1(1).Text = .Str
   Text1(2).Text = .CSV
End With

End Sub

Private Sub Text1_LostFocus(Index As Integer)
On Error GoTo EL
With CurrentStr

If CustomActive Then

If Index <= 1 Then
   Text1(Index).Text = Replace(Text1(Index).Text, " ", "_")
Else
   Text1(Index).Text = Replace(Text1(Index).Text, "_", " ")
End If

Select Case Index
      Case 0
           'Check Value
            If Left$(Text1(Index).Text, 4) <> "str_" Then
                Text1(Index).Text = "str_" & Text1(Index).Text
            End If
         .Name = Text1(Index).Text
      Case 1
         .Str = Text1(Index).Text
           'If LCase$(MnBInfo.Language) = "en" Then
           '  .CSV = Replace(.Str, "_", " ")
           'End If
      Case 2
         .CSV = Text1(Index).Text
End Select

End If
End With

Exit Sub
EL:
  Call logErr("frmStrTool", "Text1_LostFocus(" & Index & ")", Err.Number, Err.Description)
End Sub

Private Sub Timer1_Timer()
frmTip.HideTip
Timer1.Enabled = False
End Sub
