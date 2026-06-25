VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFactions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "阵营编辑器"
   ClientHeight    =   9375
   ClientLeft      =   1755
   ClientTop       =   180
   ClientWidth     =   14355
   Icon            =   "frmFactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_5"
   Begin VB.Frame FraProps 
      Caption         =   "阵营属性"
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
      Height          =   8535
      Index           =   0
      Left            =   6120
      TabIndex        =   13
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame1 
         Caption         =   "特性"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1095
         Index           =   2
         Left            =   240
         TabIndex        =   28
         Top             =   7080
         Width           =   7575
         Begin VB.TextBox txtRating 
            Height          =   375
            Left            =   5640
            TabIndex        =   30
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkFlags 
            Caption         =   "chkFlags"
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "等级:"
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
            Index           =   5
            Left            =   5040
            TabIndex        =   31
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "关系"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   4215
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   7575
         Begin VB.TextBox txtRelationShip 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   6240
            TabIndex        =   27
            Top             =   3720
            Width           =   1095
         End
         Begin MSComctlLib.ListView LstRelationShips 
            Height          =   3375
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   5953
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*关系范围(-1.0~1.0),实际游戏中*100"
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
            Index           =   6
            Left            =   120
            TabIndex        =   32
            Top             =   3840
            Width           =   3435
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "关系:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   4
            Left            =   5640
            TabIndex        =   26
            Top             =   3840
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "识别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2295
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   7575
         Begin VB.PictureBox PicColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   4440
            ScaleHeight     =   345
            ScaleWidth      =   2625
            TabIndex        =   24
            Top             =   1680
            Width           =   2655
         End
         Begin MSComDlg.CommonDialog CD 
            Left            =   6960
            Top             =   1560
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            CancelError     =   -1  'True
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   3
            Left            =   1800
            TabIndex        =   25
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   17
            Top             =   240
            Width           =   5295
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   16
            Top             =   720
            Width           =   5295
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   2
            Left            =   1800
            TabIndex        =   15
            Top             =   1200
            Width           =   5295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阵营颜色:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   2
            Left            =   810
            TabIndex        =   23
            Top             =   1760
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阵营ID:"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   180
            Index           =   0
            Left            =   960
            TabIndex        =   20
            Top             =   315
            Width           =   705
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阵营名(EN):"
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
            Left            =   600
            TabIndex        =   19
            Top             =   795
            Width           =   1110
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阵营名(NOW):"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Index           =   3
            Left            =   480
            TabIndex        =   18
            Top             =   1275
            Width           =   1215
         End
      End
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8805
      Width           =   2175
   End
   Begin VB.CommandButton CReset 
      BackColor       =   &H00C0C0C0&
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
      Left            =   9720
      TabIndex        =   4
      Top             =   8805
      Width           =   2175
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   4575
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
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   375
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
      Left            =   5520
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前阵营(&O)"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstFactions 
      Height          =   8415
      Left            =   120
      TabIndex        =   6
      Top             =   525
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   14843
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmFactions.frx":08CA
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
      Left            =   5280
      MouseIcon       =   "frmFactions.frx":4F37
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9060
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
      Left            =   4440
      MouseIcon       =   "frmFactions.frx":5241
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9060
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "阵营数:"
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
      Left            =   135
      TabIndex        =   10
      Top             =   9060
      Width           =   690
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
      Top             =   195
      Width           =   495
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
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   2
      Left            =   2760
      MouseIcon       =   "frmFactions.frx":554B
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   9060
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
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   3600
      MouseIcon       =   "frmFactions.frx":5855
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   9060
      Width           =   705
   End
End
Attribute VB_Name = "frmFactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CustomActive As Boolean

Private Sub CApply_Click()
Dim q As Boolean

If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
    CustomActive = False
    If UCase(Factions(CurrentFactionID).strID) <> UCase(CurrentFaction.strID) Then             '外引
        q = ChangeStrID(Factions(CurrentFactionID).strID, CurrentFaction.strID)
        
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentFaction.strID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    Factions(CurrentFactionID) = CurrentFaction

    LstFactions.ListItems(CurrentFactionID + 1).SubItems(1) = Factions(CurrentFactionID).csvName
    LstFactions.ListItems(CurrentFactionID + 1).SubItems(2) = Factions(CurrentFactionID).strID
    
    CurrentFaction = Factions(CurrentFactionID)
    InitRelationShipsListView2
    LoadFactionInfo
    
    CustomActive = True
End If
End Sub

Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0    '删除
           If Factions(CurrentFactionID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), Factions(CurrentFactionID).strID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
              If CurrentFactionID < N_Faction - 1 Then
                DelIndex Factions(CurrentFactionID).strID
                For i = CurrentFactionID To N_Faction - 2 Step 1
                    ChangeID Factions(i + 1).strID, Factions(i + 1).ID - 1
                    j = Factions(i).ID
                    Factions(i) = Factions(i + 1)
                    Factions(i).ID = j
                    LstFactions.ListItems(i + 1).SubItems(1) = LstFactions.ListItems(i + 2).SubItems(1)
                    LstFactions.ListItems(i + 1).SubItems(2) = LstFactions.ListItems(i + 2).SubItems(2)
                Next i
                
                LstFactions.ListItems.Remove N_Faction
                RedimFactions N_Faction - 1
                
              Else
                DelIndex Factions(CurrentFactionID).strID
                LstFactions.ListItems.Remove N_Faction
                RedimFactions N_Faction - 1
                CurrentFactionID = N_Faction - 1
                
              End If
               InitRelationShipsListView2
               LstFactions_ItemClick LstFactions.ListItems(CurrentFactionID + 1)
               LstFactions.ListItems(CurrentFactionID + 1).Selected = True
               LstFactions.ListItems(CurrentFactionID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), Factions(CurrentFactionID).strID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
       Case 1   '创建
         If MsgBox(ActiveString(PublicMsgs(5), Factions(CurrentFactionID).strID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_Faction, Factions(CurrentFactionID).strID & "_New") Then
           RedimFactions N_Faction + 1
           
           Factions(N_Faction - 1) = Factions(CurrentFactionID)
           With Factions(N_Faction - 1)
                 .ID = N_Faction - 1
                 .strID = .strID & "_New"
                 .strName = .strName & "_New"
                 .csvName = .csvName & "_New"
                 .Edit = True
           End With
           
           'Redim RelationShips
             For i = 0 To N_Faction - 1
                ReDim Preserve Factions(i).RelationShip(N_Faction - 1)
                Factions(i).RelationShip(N_Faction - 1).ID = N_Faction - 1
                Factions(i).RelationShip(N_Faction - 1).strID = Factions(N_Faction - 1).strID
             Next i
  
           Set oItem = LstFactions.ListItems.Add(, "Factions_" & Factions(N_Faction - 1).ID, Factions(N_Faction - 1).ID)
      
                 With oItem
                    .SubItems(1) = Factions(N_Faction - 1).csvName
                    .SubItems(2) = Factions(N_Faction - 1).strID
                 End With
           InitRelationShipsListView2
           LstFactions_ItemClick LstFactions.ListItems(N_Faction)
           LstFactions.ListItems(N_Faction).Selected = True
           LstFactions.ListItems(N_Faction).EnsureVisible
           
           Else
           
           MsgBox ActiveString(PublicMsgs(90), Factions(CurrentFactionID).strID & "_New"), vbCritical, PublicMsgs(19)
           End If
           
         End If
         
      Case 2       '上移
         If CurrentFactionID > 0 Then
           If Factions(CurrentFactionID - 1).Edit And Factions(CurrentFactionID).Edit Then
           
                SwapID Factions(CurrentFactionID - 1).strID, Factions(CurrentFactionID).strID
                SwapFactions CurrentFactionID - 1, CurrentFactionID
                SwapListItem LstFactions.ListItems(CurrentFactionID), LstFactions.ListItems(CurrentFactionID + 1), 2, True
                CurrentItmID = CurrentItmID - 1
               InitRelationShipsListView2
               LstFactions_ItemClick LstFactions.ListItems(CurrentFactionID)
               LstFactions.ListItems(CurrentFactionID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), Factions(CurrentFactionID - 1).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
         End If
      Case 3       '下移
        If CurrentFactionID + 1 <= N_Faction - 1 Then
           If Factions(CurrentFactionID).Edit And Factions(CurrentFactionID + 1).Edit Then
           
                SwapID Factions(CurrentFactionID).strID, Factions(CurrentFactionID + 1).strID
                SwapFactions CurrentFactionID, CurrentFactionID + 1
                SwapListItem LstFactions.ListItems(CurrentFactionID + 1), LstFactions.ListItems(CurrentFactionID + 2), 2, True
                CurrentItmID = CurrentItmID + 1
                InitRelationShipsListView2
                LstFactions_ItemClick LstFactions.ListItems(CurrentFactionID + 2)
                LstFactions.ListItems(CurrentFactionID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), Factions(CurrentFactionID).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
           
        End If
End Select
End Sub


Private Sub chkFlags_Click(Index As Integer)
Dim I64(1) As Integer64b
If CustomActive Then

    I64(0) = StrToI64(CurrentFaction.Flags)
    I64(1) = HexStrToI64(chkFlags(Index).Tag)
    
If chkFlags(Index).Value = 1 Then
    I64(0) = AddFlags64b(I64(0), I64(1))
Else
    I64(0) = DeleteFlags64b(I64(0), I64(1))
End If

CurrentFaction.Flags = I64toStrNZ(I64(0))
End If
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstFactions, LstFactions.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstFactions_ItemClick(LstFactions.SelectedItem)
End If
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

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbExclamation, PublicMsgs(0)) = vbYes Then
LstFactions_ItemClick LstFactions.ListItems(CurrentFactionID + 1)
LstFactions.ListItems(CurrentFactionID + 1).Selected = True
LstFactions.ListItems(CurrentFactionID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
CustomActive = False

InitRelationShipsListView
InitFactionsListView
InitRelationShipsListView2
LoadFactionsList
InitFlagsList

CurrentFactionID = 0
CurrentFaction = Factions(CurrentFactionID)
LoadFactionInfo

TranslateForm Me
InitCMDs

Label2.Caption = Label2.Caption & N_Faction
Label1(3).Caption = Replace(Label1(3).Caption, "(NOW):", "(" & UCase$(MnBInfo.Language) & "):", , , vbTextCompare)

If LCase$(MnBInfo.Language) = "en" Then
    txtPropBag(2).Enabled = False
End If

CustomActive = True
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

Private Sub InitFactionsListView()
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

      .ColumnHeaders.Add , , PublicMsgs(13), .Width / n / 3.6
      .ColumnHeaders.Add , , PublicEditors(5) & PublicMsgs(14), .Width / n * 1.5
      .ColumnHeaders.Add , , PublicEditors(5) & "ID", .Width / n * 1.5
End With

End Sub

Private Sub InitRelationShipsListView2()
Dim i As Long, oItem As ListItem

LstRelationShips.ListItems.Clear
For i = 0 To N_Faction - 1

      Set oItem = LstRelationShips.ListItems.Add(, , Factions(i).strID)
      
      With oItem
         .SubItems(1) = Factions(i).csvName
      End With

Next i

LstRelationShips.ListItems(1).Selected = True
LstRelationShips_ItemClick LstRelationShips.ListItems(1)
End Sub

Private Sub InitRelationShipsListView()
Dim n As Integer
n = 3
LstRelationShips.View = lvwReport
LstRelationShips.Sorted = False
LstRelationShips.ListItems.Clear
LstRelationShips.ColumnHeaders.Clear
LstRelationShips.SortOrder = lvwAscending
LstRelationShips.FullRowSelect = True
LstRelationShips.AllowColumnReorder = False
LstRelationShips.LabelEdit = lvwManual
LstRelationShips.Checkboxes = False
LstRelationShips.GridLines = True
LstRelationShips.MultiSelect = False
LstRelationShips.HideSelection = False

LstRelationShips.ColumnHeaders.Add , , PublicEditors(5) & "ID", LstRelationShips.Width / n
LstRelationShips.ColumnHeaders.Add , , PublicEditors(5) & PublicMsgs(14), LstRelationShips.Width / n
LstRelationShips.ColumnHeaders.Add , , PublicMsgs(6), LstRelationShips.Width / n


End Sub

'*************************************************************************
'**函 数 名：LoadFactionsList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-02 15:02:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFactionsList()
Dim n As Long, oItem As ListItem

LstFactions.ListItems.Clear
For n = 0 To N_Faction - 1

      Set oItem = LstFactions.ListItems.Add(, "fac_" & Factions(n).ID, Factions(n).ID)
      
      With oItem
         .SubItems(1) = Factions(n).csvName
         .SubItems(2) = Factions(n).strID
      End With

Next n

End Sub

'*************************************************************************
'**函 数 名：LoadRelationShipsList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-02 15:02:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadRelationShipsList()
Dim n As Long
CustomActive = False

StructureFactionRelationShips -1
For n = 1 To N_Faction
      With LstRelationShips.ListItems(n)
         .SubItems(2) = CurrentFaction.RelationShip(n - 1).Value
      End With
Next n

LstRelationShips.ListItems(1).Selected = True
txtRelationShip.Text = LstRelationShips.ListItems(1).SubItems(2)

CustomActive = True
End Sub

'*************************************************************************
'**函 数 名：LoadFactionInfo
'**输    入：-
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-02 14:34:47
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFactionInfo()
Dim i As Integer, oItem As ListItem, strTem As String, lngTem As Long
CustomActive = False

With CurrentFaction
       txtPropBag(0).Text = .strID
       txtPropBag(1).Text = .strName
       txtPropBag(2).Text = .csvName
       txtPropBag(3).Text = MnBtoRGBColor(.lColor)
         PicColor.BackColor = "&H" & txtPropBag(3).Text
         
       LoadRelationShipsList
       LoadFlagsList
       
End With

CustomActive = True
End Sub

'*************************************************************************
'**函 数 名：InitFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-03 22:22:12
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitFlagsList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("永不显示阵营标签")
TemArray = Array(ff_always_hide_label)
         For i = 0 To UBound(strTemArray)
             chkFlags(i).Caption = strTemArray(i)
             chkFlags(i).Tag = TemArray(i)
         Next i

End Sub



Private Sub LstFactions_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentFactionID = Val(Item.Text)

CurrentFaction = Factions(CurrentFactionID)
LoadFactionInfo
End Sub



Private Sub LstRelationShips_ItemClick(ByVal Item As MSComctlLib.ListItem)
CustomActive = False
txtRelationShip.Text = Item.SubItems(2)
CustomActive = True
End Sub

Private Sub PicColor_Click()
On Error GoTo CancelLine
If CustomActive Then

'frmColor.Show
'frmColor.InitColor = PicColor.BackColor
'Me.Enabled = False
CD.Color = PicColor.BackColor
CD.Flags = cdlCCRGBInit

CD.ShowColor
PicColor.BackColor = CD.Color
txtPropBag(3).Text = Hex(CD.Color)
CurrentFaction.lColor = RGBtoMnBColor(txtPropBag(3).Text)
Exit Sub

End If
CancelLine:
End Sub

Private Sub txtPropBag_LostFocus(Index As Integer)
If CustomActive Then
With CurrentFaction
If Index < 2 Then
   txtPropBag(Index).Text = Replace(txtPropBag(Index).Text, " ", "_")
End If

Select Case Index
      Case 0
          'Check Value
            If Left$(txtPropBag(Index).Text, 4) <> "fac_" Then
               txtPropBag(Index).Text = "fac_" & txtPropBag(Index).Text
            End If

         .strID = txtPropBag(Index).Text
      Case 1
         .strName = txtPropBag(Index).Text
           If LCase$(MnBInfo.Language) = "en" Then
             .csvName = .strName
           End If
      Case 2
          If LCase$(MnBInfo.Language) <> "en" Then
             .csvName = txtPropBag(Index).Text
          End If
      Case 3
         .lColor = RGBtoMnBColor(txtPropBag(3).Text)
         PicColor.BackColor = "&H" & txtPropBag(3).Text
End Select

End With
End If
End Sub



Private Sub txtRating_LostFocus()
Dim a As Integer64b, TemFlags As Long, tI(3) As Integer64b
If CustomActive Then

With CurrentFaction
    tI(1) = StrToI64(CStr(.Flags))     '原阵营Flags
    tI(2) = StrToI64(CStr(ff_max_rating_mask))   '所有阵营类型Flags
    
    tI(3) = StrToI64(txtRating.Text)
      LeftMv64bEx tI(3), ff_max_rating_bits        '要添加的阵营类型Flags
    
    tI(0) = DeleteFlagsI64(tI(1), tI(2))         '清空所有阵营类型Flags
    Frame1(2).Caption = DetoBinString(Val(I64toStrNZ(tI(0))))
    tI(0) = AddFlagsI64(tI(0), tI(3))             '添加阵营类型Flags
    .Flags = Val(I64toStrNZ(tI(0)))
    
End With

End If
End Sub

Private Sub txtRelationShip_Change()
If CustomActive Then
LstRelationShips.SelectedItem.SubItems(2) = txtRelationShip.Text
CurrentFaction.RelationShip(LstRelationShips.SelectedItem.Index - 1).Value = Val(txtRelationShip.Text)
End If
End Sub

'*************************************************************************
'**函 数 名：LoadFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-04 07:45:31
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFlagsList()
Dim i As Integer, StrBin As String, k As Long, tI As Integer64b

With CurrentFaction

StrBin = DetoBinString_15(.Flags)
For i = 0 To chkFlags.UBound
       If And_15(.Flags, Val(chkFlags(i).Tag)) Then
          chkFlags(i).Value = 1
       Else
          chkFlags(i).Value = 0
       End If
Next i

tI = StrToI64(CStr(.Flags))
RightMv64bEx tI, ff_max_rating_bits
txtRating.Text = I64toStrNZ(tI)
End With
End Sub

Public Sub ReLoadInfo()
LoadFactionInfo
End Sub
