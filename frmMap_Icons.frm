VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMap_Icons 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "大地图图标编辑器"
   ClientHeight    =   9375
   ClientLeft      =   1755
   ClientTop       =   180
   ClientWidth     =   14355
   Icon            =   "frmMap_Icons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_7"
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
      Caption         =   "输出当前图标(&O)"
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
   Begin MSComctlLib.ListView LstMapIcons 
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
      Picture         =   "frmMap_Icons.frx":08FF
   End
   Begin VB.Frame FraProps 
      Caption         =   "图标属性"
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
         Caption         =   "模型"
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
         Height          =   2535
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   7455
         Begin VB.Frame Frame1 
            Caption         =   "偏移"
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
            Height          =   855
            Index           =   3
            Left            =   600
            TabIndex        =   24
            Top             =   1440
            Width           =   5535
            Begin VB.TextBox txtPropBag 
               Height          =   375
               Index           =   4
               Left            =   2400
               TabIndex        =   26
               Top             =   280
               Width           =   1095
            End
            Begin VB.TextBox txtPropBag 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   840
               TabIndex        =   25
               Top             =   280
               Width           =   1095
            End
            Begin VB.TextBox txtPropBag 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2052
                  SubFormatType   =   1
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   3960
               TabIndex        =   27
               Top             =   280
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Z:"
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
               Left            =   3720
               TabIndex        =   33
               Top             =   360
               Width           =   210
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Y:"
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
               Index           =   4
               Left            =   2160
               TabIndex        =   32
               Top             =   360
               Width           =   210
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "X:"
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
               Index           =   3
               Left            =   525
               TabIndex        =   31
               Top             =   360
               Width           =   210
            End
         End
         Begin VB.CheckBox chkOffsets 
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
            Left            =   360
            TabIndex        =   30
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox txtPropBag 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   22
            Top             =   840
            Width           =   5655
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   20
            Top             =   360
            Width           =   5655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "比例:"
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
            Index           =   2
            Left            =   795
            TabIndex        =   23
            Top             =   915
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "网格名:"
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
            TabIndex        =   21
            Top             =   435
            Width           =   690
         End
      End
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
         Height          =   1695
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   4920
         Width           =   7455
         Begin VB.CheckBox chkCustom 
            Caption         =   "设置为自定义旗帜的大地图图标"
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
            Left            =   480
            TabIndex        =   34
            Top             =   960
            Width           =   6615
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
            Left            =   480
            TabIndex        =   18
            Top             =   480
            Width           =   6615
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
         Height          =   1815
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   7455
         Begin VB.ComboBox cbSounds 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1080
            Width           =   5655
         End
         Begin VB.TextBox txtPropBag 
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   15
            Top             =   480
            Width           =   5655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "声音:"
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
            Index           =   6
            Left            =   795
            TabIndex        =   28
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "图标名:"
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
            Left            =   600
            TabIndex        =   16
            Top             =   600
            Width           =   690
         End
      End
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
      MouseIcon       =   "frmMap_Icons.frx":4F6C
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
      MouseIcon       =   "frmMap_Icons.frx":5276
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9060
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图标数:"
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
      MouseIcon       =   "frmMap_Icons.frx":5580
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
      MouseIcon       =   "frmMap_Icons.frx":588A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   9060
      Width           =   705
   End
End
Attribute VB_Name = "frmMap_Icons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CustomActive As Boolean
Dim TemTriggersCount As Long, TemTriggers As Type_Trigger

Private Sub CApply_Click()
Dim q As Boolean

If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then

    If UCase(MapIcons(CurrentMapIconID).strID) <> UCase(CurrentMapIcon.strID) Then             '外引
        q = ChangeStrID(MapIcons(CurrentMapIconID).strID, CurrentMapIcon.strID)
        
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentMapIcon.strID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    MapIcons(CurrentMapIconID) = CurrentMapIcon
    LstMapIcons.ListItems(CurrentMapIconID + 1).SubItems(1) = MapIcons(CurrentMapIconID).strID
    
    CurrentMapIcon = MapIcons(CurrentMapIconID)
    LoadMapIconInfo
End If
End Sub

Private Sub cbSounds_Click()
If CustomActive Then
    CurrentMapIcon.Sound = cbSounds.ListIndex
    CurrentMapIcon.Sound_sndName = Sounds(CurrentMapIcon.Sound).sndName
End If
End Sub

Private Sub cbSounds_Scroll()
Call cbSounds_Click
End Sub

Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If MapIcons(CurrentMapIconID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), MapIcons(CurrentMapIconID).strID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
              DelIndex MapIcons(CurrentMapIconID).strID
              
              If CurrentMapIconID < N_MapIcon - 1 Then
                For i = CurrentMapIconID To N_MapIcon - 2 Step 1
                    ChangeID MapIcons(i + 1).strID, MapIcons(i + 1).ID - 1
                    j = MapIcons(i).ID
                    MapIcons(i) = MapIcons(i + 1)
                    MapIcons(i).ID = j
                    LstMapIcons.ListItems(i + 1).SubItems(1) = LstMapIcons.ListItems(i + 2).SubItems(1)
                Next i
                
                ReDim Preserve MapIcons(N_MapIcon - 2)
                LstMapIcons.ListItems.Remove N_MapIcon
                N_MapIcon = N_MapIcon - 1
                
              Else
                ReDim Preserve MapIcons(N_MapIcon - 2)
                LstMapIcons.ListItems.Remove N_MapIcon
                
                N_MapIcon = N_MapIcon - 1
                CurrentMapIconID = N_MapIcon - 1
                
              End If
               
               LstMapIcons_ItemClick LstMapIcons.ListItems(CurrentMapIconID + 1)
               LstMapIcons.ListItems(CurrentMapIconID + 1).Selected = True
               LstMapIcons.ListItems(CurrentMapIconID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), MapIcons(CurrentMapIconID).strID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
       Case 1
       
       If N_MapIcon < Val("&H" & pf_icon_mask) + 1 Then
         If MsgBox(ActiveString(PublicMsgs(5), MapIcons(CurrentMapIconID).strID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
           
           If AddIndex(N_MapIcon, MapIcons(CurrentMapIconID).strID & "_New") Then
           ReDim Preserve MapIcons(N_MapIcon)
           N_MapIcon = N_MapIcon + 1
           MapIcons(N_MapIcon - 1) = MapIcons(CurrentMapIconID)
           With MapIcons(N_MapIcon - 1)
                 .ID = N_MapIcon - 1
                 .strID = .strID & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstMapIcons.ListItems.Add(, "mis_" & MapIcons(N_MapIcon - 1).ID, MapIcons(N_MapIcon - 1).ID)
      
                 With oItem
                    .SubItems(1) = MapIcons(N_MapIcon - 1).strID
                 End With
           LstMapIcons_ItemClick LstMapIcons.ListItems(N_MapIcon)
           LstMapIcons.ListItems(N_MapIcon).Selected = True
           LstMapIcons.ListItems(N_MapIcon).EnsureVisible
           
           Else
              MsgBox ActiveString(PublicMsgs(90), MapIcons(CurrentMapIconID).strID & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
      Else
         MsgBox ActiveString(PublicMsgs(18), PublicEditors(GetEditorIndex(Me.Tag)), Val("&H" & pf_icon_mask) + 1), vbCritical, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))
      End If
      Case 2
         If CurrentMapIconID > 0 Then
           If MapIcons(CurrentMapIconID - 1).Edit And MapIcons(CurrentMapIconID).Edit Then
           
                SwapID MapIcons(CurrentMapIconID - 1).strID, MapIcons(CurrentMapIconID).strID
                SwapMapIcons CurrentMapIconID - 1, CurrentMapIconID
                SwapListItem LstMapIcons.ListItems(CurrentMapIconID), LstMapIcons.ListItems(CurrentMapIconID + 1), 1, True
                
               LstMapIcons_ItemClick LstMapIcons.ListItems(CurrentMapIconID)
               LstMapIcons.ListItems(CurrentMapIconID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), MapIcons(CurrentMapIconID - 1).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
         End If
      Case 3
        If CurrentMapIconID + 1 <= N_MapIcon - 1 Then
           If MapIcons(CurrentMapIconID).Edit And MapIcons(CurrentMapIconID + 1).Edit Then
           
                SwapID MapIcons(CurrentMapIconID).strID, MapIcons(CurrentMapIconID + 1).strID
                SwapMapIcons CurrentMapIconID, CurrentMapIconID + 1
                SwapListItem LstMapIcons.ListItems(CurrentMapIconID + 1), LstMapIcons.ListItems(CurrentMapIconID + 2), 1, True
                
                LstMapIcons_ItemClick LstMapIcons.ListItems(CurrentMapIconID + 2)
                LstMapIcons.ListItems(CurrentMapIconID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), MapIcons(CurrentMapIconID).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))

           End If
           
        End If
End Select
End Sub



'*************************************************************************
'**函 数 名：InitTempTrggier
'**输    入：-
'**输    出：-
'**功能描述：装载自定义旗帜的大地图图标的触发器
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 10:48:03
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitTempTrggier()

TemTriggersCount = 1

With TemTriggers
     .tiOn = ti_on_init_map_icon
     .ActNum = 6
     ReDim .tiAct(1 To .ActNum)
     
     .tiAct(1).Op = store_trigger_param_1
     .tiAct(1).ParaNum = 1
      ReDim .tiAct(1).Para(1 To .tiAct(1).ParaNum)
     .tiAct(1).Para(1).Value = "1224979098644774912"
     
     .tiAct(2).Op = party_get_slot
     .tiAct(2).ParaNum = 3
      ReDim .tiAct(2).Para(1 To .tiAct(2).ParaNum)
     .tiAct(2).Para(1).Value = "1224979098644774913"
     .tiAct(2).Para(2).Value = "1224979098644774912"
     .tiAct(2).Para(3).Value = "7"
     
     .tiAct(3).Op = try_begin
     .tiAct(3).ParaNum = 0

     .tiAct(4).Op = ge
     .tiAct(4).ParaNum = 2
      ReDim .tiAct(4).Para(1 To .tiAct(4).ParaNum)
     .tiAct(4).Para(1).Value = "1224979098644774913"
     .tiAct(4).Para(2).Value = "0"
     
     .tiAct(5).Op = cur_map_icon_set_tableau_material
     .tiAct(5).ParaNum = 2
      ReDim .tiAct(5).Para(1 To .tiAct(5).ParaNum)
     .tiAct(5).Para(1).Value = "1729382256910270509"
     .tiAct(5).Para(2).Value = "1224979098644774913"
     
     .tiAct(6).Op = try_end
     .tiAct(6).ParaNum = 0
     
End With
End Sub

Private Sub chkCustom_Click()
If CustomActive Then
    With CurrentMapIcon
      If chkCustom.Value = 1 Then
         .TriggerCount = TemTriggersCount
         ReDim .Triggers(1 To .TriggerCount)
         
         .Triggers(1) = TemTriggers
      Else
         .TriggerCount = 0
      End If
    End With
End If
End Sub

Private Sub chkFlags_Click(Index As Integer)
If CustomActive Then
   With CurrentMapIcon
        .Flags = chkFlags(Index).Value
   End With
End If
End Sub

Private Sub chkOffsets_Click()
Dim i As Integer
If CustomActive Then
  With CurrentMapIcon
   If chkOffsets.Value = 1 Then
     Frame1(3).Enabled = True
     
     For i = 0 To 2
         .Offset(i) = Format(.Offset(i), "0.000000")
         txtPropBag(i + 3).Text = .Offset(i)
     Next i
   Else
     Frame1(3).Enabled = False
     
     For i = 0 To 2
         .Offset(i) = "0"
         txtPropBag(i + 3).Text = .Offset(i)
     Next i
   End If
  End With
End If
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstMapIcons, LstMapIcons.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstMapIcons_ItemClick(LstMapIcons.SelectedItem)
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
    Set oItem = FindItem(oLV, Start, "0|1", QueryString, True, vbTextCompare, bReverse)
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
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
LstMapIcons_ItemClick LstMapIcons.ListItems(CurrentMapIconID + 1)
LstMapIcons.ListItems(CurrentMapIconID + 1).Selected = True
LstMapIcons.ListItems(CurrentMapIconID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
CustomActive = False



InitSoundsCombo
InitMapIconsListView
LoadMapIconsList
InitFlagsList
InitTempTrggier

CurrentMapIconID = 0
CurrentMapIcon = MapIcons(CurrentMapIconID)
LoadMapIconInfo

TranslateForm Me
InitCMDs

Label2.Caption = Label2.Caption & N_MapIcon

CustomActive = True
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

Private Sub InitMapIconsListView()
Dim n As Integer
n = 2
LstMapIcons.View = lvwReport
LstMapIcons.Sorted = False
LstMapIcons.ListItems.Clear
LstMapIcons.ColumnHeaders.Clear
LstMapIcons.SortOrder = lvwAscending
LstMapIcons.FullRowSelect = True
LstMapIcons.AllowColumnReorder = False
LstMapIcons.LabelEdit = lvwManual
LstMapIcons.Checkboxes = False
LstMapIcons.GridLines = True
LstMapIcons.MultiSelect = False
LstMapIcons.HideSelection = False

LstMapIcons.ColumnHeaders.Add , , PublicMsgs(13), LstMapIcons.Width / n / 3.6
LstMapIcons.ColumnHeaders.Add , , PublicEditors_Simplified(7) & PublicMsgs(14), LstMapIcons.Width / n * 1.5

End Sub


'*************************************************************************
'**函 数 名：LoadMapIconsList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 00:50:24
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadMapIconsList()
Dim n As Long, oItem As ListItem

LstMapIcons.ListItems.Clear
For n = 0 To N_MapIcon - 1

      Set oItem = LstMapIcons.ListItems.Add(, "mis_" & MapIcons(n).ID, MapIcons(n).ID)
      
      With oItem
         .SubItems(1) = MapIcons(n).strID
      End With

Next n

End Sub

'*************************************************************************
'**函 数 名：LoadMapIconInfo
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
Private Sub LoadMapIconInfo()
Dim i As Integer, oItem As ListItem, strTem As String, lngTem As Long, NeedOffset As Boolean
CustomActive = False

With CurrentMapIcon
       txtPropBag(0).Text = .strID
       txtPropBag(1).Text = .MeshName
       txtPropBag(2).Text = .mScale
       
   NeedOffset = CheckUseOffset(.Offset(i))
   If NeedOffset Then
      chkOffsets.Value = 1
      Frame1(3).Enabled = True
   Else
      chkOffsets.Value = 0
      Frame1(3).Enabled = False
   End If
     For i = 0 To 2
       If NeedOffset Then
         txtPropBag(i + 3).Text = Format(.Offset(i), "0.000000")
       Else
         txtPropBag(i + 3).Text = "0"
       End If
     Next i
     
       LoadFlagsList
       
       '.Sound = IIf(CheckExist(EditInfo_SoundsCount, .Sound), .Sound, 0)
       .Sound = GetID(.Sound_sndName, , Sounds(0).sndName)
       cbSounds.ListIndex = .Sound
       
     If .TriggerCount > 0 Then
         chkCustom.Value = 1
     Else
         chkCustom.Value = 0
     End If
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
'**日    期：2010-12-11 00:48:02
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitFlagsList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("无阴影")
TemArray = Array(mcn_no_shadow)
         For i = 0 To UBound(strTemArray)
             chkFlags(i).Caption = strTemArray(i)
             chkFlags(i).Tag = TemArray(i)
         Next i

End Sub



Private Sub LstMapIcons_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentMapIconID = Val(Item.Text)

CurrentMapIcon = MapIcons(CurrentMapIconID)
LoadMapIconInfo
End Sub


Private Sub txtPropBag_LostFocus(Index As Integer)
If CustomActive Then
With CurrentMapIcon

If Index < 2 Then
   txtPropBag(Index).Text = Replace(txtPropBag(Index).Text, " ", "_")
End If

Select Case Index
      Case 0
         .strID = txtPropBag(Index).Text
      Case 1
         .MeshName = txtPropBag(Index).Text
      Case 2
         .mScale = Format(txtPropBag(Index).Text, "0.000000")
      Case 3
         If chkOffsets.Value = 1 Then
            .Offset(Index - 3) = Format(txtPropBag(Index).Text, "0.000000")
         Else
            .Offset(Index - 3) = "0"
         End If
      Case 4
         If chkOffsets.Value = 1 Then
            .Offset(Index - 3) = Format(txtPropBag(Index).Text, "0.000000")
         Else
            .Offset(Index - 3) = "0"
         End If
      Case 5
         If chkOffsets.Value = 1 Then
            .Offset(Index - 3) = Format(txtPropBag(Index).Text, "0.000000")
         Else
            .Offset(Index - 3) = "0"
         End If
End Select

End With
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
Dim i As Integer

With CurrentMapIcon
    chkFlags(0).Value = .Flags
    
End With

End Sub

'*************************************************************************
'**函 数 名：InitSoundsCombo
'**输    入：
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 09:09:56
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitSoundsCombo()
Dim j As Long

   cbSounds.Clear
   
   For j = 0 To N_Sound - 1
         cbSounds.AddItem "(" & j & ")" & Sounds(j).sndName
   Next j
   
End Sub

'*************************************************************************
'**函 数 名：CheckUseOffset
'**输    入：(String)Offset
'**输    出：(Boolean)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 09:29:08
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Function CheckUseOffset(ByVal Offset As String) As Boolean
Dim n As Long

n = InStr(1, Offset, ".")

CheckUseOffset = n > 0
End Function

Public Sub ReLoadInfo()
LoadMapIconInfo
End Sub
