VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMesh 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "网格编辑器"
   ClientHeight    =   9375
   ClientLeft      =   3105
   ClientTop       =   960
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14355
   ShowInTaskbar   =   0   'False
   Tag             =   "edit_12"
   Begin MSComctlLib.ImageList IL1 
      Left            =   960
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesh.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton COutputLine 
      BackColor       =   &H0080FF80&
      Caption         =   "输出当前网格模型(&O)"
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
      TabIndex        =   14
      Top             =   8760
      Width           =   2415
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
      TabIndex        =   7
      Top             =   8760
      Width           =   2175
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
      Left            =   9720
      TabIndex        =   6
      Top             =   8760
      Width           =   2175
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtQuery 
      Height          =   330
      Left            =   720
      TabIndex        =   3
      Top             =   100
      Width           =   4455
   End
   Begin MSComctlLib.ListView LstMesh 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   15055
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
      Picture         =   "frmMesh.frx":039A
   End
   Begin VB.Frame FraProps 
      Caption         =   "FraProps0"
      Height          =   7935
      Index           =   0
      Left            =   6240
      TabIndex        =   9
      Top             =   480
      Width           =   7815
      Begin VB.Frame FraScale 
         Caption         =   "尺寸"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   1815
         Left            =   5280
         TabIndex        =   36
         Top             =   1800
         Width           =   2055
         Begin VB.TextBox txtScale 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   2
            Left            =   600
            TabIndex        =   39
            Text            =   "Z"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtScale 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   1
            Left            =   600
            TabIndex        =   38
            Text            =   "Y"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtScale 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   0
            Left            =   600
            TabIndex        =   37
            Text            =   "X"
            Top             =   435
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Z:"
            Height          =   180
            Left            =   360
            TabIndex        =   42
            Top             =   1245
            Width           =   180
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Y:"
            Height          =   180
            Left            =   360
            TabIndex        =   41
            Top             =   885
            Width           =   180
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "X:"
            Height          =   180
            Left            =   360
            TabIndex        =   40
            Top             =   480
            Width           =   180
         End
      End
      Begin VB.Frame FraRotation 
         Caption         =   "旋转"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   1815
         Left            =   2810
         TabIndex        =   29
         Top             =   1800
         Width           =   2055
         Begin VB.TextBox txtRotation 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   0
            Left            =   600
            TabIndex        =   32
            Text            =   "X"
            Top             =   435
            Width           =   1095
         End
         Begin VB.TextBox txtRotation 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   1
            Left            =   600
            TabIndex        =   31
            Text            =   "Y"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtRotation 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   2
            Left            =   600
            TabIndex        =   30
            Text            =   "Z"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "X轴:"
            Height          =   180
            Left            =   230
            TabIndex        =   35
            Top             =   480
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Y轴:"
            Height          =   180
            Left            =   230
            TabIndex        =   34
            Top             =   885
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Z轴:"
            Height          =   180
            Left            =   230
            TabIndex        =   33
            Top             =   1245
            Width           =   360
         End
      End
      Begin VB.Frame FraTranslation 
         Caption         =   "移动"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   1815
         Left            =   360
         TabIndex        =   22
         Top             =   1800
         Width           =   2055
         Begin VB.TextBox txtTranslation 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   2
            Left            =   600
            TabIndex        =   28
            Text            =   "Z"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtTranslation 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   1
            Left            =   600
            TabIndex        =   27
            Text            =   "Y"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtTranslation 
            Alignment       =   1  'Right Justify
            Height          =   270
            Index           =   0
            Left            =   600
            TabIndex        =   26
            Text            =   "X"
            Top             =   435
            Width           =   1095
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Z轴:"
            Height          =   180
            Left            =   230
            TabIndex        =   25
            Top             =   1245
            Width           =   360
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Y轴:"
            Height          =   180
            Left            =   230
            TabIndex        =   24
            Top             =   885
            Width           =   360
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "X轴:"
            Height          =   180
            Left            =   230
            TabIndex        =   23
            Top             =   480
            Width           =   360
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "基础信息"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   1335
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   6975
         Begin VB.TextBox TxtResourceName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            TabIndex        =   20
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox TxtMeshName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            TabIndex        =   18
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "网格资源名:"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   21
            Top             =   885
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "网格模型名:"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   19
            Top             =   405
            Width           =   1245
         End
      End
      Begin VB.Frame FraFlags 
         Caption         =   "特性"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   2775
         Left            =   360
         TabIndex        =   15
         Top             =   3840
         Width           =   3495
         Begin VB.ListBox LstFlags 
            Height          =   2160
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   360
            Width           =   3015
         End
      End
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   8415
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   14843
      MultiRow        =   -1  'True
      ImageList       =   "IL1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "基本信息(&I)"
            Key             =   "Info"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
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
      Left            =   3480
      MouseIcon       =   "frmMesh.frx":4A07
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   9120
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
      Left            =   2640
      MouseIcon       =   "frmMesh.frx":4D11
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9120
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
      Left            =   5160
      MouseIcon       =   "frmMesh.frx":501B
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   9120
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
      ForeColor       =   &H00008000&
      Height          =   180
      Index           =   1
      Left            =   4320
      MouseIcon       =   "frmMesh.frx":5325
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   9120
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "网格模型数:"
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
      Top             =   9120
      Width           =   1080
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
      Left            =   200
      TabIndex        =   2
      Top             =   165
      Width           =   495
   End
End
Attribute VB_Name = "frmMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CustomActive As Boolean

Private Sub InitMeshListView()
Dim n As Integer
n = 2
LstMesh.View = lvwReport
LstMesh.Sorted = False
LstMesh.ListItems.Clear
LstMesh.ColumnHeaders.Clear
LstMesh.SortOrder = lvwAscending
LstMesh.FullRowSelect = True
LstMesh.AllowColumnReorder = False
LstMesh.LabelEdit = lvwManual
LstMesh.Checkboxes = False
LstMesh.GridLines = True
LstMesh.MultiSelect = False
LstMesh.HideSelection = False

LstMesh.ColumnHeaders.Add , , PublicMsgs(13), LstMesh.Width / n / 4
LstMesh.ColumnHeaders.Add , , PublicEditors(12) & PublicMsgs(14), LstMesh.Width / n * 1.5

End Sub

'*************************************************************************
'**函 数 名：LoadMeshList
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2010-11-26 23:16:41
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadMeshList()
Dim n As Long, oItem As ListItem, tI As Integer64b, H As Integer

LstMesh.ListItems.Clear  '清空列表

For n = 0 To UBound(Mesh)

  H = n - 1      '上一网格模型
  If H < 0 Then H = 0

      Set oItem = LstMesh.ListItems.Add(, "Mesh_" & CStr(n), n)
      
      With oItem
       .SubItems(1) = Mesh(n).strID
      End With

Next n

End Sub

Private Sub CApply_Click()
Dim q As Boolean
If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
    
    If UCase(Mesh(CurrentMeshID).strID) <> UCase(CurrentMesh.strID) Then               '外引
        q = ChangeStrID(Mesh(CurrentMeshID).strID, CurrentMesh.strID)
        If Not q Then
           MsgBox ActiveString(PublicMsgs(91), CurrentMesh.strID), vbCritical, PublicMsgs(19)
           Exit Sub
        End If
    End If
    
    Mesh(CurrentMeshID) = CurrentMesh
    
    LstMesh.ListItems(CurrentMeshID + 1).SubItems(1) = Mesh(CurrentMeshID).strID

End If

End Sub

Private Sub cCMD_Click(Index As Integer)
Dim i As Long, oItem As ListItem, j As Long
Select Case Index
       Case 0
           If Mesh(CurrentMeshID).Edit Then
           
           If MsgBox(ActiveString(PublicMsgs(2), Mesh(CurrentMeshID).strID), vbYesNo + vbExclamation, ActiveString(PublicMsgs(3), PublicMsgs(GetEditorIndex(Me.Tag)))) = vbYes Then
              
              DelIndex Mesh(CurrentMeshID).strID
              If CurrentMeshID < N_Mesh - 1 Then
                For i = CurrentMeshID To N_Mesh - 2 Step 1
                    ChangeID Mesh(i + 1).strID, Mesh(i + 1).ID - 1
                    j = Mesh(i).ID
                    Mesh(i) = Mesh(i + 1)
                    Mesh(i).ID = j
                    LstMesh.ListItems(i + 1).SubItems(1) = LstMesh.ListItems(i + 2).SubItems(1)
 
                Next i
                
                ReDim Preserve Mesh(N_Mesh - 2)
                LstMesh.ListItems.Remove N_Mesh
                N_Mesh = N_Mesh - 1
                
              Else
                ReDim Preserve Mesh(N_Mesh - 2)
                LstMesh.ListItems.Remove N_Mesh
                
                N_Mesh = N_Mesh - 1
                CurrentMeshID = N_Mesh - 1
                
              End If
               
               LstMesh_ItemClick LstMesh.ListItems(CurrentMeshID + 1)
               LstMesh.ListItems(CurrentMeshID + 1).Selected = True
               LstMesh.ListItems(CurrentMeshID + 1).EnsureVisible
           
           End If
           Else
              MsgBox ActiveString(PublicMsgs(4), Mesh(CurrentMeshID).strID), vbCritical, ActiveString(PublicMsgs(3), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
       Case 1
         If MsgBox(ActiveString(PublicMsgs(5), Mesh(CurrentMeshID).strID, PublicEditors(GetEditorIndex(Me.Tag))), vbYesNo + vbInformation, ActiveString(PublicMsgs(9), PublicEditors(GetEditorIndex(Me.Tag)))) = vbYes Then
 
           If AddIndex(N_Mesh, Mesh(CurrentMeshID).strID & "_New") Then
           ReDim Preserve Mesh(N_Mesh)
           N_Mesh = N_Mesh + 1
           Mesh(N_Mesh - 1) = Mesh(CurrentMeshID)
           With Mesh(N_Mesh - 1)
                 .ID = N_Mesh - 1
                 .strID = .strID & "_New"
                 .Edit = True
           End With
           
           Set oItem = LstMesh.ListItems.Add(, "Mesh_" & Mesh(N_Mesh - 1).ID, Mesh(N_Mesh - 1).ID)
           
                 With oItem
                    .SubItems(1) = Mesh(N_Mesh - 1).strID
                 End With
                 
           LstMesh_ItemClick LstMesh.ListItems(N_Mesh)
           LstMesh.ListItems(N_Mesh).Selected = True
           LstMesh.ListItems(N_Mesh).EnsureVisible
           
           Else
           
           MsgBox ActiveString(PublicMsgs(90), Mesh(CurrentMeshID).strID & "_New"), vbCritical, PublicMsgs(19)
           End If
         End If
         
      Case 2
         If CurrentMeshID > 0 Then
           If Mesh(CurrentMeshID - 1).Edit And Mesh(CurrentMeshID).Edit Then
                SwapID Mesh(CurrentMeshID - 1).strID, Mesh(CurrentMeshID).strID
                SwapMesh CurrentMeshID - 1, CurrentMeshID
                SwapListItem LstMesh.ListItems(CurrentMeshID), LstMesh.ListItems(CurrentMeshID + 1), 1, True
                
               LstMesh_ItemClick LstMesh.ListItems(CurrentMeshID)
               LstMesh.ListItems(CurrentMeshID + 1).Selected = True
           Else
            MsgBox ActiveString(PublicMsgs(7), Mesh(CurrentMeshID - 1).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
         
         End If
      Case 3
        If CurrentMeshID + 1 <= N_Mesh - 1 Then
           If Mesh(CurrentMeshID).Edit And Mesh(CurrentMeshID + 1).Edit Then
                SwapID Mesh(CurrentMeshID).strID, Mesh(CurrentMeshID + 1).strID
                SwapMesh CurrentMeshID, CurrentMeshID + 1
                SwapListItem LstMesh.ListItems(CurrentMeshID + 1), LstMesh.ListItems(CurrentMeshID + 2), 1, True
                
                LstMesh_ItemClick LstMesh.ListItems(CurrentMeshID + 2)
                LstMesh.ListItems(CurrentMeshID + 1).Selected = True
           Else
              MsgBox ActiveString(PublicMsgs(7), Mesh(CurrentMeshID).strID), vbCritical, ActiveString(PublicMsgs(8), PublicEditors(GetEditorIndex(Me.Tag)))
           End If
           
        End If
End Select
End Sub

Private Sub COutputLine_Click()
frmLine.ShowTxtLine Me.Tag, -1
End Sub

Private Sub CQuery_Click(Index As Integer)
Dim q As Boolean
q = QueryItem(LstMesh, LstMesh.SelectedItem.Index, txtQuery.Text, CBool(Index))

If q Then
   Call LstMesh_ItemClick(LstMesh.SelectedItem)
End If
End Sub

Private Sub CReset_Click()
If MsgBox(PublicMsgs(10), vbYesNo + vbDefaultButton2 + vbExclamation, PublicMsgs(0)) = vbYes Then
LstMesh_ItemClick LstMesh.ListItems(CurrentMeshID + 1)
LstMesh.ListItems(CurrentMeshID + 1).Selected = True
LstMesh.ListItems(CurrentMeshID + 1).EnsureVisible
End If
End Sub

Private Sub Form_Load()
'Inits
CustomActive = False



InitMeshListView
LoadMeshList
InitFrames
InitLstFlags

LstMesh.ListItems(1).Selected = True
CurrentMeshID = 0
CurrentMesh = Mesh(0)
Call LoadMeshInfo(CurrentMesh)

TranslateForm Me
InitCMDs
Label2.Caption = Label2.Caption & N_Mesh

CustomActive = True
End Sub

Private Sub InitCMDs()
Dim i As Integer

For i = 1 To cCMD.UBound
    cCMD(i).Left = cCMD(i - 1).Left - cCMD(i).Width - 135
Next i

End Sub

Private Sub LstFlags_LostFocus()
Dim i As Integer, tI As Integer64b

With CurrentMesh
    'CurrentMesh.Flags = ""
    For i = 0 To UBound(MeshFlag)
          If LstFlags.Selected(i) Then
             tI = AddFlagsI64(tI, HexStrToI64(MeshFlag(i).X))
          End If
    Next i
    .Flags = I64toStrNZ(tI)
End With

End Sub

Private Sub LstMesh_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim n As Integer

CurrentMeshID = Val(Item.Text)
CurrentMesh = Mesh(CurrentMeshID)

LoadMeshInfo CurrentMesh

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


Private Sub InitFrames()
Dim i As Integer

For i = 0 To FraProps.UBound
    With FraProps(i)
         .BorderStyle = 0
         .Top = Tab1.ClientTop
         .Left = Tab1.ClientLeft
         .Width = Tab1.ClientWidth
         .Height = Tab1.ClientHeight
         .ZOrder
           If i <> 0 Then
            .Visible = False
           End If
    End With
Next i

'If LCase$(MnBInfo.Language) = "en" Then
'     Text1(2).Enabled = False
'End If

End Sub


Private Sub Tab1_Click()
Dim i As Integer, n As Integer

If CustomActive Then
For i = 0 To FraProps.UBound
    With FraProps(i)
         .Visible = i + 1 = Tab1.SelectedItem.Index
         If i = 1 Then
             'InitPropFrames
         End If
    End With
Next i

LoadMeshInfo CurrentMesh
End If

End Sub

Private Sub InitLstFlags()
Dim i As Integer

LstFlags.Clear
For i = 0 To UBound(MeshFlag)
    LstFlags.AddItem MeshFlag(i).Z
Next i

End Sub

'*************************************************************************
'**函 数 名：LoadMeshInfo
'**输    入：Mesh As Type_Particle_System
'**输    出：无
'**功能描述：载入网格模型信息
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-3 10:13:55
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadMeshInfo(Mesh As Type_Mesh)
Dim i As Integer, tStr As String, tI As Integer64b, H As Byte

CustomActive = False
TxtMeshName.Text = ""

'Select Case Tab1.SelectedItem.Index

'Case 1
    '显示信息
    TxtMeshName.Text = Mesh.strID
    TxtResourceName.Text = Mesh.Resource_Name
    
    'Flags
    For i = 0 To LstFlags.ListCount - 1
         LstFlags.Selected(i) = False
    Next i
    LoadLstFlags Mesh.Flags

    '属性
    txtTranslation(0).Text = Format(Mesh.Translation.X, "0.000000")
    txtTranslation(1).Text = Format(Mesh.Translation.Y, "0.000000")
    txtTranslation(2).Text = Format(Mesh.Translation.Z, "0.000000")
    
    txtRotation(0).Text = Format(Mesh.Rotation_Angle.X, "0.000000")
    txtRotation(1).Text = Format(Mesh.Rotation_Angle.Y, "0.000000")
    txtRotation(2).Text = Format(Mesh.Rotation_Angle.Z, "0.000000")
    
    txtScale(0).Text = Format(Mesh.Scale.X, "0.000000")
    txtScale(1).Text = Format(Mesh.Scale.Y, "0.000000")
    txtScale(2).Text = Format(Mesh.Scale.Z, "0.000000")
 
'End Select

CustomActive = True
End Sub

'*************************************************************************
'**函 数 名：LoadLstFlags
'**输    入：-(String)Flags
'**输    出：-
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：Ser_Charles
'**日    期：2011-1-29 20:51:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadLstFlags(Flags As String)
Dim tI As Integer64b, i As Byte, tI2 As Integer64b, n As Integer
         tI = StrToI64(Flags)
    With LstFlags
        For i = 0 To LstFlags.ListCount - 1
            tI2 = And64b(tI, HexStrToI64(MeshFlag(i).X))
            If I64toStrNZ(tI2) = I64toStrNZ(HexStrToI64(MeshFlag(i).X)) Then
                LstFlags.Selected(i) = True
            End If
        Next i
    End With
End Sub

Public Sub ReLoadInfo()
LoadMeshInfo CurrentMesh
End Sub

Private Sub TxtMeshName_LostFocus()

TxtMeshName.Text = Replace(TxtMeshName.Text, " ", "_")
CurrentMesh.strID = CStr(TxtMeshName.Text)

End Sub


Private Sub TxtResourceName_LostFocus()

TxtResourceName.Text = Replace(TxtResourceName.Text, " ", "_")
CurrentMesh.Resource_Name = CStr(TxtResourceName.Text)

End Sub


Private Sub txtRotation_LostFocus(Index As Integer)

Select Case Index
       Case 0
           CurrentMesh.Rotation_Angle.X = Val(txtRotation(Index).Text)
       Case 1
           CurrentMesh.Rotation_Angle.Y = Val(txtRotation(Index).Text)
       Case 2
           CurrentMesh.Rotation_Angle.Z = Val(txtRotation(Index).Text)
End Select

End Sub

Private Sub txtScale_LostFocus(Index As Integer)

Select Case Index
       Case 0
           CurrentMesh.Scale.X = Val(txtScale(Index).Text)
       Case 1
           CurrentMesh.Scale.Y = Val(txtScale(Index).Text)
       Case 2
           CurrentMesh.Scale.Z = Val(txtScale(Index).Text)
End Select

End Sub

Private Sub txtTranslation_LostFocus(Index As Integer)

Select Case Index
       Case 0
           CurrentMesh.Translation.X = Val(txtTranslation(Index).Text)
       Case 1
           CurrentMesh.Translation.Y = Val(txtTranslation(Index).Text)
       Case 2
           CurrentMesh.Translation.Z = Val(txtTranslation(Index).Text)
End Select

End Sub
