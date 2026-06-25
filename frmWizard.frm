VERSION 5.00
Begin VB.Form frmWizard 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "物品导入向导"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13725
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13725
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      Height          =   7335
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   10335
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frame2(4)"
         Height          =   7335
         Index           =   4
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   10335
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   9
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   28
            Text            =   "frmWizard.frx":08CA
            Top             =   360
            Width           =   9735
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5895
            Index           =   8
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   27
            Text            =   "frmWizard.frx":08D1
            Top             =   1080
            Width           =   9735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frame2(3)"
         Height          =   3495
         Index           =   3
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   6615
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5895
            Index           =   6
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   25
            Text            =   "frmWizard.frx":08D8
            Top             =   1080
            Width           =   9735
         End
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   7
            Left            =   240
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   24
            Text            =   "frmWizard.frx":08DF
            Top             =   360
            Width           =   9735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frame2(0)"
         Height          =   5535
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   7815
         Begin VB.FileListBox File2 
            Height          =   3330
            Index           =   1
            Left            =   6840
            Pattern         =   "*.dds"
            TabIndex        =   12
            Top             =   2280
            Width           =   2775
         End
         Begin VB.FileListBox File2 
            Height          =   3330
            Index           =   0
            Left            =   3960
            Pattern         =   "*.brf"
            TabIndex        =   11
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Index           =   0
            Left            =   360
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "frmWizard.frx":08E6
            Top             =   360
            Width           =   9375
         End
         Begin VB.DirListBox Dir1 
            Height          =   3450
            Left            =   360
            TabIndex        =   9
            Top             =   2280
            Width           =   3375
         End
         Begin VB.DriveListBox Drive1 
            Height          =   300
            Left            =   360
            TabIndex        =   8
            Top             =   1920
            Width           =   3375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "贴图文件(*.dds)[未找到]:"
            Height          =   180
            Index           =   1
            Left            =   6840
            TabIndex        =   14
            Top             =   2040
            Width           =   2160
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "模型文件(*.brf)[未找到]:"
            Height          =   180
            Index           =   0
            Left            =   3960
            TabIndex        =   13
            Top             =   2040
            Width           =   2160
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frame2(2)"
         Height          =   3135
         Index           =   2
         Left            =   720
         TabIndex        =   18
         Top             =   840
         Width           =   5055
         Begin VB.TextBox Text2 
            Height          =   375
            Index           =   5
            Left            =   8880
            TabIndex        =   21
            Top             =   6720
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   360
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   20
            Text            =   "frmWizard.frx":08ED
            Top             =   360
            Width           =   9375
         End
         Begin VB.TextBox Text2 
            Height          =   5535
            Index           =   3
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   19
            Top             =   1080
            Width           =   9615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "语言:"
            Height          =   180
            Index           =   0
            Left            =   8400
            TabIndex        =   22
            Top             =   6840
            Width           =   450
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frame2(1)"
         Height          =   3735
         Index           =   1
         Left            =   2040
         TabIndex        =   15
         Top             =   960
         Width           =   6255
         Begin VB.TextBox Text2 
            Height          =   5775
            Index           =   2
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   17
            Top             =   1080
            Width           =   9615
         End
         Begin VB.TextBox Text2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   360
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   16
            Text            =   "frmWizard.frx":08F4
            Top             =   360
            Width           =   9375
         End
      End
   End
   Begin VB.PictureBox PicButton 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   13665
      TabIndex        =   1
      Top             =   8670
      Width           =   13725
      Begin VB.CommandButton CFore 
         Caption         =   "<上一步(&F)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6000
         TabIndex        =   6
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton CNext 
         BackColor       =   &H00C0FFC0&
         Caption         =   "下一步>(&N)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   2415
      End
      Begin VB.CommandButton CCancel 
         BackColor       =   &H00C0C0FF&
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
         Height          =   615
         Left            =   10920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox PicBanner 
      Align           =   3  'Align Left
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8670
      Left            =   0
      Picture         =   "frmWizard.frx":08FB
      ScaleHeight     =   8610
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   0
      Width           =   2685
   End
   Begin VB.Label LCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public wTag As Long
Dim Path As String, Language As String
Dim Choice As Integer

Private Sub CCancel_Click()
UnLoad Me
End Sub

Private Sub CFore_Click()
   CFore.Enabled = GoFore()
   
   LoadWizard
End Sub

Private Sub CNext_Click()
Dim q As Boolean, n As Long, s As Long

Select Case wTag
       Case Tag_Item
            If Pro.NodeNow = 1 Then                  'txt行
               q = IsEmptyStr(Text2(2).Text)
               
               If Not q Then
                     Choice = 0       '有txt行
                     ReadItems Text2(2).Text
                     SetItemsDefaultCSV
               Else
                 If MsgBox(PublicWizards(17), vbYesNo + vbExclamation, PublicTools(3)) = vbYes Then
                   Choice = 1       '无txt行
                   TemItemCount = 0
                 Else
                     Exit Sub
                 End If
               End If
               
            ElseIf Pro.NodeNow = 0 Then              '模型、贴图
               Path = Dir1.Path
               
            ElseIf Pro.NodeNow = 2 Then              '语言信息
               
               If IsEmptyStr(Text2(5).Text) Then
                  Language = MnBInfo.Language
               Else
                  Language = Text2(5).Text
               End If
               
               n = ReadItemCSVLine(Text2(3).Text)
               If n <= 0 Then
                  If MsgBox(PublicWizards(18), vbYesNo + vbExclamation, PublicTools(3)) = vbNo Then
                      Exit Sub
                  End If

               End If
               
            ElseIf Pro.NodeNow = 3 Then              '导入开始!
               
            ElseIf Pro.NodeNow = 4 Or Pro.NodeNow = 5 Then              '导入完成
               UnLoad Me
               Exit Sub
               
            End If
            
            Pro.ForeNode = Pro.ForeNode & "|" & Pro.NodeNow
            
            Pro.NodeNow = ItemNodes(Pro.NodeNow).NextID(Choice)
            Pro.DefaultChoice = Choice
            CFore.Enabled = True
End Select

LoadWizard
End Sub

Private Sub Dir1_Change()
On Error Resume Next
Dim q As Boolean

SwtichNext False

If Not DirExists(Dir1.Path & "\Textures") Then                      '贴图
    Label2(1).Caption = PublicWizards(15) & "(*.dds):" & PublicWizards(14) & ":"
Else
    File2(1).Path = Dir1.Path & "\Textures"
    Label2(1).Caption = PublicWizards(15) & "(*.dds):"
End If

If Not DirExists(Dir1.Path & "\Resource") Then                      '模型
    Label2(0).Caption = PublicWizards(16) & "(*.brf):" & PublicWizards(14) & ":"
Else
    File2(0).Path = Dir1.Path & "\Resource"
    Label2(0).Caption = PublicWizards(16) & "(*.brf):"
    
    If File2(0).ListCount > 0 Then
       SwtichNext True
    End If
End If

End Sub

Private Sub Drive1_Change()
On Error GoTo Errline
Dir1.Path = Drive1.Drive

Exit Sub

Errline:
MsgBox Err.Description, vbCritical, PublicMsgs(89)
End Sub



Private Sub Form_Load()

InitFrame
InitItemNodes

SwtichFore False
LoadWizard

Me.Caption = PublicTools(3)
CNext.Caption = PublicWizards(0)
CFore.Caption = PublicWizards(1)
CCancel.Caption = PublicWizards(2)
CCancel.Enabled = True
End Sub

Private Sub InitFrame()
Dim i As Integer

Frame1.BorderStyle = 0

For i = Frame2.LBound To Frame2.UBound
    Frame2(i).Visible = False
    Frame2(i).BorderStyle = 0
Next i

End Sub

Private Sub ShowFrame2(ByVal Index As Integer)
Dim i As Integer

For i = Frame2.LBound To Frame2.UBound
    Frame2(i).Visible = i = Index
Next i

Frame2(Index).Move 0, 0, Frame1.Width, Frame1.Height
End Sub

Private Sub LoadWizard()
Dim ErrReport(1) As Type_ErrReport, p(1) As Boolean

If Pro.NodeNow = -1 Then Exit Sub

SwtichNext True

Choice = 0
If Choice >= 0 And Choice <= UBound(ItemNodes(Pro.NodeNow).NextID) Then
   If ItemNodes(Pro.NodeNow).AllowDefault Then
      Choice = Pro.DefaultChoice
   End If
End If

ShowFrame2 ItemNodes(Pro.NodeNow).Frame_Idx

Select Case wTag
       Case Tag_Item
            If Pro.NodeNow = 1 Then
               LCaption.Caption = PublicWizards(7)               '导入物品txt信息(item_kinds1.txt)
               Text2(1).Text = PublicWizards(19)
               Text2(2).SetFocus
               
            ElseIf Pro.NodeNow = 0 Then
               LCaption.Caption = PublicWizards(8)               '选择需导入装备的文件夹
               Text2(0).Text = PublicWizards(20) & PublicWizards(21) & PublicWizards(22) & PublicWizards(23) & PublicWizards(24)
               
               Call Dir1_Change
               
            ElseIf Pro.NodeNow = 2 Then
               LCaption.Caption = PublicWizards(9)                '导入物品语言信息(item_kinds.csv)
               Text2(4).Text = PublicWizards(25) & PublicWizards(26)
               Text2(5).Text = MnBInfo.Language
               Text2(3).SetFocus
               
            ElseIf Pro.NodeNow = 3 Then
               LCaption.Caption = PublicWizards(10)               '确定导入物品
               Text2(7).Text = PublicWizards(27) & vbCrLf & vbCrLf
                Text2(6).Text = ""
                WriteItemReports Text2(6)
                Text2(6).SetFocus
                
            ElseIf Pro.NodeNow = 4 Or Pro.NodeNow = 5 Then
               
               frmTip.ShowTip PublicTips(5)
               p(0) = ImportItems(ErrReport(0))
               p(1) = ImportResource(ErrReport(1))
               frmTip.HideTip
               
               If p(0) And p(1) Then
                  LCaption.Caption = PublicWizards(4)             '导入成功
                  If Pro.NodeNow = 4 Then
                     Text2(9).Text = PublicWizards(28) & PublicWizards(29)
                  Else
                     Text2(9).Text = PublicWizards(28) & PublicWizards(30)
                  End If
                  
                  SwtichFore False
                  CNext.Caption = PublicWizards(3)
                  CCancel.Enabled = False
                  Text2(8).Visible = False
               Else
                  LCaption.Caption = PublicWizards(5)             '导入失败
                  Text2(9).Text = PublicWizards(31)
                  Text2(8).Text = PublicWizards(6) & ":" & vbCrLf & _
                                   "(" & ErrReport(0).Number & ")" & ErrReport(0).Description & vbCrLf & _
                                   "(" & ErrReport(1).Number & ")" & ErrReport(1).Description
                  Text2(8).SetFocus
               End If

            End If
End Select

End Sub

Private Sub SwtichNext(ByVal Switch As Boolean)
CNext.Enabled = Switch
End Sub

Private Sub SwtichFore(ByVal Switch As Boolean)
CFore.Enabled = Switch
End Sub

Private Sub ReadItems(ByVal Text As String)
Dim n As Long, lP As Long

n = -1
lP = 1
ReDim TemItems(0)

Do While lP <= Len(Text)
   n = n + 1
   ReDim Preserve TemItems(n)
   
   ReadItemLine Text, TemItems(n), lP
Loop

If Trim(TemItems(n).dbName) = "" Then
    n = n - 1
End If

TemItemCount = n + 1

End Sub

Private Sub WriteItemReports(Board As TextBox)
Dim i As Long

If TemItemCount > 0 Then
Board.Text = Board.Text & ActiveString(PublicWizards(11), TemItemCount) & vbCrLf

For i = 0 To TemItemCount - 1
   With TemItems(i)
       Board.Text = Board.Text & .dbName & "|" & .disname & "|" & .csvName & vbCrLf
       
   End With
Next i
End If

Board.Text = Board.Text & vbCrLf & ActiveString(PublicWizards(12), File2(0).ListCount) & vbCrLf

For i = 0 To File2(0).ListCount - 1
       Board.Text = Board.Text & File2(0).List(i) & vbCrLf
Next i


Board.Text = Board.Text & vbCrLf & ActiveString(PublicWizards(13), File2(1).ListCount) & vbCrLf

For i = 0 To File2(1).ListCount - 1
       Board.Text = Board.Text & File2(1).List(i) & vbCrLf
Next i

End Sub

Private Function ImportItems(ErrInfo As Type_ErrReport) As Boolean
On Error GoTo Errline

Dim i As Long

If TemItemCount > 0 Then

ReDim Preserve itm(N_Item + TemItemCount - 1)
itm(N_Item + TemItemCount - 1) = itm(N_Item - 1)
itm(N_Item + TemItemCount - 1).ID = N_Item + TemItemCount - 1

N_Item = N_Item + TemItemCount

For i = 0 To TemItemCount - 1
    TemItems(i).Edit = True
    itm(i + N_Item - TemItemCount - 1) = TemItems(i)
    itm(i + N_Item - TemItemCount - 1).ID = i + N_Item - TemItemCount - 1
    AddIndex i + N_Item - TemItemCount - 1, itm(i + N_Item - TemItemCount - 1).dbName
Next i

End If

ImportItems = True
Exit Function

Errline:
ImportItems = False
ErrInfo.Number = Err.Number
ErrInfo.Description = Err.Description


End Function


Private Function ImportResource(ErrInfo As Type_ErrReport) As Boolean
On Error GoTo Errline
Dim i As Long, F As Integer, ResourceName As String, j As Long, q As Boolean

F = FreeFile
Open MnBInfo.ModIniFileName For Append As #F
  Print #F, ""
  
  For i = 0 To File2(0).ListCount - 1
    FileCopy File2(0).Path & "\" & File2(0).List(i), MnBInfo.ModPath & "\Resource\" & File2(0).List(i)
    
    ResourceName = Left(File2(0).List(i), Len(File2(0).List(i)) - 4)
    
    q = True
    For j = 1 To ModSets.ModResourceCount
      If LCase(ResourceName) = LCase(ModSets.ModResource(j)) Then
        q = False
        Exit For
      End If
    Next j
    
    If q Then
       Print #F, "load_module_resource = " & ResourceName
       
       ModSets.ModResourceCount = ModSets.ModResourceCount + 1
       
       ReDim Preserve ModSets.ModResource(ModSets.ModResourceCount)
       
       ModSets.ModResource(ModSets.ModResourceCount) = ResourceName
    End If
    
    DoEvents
  Next i
Close #F

For i = 0 To File2(1).ListCount - 1
    FileCopy File2(1).Path & "\" & File2(1).List(i), MnBInfo.ModPath & "\Textures\" & File2(1).List(i)
    DoEvents
Next i


ImportResource = True
Exit Function

Errline:
ImportResource = False
ErrInfo.Number = Err.Number
ErrInfo.Description = Err.Description

End Function
