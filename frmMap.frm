VERSION 5.00
Begin VB.Form frmMap 
   BackColor       =   &H00FFFFFF&
   Caption         =   "卡拉迪亚地图"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12435
   FillStyle       =   0  'Solid
   Icon            =   "frmMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   571
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   829
   StartUpPosition =   2  '屏幕中心
   Tag             =   "tool_1"
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   827
      TabIndex        =   8
      Top             =   8160
      Width           =   12435
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "坐标:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox PicBox 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   8160
      Left            =   9360
      MousePointer    =   2  'Cross
      ScaleHeight     =   542
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   203
      TabIndex        =   1
      Top             =   0
      Width           =   3075
      Begin VB.PictureBox PicFrame 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1920
         Index           =   0
         Left            =   0
         MousePointer    =   2  'Cross
         ScaleHeight     =   126
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   203
         TabIndex        =   10
         Top             =   6240
         Width           =   3075
         Begin VB.CommandButton CCMD 
            BackColor       =   &H00C0C0C0&
            Caption         =   "撤销(&B)"
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
            Index           =   0
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton CCMD 
            BackColor       =   &H000000C0&
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
            Index           =   1
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   960
            Width           =   2655
         End
      End
      Begin VB.CheckBox chkInvertSelection 
         BackColor       =   &H00E0E0E0&
         Caption         =   "反选(&S)"
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
         Left            =   120
         TabIndex        =   7
         Top             =   3480
         Width           =   2775
      End
      Begin VB.CheckBox chkLogic 
         BackColor       =   &H00E0E0E0&
         Caption         =   "反向过滤(&I)"
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Width           =   2775
      End
      Begin VB.CheckBox chkLogic 
         BackColor       =   &H00E0E0E0&
         Caption         =   "存在一项即匹配(&A)"
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
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.ListBox LstFilter 
         Appearance      =   0  'Flat
         Height          =   2970
         Index           =   1
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ListBox LstFilter 
         Appearance      =   0  'Flat
         Height          =   2970
         Index           =   0
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "过滤器:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   690
      End
   End
   Begin VB.Timer CPU 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8880
      Top             =   3120
   End
   Begin VB.Timer Painter 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8880
      Top             =   2640
   End
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   479
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   551
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MosP As tPoint, AbMosP As tPoint
Dim RedrawRequest As Boolean
Dim CustomActive As Boolean
Dim CurrentMOID As Long, CurrentMO As Type_MapObject, LastMOID As Long


Private Sub cCMD_Click(Index As Integer)
Select Case Index
    Case 0
      If LastMOID > -1 Then
         MapObj(LastMOID) = CurrentMO
      End If
    
    Case 1
      If MsgBox(PublicMsgs(1), vbYesNo + vbDefaultButton2 + vbInformation, PublicMsgs(0)) = vbYes Then
         ApplyMapObjects
         MsgBox "确定成功!", vbInformation, PublicMsgs(0)
      End If
      
End Select

RedrawRequest = True
End Sub

Private Sub chkInvertSelection_Click()
Dim i As Integer
CustomActive = False

For i = 0 To LstFilter(0).ListCount - 1
     LstFilter(0).Selected(i) = Not LstFilter(0).Selected(i)
Next i

CustomActive = True

Call LstFilter_ItemCheck(0, 0)

End Sub

Private Sub chkLogic_Click(Index As Integer)
LstFilter_ItemCheck 0, LstFilter(0).ListIndex
End Sub

Private Sub CPU_Timer()

CustomActive = False
With Custom
      If .Action.Scroll.X <> 0 Or .Action.Scroll.Y <> 0 Or .Action.dSCL <> 1 Then
         RedrawRequest = True
      End If
      .ViewP = PointAdd(.Action.Scroll, .ViewP)
      .SCL = .SCL * .Action.dSCL
End With

CustomActive = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
With Custom.Action
Select Case KeyCode
      Case vbKeyW
           .Scroll.Y = 1
      Case vbKeyS
           .Scroll.Y = -1
      Case vbKeyA
           .Scroll.X = -1
      Case vbKeyD
           .Scroll.X = 1
      Case vbKeyZ
           .dSCL = 1.03
      Case vbKeyX
           .dSCL = 0.97
End Select
End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
With Custom.Action
Select Case KeyCode
      Case vbKeyW
           .Scroll.Y = 0
      Case vbKeyS
           .Scroll.Y = 0
      Case vbKeyA
           .Scroll.X = 0
      Case vbKeyD
           .Scroll.X = 0
      Case vbKeyZ
           .dSCL = 1
      Case vbKeyX
           .dSCL = 1
End Select
End With
End Sub

Private Sub Form_Load()
CustomActive = False
CurrentMOID = -1
LastMOID = -1

InitModMap
InitFlagsList
Custom.SCL = 2

TranslateForm Me
'PicMap.BackColor = vbGreen
Me.Show
CntP = SetPoint(PicMap.ScaleWidth / 2, PicMap.ScaleHeight / 2)

RedrawRequest = True
Painter.Enabled = True
CPU.Enabled = True

CustomActive = True
End Sub

Private Sub Form_Resize()
PicMap.Move 0, 0, Me.ScaleWidth - PicBox.Width, Me.ScaleHeight

CntP = SetPoint(PicMap.ScaleWidth / 2, PicMap.ScaleHeight / 2)

RedrawRequest = True
End Sub

'*************************************************************************
'**函 数 名：LoadMap
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-09 15:08:00
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadMap()
Dim tP(3) As tPoint, l As Single, dL As Single, DrawDegree As Single, i As Long
dL = MO_Radius_Medium
For i = 0 To N_MOs - 1
  With MapObj(i)
      If .Visible Then
         tP(0) = SetPoint(.Body.X, .Body.Y)
         tP(1) = GetRelativePoint(CntP, tP(0), Custom.ViewP, Custom.SCL)
         
         tP(2) = SetPoint(tP(1).X - dL, tP(1).Y - dL - PicMap.TextHeight(.Label))
         
         tP(3) = PoltoRec(dL, .Degree - Pi / 2, tP(1))
               
         'If PicMap.DrawWidth <> 2 Then PicMap.DrawWidth = 2
         If PicMap.ForeColor <> .lColor Then PicMap.ForeColor = .lColor
         If PicMap.FillStyle <> 1 Then PicMap.FillStyle = 1
         PicMap.Circle (tP(1).X, tP(1).Y), dL
         PicMap.Line (tP(1).X, tP(1).Y)-(tP(3).X, tP(3).Y)
         
         If PicMap.CurrentX <> tP(2).X Then PicMap.CurrentX = tP(2).X
         If PicMap.CurrentY <> tP(2).Y Then PicMap.CurrentY = tP(2).Y
         If PicMap.FontSize <> .LabelSize Then PicMap.FontSize = .LabelSize
         PicMap.Print .Label
         'If .LabelSize = 15 Then MsgBox "a"
      End If
  End With
Next i

End Sub

Private Sub LstFilter_ItemCheck(Index As Integer, Item As Integer)
If CustomActive Then
    If Index = 0 Then
        LoadFlagsList chkLogic(0).Value, chkLogic(1).Value
        RedrawRequest = True
    End If
End If
End Sub

Private Sub Painter_Timer()

If RedrawRequest Then
    'Reset Paper
    PicMap.Cls
    
    'Draw
    LoadMap
    
    If CurrentMOID > -1 Then
      With MapObj(CurrentMOID)
       PicMap.CurrentX = MosP.X
       PicMap.CurrentY = MosP.Y
       PicMap.ForeColor = vbBlack
       PicMap.Print "[" & Round(CDbl(.Body.X), 2) & "," & Round(CDbl(.Body.Y), 2) & "," & Round(CDbl(RadToDeg(FuncDegreeStandardize2(.Degree))), 2) & "°]"
      End With
    End If
    
    'Frame Work
    With PicFrame(0)
      .Left = 0
      .Top = PicBox.ScaleHeight - .Height
      .Width = PicBox.ScaleWidth
    End With
    
    If RedrawRequest Then RedrawRequest = False
End If

Label1(1).Caption = PublicMsgs(55) & ":[" & Round(CDbl(AbMosP.X), 2) & " , " & Round(CDbl(AbMosP.Y), 2) & "]"


       Label1(1).Caption = Label1(1).Caption

End Sub

Private Sub PicMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tP As tPoint
MosP = SetPoint(X, Y)
AbMosP = GetAbsolutePoint(Custom.ViewP, MosP, CntP, Custom.SCL)     '获得鼠标在卡拉迪亚的坐标

If Button = vbLeftButton Then
    CurrentMOID = GetClickedMapObj(MosP)
    If CurrentMOID > -1 Then
       CurrentMO = MapObj(CurrentMOID)
       LastMOID = CurrentMOID
    End If
End If

End Sub

Private Sub PicMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dP As tPoint, TemD As Single

'Drag Map
If Button = vbRightButton Then
    dP = SetPoint(X - MosP.X, Y - MosP.Y)
    Custom.ViewP = SetPoint(Custom.ViewP.X - dP.X / Custom.SCL, Custom.ViewP.Y + dP.Y / Custom.SCL)
    RedrawRequest = True
    
    PicMap.MousePointer = 5
End If
MosP = SetPoint(X, Y)

'Drag MapObject
AbMosP = GetAbsolutePoint(Custom.ViewP, MosP, CntP, Custom.SCL)     '获得鼠标在卡拉迪亚的坐标

If Button = vbLeftButton Then
    If CurrentMOID > -1 Then
        If Shift = 0 Then
           MapObj(CurrentMOID).Body = AbMosP
        ElseIf Shift = 2 Then
           TemD = GetDegree(AbMosP, MapObj(CurrentMOID).Body)
           MapObj(CurrentMOID).Degree = Pi / 2 - TemD
        End If
        
        RedrawRequest = True
    End If
End If


End Sub

Private Sub PicMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MosP = SetPoint(X, Y)
    AbMosP = GetAbsolutePoint(Custom.ViewP, MosP, CntP, Custom.SCL)     '获得鼠标在卡拉迪亚的坐标
    
    If Button = vbLeftButton Then
        CurrentMOID = -1
    End If
  PicMap.MousePointer = 2
  RedrawRequest = True
End Sub

'*************************************************************************
'**函 数 名：InitFlagsList
'**输    入：
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 21:22:43
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub InitFlagsList()
Dim i As Integer, strTemArray() As Variant, TemArray() As Variant
strTemArray = Array("不可用", "船", "静态", "小标签", "中标签", "大标签", "总是可见", "行为默认", "在城镇中自动去除", "任务特设", _
               "无标签", "人数有限", "藏身处", "显示阵营", "不可见", "不攻击平民", "平民")
TemArray = Array(pf_disabled, pf_is_ship, pf_is_static, pf_label_small, pf_label_medium, pf_label_large, pf_always_visible, pf_default_behavior, pf_auto_remove_in_town, pf_quest_party, _
               pf_no_label, pf_limit_members, pf_hide_defenders, pf_show_faction, pf_is_hidden, pf_dont_attack_civilians, pf_civilian)
         
LstFilter(0).Clear
LstFilter(1).Clear
         For i = 0 To UBound(strTemArray)
            If i < 3 Or i > 5 Then
               LstFilter(0).AddItem strTemArray(i)
               LstFilter(0).Selected(LstFilter(0).ListCount - 1) = True
               LstFilter(1).AddItem TemArray(i)
            End If
         Next i

End Sub

'*************************************************************************
'**函 数 名：LoadFlagsList
'**输    入：(Long)Logic,(Long)Invert
'**输    出：
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2010-12-11 21:28:57
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Private Sub LoadFlagsList(ByVal Logic As Long, ByVal Invert As Long)
Dim i As Integer, j As Long, fI64 As Integer64b, fI64_NOW As Integer64b, resI64 As Integer64b

CustomActive = False
For j = 0 To N_MOs - 1
With MapObj(j)

   If Invert = 0 Then
     .Visible = False
   ElseIf Invert = 1 Then
     .Visible = True
   End If
   
fI64_NOW = StrToI64(.Flags)

   For i = 0 To LstFilter(0).ListCount - 1

        If LstFilter(0).Selected(i) Then
               fI64 = HexStrToI64(LstFilter(1).List(i))
               resI64 = And64b(fI64_NOW, fI64)

               If Not IsEqual64b(resI64, fI64) Then
                 If Invert = 0 Then
                   .Visible = False
                 ElseIf Invert = 1 Then
                   .Visible = True
                 End If
                 
                 If Logic = 0 Then
                    Exit For
                 End If
               Else
                 If Invert = 0 Then
                   .Visible = True
                 ElseIf Invert = 1 Then
                   .Visible = False
                 End If
                 
                 If Logic = 1 Then
                    Exit For
                 End If
               End If
        End If
   Next i


End With
Next j

CustomActive = True
End Sub

'*************************************************************************
'**函 数 名：GetClickedMapObj
'**输    入：(tPoint)MosPos
'**输    出：(Long)
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：
'**日    期：
'**修 改 人：SSgt_Edward
'**日    期：2010-12-11 15:45:02
'**版    本：V1.1321
'*************************************************************************
Private Function GetClickedMapObj(MosPos As tPoint) As Long
Dim i As Long, tP As tPoint, D As Single

GetClickedMapObj = -1
For i = 0 To N_MOs - 1
    With MapObj(i)
       tP = GetRelativePoint(CntP, .Body, Custom.ViewP, Custom.SCL)
       
       D = GetDistance(tP, MosPos)
       
       If D <= MO_Radius_Medium Then
          GetClickedMapObj = i
          Exit For
       End If
    End With
Next i

End Function
