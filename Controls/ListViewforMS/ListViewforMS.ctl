VERSION 5.00
Begin VB.UserControl ListViewforMS 
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   KeyPreview      =   -1  'True
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   ToolboxBitmap   =   "ListViewforMS.ctx":0000
   Begin MnBWarband_Editor.MenuforMS MSMenu 
      Left            =   3960
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   5055
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4215
      Left            =   5040
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox PicFrame 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin MnBWarband_Editor.ComboforOp OpCombo 
         Left            =   3360
         Top             =   3360
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin VB.PictureBox PicBox 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   0
         ScaleHeight     =   191
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   271
         TabIndex        =   1
         Top             =   0
         Width           =   4095
         Begin VB.TextBox txtEnter 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   480
            TabIndex        =   4
            Top             =   120
            Visible         =   0   'False
            Width           =   540
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Menu"
      Begin VB.Menu mEditor 
         Caption         =   "Null"
         Index           =   0
      End
      Begin VB.Menu mEditor 
         Caption         =   "Register"
         Index           =   1
      End
      Begin VB.Menu mEditor 
         Caption         =   "Variable"
         Index           =   2
         Begin VB.Menu mVar 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "String"
         Index           =   3
         Begin VB.Menu mStr 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Item"
         Index           =   4
         Begin VB.Menu mItm 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Troop"
         Index           =   5
         Begin VB.Menu mTrp 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Faction"
         Index           =   6
         Begin VB.Menu mFac 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Quest"
         Index           =   7
         Begin VB.Menu mQst 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Party_Tpl"
         Index           =   8
         Begin VB.Menu mPT 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Party"
         Index           =   9
         Begin VB.Menu mParty 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Scene"
         Index           =   10
         Begin VB.Menu mScene 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Mission_tpl"
         Index           =   11
         Begin VB.Menu mMT 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Menu"
         Index           =   12
         Begin VB.Menu mMnu 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Script"
         Index           =   13
         Begin VB.Menu mScript 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Particle_Sys"
         Index           =   14
         Begin VB.Menu mPSys 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Scene_Prop"
         Index           =   15
         Begin VB.Menu mSP 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Sound"
         Index           =   16
         Begin VB.Menu mSnd 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Local_Variable"
         Index           =   17
         Begin VB.Menu mLvar 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Map_Icon"
         Index           =   18
         Begin VB.Menu mMI 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Skill"
         Index           =   19
         Begin VB.Menu mSkl 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Mesh"
         Index           =   20
         Begin VB.Menu mMesh 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Presentation"
         Index           =   21
         Begin VB.Menu mPrsnt 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Quick_String"
         Index           =   22
         Begin VB.Menu mQStr 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Track"
         Index           =   23
         Begin VB.Menu mTrack 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Tableau"
         Index           =   24
         Begin VB.Menu mTabMat 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "Animation"
         Index           =   25
         Begin VB.Menu mAni 
            Caption         =   "_"
            Index           =   0
         End
      End
      Begin VB.Menu mEditor 
         Caption         =   "End"
         Index           =   26
      End
   End
   Begin VB.Menu mNegs 
      Caption         =   "Neg"
      Begin VB.Menu mNeg 
         Caption         =   "_"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mNeg 
         Caption         =   "_"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mNeg 
         Caption         =   "_"
         Checked         =   -1  'True
         Index           =   2
      End
   End
End
Attribute VB_Name = "ListViewforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ListItems As New ListItemsforMS

Private lIndex As Integer
Private sIndex As Integer
Private AutoRef As Boolean
Private p_Spacing As Single
Private CustomActive As Boolean

Public Event ItemClick(ItemIndex As Integer, SubIndex As Integer)
Public Event FillBlank(ItemIndex As Integer, SubIndex As Integer, TipIndex As Integer)
Public Event ItemChoose(ItemIndex As Integer, SubIndex As Integer, TipIndex As Integer)
Public Event LostSelection()
Public Event Change()

Public Sub Initialize(UserList As ListViewforMS)
ListItems.Initialize UserList

FillMenu
'Hook txtEnter.hWnd
MSMenu.Initialize
OpCombo.Initialize

CustomActive = True
End Sub

Private Sub HScroll1_Change()
PicBox.Left = -HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Call HScroll1_Change
End Sub


Private Sub Label1_Click(Index As Integer)
Dim i As Integer, q As Boolean, Text As String, Tag As String, l As String, r As String

If Not CustomActive Then Exit Sub
'i = lIndex
Call AssignText
For i = 1 To Label1.UBound
  If i = Index Then
    Label1(i).BorderStyle = 1
  Else
    Label1(i).BorderStyle = 0
  End If
Next i

lIndex = GetTagValue(Label1(Index).Tag, 0)
sIndex = GetTagValue(Label1(Index).Tag, 1)
'------------------------------------------
'MsgBox Index & "," & Label1(Index).Tag
'------------------------------------------
If lIndex > 0 Then
  If sIndex > 0 Then
    'Label1(Index).ToolTipText = ListItems(lIndex).SubItems(sIndex).KeyWord
    q = ListItems(lIndex).SubItems(sIndex).Locked
    Text = ListItems(lIndex).SubItems(sIndex).Text
    Tag = ListItems(lIndex).SubItems(sIndex).Value
    
    'Call_script 增添参数
    If ListItems(lIndex).Value = CStr(Call_Script) And sIndex = ListItems(lIndex).SubItems.Count Then
      CustomActive = False
      i = InStrRev(ListItems(lIndex).Text, ListItems(lIndex).SubItems(ListItems(lIndex).SubItems.Count).KeyWord)
      If i > 0 Then
        l = Left(ListItems(lIndex).Text, Len(ListItems(lIndex).Text) - Len(ListItems(lIndex).SubItems(ListItems(lIndex).SubItems.Count).KeyWord))
        r = Right(ListItems(lIndex).Text, Len(ListItems(lIndex).SubItems(ListItems(lIndex).SubItems.Count).KeyWord))
      End If
      ListItems(lIndex).SubItems.Count = ListItems(lIndex).SubItems.Count + 1
      
      With ListItems(lIndex).SubItems(ListItems(lIndex).SubItems.Count)
        .Text = ""
        .Value = ""
        .ParaType = ""
        .Locked = True
        .KeyWord = ActiveString(PublicMsgs(156), ListItems(lIndex).SubItems.Count - 2)
        .ForeColor = MENU_COLOR_DEFAULT
        ListItems(lIndex).Text = l & .KeyWord & "," & r
      End With
      
      ListItems(lIndex).SubItems.Swap ListItems(lIndex).SubItems.Count - 1, ListItems(lIndex).SubItems.Count
      ListItems(lIndex).Refresh
      CustomActive = True
    End If
  Else
    'Label1(Index).ToolTipText = ListItems(lIndex).Text
    q = ListItems(lIndex).Locked
    Text = ListItems(lIndex).Text
    Tag = ListItems(lIndex).Value
  End If
  
  If Not q Then
    With txtEnter
      .Width = Label1(Index).Width + TextWidth(" ")
      .Height = Label1(Index).Height
      .Top = Label1(Index).Top + Label1(Index).Height / 2 - .Height / 2
      .Left = Label1(Index).Left + Label1(Index).Width / 2 - .Width / 2
      
      .Tag = Tag
      .Text = Text
      .Visible = True
      .SetFocus
      .SelStart = Len(.Text)
      .FontBold = Label1(Index).FontBold
      .ForeColor = Label1(Index).ForeColor
    End With
  End If
  
  If sIndex = 0 Then
    'ShowTip True
  End If

End If

RaiseEvent ItemClick(lIndex, sIndex)

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Idx(1) As Integer
Idx(0) = GetTagValue(Label1(Index).Tag, 0)
Idx(1) = GetTagValue(Label1(Index).Tag, 1)

If Idx(0) > 0 Then
  If Idx(1) > 0 Then
    Label1(Index).ToolTipText = ListItems(Idx(0)).SubItems(Idx(1)).KeyWord
  Else
    Label1(Index).ToolTipText = ListItems(Idx(0)).Text
  End If
End If
End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Integer, i As Integer, V As Byte, MosP As POINTAPI
'Call Label1_Click(Index)

If Button = vbRightButton And txtEnter.Visible = False Then
  Call AssignText
  For i = 1 To Label1.UBound
    If i = Index Then
      Label1(i).BorderStyle = 1
    Else
      Label1(i).BorderStyle = 0
    End If
  Next i
  
  lIndex = GetTagValue(Label1(Index).Tag, 0)
  sIndex = GetTagValue(Label1(Index).Tag, 1)
  
  If lIndex > 0 Then
    If sIndex > 0 Then
      If ListItems(lIndex).SubItems(sIndex).Negation Then
        mNeg(0).Checked = True
        For i = 1 To 2
          V = Val(ListItems(lIndex).SubItems(sIndex).Value)

          mNeg(i).Checked = ((V And i) = i)
          If mNeg(i).Checked Then
            mNeg(0).Checked = False
          End If
        Next i
      
        PopupMenu mNegs
      Else
        If ListItems(lIndex).SubItems(sIndex).ParaType <> "NA" Then
          GetCursorPos MosP
          MSMenu.Value = ListItems(lIndex).SubItems(sIndex).Value
          MSMenu.ParaType = ListItems(lIndex).SubItems(sIndex).ParaType
          MSMenu.TemType = ListItems(lIndex).SubItems(sIndex).TemType
          MSMenu.ShowMenu MosP.X, MosP.Y
        End If
      End If
    Else
      GetCursorPos MosP
      OpCombo.Value = ListItems(lIndex).Value
      OpCombo.Text = ListItems(lIndex).Text
      OpCombo.ShowMenu MosP.X, MosP.Y
    End If
  End If
End If

End Sub
Private Function AssignText() As Boolean
Dim q As Boolean, oItem As ListItem, temStr As String, TemVal As Integer

If lIndex > 0 Then
  If sIndex > 0 Then '参数
    q = ListItems(lIndex).SubItems(sIndex).Locked
    If Not q Then
       ListItems(lIndex).SubItems(sIndex).Text = txtEnter.Text
       ListItems(lIndex).SubItems(sIndex).Value = txtEnter.Tag
    End If
  Else '主项
    q = ListItems(lIndex).Locked
    If Not q Then
       'ListItems(lIndex).Text = txtEnter.Text
       ListItems(lIndex).AssignValue oItem.Tag
    End If
  End If
  
  txtEnter.Visible = False
  
  If Not q Then
    ListItems(lIndex).Refresh False
    MeasureFrame
  End If
End If
End Function

Private Sub mFac_Click(Index As Integer)
txtEnter.Text = mFac(Index).Caption
End Sub

Private Sub mNeg_Click(Index As Integer)
Dim V As Byte, i As Integer

If Index = 0 Then
  If Not mNeg(0).Checked Then
    mNeg(0).Checked = True
    V = 0
    GoTo Last
  End If
End If

V = Val(ListItems(lIndex).SubItems(sIndex).Value)
If mNeg(Index).Checked Then
  V = V - Index
ElseIf Not mNeg(Index).Checked Then
  V = V + Index
End If

mNeg(Index).Checked = Not mNeg(Index).Checked

If V = 0 Then
  mNeg(0).Checked = True
  mNeg(1).Checked = False
  mNeg(2).Checked = False
ElseIf V > 0 Then
  mNeg(0).Checked = False
  ListItems(lIndex).SubItems(sIndex).Text = ""
End If

Last:
ListItems(lIndex).SubItems(sIndex).Text = GetNegWord(CInt(V))
ListItems(lIndex).SubItems(sIndex).Value = CStr(V)
ListItems(lIndex).Refresh True
Call ListItems(lIndex).ListView.Event_Change
End Sub

Private Sub MSMenu_ItemSelect(Value As String, PType As String)
If lIndex > 0 And sIndex > 0 Then
  ListItems(lIndex).SubItems(sIndex).AssignValue Value, PType
  ListItems(lIndex).Refresh True
End If

End Sub

Private Sub OpCombo_ItemSelect(strValue As String, strText As String)
Dim q As Boolean

q = ListItems(lIndex).AssignValue(strValue)
If q Then
  'ListItems(lIndex).Text = strText
  ListItems(lIndex).Refresh True
End If
End Sub

Private Sub PicBox_Click()
Dim i As Integer

For i = 1 To Label1.UBound
  Label1(i).BorderStyle = 0
Next i

Call AssignText

lIndex = 0
sIndex = 0

RaiseEvent LostSelection
End Sub

Private Sub PicFrame_Click()
Call PicBox_Click
End Sub


Private Sub txtEnter_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Call PicBox_Click
End If
End Sub


Private Sub UserControl_Initialize()
p_Spacing = 1

PicBox.BorderStyle = 0
PicBox.Move 0, 0, 0, 0

ReDim ListItem(0)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
With MSBoard
  If lIndex > 0 Then
    If Shift = 2 Then
      If KeyCode = MENU_KEY_COPY Then
        If sIndex > 0 Then
          If ListItems(lIndex).SubItems(sIndex).ParaType <> "NEG" Then
            If ListItems(lIndex).SubItems(sIndex).ParaType <> "MORE" Then
              .ContentType = BOARD_PARAM
              .Value = ListItems(lIndex).SubItems(sIndex).Value
              .TemType = ListItems(lIndex).SubItems(sIndex).TemType
            End If
          End If
        Else
          .ContentType = BOARD_OP
          .Value = ListItems(lIndex).Value
          .TemType = ""
        End If
      ElseIf KeyCode = MENU_KEY_PASTE Then
        If .ContentType = BOARD_PARAM Then
          If ListItems(lIndex).SubItems(sIndex).ParaType <> "NEG" Then
            If ListItems(lIndex).SubItems(sIndex).ParaType <> "MORE" Then
                ListItems(lIndex).SubItems(sIndex).AssignValue .Value, .TemType
                ListItems(lIndex).Refresh
            End If
          End If
        ElseIf .ContentType = BOARD_OP Then
            ListItems(lIndex).AssignValue .Value
            ListItems(lIndex).Refresh
        End If
      End If
    End If
  End If
End With
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
PicFrame.Move 0, 0, UserControl.ScaleWidth - VScroll1.Width, UserControl.ScaleHeight - HScroll1.Height
VScroll1.Move UserControl.ScaleWidth - VScroll1.Width, 0, VScroll1.Width, UserControl.ScaleHeight - HScroll1.Height
HScroll1.Move 0, UserControl.ScaleHeight - HScroll1.Height, UserControl.ScaleWidth - VScroll1.Width

NeedScroll

End Sub

Private Sub NeedScroll()
If PicBox.Height > PicFrame.ScaleHeight Then
  VScroll1.Enabled = True
  VScroll1.Max = PicBox.Height - PicFrame.ScaleHeight
  VScroll1.SmallChange = PicFrame.ScaleHeight / 2
  VScroll1.LargeChange = PicFrame.ScaleHeight
Else
  VScroll1.Enabled = False
  VScroll1.Max = 0
End If

If PicBox.Width > PicFrame.ScaleWidth Then
  HScroll1.Enabled = True
  HScroll1.Max = PicBox.Width - PicFrame.ScaleWidth
  HScroll1.SmallChange = PicFrame.ScaleWidth / 2
  HScroll1.LargeChange = PicFrame.ScaleWidth
Else
  HScroll1.Enabled = False
  HScroll1.Max = 0
End If
End Sub

Private Sub UserControl_Terminate()
'UnHook txtEnter.hWnd
End Sub

Private Sub VScroll1_Change()
PicBox.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Call VScroll1_Change
End Sub

Public Function FreeLabel() As Label
Attribute FreeLabel.VB_MemberFlags = "40"
Dim i As Integer

For i = 1 To Label1.UBound
  With Label1(i)
    If Not .Visible Then
      Exit For
    End If
  End With
Next i

If i > Label1.UBound Then
  Load Label1(i)
  With Label1(i)
    .AutoSize = True
    
  End With
End If


Set FreeLabel = Label1(i)
End Function

Public Sub MeasureFrame()
Dim i As Integer, MaxWidth As Long, MainLabel As Label, tHeight As Integer

For i = 1 To ListItems.Count
  Set MainLabel = ListItems(i).Label
  tHeight = tHeight + MainLabel.Height * p_Spacing
  If MainLabel.Width + MainLabel.Left > MaxWidth Then MaxWidth = MainLabel.Width + MainLabel.Left
Next i

PicBox.Width = MaxWidth + 4
PicBox.Height = tHeight + 4

NeedScroll
End Sub


Public Sub SwapLabels(Lab1 As Integer, Lab2 As Integer)
Attribute SwapLabels.VB_MemberFlags = "40"
Dim TemString As String, TemSingle As Single, TemLong As Long

TemString = Label1(Lab1).Tag
Label1(Lab1).Tag = Label1(Lab2).Tag
Label1(Lab2).Tag = TemString

TemSingle = Label1(Lab1).Top
Label1(Lab1).Top = Label1(Lab2).Top
Label1(Lab2).Top = TemSingle

TemSingle = Label1(Lab1).Left
Label1(Lab1).Left = Label1(Lab2).Left
Label1(Lab2).Left = TemSingle

End Sub

Public Sub SwapLabelPos(Lab1 As Integer, Lab2 As Integer)
Dim TemSingle As Single, TemLong As Long

TemSingle = Label1(Lab1).Top
Label1(Lab1).Top = Label1(Lab2).Top
Label1(Lab2).Top = TemSingle

TemSingle = Label1(Lab1).Left
Label1(Lab1).Left = Label1(Lab2).Left
Label1(Lab2).Left = TemSingle

End Sub


Public Property Get ListIndex() As Integer
ListIndex = lIndex
End Property

Public Property Let ListIndex(ByVal vNewValue As Integer)
Dim i As Integer, MainLabel As Label

For i = 1 To ListItems.Count
  Set MainLabel = ListItems(i).Label
  If i = vNewValue And sIndex = 0 Then
    MainLabel.BorderStyle = 1
  Else
    MainLabel.BorderStyle = 0
  End If
Next i

lIndex = vNewValue

PropertyChanged (ListIndex)
End Property

Public Property Get SubIndex() As Integer
SubIndex = sIndex
End Property

Public Property Let SubIndex(ByVal vNewValue As Integer)
Dim i As Integer, ParamLabel As Label

If lIndex > 0 Then
  For i = 1 To ListItems(lIndex).SubItems.Count
    Set ParamLabel = ListItems(lIndex).SubItems(i).Label
    If i = vNewValue Then
      ParamLabel.BorderStyle = 1
    Else
      ParamLabel.BorderStyle = 0
    End If
  Next i
End If

sIndex = vNewValue

PropertyChanged (SubIndex)
End Property

Public Property Get AutoRefresh() As Boolean
AutoRefresh = AutoRef
End Property

Public Property Let AutoRefresh(ByVal vNewValue As Boolean)
AutoRef = vNewValue

PropertyChanged AutoRefresh
End Property


Public Sub Refresh(Optional LabelOnly As Boolean = False)
Dim i As Integer, TemInt As Integer, TemInt2 As Integer, Selected As Long, LLeft As Single

If Not LabelOnly Then
  For i = 1 To ListItems.Count
    ListItems(i).Refresh False
  Next i
Else
  For i = 1 To Label1.UBound
    If Label1(i).Visible Then
      TemInt = GetTagValue(Label1(i).Tag, 0)
      TemInt2 = GetTagValue(Label1(i).Tag, 1)
      Selected = Label1(i).BorderStyle
      If Selected = 1 Then Label1(i).BorderStyle = 0
      Label1(i).Top = (TemInt - 1) * p_Spacing * Label1(i).Height
      If TemInt2 = 0 Then
        If ListItems(TemInt).Indention > 0 Then
          Label1(i).Left = TextWidth(Space(ListItems(TemInt).Indention))
        Else
          Label1(i).Left = 0
        End If
      Else
        If ListItems(TemInt).SubItems(TemInt2).Start > 0 Then
          Label1(i).Left = ListItems(TemInt).Label.Left + TextWidth(Space(ListItems(TemInt).SubItems(TemInt2).Start - 1))
        End If
      End If
      If Selected = 1 Then Label1(i).BorderStyle = 1
    End If
  Next i
End If
MeasureFrame

End Sub

'*************************************************************************
'**函 数 名：LenW
'**输    入：(String)iStr
'**输    出：(Long)
'**功能描述：获得字符串显示长度
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-10 21:47:28
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function LenW(iStr As String) As Long
Dim i As Long, sW As Single, k As String

sW = TextWidth(" ")
For i = 1 To Len(iStr)
   k = Mid(iStr, i, 1)
   If TextWidth(k) > sW Then
     LenW = LenW + 2
   Else
     LenW = LenW + 1
   End If
Next i
End Function

'*************************************************************************
'**函 数 名：TextWidth
'**输    入：(String)iStr
'**输    出：(Long)
'**功能描述：获得字符串显示尺寸
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-10 21:48:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function TextWidth(iStr As String) As Long

TextWidth = UserControl.TextWidth(iStr)

End Function

'*************************************************************************
'**函 数 名：FillSpace
'**输    入：(String)iStr
'**输    出：(Long)
'**功能描述：获得字符串显示尺寸
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-06-10 21:48:33
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Function FillSpace(iStr As String, SpcFont As StdFont, iStrFont As StdFont) As String
Attribute FillSpace.VB_MemberFlags = "40"
Dim Wd As Single, i As Long, w As Single, strTem As String

If iStr = "" Then Exit Function
UserControl.Font = iStrFont
Wd = UserControl.TextWidth(iStr)

UserControl.Font = SpcFont
Do While w < Wd
  strTem = strTem & " "
  w = UserControl.TextWidth(strTem)
Loop

FillSpace = strTem
End Function

'*************************************************************************
'**函 数 名：FillMenu
'**输    入：-
'**输    出：()
'**功能描述：填充自动完成菜单
'**全局变量：
'**调用模块：
'**作    者：SSgt_Edward
'**日    期：2011-07-09 17:44:26
'**修 改 人：
'**日    期：
'**版    本：V1.1321
'*************************************************************************
Public Sub FillMenu()

mNeg(0).Caption = PublicMsgs(157)
mNeg(1).Caption = PublicMsgs(153)
mNeg(2).Caption = PublicMsgs(154)
End Sub

Public Property Get Label() As Object
Attribute Label.VB_MemberFlags = "40"
Set Label = Label1
End Property

Public Property Let Label(ByVal vNewValue As Object)
Set Label = vNewValue

PropertyChanged (Label)
End Property


Public Property Get Spacing() As Single
Spacing = p_Spacing
End Property

Public Property Let Spacing(ByVal vNewValue As Single)
p_Spacing = vNewValue

PropertyChanged (Spacing)
End Property

Public Sub Event_Change()
RaiseEvent Change
End Sub
