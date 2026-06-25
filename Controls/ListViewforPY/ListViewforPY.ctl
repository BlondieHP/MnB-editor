VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ListViewforPY 
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
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
      Begin VB.TextBox txtEnter 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1200
         TabIndex        =   6
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComctlLib.ListView lstTips 
         Height          =   1695
         Left            =   960
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2990
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
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
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
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
End
Attribute VB_Name = "ListViewforPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ListItem() As New Type_ListItemforPY
Private lIndex As Integer
Private sIndex As Integer
Private AutoRef As Boolean

Public Event ItemClick(ItemIndex As Integer, SubIndex As Integer)
Public Event FillBlank(ItemIndex As Integer, SubIndex As Integer, TipIndex As Integer)
Public Event ItemChoose(ItemIndex As Integer, SubIndex As Integer, TipIndex As Integer)
Public Event LostSelection()

Private Sub HScroll1_Change()
PicBox.Left = -HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
Call HScroll1_Change
End Sub

Private Sub InitLstOp()
Dim n As Integer
n = 1
With lstTips
      .Sorted = True
      .ListItems.Clear
      .ColumnHeaders.Clear
      .SortOrder = lvwAscending
      .FullRowSelect = True
      .AllowColumnReorder = False
      .LabelEdit = lvwManual
      .Checkboxes = False
      .GridLines = False
      .MultiSelect = False
      .HideSelection = False
      .Visible = False

      .ColumnHeaders.Add , , "提示"
End With

End Sub

Private Sub Label1_Click(Index As Integer)
Dim i As Integer, q As Boolean, Text As String
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
    q = ListItem(lIndex).Params(sIndex).Locked
    Text = ListItem(lIndex).Params(sIndex).Text
  Else
    q = ListItem(lIndex).Locked
    Text = ListItem(lIndex).Text
  End If
  
  If Not q Then
    txtEnter.Move Label1(Index).Left, Label1(Index).Top, Label1(Index).Width + TextWidth(" "), Label1(Index).Height
    txtEnter.Text = Text
    txtEnter.Visible = True
    txtEnter.SetFocus
    txtEnter.SelStart = Len(txtEnter.Text)
  End If
  
End If

RaiseEvent ItemClick(lIndex, sIndex)

End Sub


Private Sub lstTips_DblClick()
Dim q As Boolean

If Not lstTips.SelectedItem Is Nothing Then

  If lIndex > 0 Then
    If sIndex = 0 Then
      q = ListItem(lIndex).Locked
      If q Then ListItem(lIndex).Text = lstTips.SelectedItem.Text
    Else
      q = ListItem(lIndex).Params(sIndex).Locked
      If q Then ListItem(lIndex).Params(sIndex).Text = lstTips.SelectedItem.Text
    End If
    
    If q Then
      lstTips.Visible = False
      PurseItemText lIndex
    Else
      txtEnter.Text = lstTips.SelectedItem.Text
      txtEnter.SetFocus
      txtEnter.SelStart = Len(txtEnter.Text)
    End If
    
    RaiseEvent ItemChoose(lIndex, sIndex, lstTips.SelectedItem.Index)
  End If
  
End If

End Sub

Private Sub AssignText()
Dim q As Boolean

If lIndex > 0 Then
  If sIndex > 0 Then
    q = ListItem(lIndex).Params(sIndex).Locked
    If Not q Then ListItem(lIndex).Params(sIndex).Text = txtEnter.Text
  Else
    q = ListItem(lIndex).Locked
    If Not q Then ListItem(lIndex).Text = txtEnter.Text
  End If
  
  txtEnter.Visible = False
  
  If Not q Then
    PurseItemText lIndex
  End If
End If
End Sub

Private Sub mFac_Click(Index As Integer)
txtEnter.Text = mFac(Index).Caption
End Sub

Private Sub PicBox_Click()
Dim i As Integer

For i = 1 To Label1.UBound
  Label1(i).BorderStyle = 0
Next i

Call AssignText

lIndex = 0
sIndex = 0

lstTips.Visible = False
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

Private Sub txtEnter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
  UserControl.PopupMenu mnuMain
End If
End Sub

Private Sub UserControl_Initialize()

PicBox.BorderStyle = 0
PicBox.Move 0, 0, 0, 0

ReDim ListItem(0)

InitLstOp
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

Private Sub VScroll1_Change()
PicBox.Top = -VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
Call VScroll1_Change
End Sub

Private Function FreeLabel()
Dim i As Integer

For i = 1 To Label1.UBound
  With Label1(i)
    If Not .Visible Then
      Exit For
    End If
  End With
Next i

FreeLabel = i
If FreeLabel > Label1.UBound Then
  Load Label1(FreeLabel)
  With Label1(FreeLabel)
    .AutoSize = True
    
  End With
End If
End Function

Public Function AddItem_Trg()

End Function

Public Function AddItem(Text As String, Optional ListIndex As Integer) As Integer
Dim New_Index As Integer, i As Integer, j As Integer, MainLabel As Integer, ParamLabel As Integer

If ListIndex > 0 Then
   New_Index = ListIndex
Else
   New_Index = UBound(ListItem) + 1
End If

ReDim Preserve ListItem(UBound(ListItem) + 1)
ReDim ListItem(UBound(ListItem)).Params(0)

For i = UBound(ListItem) - 1 To New_Index Step -1
  ListItem(i + 1) = ListItem(i)
Next i

With ListItem(New_Index)
  .OpLabel = FreeLabel()
  .ParamCount = 0
  ReDim .Params(0)
End With

With Label1(ListItem(New_Index).OpLabel)
  ListItem(New_Index).Text = Text
  .Caption = Text
  .Visible = True
  .Tag = CStr(New_Index) & ",0"
  .Left = 0
  .Top = (New_Index - 1) * .Height
End With

For i = New_Index + 1 To UBound(ListItem)
  MainLabel = ListItem(i).OpLabel
  Label1(MainLabel).Left = 0
  Label1(MainLabel).Top = (i - 1) * Label1(MainLabel).Height
  Label1(MainLabel).Tag = TagPlus(Label1(MainLabel).Tag, 0, 1)
  For j = 1 To ListItem(i).ParamCount
    ParamLabel = ListItem(i).Params(j).Label
    Label1(ParamLabel).Top = Label1(MainLabel).Top
  Next j
Next i

AddItem = New_Index

If AutoRef Then
  PurseItemText New_Index
End If
MeasureFrame

NeedScroll

End Function

Public Sub RemoveItem(ListIndex As Integer)
Dim i As Integer, j As Integer, MainLabel As Integer, ParamLabel As Integer, Label(1) As Integer

MainLabel = ListItem(ListIndex).OpLabel
Label1(MainLabel).Visible = False
For j = 1 To ListItem(ListIndex).ParamCount
   ParamLabel = ListItem(ListIndex).Params(j).Label
   Label1(ParamLabel).Visible = False
Next j
  
For i = UBound(ListItem) - 1 To ListIndex Step -1
  Label(0) = ListItem(i).OpLabel
  Label(1) = ListItem(i + 1).OpLabel
  Label1(Label(1)).Top = Label1(Label(0)).Top
  Label1(Label(1)).Tag = TagPlus(Label1(Label(1)).Tag, 0, -1)
  
  For j = 1 To ListItem(i).ParamCount
    ParamLabel = ListItem(i).Params(j).Label
    Label1(ParamLabel).Top = Label1(MainLabel).Top
  Next j
Next i

For i = ListIndex To UBound(ListItem) - 1
  ListItem(i) = ListItem(i + 1)
Next i

ReDim Preserve ListItem(UBound(ListItem) - 1)

MeasureFrame

End Sub

Private Sub MeasureFrame()
Dim i As Integer, MaxWidth As Long, MainLabel As Integer, tHeight As Integer

For i = 1 To UBound(ListItem)
  MainLabel = ListItem(i).OpLabel
  tHeight = tHeight + Label1(MainLabel).Height
  If Label1(MainLabel).Width + Label1(MainLabel).Left > MaxWidth Then MaxWidth = Label1(MainLabel).Width + Label1(MainLabel).Left
Next i

PicBox.Width = MaxWidth + 4
PicBox.Height = tHeight + 4

End Sub

Public Sub SetText(Index As Integer, SubIndex As Integer, Text As String)

If SubIndex > 0 Then
  ListItem(Index).Params(SubIndex).Text = Text
Else
  ListItem(Index).Text = Text
End If

If AutoRef Then PurseItemText Index
End Sub

Public Function GetText(Index As Integer, SubIndex As Integer) As String

If SubIndex > 0 Then
  GetText = ListItem(Index).Params(SubIndex).Text
Else
  GetText = ListItem(Index).Text
End If

End Function


Public Sub SetColor(Index As Integer, SubIndex As Integer, lColor As Long)
Dim MainLabel As Integer, ParamLabel As Integer

MainLabel = ListItem(Index).OpLabel

If SubIndex > 0 Then
  ParamLabel = ListItem(Index).Params(SubIndex).Label
  Label1(ParamLabel).ForeColor = lColor
Else
  Label1(MainLabel).ForeColor = lColor
End If

End Sub

Public Function GetColor(Index As Integer, SubIndex As Integer) As Long
Dim MainLabel As Integer, ParamLabel As Integer

MainLabel = ListItem(Index).OpLabel

If SubIndex > 0 Then
  ParamLabel = ListItem(Index).Params(SubIndex).Label
  GetColor = Label1(ParamLabel).ForeColor
Else
  GetColor = Label1(MainLabel).ForeColor
End If

End Function

Private Sub SwapLabels(Lab1 As Integer, Lab2 As Integer)
Dim TemString As String, TemSingle As Single, TemLong As Long
'Label1(0) = Label1(Lab1)
'Label1(Lab1) = Label1(Lab2)
'Label1(Lab2) = Label1(0)

TemString = Label1(Lab1).Tag
Label1(Lab1).Tag = Label1(Lab2).Tag
Label1(Lab2).Tag = TemString

TemLong = Label1(Lab1).ForeColor
Label1(Lab1).ForeColor = Label1(Lab2).ForeColor
Label1(Lab2).ForeColor = TemLong

TemSingle = Label1(Lab1).Top
Label1(Lab1).Top = Label1(Lab2).Top
Label1(Lab2).Top = TemSingle

TemSingle = Label1(Lab1).Left
Label1(Lab1).Left = Label1(Lab2).Left
Label1(Lab2).Left = TemSingle

End Sub

Public Sub SwapListItems(Item1 As Integer, Item2 As Integer)
Dim tItem As Type_ListItemforPY, j As Integer, TemString As String

SwapLabels ListItem(Item1).OpLabel, ListItem(Item2).OpLabel


tItem = ListItem(Item1)
ListItem(Item1) = ListItem(Item2)
ListItem(Item2) = tItem

End Sub

Public Property Get ListIndex() As Integer
ListIndex = lIndex
End Property

Public Property Let ListIndex(ByVal vNewValue As Integer)
Dim i As Integer, MainLabel As Integer

For i = 1 To UBound(ListItem)
  MainLabel = ListItem(i).OpLabel
  If i = vNewValue And sIndex = 0 Then
    Label1(MainLabel).BorderStyle = 1
  Else
    Label1(MainLabel).BorderStyle = 0
  End If
Next i

lIndex = vNewValue

PropertyChanged (ListIndex)
End Property

Public Property Get SubIndex() As Integer
SubIndex = sIndex
End Property

Public Property Let SubIndex(ByVal vNewValue As Integer)
Dim i As Integer, MainLabel As Integer, ParamLabel As Integer

If lIndex > 0 Then
  For i = 1 To ListItem(lIndex).ParamCount
    MainLabel = ListItem(lIndex).OpLabel
    ParamLabel = ListItem(lIndex).Params(i).Label
    If i = vNewValue Then
      Label1(ParamLabel).BorderStyle = 1
    Else
      Label1(ParamLabel).BorderStyle = 0
    End If
  Next i
End If

sIndex = vNewValue

PropertyChanged (SubIndex)
End Property

Private Function TagPlus(Tag As String, Index As Integer, Value As Integer) As String
Dim strTem() As String, lngTem As Integer, i As Integer

If Tag = "" Then Exit Function
strTem = Split(Tag, ",")

lngTem = Val(strTem(Index)) + Value
strTem(Index) = CStr(lngTem)

For i = 0 To UBound(strTem)
  TagPlus = TagPlus & strTem(i) & ","
Next i

TagPlus = Left(TagPlus, Len(TagPlus) - 1)
End Function

Private Function GetTagValue(Tag As String, Index As Integer) As Integer
Dim strTem() As String, lngTem As Integer, i As Integer

If Tag = "" Then Exit Function
strTem = Split(Tag, ",")

GetTagValue = Val(strTem(Index))

End Function

Public Sub SetParamCount(Index As Integer, Count As Integer)

ListItem(Index).ParamCount = Count
ReDim Preserve ListItem(Index).Params(Count)

If AutoRef Then PurseItemText Index
End Sub

Public Function GetParamCount(Index As Integer) As Integer
GetParamCount = ListItem(Index).ParamCount

End Function

Public Sub SetIndention(Index As Integer, Value As Integer)

ListItem(Index).Indention = Value

If AutoRef Then PurseItemText Index
End Sub

Public Function GetIndention(Index As Integer) As Integer
GetParamCount = ListItem(Index).Indention

End Function

Public Sub PurseItemText(Index As Integer)
Dim i As Integer, KeyWord As String, n As Integer

With ListItem(Index)
  For i = 1 To ListItem(Index).ParamCount
    KeyWord = "{arg_" & i & "}"
    n = InStr(1, .Text, KeyWord)
    .Params(i).Start = LenW(Left(.Text, n))
    If n > 0 Then
      If .Params(i).Label = 0 Then
        .Params(i).Label = FreeLabel()
        Label1(.Params(i).Label).Tag = Index & "," & i
        Label1(.Params(i).Label).Visible = True
      End If
      Label1(.Params(i).Label).Caption = .Params(i).Text
    End If
  Next i

End With
ShowLabel Index

End Sub

Private Sub ShowLabel(Index As Integer)
Dim i As Integer, j As Integer, MainLabel As Integer, KeyWord As String, ParamLabel As Integer, LastCaption As String

MainLabel = ListItem(Index).OpLabel
Label1(MainLabel).Caption = ListItem(Index).Text
Label1(MainLabel).Left = TextWidth(Space(ListItem(Index).Indention))

For i = 1 To ListItem(Index).ParamCount
  ParamLabel = ListItem(Index).Params(i).Label
  KeyWord = "{arg_" & i & "}"
  
  Label1(MainLabel).Caption = Replace(Label1(MainLabel).Caption, KeyWord, Space(LenW(ListItem(Index).Params(i).Text)), , 1)

  For j = i + 1 To ListItem(Index).ParamCount
    ListItem(Index).Params(j).Start = ListItem(Index).Params(j).Start + LenW(ListItem(Index).Params(i).Text) - LenW(KeyWord)
  Next j
Next i

For i = 1 To ListItem(Index).ParamCount
  ParamLabel = ListItem(Index).Params(i).Label
  If ParamLabel > 0 Then
    Label1(ParamLabel).Top = Label1(MainLabel).Top
    
    If ListItem(Index).Params(i).Start > 0 Then
      Label1(ParamLabel).Left = Label1(MainLabel).Left + TextWidth(Space(ListItem(Index).Params(i).Start - 1))
      Label1(ParamLabel).Visible = True
      Label1(ParamLabel).ZOrder
      'Label1(ParamLabel).ForeColor = vbRed
    Else
      Label1(ParamLabel).Visible = False
    End If
  End If
Next i

MeasureFrame
End Sub

Public Property Get AutoRefresh() As Boolean
AutoRefresh = AutoRef
End Property

Public Property Let AutoRefresh(ByVal vNewValue As Boolean)
AutoRef = vNewValue

PropertyChanged AutoRefresh
End Property

Public Function AddComboItem(Text As String, Optional Index As Integer) As Integer
Dim n As ListItem
If Index > 0 Then
  Set n = lstTips.ListItems.Add(Index, , Text)
Else
  Set n = lstTips.ListItems.Add(, , Text)
End If

AddComboItem = n.Index

If AutoRef Then MeasureListTip
End Function

Public Sub RemoveComboItem(Index As Integer)
lstTips.ListItems.Remove Index

If AutoRef Then MeasureListTip
End Sub

Public Sub ClearComboItem()
lstTips.ListItems.Clear
End Sub

Public Sub ClearItems()
Dim i As Long

For i = 1 To Label1.Count
  UnLoad Label1(i)
Next i

ReDim ListItem(0)

End Sub

Public Property Get Sorted() As Boolean
Sorted = lstTips.Sorted
End Property

Public Property Let Sorted(ByVal vNewValue As Boolean)
lstTips.Sorted = vNewValue
PropertyChanged Sorted
End Property

Public Sub ShowTip(Optional Visible As Boolean = True)
Dim MainLabel As Integer, ParamLabel As Integer


With lstTips
  If lIndex > 0 Then
    MainLabel = ListItem(lIndex).OpLabel
    ParamLabel = ListItem(lIndex).Params(sIndex).Label
  
    If MainLabel > 0 Then
      
      If ParamLabel > 0 Then
        If ListItem(lIndex).Params(sIndex).Locked Then
          .Top = Label1(MainLabel).Top + Label1(MainLabel).Height
        Else
          .Top = txtEnter.Top + txtEnter.Height
        End If
        
        .Left = Label1(ParamLabel).Left
      Else
        .Left = Label1(MainLabel).Left
        
        If ListItem(lIndex).Locked Then
          .Top = Label1(MainLabel).Top + Label1(MainLabel).Height
        Else
          .Top = txtEnter.Top + txtEnter.Height
        End If
        
      End If
    
      If lstTips.Left + lstTips.Width > PicFrame.ScaleWidth And lstTips.Width < PicFrame.ScaleWidth Then
        lstTips.Left = PicFrame.ScaleWidth - lstTips.Width
      End If
    
    End If
  
  End If
  .Visible = Visible
End With


End Sub

Public Sub MeasureListTip()
Dim i As Integer, MaxWidth As Long, MainLabel As Integer, tHeight As Integer

With lstTips
  For i = 1 To .ListItems.Count
    If i = 0 Then
       MaxWidth = TextWidth(.ListItems(i).Text)
    Else
       If MaxWidth < TextWidth(.ListItems(i).Text) Then
         MaxWidth = TextWidth(.ListItems(i).Text)
       End If
    End If
  Next i
  
  If .ListItems.Count > 5 Then
    .Height = 5 * .ListItems(1).Height + 4
  ElseIf .ListItems.Count > 0 Then
    .Height = .ListItems.Count * .ListItems(1).Height + 4
  End If
  
  .Width = MaxWidth + TextWidth(Space(2))
  .ColumnHeaders(1).Width = .Width * 0.9
End With

End Sub

Public Property Get ComboVisible() As Boolean
ComboVisible = lstTips.Visible
End Property

Public Property Let ComboVisible(ByVal vNewValue As Boolean)
lstTips.Visible = vNewValue

PropertyChanged ComboVisible
End Property

Public Sub PurseItemsText()
Dim i As Integer

For i = 1 To UBound(ListItem)
  PurseItemText i
Next i
End Sub


Public Sub SetLocked(Index As Integer, SubIndex As Integer, Locked As Boolean)

If SubIndex > 0 Then
  ListItem(Index).Params(SubIndex).Locked = Locked
Else
  ListItem(Index).Locked = Locked
End If

End Sub

Public Function GetLocked(Index As Integer, SubIndex As Integer) As Boolean

If SubIndex > 0 Then
  GetLocked = ListItem(Index).Params(SubIndex).Locked
Else
  GetLocked = ListItem(Index).Locked
End If

End Function


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
Private Function LenW(iStr As String) As Long
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
Dim i As Long, j As Long, strTem() As String, q As Boolean, mArray As Menu

For i = 0 To 26
   mEditor(i).Caption = PublicTags(i)
   
   If i = 0 Or i = Tag_Register Or i = Tag_Variable Or i = Tag_Local_Variable Or i = Tags_End Then
     mEditor(i).Visible = False
   Else
     q = EnumEntities(i, PublicTags(i) & ":([index])[csvname]", strTem)
     
     If q Then
       Select Case i
         Case Tag_String
           mStr(0).Caption = PublicTips(0)  'A-E
         Case Tag_Item
           For j = N_Item To mItm.UBound
             UnLoad mItm(j)
           Next j
           
           For j = mItm.Count To N_Item - 1
             Load mItm(j)
             mItm(j).Visible = True
           Next j
           
           For j = 0 To N_Item - 1
             mItm(j).Caption = strTem(j)
           Next j
         Case Tag_Troop
           For j = N_Troop To mTrp.UBound
             UnLoad mTrp(j)
           Next j
           
           For j = mTrp.Count To N_Troop - 1
             Load mTrp(j)
             mTrp(j).Visible = True
           Next j
           
           For j = 0 To N_Troop - 1
             mTrp(j).Caption = strTem(j)
           Next j
         Case Tag_Faction
           For j = N_Faction To mFac.UBound
             UnLoad mFac(j)
           Next j
           
           For j = mFac.Count To N_Faction - 1
             Load mFac(j)
             mFac(j).Visible = True
           Next j
           
           For j = 0 To N_Faction - 1
             mFac(j).Caption = strTem(j)
           Next j
         Case Tag_Quest
           mQst(0).Caption = PublicTips(0)   'A-E
         Case Tag_Party_Tpl
           For j = N_PT To mPT.UBound
             UnLoad mPT(j)
           Next j
           
           For j = mPT.Count To N_PT - 1
             Load mPT(j)
             mPT(j).Visible = True
           Next j
           
           For j = 0 To N_PT - 1
             mPT(j).Caption = strTem(j)
           Next j
         Case Tag_Party
           For j = N_Party To mParty.UBound
             UnLoad mParty(j)
           Next j
           
           For j = mParty.Count To N_Party - 1
             Load mParty(j)
             mParty(j).Visible = True
           Next j
           
           For j = 0 To N_Party - 1
             mParty(j).Caption = strTem(j)
           Next j
         Case Tag_Scene
           For j = N_Scene To mScene.UBound
             UnLoad mScene(j)
           Next j
           
           For j = mScene.Count To N_Scene - 1
             Load mScene(j)
             mScene(j).Visible = True
           Next j
           
           For j = 0 To N_Scene - 1
             mScene(j).Caption = strTem(j)
           Next j
         Case Tag_Mission_tpl
           mMT(0).Caption = PublicTips(0)   'A-E
         Case Tag_Menu
           mMnu(0).Caption = PublicTips(0)   'A-E
         Case Tag_Script
           mScript(0).Caption = PublicTips(0)   'A-E
         Case Tag_Particle_Sys
           For j = N_PSys To mPSys.UBound
             UnLoad mPSys(j)
           Next j
           
           For j = mPSys.Count To N_PSys - 1
             Load mPSys(j)
             mPSys(j).Visible = True
           Next j
           
           For j = 0 To N_PSys - 1
             mPSys(j).Caption = strTem(j)
           Next j
         Case Tag_Scene_Prop
           mSP(0).Caption = PublicTips(0)   'A-E
         Case Tag_Sound
           For j = N_Sound To mSnd.UBound
             UnLoad mSnd(j)
           Next j
           
           For j = mSnd.Count To N_Sound - 1
             Load mSnd(j)
             mSnd(j).Visible = True
           Next j
           
           For j = 0 To N_Sound - 1
             mSnd(j).Caption = strTem(j)
           Next j
         Case Tag_Local_Variable
           mLvar(0).Caption = PublicTips(0)   'A-E
         Case Tag_Map_Icon
           For j = N_MapIcon To mMI.UBound
             UnLoad mMI(j)
           Next j
           
           For j = mMI.Count To N_MapIcon - 1
             Load mMI(j)
             mMI(j).Visible = True
           Next j
           
           For j = 0 To N_MapIcon - 1
             mMI(j).Caption = strTem(j)
           Next j
         Case Tag_Skill
           For j = 0 To UBound(strTem)
             mSkl(j).Caption = strTem(j)
           Next j
         Case Tag_Mesh
           mMesh(0).Caption = PublicTips(0)   'A-E
         Case Tag_Presentation
           mPrsnt(0).Caption = PublicTips(0)   'A-E
         Case Tag_Quick_String
           mQStr(0).Caption = PublicTips(0)   'A-E
         Case Tag_Track
           mTrack(0).Caption = PublicTips(0)   'A-E
         Case Tag_Tableau
           For j = N_TabMat To mTabMat.UBound
             UnLoad mTabMat(j)
           Next j
           
           For j = mTabMat.Count To N_TabMat - 1
             Load mTabMat(j)
             mTabMat(j).Visible = True
           Next j
           
           For j = 0 To N_TabMat - 1
             mTabMat(j).Caption = strTem(j)
           Next j
         Case Tag_Animation
           mAni(0).Caption = PublicTips(0)   'A-E
       End Select
     
     End If
   End If
Next i

End Sub
