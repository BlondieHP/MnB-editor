VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl RichforPY 
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8685
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   579
   Begin MSComctlLib.ListView lstTips 
      Height          =   1695
      Left            =   480
      TabIndex        =   0
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
   Begin VB.PictureBox PicTip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Timer Delayer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   10398
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"RichforPY.ctx":0000
   End
End
Attribute VB_Name = "RichforPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim CustomActive As Boolean, ColorRemark As Boolean
Dim MainRich_SelStart As Long, ListTips_DblClick As Boolean, LastLoadTag As Long, LastLoadRes As Boolean
Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event TipShow(Visible)
Public Event ListShow(Visible)

Private Sub Delayer_Timer()
ShowParamTip
Delayer.Enabled = False
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

      .ColumnHeaders.Add , , "提示"
End With

End Sub

Private Sub FillOperations()
Dim i As Long

With lstTips
   .ListItems.Clear
   
   For i = 0 To UBound(Operation)
     .ListItems.Add , "Op_" & i, Operation(i).Op_name
   Next i
   
   .Width = TextWidth(Space(Max_Op_Len + 5))
   .ColumnHeaders(1).Width = 0.9 * .Width
   
   .ListItems(1).Selected = True
End With

LastLoadTag = 0

End Sub

Private Function FillItems(ByVal TagNo As Long) As Boolean
Dim i As Long, MaxLen As Long

With lstTips
   .ListItems.Clear
   
   Select Case TagNo   'A-E
     Case Tag_Troop  '兵种
       For i = 0 To N_Troop - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & Trps(i).strID & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(Trps(i).strID)
         Else
           If MaxLen < Len(Trps(i).strID) Then MaxLen = Len(Trps(i).strID)
         End If
       Next i
     
     Case Tag_Item   '物品
       For i = 0 To N_Item - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & itm(i).dbName & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(itm(i).dbName)
         Else
           If MaxLen < Len(itm(i).dbName) Then MaxLen = Len(itm(i).dbName)
         End If
       Next i
     
     Case Tag_Party  '部队
       For i = 0 To N_Party - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & Parties(i).strID & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(Parties(i).strID)
         Else
           If MaxLen < Len(Parties(i).strID) Then MaxLen = Len(Parties(i).strID)
         End If
       Next i
       
     Case Tag_Party_Tpl  '部队模板
       For i = 0 To N_PT - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & PTs(i).ptID & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(PTs(i).ptID)
         Else
           If MaxLen < Len(PTs(i).ptID) Then MaxLen = Len(PTs(i).ptID)
         End If
       Next i
       
     Case Tag_Faction  '阵营
       For i = 0 To N_Faction - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & Factions(i).strID & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(Factions(i).strID)
         Else
           If MaxLen < Len(Factions(i).strID) Then MaxLen = Len(Factions(i).strID)
         End If
       Next i
       
     Case Tag_Scene  '场景
       For i = 0 To N_Scene - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & Scenes(i).strID & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(Scenes(i).strID)
         Else
           If MaxLen < Len(Scenes(i).strID) Then MaxLen = Len(Scenes(i).strID)
         End If
       Next i
       
     Case Tag_Mesh  '网格
       For i = 0 To N_Mesh - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & Mesh(i).strID & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(Mesh(i).strID)
         Else
           If MaxLen < Len(Mesh(i).strID) Then MaxLen = Len(Mesh(i).strID)
         End If
       Next i
       
     Case Tag_Particle_Sys  '粒子系统
       For i = 0 To N_PSys - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & PSys(i).strID & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(PSys(i).strID)
         Else
           If MaxLen < Len(PSys(i).strID) Then MaxLen = Len(PSys(i).strID)
         End If
       Next i
       
     Case Tag_Tableau  '可变素材
       For i = 0 To N_TabMat - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & TabMat(i).strID & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(TabMat(i).strID)
         Else
           If MaxLen < Len(TabMat(i).strID) Then MaxLen = Len(TabMat(i).strID)
         End If
       Next i
       
     Case Tag_Sound  '声音
       For i = 0 To N_Sound - 1
         .ListItems.Add , Tags(TagNo) & "_" & i, Chr(34) & Sounds(i).sndName & Chr(34)
         
         If i = 0 Then
           MaxLen = Len(Sounds(i).sndName)
         Else
           If MaxLen < Len(Sounds(i).sndName) Then MaxLen = Len(Sounds(i).sndName)
         End If
       Next i
       
     Case Else
       LastLoadTag = TagNo
       Exit Function
   End Select
       
   FillItems = True
   LastLoadTag = TagNo
   .Width = TextWidth(Space(MaxLen + 5))
   .ColumnHeaders(1).Width = 0.9 * .Width
   .ListItems(1).Selected = True
End With

End Function





Private Sub lstTips_DblClick()
Dim n As Long, i As Long, j As Long, l As Single, Remark As String, Remark_Start As Long, Params() As String, Params_Start() As Long, CodeLine As String, Pointer As Long, head As Long
Dim oItem As ListItem

If CustomActive Then

    Pointer = txtMain.SelStart
    CodeLine = SplitCodeLine(txtMain.Text, Pointer, head)
    n = PursePYLine(CodeLine, Pointer, Params(), Params_Start(), Remark, Remark_Start)
  
    If n = -1 Then Exit Sub
    If lstTips.SelectedItem Is Nothing Then Exit Sub
    Set oItem = lstTips.SelectedItem
    ReplaceRichText head + Params_Start(n) - 2, Len(Params(n)), oItem.Text, True
    MainRich_SelStart = head + Params_Start(n) - 2 + Len(oItem.Text)
    
    ListTips_DblClick = True
End If
End Sub


Private Sub lstTips_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If lstTips.Visible Then
    Call lstTips_DblClick
  End If
  
ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
  
Else
  txtMain.SetFocus
End If
End Sub


Private Sub lstTips_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If lstTips.Visible Then
    If ListTips_DblClick Then
      txtMain.SetFocus
      txtMain.SelStart = MainRich_SelStart
      ShowParamTip
      ListTips_DblClick = False
    End If
  End If
End If
End Sub

Private Sub lstTips_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ListTips_DblClick Then
  txtMain.SetFocus
  txtMain.SelStart = MainRich_SelStart
  ShowParamTip
  ListTips_DblClick = False
End If
End Sub

Private Sub txtMain_Change()
If CustomActive Then
  ColorRemark = True
  Delayer.Enabled = True
  
  RaiseEvent Change
End If
End Sub

Private Sub txtMain_Click()
Delayer.Enabled = True

RaiseEvent Click
End Sub

Private Sub ShowParamTip()
Dim n As Long, i As Long, j As Long, l As Single, Remark As String, Remark_Start As Long, Params() As String, Params_Start() As Long, CodeLine As String, Pointer As Long, head As Long
Dim CaretPos As POINTAPI, s As Long, Comma As String, TagNo As Long, oItem As ListItem
Dim Op_Index As Long, strTem As String, q As Boolean
Pointer = txtMain.SelStart
CodeLine = SplitCodeLine(txtMain.Text, Pointer, head)
n = PursePYLine(CodeLine, Pointer, Params(), Params_Start(), Remark, Remark_Start)

  '注释
  If Remark_Start > 0 And ColorRemark Then
    SetRichTextColor head + Remark_Start - 2, Len(Remark) + 1, RGB(0, 128, 0), vbBlack
    ColorRemark = False
  End If
  
If n = -1 Then
  PicTip.Visible = False
  lstTips.Visible = False
  RaiseEvent TipShow(PicTip.Visible)
  RaiseEvent ListShow(lstTips.Visible)
  Exit Sub
End If

s = GetCaretPos(CaretPos)

If s <> 0 Then
  PicTip.Cls
  With PicTip
    '参数提示
    Op_Index = GetOpIndexbyName(Params(0))
    If Op_Index <> -1 Then
    
      .Top = CaretPos.Y + txtMain.Top + TextHeight(CodeLine) * 2
      .Left = txtMain.Left + TextWidth(Space(Params_Start(0)))
      ReLocatePicTip
      
      lstTips.Visible = False
      .Visible = True
      RaiseEvent TipShow(.Visible)
      RaiseEvent ListShow(lstTips.Visible)
      
      Comma = IIf(Operation(Op_Index).ParaNum > 0, ",", "")
      
      strTem = "(" & Operation(Op_Index).Op_name & Comma
      l = TextWidth(strTem)
    
      PicTip.Print strTem;
      For i = 1 To Operation(Op_Index).ParaNum
        .FontBold = i = n
        Comma = IIf(i < Operation(Op_Index).ParaNum, ",", "")
        
        strTem = "<" & Operation(Op_Index).Para(i).Value & ">" & Comma
        
        l = l + TextWidth(strTem)
        PicTip.Width = l + TextWidth(Space(2))
        PicTip.Print strTem;
        
        If .FontBold Then .FontBold = False
      Next i
      
       .Width = l + TextWidth(Space(4))
       PicTip.Print "),"
       
    Else
      .Visible = False
      RaiseEvent TipShow(.Visible)
      '自动匹配（操作）
      If n = 0 Then
        If LastLoadTag <> 0 Then
          FillOperations
          LastLoadRes = True
        End If
        lstTips.Top = CaretPos.Y + txtMain.Top + TextHeight(CodeLine) * 2
        lstTips.Left = txtMain.Left + TextWidth(Space(Params_Start(n)))
        ReLocateListTip
        
        lstTips.Visible = True
        RaiseEvent ListShow(lstTips.Visible)
        
        Set oItem = lstTips.FindItem(Params(0), , , 1)
        If Not (oItem Is Nothing) Then
          oItem.EnsureVisible
          oItem.Selected = True
        End If
        
      End If
    End If
  End With
  
  '自动匹配（参数）
  If n > 0 Then
    TagNo = GetPYTag(Params(n))
    If TagNo = 0 Then Exit Sub
        If LastLoadTag <> TagNo Then
          LastLoadRes = FillItems(TagNo)
        End If
        
        If LastLoadRes Then
          Set oItem = lstTips.FindItem(Params(n))
          If oItem Is Nothing Then
            lstTips.Top = CaretPos.Y + txtMain.Top + TextHeight(CodeLine) * 2
            lstTips.Left = txtMain.Left + TextWidth(Space(Params_Start(n)))
            ReLocateListTip
      
            lstTips.Visible = True
            RaiseEvent ListShow(lstTips.Visible)
      
            Set oItem = lstTips.FindItem(Params(n), , , 1)
            If Not (oItem Is Nothing) Then
             oItem.EnsureVisible
             oItem.Selected = True
            End If
          End If
        Else
          lstTips.Visible = False
          RaiseEvent ListShow(lstTips.Visible)
        End If
  End If
End If

End Sub


Private Sub ReplaceRichText(ByVal Start As Long, ByVal Length As Long, ByVal Text As String, Optional RestorePointer As Boolean = True)
Dim i As Long, j As Long, l As Long

CustomActive = False
If RestorePointer Then
  j = txtMain.SelStart
  l = txtMain.SelLength
End If

txtMain.SelStart = Start
txtMain.SelLength = Length
txtMain.SelText = Text

If RestorePointer Then
  txtMain.SelStart = j
  txtMain.SelLength = l
End If

CustomActive = True
End Sub

Private Sub SetRichTextColor(ByVal Start As Long, ByVal Length As Long, ByVal lColor As Long, Optional DefColor As Long)
Dim i As Long, j As Long, l As Long

CustomActive = False
j = txtMain.SelStart
l = txtMain.SelLength

txtMain.SelStart = Start
txtMain.SelLength = Length
txtMain.SelColor = lColor

txtMain.SelStart = j
txtMain.SelLength = l

If Not IsMissing(DefColor) Then
   txtMain.SelColor = DefColor
End If

CustomActive = True
End Sub

Private Sub txtMain_DblClick()
RaiseEvent DblClick
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
  Delayer.Enabled = True
ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
  If Not lstTips.Visible Then
    Delayer.Enabled = True
  Else
    lstTips.SetFocus
  End If
  
End If

End Sub

Public Sub Initialize()

PicTip.Height = TextHeight("a") + 4

CustomActive = True
ColorRemark = False
ListTips_DblClick = False

InitLstOp

LastLoadTag = 0
FillOperations
AutoSwitchLine txtMain, False
End Sub

Private Sub UserControl_Resize()
txtMain.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Property Get Text() As String
Text = txtMain.Text
End Property

Public Property Let Text(ByVal vNewValue As String)
txtMain.Text = vNewValue

PropertyChanged Text
End Property

Public Property Get SelText() As String
SelText = txtMain.SelText
End Property

Public Property Let SelText(ByVal vNewValue As String)
txtMain.SelText = vNewValue

PropertyChanged SelText
End Property

Public Property Get SelStart() As Long
SelStart = txtMain.SelStart
End Property

Public Property Let SelStart(ByVal vNewValue As Long)
txtMain.SelStart = vNewValue

PropertyChanged SelStart
End Property

Public Property Get SelLength() As Long
SelLength = txtMain.SelLength
End Property

Public Property Let SelLength(ByVal vNewValue As Long)
txtMain.SelLength = vNewValue

PropertyChanged SelLength
End Property

Public Property Get SelColor() As Long
SelColor = txtMain.SelColor
End Property

Public Property Let SelColor(ByVal vNewValue As Long)
txtMain.SelColor = vNewValue

PropertyChanged SelColor
End Property

Private Sub ReLocateListTip()
If lstTips.Left + lstTips.Width > UserControl.ScaleWidth Then
  If lstTips.Width < UserControl.ScaleWidth Then
    lstTips.Left = UserControl.ScaleWidth - lstTips.Width   '定位
  End If
End If
End Sub

Private Sub ReLocatePicTip()
With PicTip
If .Left + .Width > UserControl.ScaleWidth Then
  If .Width < UserControl.ScaleWidth Then
     .Left = UserControl.ScaleWidth - .Width    '定位
  End If
End If
End With
End Sub

