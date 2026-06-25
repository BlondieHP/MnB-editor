VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubMenuforMS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "子菜单"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtEnter 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin MSComctlLib.ListView lstMenu 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9975
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Image ImgEnter 
      Height          =   360
      Left            =   2760
      MouseIcon       =   "frmSubMenuforMS.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmSubMenuforMS.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmSubMenuforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TagNo As Integer, Index As String, ParaType As String, TemType As String, Value As String
Dim TemItem As ListItem
Dim ctlParent As MenuforMS
Dim frmSub As New frmSubMenuforMS

Private Sub Form_Deactivate()
ctlParent.HideMenu
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim TemLong As Long, temStr As String
If KeyCode = MENU_KEY_OK Then
  temStr = AssignValue()
  ctlParent.Event_ItemSelect temStr, ParaType
ElseIf KeyCode = MENU_KEY_CANCEL Then
  Call Form_Deactivate
End If
End Sub

Private Sub Form_Load()
InitlstMenu
FillMenu
End Sub

Private Sub InitlstMenu()
Dim n As Integer
n = 2
With lstMenu
      .Sorted = False
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
      .View = lvwReport
        
      .HideColumnHeaders = True
      .ColumnHeaders.Add , , PublicMsgs(13)
      .ColumnHeaders.Add , , PublicMsgs(163)

End With

End Sub

Private Sub Form_Resize()
'lstMenu.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub FillMenu()
Dim i As Integer, oItem As ListItem, temStr() As String, MaxWidth(1) As Single, q As Boolean, mType As Long, Reged As Boolean, values() As String

q = EnumEntities(TagNo, "[csvname]", temStr)
lstMenu.ListItems.Clear
txtEnter.Text = ""
mType = MenuType(ParaType)

If TagNo > 0 Then
  Reged = TagReg(TagNo)
Else
  Reged = True
End If

If TagNo = Tag_Register Or TagNo = Tag_Variable Or TagNo = Tag_Local_Variable Or (Not Reged) Or (TagNo = Tags_End And mType <> MENU_LIST_ONLY) Then
  txtEnter.Top = 0
  txtEnter.Left = 0
  txtEnter.Visible = True
  ImgEnter.Top = 0
  ImgEnter.Left = txtEnter.Width
  lstMenu.Visible = False
  
  If TagNo = Tag_Register And Index > 0 Then
    txtEnter.Text = PublicTags(TagNo) & Index
  ElseIf TagNo = Tag_Variable And Index >= 0 Then
    txtEnter.Text = PYTags(Tag_Variable) & GetVariableName(CLng(Index), True)
  ElseIf TagNo = Tag_Local_Variable And Index >= 0 Then
    txtEnter.Text = PYTags(Tag_Local_Variable) & GetVariableName(CLng(Index), False)
  ElseIf Not TagReg(TagNo) And Index >= 0 Then
    txtEnter.Text = PublicTags(TagNo) & Index
  End If
  
  If TagNo = Tags_End Then
    If TemType = "pos" Then     'ends-add
      txtEnter.Text = PublicTags(Tags_End + 1) & Value
    ElseIf TemType = "s" Then
      txtEnter.Text = PublicTags(Tags_End + 10) & Value
    Else
      txtEnter.Text = Value
    End If
  End If
Else
  lstMenu.Top = 0
  lstMenu.Left = 0
  lstMenu.Visible = True
  txtEnter.Visible = False
End If

If TagNo = Tag_Variable Or TagNo = Tag_Local_Variable Or (TagNo = Tags_End And mType = MENU_TEXT_AND_LIST) Then
  lstMenu.Top = txtEnter.Height
  lstMenu.Visible = True
End If

If q Then
  LoadItems temStr, values, False
ElseIf TagNo = Tags_End Then
  q = EnumConsts(ParaType, "[csvname]", temStr)
  If q Then
    EnumConsts ParaType, "[value]", values
    LoadItems temStr, values, True
  End If
End If

Me.Width = lstMenu.Width
Me.Height = IIf(lstMenu.Visible, lstMenu.Height + lstMenu.Top, txtEnter.Height)
End Sub

Private Sub LoadItems(strAry() As String, varAry() As String, Optional hasValues As Boolean = False)
Dim i As Integer, oItem As ListItem, MaxWidth(1) As Single, TemWidth As Single, q As Boolean
  For i = LBound(strAry) To UBound(strAry)
    If (TagNo <> Tag_Variable And TagNo <> Tag_Local_Variable) Or Len(strAry(i)) > Len(PYTags(TagNo)) Then
      Set oItem = lstMenu.ListItems.Add(, , i)
      oItem.SubItems(1) = strAry(i)
      If hasValues Then
        oItem.Tag = varAry(i)
      Else
        oItem.Tag = i
      End If

      If (TagNo <> Tags_End And oItem.Tag = Index) Or (TagNo = Tags_End And CStr(i) = Value) Then
          txtEnter.Text = oItem.SubItems(1)
          oItem.Selected = True
          oItem.EnsureVisible
          oItem.Bold = True
          oItem.ListSubItems(1).Bold = True
      End If
    
      If Not q Then
        MaxWidth(0) = TextWidth(CStr(i) & Space(2))
      
        TemWidth = TextWidth(strAry(i) & Space(2))
        If oItem.Bold Then TemWidth = TemWidth * 1.2
        MaxWidth(1) = TemWidth
        q = True
      Else
        TemWidth = TextWidth(CStr(i) & Space(2))
        If MaxWidth(0) < TemWidth Then MaxWidth(0) = TemWidth
      
        TemWidth = TextWidth(strAry(i) & Space(2))
        If oItem.Bold Then TemWidth = TemWidth * 1.2
        If MaxWidth(1) < TemWidth Then MaxWidth(1) = TemWidth
      End If
    End If
  Next i

  For i = 0 To 1
    lstMenu.ColumnHeaders(i + 1).Width = MaxWidth(i)
  Next i
End Sub

Private Sub ImgEnter_Click()
Call Form_KeyUp(MENU_KEY_OK, 0)
End Sub

Private Sub lstMenu_DblClick()
Dim temStr As String
If Not TemItem Is Nothing Then
  Index = TemItem.Tag
  temStr = AssignValue()
  ctlParent.Event_ItemSelect temStr, ParaType
  
  Me.Hide
  Set TemItem = Nothing
End If
End Sub

Private Sub lstMenu_ItemClick(ByVal Item As MSComctlLib.ListItem)

If TagNo = Tag_Variable Or TagNo = Tag_Local_Variable Then
  txtEnter.Text = Item.SubItems(1)
End If

Set TemItem = Item
End Sub

Private Function AssignValue() As String
Dim TemLong As Long, temStr As String, n As Long, mType As Long, q As Boolean, DefIndex As Long, s As Long
mType = MenuType(ParaType)
If txtEnter.Text = "" Then
  If TagNo = Tag_Variable Or TagNo = Tag_Local_Variable Or TagNo = Tag_Register Or (TagNo = Tags_End And mType <> MENU_LIST_ONLY) Then
    AssignValue = ""
    Exit Function
  End If
End If
  If TagNo = Tag_Variable Or TagNo = Tag_Local_Variable Then
    TemLong = IsVariableIDExist(txtEnter.Text)
    If TemLong = -1 Then
      s = MsgBox(PublicMsgs(167), vbInformation + vbYesNo, PublicMsgs(0))
      If s = vbYes Then
        DefIndex = -1
      Else
        DefIndex = Index
      End If
    Else
      DefIndex = -1
    End If
    TemLong = GetVariableID(txtEnter.Text, DefIndex)
    temStr = getTXTID(TagNo, TemLong)
  ElseIf TagNo = Tag_Register Or TagReg(TagNo) = False Then
    temStr = Replace(txtEnter.Text, PublicTags(TagNo), "", , , vbTextCompare)
    temStr = Replace(temStr, PYTags(TagNo), "", , , vbTextCompare)
    If IsNumeric(temStr) Then
      TemLong = Val(temStr)
      temStr = getTXTID(TagNo, TemLong)
    Else
      temStr = "NA"
    End If
  ElseIf TagNo = Tags_End Then
    If mType = MENU_LIST_ONLY Then
      temStr = Index
    Else
      q = GetNoTagParamInfo(txtEnter.Text, ParaType, temStr)
      If Not q Then
        temStr = "NA"
        Exit Function
      End If
    End If

  Else
    If Index >= 0 Then
      temStr = getTXTID(TagNo, CLng(Val(Index)))
    Else
      temStr = "NA"
    End If
  End If
AssignValue = temStr
End Function

Public Sub Initialize(lParent As MenuforMS)
Set ctlParent = lParent
End Sub

