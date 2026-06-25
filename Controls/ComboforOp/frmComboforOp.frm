VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComboforOp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtEnter 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin MSComctlLib.ListView LstData 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
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
   Begin MSComctlLib.ListView LstTips 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9340
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
   Begin VB.Image ImgEnter 
      Height          =   360
      Left            =   2760
      MouseIcon       =   "frmComboforOp.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmComboforOp.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "frmComboforOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Text As String, Value As String, DefText As String
Dim ctlParent As ComboforOp
Dim CustomActive As Boolean

Private Sub Form_Deactivate()
Me.Hide
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = MENU_KEY_CANCEL Then
  Call Form_Deactivate
End If
End Sub

Private Sub Form_Load()
InitLstOp

InitItems

Me.Hide
CustomActive = True
End Sub

Private Sub Form_Resize()
txtEnter.Move 0, 0, Me.ScaleWidth - ImgEnter.Width - 20
ImgEnter.Move txtEnter.Width + 20, 0
LstTips.Move 0, txtEnter.Height, Me.ScaleWidth, Me.ScaleHeight - txtEnter.Height

End Sub

Private Sub InitLstOp()
Dim n As Integer
n = 1
With LstTips
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
      .Visible = True

      .ColumnHeaders.Add , , "提示"
End With

With LstData
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

      .ColumnHeaders.Add , , "数据"
End With
End Sub

Private Sub InitItems()
Dim i As Integer, oItem As ListItem
LstData.ListItems.Clear
For i = 0 To UBound(Operation)
  Set oItem = LstData.ListItems.Add(, "Op_" & i, Operation(i).Pseudo)
  oItem.Tag = Operation(i).OpID
Next i
End Sub

Public Sub InitMenu()
txtEnter.Text = Text
AssignText Text
txtEnter.SelStart = Len(Text)
txtEnter.SetFocus
End Sub

Public Sub AssignText(strText As String)
Dim n As Long, oItem As ListItem, oItem2 As ListItem, MaxWidth As Single
n = 0
  CustomActive = False
  txtEnter.Text = strText
  LstTips.ListItems.Clear
  Do While n + 1 <= LstData.ListItems.Count
    'Set oItem = LstData.FindItem(txtEnter.Text, lvwText, n + 1, lvwPartial)
    Set oItem = FindItem(LstData, n, "0", txtEnter.Text, True, vbTextCompare)
    
    If oItem Is Nothing Then
      Exit Do
    Else
      If MaxWidth = 0 Then
        MaxWidth = TextWidth(oItem.Text & Space(2))
      Else
        If MaxWidth < TextWidth(oItem.Text & Space(2)) Then
          MaxWidth = TextWidth(oItem.Text & Space(2))
        End If
      End If
      n = oItem.Index
      If Trim(oItem.Text) <> "" Then
        Set oItem2 = LstTips.ListItems.Add(, oItem.Key, oItem.Text)
        oItem2.Tag = oItem.Tag
        oItem2.Bold = IsControlOp(Val(oItem.Tag))
      End If
      'DoEvents
    End If
  Loop
  
  LstTips.ColumnHeaders(1).Width = MaxWidth
  'MeasureListTip
  CustomActive = True
End Sub

Private Sub ImgEnter_Click()
Dim oItem As ListItem

If DefText = txtEnter.Text Then
  Me.Hide
  Exit Sub
End If

Set oItem = LstTips.FindItem(txtEnter.Text, lvwText, , lvwWhole)
If Not oItem Is Nothing Then
  ctlParent.Event_ItemSelect oItem.Tag, oItem.Text
Else
  MsgBox ActiveString(PublicMsgs(162), txtEnter.Text), vbExclamation, PublicMsgs(19)
End If

Me.Hide
End Sub

Private Sub lstTips_DblClick()
If Not LstTips.SelectedItem Is Nothing Then
  ctlParent.Event_ItemSelect LstTips.SelectedItem.Tag, LstTips.SelectedItem.Text
  Me.Hide
End If

End Sub

Public Sub Initialize(lParent As ComboforOp)
Set ctlParent = lParent
End Sub

Private Sub LstTips_ItemClick(ByVal Item As MSComctlLib.ListItem)
CustomActive = False
txtEnter.Text = Item.Text
CustomActive = True
End Sub

Private Sub txtEnter_Change()
If CustomActive Then
  AssignText txtEnter.Text
End If
End Sub

Private Sub txtEnter_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 Then
  If KeyCode = vbKeyA Then
    txtEnter.SelStart = 0
    txtEnter.SelLength = Len(txtEnter.Text)
  End If
End If
End Sub

Private Sub txtEnter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call ImgEnter_Click
End If
End Sub

