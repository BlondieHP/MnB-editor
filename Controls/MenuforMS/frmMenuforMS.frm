VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenuforMS 
   BorderStyle     =   0  'None
   Caption         =   "菜单"
   ClientHeight    =   6060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Hider 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2760
      Top             =   1080
   End
   Begin MSComctlLib.ListView lstMenu 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
End
Attribute VB_Name = "frmMenuforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Mode As Long
Dim frmSub As New frmSubMenuforMS, frmSub2 As New frmSubMenuforMS2
Dim ctlParent As MenuforMS
Dim NoHide As Boolean, Focused As Boolean
Public TagNo As Integer, Index As Integer, ParaType As String
Public Value As String, TemType As String             'for overflaw value
Dim TemItem As ListItem

Public Sub Initialize(lParent As MenuforMS)
Set ctlParent = lParent
frmSub.Initialize lParent
frmSub2.Initialize lParent, frmSub
End Sub

Private Sub InitlstMenu()
Dim n As Integer
n = 1
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
      .Visible = True
      .View = lvwReport
        
      .HideColumnHeaders = True
      .ColumnHeaders.Add , , "菜单"
End With

End Sub

Private Sub Form_Activate()
Focused = True
End Sub

Private Sub Form_Deactivate()
If Not NoHide Then
  frmSub2.Hide
  frmSub.Hide
  Me.Hide
Else
  NoHide = False
End If

Focused = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = MENU_KEY_CANCEL Then
  Call Form_Deactivate
End If
End Sub

Private Sub Form_Load()
InitlstMenu

End Sub

Public Sub HideMenu(Optional HideNo As String = "0|1|2")
Dim strTem() As String, i As Integer

strTem = Split(HideNo, "|")

For i = 0 To UBound(strTem)
  If strTem(i) = "0" Then
    Me.Hide
    NoHide = False
  ElseIf strTem(i) = "1" Then
    frmSub.Hide
  ElseIf strTem(i) = "2" Then
    frmSub2.Hide
  End If
Next i

End Sub

Public Sub HideMenu2()
'frmSub.Hide
Hider.Enabled = True
frmSub2.Hider.Enabled = True
End Sub

Private Sub Form_Resize()
'lstMenu.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub ShowMenu()
Dim i As Integer, oItem As ListItem, MaxWidth As Single, CanShow As Boolean

lstMenu.ListItems.Clear
If TagNo = 0 Then TagNo = Tags_End

  For i = 1 To 26
    If ParaType <> "" And i <> Tag_Register And i <> Tag_Variable And i <> Tag_Local_Variable And i <> Tags_End Then
      CanShow = Val(ParaType) = i
    Else
      CanShow = True
    End If
      
    If CanShow Then
      Set oItem = lstMenu.ListItems.Add(, , PublicTags(i))
      oItem.Tag = i
      
      If i = 26 Then
        oItem.ForeColor = vbBlue
      End If
      
      If i = TagNo Then
        oItem.Selected = True
        oItem.Bold = True
      End If

      If i = 1 Then
        MaxWidth = TextWidth(PublicTags(i) & Space(2))
      Else
        If MaxWidth < TextWidth(PublicTags(i) & Space(2)) Then
          MaxWidth = TextWidth(PublicTags(i) & Space(2))
        End If
      End If
    End If
  Next i

Me.Show
lstMenu.ColumnHeaders(1).Width = MaxWidth
lstMenu.Width = lstMenu.ColumnHeaders(1).Width
lstMenu.Height = (lstMenu.ListItems.Count + 0.5) * lstMenu.ListItems(1).Height
Me.Width = lstMenu.Width
Me.Height = lstMenu.Height


End Sub


Private Sub Hider_Timer()
Dim Msg As Long

Msg = IIf(Focused, MENU_MSG_ACTIVE, MENU_MSG_DEACTIVE)
ctlParent.ShutMenu 0, Msg
NoHide = False

Hider.Enabled = False
End Sub

Private Sub lstMenu_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set TemItem = Item


End Sub

Private Sub lstMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Item As ListItem
If TemItem Is Nothing Then
  Exit Sub
End If
NoHide = True

Set Item = TemItem

  frmSub.TagNo = Val(Item.Tag)
  
  If frmSub.TagNo <> Tags_End Then
    If frmSub.TagNo = TagNo Then
      frmSub.Index = Index
    Else
      frmSub.Index = -1
    End If
    frmSub.ParaType = ParaType
    frmSub.Show
    frmSub.FillMenu
    
    frmSub.Top = Me.Top + Item.Top
    frmSub.Left = Me.Left + Me.Width

    If frmSub.Top + frmSub.Height > Screen.Height Then
      frmSub.Top = frmSub.Top - frmSub.Height
    End If

    If frmSub.Left + frmSub.Width > Screen.Width Then
      frmSub.Left = frmSub.Left - frmSub.Width - Me.Width
    End If
    frmSub.SetFocus
  Else
    frmSub2.TagNo = Tags_End
    frmSub2.ParaType = ParaType
    frmSub2.TemType = TemType
    frmSub2.Value = Value
    frmSub2.Show
    frmSub2.ShowMenu
    frmSub2.Top = Me.Top + Item.Top
    frmSub2.Left = Me.Left + Me.Width

    If frmSub2.Top + frmSub2.Height > Screen.Height Then
      frmSub2.Top = frmSub2.Top - frmSub2.Height
    End If

    If frmSub2.Left + frmSub2.Width > Screen.Width Then
      frmSub2.Left = frmSub2.Left - frmSub2.Width
    End If
    frmSub2.SetFocus
  End If


Set TemItem = Nothing

End Sub
