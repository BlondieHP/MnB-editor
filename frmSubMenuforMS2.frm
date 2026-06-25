VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSubMenuforMS2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Hider 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   2160
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
Attribute VB_Name = "frmSubMenuforMS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctlParent As MenuforMS, frmSub As frmSubMenuforMS
Dim NoHide As Boolean, Focused As Boolean
Public TagNo As Integer, Value As String, Index As Integer, ParaType As String, TemType As String
Dim TemItem As ListItem

Private Sub Form_Activate()
Focused = True
End Sub

Public Sub Initialize(lParent As MenuforMS, lSub As frmSubMenuforMS)
Set ctlParent = lParent
Set frmSub = lSub
End Sub

Private Sub Form_Deactivate()
Focused = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = MENU_KEY_CANCEL Then
  Call Form_Deactivate
  ctlParent.HideMenu
End If
End Sub

Private Sub Form_Load()
InitlstMenu
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
      .ColumnHeaders.Add , , "²Ëµ¥"
End With

End Sub

Public Sub ShowMenu()
Dim i As Integer, oItem As ListItem, MaxWidth As Single, CanShow As Boolean

lstMenu.ListItems.Clear
If TagNo = 0 Then TagNo = Tags_End
  
  Set oItem = lstMenu.ListItems.Add(, , PublicMsgs(166))
  oItem.Tag = ""
  MaxWidth = TextWidth(PublicMsgs(166) & Space(2))
  If ParaType = oItem.Tag Then oItem.Bold = True
  
  'ends_add
  If ParaType <> "pos" And ParaType <> "s" Then
    'itp
    Set oItem = lstMenu.ListItems.Add(, , PublicTags(Tags_End + 2))
    oItem.Tag = "itp"
      If MaxWidth < TextWidth(PublicTags(Tags_End + 2) & Space(2)) Then
      MaxWidth = TextWidth(PublicTags(Tags_End + 2) & Space(2))
    End If
    If ParaType = oItem.Tag Then oItem.Bold = True
    
    'tf
    Set oItem = lstMenu.ListItems.Add(, , PublicTags(Tags_End + 3))
    oItem.Tag = "tf"
      If MaxWidth < TextWidth(PublicTags(Tags_End + 3) & Space(2)) Then
      MaxWidth = TextWidth(PublicTags(Tags_End + 3) & Space(2))
    End If
    If ParaType = oItem.Tag Then oItem.Bold = True
    
    'bs
    Set oItem = lstMenu.ListItems.Add(, , PublicTags(Tags_End + 4))
    oItem.Tag = "bs"
      If MaxWidth < TextWidth(PublicTags(Tags_End + 4) & Space(2)) Then
      MaxWidth = TextWidth(PublicTags(Tags_End + 4) & Space(2))
    End If
    If ParaType = oItem.Tag Then oItem.Bold = True
    
    'ai_bhvr
    Set oItem = lstMenu.ListItems.Add(, , PublicTags(Tags_End + 5))
    oItem.Tag = "ai_bhvr"
      If MaxWidth < TextWidth(PublicTags(Tags_End + 5) & Space(2)) Then
      MaxWidth = TextWidth(PublicTags(Tags_End + 5) & Space(2))
    End If
    If ParaType = oItem.Tag Then oItem.Bold = True
    
    'po
    Set oItem = lstMenu.ListItems.Add(, , PublicTags(Tags_End + 6))
    oItem.Tag = "po"
      If MaxWidth < TextWidth(PublicTags(Tags_End + 6) & Space(2)) Then
      MaxWidth = TextWidth(PublicTags(Tags_End + 6) & Space(2))
    End If
    If ParaType = oItem.Tag Then oItem.Bold = True

    'pf
    Set oItem = lstMenu.ListItems.Add(, , PublicTags(Tags_End + 7))
    oItem.Tag = "pf"
      If MaxWidth < TextWidth(PublicTags(Tags_End + 7) & Space(2)) Then
      MaxWidth = TextWidth(PublicTags(Tags_End + 7) & Space(2))
    End If
    If ParaType = oItem.Tag Then oItem.Bold = True

    'ap
    Set oItem = lstMenu.ListItems.Add(, , PublicTags(Tags_End + 8))
    oItem.Tag = "ap"
      If MaxWidth < TextWidth(PublicTags(Tags_End + 8) & Space(2)) Then
      MaxWidth = TextWidth(PublicTags(Tags_End + 8) & Space(2))
    End If
    If ParaType = oItem.Tag Then oItem.Bold = True
    
    'as
    Set oItem = lstMenu.ListItems.Add(, , PublicTags(Tags_End + 9))
    oItem.Tag = "as"
      If MaxWidth < TextWidth(PublicTags(Tags_End + 9) & Space(2)) Then
      MaxWidth = TextWidth(PublicTags(Tags_End + 9) & Space(2))
    End If
    If ParaType = oItem.Tag Then oItem.Bold = True
  End If
  
  TemType = ParaType
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
ctlParent.ShutMenu 2, Msg
NoHide = False
Hider.Enabled = False
End Sub

Private Sub lstMenu_ItemClick(ByVal Item As MSComctlLib.ListItem)
Set TemItem = Item


End Sub

Private Sub lstMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo EL

Dim Item As ListItem
If TemItem Is Nothing Then
  Exit Sub
End If
NoHide = True

Set Item = TemItem

    frmSub.TagNo = Tags_End
    frmSub.ParaType = Item.Tag
    frmSub.TemType = TemType
    frmSub.Value = Value
    
    If frmSub.TagNo = TagNo Then
      frmSub.Index = Index
    Else
      frmSub.Index = -1
    End If
    
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


Set TemItem = Nothing

Exit Sub
EL:
    Call logErr("lstMenu", "lstMenu_MouseUp", Err.Number, Err.Description)
End Sub
