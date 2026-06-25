VERSION 5.00
Begin VB.UserControl MenuforMS 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   615
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "MenuforMS.ctx":0000
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "MenuforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim frmMenu As New frmMenuforMS
Public Value As String
Public ParaType As String
Public TemType As String
Public Event ItemSelect(Value As String, PType As String)
Private MnuMsg(2) As Long

Public Sub ShutMenu(From As Long, Msg As Long)
MnuMsg(From) = Msg

If MnuMsg(0) = MENU_MSG_DEACTIVE And MnuMsg(2) = MENU_MSG_DEACTIVE Then
  frmMenu.HideMenu
  MnuMsg(0) = 0
  MnuMsg(2) = 0
End If

If MnuMsg(0) = MENU_MSG_ACTIVE And MnuMsg(2) = MENU_MSG_DEACTIVE Then
  frmMenu.HideMenu "1|2"
  MnuMsg(0) = 0
  MnuMsg(2) = 0
End If

If MnuMsg(0) = MENU_MSG_DEACTIVE And MnuMsg(2) = MENU_MSG_ACTIVE Then
  frmMenu.HideMenu "1"
  MnuMsg(0) = 0
  MnuMsg(2) = 0
End If
End Sub

Public Sub ShowMenu(X As Long, Y As Long)
Dim TagNo As Integer, Pid As String, Idx As Integer
GetParamCodeInfo Value, TagNo, Pid

If TagNo > 0 Then
  Idx = Val(Pid)
  frmMenu.Value = ""
Else
  frmMenu.Value = Value
  Idx = -1
End If

frmMenu.Top = Y * Screen.TwipsPerPixelY
frmMenu.Left = X * Screen.TwipsPerPixelX
frmMenu.TagNo = TagNo
frmMenu.Index = Idx
frmMenu.ParaType = ParaType
frmMenu.TemType = TemType

frmMenu.ShowMenu

If frmMenu.Top + frmMenu.Height > Screen.Height Then
  frmMenu.Top = Y * Screen.TwipsPerPixelY - frmMenu.Height
End If

If frmMenu.Left + frmMenu.Width > Screen.Width Then
  frmMenu.Left = X * Screen.TwipsPerPixelY - frmMenu.Width
End If

End Sub

Public Sub Event_ItemSelect(Value As String, PType As String)
If Value <> "NA" Then
  RaiseEvent ItemSelect(Value, PType)
End If
frmMenu.HideMenu
End Sub

Public Sub HideMenu()
  frmMenu.HideMenu2
End Sub

Public Sub Initialize()
frmMenu.Initialize Me
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Image1.Height
UserControl.Width = Image1.Width
End Sub
