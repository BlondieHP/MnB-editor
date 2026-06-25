VERSION 5.00
Begin VB.UserControl ComboEx 
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ScaleHeight     =   435
   ScaleWidth      =   4950
   Begin VB.CommandButton CExpand 
      Caption         =   "¡ý"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtMain 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "ComboEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim frmCombo As New frmComboforOp
Public Value As String
Dim p_Text As String
Dim p_Height As Single
Dim Initialized As Boolean
Public Event ItemSelect(strValue As String, strText As String)


Private Sub UserControl_Resize()
UserControl.Height = txtMain.Height
CExpand.Move UserControl.ScaleWidth - CExpand.Width, 0
txtMain.Move 0, 0, CExpand.Left
End Sub

Public Sub ShowMenu(Optional Height As Long)
Dim X As Long, Y As Long

X = UserControl.Extender.Left
Y = UserControl.Extender.Top
frmCombo.Width = UserControl.Width
If p_Height > 0 Then frmCombo.Height = p_Height

frmCombo.Top = Y * Screen.TwipsPerPixelY
frmCombo.Left = X * Screen.TwipsPerPixelX

If frmCombo.Top + frmCombo.Height > Screen.Height Then
  frmCombo.Top = Y * Screen.TwipsPerPixelY - frmCombo.Height
End If

If frmCombo.Left + frmCombo.Width > Screen.Width Then
  frmCombo.Left = X * Screen.TwipsPerPixelY - frmCombo.Width
End If

frmCombo.Value = Value
frmCombo.DefText = p_Text
frmCombo.Text = p_Text
frmCombo.Show
frmCombo.InitMenu

End Sub

Public Property Get MenuHeight() As Single
If Initialized Then MenuHeight = p_Height
End Property

Public Property Let MenuHeight(ByVal vNewValue As Single)
If Initialized Then
  p_Height = vNewValue

  PropertyChanged "MenuHeight"
End If
End Property

Public Property Get Text() As String
Text = p_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
p_Text = vNewValue
If Initialized Then frmCombo.AssignText p_Text

PropertyChanged "Text"
End Property

Public Sub Initialize()
frmCombo.Initialize Me
Initialized = True
End Sub

Public Sub Event_ItemSelect(sValue As String, sText As String)
RaiseEvent ItemSelect(sValue, sText)
End Sub

Private Sub UserControl_Terminate()
Initialized = False
End Sub

