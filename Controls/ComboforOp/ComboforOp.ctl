VERSION 5.00
Begin VB.UserControl ComboforOp 
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1230
   ScaleWidth      =   1470
   Begin VB.Timer Delayer 
      Enabled         =   0   'False
      Index           =   10
      Left            =   240
      Top             =   480
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   0
      Picture         =   "ComboforOp.ctx":0000
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "ComboforOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim frmCombo As New frmComboforOp
Public Value As String
Dim p_Text As String
Dim Initialized As Boolean

Public Event ItemSelect(strValue As String, strText As String)

Public Sub ShowMenu(X As Long, Y As Long, Optional Width As Long, Optional Height As Long)

If Width > 0 Then frmCombo.Width = Width
If Height > 0 Then frmCombo.Height = Height

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

Public Property Get MenuWidth() As Single
If Initialized Then MenuWidth = frmCombo.Width
End Property

Public Property Let MenuWidth(ByVal vNewValue As Single)
If Initialized Then
  frmCombo.Width = vNewValue

  PropertyChanged "MenuWidth"
End If
End Property

Public Property Get MenuHeight() As Single
If Initialized Then MenuHeight = frmCombo.Height
End Property

Public Property Let MenuHeight(ByVal vNewValue As Single)
If Initialized Then
  frmCombo.Height = vNewValue

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

Private Sub UserControl_Resize()
UserControl.Height = Image1.Height
UserControl.Width = Image1.Width
End Sub
