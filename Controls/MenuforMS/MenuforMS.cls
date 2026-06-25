VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuforMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim frmMenu As New frmMenuforMS
Public Value As String

Public Sub ShowMenu(X As Long, Y As Long)
Dim TagNo As Integer, Pid As String, Idx As Integer
GetParamCodeInfo Value, TagNo, Pid

If TagNo > 0 Then Idx = Val(Pid)

frmMenu.Top = Y * Screen.TwipsPerPixelY
frmMenu.Left = X * Screen.TwipsPerPixelX
frmMenu.TagNo = TagNo
frmMenu.Index = Idx

frmMenu.ShowMenu

If frmMenu.Top + frmMenu.Height > Screen.Height Then
  frmMenu.Top = Y * Screen.TwipsPerPixelY - frmMenu.Height
End If

If frmMenu.Left + frmMenu.Width > Screen.Width Then
  frmMenu.Left = X * Screen.TwipsPerPixelY - frmMenu.Width
End If

End Sub
