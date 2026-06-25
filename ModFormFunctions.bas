Attribute VB_Name = "ModFormFunctions"

Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private hBitmap As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

 Private Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
  Private Const GWL_STYLE = (-16)
  Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Dim Xtime As Integer
Dim metop, meheight


Public Sub FormTransparent(Frm As Form)

  Dim TempLng As Long
    
  TempLng = GetWindowLong(Frm.hWnd, GWL_STYLE)
  TempLng = TempLng And Not WS_CAPTION
  SetWindowLong Frm.hWnd, GWL_STYLE, TempLng
    
hBitmap = CreateCompatibleBitmap(Frm.hdc, 0, 0)
SelectObject Frm.hdc, hBitmap
Frm.Refresh
End Sub



Public Sub Release()
DeleteObject hBitmap

End Sub

